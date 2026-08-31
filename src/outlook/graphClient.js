import { config } from '../config/env.js';
import fs from 'fs';
import path from 'path';

function toRecipients(value) {
  if (!value) return [];
  const addresses = Array.isArray(value) ? value : value.split(',');
  return addresses.map(addr => addr.trim()).filter(Boolean).map(address => ({ emailAddress: { address } }));
}

// Direct-attach ceiling per Microsoft Graph (base64 inflates ~33%, JSON body cap is 4MB).
// This is a COMBINED total across all attachments in one call, not per-file —
// two 2.9MB files would each pass individually but blow the real request-size limit together.
// Above this, attachFilesToMessage() below routes through createUploadSession instead.
const MAX_DIRECT_ATTACH_BYTES = 3 * 1024 * 1024;
// Outer ceiling for the createUploadSession path (Graph itself allows up to 150MB per
// file/message) — kept conservative since the chunked upload runs synchronously inside
// one MCP tool call; raise if a real need for bigger files shows up.
const MAX_UPLOAD_SESSION_BYTES = 25 * 1024 * 1024;
// Must be a multiple of 320 KiB per Graph's upload-session contract (except the final
// chunk). 5 MiB = 16 * 320 KiB.
const UPLOAD_CHUNK_BYTES = 5 * 1024 * 1024;
const ATTACHMENT_ROOT = '/opt/data';
const MIME_TYPES = {
  '.pdf': 'application/pdf',
  '.doc': 'application/msword',
  '.docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  '.xls': 'application/vnd.ms-excel',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  '.ppt': 'application/vnd.ms-powerpoint',
  '.pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
  '.csv': 'text/csv',
  '.txt': 'text/plain',
  '.zip': 'application/zip',
};

function resolveAttachmentPaths(filePaths) {
  if (!filePaths || filePaths.length === 0) return [];
  return filePaths.map(filePath => {
    let resolved;
    try { resolved = fs.realpathSync(filePath); }
    catch { throw new Error(`Attachment file not found: ${filePath}`); }
    if (!resolved.startsWith(ATTACHMENT_ROOT + path.sep)) {
      throw new Error(`Attachment path '${filePath}' is outside the shared data directory (${ATTACHMENT_ROOT}) — refusing to read it.`);
    }
    return { resolved, size: fs.statSync(resolved).size };
  });
}

// Small-attachment path: base64-inline in the same request that creates the
// message. Caller must already have confirmed the combined size fits under
// MAX_DIRECT_ATTACH_BYTES — this function does not check.
function buildInlineAttachments(resolvedFiles) {
  return resolvedFiles.map(({ resolved }) => {
    const contentBytes = fs.readFileSync(resolved).toString('base64');
    const contentType = MIME_TYPES[path.extname(resolved).toLowerCase()] || 'application/octet-stream';
    return { '@odata.type': '#microsoft.graph.fileAttachment', name: path.basename(resolved), contentType, contentBytes };
  });
}

// Large-attachment path: the message must already exist (a draft with a real
// id) — createUploadSession attaches to a specific message, it cannot create
// one. Uploads one file in sequential chunks per Graph's resumable-upload
// contract (https://learn.microsoft.com/graph/outlook-large-attachments).
async function uploadAttachmentToMessage(messageId, resolvedFile, userEmail) {
  const { resolved, size } = resolvedFile;
  const name = path.basename(resolved);
  const session = await graphClient.graphFetch(
    `/messages/${messageId}/attachments/createUploadSession`,
    { method: 'POST', body: JSON.stringify({ AttachmentItem: { attachmentType: 'file', name, size } }) },
    userEmail,
  );
  if (!session?.uploadUrl) {
    throw new Error(`Failed to create upload session for '${name}' — Graph API did not return an uploadUrl.`);
  }

  const fd = fs.openSync(resolved, 'r');
  try {
    let offset = 0;
    let lastResponseBody = null;
    while (offset < size) {
      const chunkSize = Math.min(UPLOAD_CHUNK_BYTES, size - offset);
      const buffer = Buffer.alloc(chunkSize);
      fs.readSync(fd, buffer, 0, chunkSize, offset);
      // Upload-session URLs are pre-authorized — no Authorization header, and
      // Graph rejects the request if one is sent alongside a mismatched host.
      const res = await fetch(session.uploadUrl, {
        method: 'PUT',
        headers: {
          'Content-Length': String(chunkSize),
          'Content-Range': `bytes ${offset}-${offset + chunkSize - 1}/${size}`,
        },
        body: buffer,
      });
      if (!res.ok) {
        const errBody = await res.text().catch(() => '');
        throw new Error(`Upload chunk failed for '${name}' at offset ${offset} (HTTP ${res.status}): ${errBody.slice(0, 300)}`);
      }
      offset += chunkSize;
      // Final chunk returns the completed attachment object (200/201); every
      // earlier chunk returns 202 with nextExpectedRanges — only the last
      // response body is the one we want.
      if (offset >= size) {
        const text = await res.text();
        lastResponseBody = text ? JSON.parse(text) : null;
      }
    }
    return lastResponseBody;
  } finally {
    fs.closeSync(fd);
  }
}

// Routes each file through the size-appropriate path and returns the count
// actually attached. Used by the >3MB branch of create/send functions below,
// where the message already exists as a draft (own or send-via-draft).
async function attachFilesToExistingMessage(messageId, resolvedFiles, userEmail) {
  let count = 0;
  for (const file of resolvedFiles) {
    await uploadAttachmentToMessage(messageId, file, userEmail);
    count += 1;
  }
  return count;
}

async function sendExistingDraft(draftId, userEmail) {
  // POST .../send returns 202 Accepted with an empty body on success.
  const resData = await graphClient.graphFetch(`/messages/${draftId}/send`, { method: 'POST' }, userEmail);
  return resData?.success === true;
}

// Folders excluded from getRecentEmailsAllFolders — optional, configured via
// OUTLOOK_EXCLUDED_FOLDER_IDS (comma-separated Graph folder IDs) in env.js.
// Empty/unset means no folders are skipped.
const EXCLUDED_FOLDER_IDS = new Set(config.excludedFolderIds);

class GraphClient {
  constructor() { this.accessToken = null; this.tokenExpiresAt = 0; }

  get isConfigured() {
    const { tenantId, clientId, clientSecret, userEmail } = config.azure;
    return Boolean(tenantId && clientId && clientSecret && userEmail && !tenantId.includes('your_'));
  }

  async getAccessToken() {
    if (this.accessToken && Date.now() < this.tokenExpiresAt - 60000) return this.accessToken;
    const { tenantId, clientId, clientSecret } = config.azure;
    if (!this.isConfigured) return null;
    try {
      const body = new URLSearchParams({ client_id: clientId, client_secret: clientSecret, scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials' });
      const res = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, { method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body });
      const data = await res.json();
      if (res.ok && data.access_token) { this.accessToken = data.access_token; this.tokenExpiresAt = Date.now() + (data.expires_in * 1000); return this.accessToken; }
      console.error('[GraphClient] Token error:', data.error_description || data.error);
    } catch (err) { console.error('[GraphClient] Token exception:', err.message); }
    return null;
  }

  async graphFetch(endpoint, options = {}, userEmail = null) {
    const token = await this.getAccessToken();
    const targetEmail = userEmail || config.azure.userEmail;
    if (!token || !targetEmail) return null;
    const url = endpoint.startsWith('http') ? endpoint : `https://graph.microsoft.com/v1.0/users/${targetEmail}${endpoint}`;
    try {
      const res = await fetch(url, { ...options, headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json', ...(options.headers || {}) } });
      if (res.ok) {
        // 202 Accepted and 204 No Content both have empty bodies — return sentinel
        if (res.status === 202 || res.status === 204) return { success: true };
        const text = await res.text();
        return text ? JSON.parse(text) : { success: true };
      }
      const errData = await res.json().catch(() => ({}));
      console.error(`[GraphClient] ${options.method || 'GET'} ${endpoint} (${res.status}):`, errData.error?.message || 'Unknown error');
    } catch (err) { console.error(`[GraphClient] Fetch exception ${endpoint}:`, err.message); }
    return null;
  }

  async getRecentEmails(limit = 10, userEmail = null) {
    const data = await this.graphFetch(`/mailFolders/inbox/messages?$orderby=receivedDateTime desc&$top=${limit}&$select=id,conversationId,sender,subject,body,receivedDateTime,isRead`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch recent emails from Microsoft Graph API. Please check configuration.' };
    return (data.value || []).map(msg => ({ id: msg.id, conversationId: msg.conversationId, sender: msg.sender?.emailAddress?.address || 'unknown', senderName: msg.sender?.emailAddress?.name || 'Unknown', subject: msg.subject || 'No Subject', body: msg.body?.content || '', receivedDateTime: msg.receivedDateTime, isRead: msg.isRead }));
  }

  async getRecentEmailsAllFolders(limit = 10, userEmail = null) {
    const overfetch = Math.max(limit * 3, 30);
    const data = await this.graphFetch(`/messages?$orderby=receivedDateTime desc&$top=${overfetch}&$select=id,conversationId,sender,subject,bodyPreview,receivedDateTime,isRead,parentFolderId`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch recent emails from Microsoft Graph API. Please check configuration.' };
    return (data.value || [])
      .filter(msg => !config.skipExcludedFolders || !EXCLUDED_FOLDER_IDS.has(msg.parentFolderId))
      .slice(0, limit)
      .map(msg => ({ id: msg.id, conversationId: msg.conversationId, sender: msg.sender?.emailAddress?.address || 'unknown', senderName: msg.sender?.emailAddress?.name || 'Unknown', subject: msg.subject || 'No Subject', bodyPreview: msg.bodyPreview || '', receivedDateTime: msg.receivedDateTime, isRead: msg.isRead }));
  }

  async getUnreadEmails(userEmail = null) {
    const data = await this.graphFetch(`/mailFolders/inbox/messages?$filter=isRead eq false&$orderby=receivedDateTime desc&$top=25&$select=id,conversationId,sender,subject,body,receivedDateTime`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch unread emails from Microsoft Graph API. Please check configuration.' };
    return (data.value || []).map(msg => ({ id: msg.id, conversationId: msg.conversationId, sender: msg.sender?.emailAddress?.address || 'unknown', senderName: msg.sender?.emailAddress?.name || 'Unknown', subject: msg.subject || 'No Subject', body: msg.body?.content || '', receivedDateTime: msg.receivedDateTime }));
  }

  async getConversationThread(conversationId, limit = 10, userEmail = null) {
    const data = await this.graphFetch(`/messages?$filter=conversationId eq '${conversationId}'&$top=${limit}&$select=id,sender,subject,body,sentDateTime,isDraft`, { headers: { 'ConsistencyLevel': 'eventual' } }, userEmail);
    if (!data) return { error: `Failed to fetch conversation thread for ID '${conversationId}'.` };
    return (data.value || [])
      .map(msg => ({ id: msg.id, sender: msg.sender?.emailAddress?.address || '', senderName: msg.sender?.emailAddress?.name || '', subject: msg.subject || '', body: msg.body?.content || '', sentDateTime: msg.sentDateTime, isDraft: msg.isDraft }))
      .sort((a, b) => new Date(a.sentDateTime) - new Date(b.sentDateTime));
  }

  async getAttachmentsForMessage(messageId, userEmail = null) {
    if (!messageId) throw new Error("Missing messageId argument");
    const endpoint = `/messages/${messageId}/attachments`;
    const resp = await this.graphFetch(endpoint, {}, userEmail);
    return resp?.value || [];
  }

  async getAttachmentContent({ messageId, attachmentId }, userEmail = null) {
    if (!messageId || !attachmentId) throw new Error("Missing required arguments: messageId, attachmentId");
    const endpoint = `/messages/${messageId}/attachments/${attachmentId}`;
    const resp = await this.graphFetch(endpoint, {}, userEmail);
    return resp || {};
  }

  async searchEmails(query, folder = 'inbox', limit = 20, userEmail = null) {
    const data = await this.graphFetch(`/messages?$search="${encodeURIComponent(query)}"&$top=${limit}&$select=id,conversationId,sender,subject,body,receivedDateTime,isRead`, { headers: { 'ConsistencyLevel': 'eventual' } }, userEmail);
    if (!data) return { error: `Search query failed for '${query}'.` };
    return (data.value || []).map(msg => ({ id: msg.id, conversationId: msg.conversationId, sender: msg.sender?.emailAddress?.address || 'unknown', senderName: msg.sender?.emailAddress?.name || 'Unknown', subject: msg.subject || 'No Subject', bodyPreview: (msg.body?.content || '').substring(0, 300), receivedDateTime: msg.receivedDateTime, isRead: msg.isRead }));
  }

  async getSentEmails(limit = 25, userEmail = null) {
    const data = await this.graphFetch(`/mailFolders/sentItems/messages?$orderby=sentDateTime desc&$top=${limit}&$select=id,conversationId,toRecipients,subject,body,sentDateTime`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch sent emails from Microsoft Graph API.' };
    return (data.value || []).map(msg => ({ id: msg.id, conversationId: msg.conversationId, to: (msg.toRecipients || []).map(r => r.emailAddress?.address).join(', '), subject: msg.subject || 'No Subject', bodyPreview: (msg.body?.content || '').substring(0, 300), sentDateTime: msg.sentDateTime }));
  }

  async getEmailFolders(userEmail = null) {
    const data = await this.graphFetch(`/mailFolders?$top=50`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch email folders from Microsoft Graph API.' };
    return (data.value || []).map(f => ({ id: f.id, name: f.displayName, unreadCount: f.unreadItemCount, totalCount: f.totalItemCount }));
  }

  async moveEmail(emailId, destinationFolderId, userEmail = null) {
    const resData = await this.graphFetch(`/messages/${emailId}/move`, { method: 'POST', body: JSON.stringify({ destinationId: destinationFolderId }) }, userEmail);
    if (resData?.id || resData?.success) return { success: true, emailId, destinationFolderId };
    return { success: false, error: 'Failed to move email — Graph API returned no confirmation.' };
  }

  async markEmailRead(emailId, isRead = true, userEmail = null) {
    const resData = await this.graphFetch(`/messages/${emailId}`, { method: 'PATCH', body: JSON.stringify({ isRead }) }, userEmail);
    if (resData?.id || resData?.isRead !== undefined || resData?.success) {
      return { success: true, emailId, isRead };
    }
    return { success: false, error: 'Failed to mark email read/unread — Graph API returned no confirmation.' };
  }

  async createDraft({ to, cc, bcc, subject, body, conversationId, attachments }, userEmail = null) {
    const resolvedFiles = resolveAttachmentPaths(attachments);
    const totalBytes = resolvedFiles.reduce((sum, f) => sum + f.size, 0);
    if (totalBytes > MAX_UPLOAD_SESSION_BYTES) {
      return { success: false, error: `Combined attachment size is ${(totalBytes / 1024 / 1024).toFixed(1)}MB — over the ${MAX_UPLOAD_SESSION_BYTES / 1024 / 1024}MB ceiling for this tool.` };
    }
    const useUploadSession = totalBytes > MAX_DIRECT_ATTACH_BYTES;
    const payload = { subject, body: { contentType: 'HTML', content: body }, toRecipients: toRecipients(to), ccRecipients: toRecipients(cc), bccRecipients: toRecipients(bcc) };
    if (resolvedFiles.length && !useUploadSession) payload.attachments = buildInlineAttachments(resolvedFiles);
    const resData = await this.graphFetch(`/messages`, { method: 'POST', body: JSON.stringify(payload) }, userEmail);
    if (!resData?.id) return { success: false, error: 'Failed to create draft in Outlook — Graph API did not return a draft ID. The draft was NOT saved.' };
    let attachmentCount = payload.attachments?.length || 0;
    if (useUploadSession) attachmentCount = await attachFilesToExistingMessage(resData.id, resolvedFiles, userEmail);
    return { draftId: resData.id, to, cc, bcc, subject, body, conversationId, attachmentCount, status: 'created_in_outlook_drafts' };
  }

  async createReplyDraft({ messageId, body = '', replyAll = false, cc, bcc, attachments }, userEmail = null) {
    if (!messageId) throw new Error("Missing required argument: messageId");
    const resolvedFiles = resolveAttachmentPaths(attachments);
    const totalBytes = resolvedFiles.reduce((sum, f) => sum + f.size, 0);
    if (totalBytes > MAX_UPLOAD_SESSION_BYTES) {
      return { success: false, error: `Combined attachment size is ${(totalBytes / 1024 / 1024).toFixed(1)}MB — over the ${MAX_UPLOAD_SESSION_BYTES / 1024 / 1024}MB ceiling for this tool.` };
    }
    const useUploadSession = totalBytes > MAX_DIRECT_ATTACH_BYTES;
    const endpoint = replyAll ? `/messages/${messageId}/createReplyAll` : `/messages/${messageId}/createReply`;
    const payload = body ? { comment: body } : {};
    if (cc || bcc || (resolvedFiles.length && !useUploadSession)) {
      payload.message = { ccRecipients: toRecipients(cc), bccRecipients: toRecipients(bcc) };
      if (resolvedFiles.length && !useUploadSession) payload.message.attachments = buildInlineAttachments(resolvedFiles);
    }
    const resData = await this.graphFetch(endpoint, { method: 'POST', body: JSON.stringify(payload) }, userEmail);
    if (!resData?.id) return { success: false, error: 'Failed to create reply draft in Outlook — Graph API did not return a draft ID.' };
    let attachmentCount = payload.message?.attachments?.length || 0;
    if (useUploadSession) attachmentCount = await attachFilesToExistingMessage(resData.id, resolvedFiles, userEmail);
    return {
      draftId: resData.id,
      messageId,
      conversationId: resData.conversationId || null,
      subject: resData.subject || '',
      toRecipients: (resData.toRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(', '),
      ccRecipients: (resData.ccRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(', '),
      bccRecipients: (resData.bccRecipients || []).map(r => r.emailAddress?.address).filter(Boolean).join(', '),
      attachmentCount,
      status: 'created_reply_draft_in_outlook'
    };
  }

  async updateDraft({ draftId, subject, body, to, cc, bcc }, userEmail = null) {
    if (!draftId) throw new Error("Missing required argument: draftId");
    const payload = {};
    if (subject !== undefined) payload.subject = subject;
    if (body !== undefined) payload.body = { contentType: 'HTML', content: body };
    if (to !== undefined) payload.toRecipients = toRecipients(to);
    if (cc !== undefined) payload.ccRecipients = toRecipients(cc);
    if (bcc !== undefined) payload.bccRecipients = toRecipients(bcc);
    const resData = await this.graphFetch(`/messages/${draftId}`, { method: 'PATCH', body: JSON.stringify(payload) }, userEmail);
    if (resData?.id || resData?.success) return { success: true, draftId, updatedFields: Object.keys(payload) };
    return { success: false, error: 'Failed to update draft in Outlook — Graph API returned no confirmation.' };
  }

  async replyToEmail(emailId, body, replyAll = false, cc = null, bcc = null, attachments = null, userEmail = null) {
    const resolvedFiles = resolveAttachmentPaths(attachments);
    const totalBytes = resolvedFiles.reduce((sum, f) => sum + f.size, 0);
    if (totalBytes > MAX_UPLOAD_SESSION_BYTES) {
      return { success: false, error: `Combined attachment size is ${(totalBytes / 1024 / 1024).toFixed(1)}MB — over the ${MAX_UPLOAD_SESSION_BYTES / 1024 / 1024}MB ceiling for this tool.` };
    }
    const useUploadSession = totalBytes > MAX_DIRECT_ATTACH_BYTES;

    if (useUploadSession) {
      // /reply and /replyAll send immediately and never hand back a message id
      // to attach to — createUploadSession needs a real id. Route through the
      // unsent-draft variant instead: create the reply draft, attach in
      // chunks, then send that draft explicitly.
      const draftResult = await this.createReplyDraft({ messageId: emailId, body, replyAll, cc, bcc, attachments }, userEmail);
      if (!draftResult?.draftId) return { success: false, error: draftResult?.error || 'Failed to create reply draft for large-attachment send.' };
      const sent = await sendExistingDraft(draftResult.draftId, userEmail);
      if (!sent) return { success: false, error: `Reply draft ${draftResult.draftId} was created with ${draftResult.attachmentCount} attachment(s) but sending it failed — it is still sitting in Drafts, not lost.` };
      return { success: true, emailId, replyAll, cc, bcc, attachmentCount: draftResult.attachmentCount, sentAt: new Date().toISOString() };
    }

    const endpoint = replyAll ? `/messages/${emailId}/replyAll` : `/messages/${emailId}/reply`;
    const payload = { comment: body };
    if (cc || bcc || resolvedFiles.length) {
      payload.message = { ccRecipients: toRecipients(cc), bccRecipients: toRecipients(bcc) };
      if (resolvedFiles.length) payload.message.attachments = buildInlineAttachments(resolvedFiles);
    }
    // Graph API returns 202 Accepted (empty body) on success
    const resData = await this.graphFetch(endpoint, { method: 'POST', body: JSON.stringify(payload) }, userEmail);
    if (resData?.success) return { success: true, emailId, replyAll, cc, bcc, attachmentCount: payload.message?.attachments?.length || 0, sentAt: new Date().toISOString() };
    return { success: false, error: 'Failed to send reply — Graph API did not return 202 Accepted.' };
  }

  async sendEmail({ to, cc, bcc, subject, body, attachments }, userEmail = null) {
    const resolvedFiles = resolveAttachmentPaths(attachments);
    const totalBytes = resolvedFiles.reduce((sum, f) => sum + f.size, 0);
    if (totalBytes > MAX_UPLOAD_SESSION_BYTES) {
      return { success: false, error: `Combined attachment size is ${(totalBytes / 1024 / 1024).toFixed(1)}MB — over the ${MAX_UPLOAD_SESSION_BYTES / 1024 / 1024}MB ceiling for this tool.` };
    }
    const useUploadSession = totalBytes > MAX_DIRECT_ATTACH_BYTES;

    if (useUploadSession) {
      // /sendMail is fire-and-forget — no message id comes back, so
      // createUploadSession has nothing to attach to. Route through a real
      // draft instead: create it, attach in chunks, then send that draft.
      const draftResult = await this.createDraft({ to, cc, bcc, subject, body, attachments }, userEmail);
      if (!draftResult?.draftId) return { success: false, error: draftResult?.error || 'Failed to create draft for large-attachment send.' };
      const sent = await sendExistingDraft(draftResult.draftId, userEmail);
      if (!sent) return { success: false, error: `Draft ${draftResult.draftId} was created with ${draftResult.attachmentCount} attachment(s) but sending it failed — it is still sitting in Drafts, not lost.` };
      return { success: true, attachmentCount: draftResult.attachmentCount, sentAt: new Date().toISOString() };
    }

    const payload = { message: { subject, body: { contentType: 'HTML', content: body }, toRecipients: toRecipients(to), ccRecipients: toRecipients(cc), bccRecipients: toRecipients(bcc) } };
    if (resolvedFiles.length) payload.message.attachments = buildInlineAttachments(resolvedFiles);
    const resData = await this.graphFetch(`/sendMail`, { method: 'POST', body: JSON.stringify(payload) }, userEmail);
    // /sendMail returns 202 Accepted with empty body on success; graphFetch returns { success: true }
    if (resData?.success === true) return { success: true, attachmentCount: payload.message.attachments?.length || 0, sentAt: new Date().toISOString() };
    return { success: false, error: 'Failed to send email — Graph API did not return 202 Accepted.' };
  }

  async getCalendarEvents(startTime, endTime, userEmail = null) {
    const data = await this.graphFetch(`/calendarView?startDateTime=${encodeURIComponent(startTime)}&endDateTime=${encodeURIComponent(endTime)}&$select=id,subject,start,end,attendees,location,organizer,bodyPreview,isOnlineMeeting,onlineMeetingUrl&$orderby=start/dateTime`, {}, userEmail);
    if (!data) return { error: 'Failed to fetch calendar events from Microsoft Graph API.' };
    return (data.value || []).map(evt => ({ id: evt.id, subject: evt.subject, start: evt.start?.dateTime, end: evt.end?.dateTime, location: evt.location?.displayName || '', organizer: evt.organizer?.emailAddress?.address || '', attendees: (evt.attendees || []).map(a => a.emailAddress?.address), bodyPreview: evt.bodyPreview || '', isOnlineMeeting: evt.isOnlineMeeting, onlineMeetingUrl: evt.onlineMeetingUrl || '' }));
  }

  async getUpcomingEvents(days = 7, userEmail = null) {
    const start = new Date().toISOString();
    const end = new Date(Date.now() + days * 86400000).toISOString();
    return await this.getCalendarEvents(start, end, userEmail);
  }

  async createEvent({ subject, startTime, endTime, attendees = [], location = 'Microsoft Teams', body = '', force = false }, userEmail = null) {
    if (!force) {
      const existing = await this.getCalendarEvents(startTime, endTime, userEmail);
      if (Array.isArray(existing) && existing.length > 0) {
        return {
          success: false,
          conflict: true,
          conflictingEvents: existing.map(e => ({ id: e.id, subject: e.subject, start: e.start, end: e.end })),
          error: `Conflict: ${existing.length} existing event(s) overlap this time range. Present the conflict to the user before proceeding — pass force:true only after explicit approval to double-book.`
        };
      }
    }
    const payload = { subject, body: { contentType: 'HTML', content: body }, start: { dateTime: startTime, timeZone: 'W. Australia Standard Time' }, end: { dateTime: endTime, timeZone: 'W. Australia Standard Time' }, location: { displayName: location }, attendees: attendees.map(addr => ({ emailAddress: { address: addr }, type: 'required' })) };
    const resData = await this.graphFetch(`/events`, { method: 'POST', body: JSON.stringify(payload) }, userEmail);
    if (resData?.id) return { eventId: resData.id, subject, startTime, endTime, attendees, location, status: 'scheduled' };
    return { success: false, error: 'Failed to create calendar event — Graph API did not return an event ID. The event was NOT created.' };
  }

  async updateEvent(eventId, updates, userEmail = null) {
    if (updates.startTime && updates.endTime && !updates.force) {
      const existing = await this.getCalendarEvents(updates.startTime, updates.endTime, userEmail);
      const conflicts = (Array.isArray(existing) ? existing : []).filter(e => e.id !== eventId);
      if (conflicts.length > 0) {
        return {
          success: false,
          conflict: true,
          conflictingEvents: conflicts.map(e => ({ id: e.id, subject: e.subject, start: e.start, end: e.end })),
          error: `Conflict: ${conflicts.length} existing event(s) overlap the new time range. Present the conflict to the user before proceeding — pass force:true only after explicit approval to double-book.`
        };
      }
    }
    const payload = {};
    if (updates.subject) payload.subject = updates.subject;
    if (updates.startTime) payload.start = { dateTime: updates.startTime, timeZone: 'W. Australia Standard Time' };
    if (updates.endTime) payload.end = { dateTime: updates.endTime, timeZone: 'W. Australia Standard Time' };
    if (updates.location) payload.location = { displayName: updates.location };
    if (updates.body) payload.body = { contentType: 'HTML', content: updates.body };
    const resData = await this.graphFetch(`/events/${eventId}`, { method: 'PATCH', body: JSON.stringify(payload) }, userEmail);
    if (resData?.id || resData?.success) return { success: true, eventId, updates };
    return { success: false, error: 'Failed to update calendar event — Graph API returned no confirmation.' };
  }

  async cancelEvent(eventId, comment = '', userEmail = null) {
    // Graph API returns 202 Accepted (empty body) on success
    const resData = await this.graphFetch(`/events/${eventId}/cancel`, { method: 'POST', body: JSON.stringify({ comment }) }, userEmail);
    if (resData?.success) return { success: true, eventId, cancelledAt: new Date().toISOString() };
    return { success: false, error: 'Failed to cancel calendar event — Graph API did not return 202 Accepted.' };
  }

  async getTodaysBriefing(userEmail = null) {
    const start = new Date(); start.setHours(0, 0, 0, 0);
    const end = new Date(); end.setHours(23, 59, 59, 999);
    const [emails, events] = await Promise.all([this.getUnreadEmails(userEmail), this.getCalendarEvents(start.toISOString(), end.toISOString(), userEmail)]);
    return {
      date: new Date().toISOString(),
      unreadEmails: Array.isArray(emails) ? emails : [],
      todaysEvents: Array.isArray(events) ? events : [],
      ...(emails?.error ? { emailError: emails.error } : {}),
      ...(events?.error ? { eventError: events.error } : {})
    };
  }
}

export const graphClient = new GraphClient();

