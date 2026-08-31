    import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
    import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
    import { z } from 'zod';
    import { graphClient } from '../outlook/graphClient.js';
    import { triageEmail } from '../tools/emailTool.js';
    import { getEmailAttachments } from '../tools/getAttachments.js';

    const server = new McpServer({ name: 'OutlookMCPServer', version: '2.2.0' });

    const mailboxSchema = z.string().optional().describe('Target Outlook email address / mailbox. Defaults to primary configured user email if omitted.');
    const ccSchema = z.string().optional().describe('Comma-separated CC email address(es)');
    const bccSchema = z.string().optional().describe('Comma-separated BCC email address(es)');
    const attachmentsSchema = z.array(z.string()).optional().describe('Absolute file paths under /opt/data to attach. Up to 3MB combined attaches instantly; over 3MB (up to 25MB combined) is uploaded in chunks automatically — no difference in how you call this, just slower for the larger case. Over 25MB combined is rejected.');

    server.tool('get_recent_emails', 'Fetch recently received emails from Outlook inbox',
    { limit: z.number().optional().describe('Max emails (default: 10)'), mailbox: mailboxSchema },
    async ({ limit = 10, mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getRecentEmails(limit, mailbox), null, 2) }] }));

    server.tool('get_recent_emails_all_folders', 'Fetch recently received emails across the entire mailbox (all folders, not just Inbox), excluding Deleted Items (and its subfolders), Junk, Junk Emails, and Drafts. Returns bodyPreview (short text snippet) instead of full HTML body to keep payload small.',
    { limit: z.number().optional().describe('Max emails (default: 10)'), mailbox: mailboxSchema },
    async ({ limit = 10, mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getRecentEmailsAllFolders(limit, mailbox), null, 2) }] }));

    server.tool('get_unread_emails', 'Fetch all unread emails from Outlook inbox for triage',
    { mailbox: mailboxSchema },
    async ({ mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getUnreadEmails(mailbox), null, 2) }] }));

    server.tool('get_conversation_thread', 'Retrieve full email conversation history by conversation ID',
    { conversationId: z.string().describe('Outlook Conversation ID'), limit: z.number().optional(), mailbox: mailboxSchema },
    async ({ conversationId, limit = 10, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getConversationThread(conversationId, limit, mailbox), null, 2) }] }));

    server.tool('search_emails', 'Search emails by keyword, sender, subject or topic',
    { query: z.string().describe('Search query'), limit: z.number().optional(), mailbox: mailboxSchema },
    async ({ query, limit = 20, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.searchEmails(query, 'inbox', limit, mailbox), null, 2) }] }));

    server.tool('get_sent_emails', 'Fetch sent emails to track outbox and detect missing replies',
    { limit: z.number().optional(), mailbox: mailboxSchema },
    async ({ limit = 25, mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getSentEmails(limit, mailbox), null, 2) }] }));

    server.tool('get_email_folders', 'List all Outlook mail folders with unread counts',
    { mailbox: mailboxSchema },
    async ({ mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getEmailFolders(mailbox), null, 2) }] }));

    server.tool('triage_email', 'Classify an email for urgency, category and suggest a reply',
    { id: z.string(), sender: z.string(), senderName: z.string().optional(), subject: z.string(), body: z.string() },
    async (emailData) => ({ content: [{ type: 'text', text: JSON.stringify(await triageEmail(emailData), null, 2) }] }));

    server.tool('create_email_draft', 'Create a standalone draft email in Outlook Drafts folder for review before sending. Supports file attachments up to 25MB combined.',
    { to: z.string(), cc: ccSchema, bcc: bccSchema, subject: z.string(), body: z.string(), conversationId: z.string().optional(), attachments: attachmentsSchema, mailbox: mailboxSchema },
    async ({ to, cc, bcc, subject, body, conversationId, attachments, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.createDraft({ to, cc, bcc, subject, body, conversationId, attachments }, mailbox), null, 2) }] }));

    server.tool('create_reply_draft', 'Create an unsent draft reply attached to an existing email thread using Microsoft Graph createReply API (POST /messages/{id}/createReply). Supports file attachments up to 25MB combined.',
    { messageId: z.string().describe('Outlook message ID of the email to reply to'), body: z.string().optional().describe('Text/HTML reply content to populate in the draft'), replyAll: z.boolean().optional().describe('Reply to all recipients (default: false)'), cc: ccSchema, bcc: bccSchema, attachments: attachmentsSchema, mailbox: mailboxSchema },
    async ({ messageId, body = '', replyAll = false, cc, bcc, attachments, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.createReplyDraft({ messageId, body, replyAll, cc, bcc, attachments }, mailbox), null, 2) }] }));

    server.tool('update_email_draft', 'Update/edit an existing email draft subject, body, or recipient before sending (PATCH /messages/{draftId})',
    { draftId: z.string().describe('Outlook draft message ID'), body: z.string().optional().describe('Updated body content (HTML)'), subject: z.string().optional().describe('Updated subject line'), to: z.string().optional().describe('Updated recipient email address'), cc: ccSchema, bcc: bccSchema, mailbox: mailboxSchema },
    async ({ draftId, body, subject, to, cc, bcc, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.updateDraft({ draftId, subject, body, to, cc, bcc }, mailbox), null, 2) }] }));

    server.tool('reply_to_email', 'Reply to an existing email thread. Requires prior approval. Supports file attachments up to 25MB combined (over 3MB sends via an internal draft-then-send step — same result, no different usage).',
    { emailId: z.string(), body: z.string(), replyAll: z.boolean().optional(), cc: ccSchema, bcc: bccSchema, attachments: attachmentsSchema, mailbox: mailboxSchema },
    async ({ emailId, body, replyAll = false, cc, bcc, attachments, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.replyToEmail(emailId, body, replyAll, cc, bcc, attachments, mailbox), null, 2) }] }));

    server.tool('send_email', 'Send an approved email. Requires explicit approval from the client first. Supports file attachments up to 25MB combined (over 3MB sends via an internal draft-then-send step — same result, no different usage).',
    { to: z.string(), cc: ccSchema, bcc: bccSchema, subject: z.string(), body: z.string(), attachments: attachmentsSchema, mailbox: mailboxSchema },
    async ({ to, cc, bcc, subject, body, attachments, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.sendEmail({ to, cc, bcc, subject, body, attachments }, mailbox), null, 2) }] }));

    server.tool('move_email', 'Move an email to a specific folder',
    { emailId: z.string(), destinationFolderId: z.string(), mailbox: mailboxSchema },
    async ({ emailId, destinationFolderId, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.moveEmail(emailId, destinationFolderId, mailbox), null, 2) }] }));

    server.tool('get_email_attachments', 'Fetches the content of all the file attachments for a given Outlook messageId. Returns array [{id, name, contentType, size, isInline, contentBytes}].',
    { id: z.string().describe('Outlook/Graph message ID'), mailbox: mailboxSchema },
    async ({ id, mailbox }) => {
        try {
        const data = await getEmailAttachments({ id, mailbox });
        return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
        } catch (err) {
        return { content: [{ type: 'text', text: JSON.stringify({ error: err.message || String(err) }, null, 2) }], isError: true };
        }
    });

    server.tool('mark_email_read', 'Mark an email as read or unread',
    { emailId: z.string(), isRead: z.boolean().optional(), mailbox: mailboxSchema },
    async ({ emailId, isRead = true, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.markEmailRead(emailId, isRead, mailbox), null, 2) }] }));

    server.tool('get_upcoming_events', 'Get upcoming calendar events for the next N days',
    { days: z.number().optional(), mailbox: mailboxSchema },
    async ({ days = 7, mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getUpcomingEvents(days, mailbox), null, 2) }] }));

    server.tool('get_calendar_events', 'Get calendar events between two datetimes',
    { startTime: z.string().describe('ISO 8601 start'), endTime: z.string().describe('ISO 8601 end'), mailbox: mailboxSchema },
    async ({ startTime, endTime, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getCalendarEvents(startTime, endTime, mailbox), null, 2) }] }));

    server.tool('create_calendar_event', 'Create a new calendar event. Requires approval. Automatically checks for conflicting events in that time range first and refuses (returning conflict:true and the conflicting events) unless force:true is passed — only pass force:true after the user has explicitly approved double-booking.',
    { subject: z.string(), startTime: z.string(), endTime: z.string(), attendees: z.array(z.string()).optional(), location: z.string().optional(), body: z.string().optional(), force: z.boolean().optional().describe('Skip the conflict check and create anyway. Only set true after explicit user approval to double-book.'), mailbox: mailboxSchema },
    async ({ subject, startTime, endTime, attendees = [], location = 'Microsoft Teams', body = '', force = false, mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.createEvent({ subject, startTime, endTime, attendees, location, body, force }, mailbox), null, 2) }] }));

    server.tool('update_calendar_event', 'Update an existing calendar event. If both startTime and endTime are being changed, automatically checks for conflicting events in the new time range first and refuses (returning conflict:true and the conflicting events) unless force:true is passed — only pass force:true after the user has explicitly approved double-booking.',
    { eventId: z.string(), subject: z.string().optional(), startTime: z.string().optional(), endTime: z.string().optional(), location: z.string().optional(), body: z.string().optional(), force: z.boolean().optional().describe('Skip the conflict check and update anyway. Only set true after explicit user approval to double-book.'), mailbox: mailboxSchema },
    async ({ eventId, mailbox, ...updates }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.updateEvent(eventId, updates, mailbox), null, 2) }] }));

    server.tool('cancel_calendar_event', 'Cancel a calendar event and notify attendees',
    { eventId: z.string(), comment: z.string().optional(), mailbox: mailboxSchema },
    async ({ eventId, comment = '', mailbox }) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.cancelEvent(eventId, comment, mailbox), null, 2) }] }));

    server.tool('get_todays_briefing', "Get today's unread emails and calendar events in one call",
    { mailbox: mailboxSchema },
    async ({ mailbox } = {}) => ({ content: [{ type: 'text', text: JSON.stringify(await graphClient.getTodaysBriefing(mailbox), null, 2) }] }));

    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error('[OutlookMCPServer] v2.2.0 started — multi-mailbox support; attachments now up to 25MB combined (3-25MB auto-chunked via createUploadSession, was hard-capped at 3MB)');
