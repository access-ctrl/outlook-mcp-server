import { graphClient } from '../outlook/graphClient.js';

/**
 * Fetch content for all file attachments of a given Outlook email (messageId).
 * @param {object} param0
 * @param {string} param0.id - Outlook/Graph message id
 * @returns {Promise<object>} { id, attachments: [ {id, name, contentType, size, contentBytes, isInline} ], count }
 */
export async function getEmailAttachments({ id, mailbox }) {
  if (!id) throw new Error("Missing required argument: id (Outlook message id)");

  // 1. Fetch all attachments metadata for this email
  const metadataList = await graphClient.getAttachmentsForMessage(id, mailbox);

  if (!Array.isArray(metadataList)) {
    return { id, attachments: [], count: 0 };
  }

  // 2. For each attachment, fetch content if file attachment
  const attachments = [];
  for (const meta of metadataList) {
    if (!meta) continue;

    const odataType = meta['@odata.type'] || '';
    const isFileAttachment = !odataType || odataType.includes('fileAttachment') || Boolean(meta.contentBytes);

    if (isFileAttachment) {
      let content = meta.contentBytes;
      if (!content && meta.id) {
        try {
          const full = await graphClient.getAttachmentContent({ messageId: id, attachmentId: meta.id }, mailbox);
          content = full?.contentBytes || null;
        } catch (err) {
          console.error(`[getEmailAttachments] Failed to fetch content for attachment ${meta.id}:`, err.message);
        }
      }

      attachments.push({
        id: meta.id || '',
        name: meta.name || 'unnamed_attachment',
        contentType: meta.contentType || 'application/octet-stream',
        size: meta.size || 0,
        isInline: Boolean(meta.isInline),
        contentBytes: content || null
      });
    } else {
      attachments.push({
        id: meta.id || '',
        name: meta.name || 'item_attachment',
        contentType: odataType || 'itemAttachment',
        size: meta.size || 0,
        isInline: Boolean(meta.isInline),
        contentBytes: null
      });
    }
  }

  return { id, attachments, count: attachments.length };
}

