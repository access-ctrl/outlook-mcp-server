import { graphClient } from '../outlook/graphClient.js';
import { config } from '../config/env.js';

export async function createEmailDraft({ to, subject, body }) {
  return await graphClient.createDraft({ to, subject, body });
}

export async function triageEmail(emailData) {
  const text = (emailData.subject + ' ' + emailData.body).toLowerCase();
  const isMeetingRequest = /meeting|schedule|call|discuss|zoom|teams|catch up|availability/i.test(text);
  const isUrgent = /urgent|asap|immediately|critical|emergency|deadline|today|by end of/i.test(text);
  const isDocReview = /doc|guide|api|review|documentation|feedback|draft|contract|agreement/i.test(text);
  const isAddressRequest = /address|shipping|order|delivery|postal|zip/i.test(text);
  const requiresReply = /please|kindly|reply|let me know|confirm|update|review|could you|would you/i.test(text);
  const isIntroduction = /introduce|introduction|pleasure|reach out|connect|opportunity/i.test(text);
  const isFYI = /fyi|for your information|no action|just to let you know/i.test(text);

  let category = 'General', suggestedAction = 'Review & Draft Response', urgency = 'Normal';
  if (isUrgent) urgency = 'High';
  if (isMeetingRequest) { category = 'Meeting Request'; suggestedAction = 'Check Calendar & Propose Time'; }
  else if (isDocReview) { category = 'Document Review'; suggestedAction = 'Review & Provide Feedback'; }
  else if (isAddressRequest) { category = 'Information Request'; suggestedAction = 'Provide Requested Information'; }
  else if (isIntroduction) { category = 'Introduction'; suggestedAction = 'Draft Acknowledgement'; }
  else if (isFYI) { category = 'FYI'; suggestedAction = 'No Reply Needed'; urgency = 'Low'; }

  const name = emailData.senderName || 'there';
  const signoff = config.signoffName;
  let suggestedReply = '';
  if (isMeetingRequest) suggestedReply = `Hi ${name},\n\nThank you for reaching out. I'm checking availability and will come back to you shortly with a proposed time.\n\nBest regards,\n${signoff}`;
  else if (isDocReview) suggestedReply = `Hi ${name},\n\nThank you for sharing this. I'll review the material and get back to you.\n\nBest regards,\n${signoff}`;
  else if (isUrgent) suggestedReply = `Hi ${name},\n\nThank you — I've received your message and will attend to this shortly.\n\nBest regards,\n${signoff}`;
  else if (!isFYI) suggestedReply = `Hi ${name},\n\nThank you for your message. I'll review and get back to you.\n\nBest regards,\n${signoff}`;

  return { id: emailData.id, subject: emailData.subject, sender: emailData.sender, senderName: emailData.senderName, category, urgency, isMeetingRequest, isUrgent, requiresReply, suggestedAction, suggestedReply };
}
