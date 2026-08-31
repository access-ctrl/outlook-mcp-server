// Optional: comma-separated Graph folder IDs to exclude from
// getRecentEmailsAllFolders (e.g. deployment-specific "Deleted Items"
// subfolders). Unset/empty means no folders are skipped.
function parseExcludedFolderIds(raw) {
  if (!raw) return [];
  return raw.split(',').map(id => id.trim()).filter(Boolean);
}

export const config = {
  azure: {
    tenantId: process.env.AZURE_TENANT_ID || '',
    clientId: process.env.AZURE_CLIENT_ID || '',
    clientSecret: process.env.AZURE_CLIENT_SECRET || '',
    userEmail: process.env.OUTLOOK_USER_EMAIL || ''
  },
  excludedFolderIds: parseExcludedFolderIds(process.env.OUTLOOK_EXCLUDED_FOLDER_IDS),
  // Master on/off switch for folder skipping in getRecentEmailsAllFolders.
  // Defaults to false — filtering only happens when explicitly enabled.
  skipExcludedFolders: process.env.OUTLOOK_SKIP_EXCLUDED_FOLDERS === 'true',
  // Sign-off name used in triageEmail's suggested-reply templates.
  // Defaults to a generic placeholder so the package isn't tied to one person.
  signoffName: process.env.OUTLOOK_SIGNOFF_NAME || 'The Team'
};
