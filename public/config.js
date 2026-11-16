// Portfolio example: replace zero GUIDs with your own development Azure app/tenant IDs.
// Azure AD app registration settings
// clientId and tenantId are non-secret identifiers — safe to commit.
// Never commit ANTHROPIC_API_KEY, HUBSPOT_TOKEN, or other secrets here; those belong in .env
//
// MULTI-USER: Every colleague uses the same config.js — each signs in
// with their own Microsoft account. The add-in reads *their* sent mail
// and builds *their* personal style profile automatically.
//
// DEPLOYMENT: The add-in is served from https://example.invalid.
// The redirect URI (SPA type) in Azure AD must be:
//   https://example.invalid/auth-complete.html
// Upload the manifest.xml via Microsoft 365 Admin Center
// (Settings > Integrated apps) to push the add-in to all users at once.
const _tenantId = "00000000-0000-0000-0000-000000000001";
const msalConfig = {
  auth: {
    clientId:    "00000000-0000-0000-0000-000000000002",
    tenantId:    _tenantId,
    authority:   "https://login.microsoftonline.com/" + _tenantId,
    redirectUri: window.location.origin + "/auth-complete.html"
  }
};

// Required Microsoft Graph scopes
const loginScopes = [
  "Mail.Read",
  "Mail.ReadWrite",
  "Mail.Send",
  "MailboxSettings.Read",
  "User.Read",
  // "People.Read" — needs to be added to the Azure app registration first.
  // Until the admin adds it under App registrations → API permissions → Microsoft Graph,
  // including it here causes AADSTS65005 and breaks the entire sign-in flow.
];
