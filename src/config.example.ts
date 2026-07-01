// Copy this file to src/config.ts and paste your Entra Application (client) ID.
// src/config.ts is git-ignored so your ID is never committed.
export const appConfig = {
  clientId: 'YOUR_ENTRA_APPLICATION_CLIENT_ID',
  redirectUri: 'http://localhost:5173',
  authority: 'https://login.microsoftonline.com/consumers',
  scopes: ['User.Read', 'Contacts.ReadWrite'],
};

export type AppConfig = typeof appConfig;
