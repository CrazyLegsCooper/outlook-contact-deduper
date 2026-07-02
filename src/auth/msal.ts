import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { appConfig } from '../config';

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: appConfig.clientId,
    authority: appConfig.authority,
    redirectUri: appConfig.redirectUri,
  },
  cache: { cacheLocation: 'sessionStorage' },
});

export const loginRequest = { scopes: appConfig.scopes };

export async function getAccessToken(): Promise<string> {
  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0];
  if (!account) throw new Error('No signed-in account');
  try {
    const res = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return res.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const res = await msalInstance.acquireTokenPopup(loginRequest);
      return res.accessToken;
    }
    throw err;
  }
}
