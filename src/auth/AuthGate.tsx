import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import type { ReactNode } from 'react';
import { loginRequest } from './msal';

export function AuthGate({ children }: { children: ReactNode }) {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  if (!isAuthenticated) {
    return (
      <div style={{ padding: 24 }}>
        <h1>Outlook Contact Deduper</h1>
        <p>Sign in with your personal Microsoft account to begin.</p>
        <button
          onClick={() =>
            instance.loginPopup(loginRequest).then((r) => instance.setActiveAccount(r.account))
          }
        >
          Sign in
        </button>
      </div>
    );
  }
  return <>{children}</>;
}
