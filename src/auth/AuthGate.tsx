import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import type { ReactNode } from 'react';
import { loginRequest } from './msal';
import { Button } from '../ui/components/Button';

export function AuthGate({ children }: { children: ReactNode }) {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  if (!isAuthenticated) {
    return (
      <div className="mx-auto max-w-3xl p-6">
        <h1 className="text-3xl font-bold text-primary">Outlook Contact Deduper</h1>
        <p className="mt-1 text-muted-fg">Sign in with your personal Microsoft account to begin.</p>
        <Button className="mt-4"
          onClick={() => instance.loginPopup(loginRequest).then((r) => instance.setActiveAccount(r.account))}>
          Sign in
        </Button>
      </div>
    );
  }
  return <>{children}</>;
}
