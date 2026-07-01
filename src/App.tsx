import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { getAccessToken } from './auth/msal';

export default function App() {
  const [status, setStatus] = useState('');
  return (
    <AuthGate>
      <div style={{ padding: 24 }}>
        <h1>Outlook Contact Deduper</h1>
        <button
          onClick={async () => {
            try {
              const t = await getAccessToken();
              setStatus(`Got token, length ${t.length}`);
            } catch (e) {
              setStatus(`Error: ${(e as Error).message}`);
            }
          }}
        >
          Get token (debug)
        </button>
        <p>{status}</p>
      </div>
    </AuthGate>
  );
}
