import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { getAccessToken } from './auth/msal';
import { listAllContacts } from './graph/contacts';
import { downloadBackup } from './graph/backup';

export default function App() {
  const [count, setCount] = useState<number | null>(null);
  return (
    <AuthGate>
      <div style={{ padding: 24 }}>
        <h1>Outlook Contact Deduper</h1>
        <button
          onClick={async () => {
            const token = await getAccessToken();
            const contacts = await listAllContacts(token);
            downloadBackup(contacts);
            setCount(contacts.length);
          }}
        >
          Load contacts + backup
        </button>
        {count !== null && <p>Loaded {count} contacts; backup downloaded.</p>}
      </div>
    </AuthGate>
  );
}
