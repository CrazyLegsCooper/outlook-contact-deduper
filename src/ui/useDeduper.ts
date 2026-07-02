import { useState, useCallback } from 'react';
import type { Contact } from '../engine/types';
import { analyze, type AnalyzeResult } from '../engine/analyze';
import { getAccessToken } from '../auth/msal';
import { listAllContacts } from '../graph/contacts';
import { downloadBackup } from '../graph/backup';

export type Phase = 'idle' | 'loading' | 'ready' | 'applying' | 'done';

export function useDeduper() {
  const [phase, setPhase] = useState<Phase>('idle');
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [result, setResult] = useState<AnalyzeResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setPhase('loading');
    setError(null);
    try {
      const token = await getAccessToken();
      const all = await listAllContacts(token);
      downloadBackup(all);
      setContacts(all);
      setResult(analyze(all));
      setPhase('ready');
    } catch (e) {
      setError((e as Error).message);
      setPhase('idle');
    }
  }, []);

  const byId = useCallback((id: string) => contacts.find((c) => c.id === id)!, [contacts]);

  return { phase, setPhase, contacts, setContacts, result, setResult, error, load, byId };
}
