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

  const [appliedOutcomes, setAppliedOutcomes] = useState<import('../graph/apply').ApplyOutcome[]>([]);

  const applyPlans = useCallback(async (plans: import('../engine/merge').MergePlan[]) => {
    setPhase('applying');
    const token = await getAccessToken();
    const { applyMergePlans } = await import('../graph/apply');
    const outcomes = await applyMergePlans(token, plans);
    const removed = new Set(plans.flatMap((p) => p.deleteIds));
    setContacts((prev) => {
      const next = prev.filter((c) => !removed.has(c.id));
      setResult(analyze(next));
      return next;
    });
    setAppliedOutcomes((prev) => [...prev, ...outcomes]);
    setPhase('ready');
  }, []);

  const undoLast = useCallback(async () => {
    const last = appliedOutcomes[appliedOutcomes.length - 1];
    if (!last) return;
    const token = await getAccessToken();
    const { undoMerge } = await import('../graph/apply');
    await undoMerge(token, last.plan);
    setAppliedOutcomes((prev) => prev.slice(0, -1));
    // Reload is the simplest correct way to resync ids after a recreate.
  }, [appliedOutcomes]);

  return {
    phase, setPhase, contacts, setContacts, result, setResult, error, load, byId,
    applyPlans, appliedOutcomes, setAppliedOutcomes, undoLast,
  };
}
