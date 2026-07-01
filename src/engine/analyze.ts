import type { Contact } from './types';
import { findCandidatePairs, type Bucket, type CandidatePair } from './match';
import { clusterPairs } from './cluster';

export interface Group {
  ids: string[];
  bucket: Bucket;
  reasons: string[];
}

export interface AnalyzeResult {
  totalContacts: number;
  veryLikely: Group[];
  notSure: Group[];
}

function reasonsFor(ids: string[], pairs: CandidatePair[]): string[] {
  const idSet = new Set(ids);
  const reasons = new Set<string>();
  for (const p of pairs) {
    if (idSet.has(p.aId) && idSet.has(p.bId)) p.reasons.forEach((r) => reasons.add(r));
  }
  return [...reasons];
}

export function analyze(contacts: Contact[]): AnalyzeResult {
  const pairs = findCandidatePairs(contacts);
  const vlPairs = pairs.filter((p) => p.bucket === 'very-likely');
  const nsPairs = pairs.filter((p) => p.bucket === 'not-sure');

  const veryLikely: Group[] = clusterPairs(vlPairs).map((ids) => ({
    ids,
    bucket: 'very-likely' as const,
    reasons: reasonsFor(ids, vlPairs),
  }));

  const claimed = new Set(veryLikely.flatMap((g) => g.ids));

  const notSure: Group[] = clusterPairs(nsPairs)
    .map((ids) => ids.filter((id) => !claimed.has(id)))
    .filter((ids) => ids.length >= 2)
    .map((ids) => ({ ids, bucket: 'not-sure' as const, reasons: reasonsFor(ids, nsPairs) }));

  return { totalContacts: contacts.length, veryLikely, notSure };
}
