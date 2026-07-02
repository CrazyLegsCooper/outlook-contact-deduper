import type { Contact } from '../engine/types';
import type { MergePlan } from '../engine/merge';

const READ_ONLY: Array<keyof Contact> = ['id', 'parentFolderId', 'lastModifiedDateTime'];

export interface BatchStep {
  id: string;
  method: 'PATCH' | 'DELETE' | 'POST';
  url: string;
  body?: unknown;
  headers?: Record<string, string>;
}

export function chunk<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}

export function contactPatchBody(c: Contact): Partial<Contact> {
  const body: Partial<Contact> = { ...c };
  for (const f of READ_ONLY) delete (body as Record<string, unknown>)[f];
  return body;
}

export function parseRetryAfter(header: string | null): number {
  if (!header) return 5;
  const n = parseInt(header, 10);
  return Number.isFinite(n) && n >= 1 ? n : 5;
}

export function buildBatchSteps(plans: MergePlan[]): BatchStep[] {
  const steps: BatchStep[] = [];
  let n = 0;
  for (const plan of plans) {
    steps.push({
      id: String(++n),
      method: 'PATCH',
      url: `/me/contacts/${plan.survivorId}`,
      body: contactPatchBody(plan.survivor),
      headers: { 'Content-Type': 'application/json' },
    });
    for (const del of plan.deleteIds) {
      steps.push({ id: String(++n), method: 'DELETE', url: `/me/contacts/${del}` });
    }
  }
  return steps;
}
