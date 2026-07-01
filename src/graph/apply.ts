import type { MergePlan } from '../engine/merge';
import type { Contact } from '../engine/types';
import { buildBatchSteps, chunk, contactPatchBody, parseRetryAfter, type BatchStep } from './batch';

const GRAPH = 'https://graph.microsoft.com/v1.0';

export interface ApplyOutcome {
  plan: MergePlan;
  ok: boolean;
  error?: string;
}

interface BatchResponse {
  responses: Array<{ id: string; status: number; body?: unknown; headers?: Record<string, string> }>;
}

async function postBatch(token: string, steps: BatchStep[]): Promise<BatchResponse> {
  const res = await fetch(`${GRAPH}/$batch`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests: steps }),
  });
  if (!res.ok) throw new Error(`$batch ${res.status}: ${await res.text()}`);
  return (await res.json()) as BatchResponse;
}

/** Apply each plan's PATCH + DELETEs. Retries throttled (429) sub-requests once, honoring Retry-After. */
export async function applyMergePlans(token: string, plans: MergePlan[]): Promise<ApplyOutcome[]> {
  const outcomes = new Map<string, ApplyOutcome>();
  // Track which batch-step ids belong to which plan so we can roll status up per plan.
  const stepPlan = new Map<string, MergePlan>();
  const allSteps: BatchStep[] = [];
  for (const plan of plans) {
    outcomes.set(plan.survivorId, { plan, ok: true });
    const steps = buildBatchSteps([plan]);
    for (const s of steps) stepPlan.set(s.id, plan);
    allSteps.push(...steps);
  }

  for (const group of chunk(allSteps, 20)) {
    let pending = group;
    for (let attempt = 0; attempt < 2 && pending.length > 0; attempt++) {
      const { responses } = await postBatch(token, pending);
      const retry: BatchStep[] = [];
      let waitSec = 0;
      for (const r of responses) {
        const plan = stepPlan.get(r.id)!;
        if (r.status === 429) {
          waitSec = Math.max(waitSec, parseRetryAfter(r.headers?.['Retry-After'] ?? null));
          const original = pending.find((s) => s.id === r.id)!;
          retry.push(original);
        } else if (r.status >= 400) {
          outcomes.set(plan.survivorId, { plan, ok: false, error: `status ${r.status}` });
        }
      }
      if (retry.length && attempt === 0) {
        await new Promise((res) => setTimeout(res, waitSec * 1000));
        pending = retry;
      } else {
        pending = [];
      }
    }
  }
  return [...outcomes.values()];
}

/** Revert a single applied merge: restore the survivor's prior state and recreate deleted contacts. */
export async function undoMerge(token: string, plan: MergePlan): Promise<void> {
  const patch = await fetch(`${GRAPH}/me/contacts/${plan.survivorId}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(contactPatchBody(plan.original)),
  });
  if (!patch.ok) throw new Error(`undo PATCH ${patch.status}: ${await patch.text()}`);
  // Deleted contacts are recreated (they receive new ids). The full JSON came from the backup/analyze set.
  for (const c of plan.deletedContacts ?? []) {
    const post = await fetch(`${GRAPH}/me/contacts`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(contactPatchBody(c as Contact)),
    });
    if (!post.ok) throw new Error(`undo POST ${post.status}: ${await post.text()}`);
  }
}
