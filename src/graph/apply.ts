import type { MergePlan } from '../engine/merge';
import type { Contact } from '../engine/types';
import { buildBatchSteps, chunk, contactPatchBody, parseRetryAfter, type BatchStep } from './batch';

const GRAPH = 'https://graph.microsoft.com/v1.0';

export interface ApplyOutcome {
  plan: MergePlan;
  ok: boolean;
  error?: string;
}

/** Contact ids to remove from local state after applying: only the deleted ids
 *  of plans whose merge actually succeeded (ok). Failed merges leave their
 *  contacts in place so the duplicate stays visible for retry. */
export function removedContactIds(outcomes: ApplyOutcome[]): Set<string> {
  return new Set(outcomes.filter((o) => o.ok).flatMap((o) => o.plan.deleteIds));
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

/**
 * Build batch steps for multiple plans with ids unique across the ENTIRE batch,
 * plus a map from each step id back to its owning plan. buildBatchSteps restarts
 * its counter per call, so we renumber here to avoid cross-plan id collisions.
 */
export function buildPlanSteps(plans: MergePlan[]): {
  steps: BatchStep[];
  stepPlan: Map<string, MergePlan>;
  planSteps: BatchStep[][];
} {
  const planSteps: BatchStep[][] = [];
  const stepPlan = new Map<string, MergePlan>();
  let n = 0;
  for (const plan of plans) {
    const group: BatchStep[] = [];
    for (const s of buildBatchSteps([plan])) {
      s.id = String(++n);
      stepPlan.set(s.id, plan);
      group.push(s);
    }
    planSteps.push(group);
  }
  return { steps: planSteps.flat(), stepPlan, planSteps };
}

/** Pack whole plans' step-groups into $batch chunks of at most `size` ops so a
 *  single plan never straddles a boundary. A lone plan larger than `size`
 *  (a cluster with >size contacts) is unavoidably split, but never dragged in
 *  with others. */
export function packSteps(planSteps: BatchStep[][], size = 20): BatchStep[][] {
  const groups: BatchStep[][] = [];
  let current: BatchStep[] = [];
  for (const ps of planSteps) {
    if (ps.length > size) {
      if (current.length) { groups.push(current); current = []; }
      for (const c of chunk(ps, size)) groups.push(c);
      continue;
    }
    if (current.length + ps.length > size) { groups.push(current); current = []; }
    current.push(...ps);
  }
  if (current.length) groups.push(current);
  return groups;
}

/** Apply each plan's PATCH + DELETEs. Retries throttled (429) sub-requests once, honoring Retry-After. */
export async function applyMergePlans(token: string, plans: MergePlan[]): Promise<ApplyOutcome[]> {
  const outcomes = new Map<string, ApplyOutcome>();
  for (const plan of plans) outcomes.set(plan.survivorId, { plan, ok: true });
  // Step ids must be unique across the whole batch (Graph requires it); buildPlanSteps
  // renumbers across plans and returns the id -> plan map for status roll-up.
  const { stepPlan, planSteps } = buildPlanSteps(plans);

  for (const group of packSteps(planSteps)) {
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
