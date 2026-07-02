import { describe, it, expect } from 'vitest';
import { buildPlanSteps, packSteps, removedContactIds } from '../src/graph/apply';
import type { MergePlan } from '../src/engine/merge';

const plan = (survivorId: string, deleteIds: string[]): MergePlan => ({
  survivorId,
  survivor: { id: survivorId, displayName: survivorId },
  deleteIds,
  original: { id: survivorId, displayName: survivorId },
});

describe('buildPlanSteps', () => {
  it('assigns ids unique across all plans in the batch', () => {
    const { steps } = buildPlanSteps([plan('1', ['2', '3']), plan('4', ['5'])]);
    const ids = steps.map((s) => s.id);
    expect(new Set(ids).size).toBe(ids.length); // no collisions across plans
    expect(steps).toHaveLength(5); // 2 PATCH + 3 DELETE
  });

  it('maps each step id back to its owning plan', () => {
    const { steps, stepPlan } = buildPlanSteps([plan('1', ['2']), plan('4', ['5'])]);
    for (const s of steps) {
      const owner = stepPlan.get(s.id)!;
      const contactId = s.url.split('/').pop()!;
      if (s.method === 'PATCH') expect(owner.survivorId).toBe(contactId);
      else expect(owner.deleteIds).toContain(contactId);
    }
    const patchOwners = steps
      .filter((s) => s.method === 'PATCH')
      .map((s) => stepPlan.get(s.id)!.survivorId);
    expect(new Set(patchOwners)).toEqual(new Set(['1', '4']));
  });
});

describe('removedContactIds', () => {
  it('removes deleteIds only from successful outcomes', () => {
    const okPlan = plan('1', ['2', '3']);
    const failPlan = plan('4', ['5']);
    const removed = removedContactIds([
      { plan: okPlan, ok: true },
      { plan: failPlan, ok: false, error: 'status 403' },
    ]);
    expect([...removed].sort()).toEqual(['2', '3']);
    expect(removed.has('5')).toBe(false);
  });
});

describe('packSteps', () => {
  it("keeps each plan's steps within a single group (no straddle)", () => {
    const plans = Array.from({ length: 7 }, (_, i) =>
      plan(String(i * 10), [String(i * 10 + 1), String(i * 10 + 2)]),
    ); // 7 plans x 3 steps = 21 steps -> must span >1 group without splitting a plan
    const { planSteps } = buildPlanSteps(plans);
    const groups = packSteps(planSteps, 20);
    for (const g of groups) expect(g.length).toBeLessThanOrEqual(20);
    for (const ps of planSteps) {
      const ids = new Set(ps.map((s) => s.id));
      const containing = groups.filter((g) => g.some((s) => ids.has(s.id)));
      expect(containing).toHaveLength(1); // whole plan in exactly one group
      expect(containing[0].filter((s) => ids.has(s.id))).toHaveLength(ps.length);
    }
  });

  it('splits a single oversized plan (>size steps) across groups', () => {
    const big = plan('1', Array.from({ length: 25 }, (_, i) => String(100 + i))); // 26 steps
    const { planSteps } = buildPlanSteps([big]);
    const groups = packSteps(planSteps, 20);
    expect(groups.length).toBeGreaterThan(1);
    for (const g of groups) expect(g.length).toBeLessThanOrEqual(20);
  });
});
