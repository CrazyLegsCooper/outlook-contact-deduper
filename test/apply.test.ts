import { describe, it, expect } from 'vitest';
import { buildPlanSteps } from '../src/graph/apply';
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
