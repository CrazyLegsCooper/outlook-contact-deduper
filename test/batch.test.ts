import { describe, it, expect } from 'vitest';
import { chunk, contactPatchBody, parseRetryAfter, buildBatchSteps } from '../src/graph/batch';
import type { MergePlan } from '../src/engine/merge';
import type { Contact } from '../src/engine/types';

describe('chunk', () => {
  it('splits into groups of at most size', () => {
    expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
  });
});

describe('contactPatchBody', () => {
  it('strips read-only fields', () => {
    const c: Contact = { id: '1', displayName: 'A', parentFolderId: 'p', lastModifiedDateTime: 't', jobTitle: 'X' };
    const body = contactPatchBody(c);
    expect(body).not.toHaveProperty('id');
    expect(body).not.toHaveProperty('parentFolderId');
    expect(body).not.toHaveProperty('lastModifiedDateTime');
    expect(body.jobTitle).toBe('X');
  });
});

describe('parseRetryAfter', () => {
  it('parses seconds', () => {
    expect(parseRetryAfter('12')).toBe(12);
  });
  it('defaults when missing', () => {
    expect(parseRetryAfter(null)).toBe(5);
  });
});

describe('buildBatchSteps', () => {
  it('emits a PATCH for the survivor and a DELETE per removed id', () => {
    const plan: MergePlan = {
      survivorId: '1',
      survivor: { id: '1', displayName: 'A', jobTitle: 'X' },
      deleteIds: ['2', '3'],
      original: { id: '1', displayName: 'A' },
    };
    const steps = buildBatchSteps([plan]);
    const methods = steps.map((s) => s.method);
    expect(methods.filter((m) => m === 'PATCH')).toHaveLength(1);
    expect(methods.filter((m) => m === 'DELETE')).toHaveLength(2);
    expect(steps.find((s) => s.method === 'PATCH')!.url).toBe('/me/contacts/1');
    expect(steps.filter((s) => s.method === 'DELETE').map((s) => s.url).sort())
      .toEqual(['/me/contacts/2', '/me/contacts/3']);
    // ids must be unique across the batch
    expect(new Set(steps.map((s) => s.id)).size).toBe(steps.length);
  });
});
