import { describe, it, expect } from 'vitest';
import { mergeContacts, chooseSurvivor } from '../src/engine/merge';
import type { Contact } from '../src/engine/types';

describe('chooseSurvivor', () => {
  it('picks the most complete contact', () => {
    const a: Contact = { id: '1', displayName: 'Chris' };
    const b: Contact = {
      id: '2', displayName: 'Chris Cooper', givenName: 'Chris', surname: 'Cooper',
      emailAddresses: [{ address: 'c@x.com' }], mobilePhone: '5551234567',
    };
    expect(chooseSurvivor([a, b]).id).toBe('2');
  });
});

describe('mergeContacts', () => {
  it('unions emails and fills empty fields, deleting the others', () => {
    const a: Contact = {
      id: '1', displayName: 'Christopher Cooper', givenName: 'Christopher', surname: 'Cooper',
      emailAddresses: [{ address: 'chris@x.com' }], mobilePhone: '555-123-4567',
    };
    const b: Contact = {
      id: '2', displayName: 'Chris Cooper',
      emailAddresses: [{ address: 'chris.cooper@work.com' }], businessPhones: ['555-999-0000'],
      jobTitle: 'Engineer',
    };
    const plan = mergeContacts([a, b]);
    expect(plan.survivorId).toBe('1'); // a is more complete
    expect(plan.deleteIds).toEqual(['2']);
    expect(plan.survivor.emailAddresses!.map((e) => e.address).sort())
      .toEqual(['chris.cooper@work.com', 'chris@x.com']);
    expect(plan.survivor.mobilePhone).toBe('555-123-4567');
    expect(plan.survivor.businessPhones).toEqual(['555-999-0000']);
    expect(plan.survivor.jobTitle).toBe('Engineer'); // filled from b
    expect(plan.original.id).toBe('1'); // pre-merge snapshot
  });

  it('does not duplicate a phone that appears in two categories', () => {
    const a: Contact = { id: '1', displayName: 'Mum', mobilePhone: '5551112222' };
    const b: Contact = { id: '2', displayName: 'Mum', homePhones: ['(555) 111-2222'] };
    const plan = mergeContacts([a, b]);
    expect(plan.survivor.mobilePhone).toBe('5551112222');
    expect(plan.survivor.homePhones ?? []).toEqual([]); // already covered by mobile
  });

  it('respects an explicit survivorId', () => {
    const a: Contact = { id: '1', displayName: 'A', jobTitle: 'X' };
    const b: Contact = { id: '2', displayName: 'B' };
    const plan = mergeContacts([a, b], '2');
    expect(plan.survivorId).toBe('2');
    expect(plan.deleteIds).toEqual(['1']);
    expect(plan.survivor.jobTitle).toBe('X'); // still filled from a
  });

  it('appends differing personal notes', () => {
    const a: Contact = { id: '1', displayName: 'A', personalNotes: 'note one' };
    const b: Contact = { id: '2', displayName: 'A', personalNotes: 'note two' };
    const plan = mergeContacts([a, b]);
    expect(plan.survivor.personalNotes).toContain('note one');
    expect(plan.survivor.personalNotes).toContain('note two');
  });
});
