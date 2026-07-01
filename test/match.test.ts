import { describe, it, expect } from 'vitest';
import { findCandidatePairs } from '../src/engine/match';
import type { Contact } from '../src/engine/types';

const c = (id: string, extra: Partial<Contact>): Contact => ({ id, ...extra });

describe('findCandidatePairs', () => {
  it('flags shared email as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Chris Cooper', emailAddresses: [{ address: 'chris@x.com' }] }),
      c('2', { displayName: 'Christopher Cooper', emailAddresses: [{ address: 'CHRIS@x.com' }] }),
    ]);
    expect(pairs).toHaveLength(1);
    expect(pairs[0].bucket).toBe('very-likely');
    expect(pairs[0].reasons.join(' ')).toMatch(/email/i);
  });

  it('flags exact name + shared phone as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Mum', mobilePhone: '555-111-2222' }),
      c('2', { displayName: 'Mum', homePhones: ['(555) 111 2222'] }),
    ]);
    expect(pairs[0].bucket).toBe('very-likely');
  });

  it('flags fuzzy name + shared phone as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Chris Cooper', mobilePhone: '5551234567' }),
      c('2', { displayName: 'Christopher Cooper', mobilePhone: '555-123-4567' }),
    ]);
    expect(pairs[0].bucket).toBe('very-likely');
  });

  it('flags same name with no shared contact info as not-sure', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'john1@x.com' }] }),
      c('2', { displayName: 'John Smith', emailAddresses: [{ address: 'john2@y.com' }] }),
    ]);
    expect(pairs).toHaveLength(1);
    expect(pairs[0].bucket).toBe('not-sure');
  });

  it('emits no pair for unrelated contacts', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'john@x.com' }] }),
      c('2', { displayName: 'Peter Baker', emailAddresses: [{ address: 'peter@y.com' }] }),
    ]);
    expect(pairs).toHaveLength(0);
  });
});
