import { describe, it, expect } from 'vitest';
import { analyze } from '../src/engine/analyze';
import type { Contact } from '../src/engine/types';

const c = (id: string, extra: Partial<Contact>): Contact => ({ id, ...extra });

describe('analyze', () => {
  it('groups three copies of the same person into one very-likely group', () => {
    const result = analyze([
      c('1', { displayName: 'Mum', mobilePhone: '5551112222' }),
      c('2', { displayName: 'Mum', homePhones: ['555-111-2222'] }),
      c('3', { displayName: 'Mum', businessPhones: ['(555) 111 2222'] }),
    ]);
    expect(result.totalContacts).toBe(3);
    expect(result.veryLikely).toHaveLength(1);
    expect([...result.veryLikely[0].ids].sort()).toEqual(['1', '2', '3']);
    expect(result.notSure).toHaveLength(0);
  });

  it('separates a not-sure group from very-likely', () => {
    const result = analyze([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'a@x.com' }] }),
      c('2', { displayName: 'John Smith', emailAddresses: [{ address: 'b@y.com' }] }),
    ]);
    expect(result.veryLikely).toHaveLength(0);
    expect(result.notSure).toHaveLength(1);
  });

  it('excludes a very-likely contact from not-sure groups', () => {
    // 1 & 2 share an email (very-likely). 2 & 3 share only a name (not-sure).
    const result = analyze([
      c('1', { displayName: 'Ann Lee', emailAddresses: [{ address: 'ann@x.com' }] }),
      c('2', { displayName: 'Ann Lee', emailAddresses: [{ address: 'ann@x.com' }] }),
      c('3', { displayName: 'Ann Lee', emailAddresses: [{ address: 'other@z.com' }] }),
    ]);
    expect(result.veryLikely).toHaveLength(1);
    // id 3 pairs by name with 1 and 2, but those are already in very-likely,
    // so no not-sure group should reference them.
    for (const g of result.notSure) {
      expect(g.ids).not.toContain('1');
      expect(g.ids).not.toContain('2');
    }
  });
});
