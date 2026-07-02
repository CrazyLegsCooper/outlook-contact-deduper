import { describe, it, expect } from 'vitest';
import {
  normalizeName, normalizeEmail, normalizePhone,
  contactEmails, contactPhones, contactFullName,
} from '../src/engine/normalize';
import type { Contact } from '../src/engine/types';

describe('normalizeName', () => {
  it('lowercases, trims, collapses whitespace, strips punctuation', () => {
    expect(normalizeName('  Chris   Cooper! ')).toBe('chris cooper');
  });
  it('strips diacritics', () => {
    expect(normalizeName('José')).toBe('jose');
  });
  it('handles null/undefined', () => {
    expect(normalizeName(null)).toBe('');
    expect(normalizeName(undefined)).toBe('');
  });
});

describe('normalizeEmail', () => {
  it('lowercases and trims', () => {
    expect(normalizeEmail('  Chris@X.COM ')).toBe('chris@x.com');
  });
});

describe('normalizePhone', () => {
  it('keeps last 10 digits, dropping formatting and country code', () => {
    expect(normalizePhone('+1 (555) 123-4567')).toBe('5551234567');
  });
  it('returns short numbers as-is', () => {
    expect(normalizePhone('123-456')).toBe('123456');
  });
});

describe('contact aggregates', () => {
  const c: Contact = {
    id: '1',
    displayName: 'Chris Cooper',
    emailAddresses: [{ address: 'Chris@X.com' }, { address: 'chris@x.com' }],
    mobilePhone: '555-123-4567',
    homePhones: ['(555) 123 4567', '555-999-0000'],
  };
  it('de-dupes normalized emails', () => {
    expect(contactEmails(c)).toEqual(['chris@x.com']);
  });
  it('de-dupes normalized phones across fields', () => {
    expect(contactPhones(c)).toEqual(['5551234567', '5559990000']);
  });
  it('derives normalized full name from displayName', () => {
    expect(contactFullName(c)).toBe('chris cooper');
  });
  it('falls back to given + surname', () => {
    expect(contactFullName({ id: '2', givenName: 'John', surname: 'Smith' })).toBe('john smith');
  });
});
