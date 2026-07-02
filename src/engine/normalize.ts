import type { Contact } from './types';

export function normalizeName(raw?: string | null): string {
  if (!raw) return '';
  return raw
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function normalizeEmail(raw?: string | null): string {
  if (!raw) return '';
  return raw.trim().toLowerCase();
}

export function normalizePhone(raw?: string | null): string {
  if (!raw) return '';
  const digits = raw.replace(/\D/g, '');
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function uniq(values: string[]): string[] {
  return [...new Set(values.filter((v) => v.length > 0))];
}

export function contactEmails(c: Contact): string[] {
  return uniq((c.emailAddresses ?? []).map((e) => normalizeEmail(e.address)));
}

export function contactPhones(c: Contact): string[] {
  const all = [c.mobilePhone, ...(c.homePhones ?? []), ...(c.businessPhones ?? [])];
  return uniq(all.map((p) => normalizePhone(p ?? '')));
}

export function contactFullName(c: Contact): string {
  if (c.displayName && c.displayName.trim()) return normalizeName(c.displayName);
  return normalizeName([c.givenName, c.surname].filter(Boolean).join(' '));
}
