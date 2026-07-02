import type { Contact, EmailAddress, PhysicalAddress } from './types';
import { normalizeEmail, normalizePhone } from './normalize';

export interface MergePlan {
  survivorId: string;
  survivor: Contact;
  deleteIds: string[];
  original: Contact;
  deletedContacts?: Contact[];
}

const SCALAR_FIELDS: Array<keyof Contact> = [
  'displayName', 'givenName', 'surname', 'nickName',
  'companyName', 'jobTitle', 'department', 'birthday',
];

function addressHasData(a?: PhysicalAddress): boolean {
  return !!a && Object.values(a).some((v) => v && String(v).trim().length > 0);
}

export function completenessScore(c: Contact): number {
  let score = 0;
  for (const f of SCALAR_FIELDS) if (c[f] && String(c[f]).trim().length > 0) score += 1;
  score += (c.emailAddresses ?? []).length;
  score += (c.mobilePhone ? 1 : 0) + (c.homePhones ?? []).length + (c.businessPhones ?? []).length;
  if (addressHasData(c.homeAddress)) score += 1;
  if (addressHasData(c.businessAddress)) score += 1;
  if (addressHasData(c.otherAddress)) score += 1;
  if (c.personalNotes && c.personalNotes.trim()) score += 1;
  return score;
}

export function chooseSurvivor(contacts: Contact[]): Contact {
  return [...contacts].sort((a, b) => {
    const s = completenessScore(b) - completenessScore(a);
    if (s !== 0) return s;
    const t = (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? '');
    if (t !== 0) return t;
    return 0;
  })[0];
}

function mergeEmails(ordered: Contact[]): EmailAddress[] {
  const seen = new Set<string>();
  const out: EmailAddress[] = [];
  for (const c of ordered) {
    for (const e of c.emailAddresses ?? []) {
      const key = normalizeEmail(e.address);
      if (key && !seen.has(key)) { seen.add(key); out.push({ name: e.name, address: e.address }); }
    }
  }
  return out;
}

function mergePhones(survivor: Contact, others: Contact[]): Pick<Contact, 'mobilePhone' | 'homePhones' | 'businessPhones'> {
  const seen = new Set<string>();
  const take = (raw?: string | null): boolean => {
    const key = normalizePhone(raw ?? '');
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  };
  const mobileSource = survivor.mobilePhone ?? others.find((c) => c.mobilePhone)?.mobilePhone ?? undefined;
  const mobilePhone = mobileSource ?? null;
  if (mobilePhone) take(mobilePhone);

  const businessPhones: string[] = [];
  for (const c of [survivor, ...others]) for (const p of c.businessPhones ?? []) if (take(p)) businessPhones.push(p);

  const homePhones: string[] = [];
  for (const c of [survivor, ...others]) {
    for (const p of c.homePhones ?? []) if (take(p)) homePhones.push(p);
    if (c.mobilePhone && c.mobilePhone !== mobilePhone && take(c.mobilePhone)) homePhones.push(c.mobilePhone);
  }
  return { mobilePhone, homePhones, businessPhones };
}

function firstNonEmpty<T>(values: (T | undefined | null)[]): T | undefined {
  return values.find((v) => v !== undefined && v !== null && String(v).trim().length > 0) ?? undefined;
}

function mergeNotes(ordered: Contact[]): string | undefined {
  const notes = [...new Set(ordered.map((c) => c.personalNotes?.trim()).filter(Boolean) as string[])];
  return notes.length ? notes.join('\n---\n') : undefined;
}

export function mergeContacts(contacts: Contact[], survivorId?: string): MergePlan {
  if (contacts.length === 0) {
    throw new Error('mergeContacts requires at least one contact');
  }
  const survivor = survivorId
    ? contacts.find((c) => c.id === survivorId)
    : chooseSurvivor(contacts);
  if (!survivor) {
    throw new Error(`mergeContacts: survivorId ${survivorId} not found in group`);
  }
  const others = contacts.filter((c) => c.id !== survivor.id);
  const ordered = [survivor, ...others];

  const merged: Contact = { ...survivor };
  for (const f of SCALAR_FIELDS) {
    if (!merged[f] || String(merged[f]).trim().length === 0) {
      const v = firstNonEmpty(ordered.map((c) => c[f] as string | undefined));
      if (v !== undefined) (merged as unknown as Record<string, unknown>)[f] = v;
    }
  }
  merged.emailAddresses = mergeEmails(ordered);
  Object.assign(merged, mergePhones(survivor, others));
  if (!addressHasData(merged.homeAddress)) merged.homeAddress = ordered.map((c) => c.homeAddress).find(addressHasData);
  if (!addressHasData(merged.businessAddress)) merged.businessAddress = ordered.map((c) => c.businessAddress).find(addressHasData);
  if (!addressHasData(merged.otherAddress)) merged.otherAddress = ordered.map((c) => c.otherAddress).find(addressHasData);
  const notes = mergeNotes(ordered);
  if (notes) merged.personalNotes = notes;

  return {
    survivorId: survivor.id,
    survivor: merged,
    deleteIds: others.map((c) => c.id),
    original: survivor,
    deletedContacts: others,
  };
}
