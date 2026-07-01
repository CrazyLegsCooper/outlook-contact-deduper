import type { Contact } from './types';
import { contactEmails, contactPhones, contactFullName } from './normalize';
import { namesMatchFuzzy } from './names';

export type Bucket = 'very-likely' | 'not-sure';

export interface CandidatePair {
  aId: string;
  bId: string;
  bucket: Bucket;
  reasons: string[];
}

function intersect(a: string[], b: string[]): string[] {
  const setB = new Set(b);
  return a.filter((x) => setB.has(x));
}

export function findCandidatePairs(contacts: Contact[]): CandidatePair[] {
  const pairs: CandidatePair[] = [];
  const emails = contacts.map(contactEmails);
  const phones = contacts.map(contactPhones);
  const names = contacts.map(contactFullName);

  for (let i = 0; i < contacts.length; i++) {
    for (let j = i + 1; j < contacts.length; j++) {
      const sharedEmails = intersect(emails[i], emails[j]);
      const sharedPhones = intersect(phones[i], phones[j]);
      const nameExact = names[i].length > 0 && names[i] === names[j];
      const nameFuzzy = namesMatchFuzzy(names[i], names[j]);
      const reasons: string[] = [];

      if (sharedEmails.length) reasons.push(`shared email: ${sharedEmails.join(', ')}`);
      if (sharedPhones.length) reasons.push(`shared phone: ${sharedPhones.join(', ')}`);
      if (nameExact) reasons.push('identical name');
      else if (nameFuzzy) reasons.push('similar name');

      const veryLikely =
        sharedEmails.length > 0 ||
        (nameExact && sharedPhones.length > 0) ||
        (nameFuzzy && (sharedEmails.length > 0 || sharedPhones.length > 0));

      const notSure =
        !veryLikely &&
        (nameExact || nameFuzzy) &&
        sharedEmails.length === 0 &&
        sharedPhones.length === 0;

      if (veryLikely) pairs.push({ aId: contacts[i].id, bId: contacts[j].id, bucket: 'very-likely', reasons });
      else if (notSure) pairs.push({ aId: contacts[i].id, bId: contacts[j].id, bucket: 'not-sure', reasons });
    }
  }
  return pairs;
}
