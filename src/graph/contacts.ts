import type { Contact } from '../engine/types';
import { CONTACT_SELECT } from './select';

const GRAPH = 'https://graph.microsoft.com/v1.0';

interface ContactsPage {
  value: Contact[];
  '@odata.nextLink'?: string;
}

export async function listAllContacts(token: string): Promise<Contact[]> {
  let url: string | undefined = `${GRAPH}/me/contacts?$top=100&$select=${CONTACT_SELECT}`;
  const all: Contact[] = [];
  while (url) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
    const page = (await res.json()) as ContactsPage;
    all.push(...page.value);
    url = page['@odata.nextLink'];
  }
  return all;
}
