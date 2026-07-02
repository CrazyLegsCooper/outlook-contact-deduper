import type { Contact } from '../engine/types';

function phonesOf(c: Contact): string[] {
  return [c.mobilePhone, ...(c.homePhones ?? []), ...(c.businessPhones ?? [])].filter(Boolean) as string[];
}

function nameOf(c: Contact): string {
  return c.displayName || `${c.givenName ?? ''} ${c.surname ?? ''}`.trim() || '(no name)';
}

/**
 * Read-only list of the actual contacts feeding a merge, so the reviewer can
 * see *why* a group is a duplicate — including for identical copies that the
 * merged result alone makes look like a single ordinary contact. Phone numbers
 * keep their original formatting and the last-modified date is shown, which is
 * what distinguishes otherwise-identical duplicates.
 */
export function SourceContacts({ contacts }: { contacts: Contact[] }) {
  return (
    <div className="rounded-lg border border-border bg-bg/60 p-2">
      <p className="mb-1 px-1 text-xs font-medium text-muted-fg">
        {contacts.length} source contact{contacts.length === 1 ? '' : 's'}
      </p>
      <ul className="divide-y divide-border">
        {contacts.map((c) => {
          const phones = phonesOf(c);
          const emails = (c.emailAddresses ?? []).map((e) => e.address);
          const modified = c.lastModifiedDateTime?.slice(0, 10);
          return (
            <li key={c.id} className="flex flex-wrap items-baseline gap-x-2 px-1 py-1 text-sm">
              <span className="font-medium text-fg">{nameOf(c)}</span>
              {phones.map((p) => <span key={p} className="text-muted-fg">{p}</span>)}
              {emails.map((e) => <span key={e} className="text-muted-fg">{e}</span>)}
              {modified && <span className="ml-auto text-xs text-muted-fg tabular-nums">{modified}</span>}
            </li>
          );
        })}
      </ul>
    </div>
  );
}
