import { describe, it, expect } from 'vitest';
import { renderToStaticMarkup } from 'react-dom/server';
import { SourceContacts } from '../src/ui/SourceContacts';
import type { Contact } from '../src/engine/types';

describe('SourceContacts', () => {
  it('shows every source contact — with original phone formatting and modified date', () => {
    // Two byte-identical duplicates that differ only in phone formatting + date:
    // the merged result alone would hide the duplication, which is the whole bug.
    const contacts: Contact[] = [
      { id: '1', displayName: 'Dave Barysh', mobilePhone: '+1 201-245-7127', lastModifiedDateTime: '2026-02-15T18:55:31Z' },
      { id: '2', displayName: 'Dave Barysh', mobilePhone: '+1-201-245-7127', lastModifiedDateTime: '2025-04-07T17:37:00Z' },
    ];
    const html = renderToStaticMarkup(<SourceContacts contacts={contacts} />);

    expect(html).toContain('2 source contacts');
    // both copies rendered, each keeping its ORIGINAL phone format (the disambiguator)
    expect(html).toContain('+1 201-245-7127');
    expect(html).toContain('+1-201-245-7127');
    // dates that distinguish the otherwise-identical copies
    expect(html).toContain('2026-02-15');
    expect(html).toContain('2025-04-07');
  });

  it('renders emails and singular label for a one-contact list', () => {
    const html = renderToStaticMarkup(
      <SourceContacts contacts={[{ id: '1', displayName: 'X', emailAddresses: [{ address: 'x@y.com' }] }]} />,
    );
    expect(html).toContain('1 source contact');
    expect(html).not.toContain('1 source contacts');
    expect(html).toContain('x@y.com');
  });
});
