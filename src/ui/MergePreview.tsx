import type { MergePlan } from '../engine/merge';
import type { Bucket } from '../engine/match';
import { normalizeEmail, normalizePhone } from '../engine/normalize';
import { Card } from './components/Card';
import { ConfidenceBadge } from './components/ConfidenceBadge';

function Chip({ label, added }: { label: string; added: boolean }) {
  return (
    <span className={`inline-block rounded px-1.5 py-0.5 text-sm ${added ? 'bg-emerald-50 text-emerald-800' : 'text-fg'}`}>
      {added ? '+ ' : ''}{label}
    </span>
  );
}

export function MergePreview({ plan, bucket }: { plan: MergePlan; bucket?: Bucket }) {
  const s = plan.survivor;
  const o = plan.original;
  const origEmails = new Set((o.emailAddresses ?? []).map((e) => normalizeEmail(e.address)));
  const origPhones = new Set(
    [o.mobilePhone, ...(o.homePhones ?? []), ...(o.businessPhones ?? [])].map((p) => normalizePhone(p ?? '')),
  );
  const emails = s.emailAddresses ?? [];
  const phones = [s.mobilePhone, ...(s.homePhones ?? []), ...(s.businessPhones ?? [])].filter(Boolean) as string[];
  const name = s.displayName || `${s.givenName ?? ''} ${s.surname ?? ''}`.trim() || '(no name)';
  return (
    <Card>
      <div className="flex items-center justify-between gap-2">
        <h3 className="text-lg font-semibold">{name}</h3>
        {bucket && <ConfidenceBadge bucket={bucket} />}
      </div>
      <dl className="mt-3 space-y-2 text-sm">
        <div className="flex gap-2">
          <dt className="w-20 shrink-0 text-muted-fg">Emails</dt>
          <dd className="flex flex-wrap gap-1">
            {emails.length
              ? emails.map((e) => <Chip key={e.address} label={e.address} added={!origEmails.has(normalizeEmail(e.address))} />)
              : <span className="text-muted-fg">—</span>}
          </dd>
        </div>
        <div className="flex gap-2">
          <dt className="w-20 shrink-0 text-muted-fg">Phones</dt>
          <dd className="flex flex-wrap gap-1">
            {phones.length
              ? phones.map((p) => <Chip key={p} label={p} added={!origPhones.has(normalizePhone(p))} />)
              : <span className="text-muted-fg">—</span>}
          </dd>
        </div>
        {s.jobTitle && <div className="flex gap-2"><dt className="w-20 shrink-0 text-muted-fg">Title</dt><dd>{s.jobTitle}</dd></div>}
        {s.companyName && <div className="flex gap-2"><dt className="w-20 shrink-0 text-muted-fg">Company</dt><dd>{s.companyName}</dd></div>}
      </dl>
      <p className="mt-3 text-xs text-muted-fg tabular-nums">merging {plan.deleteIds.length + 1} → 1</p>
    </Card>
  );
}
