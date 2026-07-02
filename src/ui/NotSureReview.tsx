import { useState } from 'react';
import type { Group } from '../engine/analyze';
import type { Contact } from '../engine/types';
import { mergeContacts, chooseSurvivor, type MergePlan } from '../engine/merge';
import { MergePreview } from './MergePreview';
import { Card } from './components/Card';
import { Button } from './components/Button';

export function NotSureReview({
  group, byId, index, total, onMerge, onSkip, onBack,
}: {
  group: Group;
  byId: (id: string) => Contact;
  index: number;
  total: number;
  onMerge: (plan: MergePlan) => void;
  onSkip: () => void;
  onBack: () => void;
}) {
  const members = group.ids.map(byId);
  const [survivorId, setSurvivorId] = useState<string>(chooseSurvivor(members).id);
  const plan = mergeContacts(members, survivorId);

  return (
    <div className="mx-auto max-w-3xl p-6 pb-24">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold">Not sure — {index + 1} of {total}</h1>
        <Button variant="ghost" onClick={onBack}>Back</Button>
      </div>
      <p className="mt-1 text-sm text-muted-fg">{group.reasons.join(' · ') || 'same/similar name only'}</p>
      <div className="mt-4 grid gap-4 sm:grid-cols-2">
        {members.map((m) => {
          const selected = survivorId === m.id;
          const phones = [m.mobilePhone, ...(m.homePhones ?? []), ...(m.businessPhones ?? [])].filter(Boolean) as string[];
          return (
            <Card key={m.id} className={selected ? 'ring-2 ring-primary' : ''}>
              <label className="flex cursor-pointer items-center justify-between gap-2">
                <span className="font-semibold">{m.displayName || '(no name)'}</span>
                <span className="inline-flex items-center gap-1 text-xs text-muted-fg">
                  <input
                    type="radio"
                    name="survivor"
                    className="h-4 w-4 accent-[color:var(--color-primary)]"
                    checked={selected}
                    onChange={() => setSurvivorId(m.id)}
                  />
                  keep this
                </span>
              </label>
              <div className="mt-2 text-sm text-muted-fg">{(m.emailAddresses ?? []).map((e) => e.address).join(', ') || '—'}</div>
              <div className="text-sm text-muted-fg">{phones.join(', ') || '—'}</div>
            </Card>
          );
        })}
      </div>
      <h3 className="mt-6 text-sm font-semibold text-muted-fg">Merged result</h3>
      <div className="mt-2"><MergePreview plan={plan} bucket="not-sure" /></div>
      <div className="mt-4 flex gap-3">
        <Button onClick={() => onMerge(plan)}>Merge</Button>
        <Button variant="ghost" onClick={onSkip}>Skip</Button>
      </div>
    </div>
  );
}
