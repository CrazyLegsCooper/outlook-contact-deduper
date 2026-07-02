import { useMemo, useState } from 'react';
import type { Group } from '../engine/analyze';
import type { Contact } from '../engine/types';
import { mergeContacts, type MergePlan } from '../engine/merge';
import { MergePreview } from './MergePreview';
import { SourceContacts } from './SourceContacts';
import { Button } from './components/Button';

export function VeryLikelyReview({
  groups, byId, onApply, onBack,
}: {
  groups: Group[];
  byId: (id: string) => Contact;
  onApply: (plans: MergePlan[]) => void;
  onBack: () => void;
}) {
  const plans = useMemo(
    () => groups.map((g) => mergeContacts(g.ids.map(byId))),
    [groups, byId],
  );
  const [checked, setChecked] = useState<boolean[]>(() => plans.map(() => true));
  const toggle = (i: number) => setChecked((c) => c.map((v, j) => (j === i ? !v : v)));
  const selected = plans.filter((_, i) => checked[i]);

  return (
    <div className="mx-auto max-w-3xl p-6 pb-24">
      <div className="flex items-center justify-between">
        <h1 className="text-2xl font-bold">Very likely duplicates</h1>
        <Button variant="ghost" onClick={onBack}>Back</Button>
      </div>
      <div className="mt-2 flex gap-2">
        <Button variant="secondary" onClick={() => setChecked(plans.map(() => true))}>Approve all</Button>
        <Button variant="ghost" onClick={() => setChecked(plans.map(() => false))}>Clear</Button>
      </div>
      <div className="mt-4 space-y-3">
        {plans.map((plan, i) => (
          <div key={plan.survivorId} className="flex items-start gap-3">
            <input
              type="checkbox"
              className="mt-5 h-4 w-4 accent-[color:var(--color-primary)]"
              checked={checked[i]}
              onChange={() => toggle(i)}
              aria-label={`Approve merge for ${plan.survivor.displayName ?? plan.survivorId}`}
            />
            <div className="flex-1">
              <p className="mb-1 text-xs text-muted-fg">{groups[i].reasons.join(' · ')}</p>
              <SourceContacts contacts={groups[i].ids.map(byId)} />
              <p className="my-1 text-xs font-medium text-muted-fg">merges into ↓</p>
              <MergePreview plan={plan} bucket="very-likely" />
            </div>
          </div>
        ))}
      </div>
      <div className="mt-6 flex items-center gap-3">
        <Button disabled={!selected.length} onClick={() => onApply(selected)}>
          Apply {selected.length} merge(s)
        </Button>
        <span className="text-sm text-muted-fg tabular-nums">{selected.length} of {plans.length} selected</span>
      </div>
    </div>
  );
}
