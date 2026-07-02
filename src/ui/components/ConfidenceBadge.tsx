import { CheckCircle2, HelpCircle } from 'lucide-react';
import type { Bucket } from '../../engine/match';

export function ConfidenceBadge({ bucket }: { bucket: Bucket }) {
  const isVL = bucket === 'very-likely';
  const cls = isVL
    ? 'bg-teal-50 text-teal-700 border-teal-200'
    : 'bg-orange-50 text-orange-700 border-orange-200';
  const Icon = isVL ? CheckCircle2 : HelpCircle;
  return (
    <span className={`inline-flex items-center gap-1 rounded-lg border px-2 py-0.5 text-xs font-medium ${cls}`}>
      <Icon size={14} aria-hidden />
      {isVL ? 'Very likely' : 'Not sure'}
    </span>
  );
}
