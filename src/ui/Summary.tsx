import type { AnalyzeResult } from '../engine/analyze';
import { Card } from './components/Card';
import { Button } from './components/Button';
import { ConfidenceBadge } from './components/ConfidenceBadge';

export function Summary({
  result, onReviewVeryLikely, onReviewNotSure,
}: {
  result: AnalyzeResult;
  onReviewVeryLikely: () => void;
  onReviewNotSure: () => void;
}) {
  const vlContacts = result.veryLikely.reduce((n, g) => n + g.ids.length, 0);
  const nsContacts = result.notSure.reduce((n, g) => n + g.ids.length, 0);
  return (
    <div className="mx-auto max-w-3xl p-6">
      <h1 className="text-3xl font-bold">Analysis complete</h1>
      <p className="mt-1 text-muted-fg tabular-nums">
        {result.totalContacts} contacts scanned · backup downloaded
      </p>
      <div className="mt-6 grid gap-4 sm:grid-cols-2">
        <Card>
          <ConfidenceBadge bucket="very-likely" />
          <p className="mt-2 text-2xl font-semibold tabular-nums">{result.veryLikely.length}</p>
          <p className="text-sm text-muted-fg">groups · {vlContacts} contacts</p>
          <Button className="mt-3 w-full" disabled={!result.veryLikely.length} onClick={onReviewVeryLikely}>
            Review &amp; bulk-approve
          </Button>
        </Card>
        <Card>
          <ConfidenceBadge bucket="not-sure" />
          <p className="mt-2 text-2xl font-semibold tabular-nums">{result.notSure.length}</p>
          <p className="text-sm text-muted-fg">groups · {nsContacts} contacts</p>
          <Button variant="secondary" className="mt-3 w-full" disabled={!result.notSure.length} onClick={onReviewNotSure}>
            Review one-by-one
          </Button>
        </Card>
      </div>
    </div>
  );
}
