import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { useDeduper } from './ui/useDeduper';
import { Summary } from './ui/Summary';
import { VeryLikelyReview } from './ui/VeryLikelyReview';
import { NotSureReview } from './ui/NotSureReview';
import { UndoBar } from './ui/UndoBar';
import { Button } from './ui/components/Button';

type Screen = 'summary' | 'very-likely' | 'not-sure';

export default function App() {
  const d = useDeduper();
  const [screen, setScreen] = useState<Screen>('summary');
  const [nsIndex, setNsIndex] = useState(0);

  return (
    <AuthGate>
      {d.phase === 'idle' && (
        <div className="mx-auto max-w-3xl p-6">
          <h1 className="text-3xl font-bold text-primary">Outlook Contact Deduper</h1>
          <p className="mt-1 text-muted-fg">Scan your contacts, back them up, and merge duplicates safely.</p>
          {d.error && (
            <p className="mt-3 rounded-lg border border-danger/30 bg-red-50 px-3 py-2 text-sm text-danger" role="alert">
              Error: {d.error}
            </p>
          )}
          <Button className="mt-4" onClick={d.load}>Load contacts + analyze</Button>
        </div>
      )}
      {d.phase === 'loading' && (
        <div className="mx-auto max-w-3xl p-6 text-muted-fg">Loading contacts…</div>
      )}
      {d.phase === 'ready' && d.result && screen === 'summary' && (
        <Summary
          result={d.result}
          onReviewVeryLikely={() => setScreen('very-likely')}
          onReviewNotSure={() => setScreen('not-sure')}
        />
      )}
      {d.phase === 'ready' && d.result && screen === 'very-likely' && (
        <VeryLikelyReview
          groups={d.result.veryLikely}
          byId={d.byId}
          onApply={async (plans) => { await d.applyPlans(plans); setScreen('summary'); }}
          onBack={() => setScreen('summary')}
        />
      )}
      {d.phase === 'applying' && <div className="mx-auto max-w-3xl p-6 text-muted-fg">Applying merges…</div>}
      {d.phase === 'ready' && d.result && screen === 'not-sure' && d.result.notSure.length > 0 && (
        <NotSureReview
          group={d.result.notSure[Math.min(nsIndex, d.result.notSure.length - 1)]}
          byId={d.byId}
          index={Math.min(nsIndex, d.result.notSure.length - 1)}
          total={d.result.notSure.length}
          onMerge={async (plan) => { await d.applyPlans([plan]); setNsIndex(0); }}
          onSkip={() => setNsIndex((i) => (i + 1 < d.result!.notSure.length ? i + 1 : 0) )}
          onBack={() => { setNsIndex(0); setScreen('summary'); }}
        />
      )}
      {d.phase === 'ready' && d.result && screen === 'not-sure' && d.result.notSure.length === 0 && (
        <div className="mx-auto max-w-3xl p-6">
          <p className="text-muted-fg">No more not-sure groups.</p>
          <Button className="mt-3" variant="secondary" onClick={() => setScreen('summary')}>Back to summary</Button>
        </div>
      )}
      <UndoBar outcomes={d.appliedOutcomes} onUndo={d.undoLast} />
    </AuthGate>
  );
}
