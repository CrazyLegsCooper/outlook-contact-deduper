import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { useDeduper } from './ui/useDeduper';
import { Summary } from './ui/Summary';
import { VeryLikelyReview } from './ui/VeryLikelyReview';
import { Button } from './ui/components/Button';

type Screen = 'summary' | 'very-likely' | 'not-sure';

export default function App() {
  const d = useDeduper();
  const [screen, setScreen] = useState<Screen>('summary');

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
      {/* not-sure screen is added in Task 12 */}
    </AuthGate>
  );
}
