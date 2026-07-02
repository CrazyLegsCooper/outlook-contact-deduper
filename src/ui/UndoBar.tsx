import { Undo2 } from 'lucide-react';
import type { ApplyOutcome } from '../graph/apply';
import { Button } from './components/Button';

export function UndoBar({ outcomes, onUndo }: { outcomes: ApplyOutcome[]; onUndo: () => void }) {
  if (!outcomes.length) return null;
  const merged = outcomes.filter((o) => o.ok).length;
  const failed = outcomes.filter((o) => !o.ok).length;
  return (
    <div className="fixed inset-x-0 bottom-0 border-t border-border bg-white/95 px-6 py-3 backdrop-blur">
      <div className="mx-auto flex max-w-3xl items-center justify-between">
        <span className="text-sm tabular-nums">
          <span className="font-semibold text-success">{merged} merged</span>
          {failed ? <span className="ml-2 text-danger">{failed} failed</span> : null}
        </span>
        <Button variant="secondary" onClick={onUndo}>
          <Undo2 size={16} className="mr-1 inline" aria-hidden />
          Undo last merge
        </Button>
      </div>
    </div>
  );
}
