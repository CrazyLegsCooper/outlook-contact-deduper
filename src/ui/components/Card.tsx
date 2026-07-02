import type { HTMLAttributes } from 'react';

export function Card(
  { interactive = false, className = '', ...props }:
  HTMLAttributes<HTMLDivElement> & { interactive?: boolean },
) {
  return (
    <div
      className={`bg-surface border border-border rounded-xl shadow-sm p-4 ${interactive ? 'hover:shadow-md transition cursor-pointer' : ''} ${className}`}
      {...props}
    />
  );
}
