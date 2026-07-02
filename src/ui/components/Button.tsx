import type { ButtonHTMLAttributes } from 'react';

type Variant = 'primary' | 'secondary' | 'ghost' | 'danger';

const VARIANTS: Record<Variant, string> = {
  primary: 'bg-primary text-primary-fg hover:brightness-95',
  secondary: 'bg-white text-fg border border-border hover:bg-slate-50',
  ghost: 'text-muted-fg hover:bg-slate-100',
  danger: 'bg-danger text-white hover:brightness-95',
};

export function Button(
  { variant = 'primary', className = '', ...props }:
  ButtonHTMLAttributes<HTMLButtonElement> & { variant?: Variant },
) {
  return (
    <button
      className={`h-10 px-4 rounded-lg font-semibold text-sm transition disabled:opacity-50 disabled:cursor-not-allowed focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring cursor-pointer ${VARIANTS[variant]} ${className}`}
      {...props}
    />
  );
}
