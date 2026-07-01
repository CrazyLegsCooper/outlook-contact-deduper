# Outlook Contact Deduper — Design System

Source: generated via `ui-ux-pro-max` (2026-07-01) and adapted from a
landing-page pattern to a **data-review workbench**. This is the single source
of truth for the UI tasks (Tasks 10–12). Implementers MUST follow it; do not
invent alternate colors, fonts, or spacing.

## Principles

- **Trustworthy over flashy.** This app deletes contacts. The UI must make it
  obvious what will change before it changes: previews, confidence badges,
  diff highlighting, and an always-visible undo.
- **Calm neutral canvas, semantic color for meaning.** White cards on a light
  neutral background. Color communicates state, never decoration.
- **Micro-interactions.** Subtle 150–200ms transitions on hover/press; respect
  `prefers-reduced-motion`.

## Styling approach

**Tailwind CSS.** Design tokens are CSS variables defined in `src/index.css`
and mapped into Tailwind via `tailwind.config.js` so components use semantic
class names (`bg-primary`, `text-danger`) rather than raw hex.

## Color tokens

| Token | Hex | Meaning / usage |
|---|---|---|
| `--color-primary` | `#0D9488` | Teal. Primary/safe actions, "very likely" accents, primary buttons. |
| `--color-primary-fg` | `#FFFFFF` | Text/icon on primary. |
| `--color-accent` | `#EA580C` | Orange. "Not sure" review emphasis, attention. |
| `--color-danger` | `#DC2626` | Red. Delete indicators, destructive confirms. |
| `--color-success` | `#059669` | Green. Fields merged IN (diff additions), success toasts. |
| `--color-bg` | `#F8FAFC` | App background (slate-50; neutral, not teal-tinted). |
| `--color-surface` | `#FFFFFF` | Cards, panels. |
| `--color-fg` | `#0F172A` | Primary text (slate-900), contrast ≥ 4.5:1. |
| `--color-muted-fg` | `#64748B` | Secondary text (slate-500). |
| `--color-border` | `#E2E8F0` | Card/divider borders (slate-200). |
| `--color-ring` | `#0D9488` | Focus ring (2px, always visible). |

Confidence: **very-likely** = teal (`primary`); **not-sure** = amber/orange
(`accent`). Never rely on color alone — pair with a text label and an icon.

## Typography

- **Font:** Inter (headings + body), loaded via `@fontsource/inter` (bundled,
  no CDN dependency at runtime).
- **Scale (px):** 12 (meta), 14 (secondary), 16 (body/base), 18 (card title),
  24 (section heading), 32 (screen title).
- **Weights:** 400 body, 500 labels, 600 headings/buttons, 700 screen titles.
- **Line height:** 1.5 body. **Tabular numbers** (`tabular-nums`) for counts.

## Spacing & radius

- 8pt rhythm: `2/3/4/6/8` Tailwind units (8/12/16/24/32 px).
- Card radius `rounded-xl` (12px); buttons/badges `rounded-lg` (8px).
- Card shadow: `shadow-sm` at rest, `shadow-md` on hover for interactive cards.

## Component recipes

Implementers build these as small reusable components in `src/ui/components/`.

### Button (`components/Button.tsx`)
Variants via a `variant` prop:
- `primary` — `bg-primary text-primary-fg hover:brightness-95` (Apply, Sign in)
- `secondary` — `bg-white text-fg border border-border hover:bg-slate-50`
- `ghost` — `text-muted-fg hover:bg-slate-100` (Back, Skip)
- `danger` — `bg-danger text-white hover:brightness-95` (destructive confirms)
All: `h-10 px-4 rounded-lg font-semibold text-sm transition disabled:opacity-50
disabled:cursor-not-allowed focus-visible:outline-none focus-visible:ring-2
focus-visible:ring-ring cursor-pointer`. Min height 40px (touch-friendly).

### ConfidenceBadge (`components/ConfidenceBadge.tsx`)
Props `bucket: 'very-likely' | 'not-sure'`.
- very-likely: `bg-teal-50 text-teal-700 border border-teal-200` + check icon + "Very likely".
- not-sure: `bg-orange-50 text-orange-700 border border-orange-200` + question icon + "Not sure".
`inline-flex items-center gap-1 rounded-lg px-2 py-0.5 text-xs font-medium`.

### Card (`components/Card.tsx`)
`bg-surface border border-border rounded-xl shadow-sm p-4`. Interactive variant
adds `hover:shadow-md transition cursor-pointer`.

### FieldRow (used by MergePreview)
A label + value row. When the value was **merged in** from a non-survivor
record, wrap the added token with the diff-add style so the user sees what's
being combined:
- added token: `bg-emerald-50 text-emerald-800 rounded px-1` + a small `+` prefix.
- survivor's own token: plain `text-fg`.
Emails and phones render as chips; added chips use the diff-add style.

### MergePreview (`ui/MergePreview.tsx`)
A `Card` titled with the survivor display name + `ConfidenceBadge`. Shows rows:
Emails, Phones, Company/Title, Notes — each a `FieldRow`, with merged-in values
diff-highlighted. Footer meta (`text-xs text-muted-fg`): "merging N → 1".

### Side-by-side compare (Not-sure screen)
Grid of candidate `Card`s (`grid gap-4 sm:grid-cols-2`). Each card has a radio
"Keep this" in the top-right; the selected survivor card gets a
`ring-2 ring-primary` highlight. Below the grid, the live `MergePreview` for the
chosen survivor. Primary `Merge` button + `ghost` `Skip`.

### Sticky UndoBar (`ui/UndoBar.tsx`)
Fixed bottom bar, full width: `fixed inset-x-0 bottom-0 bg-white/95 backdrop-blur
border-t border-border px-4 py-3 flex items-center justify-between`. Left:
`N merged` (tabular-nums), failures in `text-danger` if any. Right: `secondary`
"Undo last merge". Screens add `pb-20` so content clears the bar.

### Toast (optional, `components/Toast.tsx`)
`role="status" aria-live="polite"`, auto-dismiss 3–4s, success = `success`
color. Used for "Merged", "Undone", errors (error stays until dismissed).

## Accessibility (must pass)

- Contrast ≥ 4.5:1 for text (tokens above are checked).
- Every control keyboard-reachable; visible `ring-2 ring-ring` focus.
- Icon-only controls get `aria-label`. Icons from **lucide-react** (SVG), never emoji.
- `prefers-reduced-motion`: disable non-essential transitions.
- Confirm before destructive apply (the review action IS the confirmation);
  undo is always available in the sticky bar.

## Anti-patterns (from the generator) — avoid

- Complex onboarding; slow performance.
- Emoji as icons; raw hex in components; gray-on-gray low-contrast text.
