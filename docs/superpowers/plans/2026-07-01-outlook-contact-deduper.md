# Outlook Contact Deduper Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** A local static React SPA that finds and merges duplicate contacts in a personal Microsoft account via Microsoft Graph, previewing every change and keeping a full backup plus session undo.

**Architecture:** Vite + React + TypeScript single-page app, no backend. A pure-TypeScript dedupe engine (`src/engine/`) does normalization, matching, clustering, and merging and holds all the tested correctness logic. Thin `auth/` (MSAL, PKCE) and `graph/` (fetch, paging, `$batch`) layers talk to Microsoft Graph directly from the browser. React `ui/` renders a bucketed review flow.

**Tech Stack:** Vite, React 18, TypeScript, Vitest, `@azure/msal-browser`, `@azure/msal-react`, Microsoft Graph v1.0.

## Global Constraints

- **Node:** 18+ (Vite 5 requirement). Package manager: npm.
- **Dev server port:** `http://localhost:5173` — MUST match the Entra redirect URI exactly.
- **Auth:** MSAL public client, **PKCE, no client secret**. Authority `https://login.microsoftonline.com/consumers` (personal Microsoft accounts only). Scopes `['User.Read', 'Contacts.ReadWrite']`.
- **Client ID:** read from `src/config.ts` (git-ignored), never hardcoded elsewhere. `src/config.example.ts` is the committed template.
- **Never commit personal data:** contact backups (`contacts-backup-*.json`), any exported contacts, and `src/config.ts` are git-ignored (already in `.gitignore`). Do not add exceptions.
- **Graph write batching:** `$batch` with ≤20 sub-requests per call; honor `429 Retry-After`.
- **No deletes without consent:** a contact is only `DELETE`d after a bulk-approve click or an individual confirm in the UI.
- **UI design system:** all UI (Tasks 10–12) MUST follow `docs/design-system.md` — Tailwind + the CSS-variable tokens, Inter font, `lucide-react` icons (never emoji), semantic color classes (no raw hex in components), confidence badges, diff-highlighted merge previews, and the sticky undo bar. Meet the accessibility bar there (contrast ≥ 4.5:1, visible focus rings, `prefers-reduced-motion`).
- **Discipline:** TDD for the engine and all pure helpers; DRY; YAGNI; commit after every task.

---

### Task 1: Scaffold Vite + React + TS + Vitest

**Files:**
- Create: `package.json`, `vite.config.ts`, `tsconfig.json`, `tsconfig.node.json`, `index.html`, `src/main.tsx`, `src/App.tsx`, `src/vite-env.d.ts`
- Create: `tailwind.config.js`, `postcss.config.js`, `src/index.css` (design tokens)
- Create: `src/config.example.ts`
- Create: `src/engine/.gitkeep`, `src/auth/.gitkeep`, `src/graph/.gitkeep`, `src/ui/.gitkeep`, `src/ui/components/.gitkeep`, `test/.gitkeep`
- Create: `test/smoke.test.ts`

**Interfaces:**
- Consumes: `docs/design-system.md` (color tokens, typography, Tailwind approach).
- Produces: a runnable dev server at `http://localhost:5173`; Tailwind + Inter + design tokens available app-wide; `npm test` runs Vitest; `config.example.ts` exporting `appConfig` shape `{ clientId: string; redirectUri: string; authority: string; scopes: string[] }`.

- [ ] **Step 1: Create `package.json`**

```json
{
  "name": "outlook-contact-deduper",
  "private": true,
  "version": "0.1.0",
  "type": "module",
  "scripts": {
    "dev": "vite --port 5173 --strictPort",
    "build": "tsc && vite build",
    "preview": "vite preview --port 5173 --strictPort",
    "test": "vitest run",
    "test:watch": "vitest"
  },
  "dependencies": {
    "@azure/msal-browser": "^3.10.0",
    "@azure/msal-react": "^2.0.12",
    "@fontsource/inter": "^5.0.0",
    "lucide-react": "^0.379.0",
    "react": "^18.2.0",
    "react-dom": "^18.2.0"
  },
  "devDependencies": {
    "@types/react": "^18.2.0",
    "@types/react-dom": "^18.2.0",
    "@vitejs/plugin-react": "^4.2.0",
    "autoprefixer": "^10.4.0",
    "postcss": "^8.4.0",
    "tailwindcss": "^3.4.0",
    "typescript": "^5.4.0",
    "vite": "^5.2.0",
    "vitest": "^1.4.0"
  }
}
```

- [ ] **Step 2: Create config files**

`vite.config.ts`:
```ts
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';

export default defineConfig({
  plugins: [react()],
  server: { port: 5173, strictPort: true },
});
```

`tsconfig.json`:
```json
{
  "compilerOptions": {
    "target": "ES2020",
    "useDefineForClassFields": true,
    "lib": ["ES2020", "DOM", "DOM.Iterable"],
    "module": "ESNext",
    "skipLibCheck": true,
    "moduleResolution": "bundler",
    "resolveJsonModule": true,
    "isolatedModules": true,
    "noEmit": true,
    "jsx": "react-jsx",
    "strict": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "noFallthroughCasesInSwitch": true
  },
  "include": ["src", "test"],
  "references": [{ "path": "./tsconfig.node.json" }]
}
```

`tsconfig.node.json`:
```json
{
  "compilerOptions": {
    "composite": true,
    "skipLibCheck": true,
    "module": "ESNext",
    "moduleResolution": "bundler",
    "allowSyntheticDefaultImports": true,
    "strict": true
  },
  "include": ["vite.config.ts"]
}
```

- [ ] **Step 3: Create app entry files**

`index.html`:
```html
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Outlook Contact Deduper</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.tsx"></script>
  </body>
</html>
```

`src/vite-env.d.ts`:
```ts
/// <reference types="vite/client" />
```

`src/App.tsx`:
```tsx
export default function App() {
  return <h1>Outlook Contact Deduper</h1>;
}
```

`src/main.tsx`:
```tsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import '@fontsource/inter/400.css';
import '@fontsource/inter/500.css';
import '@fontsource/inter/600.css';
import '@fontsource/inter/700.css';
import './index.css';
import App from './App';

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
```

> Note: `src/main.tsx` is rewritten in Task 7 to add `MsalProvider`; keep these
> font/CSS imports when you do.

- [ ] **Step 3b: Set up Tailwind + design tokens (per `docs/design-system.md`)**

`tailwind.config.js`:
```js
/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        primary: 'var(--color-primary)',
        'primary-fg': 'var(--color-primary-fg)',
        accent: 'var(--color-accent)',
        danger: 'var(--color-danger)',
        success: 'var(--color-success)',
        bg: 'var(--color-bg)',
        surface: 'var(--color-surface)',
        fg: 'var(--color-fg)',
        'muted-fg': 'var(--color-muted-fg)',
        border: 'var(--color-border)',
        ring: 'var(--color-ring)',
      },
      fontFamily: { sans: ['Inter', 'system-ui', 'sans-serif'] },
    },
  },
  plugins: [],
};
```

`postcss.config.js`:
```js
export default { plugins: { tailwindcss: {}, autoprefixer: {} } };
```

`src/index.css`:
```css
@tailwind base;
@tailwind components;
@tailwind utilities;

:root {
  --color-primary: #0D9488;
  --color-primary-fg: #FFFFFF;
  --color-accent: #EA580C;
  --color-danger: #DC2626;
  --color-success: #059669;
  --color-bg: #F8FAFC;
  --color-surface: #FFFFFF;
  --color-fg: #0F172A;
  --color-muted-fg: #64748B;
  --color-border: #E2E8F0;
  --color-ring: #0D9488;
}

html, body, #root { height: 100%; }
body { background: var(--color-bg); color: var(--color-fg); font-family: 'Inter', system-ui, sans-serif; }

@media (prefers-reduced-motion: reduce) {
  * { transition-duration: 0.01ms !important; animation-duration: 0.01ms !important; }
}
```

Update `src/App.tsx` from Step 3 to prove Tailwind works:
```tsx
export default function App() {
  return <h1 className="p-6 text-2xl font-bold text-primary">Outlook Contact Deduper</h1>;
}
```

- [ ] **Step 4: Create `src/config.example.ts`**

```ts
// Copy this file to src/config.ts and paste your Entra Application (client) ID.
// src/config.ts is git-ignored so your ID is never committed.
export const appConfig = {
  clientId: 'YOUR_ENTRA_APPLICATION_CLIENT_ID',
  redirectUri: 'http://localhost:5173',
  authority: 'https://login.microsoftonline.com/consumers',
  scopes: ['User.Read', 'Contacts.ReadWrite'],
};

export type AppConfig = typeof appConfig;
```

- [ ] **Step 5: Add a smoke test**

`test/smoke.test.ts`:
```ts
import { describe, it, expect } from 'vitest';

describe('toolchain', () => {
  it('runs vitest', () => {
    expect(1 + 1).toBe(2);
  });
});
```

- [ ] **Step 6: Install and verify**

Run: `cd /d D:\repos\outlook-contact-deduper && npm install && npm test`
Expected: install completes; Vitest prints `1 passed`.

Run: `npm run dev` then open `http://localhost:5173`.
Expected: page shows "Outlook Contact Deduper" in teal (`text-primary`) Inter bold — confirms Tailwind + tokens + font load. Stop the dev server (Ctrl+C) after confirming.

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "chore: scaffold Vite + React + TS + Vitest + Tailwind design tokens"
```

---

### Task 2: Contact types and normalization

**Files:**
- Create: `src/engine/types.ts`
- Create: `src/engine/normalize.ts`
- Test: `test/normalize.test.ts`

**Interfaces:**
- Consumes: nothing.
- Produces:
  - `Contact`, `EmailAddress`, `PhysicalAddress` interfaces.
  - `normalizeName(raw?: string | null): string`
  - `normalizeEmail(raw?: string | null): string`
  - `normalizePhone(raw?: string | null): string`
  - `contactEmails(c: Contact): string[]` (normalized, de-duped)
  - `contactPhones(c: Contact): string[]` (normalized, de-duped, from mobile+home+business)
  - `contactFullName(c: Contact): string` (normalized; `displayName` else `givenName surname`)

- [ ] **Step 1: Create the types**

`src/engine/types.ts`:
```ts
export interface EmailAddress {
  name?: string;
  address: string;
}

export interface PhysicalAddress {
  street?: string;
  city?: string;
  state?: string;
  countryOrRegion?: string;
  postalCode?: string;
}

/** Subset of the Microsoft Graph `contact` resource we read and write. */
export interface Contact {
  id: string;
  displayName?: string;
  givenName?: string;
  surname?: string;
  nickName?: string;
  emailAddresses?: EmailAddress[];
  mobilePhone?: string | null;
  homePhones?: string[];
  businessPhones?: string[];
  homeAddress?: PhysicalAddress;
  businessAddress?: PhysicalAddress;
  otherAddress?: PhysicalAddress;
  companyName?: string;
  jobTitle?: string;
  department?: string;
  birthday?: string | null;
  personalNotes?: string;
  parentFolderId?: string;
  lastModifiedDateTime?: string;
}
```

- [ ] **Step 2: Write failing tests**

`test/normalize.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import {
  normalizeName, normalizeEmail, normalizePhone,
  contactEmails, contactPhones, contactFullName,
} from '../src/engine/normalize';
import type { Contact } from '../src/engine/types';

describe('normalizeName', () => {
  it('lowercases, trims, collapses whitespace, strips punctuation', () => {
    expect(normalizeName('  Chris   Cooper! ')).toBe('chris cooper');
  });
  it('strips diacritics', () => {
    expect(normalizeName('José')).toBe('jose');
  });
  it('handles null/undefined', () => {
    expect(normalizeName(null)).toBe('');
    expect(normalizeName(undefined)).toBe('');
  });
});

describe('normalizeEmail', () => {
  it('lowercases and trims', () => {
    expect(normalizeEmail('  Chris@X.COM ')).toBe('chris@x.com');
  });
});

describe('normalizePhone', () => {
  it('keeps last 10 digits, dropping formatting and country code', () => {
    expect(normalizePhone('+1 (555) 123-4567')).toBe('5551234567');
  });
  it('returns short numbers as-is', () => {
    expect(normalizePhone('123-456')).toBe('123456');
  });
});

describe('contact aggregates', () => {
  const c: Contact = {
    id: '1',
    displayName: 'Chris Cooper',
    emailAddresses: [{ address: 'Chris@X.com' }, { address: 'chris@x.com' }],
    mobilePhone: '555-123-4567',
    homePhones: ['(555) 123 4567', '555-999-0000'],
  };
  it('de-dupes normalized emails', () => {
    expect(contactEmails(c)).toEqual(['chris@x.com']);
  });
  it('de-dupes normalized phones across fields', () => {
    expect(contactPhones(c)).toEqual(['5551234567', '5559990000']);
  });
  it('derives normalized full name from displayName', () => {
    expect(contactFullName(c)).toBe('chris cooper');
  });
  it('falls back to given + surname', () => {
    expect(contactFullName({ id: '2', givenName: 'John', surname: 'Smith' })).toBe('john smith');
  });
});
```

- [ ] **Step 3: Run to verify failure**

Run: `npm test -- normalize`
Expected: FAIL — cannot find module `../src/engine/normalize`.

- [ ] **Step 4: Implement `src/engine/normalize.ts`**

```ts
import type { Contact } from './types';

export function normalizeName(raw?: string | null): string {
  if (!raw) return '';
  return raw
    .normalize('NFKD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

export function normalizeEmail(raw?: string | null): string {
  if (!raw) return '';
  return raw.trim().toLowerCase();
}

export function normalizePhone(raw?: string | null): string {
  if (!raw) return '';
  const digits = raw.replace(/\D/g, '');
  return digits.length > 10 ? digits.slice(-10) : digits;
}

function uniq(values: string[]): string[] {
  return [...new Set(values.filter((v) => v.length > 0))];
}

export function contactEmails(c: Contact): string[] {
  return uniq((c.emailAddresses ?? []).map((e) => normalizeEmail(e.address)));
}

export function contactPhones(c: Contact): string[] {
  const all = [c.mobilePhone, ...(c.homePhones ?? []), ...(c.businessPhones ?? [])];
  return uniq(all.map((p) => normalizePhone(p ?? '')));
}

export function contactFullName(c: Contact): string {
  if (c.displayName && c.displayName.trim()) return normalizeName(c.displayName);
  return normalizeName([c.givenName, c.surname].filter(Boolean).join(' '));
}
```

- [ ] **Step 5: Run to verify pass**

Run: `npm test -- normalize`
Expected: PASS (all assertions).

- [ ] **Step 6: Commit**

```bash
git add src/engine/types.ts src/engine/normalize.ts test/normalize.test.ts
git commit -m "feat(engine): contact types and normalization helpers"
```

---

### Task 3: String similarity and fuzzy name matching

**Files:**
- Create: `src/engine/names.ts`
- Test: `test/names.test.ts`

**Interfaces:**
- Consumes: nothing (operates on already-normalized name strings).
- Produces:
  - `levenshtein(a: string, b: string): number`
  - `similarityRatio(a: string, b: string): number` (0..1)
  - `namesMatchFuzzy(a: string, b: string): boolean` — inputs are normalized full names; true for nickname/initial/prefix/typo matches.

- [ ] **Step 1: Write failing tests**

`test/names.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { levenshtein, similarityRatio, namesMatchFuzzy } from '../src/engine/names';

describe('levenshtein', () => {
  it('counts single edits', () => {
    expect(levenshtein('kitten', 'sitting')).toBe(3);
    expect(levenshtein('abc', 'abc')).toBe(0);
  });
});

describe('similarityRatio', () => {
  it('is 1 for identical strings', () => {
    expect(similarityRatio('john', 'john')).toBe(1);
  });
  it('is between 0 and 1 for near matches', () => {
    expect(similarityRatio('jon', 'john')).toBeGreaterThan(0.7);
  });
});

describe('namesMatchFuzzy', () => {
  it('matches diminutive/prefix names with same surname', () => {
    expect(namesMatchFuzzy('chris cooper', 'christopher cooper')).toBe(true);
  });
  it('matches initial to full first name', () => {
    expect(namesMatchFuzzy('j smith', 'john smith')).toBe(true);
  });
  it('matches known nicknames', () => {
    expect(namesMatchFuzzy('bob jones', 'robert jones')).toBe(true);
  });
  it('tolerates a typo', () => {
    expect(namesMatchFuzzy('jonh smith', 'john smith')).toBe(true);
  });
  it('does NOT match different surnames', () => {
    expect(namesMatchFuzzy('chris cooper', 'chris baker')).toBe(false);
  });
  it('does NOT match unrelated first names with same surname', () => {
    expect(namesMatchFuzzy('john smith', 'peter smith')).toBe(false);
  });
  it('does NOT match distinct short first names that differ by one substitution', () => {
    // Guards the transposition typo rule from over-matching (regression for the
    // rejected 0.5-ratio approach): these are substitutions, not transpositions.
    expect(namesMatchFuzzy('joan smith', 'john smith')).toBe(false);
    expect(namesMatchFuzzy('mark cooper', 'mary cooper')).toBe(false);
    expect(namesMatchFuzzy('anna lee', 'anne lee')).toBe(false);
  });
  it('returns false when either name is empty', () => {
    expect(namesMatchFuzzy('', 'john smith')).toBe(false);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- names`
Expected: FAIL — cannot find module `../src/engine/names`.

- [ ] **Step 3: Implement `src/engine/names.ts`**

```ts
export function levenshtein(a: string, b: string): number {
  const m = a.length, n = b.length;
  if (m === 0) return n;
  if (n === 0) return m;
  const prev = new Array(n + 1);
  const curr = new Array(n + 1);
  for (let j = 0; j <= n; j++) prev[j] = j;
  for (let i = 1; i <= m; i++) {
    curr[0] = i;
    for (let j = 1; j <= n; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      curr[j] = Math.min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + cost);
    }
    for (let j = 0; j <= n; j++) prev[j] = curr[j];
  }
  return prev[n];
}

export function similarityRatio(a: string, b: string): number {
  if (a === b) return 1;
  const maxLen = Math.max(a.length, b.length);
  if (maxLen === 0) return 1;
  return 1 - levenshtein(a, b) / maxLen;
}

// Known nickname groups (normalized, lowercase). Extend as needed.
const NICKNAME_GROUPS: string[][] = [
  ['robert', 'rob', 'bob', 'bobby'],
  ['william', 'will', 'bill', 'billy'],
  ['richard', 'rich', 'rick', 'dick'],
  ['margaret', 'maggie', 'peggy', 'meg'],
  ['elizabeth', 'liz', 'beth', 'betty', 'eliza'],
  ['james', 'jim', 'jimmy'],
  ['john', 'jack', 'johnny'],
  ['michael', 'mike', 'mick'],
  ['edward', 'ed', 'eddie', 'ted'],
  ['charles', 'charlie', 'chuck'],
];

function sameNicknameGroup(a: string, b: string): boolean {
  return NICKNAME_GROUPS.some((g) => g.includes(a) && g.includes(b));
}

// A single adjacent transposition (e.g. "jonh" vs "john"). This catches typos
// WITHOUT the false positives a low similarity ratio would allow on short names
// (a substitution like "john"/"joan" is NOT a transposition and stays rejected).
function isAdjacentTransposition(a: string, b: string): boolean {
  if (a.length !== b.length) return false;
  const diffs: number[] = [];
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) diffs.push(i);
  return (
    diffs.length === 2 &&
    diffs[1] === diffs[0] + 1 &&
    a[diffs[0]] === b[diffs[1]] &&
    a[diffs[1]] === b[diffs[0]]
  );
}

function firstNamesMatch(a: string, b: string): boolean {
  if (!a || !b) return true; // missing first name should not block a surname match
  if (a === b) return true;
  if ((a.length === 1 && b.startsWith(a)) || (b.length === 1 && a.startsWith(b))) return true;
  const [short, long] = a.length <= b.length ? [a, b] : [b, a];
  if (short.length >= 3 && long.startsWith(short)) return true; // chris -> christopher
  if (sameNicknameGroup(a, b)) return true;
  if (isAdjacentTransposition(a, b)) return true; // "jonh" -> "john" (typo)
  return similarityRatio(a, b) >= 0.8; // tight ratio: no short-name false positives
}

function tokens(name: string): string[] {
  return name.split(' ').filter(Boolean);
}

export function namesMatchFuzzy(a: string, b: string): boolean {
  const ta = tokens(a), tb = tokens(b);
  if (ta.length === 0 || tb.length === 0) return false;
  const aFirst = ta[0], aLast = ta[ta.length - 1];
  const bFirst = tb[0], bLast = tb[tb.length - 1];
  const lastOk =
    ta.length < 2 || tb.length < 2
      ? true
      : aLast === bLast || similarityRatio(aLast, bLast) >= 0.9;
  return lastOk && firstNamesMatch(aFirst, bFirst);
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- names`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/engine/names.ts test/names.test.ts
git commit -m "feat(engine): levenshtein similarity and fuzzy name matching"
```

---

### Task 4: Candidate pair detection and confidence buckets

**Files:**
- Create: `src/engine/match.ts`
- Test: `test/match.test.ts`

**Interfaces:**
- Consumes: `Contact` (Task 2); `contactEmails`, `contactPhones`, `contactFullName` (Task 2); `namesMatchFuzzy` (Task 3).
- Produces:
  - `type Bucket = 'very-likely' | 'not-sure'`
  - `interface CandidatePair { aId: string; bId: string; bucket: Bucket; reasons: string[] }`
  - `findCandidatePairs(contacts: Contact[]): CandidatePair[]`

Bucket rules (from spec §5.2):
- **very-likely** if any: shared email; OR exact normalized full name AND shared phone; OR fuzzy name match AND (shared email OR shared phone).
- **not-sure** if: (exact name OR fuzzy name) AND no shared email AND no shared phone.
- otherwise: no pair emitted.

- [ ] **Step 1: Write failing tests**

`test/match.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { findCandidatePairs } from '../src/engine/match';
import type { Contact } from '../src/engine/types';

const c = (id: string, extra: Partial<Contact>): Contact => ({ id, ...extra });

describe('findCandidatePairs', () => {
  it('flags shared email as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Chris Cooper', emailAddresses: [{ address: 'chris@x.com' }] }),
      c('2', { displayName: 'Christopher Cooper', emailAddresses: [{ address: 'CHRIS@x.com' }] }),
    ]);
    expect(pairs).toHaveLength(1);
    expect(pairs[0].bucket).toBe('very-likely');
    expect(pairs[0].reasons.join(' ')).toMatch(/email/i);
  });

  it('flags exact name + shared phone as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Mum', mobilePhone: '555-111-2222' }),
      c('2', { displayName: 'Mum', homePhones: ['(555) 111 2222'] }),
    ]);
    expect(pairs[0].bucket).toBe('very-likely');
  });

  it('flags fuzzy name + shared phone as very-likely', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'Chris Cooper', mobilePhone: '5551234567' }),
      c('2', { displayName: 'Christopher Cooper', mobilePhone: '555-123-4567' }),
    ]);
    expect(pairs[0].bucket).toBe('very-likely');
  });

  it('flags same name with no shared contact info as not-sure', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'john1@x.com' }] }),
      c('2', { displayName: 'John Smith', emailAddresses: [{ address: 'john2@y.com' }] }),
    ]);
    expect(pairs).toHaveLength(1);
    expect(pairs[0].bucket).toBe('not-sure');
  });

  it('emits no pair for unrelated contacts', () => {
    const pairs = findCandidatePairs([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'john@x.com' }] }),
      c('2', { displayName: 'Peter Baker', emailAddresses: [{ address: 'peter@y.com' }] }),
    ]);
    expect(pairs).toHaveLength(0);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- match`
Expected: FAIL — cannot find module `../src/engine/match`.

- [ ] **Step 3: Implement `src/engine/match.ts`**

```ts
import type { Contact } from './types';
import { contactEmails, contactPhones, contactFullName } from './normalize';
import { namesMatchFuzzy } from './names';

export type Bucket = 'very-likely' | 'not-sure';

export interface CandidatePair {
  aId: string;
  bId: string;
  bucket: Bucket;
  reasons: string[];
}

function intersect(a: string[], b: string[]): string[] {
  const setB = new Set(b);
  return a.filter((x) => setB.has(x));
}

export function findCandidatePairs(contacts: Contact[]): CandidatePair[] {
  const pairs: CandidatePair[] = [];
  const emails = contacts.map(contactEmails);
  const phones = contacts.map(contactPhones);
  const names = contacts.map(contactFullName);

  for (let i = 0; i < contacts.length; i++) {
    for (let j = i + 1; j < contacts.length; j++) {
      const sharedEmails = intersect(emails[i], emails[j]);
      const sharedPhones = intersect(phones[i], phones[j]);
      const nameExact = names[i].length > 0 && names[i] === names[j];
      const nameFuzzy = namesMatchFuzzy(names[i], names[j]);
      const reasons: string[] = [];

      if (sharedEmails.length) reasons.push(`shared email: ${sharedEmails.join(', ')}`);
      if (sharedPhones.length) reasons.push(`shared phone: ${sharedPhones.join(', ')}`);
      if (nameExact) reasons.push('identical name');
      else if (nameFuzzy) reasons.push('similar name');

      const veryLikely =
        sharedEmails.length > 0 ||
        (nameExact && sharedPhones.length > 0) ||
        (nameFuzzy && (sharedEmails.length > 0 || sharedPhones.length > 0));

      const notSure =
        !veryLikely &&
        (nameExact || nameFuzzy) &&
        sharedEmails.length === 0 &&
        sharedPhones.length === 0;

      if (veryLikely) pairs.push({ aId: contacts[i].id, bId: contacts[j].id, bucket: 'very-likely', reasons });
      else if (notSure) pairs.push({ aId: contacts[i].id, bId: contacts[j].id, bucket: 'not-sure', reasons });
    }
  }
  return pairs;
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- match`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/engine/match.ts test/match.test.ts
git commit -m "feat(engine): candidate pair detection with confidence buckets"
```

---

### Task 5: Clustering and analyze

**Files:**
- Create: `src/engine/cluster.ts`
- Create: `src/engine/analyze.ts`
- Test: `test/cluster.test.ts`
- Test: `test/analyze.test.ts`

**Interfaces:**
- Consumes: `CandidatePair`, `Bucket` (Task 4); `Contact` (Task 2); `findCandidatePairs` (Task 4).
- Produces:
  - `clusterPairs(pairs: Array<{ aId: string; bId: string }>): string[][]` — union-find groups of ids (size ≥ 2).
  - `interface Group { ids: string[]; bucket: Bucket; reasons: string[] }`
  - `interface AnalyzeResult { totalContacts: number; veryLikely: Group[]; notSure: Group[] }`
  - `analyze(contacts: Contact[]): AnalyzeResult` — very-likely and not-sure groups are computed separately; any contact appearing in a very-likely group is excluded from not-sure groups (very-likely takes precedence).

- [ ] **Step 1: Write failing test for clustering**

`test/cluster.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { clusterPairs } from '../src/engine/cluster';

describe('clusterPairs', () => {
  it('merges transitive pairs into one group', () => {
    const groups = clusterPairs([
      { aId: 'a', bId: 'b' },
      { aId: 'b', bId: 'c' },
    ]);
    expect(groups).toHaveLength(1);
    expect([...groups[0]].sort()).toEqual(['a', 'b', 'c']);
  });

  it('keeps disjoint pairs as separate groups', () => {
    const groups = clusterPairs([
      { aId: 'a', bId: 'b' },
      { aId: 'x', bId: 'y' },
    ]);
    expect(groups).toHaveLength(2);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- cluster`
Expected: FAIL — cannot find module `../src/engine/cluster`.

- [ ] **Step 3: Implement `src/engine/cluster.ts`**

```ts
export function clusterPairs(pairs: Array<{ aId: string; bId: string }>): string[][] {
  const parent = new Map<string, string>();
  const find = (x: string): string => {
    if (!parent.has(x)) parent.set(x, x);
    let root = x;
    while (parent.get(root) !== root) root = parent.get(root)!;
    let cur = x;
    while (parent.get(cur) !== root) {
      const next = parent.get(cur)!;
      parent.set(cur, root);
      cur = next;
    }
    return root;
  };
  const union = (a: string, b: string) => {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent.set(ra, rb);
  };

  for (const p of pairs) union(p.aId, p.bId);

  const groups = new Map<string, string[]>();
  for (const id of parent.keys()) {
    const root = find(id);
    if (!groups.has(root)) groups.set(root, []);
    groups.get(root)!.push(id);
  }
  return [...groups.values()].filter((g) => g.length >= 2);
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- cluster`
Expected: PASS.

- [ ] **Step 5: Write failing test for analyze**

`test/analyze.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { analyze } from '../src/engine/analyze';
import type { Contact } from '../src/engine/types';

const c = (id: string, extra: Partial<Contact>): Contact => ({ id, ...extra });

describe('analyze', () => {
  it('groups three copies of the same person into one very-likely group', () => {
    const result = analyze([
      c('1', { displayName: 'Mum', mobilePhone: '5551112222' }),
      c('2', { displayName: 'Mum', homePhones: ['555-111-2222'] }),
      c('3', { displayName: 'Mum', businessPhones: ['(555) 111 2222'] }),
    ]);
    expect(result.totalContacts).toBe(3);
    expect(result.veryLikely).toHaveLength(1);
    expect([...result.veryLikely[0].ids].sort()).toEqual(['1', '2', '3']);
    expect(result.notSure).toHaveLength(0);
  });

  it('separates a not-sure group from very-likely', () => {
    const result = analyze([
      c('1', { displayName: 'John Smith', emailAddresses: [{ address: 'a@x.com' }] }),
      c('2', { displayName: 'John Smith', emailAddresses: [{ address: 'b@y.com' }] }),
    ]);
    expect(result.veryLikely).toHaveLength(0);
    expect(result.notSure).toHaveLength(1);
  });

  it('excludes a very-likely contact from not-sure groups', () => {
    // 1 & 2 share an email (very-likely). 2 & 3 share only a name (not-sure).
    const result = analyze([
      c('1', { displayName: 'Ann Lee', emailAddresses: [{ address: 'ann@x.com' }] }),
      c('2', { displayName: 'Ann Lee', emailAddresses: [{ address: 'ann@x.com' }] }),
      c('3', { displayName: 'Ann Lee', emailAddresses: [{ address: 'other@z.com' }] }),
    ]);
    expect(result.veryLikely).toHaveLength(1);
    // id 3 pairs by name with 1 and 2, but those are already in very-likely,
    // so no not-sure group should reference them.
    for (const g of result.notSure) {
      expect(g.ids).not.toContain('1');
      expect(g.ids).not.toContain('2');
    }
  });
});
```

- [ ] **Step 6: Run to verify failure**

Run: `npm test -- analyze`
Expected: FAIL — cannot find module `../src/engine/analyze`.

- [ ] **Step 7: Implement `src/engine/analyze.ts`**

```ts
import type { Contact } from './types';
import { findCandidatePairs, type Bucket, type CandidatePair } from './match';
import { clusterPairs } from './cluster';

export interface Group {
  ids: string[];
  bucket: Bucket;
  reasons: string[];
}

export interface AnalyzeResult {
  totalContacts: number;
  veryLikely: Group[];
  notSure: Group[];
}

function reasonsFor(ids: string[], pairs: CandidatePair[]): string[] {
  const idSet = new Set(ids);
  const reasons = new Set<string>();
  for (const p of pairs) {
    if (idSet.has(p.aId) && idSet.has(p.bId)) p.reasons.forEach((r) => reasons.add(r));
  }
  return [...reasons];
}

export function analyze(contacts: Contact[]): AnalyzeResult {
  const pairs = findCandidatePairs(contacts);
  const vlPairs = pairs.filter((p) => p.bucket === 'very-likely');
  const nsPairs = pairs.filter((p) => p.bucket === 'not-sure');

  const veryLikely: Group[] = clusterPairs(vlPairs).map((ids) => ({
    ids,
    bucket: 'very-likely' as const,
    reasons: reasonsFor(ids, vlPairs),
  }));

  const claimed = new Set(veryLikely.flatMap((g) => g.ids));

  const notSure: Group[] = clusterPairs(nsPairs)
    .map((ids) => ids.filter((id) => !claimed.has(id)))
    .filter((ids) => ids.length >= 2)
    .map((ids) => ({ ids, bucket: 'not-sure' as const, reasons: reasonsFor(ids, nsPairs) }));

  return { totalContacts: contacts.length, veryLikely, notSure };
}
```

- [ ] **Step 8: Run to verify pass**

Run: `npm test -- analyze cluster`
Expected: PASS.

- [ ] **Step 9: Commit**

```bash
git add src/engine/cluster.ts src/engine/analyze.ts test/cluster.test.ts test/analyze.test.ts
git commit -m "feat(engine): union-find clustering and analyze"
```

---

### Task 6: Merge engine

**Files:**
- Create: `src/engine/merge.ts`
- Test: `test/merge.test.ts`

**Interfaces:**
- Consumes: `Contact`, `EmailAddress`, `PhysicalAddress` (Task 2); `normalizePhone`, `normalizeEmail` (Task 2).
- Produces:
  - `interface MergePlan { survivorId: string; survivor: Contact; deleteIds: string[]; original: Contact }`
  - `completenessScore(c: Contact): number`
  - `chooseSurvivor(contacts: Contact[]): Contact`
  - `mergeContacts(contacts: Contact[], survivorId?: string): MergePlan`

Merge rules (spec §6): survivor = most complete (ties → latest `lastModifiedDateTime` → first). Union emails (dedupe by normalized address). Union phones preserving Graph structure: keep survivor's `mobilePhone` (or first available), put all other unique business numbers in `businessPhones`, remaining unique numbers in `homePhones`. Fill empty scalar fields from others. Append differing `personalNotes`. Fill empty address slots from others.

- [ ] **Step 1: Write failing tests**

`test/merge.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { mergeContacts, chooseSurvivor } from '../src/engine/merge';
import type { Contact } from '../src/engine/types';

describe('chooseSurvivor', () => {
  it('picks the most complete contact', () => {
    const a: Contact = { id: '1', displayName: 'Chris' };
    const b: Contact = {
      id: '2', displayName: 'Chris Cooper', givenName: 'Chris', surname: 'Cooper',
      emailAddresses: [{ address: 'c@x.com' }], mobilePhone: '5551234567',
    };
    expect(chooseSurvivor([a, b]).id).toBe('2');
  });
});

describe('mergeContacts', () => {
  it('unions emails and fills empty fields, deleting the others', () => {
    const a: Contact = {
      id: '1', displayName: 'Christopher Cooper', givenName: 'Christopher', surname: 'Cooper',
      emailAddresses: [{ address: 'chris@x.com' }], mobilePhone: '555-123-4567',
    };
    const b: Contact = {
      id: '2', displayName: 'Chris Cooper',
      emailAddresses: [{ address: 'chris.cooper@work.com' }], businessPhones: ['555-999-0000'],
      jobTitle: 'Engineer',
    };
    const plan = mergeContacts([a, b]);
    expect(plan.survivorId).toBe('1'); // a is more complete
    expect(plan.deleteIds).toEqual(['2']);
    expect(plan.survivor.emailAddresses!.map((e) => e.address).sort())
      .toEqual(['chris.cooper@work.com', 'chris@x.com']);
    expect(plan.survivor.mobilePhone).toBe('555-123-4567');
    expect(plan.survivor.businessPhones).toEqual(['555-999-0000']);
    expect(plan.survivor.jobTitle).toBe('Engineer'); // filled from b
    expect(plan.original.id).toBe('1'); // pre-merge snapshot
  });

  it('does not duplicate a phone that appears in two categories', () => {
    const a: Contact = { id: '1', displayName: 'Mum', mobilePhone: '5551112222' };
    const b: Contact = { id: '2', displayName: 'Mum', homePhones: ['(555) 111-2222'] };
    const plan = mergeContacts([a, b]);
    expect(plan.survivor.mobilePhone).toBe('5551112222');
    expect(plan.survivor.homePhones ?? []).toEqual([]); // already covered by mobile
  });

  it('respects an explicit survivorId', () => {
    const a: Contact = { id: '1', displayName: 'A', jobTitle: 'X' };
    const b: Contact = { id: '2', displayName: 'B' };
    const plan = mergeContacts([a, b], '2');
    expect(plan.survivorId).toBe('2');
    expect(plan.deleteIds).toEqual(['1']);
    expect(plan.survivor.jobTitle).toBe('X'); // still filled from a
  });

  it('appends differing personal notes', () => {
    const a: Contact = { id: '1', displayName: 'A', personalNotes: 'note one' };
    const b: Contact = { id: '2', displayName: 'A', personalNotes: 'note two' };
    const plan = mergeContacts([a, b]);
    expect(plan.survivor.personalNotes).toContain('note one');
    expect(plan.survivor.personalNotes).toContain('note two');
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- merge`
Expected: FAIL — cannot find module `../src/engine/merge`.

- [ ] **Step 3: Implement `src/engine/merge.ts`**

```ts
import type { Contact, EmailAddress, PhysicalAddress } from './types';
import { normalizeEmail, normalizePhone } from './normalize';

export interface MergePlan {
  survivorId: string;
  survivor: Contact;
  deleteIds: string[];
  original: Contact;
}

const SCALAR_FIELDS: Array<keyof Contact> = [
  'displayName', 'givenName', 'surname', 'nickName',
  'companyName', 'jobTitle', 'department', 'birthday',
];

function addressHasData(a?: PhysicalAddress): boolean {
  return !!a && Object.values(a).some((v) => v && String(v).trim().length > 0);
}

export function completenessScore(c: Contact): number {
  let score = 0;
  for (const f of SCALAR_FIELDS) if (c[f] && String(c[f]).trim().length > 0) score += 1;
  score += (c.emailAddresses ?? []).length;
  score += (c.mobilePhone ? 1 : 0) + (c.homePhones ?? []).length + (c.businessPhones ?? []).length;
  if (addressHasData(c.homeAddress)) score += 1;
  if (addressHasData(c.businessAddress)) score += 1;
  if (addressHasData(c.otherAddress)) score += 1;
  if (c.personalNotes && c.personalNotes.trim()) score += 1;
  return score;
}

export function chooseSurvivor(contacts: Contact[]): Contact {
  return [...contacts].sort((a, b) => {
    const s = completenessScore(b) - completenessScore(a);
    if (s !== 0) return s;
    const t = (b.lastModifiedDateTime ?? '').localeCompare(a.lastModifiedDateTime ?? '');
    if (t !== 0) return t;
    return 0;
  })[0];
}

function mergeEmails(ordered: Contact[]): EmailAddress[] {
  const seen = new Set<string>();
  const out: EmailAddress[] = [];
  for (const c of ordered) {
    for (const e of c.emailAddresses ?? []) {
      const key = normalizeEmail(e.address);
      if (key && !seen.has(key)) { seen.add(key); out.push({ name: e.name, address: e.address }); }
    }
  }
  return out;
}

function mergePhones(survivor: Contact, others: Contact[]): Pick<Contact, 'mobilePhone' | 'homePhones' | 'businessPhones'> {
  const seen = new Set<string>();
  const take = (raw?: string | null): boolean => {
    const key = normalizePhone(raw ?? '');
    if (!key || seen.has(key)) return false;
    seen.add(key);
    return true;
  };
  const mobileSource = survivor.mobilePhone ?? others.find((c) => c.mobilePhone)?.mobilePhone ?? undefined;
  const mobilePhone = mobileSource ?? null;
  if (mobilePhone) take(mobilePhone);

  const businessPhones: string[] = [];
  for (const c of [survivor, ...others]) for (const p of c.businessPhones ?? []) if (take(p)) businessPhones.push(p);

  const homePhones: string[] = [];
  for (const c of [survivor, ...others]) {
    for (const p of c.homePhones ?? []) if (take(p)) homePhones.push(p);
    if (c.mobilePhone && c.mobilePhone !== mobilePhone && take(c.mobilePhone)) homePhones.push(c.mobilePhone);
  }
  return { mobilePhone, homePhones, businessPhones };
}

function firstNonEmpty<T>(values: (T | undefined | null)[]): T | undefined {
  return values.find((v) => v !== undefined && v !== null && String(v).trim().length > 0) ?? undefined;
}

function mergeNotes(ordered: Contact[]): string | undefined {
  const notes = [...new Set(ordered.map((c) => c.personalNotes?.trim()).filter(Boolean) as string[])];
  return notes.length ? notes.join('\n---\n') : undefined;
}

export function mergeContacts(contacts: Contact[], survivorId?: string): MergePlan {
  if (contacts.length === 0) {
    throw new Error('mergeContacts requires at least one contact');
  }
  const survivor = survivorId
    ? contacts.find((c) => c.id === survivorId)
    : chooseSurvivor(contacts);
  if (!survivor) {
    throw new Error(`mergeContacts: survivorId ${survivorId} not found in group`);
  }
  const others = contacts.filter((c) => c.id !== survivor.id);
  const ordered = [survivor, ...others];

  const merged: Contact = { ...survivor };
  for (const f of SCALAR_FIELDS) {
    if (!merged[f] || String(merged[f]).trim().length === 0) {
      const v = firstNonEmpty(ordered.map((c) => c[f] as string | undefined));
      if (v !== undefined) (merged as Record<string, unknown>)[f] = v;
    }
  }
  merged.emailAddresses = mergeEmails(ordered);
  Object.assign(merged, mergePhones(survivor, others));
  if (!addressHasData(merged.homeAddress)) merged.homeAddress = ordered.map((c) => c.homeAddress).find(addressHasData);
  if (!addressHasData(merged.businessAddress)) merged.businessAddress = ordered.map((c) => c.businessAddress).find(addressHasData);
  if (!addressHasData(merged.otherAddress)) merged.otherAddress = ordered.map((c) => c.otherAddress).find(addressHasData);
  const notes = mergeNotes(ordered);
  if (notes) merged.personalNotes = notes;

  return {
    survivorId: survivor.id,
    survivor: merged,
    deleteIds: others.map((c) => c.id),
    original: survivor,
  };
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- merge`
Expected: PASS.

- [ ] **Step 5: Run the whole engine suite**

Run: `npm test`
Expected: all engine tests pass (normalize, names, match, cluster, analyze, merge, smoke).

- [ ] **Step 6: Commit**

```bash
git add src/engine/merge.ts test/merge.test.ts
git commit -m "feat(engine): merge engine with survivor selection and field union"
```

---

### Task 7: MSAL auth setup and sign-in

**Files:**
- Create: `src/auth/msal.ts`
- Create: `src/auth/AuthGate.tsx`
- Modify: `src/main.tsx` (wrap app in `MsalProvider`)
- Modify: `src/App.tsx` (use auth state, add a "Get token (debug)" button for this task)

**Interfaces:**
- Consumes: `appConfig` from `src/config.ts` (created by the user from `config.example.ts`).
- Produces:
  - `msalInstance: PublicClientApplication`
  - `loginRequest: { scopes: string[] }`
  - `getAccessToken(): Promise<string>` — silent-first, popup fallback.
  - `<AuthGate>` component that renders a sign-in button when unauthenticated and its children when authenticated.

> Prerequisite: the user has completed the Entra registration (README §"One-time setup") and created `src/config.ts` from `src/config.example.ts` with their real client ID. If `src/config.ts` is missing, `npm run dev` will fail to import it — that is the signal to do the registration step.

- [ ] **Step 1: Create `src/auth/msal.ts`**

```ts
import { PublicClientApplication, InteractionRequiredAuthError } from '@azure/msal-browser';
import { appConfig } from '../config';

export const msalInstance = new PublicClientApplication({
  auth: {
    clientId: appConfig.clientId,
    authority: appConfig.authority,
    redirectUri: appConfig.redirectUri,
  },
  cache: { cacheLocation: 'sessionStorage' },
});

export const loginRequest = { scopes: appConfig.scopes };

export async function getAccessToken(): Promise<string> {
  const account = msalInstance.getActiveAccount() ?? msalInstance.getAllAccounts()[0];
  if (!account) throw new Error('No signed-in account');
  try {
    const res = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return res.accessToken;
  } catch (err) {
    if (err instanceof InteractionRequiredAuthError) {
      const res = await msalInstance.acquireTokenPopup(loginRequest);
      return res.accessToken;
    }
    throw err;
  }
}
```

- [ ] **Step 2: Create `src/auth/AuthGate.tsx`**

```tsx
import { useMsal, useIsAuthenticated } from '@azure/msal-react';
import type { ReactNode } from 'react';
import { loginRequest } from './msal';

export function AuthGate({ children }: { children: ReactNode }) {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  if (!isAuthenticated) {
    return (
      <div style={{ padding: 24 }}>
        <h1>Outlook Contact Deduper</h1>
        <p>Sign in with your personal Microsoft account to begin.</p>
        <button
          onClick={() =>
            instance.loginPopup(loginRequest).then((r) => instance.setActiveAccount(r.account))
          }
        >
          Sign in
        </button>
      </div>
    );
  }
  return <>{children}</>;
}
```

- [ ] **Step 3: Wire `MsalProvider` in `src/main.tsx`**

```tsx
import React from 'react';
import ReactDOM from 'react-dom/client';
import { MsalProvider } from '@azure/msal-react';
import { msalInstance } from './auth/msal';
import App from './App';

msalInstance.initialize().then(() => {
  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) msalInstance.setActiveAccount(accounts[0]);

  ReactDOM.createRoot(document.getElementById('root')!).render(
    <React.StrictMode>
      <MsalProvider instance={msalInstance}>
        <App />
      </MsalProvider>
    </React.StrictMode>,
  );
});
```

- [ ] **Step 4: Add a debug token button in `src/App.tsx`**

```tsx
import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { getAccessToken } from './auth/msal';

export default function App() {
  const [status, setStatus] = useState('');
  return (
    <AuthGate>
      <div style={{ padding: 24 }}>
        <h1>Outlook Contact Deduper</h1>
        <button
          onClick={async () => {
            try {
              const t = await getAccessToken();
              setStatus(`Got token, length ${t.length}`);
            } catch (e) {
              setStatus(`Error: ${(e as Error).message}`);
            }
          }}
        >
          Get token (debug)
        </button>
        <p>{status}</p>
      </div>
    </AuthGate>
  );
}
```

- [ ] **Step 5: Manual verification**

1. Ensure `src/config.ts` exists with your real client ID (copy from `config.example.ts`).
2. Run: `npm run dev`, open `http://localhost:5173`.
3. Click **Sign in**, complete the Microsoft popup with `christopher.p.cooper@live.com`, consent to the requested permissions.
4. Click **Get token (debug)**.
Expected: status shows `Got token, length <n>` (a few hundred+ chars). If you see a redirect-URI or unauthorized-client error, re-check the Entra SPA redirect URI is exactly `http://localhost:5173` and the delegated permissions are added.

- [ ] **Step 6: Commit**

```bash
git add src/auth/ src/main.tsx src/App.tsx
git commit -m "feat(auth): MSAL sign-in and token acquisition (PKCE, personal accounts)"
```

---

### Task 8: Graph read and backup download

**Files:**
- Create: `src/graph/select.ts`
- Create: `src/graph/contacts.ts`
- Create: `src/graph/backup.ts`
- Test: `test/backup.test.ts`
- Modify: `src/App.tsx` (add a "Load contacts" button for this task)

**Interfaces:**
- Consumes: `getAccessToken` (Task 7); `Contact` (Task 2).
- Produces:
  - `CONTACT_SELECT: string` — `$select` field list.
  - `listAllContacts(token: string): Promise<Contact[]>` — pages through `@odata.nextLink`.
  - `backupFilename(d: Date): string` → `contacts-backup-YYYY-MM-DD.json`.
  - `downloadBackup(contacts: Contact[], d?: Date): void` — triggers a browser download.

- [ ] **Step 1: Write failing test for backup filename**

`test/backup.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { backupFilename } from '../src/graph/backup';

describe('backupFilename', () => {
  it('formats the date as YYYY-MM-DD', () => {
    expect(backupFilename(new Date('2026-07-01T10:20:30Z'))).toBe('contacts-backup-2026-07-01.json');
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- backup`
Expected: FAIL — cannot find module `../src/graph/backup`.

- [ ] **Step 3: Implement `src/graph/select.ts` and `src/graph/backup.ts`**

`src/graph/select.ts`:
```ts
export const CONTACT_SELECT = [
  'id', 'displayName', 'givenName', 'surname', 'nickName',
  'emailAddresses', 'mobilePhone', 'homePhones', 'businessPhones',
  'homeAddress', 'businessAddress', 'otherAddress',
  'companyName', 'jobTitle', 'department', 'birthday', 'personalNotes',
  'parentFolderId', 'lastModifiedDateTime',
].join(',');
```

`src/graph/backup.ts`:
```ts
import type { Contact } from '../engine/types';

export function backupFilename(d: Date): string {
  const iso = d.toISOString().slice(0, 10);
  return `contacts-backup-${iso}.json`;
}

export function downloadBackup(contacts: Contact[], d: Date = new Date()): void {
  const blob = new Blob([JSON.stringify(contacts, null, 2)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = backupFilename(d);
  a.click();
  URL.revokeObjectURL(url);
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- backup`
Expected: PASS.

- [ ] **Step 5: Implement `src/graph/contacts.ts`**

```ts
import type { Contact } from '../engine/types';
import { CONTACT_SELECT } from './select';

const GRAPH = 'https://graph.microsoft.com/v1.0';

interface ContactsPage {
  value: Contact[];
  '@odata.nextLink'?: string;
}

export async function listAllContacts(token: string): Promise<Contact[]> {
  let url: string | undefined = `${GRAPH}/me/contacts?$top=100&$select=${CONTACT_SELECT}`;
  const all: Contact[] = [];
  while (url) {
    const res = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
    if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
    const page = (await res.json()) as ContactsPage;
    all.push(...page.value);
    url = page['@odata.nextLink'];
  }
  return all;
}
```

- [ ] **Step 6: Add a "Load contacts" button in `src/App.tsx`**

Replace the debug button block with:
```tsx
// inside AuthGate:
const [count, setCount] = useState<number | null>(null);
// ...
<button
  onClick={async () => {
    const token = await getAccessToken();
    const contacts = await listAllContacts(token);
    downloadBackup(contacts);
    setCount(contacts.length);
  }}
>
  Load contacts + backup
</button>
{count !== null && <p>Loaded {count} contacts; backup downloaded.</p>}
```
Add imports: `import { listAllContacts } from './graph/contacts';` and `import { downloadBackup } from './graph/backup';`

- [ ] **Step 7: Manual verification**

1. Run: `npm run dev`, sign in, click **Load contacts + backup**.
Expected: a `contacts-backup-YYYY-MM-DD.json` download appears; the page reports the count. Open the JSON and confirm it contains your contacts. (This file is git-ignored — do not commit it.)

- [ ] **Step 8: Commit**

```bash
git add src/graph/select.ts src/graph/contacts.ts src/graph/backup.ts test/backup.test.ts src/App.tsx
git commit -m "feat(graph): list all contacts (paged) and download JSON backup"
```

---

### Task 9: Graph write via $batch and apply merges

**Files:**
- Create: `src/graph/batch.ts`
- Create: `src/graph/apply.ts`
- Test: `test/batch.test.ts`

**Interfaces:**
- Consumes: `MergePlan` (Task 6); `Contact` (Task 2).
- Produces:
  - `chunk<T>(items: T[], size: number): T[][]`
  - `contactPatchBody(c: Contact): Partial<Contact>` — strips read-only fields (`id`, `parentFolderId`, `lastModifiedDateTime`) before PATCH.
  - `parseRetryAfter(header: string | null): number` — seconds (default 5, min 1).
  - `buildBatchSteps(plans: MergePlan[]): Array<{ id: string; method: string; url: string; body?: unknown; headers?: Record<string, string> }>` — one PATCH per survivor + one DELETE per deleted id.
  - `interface ApplyOutcome { plan: MergePlan; ok: boolean; error?: string }`
  - `applyMergePlans(token: string, plans: MergePlan[]): Promise<ApplyOutcome[]>`
  - `undoMerge(token: string, plan: MergePlan): Promise<void>` — PATCH survivor back to `original`, re-POST deleted contacts (new ids).

- [ ] **Step 1: Write failing tests**

`test/batch.test.ts`:
```ts
import { describe, it, expect } from 'vitest';
import { chunk, contactPatchBody, parseRetryAfter, buildBatchSteps } from '../src/graph/batch';
import type { MergePlan } from '../src/engine/merge';
import type { Contact } from '../src/engine/types';

describe('chunk', () => {
  it('splits into groups of at most size', () => {
    expect(chunk([1, 2, 3, 4, 5], 2)).toEqual([[1, 2], [3, 4], [5]]);
  });
});

describe('contactPatchBody', () => {
  it('strips read-only fields', () => {
    const c: Contact = { id: '1', displayName: 'A', parentFolderId: 'p', lastModifiedDateTime: 't', jobTitle: 'X' };
    const body = contactPatchBody(c);
    expect(body).not.toHaveProperty('id');
    expect(body).not.toHaveProperty('parentFolderId');
    expect(body).not.toHaveProperty('lastModifiedDateTime');
    expect(body.jobTitle).toBe('X');
  });
});

describe('parseRetryAfter', () => {
  it('parses seconds', () => {
    expect(parseRetryAfter('12')).toBe(12);
  });
  it('defaults when missing', () => {
    expect(parseRetryAfter(null)).toBe(5);
  });
});

describe('buildBatchSteps', () => {
  it('emits a PATCH for the survivor and a DELETE per removed id', () => {
    const plan: MergePlan = {
      survivorId: '1',
      survivor: { id: '1', displayName: 'A', jobTitle: 'X' },
      deleteIds: ['2', '3'],
      original: { id: '1', displayName: 'A' },
    };
    const steps = buildBatchSteps([plan]);
    const methods = steps.map((s) => s.method);
    expect(methods.filter((m) => m === 'PATCH')).toHaveLength(1);
    expect(methods.filter((m) => m === 'DELETE')).toHaveLength(2);
    expect(steps.find((s) => s.method === 'PATCH')!.url).toBe('/me/contacts/1');
    expect(steps.filter((s) => s.method === 'DELETE').map((s) => s.url).sort())
      .toEqual(['/me/contacts/2', '/me/contacts/3']);
    // ids must be unique across the batch
    expect(new Set(steps.map((s) => s.id)).size).toBe(steps.length);
  });
});
```

- [ ] **Step 2: Run to verify failure**

Run: `npm test -- batch`
Expected: FAIL — cannot find module `../src/graph/batch`.

- [ ] **Step 3: Implement `src/graph/batch.ts`**

```ts
import type { Contact } from '../engine/types';
import type { MergePlan } from '../engine/merge';

const READ_ONLY: Array<keyof Contact> = ['id', 'parentFolderId', 'lastModifiedDateTime'];

export interface BatchStep {
  id: string;
  method: 'PATCH' | 'DELETE' | 'POST';
  url: string;
  body?: unknown;
  headers?: Record<string, string>;
}

export function chunk<T>(items: T[], size: number): T[][] {
  const out: T[][] = [];
  for (let i = 0; i < items.length; i += size) out.push(items.slice(i, i + size));
  return out;
}

export function contactPatchBody(c: Contact): Partial<Contact> {
  const body: Partial<Contact> = { ...c };
  for (const f of READ_ONLY) delete (body as Record<string, unknown>)[f];
  return body;
}

export function parseRetryAfter(header: string | null): number {
  if (!header) return 5;
  const n = parseInt(header, 10);
  return Number.isFinite(n) && n >= 1 ? n : 5;
}

export function buildBatchSteps(plans: MergePlan[]): BatchStep[] {
  const steps: BatchStep[] = [];
  let n = 0;
  for (const plan of plans) {
    steps.push({
      id: String(++n),
      method: 'PATCH',
      url: `/me/contacts/${plan.survivorId}`,
      body: contactPatchBody(plan.survivor),
      headers: { 'Content-Type': 'application/json' },
    });
    for (const del of plan.deleteIds) {
      steps.push({ id: String(++n), method: 'DELETE', url: `/me/contacts/${del}` });
    }
  }
  return steps;
}
```

- [ ] **Step 4: Run to verify pass**

Run: `npm test -- batch`
Expected: PASS.

- [ ] **Step 5: Implement `src/graph/apply.ts`**

```ts
import type { MergePlan } from '../engine/merge';
import type { Contact } from '../engine/types';
import { buildBatchSteps, chunk, contactPatchBody, parseRetryAfter, type BatchStep } from './batch';

const GRAPH = 'https://graph.microsoft.com/v1.0';

export interface ApplyOutcome {
  plan: MergePlan;
  ok: boolean;
  error?: string;
}

interface BatchResponse {
  responses: Array<{ id: string; status: number; body?: unknown; headers?: Record<string, string> }>;
}

async function postBatch(token: string, steps: BatchStep[]): Promise<BatchResponse> {
  const res = await fetch(`${GRAPH}/$batch`, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify({ requests: steps }),
  });
  if (!res.ok) throw new Error(`$batch ${res.status}: ${await res.text()}`);
  return (await res.json()) as BatchResponse;
}

/** Apply each plan's PATCH + DELETEs. Retries throttled (429) sub-requests once, honoring Retry-After. */
export async function applyMergePlans(token: string, plans: MergePlan[]): Promise<ApplyOutcome[]> {
  const outcomes = new Map<string, ApplyOutcome>();
  // Track which batch-step ids belong to which plan so we can roll status up per plan.
  const stepPlan = new Map<string, MergePlan>();
  const allSteps: BatchStep[] = [];
  for (const plan of plans) {
    outcomes.set(plan.survivorId, { plan, ok: true });
    const steps = buildBatchSteps([plan]);
    for (const s of steps) stepPlan.set(s.id, plan);
    allSteps.push(...steps);
  }

  for (const group of chunk(allSteps, 20)) {
    let pending = group;
    for (let attempt = 0; attempt < 2 && pending.length > 0; attempt++) {
      const { responses } = await postBatch(token, pending);
      const retry: BatchStep[] = [];
      let waitSec = 0;
      for (const r of responses) {
        const plan = stepPlan.get(r.id)!;
        if (r.status === 429) {
          waitSec = Math.max(waitSec, parseRetryAfter(r.headers?.['Retry-After'] ?? null));
          const original = pending.find((s) => s.id === r.id)!;
          retry.push(original);
        } else if (r.status >= 400) {
          outcomes.set(plan.survivorId, { plan, ok: false, error: `status ${r.status}` });
        }
      }
      if (retry.length && attempt === 0) {
        await new Promise((res) => setTimeout(res, waitSec * 1000));
        pending = retry;
      } else {
        pending = [];
      }
    }
  }
  return [...outcomes.values()];
}

/** Revert a single applied merge: restore the survivor's prior state and recreate deleted contacts. */
export async function undoMerge(token: string, plan: MergePlan): Promise<void> {
  const patch = await fetch(`${GRAPH}/me/contacts/${plan.survivorId}`, {
    method: 'PATCH',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: JSON.stringify(contactPatchBody(plan.original)),
  });
  if (!patch.ok) throw new Error(`undo PATCH ${patch.status}: ${await patch.text()}`);
  // Deleted contacts are recreated (they receive new ids). The full JSON came from the backup/analyze set.
  for (const c of plan.deletedContacts ?? []) {
    const post = await fetch(`${GRAPH}/me/contacts`, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json' },
      body: JSON.stringify(contactPatchBody(c as Contact)),
    });
    if (!post.ok) throw new Error(`undo POST ${post.status}: ${await post.text()}`);
  }
}
```

> Note: `undoMerge` needs the full JSON of the deleted contacts. Extend `MergePlan` in `src/engine/merge.ts` to also carry `deletedContacts: Contact[]` (the removed source contacts), set in `mergeContacts` as `others`. Add this field now and update the Task 6 return object: `deletedContacts: others`. Re-run `npm test` to confirm the engine tests still pass (the new field is additive).

- [ ] **Step 6: Add `deletedContacts` to `MergePlan`**

In `src/engine/merge.ts`, add `deletedContacts: Contact[];` to the `MergePlan` interface and `deletedContacts: others,` to the returned object in `mergeContacts`.

Run: `npm test`
Expected: all tests still PASS (additive field).

- [ ] **Step 7: Commit**

```bash
git add src/graph/batch.ts src/graph/apply.ts test/batch.test.ts src/engine/merge.ts
git commit -m "feat(graph): $batch apply with 429 handling and single-merge undo"
```

---

### Task 10: App shell, load + analyze summary

> **Design system (MUST follow `docs/design-system.md`):** all components use
> Tailwind + the design tokens, Inter, and `lucide-react` icons. No inline
> styles, no raw hex, no emoji. This task also creates the shared primitives the
> later UI tasks reuse.

**Files:**
- Create: `src/ui/components/Button.tsx`, `src/ui/components/ConfidenceBadge.tsx`, `src/ui/components/Card.tsx`
- Create: `src/ui/useDeduper.ts` (state hook)
- Create: `src/ui/Summary.tsx`
- Rewrite: `src/App.tsx` (state machine: signin → loading → summary)

**Interfaces:**
- Consumes: `docs/design-system.md`; `getAccessToken` (Task 7); `listAllContacts`, `downloadBackup` (Task 8); `analyze`, `AnalyzeResult`, `Group` (Task 5); `Bucket` (Task 4); `Contact` (Task 2).
- Produces:
  - `<Button variant='primary'|'secondary'|'ghost'|'danger' … />`, `<ConfidenceBadge bucket={Bucket} />`, `<Card interactive? … />`
  - `type Phase = 'idle' | 'loading' | 'ready' | 'applying' | 'done'`
  - `useDeduper()` hook returning `{ phase, contacts, result, load, ... }` (extended in Tasks 11–12).
  - `<Summary result={...} onReviewVeryLikely onReviewNotSure />`

- [ ] **Step 1: Create shared UI primitives** (per `docs/design-system.md`)

`src/ui/components/Button.tsx`:
```tsx
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
```

`src/ui/components/ConfidenceBadge.tsx`:
```tsx
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
```

`src/ui/components/Card.tsx`:
```tsx
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
```

- [ ] **Step 2: Create `src/ui/useDeduper.ts`**

```ts
import { useState, useCallback } from 'react';
import type { Contact } from '../engine/types';
import { analyze, type AnalyzeResult } from '../engine/analyze';
import { getAccessToken } from '../auth/msal';
import { listAllContacts } from '../graph/contacts';
import { downloadBackup } from '../graph/backup';

export type Phase = 'idle' | 'loading' | 'ready' | 'applying' | 'done';

export function useDeduper() {
  const [phase, setPhase] = useState<Phase>('idle');
  const [contacts, setContacts] = useState<Contact[]>([]);
  const [result, setResult] = useState<AnalyzeResult | null>(null);
  const [error, setError] = useState<string | null>(null);

  const load = useCallback(async () => {
    setPhase('loading');
    setError(null);
    try {
      const token = await getAccessToken();
      const all = await listAllContacts(token);
      downloadBackup(all);
      setContacts(all);
      setResult(analyze(all));
      setPhase('ready');
    } catch (e) {
      setError((e as Error).message);
      setPhase('idle');
    }
  }, []);

  const byId = useCallback((id: string) => contacts.find((c) => c.id === id)!, [contacts]);

  return { phase, setPhase, contacts, setContacts, result, setResult, error, load, byId };
}
```

- [ ] **Step 3: Create `src/ui/Summary.tsx`**

```tsx
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
```

- [ ] **Step 4: Rewrite `src/App.tsx`**

```tsx
import { useState } from 'react';
import { AuthGate } from './auth/AuthGate';
import { useDeduper } from './ui/useDeduper';
import { Summary } from './ui/Summary';
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
      {/* very-likely and not-sure screens are added in Tasks 11 and 12 */}
    </AuthGate>
  );
}
```

Also update `src/auth/AuthGate.tsx` from Task 7 to use the `Button` primitive and Tailwind (replace the inline-styled sign-in markup):
```tsx
// keep the useMsal/useIsAuthenticated logic; render:
return (
  <div className="mx-auto max-w-3xl p-6">
    <h1 className="text-3xl font-bold text-primary">Outlook Contact Deduper</h1>
    <p className="mt-1 text-muted-fg">Sign in with your personal Microsoft account to begin.</p>
    <Button className="mt-4"
      onClick={() => instance.loginPopup(loginRequest).then((r) => instance.setActiveAccount(r.account))}>
      Sign in
    </Button>
  </div>
);
// add: import { Button } from '../ui/components/Button';
```

- [ ] **Step 5: Manual verification**

1. Run: `npm run dev`, sign in, click **Load contacts + analyze**.
Expected: backup downloads; the summary shows total count and two token-styled cards (teal "Very likely", orange "Not sure") with tabular counts; buttons enable only when groups exist; focus rings visible when tabbing.

- [ ] **Step 6: Commit**

```bash
git add src/ui/components src/ui/useDeduper.ts src/ui/Summary.tsx src/App.tsx src/auth/AuthGate.tsx
git commit -m "feat(ui): design-system primitives, app shell, and summary screen"
```

---

### Task 11: Very-likely bulk-approve review

> **Design system (MUST follow `docs/design-system.md`):** Tailwind + tokens;
> reuse `Button`, `Card`, `ConfidenceBadge`; merge previews use diff-highlighted
> chips for values merged in from the non-survivor. No inline styles / raw hex.

**Files:**
- Create: `src/ui/MergePreview.tsx`
- Create: `src/ui/VeryLikelyReview.tsx`
- Modify: `src/ui/useDeduper.ts` (add `applyPlans`, `appliedOutcomes`)
- Modify: `src/App.tsx` (render `very-likely` screen)

**Interfaces:**
- Consumes: `docs/design-system.md`; `Group` (Task 5); `Bucket` (Task 4); `mergeContacts`, `MergePlan` (Task 6); `normalizeEmail`, `normalizePhone` (Task 2); `applyMergePlans`, `ApplyOutcome` (Task 9); `getAccessToken` (Task 7); primitives `Button`/`Card`/`ConfidenceBadge` (Task 10).
- Produces:
  - `<MergePreview plan={MergePlan} bucket?={Bucket} />` — merged survivor fields with diff-highlighted merged-in chips.
  - `<VeryLikelyReview groups byId onApply onBack />` — checkbox list with "Approve all"/"Apply selected".
  - `useDeduper().applyPlans(plans: MergePlan[]): Promise<void>` (sets phase `applying` → `ready`, stores outcomes and removes merged contacts from state, re-runs `analyze`).

- [ ] **Step 1: Create `src/ui/MergePreview.tsx`**

```tsx
import type { MergePlan } from '../engine/merge';
import type { Bucket } from '../engine/match';
import { normalizeEmail, normalizePhone } from '../engine/normalize';
import { Card } from './components/Card';
import { ConfidenceBadge } from './components/ConfidenceBadge';

function Chip({ label, added }: { label: string; added: boolean }) {
  return (
    <span className={`inline-block rounded px-1.5 py-0.5 text-sm ${added ? 'bg-emerald-50 text-emerald-800' : 'text-fg'}`}>
      {added ? '+ ' : ''}{label}
    </span>
  );
}

export function MergePreview({ plan, bucket }: { plan: MergePlan; bucket?: Bucket }) {
  const s = plan.survivor;
  const o = plan.original;
  const origEmails = new Set((o.emailAddresses ?? []).map((e) => normalizeEmail(e.address)));
  const origPhones = new Set(
    [o.mobilePhone, ...(o.homePhones ?? []), ...(o.businessPhones ?? [])].map((p) => normalizePhone(p ?? '')),
  );
  const emails = s.emailAddresses ?? [];
  const phones = [s.mobilePhone, ...(s.homePhones ?? []), ...(s.businessPhones ?? [])].filter(Boolean) as string[];
  const name = s.displayName || `${s.givenName ?? ''} ${s.surname ?? ''}`.trim() || '(no name)';
  return (
    <Card>
      <div className="flex items-center justify-between gap-2">
        <h3 className="text-lg font-semibold">{name}</h3>
        {bucket && <ConfidenceBadge bucket={bucket} />}
      </div>
      <dl className="mt-3 space-y-2 text-sm">
        <div className="flex gap-2">
          <dt className="w-20 shrink-0 text-muted-fg">Emails</dt>
          <dd className="flex flex-wrap gap-1">
            {emails.length
              ? emails.map((e) => <Chip key={e.address} label={e.address} added={!origEmails.has(normalizeEmail(e.address))} />)
              : <span className="text-muted-fg">—</span>}
          </dd>
        </div>
        <div className="flex gap-2">
          <dt className="w-20 shrink-0 text-muted-fg">Phones</dt>
          <dd className="flex flex-wrap gap-1">
            {phones.length
              ? phones.map((p) => <Chip key={p} label={p} added={!origPhones.has(normalizePhone(p))} />)
              : <span className="text-muted-fg">—</span>}
          </dd>
        </div>
        {s.jobTitle && <div className="flex gap-2"><dt className="w-20 shrink-0 text-muted-fg">Title</dt><dd>{s.jobTitle}</dd></div>}
        {s.companyName && <div className="flex gap-2"><dt className="w-20 shrink-0 text-muted-fg">Company</dt><dd>{s.companyName}</dd></div>}
      </dl>
      <p className="mt-3 text-xs text-muted-fg tabular-nums">merging {plan.deleteIds.length + 1} → 1</p>
    </Card>
  );
}
```

- [ ] **Step 2: Add `applyPlans` to `src/ui/useDeduper.ts`**

Add inside `useDeduper`, before the `return`:
```ts
const [appliedOutcomes, setAppliedOutcomes] = useState<import('../graph/apply').ApplyOutcome[]>([]);

const applyPlans = useCallback(async (plans: import('../engine/merge').MergePlan[]) => {
  setPhase('applying');
  const token = await getAccessToken();
  const { applyMergePlans } = await import('../graph/apply');
  const outcomes = await applyMergePlans(token, plans);
  const removed = new Set(plans.flatMap((p) => p.deleteIds));
  setContacts((prev) => {
    const next = prev.filter((c) => !removed.has(c.id));
    setResult(analyze(next));
    return next;
  });
  setAppliedOutcomes((prev) => [...prev, ...outcomes]);
  setPhase('ready');
}, []);
```
Add `applyPlans, appliedOutcomes, setAppliedOutcomes` to the returned object.

- [ ] **Step 3: Create `src/ui/VeryLikelyReview.tsx`**

```tsx
import { useMemo, useState } from 'react';
import type { Group } from '../engine/analyze';
import type { Contact } from '../engine/types';
import { mergeContacts, type MergePlan } from '../engine/merge';
import { MergePreview } from './MergePreview';
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
```

- [ ] **Step 4: Render the screen in `src/App.tsx`**

Add below the summary block:
```tsx
{d.phase === 'ready' && d.result && screen === 'very-likely' && (
  <VeryLikelyReview
    groups={d.result.veryLikely}
    byId={d.byId}
    onApply={async (plans) => { await d.applyPlans(plans); setScreen('summary'); }}
    onBack={() => setScreen('summary')}
  />
)}
{d.phase === 'applying' && <div className="mx-auto max-w-3xl p-6 text-muted-fg">Applying merges…</div>}
```
Add import: `import { VeryLikelyReview } from './ui/VeryLikelyReview';`

- [ ] **Step 5: Manual verification (use caution — this writes to your account)**

1. Run: `npm run dev`, sign in, load + analyze.
2. Open **Review** for very-likely. Confirm previews look right. Untick anything wrong.
3. Click **Apply**. Wait for it to return to summary.
Expected: the very-likely count drops; in Outlook/Graph the merged survivors now hold the unioned info and the redundant copies are gone. The pre-merge state is preserved in your backup JSON if anything looks off.

- [ ] **Step 6: Commit**

```bash
git add src/ui/MergePreview.tsx src/ui/VeryLikelyReview.tsx src/ui/useDeduper.ts src/App.tsx
git commit -m "feat(ui): very-likely bulk-approve review and apply"
```

---

### Task 12: Not-sure one-by-one review, undo, and done

> **Design system (MUST follow `docs/design-system.md`):** Tailwind + tokens;
> reuse `Button`, `Card`, `MergePreview`; side-by-side candidate cards with the
> chosen survivor ring-highlighted; sticky bottom `UndoBar` using `lucide-react`.
> No inline styles / raw hex / emoji.

**Files:**
- Create: `src/ui/NotSureReview.tsx`
- Create: `src/ui/UndoBar.tsx`
- Modify: `src/ui/useDeduper.ts` (add `undoLast`)
- Modify: `src/App.tsx` (render `not-sure` screen + undo bar)

**Interfaces:**
- Consumes: `docs/design-system.md`; `Group` (Task 5); `mergeContacts`, `chooseSurvivor`, `MergePlan` (Task 6); `applyMergePlans`, `undoMerge`, `ApplyOutcome` (Task 9); `Contact` (Task 2); primitives `Button`/`Card` and `MergePreview` (Tasks 10–11).
- Produces:
  - `<NotSureReview group byId onMerge onSkip onBack index total />` — side-by-side members, survivor selector, Merge/Skip.
  - `<UndoBar outcomes onUndo />` — sticky bottom bar; undo the last applied merge.
  - `useDeduper().undoLast(): Promise<void>`.

- [ ] **Step 1: Create `src/ui/NotSureReview.tsx`**

```tsx
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
```

- [ ] **Step 2: Add `undoLast` to `src/ui/useDeduper.ts`**

Add inside `useDeduper`:
```ts
const undoLast = useCallback(async () => {
  const last = appliedOutcomes[appliedOutcomes.length - 1];
  if (!last) return;
  const token = await getAccessToken();
  const { undoMerge } = await import('../graph/apply');
  await undoMerge(token, last.plan);
  setAppliedOutcomes((prev) => prev.slice(0, -1));
  // Reload is the simplest correct way to resync ids after a recreate.
}, [appliedOutcomes]);
```
Add `undoLast` to the returned object.

- [ ] **Step 3: Create `src/ui/UndoBar.tsx`**

```tsx
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
```

- [ ] **Step 4: Render not-sure flow + undo bar in `src/App.tsx`**

Add a not-sure index state near the top of `App`: `const [nsIndex, setNsIndex] = useState(0);`
Add below the very-likely block:
```tsx
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
```
Add imports: `import { NotSureReview } from './ui/NotSureReview';`, `import { UndoBar } from './ui/UndoBar';` (and `Button` is already imported from Task 10).

- [ ] **Step 5: Manual verification (writes to your account)**

1. Run: `npm run dev`, sign in, load + analyze, open **not-sure** review.
2. For a group, pick which record to keep, review the merged preview, click **Merge** (or **Skip**).
3. After a merge, use **Undo last merge** in the bottom bar; then reload contacts and confirm the deleted contact reappears (with a new id) and the survivor is back to its pre-merge fields.
Expected: merges and undo both behave; skipping advances to the next group.

- [ ] **Step 6: Full test run + commit**

Run: `npm test`
Expected: all engine/helper tests PASS.

```bash
git add src/ui/NotSureReview.tsx src/ui/UndoBar.tsx src/ui/useDeduper.ts src/App.tsx
git commit -m "feat(ui): not-sure review, session undo, and apply flow"
```

---

## Self-Review Notes

**Spec coverage check (spec → task):**
- §2 architecture (static SPA, no backend, MSAL, direct Graph) → Tasks 1, 7, 8, 9.
- §3 Entra registration → README + Task 7 prerequisite.
- §4 data model (Graph contact subset) → Task 2 (`types.ts`), Task 8 (`$select`).
- §5.1 normalization → Task 2.
- §5.2 buckets → Task 4 (+ §5.3 fuzzy names in Task 3).
- §5.3 clustering (union-find) → Task 5.
- §6 merge (survivor, field union, phone structure, notes, preview, PATCH+DELETE, `$batch`) → Task 6 (engine) + Task 9 (Graph) + Task 11 (preview).
- §7 safety (backup, session undo, no silent deletes) → Task 8 (backup), Task 9 + Task 12 (undo), Tasks 11–12 (consent before delete).
- §8 UI flow → Tasks 10, 11, 12.
- **UI design system** (`docs/design-system.md`, added 2026-07-01 via ui-ux-pro-max) → Task 1 (Tailwind + tokens + Inter), Task 10 (shared primitives + summary), Tasks 11–12 (diff-highlighted previews, side-by-side compare, sticky undo). Enforced by the "UI design system" Global Constraint.
- §9 layout / §10 testing → reflected across all tasks; engine is unit-tested, wrappers have pure-helper tests + manual checks.
- §11 risks: fuzzy algorithm chosen (Task 3); throttling handled (Task 9); personal-account read/write verified live (Task 7–8 manual steps). Photo-merge is intentionally deferred (survivor keeps its photo) — noted here as the one spec §6 detail not implemented, to keep scope tight; revisit if photos matter.

**Type consistency:** `Contact`, `MergePlan` (incl. added `deletedContacts`), `Group`, `AnalyzeResult`, `Bucket`, `CandidatePair`, `BatchStep`, `ApplyOutcome`, and function names (`analyze`, `mergeContacts`, `chooseSurvivor`, `findCandidatePairs`, `clusterPairs`, `applyMergePlans`, `undoMerge`, `listAllContacts`, `downloadBackup`, `getAccessToken`) are used consistently across tasks.

**Open item to confirm during execution:** photo handling (deferred) — acceptable per scope, flagged for the user.
