# Outlook Contact Deduper — Design

- **Date:** 2026-07-01
- **Owner:** Christopher Cooper (christopher.p.cooper@live.com)
- **Repo:** github.com/CrazyLegsCooper/outlook-contact-deduper
- **Status:** Approved design, pre-implementation

## 1. Problem & Goal

The new Outlook web/app environment no longer offers the classic "clean up
duplicate contacts" tool. Duplicates remain in the account and clutter
downstream clients (notably the iPhone Contacts app, which no longer hides them
the way the old Outlook did).

**Goal:** a small, safe tool that finds duplicate contacts in a personal
Microsoft account and merges them without losing information, prompting the user
whenever a match is not certain.

**Success criteria:**

- Duplicate contacts in the account are meaningfully reduced.
- No data is lost: merging is a union of fields, never a silent overwrite.
- Nothing is deleted without either a bulk-approve action or an individual
  confirmation.
- A full backup exists before any change, and individual merges can be undone
  within the session.

**Out of scope (separate follow-up project):** investigating/fixing why new
iPhone contacts are not syncing into the Outlook account (likely the iPhone is
saving to iCloud or "On My iPhone" rather than the Outlook account). This spec
covers Outlook/Graph de-duping only.

## 2. Approach & Architecture

A **static single-page web app** served locally, talking directly to Microsoft
Graph. No backend, no client secret, no database. Contacts live only in the
browser tab and in the backup file the user downloads.

```
  Browser tab (http://localhost:5173)
    │
    ├─ MSAL (@azure/msal-browser + @azure/msal-react)
    │     └─ sign in with christopher.p.cooper@live.com (popup, PKCE)
    │        consent to Contacts.ReadWrite / User.Read / offline_access
    │
    ├─ Graph client (fetch)
    │     GET    /me/contacts            (paged — fetch all)
    │     PATCH  /me/contacts/{id}       (write merged survivor)
    │     DELETE /me/contacts/{id}       (remove redundant copies)
    │     POST   /$batch                 (≤20 ops/request, 429-aware)
    │
    └─ Dedupe engine (pure TypeScript, no DOM)
          normalize → match/cluster → merge preview
```

**Why React + MSAL:** Microsoft ships first-party `@azure/msal-react` /
`@azure/msal-browser` with a well-supported personal-account (MSA) sign-in flow,
removing the hardest part of the project. Graph supports CORS from the browser,
so the SPA calls Graph directly with no proxy. Alternatives considered (Rust +
Tauri desktop; Rust/WASM SPA) were rejected for this task: no Rust MSAL exists,
so both require hand-rolling the MSA OAuth flow for no benefit on a contacts
utility.

**Serving nuance:** the sign-in redirect URI must be a real
`http://localhost:<port>` (not `file://`), so the page is served by Vite's dev
server on a fixed port (5173) that matches the registered redirect URI. This is
static file serving only — there is still no application backend.

## 3. One-time Entra app registration (manual, by the user)

The app requires a **client ID**. This is created once, for free, in the Entra
portal; the app cannot create it. Steps (also in README):

1. entra.microsoft.com → App registrations → New registration.
2. Name `Outlook Contact Deduper`; account type **Personal Microsoft accounts
   only**.
3. Platform **Single-page application (SPA)**, redirect URI
   `http://localhost:5173`.
4. API permissions → Microsoft Graph → **Delegated**: `Contacts.ReadWrite`,
   `User.Read`, `offline_access`.
5. Copy the **Application (client) ID** into `src/config.ts` (copied from
   `src/config.example.ts`; `src/config.ts` is git-ignored).

No client secret is created (PKCE). Consent for a personal account happens at
first sign-in.

## 4. Data model (relevant Microsoft Graph `contact` fields)

Read/written per contact:

- `id`, `displayName`, `givenName`, `surname`, `nickName`
- `emailAddresses: [{ name, address }]`
- `mobilePhone` (single string), `homePhones: []`, `businessPhones: []`
- `homeAddress`, `businessAddress`, `otherAddress` (physicalAddress objects)
- `companyName`, `jobTitle`, `department`
- `birthday`, `personalNotes`
- `parentFolderId`, `lastModifiedDateTime` (used for survivor selection)
- Photo is a separate resource (`/me/contacts/{id}/photo/$value`) — not merged
  byte-for-byte; the survivor keeps its photo, or inherits the other's photo
  only if the survivor has none. (Photo handling detail finalized in the plan.)

## 5. Dedupe engine (pure, unit-tested)

Kept DOM-free in `src/engine/` so it can be tested against fabricated contact
sets independently of auth/UI.

### 5.1 Normalization

- **Name:** lowercase, trim, collapse internal whitespace, strip punctuation;
  derive a normalized full name from `displayName` and from
  `givenName + surname`.
- **Email:** lowercase, trim.
- **Phone:** strip all non-digits; compare on the last 10 digits (to ignore
  country-code / formatting differences).

### 5.2 Candidate matching → confidence buckets

Contacts are compared pairwise (blocking by shared email/normalized-name key to
avoid unnecessary O(n²) work at scale, but correctness does not depend on the
blocking).

| Bucket | Trigger | UX |
| --- | --- | --- |
| **Very likely** | Any shared normalized email; OR identical normalized name **and** a shared phone; OR fuzzy name match (nickname/initial expansion, e.g. Chris↔Christopher, J.↔John) **and** a shared email or phone; OR all key fields identical | Shown in a **bulk-approve** list ("Approve all" / untick individually) |
| **Not sure** | Same or similar name but **no** overlapping email or phone (could be two different people) | **One-by-one** review: side-by-side, pick survivor, merge/skip |
| **Ignore** | No shared identifiers and dissimilar names | Not surfaced |

Fuzzy name matching uses nickname/initial expansion plus a string-similarity
threshold (exact algorithm/threshold chosen in the plan; must treat
Chris↔Christopher and J.↔John as matches).

### 5.3 Clustering

Candidate pairs are combined with **union-find** so transitive duplicates form a
single group (e.g. three "Mum" contacts become one cluster merged into one
survivor, not three separate pairwise merges).

## 6. Merge behavior

- **Survivor selection:** the most complete record (most non-empty fields);
  ties broken by most recent `lastModifiedDateTime`, then by having a photo. In
  the one-by-one review the user can override which record is the survivor.
- **Field merge:** union `emailAddresses` (dedupe by normalized address); union
  all phones (`mobilePhone` + `homePhones` + `businessPhones`) with de-dupe;
  fill any empty scalar field on the survivor (`givenName`, `surname`,
  `companyName`, `jobTitle`, `birthday`, addresses) from the other record;
  append differing `personalNotes` with a separator; keep the survivor's photo
  unless it has none and the other does.
- **Preview before apply:** the merged result is shown before it is written
  (always for "not sure"; available per row for "very likely").
- **Apply:** `PATCH` the survivor with merged fields, then `DELETE` the
  redundant contact(s). Operations are grouped into Graph `$batch` requests
  (≤20 ops each) with handling for `429 Too Many Requests` (respect
  `Retry-After`).

## 7. Safety net

- **Full backup:** after fetching all contacts on load, the app auto-downloads
  `contacts-backup-YYYY-MM-DD.json` (raw full copies of every contact) before
  any change. Git-ignored; also useful for the later iPhone work.
- **Session undo:** each applied merge records the survivor's pre-merge state
  and the deleted contact(s)' full JSON. "Undo" re-`PATCH`es the survivor to its
  prior state and re-creates the deleted contact(s) via `POST` (a re-created
  contact receives a new `id` — expected and acceptable).
- **No silent deletes:** deletion only follows a bulk-approve click or an
  individual confirmation.

## 8. UI flow

1. **Sign in** (MSAL popup) → consent.
2. **Load:** fetch all contacts (paged, with progress); auto-download backup
   JSON.
3. **Analyze:** run the engine; show a summary (N contacts, X very-likely
   groups, Y not-sure groups).
4. **Very likely tab:** list of groups, each with a merged preview and a
   checkbox; "Approve all" or untick individually; apply selected.
5. **Not sure tab:** group-by-group side-by-side; choose survivor; merge or
   skip.
6. **Apply:** batched `PATCH`/`DELETE` with progress and per-op results; undo
   available.
7. **Done:** summary of merged/deleted counts and backup file location.

## 9. Project layout

```
outlook-contact-deduper/
  index.html
  package.json
  vite.config.ts
  src/
    main.tsx
    config.example.ts     ← template (committed)
    config.ts             ← real client ID (git-ignored)
    auth/                 ← MSAL setup, sign-in, token acquisition
    graph/                ← fetch wrapper, paging, $batch, 429 handling, backup
    engine/
      normalize.ts
      match.ts            ← buckets + union-find clustering
      merge.ts            ← survivor selection + field union
    ui/                   ← React review screens (summary, very-likely, not-sure)
  test/                   ← Vitest unit tests for engine/ with contact fixtures
  docs/superpowers/specs/
  README.md
```

## 10. Testing strategy

- **Engine (`src/engine/`):** real unit tests (Vitest) over fabricated contact
  fixtures — normalization, each confidence bucket, union-find clustering, and
  merge field-union / survivor selection. This is where correctness lives.
- **Graph / auth wrappers:** thin; exercised manually against the real account,
  plus small unit tests for pure helpers (paging cursor, `$batch` chunking,
  `Retry-After` parsing).

## 11. Risks / open items (resolve in plan)

- Exact fuzzy-name algorithm and similarity threshold.
- Photo merge details (fetch/compare cost vs. simply preferring a survivor with
  a photo).
- Personal-account Graph quirks: verify `/me/contacts` read/write early with a
  small live check before building the full flow.
- Graph throttling behavior on bulk `$batch` deletes — validate batch size and
  backoff against the real account.
