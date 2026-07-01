# outlook-contact-deduper

A local, single-page web app that finds and merges duplicate contacts in a
**personal Microsoft account** (e.g. `*@live.com`, `*@outlook.com`) via the
Microsoft Graph API.

- Runs entirely in your browser. **No backend, no client secret, no database.**
- Your contacts never leave your machine except to talk to Microsoft Graph.
- Merges duplicates *without losing data* (union of emails/phones/addresses/notes),
  previews every change, and keeps a full JSON backup + session undo.

> Status: **design phase.** See the spec in
> [`docs/superpowers/specs/2026-07-01-outlook-contact-deduper-design.md`](docs/superpowers/specs/2026-07-01-outlook-contact-deduper-design.md).
> Implementation follows an approved plan.

## Why this exists

Classic Outlook's "clean up duplicate contacts" tool is gone from the new
Outlook, but duplicates still clutter downstream clients (e.g. iPhone Contacts).
This app talks to Graph directly to find and merge them, prompting you whenever
a match isn't certain.

## Stack

- Vite + React + TypeScript
- Microsoft auth: `@azure/msal-browser` + `@azure/msal-react` (PKCE, no secret)
- Pure-TypeScript dedupe engine (`src/engine/`), unit-tested with Vitest

## One-time setup: register a free Entra app

The app needs a **client ID** so it can sign you in and ask for permission to
your contacts. This is a free, ~3-minute, one-time step. Microsoft requires it
for personal-account Graph access — the app cannot create it for you.

1. Go to **https://entra.microsoft.com** → **App registrations** → **New registration**.
2. **Name:** `Outlook Contact Deduper`.
3. **Supported account types:** *Personal Microsoft accounts only*.
4. **Redirect URI:** platform **Single-page application (SPA)**, value
   `http://localhost:5173`.
5. Register, then open **API permissions** → **Add a permission** →
   **Microsoft Graph** → **Delegated permissions** → add:
   `Contacts.ReadWrite`, `User.Read`, `offline_access`.
6. Copy the **Application (client) ID** from the Overview page.
7. Copy `src/config.example.ts` to `src/config.ts` and paste your client ID in.
   (`src/config.ts` is git-ignored so your ID never gets committed.)

No client secret is created or needed — a SPA uses PKCE.

## Run (once implemented)

```bash
npm install
npm run dev      # serves at http://localhost:5173 (must match the redirect URI)
npm test         # engine unit tests
```

## Safety

- On load, the app downloads `contacts-backup-YYYY-MM-DD.json` (a full copy of
  all your contacts) before you change anything. These files are git-ignored.
- Every merge shows a preview; certain matches are bulk-approved, uncertain ones
  are reviewed one-by-one; a session undo can revert individual merges.

## Scope

This repo currently covers **de-duping Outlook/Graph contacts only**. Syncing
new iPhone contacts back into the Outlook account is a planned, separate
follow-up.
