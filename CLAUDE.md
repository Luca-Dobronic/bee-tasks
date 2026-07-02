# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

This is the **BEE Task System** — a tiny two-file B-BBEE task tracker for Compliance Hub Consulting (CHC), a B-BBEE verification/compliance consultancy, built as a Google Apps Script backend + HTML/JS frontend:

- [`Code.gs`](Code.gs) — Apps Script backend. Deployed as a Web App bound to a Google Sheet; exposes a single `doGet(e)` JSONP/JSON API.
- [`bbbee_v5_2 (2).html`](bbbee_v5_2%20(2).html) — self-contained single-page frontend (inline CSS + vanilla JS, Chart.js from CDN). No build step, no npm, no framework.

There is no local dev server, package manager, linter, or test suite. There is no git repository initialized in this folder. "Running" this app means deploying `Code.gs` as an Apps Script Web App (data layer) and the HTML frontend is deployed separately on **GitHub Pages**, pointed at that deployment's `/exec` URL. The data layer is Google Sheets, accessed exclusively via JSONP (see `jsonpCall` below) to work around the lack of CORS support on Apps Script Web Apps.

The frontend is also deployed as a **PWA** and installed on iPhone (see the manifest/service-worker tags at the top of the HTML `<head>`) — keep PWA installability in mind (manifest.json, icons, service worker registration) when touching the `<head>` or top-level script.

**Core features:**
- Bidirectional task assignment (Luca ↔ Edrich) via the Assign tab
- Five-stage status pipeline (`not_started` → `in_progress`/`blocked`/`review` → `ready_review` → `approved`)
- Comment notifications (per-user "new comment" bell, see the comment-notifications module)
- Checkbox done-state on comments
- Email pings via `MailApp` (Apps Script side) in addition to the frontend's Gmail-compose handoff

**Known issue:** the bound Apps Script project sometimes doesn't appear under "My Projects" in the Apps Script dashboard — still unresolved. Don't assume the script is missing/deleted if a user reports this; it's a known dashboard quirk, not a data-loss symptom.

## Architecture

**Backend (`Code.gs`)** — all state lives in a Google Sheet with two tabs:
- `Tasks` sheet: one row per task. Columns are defined by `TASK_HEADERS` and resolved by name via `col(sheet, name)` — **never by hardcoded index**, since `getTaskSheet()` auto-migrates old sheets by appending any missing header columns with sensible defaults. Preserve this pattern when adding fields.
- `ActivityLog` sheet: append-only audit trail, written via `writeLog()` on every mutation.

All backend routing goes through `doGet(e)` — there is no `doPost`. Actions (`getTasks`, `addTask`, `updateTask`, `deleteTask`, `updateStatus`, `addComment`, `getLog`) are dispatched by `e.parameter.action`, with the payload passed as a JSON-encoded string in a query param (`task`, `payload`, or `id`). Responses are either raw JSON or JSONP-wrapped (`callback(...)`) depending on whether `callback` is present — the frontend always uses JSONP via dynamic `<script>` tag injection (see `jsonpCall` in the HTML) because Apps Script Web Apps don't support CORS for XHR/fetch from arbitrary origins.

Task `status` is the source of truth for workflow state (`not_started` → `in_progress`/`blocked`/`review` → `ready_review` → `approved`); the legacy boolean `done` field is derived/kept in sync for backwards compatibility (`done === true` implies `status === 'approved'`). `pillars` and `comments` are stored as JSON strings in single sheet cells and parsed/stringified at the read/write boundary.

**Frontend (single HTML file)** — five tab "panels" (`p-dash`, `p-board`, `p-assign`, `p-notify`, `p-log`) toggled via `go(panel)`, all sharing one in-memory `tasks` array fetched from the backend:
- Connection: user pastes their Apps Script `/exec` URL once; it's stored in `localStorage` (`bbbee_v5` key) as `SURL`. No auth beyond "knows the URL."
- Polling: `fetchTasks()` runs every 15s (`pollTimer`) plus after every mutation; there's no realtime push. Optimistic local updates are tracked in `pendingSaves` keyed by task id and reconciled/rolled back against the next fetch.
- Identity ("whoami"): which user (`Luca` or `Edrich`) is set once via a modal and stored in `localStorage` (`bbbee_whoami`) — used only for comment attribution and per-user "new comment" notification filtering (`bbbee_comments_seen` tracks last-seen timestamp per task id, client-side only).
- Assign/Notify tabs don't just write to the sheet — they also build a `mailto:`-style Gmail compose URL (`mail.google.com/mail/?view=cm&...`) and `window.open` it, so "sending" a task is really: persist to sheet + hand off to Gmail compose.
- Charts (`updateDash`) are redrawn from scratch (`chart.destroy()` then recreate) rather than updated in place — follow this pattern if adding new charts.

## Conventions worth preserving

- Column-by-name lookups in `Code.gs` (`col()`), never positional indices — this is what makes the auto-migration in `getTaskSheet()` safe.
- All mutations in `Code.gs` call `SpreadsheetApp.flush()` before returning and `writeLog(...)` for the audit trail.
- Frontend state colors/labels are centralized in small lookup maps near the top of the `<script>` block (`SLABELS`, `PMAP`, `PE`, `SMAP`, `SEMO`, `NPMAP`, `TEMPLATES`) — extend these rather than inlining new string logic in render functions.
- The comment-notification system (bottom of the script, clearly delimited by a comment block) is intentionally additive/read-only against `task.comments` — it doesn't call the backend beyond what `addComment`/`fetchAll` already do.
