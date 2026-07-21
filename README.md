# Job Maverick

A shareable, generic version of Rick's personal job-search tracker (`life_job`) — same Google
Apps Script + Google Sheets + Claude API architecture, minus the personal data. New users
configure their own copy through an in-sheet Setup Wizard, no CLI or code editor required.

See **SETUP_GUIDE.md** if you're setting up your own copy. See **CLAUDE.md** for the
dev-facing architecture/maintenance notes (Rick's own reference).

---

## Current State

| Field | Value |
|---|---|
| Last updated | 2026-07-21 |
| Code version | v1.0 |
| Forked from | `life_job` v2.20 |
| AI backend | Claude (model IDs configurable in Config) |

### What Works
- Everything from `life_job` v2.20: bookmarklet capture, Claude-scored job analysis, Config-driven scoring rules, tailoring/cover-letter/resume generation, Interview Prep, Gmail email monitor (hourly + daily extended + lock-guarded against overlap), sort logic (including Response Probability tie-break), conditional-formatting repair, new-Claude-model detection, orphaned-doc auto-cleanup.
- **New**: Setup Wizard (`Job Assistant → Setup Wizard`) collects personal info, resume/docs links, target role, background, and preferences, and writes them into Config — no manual Config editing required to get started.
- **New**: `seedGenericConfigDefaults()` seeds ~65 generic scoring/formatting Config rows (methodology only, no personal data) — same idempotent pattern as `life_job`'s other Config seed functions.
- **New**: `Job Assistant → Get Bookmarklet` generates a ready-to-drag bookmarklet using the copy's own deployment URL and webhook token — no manual file editing or minification needed.

### Known Differences from `life_job`
- No calibration examples pre-seeded (Calibration Summary starts blank — same starting point `life_job` effectively had before Rick's own calibration history accumulated).
- No role-specific scoring sub-rules (e.g. "Sales Engineer scoring") — those were `life_job`-specific refinements to Rick's own target role; a friend can add equivalent custom Config rows for their own target role if they want that level of nuance.
- Timezone defaults to `America/New_York` in `appsscript.json` — editable manually in the Apps Script editor, not wizard-driven.

---

## Tech Stack

Same as `life_job`: Chrome bookmarklet (chunked GET via image pixel pings) → Google Apps Script
backend → Claude API (tool use for structured output) → Google Sheets (6 tabs) → GitHub for
version control.

---

## Roadmap

| Version | Focus |
|---|---|
| v1.0 ✅ | Fork from `life_job` v2.20; genericized Config; Setup Wizard; dynamic bookmarklet generation |
