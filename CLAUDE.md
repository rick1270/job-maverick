# CLAUDE.md — Job Maverick

Job Maverick is the generic, shareable fork of Rick's personal `life_job` tracker. It's the same
Code.js architecture, stripped of Rick-specific data and fitted with a Setup Wizard so a
non-technical friend (Google Sheets + a Claude API key, no CLI/VS Code) can configure their own
copy through the sheet UI alone.

## Critical Rules

- **Never paste code back to the user.** All changes go through clasp push or direct file edits.
- **Config sheet is source of truth** for scoring rules, the user's profile, calibration examples, column definitions, and all prompt/writing-style instructions. Any plain-language instruction text belongs in Config, never as a string literal in Code.js. Code.js should only assemble/interpolate dynamic values around Config-sourced instruction text — pure plumbing, no prose. The only English allowed to remain in Code.js is mechanical/safety constraints tightly coupled to the code's own behavior (e.g. "'find' must be an exact substring", "do not invent experience/metrics").
- **Two Config categories, seeded differently.** Generic scoring/formatting methodology (the ~65 rows in `seedGenericConfigDefaults()`) is code-seeded and identical across every copy of this template. Personal fields (name, contact info, resume URL, career background, target role, compensation) are **never** hardcoded here — they're collected by the Setup Wizard (`setup_wizard.html` / `showSetupWizard()` / `saveSetupWizardAnswers()`) and written per-copy. Never add a new personal-data default to `seedGenericConfigDefaults()` — if a new Config row needs personal input, add a wizard field for it instead.
- **Never overwrite Training Notes.** Manual-only column.
- **Auto Response is written only by `checkJobApplicationEmails()`.** Never generate content for it from job analysis.
- **Apply Anyway is manual-only.** Never generate it from AI analysis. `sanitizeRecommendation()` maps it to Skip.
- **Claude API uses tool use for structured output**, not `json_schema` format (that's OpenAI).
- **The bookmarklet is served dynamically, not hand-edited.** `bookmarklet_template.html` holds the minified bookmarklet JS with `__BASE_URL__`/`__WEBHOOK_TOKEN__` placeholders (kept as `.html` — not `.js` — so Apps Script treats it as a static asset instead of trying to execute it as server code). `getBookmarkletSnippet()` substitutes the live deployment URL and webhook token at runtime; `showBookmarkletDialog()` displays a draggable link + copy-paste fallback. If you change the bookmarklet's capture/extraction logic, re-minify it into `bookmarklet_template.html` (strip the header comment and any `//` line comments, join remaining lines with spaces, verify with `node --check` against a copy with the placeholders swapped for dummy strings) — don't hand-edit the minified file directly.
- GitHub is source of truth for docs (README, CHANGELOG, SETUP_GUIDE).

---

## Two audiences, two docs

- **This file, README.md, CHANGELOG.md** — for Rick, maintaining the template via clasp. Assumes CLI/git/Apps Script editor familiarity.
- **SETUP_GUIDE.md** — for friends setting up their own copy. Zero CLI/git/clasp references; everything is Google Sheets UI + the in-browser Apps Script editor (a GUI, not a local tool, so it's fair game) + the in-sheet Setup Wizard.

---

## Deployment

| Field | Value |
|---|---|
| GitHub | `https://github.com/rick1270/job-Maverick` |
| Script ID | `1dWHxPy_B7NaoCIg8aUuhYz1pJvE-Rs9t7uT1lcSsvnIZS9vLOa5_dnsS` |
| Spreadsheet | `https://drive.google.com/open?id=17BLQw0FTUZJ90zOSxByFkunKIz_NCoLsspIfFUR3t54` |
| Script Properties | `CLAUDE_API_KEY`, `WEBHOOK_TOKEN` (both set via the Setup Wizard, or `Set Claude API Key`/`Set Webhook Token` menu items) |
| Timezone | `America/New_York` (generic default — a friend can change this in the Apps Script editor's `appsscript.json` if needed; not wizard-driven, low-frequency enough not to be worth the extra UI) |

This master copy is **Rick's own template instance** — he maintains it via clasp, same workflow as `life_job`. Friends never touch this specific copy or clasp at all: they get their own independent copy via **File > Make a Copy** on the Sheet (which auto-duplicates the bound Apps Script project), then run the Setup Wizard themselves. Updates to the master template don't propagate automatically to copies already made — see "Porting changes" below.

---

## Development Workflow

Same as `life_job`: edit `Code.js` locally, `clasp push`, test manually via the sheet, commit + push to GitHub. No production web-app deployment step is required for menu-triggered functions; only the bookmarklet's `doGet` path needs an actual **Deploy > New deployment** in the Apps Script editor (each friend does this once, themselves, for their own copy — see SETUP_GUIDE.md).

---

## Porting changes between `life_job` and `job_maverick`

These are separate repos by design (Rick's decision) — not branches of one repo, no automated sync. To bring a fix from one into the other:

```
cd job_maverick
git remote add life_job ../life_job   # one-time
git log life_job/main --oneline -- Code.js
git cherry-pick <hash>
```

Only cherry-pick **mechanical fixes** (bug fixes, refactors, new generic features) — never a commit whose diff touches Rick's personal Config content or hardcoded values. Conflicts are likely in `seedGenericConfigDefaults()`/wizard code that doesn't exist in `life_job` at all; resolve by keeping `job_maverick`'s version of anything wizard/generic-Config-related.

---

## Repo Layout

```
Code.js                    — Apps Script backend, forked from life_job, genericized
appsscript.json             — Apps Script manifest
setup_wizard.html           — Setup Wizard sidebar UI (personal-info form)
bookmarklet_template.html   — Minified bookmarklet JS with placeholder tokens (static asset, not executed)
bookmarklet_dialog.html     — "Get Bookmarklet" dialog (draggable link + copy fallback)
sidebar.html                — Claude Chat sidebar (unchanged from life_job)
training_sidebar.html       — Post-proofread training-notes sidebar (unchanged from life_job)
.clasp.json / .claspignore  — clasp config
CLAUDE.md / README.md / CHANGELOG.md / SETUP_GUIDE.md — docs (local only, excluded from clasp push)
```

---

## Config: generic vs. personal

- **Generic methodology** (~65 rows, `seedGenericConfigDefaults()`): ATS/Fit/Response-Probability/Compensation/Stability/Recommendation/Status/Next-Step definitions and thresholds, the writing-rule templates (Tailoring Notes, Cover Letter, Resume rules, Interview Prep). Identical across every copy unless a friend edits them. Idempotent — safe to re-run.
- **Personal fields** (Setup Wizard): name/contact info, Base Resume URL, Application Docs Folder, target role/industries, career background, work-arrangement and compensation preferences. Never seeded with a default value — wizard-collected per copy.
- **Deliberately omitted from the generic template** (were Rick-specific role-strategy refinements in `life_job`, not generalizable): role-specific scoring sub-rules like "Sales Engineer scoring" / "Data analyst scoring", the granular "Rick background: ..." row cluster (consolidated into one `User Background` field here), and narrowly-scoped gap rules naming specific tools/regions (e.g. a cybersecurity-skills gap rule, a Boston-specific location rule). A friend can add rows like these back manually if their own search needs that level of nuance — documented as an option in SETUP_GUIDE.md, not built into the wizard.

---

## Session Continuity

At the start of each session, read this file, then README.md, then CHANGELOG.md.
