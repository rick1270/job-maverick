# Job Maverick — Setup Guide

Everything below happens in your browser — Google Sheets and the Apps Script editor (which
opens inside Google, not a separate program). No installs, no command line.

(This same content also lives as a formatted **Instructions** tab inside the sheet itself —
that's what a friend will actually see; this file is a source/reference copy.)

## What this tool does

- One-click job capture — a browser bookmarklet grabs the job description from any posting
  (LinkedIn, company career pages, etc.) and sends it straight to your tracker.
- Automatic scoring — every captured job gets a Fit Score, ATS Score, Response Probability, and
  a Recommendation (Apply / Discuss / Apply if Easy / Skip), all driven by rules you control in
  the Config tab.
- Tailored resumes and cover letters — generates a version of your resume and a cover letter
  matched to each specific job description, with keyword and ATS-aware editing.
- Interview prep sheets — a study sheet per interview: likely questions, talking points pulled
  from your real resume, a terms/technology glossary, and gaps to watch for.
- Email monitoring — automatically checks your inbox for application responses (confirmations,
  rejections, interview requests, offers) and updates your tracker, no manual entry needed.
- Auto-sorting — your job list re-sorts by priority as things change.
- Fully customizable scoring — the Config tab is plain-English instructions, not code; edit any
  rule any time.
- Built-in help — a Claude Chat sidebar for questions, and a feedback tool to report bugs or
  suggest features directly to the developer.

## What you'll need

- A Google account with Google Sheets.
- A Claude API key from [console.anthropic.com](https://console.anthropic.com) (Plans & Billing →
  API keys). This is a paid, pay-as-you-go API — typical usage for a job search costs a few
  dollars a month, not a subscription.
- Your resume as a Google Doc (not a PDF/Word file — if you have a PDF, open it in Google Docs
  first: File > Open > Upload, or paste the text into a new Doc).

## 1. Get your own copy

Open the Job Maverick template sheet (link from whoever shared this with you), then:

**File → Make a copy**

This gives you your own independent spreadsheet *and* a copy of all the code behind it — nothing
you do in your copy affects the original template or anyone else's copy.

## 2. Deploy the web app (one-time, needed for job capture)

This step lets the bookmarklet (step 4) send job postings to your sheet.

1. In your copy, go to **Extensions → Apps Script**. This opens the code editor — you don't need
   to read or understand the code, just navigate its menus.
2. Click **Deploy** (top right) → **New deployment**.
3. Click the gear icon next to "Select type" → choose **Web app**.
4. Set "Execute as" to **Me**, and "Who has access" to **Anyone**.
5. Click **Deploy**. The first time, Google will ask you to authorize the script — click through
   the permission screens (you may see a "Google hasn't verified this app" warning; click
   **Advanced → Go to [project name] (unsafe)** — this is expected for a personal script you
   just created yourself, not a real security warning).
6. You don't need to copy the deployment URL shown here — the Setup Wizard (next step) fetches
   it automatically.

## 3. Run the Setup Wizard

Back in your spreadsheet: **Job Assistant → Setup Wizard**.

Fill in each section — personal info, your resume/docs links, target role, background, work
preferences, and your Claude API key. Nothing here needs to be perfect on the first pass; every
field is a normal Config cell you can edit later directly in the Config tab.

Click **Save Setup**.

## 4. Install your bookmarklet

**Job Assistant → Get Bookmarklet**. Drag the blue button into your browser's bookmarks bar. That's
your job-capture button — click it on any job posting page (LinkedIn, a company careers page,
etc.) and the description gets sent to your sheet automatically, scored within about 30 seconds.

If dragging doesn't work in your browser, the same dialog gives you the code as text — create a
new bookmark manually (right-click your bookmarks bar → Add page) and paste it as the URL.

## 5. Turn on the email monitor (optional but recommended)

This checks your inbox for application responses (confirmations, rejections, interview
requests) and updates your tracker automatically.

1. In the Apps Script editor (Extensions → Apps Script), find the function dropdown near the
   top toolbar.
2. Select `installEmailMonitorTrigger` and click **Run**.
3. Also run `installExtendedEmailCheckTrigger` (a once-daily backup check) and
   `installNewModelCheckTrigger` (a weekly check for newer Claude models) the same way.
4. For each of these, add an **Email Domain** to a job row (e.g. `greenhouse.io` or the
   company's own domain) so the monitor knows what to search for on that application.

## You're set up. Day to day:

- Click your bookmarklet on job postings you find.
- Check the **Jobs** tab — new rows appear scored (Fit Score, ATS Score, Recommendation) within
  about 30 seconds of capture.
- Use the **Job Assistant** menu for everything else: generating tailored resumes/cover letters,
  interview prep, sorting, checking email responses.
- The **Config** tab is your scoring rulebook — edit any row any time to change how Claude scores
  or writes for you. If you want finer-grained scoring for a specific type of role (the way a
  power user might add a rule like "score X role type highly when Y"), just add a new row to
  Config with a short label and plain-English instruction — it's picked up automatically on the
  next analysis.

## Have a question, found a bug, or have an idea?

**Job Assistant → Open Claude Chat** for questions about how the tracker works. **Job Assistant
→ Send Feedback** to report a bug or suggest a feature — it goes straight to the developer's
inbox.

## If something breaks

Check **Extensions → Apps Script → Executions** (left sidebar) for an error log of the most
recent runs.
