/***************
 * Job Maverick — Job Search Assistant
 *
 * Tabs:
 * - Jobs
 * - Input
 * - Config
 * - Training
 * - Activity Log
 * - Bugs & Features
 *
 * Input:
 * - Input!A2 = Job Link
 * - Input!B2 = Job Description
 * - Input!C2 = Ready to Analyze checkbox
 *
 * Web App:
 * - doGet supports chunked bookmarklet capture
 ***************/

const CLAUDE_DEFAULT_MODEL = "claude-haiku-4-5-20251001";
const CLAUDE_DEFAULT_TAILORING_MODEL = "claude-fable-5";

const INPUT_SHEET_NAME = "Input";
const JOBS_SHEET_NAME = "Jobs";
const CONFIG_SHEET_NAME = "Config";
const ACTIVITY_LOG_SHEET_NAME = "Activity Log";
const TRAINING_SHEET_NAME = "Training";
const BUGS_FEATURES_SHEET_NAME = "Bugs & Features";

const CALIBRATION_NOTE_HEADER = "Calibration Note";
const CALIBRATION_SUMMARY_KEY = "Calibration Summary";
// Marks a Calibration Note as already ingested into Config. Font color, not
// background — sortJobsByPriority clears cell backgrounds on every auto-sort,
// which would wipe a background-based marker almost immediately.
const CALIBRATION_PROCESSED_FONT_COLOR = "#999999";

const JOBS_HEADER_ROW = 3;   // rows 1-2 are dashboard/metrics; row 3 is column headers
const JOBS_DATA_START_ROW = JOBS_HEADER_ROW + 1;

const NEXT_STEP_SEQUENCE = [
  "Verify company site",
  "Find Contact",
  "Create Resume/Letter",
  "Apply",
  "Paste Domain",
  "Follow Up"
];

const SORT_TRIGGER_HEADERS = [
  "Fit Score",
  "ATS Score",
  "Response Probability",
  "Compensation Fit",
  "Stability",
  "Recommendation",
  "Status",
  "Auto Response",
  "Next Step",
  "Next Step Due Date",
  "Response Date"
];

const VALID_RECOMMENDATIONS = [
  "Apply",
  "Discuss",
  "Apply if Easy",
  "Skip"
];

// Manual-only resolutions for borderline recommendations (Discuss / Apply if Easy),
// never produced by Claude analysis. sortJobsByPriority and the Recommendation
// conditional formatting rule both treat "- Apply"/"- Skip" suffixes as Apply/Skip.
const BORDERLINE_RESOLUTION_RECOMMENDATIONS = [
  "Apply if Easy - Apply",
  "Apply if Easy - Skip",
  "Discuss - Apply",
  "Discuss - Skip"
];

const VALID_STATUSES = [
  "New",
  "Review",
  "Networking",
  "Applied",
  "Skip"
];

const VALID_NEXT_STEPS = [
  "Verify company site",
  "Discuss",
  "Review if easy",
  "Find Contact",
  "Create Resume/Letter",
  "Apply",
  "Paste Domain",
  "Follow Up",
  "Pass"
];

function getOrCreateActivityLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(ACTIVITY_LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(ACTIVITY_LOG_SHEET_NAME);
    log.getRange(1, 1, 1, 7).setValues([["Date", "Company", "Role", "Field", "From", "To", "Job ID"]]);
    log.setFrozenRows(1);
    log.hideColumns(7);
  }
  return log;
}

function getOrCreateBugsFeaturesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BUGS_FEATURES_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(BUGS_FEATURES_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([["Date Added", "Type", "Description", "Status"]]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getBugsFeaturesJson() {
  const sheet = getOrCreateBugsFeaturesSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return jsonResponse({ ok: true, items: [] });
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
  const items = values
    .filter(function(row) { return String(row[2] || "").trim(); }) // has a description
    .map(function(row) {
      return {
        dateAdded: (row[0] instanceof Date)
          ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "M/d/yyyy")
          : String(row[0] || "").trim(),
        type: String(row[1] || "").trim(),
        description: String(row[2] || "").trim(),
        status: String(row[3] || "").trim()
      };
    });

  return jsonResponse({ ok: true, items: items });
}

function logActivityChange(jobId, company, role, fieldName, fromValue, toValue) {
  if (!jobId || !fieldName) return;
  const from = String(fromValue || "").trim();
  const to   = String(toValue  || "").trim();
  if (from === to) return;

  const log = getOrCreateActivityLogSheet();
  const lastRow = log.getLastRow();

  if (lastRow >= 2) {
    const data = log.getRange(2, 1, lastRow - 1, 7).getValues();
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][6]) === String(jobId) && data[i][3] === fieldName) {
        if (String(data[i][5]) === to) return;
        log.getRange(i + 2, 1).setValue(new Date());
        log.getRange(i + 2, 6).setValue(to);
        return;
      }
    }
  }

  log.appendRow([new Date(), company, role, fieldName, from, to, jobId]);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Job Assistant")
    .addItem("View Instructions", "viewInstructions")
    .addItem("Setup Wizard", "showSetupWizard")
    .addItem("Get Bookmarklet", "showBookmarkletDialog")
    .addItem("Set Claude API Key", "setClaudeApiKey")
    .addItem("Set Webhook Token", "setWebhookToken")
    .addSeparator()
    .addItem("Mark Step Complete", "markNextStepComplete")
    .addItem("Sort by Priority", "sortJobsByPriority")
    .addItem("Check Emails Now", "checkJobApplicationEmails")
    .addItem("Check Emails (Extended 30 Days)", "checkJobApplicationEmailsExtended")
    .addSeparator()
    .addItem("Recalculate ATS Score", "recalculateAtsScoreForSelectedRow")
    .addItem("Analyze Resume Diff", "analyzeResumeDiffForSelectedRow")
    .addItem("Update Calibration from Notes", "updateCalibrationFromNotes")
    .addItem("Check for New Claude Models", "checkForNewClaudeModelsInteractive")
    .addSeparator()
    .addItem("Open Claude Chat", "openClaudeSidebar")
    .addToUi();
}

function doGet(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    if (!params.token && !params.jobDescription && !params.action) {
      return ContentService
        .createTextOutput("Job tracker web app is live")
        .setMimeType(ContentService.MimeType.TEXT);
    }

    if (!validateWebhookToken(params.token)) {
      return jsonResponse({
        ok: false,
        error: "Invalid or missing token."
      });
    }

    if (params.action === "getBugsFeatures") {
      return getBugsFeaturesJson();
    }

    if (params.action) {
      return handleChunkedGet(params);
    }

    if (params.jobDescription) {
      return handleSingleGet(params);
    }

    return jsonResponse({
      ok: false,
      error: "No action or jobDescription provided."
    });

  } catch (err) {
    console.error("doGet error:", err.message, err.stack);
    return jsonResponse({
      ok: false,
      error: err.message
    });
  }
}

function doPost(e) {
  try {
    const params = e && e.parameter ? e.parameter : {};
    if (!validateWebhookToken(params.token)) {
      return htmlResponse("Error: invalid token.");
    }

    const jobDescription = String(params.jobDescription || "").trim();
    const jobLink = String(params.jobLink || "").trim();
    const pageTitle = String(params.pageTitle || "").trim();

    if (!jobDescription) {
      return htmlResponse("Error: no job description provided.");
    }

    const enrichedDescription =
      "Captured page title: " + pageTitle + "\n\n" +
      "Captured source: bookmarklet\n\n" +
      jobDescription;

    analyzeAndWriteJob(enrichedDescription, jobLink);

    return htmlResponse("Job added to tracker! This tab will close shortly.");
  } catch (err) {
    console.error("doPost error:", err.message, err.stack);
    return htmlResponse("Error: " + err.message);
  }
}

function htmlResponse(msg) {
  const escaped = msg.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  return HtmlService.createHtmlOutput(
    '<html><head><title>Job Tracker</title></head>' +
    '<body style="font-family:Arial,sans-serif;padding:40px;max-width:500px;margin:0 auto;">' +
    '<h2 style="color:#1a73e8;">Job Tracker</h2>' +
    '<p>' + escaped + '</p>' +
    '<script>setTimeout(function(){window.close();},4000);</script>' +
    '</body></html>'
  );
}

function validateWebhookToken(receivedToken) {
  const expectedToken = PropertiesService
    .getScriptProperties()
    .getProperty("WEBHOOK_TOKEN");

  if (!expectedToken) return false;

  return String(receivedToken || "").trim() === expectedToken;
}

function handleSingleGet(params) {
  const jobLink = String(params.jobLink || "").trim();
  const pageTitle = String(params.pageTitle || "").trim();
  const source = String(params.source || "single-get").trim();
  const jobDescription = String(params.jobDescription || "").trim();

  if (!jobDescription) {
    return jsonResponse({
      ok: false,
      error: "No job description captured."
    });
  }

  const enrichedDescription =
    "Captured page title: " + pageTitle + "\n\n" +
    "Captured source: " + source + "\n\n" +
    jobDescription;

  const finalAnalysis = analyzeAndWriteJob(enrichedDescription, jobLink);

  return jsonResponse({
    ok: true,
    company: finalAnalysis.company,
    role: finalAnalysis.role,
    message: "Job added to tracker."
  });
}

function handleChunkedGet(params) {
  const action = String(params.action || "").trim();
  const captureId = String(params.captureId || "").trim();

  if (!captureId) {
    return jsonResponse({
      ok: false,
      error: "Missing captureId."
    });
  }

  const cache = CacheService.getScriptCache();

  if (action === "start") {
    const meta = {
      jobLink: String(params.jobLink || "").trim(),
      pageTitle: String(params.pageTitle || "").trim(),
      source: String(params.source || "chunked-bookmarklet").trim(),
      totalChunks: Number(params.totalChunks || 0)
    };

    cache.put(captureId + ":meta", JSON.stringify(meta), 21600);

    return jsonResponse({
      ok: true,
      action: "start",
      captureId: captureId
    });
  }

  if (action === "chunk") {
    const index = String(params.index || "").trim();
    const chunk = String(params.chunk || "");

    if (!index) {
      return jsonResponse({
        ok: false,
        error: "Missing chunk index."
      });
    }

    cache.put(captureId + ":chunk:" + index, chunk, 21600);

    return jsonResponse({
      ok: true,
      action: "chunk",
      index: index
    });
  }

  if (action === "finish") {
    const metaText = cache.get(captureId + ":meta");

    if (!metaText) {
      console.error("finish: cache miss for meta, captureId=" + captureId);
      return jsonResponse({
        ok: false,
        error: "Missing capture metadata."
      });
    }

    const meta = JSON.parse(metaText);
    const totalChunks = Number(meta.totalChunks || 0);

    if (!totalChunks) {
      console.error("finish: totalChunks=0 for captureId=" + captureId);
      return jsonResponse({
        ok: false,
        error: "Invalid totalChunks."
      });
    }

    let jobDescription = "";

    for (let i = 0; i < totalChunks; i++) {
      const chunk = cache.get(captureId + ":chunk:" + i);

      if (chunk === null) {
        console.error("finish: missing chunk " + i + " of " + totalChunks);
        return jsonResponse({
          ok: false,
          error: "Missing chunk " + i + "."
        });
      }

      jobDescription += chunk;
    }

    const enrichedDescription =
      "Captured page title: " + meta.pageTitle + "\n\n" +
      "Captured source: " + meta.source + "\n\n" +
      jobDescription;

    const finalAnalysis = analyzeAndWriteJob(enrichedDescription, meta.jobLink);

    return jsonResponse({
      ok: true,
      company: finalAnalysis.company,
      role: finalAnalysis.role,
      message: "Job added to tracker."
    });
  }

  return jsonResponse({
    ok: false,
    error: "Unknown action: " + action
  });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function setClaudeApiKey() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    "Set Claude API Key",
    "Paste your Anthropic Claude API key. It will be stored in Script Properties, not in the sheet.",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const apiKey = result.getResponseText().trim();

  if (!apiKey) {
    ui.alert("No API key entered.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("CLAUDE_API_KEY", apiKey);
  ui.alert("Claude API key saved.");
}

function setWebhookToken() {
  const ui = SpreadsheetApp.getUi();

  const result = ui.prompt(
    "Set Webhook Token",
    "Enter a private token for the browser bookmarklet. Example: rick-jobs-2026",
    ui.ButtonSet.OK_CANCEL
  );

  if (result.getSelectedButton() !== ui.Button.OK) return;

  const token = result.getResponseText().trim();

  if (!token) {
    ui.alert("No token entered.");
    return;
  }

  PropertiesService.getScriptProperties().setProperty("WEBHOOK_TOKEN", token);
  ui.alert("Webhook token saved.");
}

// Builds the ready-to-drag bookmarklet from bookmarklet_template.html (kept as a .html file,
// not .js, so Apps Script treats it as a static asset instead of trying to execute it as
// server code — it's the same browser JS as bookmarklet.js in the repo, just read at runtime
// via HtmlService so this deployment's own URL/token can be substituted in automatically,
// with no hand-editing required).
function getBookmarkletSnippet() {
  const webAppUrl = ScriptApp.getService().getUrl();
  if (!webAppUrl) {
    throw new Error("This script isn't deployed as a web app yet. In the Apps Script editor: Deploy > New deployment > Web app, then try again.");
  }

  const token = PropertiesService.getScriptProperties().getProperty("WEBHOOK_TOKEN") || "";
  if (!token) {
    throw new Error("No webhook token set yet. Run the Setup Wizard first (it generates one automatically), or Job Assistant → Set Webhook Token.");
  }

  // Substitute whole quoted placeholders with JSON.stringify'd values (not raw concatenation)
  // so a stray quote/backslash in either value can't break out of the string literal. Returns
  // the raw JS (no "javascript:" prefix, not URI-encoded) — the dialog builds both a draggable
  // href (needs encodeURIComponent) and a plain-text copy version (needs the raw code) from it.
  const template = HtmlService.createHtmlOutputFromFile("bookmarklet_template").getContent();
  return template
    .split('"__BASE_URL__"').join(JSON.stringify(webAppUrl))
    .split('"__WEBHOOK_TOKEN__"').join(JSON.stringify(token));
}

function showBookmarkletDialog() {
  const ui = SpreadsheetApp.getUi();
  let code;
  try {
    code = getBookmarkletSnippet();
  } catch (e) {
    ui.alert(e.message);
    return;
  }

  const template = HtmlService.createTemplateFromFile("bookmarklet_dialog");
  template.code = code;
  const html = template.evaluate().setWidth(480).setHeight(420);
  ui.showModalDialog(html, "Your Bookmarklet");
}

function analyzeInputJob() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    SpreadsheetApp.getUi().alert("Another analysis is already running. Try again in a moment.");
    return;
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);

    if (!inputSheet) throw new Error("Missing tab: Input");

    const jobLink = String(inputSheet.getRange("A2").getValue()).trim();
    const jobDescription = String(inputSheet.getRange("B2").getValue()).trim();

    if (!jobDescription) {
      SpreadsheetApp.getUi().alert("Paste a job description into Input!B2 first.");
      return;
    }

    const finalAnalysis = analyzeAndWriteJob(jobDescription, jobLink);

    resetInputRow(inputSheet);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Job added: " + finalAnalysis.company + " — " + finalAnalysis.role,
      "Job Assistant", 10
    );

  } catch (err) {
    SpreadsheetApp.getUi().alert("Error: " + err.message);
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function analyzeAndWriteJob(jobDescription, jobLink) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobsSheet = ss.getSheetByName(JOBS_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!jobsSheet) throw new Error("Missing tab: Jobs");
  if (!configSheet) throw new Error("Missing tab: Config");

  const config = getConfigValues(configSheet);
  const resumeText = getResumeText(config["Base Resume URL"] || "");
  const analysis = analyzeJobWithClaude(jobDescription, jobLink, config, resumeText);
  const finalAnalysis = applyMechanicalRules(analysis, jobDescription, jobLink, config);

  writeJobRow(jobsSheet, finalAnalysis, jobLink, jobDescription);
  sortJobsByPriority();

  return finalAnalysis;
}

function resetInputRow(inputSheet) {
  inputSheet.getRange("A2:C2").clearContent();
  inputSheet.getRange("C2").insertCheckboxes();
  inputSheet.getRange("C2").setValue(false);
}

function getConfigValues(configSheet) {
  const values = configSheet.getDataRange().getValues();
  const config = {};

  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0]).trim();
    const value = String(values[i][1]).trim();

    if (key) {
      config[key] = value;
    }
  }

  return config;
}

function getConfigNumber(config, key, fallbackValue) {
  const raw = config && Object.prototype.hasOwnProperty.call(config, key)
    ? config[key]
    : "";

  const number = Number(raw);

  if (isNaN(number)) return fallbackValue;

  return number;
}

// Substitutes {{TOKEN}} placeholders in a Config-sourced instruction string with runtime
// values (e.g. the current ATS-matched keyword list). Keeps all instructional wording in
// Config — this is pure string substitution, no prose.
function fillConfigTemplate(template, replacements) {
  let text = String(template || "");
  Object.keys(replacements || {}).forEach(function(token) {
    text = text.split("{{" + token + "}}").join(replacements[token]);
  });
  return text;
}

// Builds the Drive file name for a deliverable doc (Tailored Resume, Cover Letter) from the
// Config "File Naming Pattern" template, so the naming convention is editable without a code change.
function buildDeliverableFileName(config, company, docType) {
  const pattern = config["File Naming Pattern"] || "{{COMPANY}}_{{TYPE}}";
  return fillConfigTemplate(pattern, { COMPANY: company, TYPE: docType });
}

function getConfigForRuntime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!configSheet) return {};

  return getConfigValues(configSheet);
}

function buildConfigSystemPrompt(config) {
  const userName = config["User Display Name"] || config["User Full Name"] || "the user";
  return (
    "You are a job analysis assistant for " + userName + "'s job tracker.\n\n" +
    "Use the Config values below as the source of truth for all scoring rules, " +
    userName + "'s background and preferences, calibration examples, and column definitions. " +
    "Do not invent rules or defaults that conflict with Config.\n\n" +
    "CONFIG VALUES:\n" +
    Object.keys(config)
      .map(key => "- " + key + ": " + config[key])
      .join("\n")
  );
}

function analyzeJobWithClaude(jobDescription, jobLink, config, resumeText) {
  const apiKey = PropertiesService
    .getScriptProperties()
    .getProperty("CLAUDE_API_KEY");

  if (!apiKey) {
    throw new Error("Missing Claude API key. Use Job Assistant → Set Claude API Key.");
  }

  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const toolSchema = {
    name: "job_analysis",
    description: "Structured job analysis output for the candidate's job tracker",
    input_schema: {
      type: "object",
      additionalProperties: false,
      required: [
        "company",
        "role",
        "fit_score",
        "ats_score",
        "response_probability",
        "compensation_fit",
        "stability",
        "recommendation",
        "status",
        "referral_contact",
        "resume_version_used",
        "next_step",
        "industry_type",
        "sector"
      ],
      properties: {
        company: { type: "string" },
        role: { type: "string" },
        fit_score: { type: "integer" },
        ats_score: { type: "integer" },
        response_probability: { type: "string" },
        compensation_fit: { type: "string", description: "Compensation range verbatim from posting (e.g. '$120K-$150K OTE', '$95K base + equity'). If not stated, prefix with 'Estimated:' and infer from role/level/market." },
        stability: { type: "string" },
        recommendation: {
          type: "string",
          enum: ["Apply", "Discuss", "Apply if Easy", "Skip"]
        },
        status: {
          type: "string",
          enum: ["New", "Review", "Applied", "Networking", "Skip"]
        },
        referral_contact: { type: "string" },
        resume_version_used: { type: "string" },
        next_step: {
          type: "string",
          enum: [
            "Verify company site",
            "Discuss",
            "Review if easy",
            "Find Contact",
            "Create Resume/Letter",
            "Apply",
            "Paste Domain",
            "Follow Up",
            "Pass"
          ]
        },
        industry_type: { type: "string" },
        sector: { type: "string" }
      }
    }
  };

  const userName = config["User Display Name"] || config["User Full Name"] || "the user";
  const userPrompt =
    "Analyze the following job for the tracker.\n\n" +
    "Mechanical constraints:\n" +
    "- Recommendation can only be: Apply, Discuss, Apply if Easy, or Skip.\n" +
    "- Do not output Apply Anyway. Apply Anyway is manual only.\n" +
    "- Auto Response is email-monitored — do not generate it.\n" +
    "- Response Probability ignores referrals. Never output with-referral probability.\n" +
    "- Do not penalize for company-site verification during initial LinkedIn scoring.\n" +
    "- Preserve all compensation details exactly when listed.\n" +
    "- Be conservative. Do not inflate weak matches.\n" +
    "- Company and Role must not be blank. Use best extraction from title or text.\n\n" +
    "Scoring guidance:\n" +
    "- Job title has low weight — LinkedIn already filters by title. Score from requirements, must-have gaps, domain fit, product/tool fit, compensation, remote/location fit, and time-worthiness.\n" +
    "- Apply = clearly worth tailored resume/cover-letter time.\n" +
    "- Apply if Easy = decent role, not worth much tailoring.\n" +
    "- Discuss = non-standard, interesting, or edge-case worth manual conversation.\n" +
    "- Skip = major must-have gaps, low comp, poor location fit, or low response likelihood.\n\n" +
    "Next Step rules:\n" +
    "- Apply → 'Verify company site'\n" +
    "- Discuss → 'Discuss'\n" +
    "- Apply if Easy → 'Review if easy'\n" +
    "- Skip → 'Pass'\n\n" +
    "Status rules:\n" +
    "- Discuss → 'Review'; Skip → 'Skip'; otherwise 'New'.\n\n" +
    "Resume for ATS scoring:\n" + (resumeText ? resumeText.substring(0, 8000) : "Use " + userName + "'s background from Config.") + "\n\n" +
    "Job link:\n" + (jobLink || "No link provided") + "\n\n" +
    "Job description:\n" + jobDescription;

  const payload = {
    model: model,
    max_tokens: 4096,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "job_analysis" },
    messages: [
      { role: "user", content: userPrompt }
    ]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  const statusCode = response.getResponseCode();
  const body = response.getContentText();

  if (statusCode < 200 || statusCode >= 300) {
    throw new Error("Claude API error: " + statusCode + " — " + body);
  }

  const data = JSON.parse(body);
  const result = extractClaudeToolResult(data);

  if (!result) {
    throw new Error("No tool result returned from Claude. Response: " + body);
  }

  return result;
}

function extractClaudeToolResult(responseJson) {
  if (!responseJson.content) return null;

  for (const block of responseJson.content) {
    if (block.type === "tool_use" && block.name === "job_analysis") {
      return block.input;
    }
  }

  return null;
}

function applyMechanicalRules(analysis, jobDescription, jobLink, config) {
  const recommendation = sanitizeRecommendation(analysis.recommendation);

  let nextStep = "";

  if (recommendation === "Apply") {
    nextStep = "Verify company site";
  } else if (recommendation === "Discuss") {
    nextStep = "Discuss";
  } else if (recommendation === "Apply if Easy") {
    nextStep = "Review if easy";
  } else {
    nextStep = "Pass";
  }

  const identity = enforceJobIdentity(
    analysis.company,
    analysis.role,
    jobDescription,
    jobLink
  );

  return {
    company: identity.company,
    role: identity.role,
    fit_score: clampNumber(analysis.fit_score, 1, 10),
    ats_score: clampNumber(analysis.ats_score, 1, 100),
    response_probability: analysis.response_probability || "",
    compensation_fit: analysis.compensation_fit || "",
    stability: analysis.stability || "",
    recommendation: recommendation,
    status: sanitizeStatus(analysis.status, recommendation),
    company_response: "",
    referral_contact: analysis.referral_contact || "",
    notes: identity.note,
    resume_version_used: resumeHyperlinkFormula(
      analysis.resume_version_used || config["Base Resume URL"] || "",
      "Base Resume"
    ),
    next_step: nextStep,
    next_step_due_date: getDueDateForStep(nextStep, new Date(), config),
    industry_type: analysis.industry_type || "",
    sector: analysis.sector || ""
  };
}

function enforceJobIdentity(company, role, jobDescription, jobLink) {
  let finalCompany = String(company || "").trim();
  let finalRole = String(role || "").trim();
  let note = "";

  if (!finalCompany || finalCompany.toLowerCase() === "unknown") {
    finalCompany = extractCompanyFallback(jobDescription, jobLink);
  }

  if (!finalRole || finalRole.toLowerCase() === "unknown") {
    finalRole = extractRoleFallback(jobDescription);
  }

  if (!finalCompany) {
    finalCompany = "Unknown Company";
    note = appendNote(note, "Company could not be extracted automatically; review manually.");
  }

  if (!finalRole) {
    finalRole = "Unknown Role";
    note = appendNote(note, "Role could not be extracted automatically; review manually.");
  }

  return {
    company: finalCompany,
    role: finalRole,
    note: note
  };
}

function extractCompanyFallback(jobDescription, jobLink) {
  const text = String(jobDescription || "");

  const titleMatch = text.match(/Captured page title:\s*(.+)/i);
  if (titleMatch && titleMatch[1]) {
    const title = titleMatch[1].trim();

    const atMatch = title.match(/\bat\s+(.+?)(\s*\||\s*-|$)/i);
    if (atMatch && atMatch[1]) {
      return cleanupExtractedText(atMatch[1]);
    }
  }

  const companyLineMatch = text.match(/Company\s*[:\-]\s*(.+)/i);
  if (companyLineMatch && companyLineMatch[1]) {
    return cleanupExtractedText(companyLineMatch[1]);
  }

  const clientMatch = text.match(/Our client\s+is\s+(.+?)(\.|\n)/i);
  if (clientMatch && clientMatch[1]) {
    return cleanupExtractedText(clientMatch[1]);
  }

  const linkText = String(jobLink || "").toLowerCase();

  if (linkText.includes("linkedin")) return "LinkedIn Posting";
  if (linkText.includes("greenhouse")) return "Greenhouse Posting";
  if (linkText.includes("ashby")) return "Ashby Posting";
  if (linkText.includes("lever")) return "Lever Posting";
  if (linkText.includes("workday")) return "Workday Posting";

  return "";
}

function extractRoleFallback(jobDescription) {
  const text = String(jobDescription || "");

  const titleMatch = text.match(/Captured page title:\s*(.+)/i);
  if (titleMatch && titleMatch[1]) {
    let title = titleMatch[1].trim();
    title = title.replace(/\s*\|\s*LinkedIn.*$/i, "");
    title = title.replace(/\s*-\s*LinkedIn.*$/i, "");

    const atIndex = title.toLowerCase().indexOf(" at ");
    if (atIndex > 0) {
      return cleanupExtractedText(title.substring(0, atIndex));
    }

    if (title.length <= 100) {
      return cleanupExtractedText(title);
    }
  }

  const roleLineMatch = text.match(/(?:Role|Job Title|Title)\s*[:\-]\s*(.+)/i);
  if (roleLineMatch && roleLineMatch[1]) {
    return cleanupExtractedText(roleLineMatch[1]);
  }

  const salesEngineerMatch = text.match(/([A-Za-z\s\/,-]*Sales Engineer[A-Za-z\s\/,-]*)/i);
  if (salesEngineerMatch && salesEngineerMatch[1]) {
    return cleanupExtractedText(salesEngineerMatch[1]);
  }

  const solutionsEngineerMatch = text.match(/([A-Za-z\s\/,-]*Solutions Engineer[A-Za-z\s\/,-]*)/i);
  if (solutionsEngineerMatch && solutionsEngineerMatch[1]) {
    return cleanupExtractedText(solutionsEngineerMatch[1]);
  }

  return "";
}

function cleanupExtractedText(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .replace(/^[\-–—:|]+/, "")
    .replace(/[\-–—:|]+$/, "")
    .trim()
    .substring(0, 120);
}

function appendNote(existing, newNote) {
  const current = String(existing || "").trim();
  const add = String(newNote || "").trim();

  if (!current) return add;
  if (!add) return current;

  return current + " " + add;
}

function sanitizeRecommendation(value) {
  const text = String(value || "").trim();

  if (text === "Apply") return "Apply";
  if (text === "Discuss") return "Discuss";
  if (text === "Apply if Easy") return "Apply if Easy";
  if (text === "Skip") return "Skip";

  if (text.toLowerCase() === "apply anyway") return "Skip";

  return "Skip";
}

function sanitizeStatus(statusValue, recommendationValue) {
  if (recommendationValue === "Skip") return "Skip";
  if (recommendationValue === "Discuss") return "Review";

  const text = String(statusValue || "").trim();

  if (["New", "Review", "Applied", "Networking", "Skip"].indexOf(text) !== -1) {
    return text;
  }

  return "New";
}

function clampNumber(value, min, max) {
  const number = Number(value);

  if (isNaN(number)) return "";

  return Math.max(min, Math.min(max, Math.round(number)));
}

function extractMaxProbabilityPercent(value) {
  const text = String(value || "");
  const matches = text.match(/\d+(\.\d+)?/g);

  if (!matches) return 0;

  return Math.max.apply(null, matches.map(Number));
}

function normalizeNextStep(value) {
  const text = String(value || "").trim().toLowerCase();

  if (!text) return "";

  if (text === "pass" || text === "skip") return "Pass";
  if (text.includes("closed")) return "Closed";
  if (text.includes("reject")) return "Closed";
  if (text.includes("discuss")) return "Discuss";
  if (text.includes("verify")) return "Verify company site";
  if (text.includes("review")) return "Review if easy";
  if (text.includes("find")) return "Find Contact";
  if (text.includes("resume") || text.includes("letter")) return "Create Resume/Letter";
  if (text.includes("paste")) return "Paste Domain";
  if (text.includes("apply")) return "Apply";
  if (text.includes("follow")) return "Follow Up";

  return "";
}

function isActionNextStep(value) {
  const normalized = normalizeNextStep(value);

  return [
    "Verify company site",
    "Review if easy",
    "Discuss",
    "Find Contact",
    "Create Resume/Letter",
    "Apply",
    "Paste Domain",
    "Follow Up"
  ].indexOf(normalized) !== -1;
}

function getDueDateForStep(step, baseDate, config) {
  const normalized = normalizeNextStep(step);
  const startDate = baseDate ? new Date(baseDate) : new Date();
  const runtimeConfig = config || getConfigForRuntime();

  if (!normalized || normalized === "Pass" || normalized === "Closed") return "";

  if (normalized === "Verify company site") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Initial Apply Due Days", 7)));
  }

  if (normalized === "Review if easy") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Initial Apply If Easy Due Days", 7)));
  }

  if (normalized === "Discuss") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Initial Discuss Due Days", 7)));
  }

  if (normalized === "Find Contact") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Find Contact Due Days", 2)));
  }

  if (normalized === "Create Resume/Letter") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Create Resume Letter Due Days", 3)));
  }

  if (normalized === "Apply") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Apply Due Days", 4)));
  }

  if (normalized === "Paste Domain") {
    return formatDate(addDays(startDate, 0));
  }

  if (normalized === "Follow Up") {
    return formatDate(addDays(startDate, getConfigNumber(runtimeConfig, "Follow Up Due Days", 7)));
  }

  return "";
}

function addDays(date, days) {
  const copy = new Date(date);
  copy.setDate(copy.getDate() + Number(days || 0));
  return copy;
}

function formatDate(date) {
  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function writeJobRow(jobsSheet, analysis, jobLink, jobDescription) {
  const lastColumn = jobsSheet.getLastColumn();
  const headers = jobsSheet.getRange(JOBS_HEADER_ROW, 1, 1, lastColumn).getValues()[0];

  const duplicate = findDuplicateJobAcrossSheets(
    analysis.company,
    analysis.role,
    jobLink
  );

  if (duplicate) {
    throw new Error(
      "Duplicate job found in " + duplicate.sheetName + " row " + duplicate.rowNumber + ". Job was not added."
    );
  }

  const nextId = getNextJobId();

  const rowObj = {
    "ID": nextId,
    "Company": analysis.company,
    "Role": analysis.role,
    "Job Link": jobLink,
    "Fit Score": analysis.fit_score,
    "ATS Score": analysis.ats_score,
    "Response Probability": analysis.response_probability,
    "Compensation Fit": analysis.compensation_fit,
    "Stability": analysis.stability,
    "Recommendation": analysis.recommendation,
    "Status": analysis.status,
    "Referral / Contact": analysis.referral_contact,
    "Next Step": analysis.next_step,
    "Next Step Due Date": analysis.next_step_due_date,
    "Industry Type": analysis.industry_type,
    "Sector": analysis.sector,
    "Last Updated": formatDate(new Date()),
    "Response Date": "",
    "Auto Response": analysis.company_response,
    "ATS / Platform": inferAtsPlatform(jobLink, jobDescription),
    "Job Description": String(jobDescription || "").substring(0, 45000),
    "Job Discription": String(jobDescription || "").substring(0, 45000)
  };

  const normalizedRowObj = {};
  Object.keys(rowObj).forEach(k => { normalizedRowObj[normalizeHeader(k)] = rowObj[k]; });
  const row = headers.map(header => normalizedRowObj[normalizeHeader(header)] ?? "");
  const targetRow = getFirstEmptyJobRow(jobsSheet, headers);

  jobsSheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
  jobsSheet.getBandings().forEach(function(b) { b.remove(); });
}

function getNextJobId() {
  const now = new Date();
  const pad = n => String(n).padStart(2, "0");
  const ms = String(now.getMilliseconds()).padStart(3, "0");
  return (
    now.getFullYear() +
    pad(now.getMonth() + 1) +
    pad(now.getDate()) +
    "-" +
    pad(now.getHours()) +
    pad(now.getMinutes()) +
    pad(now.getSeconds()) +
    ms
  );
}

function fixDuplicateIds() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { SpreadsheetApp.getUi().alert("No jobs to check."); return; }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = getColumnIndexByHeader(headers, "ID");
  if (!idCol) { SpreadsheetApp.getUi().alert("No ID column found."); return; }

  const idRange = sheet.getRange(JOBS_DATA_START_ROW, idCol, lastRow - JOBS_HEADER_ROW, 1);
  const idValues = idRange.getValues();

  const seen = {};
  const fixes = [];

  for (let i = 0; i < idValues.length; i++) {
    const id = String(idValues[i][0] || "").trim();
    if (!id) continue;
    if (seen[id] === undefined) {
      seen[id] = 1;
    } else {
      seen[id]++;
      const suffix = String.fromCharCode(64 + seen[id]); // B, C, D...
      const newId = id + suffix;
      fixes.push({ row: i + 2, newId: newId });
      idValues[i][0] = newId;
    }
  }

  if (fixes.length === 0) {
    SpreadsheetApp.getActiveSpreadsheet().toast("No duplicate IDs found.", "Job Assistant", 8);
    return;
  }

  idRange.setValues(idValues);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    fixes.length + " duplicate ID(s) fixed:\n\n" +
    fixes.map(f => "Row " + f.row + " → " + f.newId).join("\n"),
    "Job Assistant", 15
  );
}

function backfillMissingJobIds() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);

  if (!sheet) throw new Error("Missing tab: Jobs");

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = getColumnIndexByHeader(headers, "ID");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");

  if (!idCol) {
    SpreadsheetApp.getUi().alert("No ID column found.");
    return;
  }

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert("No jobs to update.");
    return;
  }

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, sheet.getLastColumn()).getValues();

  const today = (() => {
    const d = new Date();
    const pad = n => String(n).padStart(2, "0");
    return d.getFullYear() + pad(d.getMonth() + 1) + pad(d.getDate());
  })();

  const isTimestampId = id => /^\d{8}-/.test(id);

  let updated = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const existingId = String(values[i][idCol - 1] || "").trim();
    const company = companyCol ? String(values[i][companyCol - 1] || "").trim() : "";
    const role = roleCol ? String(values[i][roleCol - 1] || "").trim() : "";

    if (!(company || role)) continue;

    if (!existingId) {
      // Missing ID — assign timestamp placeholder
      sheet.getRange(rowNumber, idCol).setValue(today + "-000000");
      updated++;
    } else if (!isTimestampId(existingId)) {
      // Old integer ID — convert to today-NNNNNN format
      const padded = String(existingId).padStart(6, "0");
      sheet.getRange(rowNumber, idCol).setValue(today + "-" + padded);
      updated++;
    }
  }

  sortJobsByPriority();

  SpreadsheetApp.getActiveSpreadsheet().toast("Updated " + updated + " job IDs to timestamp format.", "Job Assistant", 10);
}

function findDuplicateJobAcrossSheets(company, role, jobLink) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = [
    JOBS_SHEET_NAME
  ];

  for (const sheetName of sheetNames) {
    const sheet = ss.getSheetByName(sheetName);

    if (!sheet) continue;

    const result = findDuplicateJobRowInSheet(sheet, company, role, jobLink);

    if (result) {
      return {
        sheetName: sheetName,
        rowNumber: result
      };
    }
  }

  return null;
}

function findDuplicateJobRowInSheet(sheet, company, role, jobLink) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2 || lastCol < 1) return null;

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const jobLinkCol = getColumnIndexByHeader(headers, "Job Link");

  if (!companyCol || !roleCol) return null;

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, lastCol).getValues();

  const newCompany = normalizeForDuplicate(company);
  const newRole = normalizeForDuplicate(role);
  const newLink = normalizeJobLink(jobLink);

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    const existingCompany = normalizeForDuplicate(row[companyCol - 1]);
    const existingRole = normalizeForDuplicate(row[roleCol - 1]);
    const existingLink = jobLinkCol ? normalizeJobLink(row[jobLinkCol - 1]) : "";

    if (newLink && existingLink && newLink === existingLink) {
      return i + 2;
    }

    if (newCompany && newRole && existingCompany === newCompany && existingRole === newRole) {
      return i + 2;
    }
  }

  return null;
}

function scanForDuplicates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 3) { SpreadsheetApp.getUi().alert("Not enough rows to check for duplicates."); return; }

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const jobLinkCol = getColumnIndexByHeader(headers, "Job Link");

  if (!companyCol || !roleCol) { SpreadsheetApp.getUi().alert("Company and Role columns required."); return; }

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, lastCol).getValues();
  const duplicateRows = new Set();
  const findings = [];

  for (let i = 0; i < values.length; i++) {
    if (duplicateRows.has(i)) continue;
    const compA = normalizeForDuplicate(values[i][companyCol - 1]);
    const roleA = normalizeForDuplicate(values[i][roleCol - 1]);
    const linkA = jobLinkCol ? normalizeJobLink(values[i][jobLinkCol - 1]) : "";
    if (!compA && !linkA) continue;

    for (let j = i + 1; j < values.length; j++) {
      const compB = normalizeForDuplicate(values[j][companyCol - 1]);
      const roleB = normalizeForDuplicate(values[j][roleCol - 1]);
      const linkB = jobLinkCol ? normalizeJobLink(values[j][jobLinkCol - 1]) : "";

      const linkMatch = linkA && linkB && linkA === linkB;
      const nameMatch = compA && roleA && compA === compB && roleA === roleB;

      if (linkMatch || nameMatch) {
        duplicateRows.add(j);
        const label = String(values[i][companyCol - 1] || "").trim() + " — " + String(values[i][roleCol - 1] || "").trim();
        findings.push("Rows " + (i + 2) + " & " + (j + 2) + ": " + label + (linkMatch ? " (same link)" : " (same name)"));
      }
    }
  }

  if (findings.length === 0) {
    SpreadsheetApp.getUi().alert("No duplicates found.");
  } else {
    SpreadsheetApp.getUi().alert(
      findings.length + " duplicate pair(s) found:\n\n" + findings.join("\n") +
      "\n\nReview these rows and delete the one you want to keep fewer of."
    ); // deliberately blocking — requires the user to review and act, unlike a routine "Done!"
  }
}

function normalizeForDuplicate(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

function normalizeJobLink(value) {
  const text = String(value || "").trim().toLowerCase();

  if (!text) return "";
  if (text.includes("url not pasted")) return "";
  if (text.includes("no link")) return "";
  if (text === "linkedin easy apply") return "";

  const linkedInJobId = extractLinkedInJobId(text);

  if (linkedInJobId) {
    return "linkedin-job-" + linkedInJobId;
  }

  return text
    .replace(/^https?:\/\//, "")
    .replace(/^www\./, "")
    .replace(/[?#].*$/, "")
    .replace(/\/$/, "");
}

function extractLinkedInJobId(value) {
  const text = String(value || "");

  const viewMatch = text.match(/\/jobs\/view\/(\d+)/i);
  if (viewMatch && viewMatch[1]) return viewMatch[1];

  const currentJobMatch = text.match(/[?&]currentJobId=(\d+)/i);
  if (currentJobMatch && currentJobMatch[1]) return currentJobMatch[1];

  return "";
}

function inferAtsPlatform(jobLink, notes) {
  const combined = (String(jobLink || "") + " " + String(notes || "")).toLowerCase();

  if (combined.includes("ashby")) return "Ashby";
  if (combined.includes("greenhouse")) return "Greenhouse";
  if (combined.includes("lever")) return "Lever";
  if (combined.includes("workday")) return "Workday";
  if (combined.includes("smartrecruiters")) return "SmartRecruiters";
  if (combined.includes("icims")) return "iCIMS";
  if (combined.includes("taleo")) return "Taleo";
  if (combined.includes("successfactors")) return "SAP SuccessFactors";
  if (combined.includes("bamboohr")) return "BambooHR";
  if (combined.includes("jazzhr") || combined.includes("jazz.co")) return "JazzHR";
  if (combined.includes("rippling")) return "Rippling";
  if (combined.includes("jobvite")) return "Jobvite";
  if (combined.includes("recruitee")) return "Recruitee";
  if (combined.includes("comeet")) return "Comeet";
  if (combined.includes("dover.com")) return "Dover";
  if (combined.includes("linkedin") && combined.includes("easy apply")) return "LinkedIn Easy Apply";
  if (combined.includes("easy apply")) return "LinkedIn Easy Apply";
  if (combined.includes("linkedin")) return "LinkedIn";

  if (String(jobLink || "").trim()) return "Unavailable";
  return "";
}

function getFirstEmptyJobRow(jobsSheet, headers) {
  const maxRows = jobsSheet.getMaxRows();
  const companyCol = getColumnIndexByHeader(headers, "Company") || 1;
  const companyValues = jobsSheet.getRange(JOBS_DATA_START_ROW, companyCol, maxRows - JOBS_HEADER_ROW, 1).getValues();

  for (let i = 0; i < companyValues.length; i++) {
    const company = String(companyValues[i][0]).trim();

    if (!company) {
      return i + JOBS_DATA_START_ROW;
    }
  }

  return maxRows + 1;
}


function installAutoAnalyzeTrigger() {
  const ss = SpreadsheetApp.getActive();

  const triggers = ScriptApp.getProjectTriggers();

  for (const trigger of triggers) {
    if (
      trigger.getHandlerFunction() === "handleEdit" ||
      trigger.getHandlerFunction() === "handleInputEdit"
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  }

  ScriptApp.newTrigger("handleEdit")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast("Auto-analyze, checkbox, ID, duplicate-check, due-date, validation, and lifecycle priority-sort trigger installed.", "Job Assistant", 10);
}

function handleEdit(e) {
  try {
    if (!e || !e.range) return;

    const range = e.range;
    const sheet = range.getSheet();

    if (sheet.getName() === INPUT_SHEET_NAME) {
      handleInputSheetEdit(e);
      return;
    }

    if (sheet.getName() === JOBS_SHEET_NAME) {
      handleJobsSheetEdit(e);
      return;
    }

  } catch (err) {
    console.error(err);
  }
}

function handleInputSheetEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();

  const row = range.getRow();
  const col = range.getColumn();

  if (row !== 2 || col !== 3) return;

  const ready = sheet.getRange("C2").getValue() === true;
  if (!ready) return;

  const jobDescription = String(sheet.getRange("B2").getValue()).trim();

  if (!jobDescription) {
    sheet.getRange("C2").setValue(false);
    SpreadsheetApp.getUi().alert("Paste a job description into Input!B2 first.");
    return;
  }

  analyzeInputJob();

  sheet.getRange("C2").setValue(false);
}

function handleJobsSheetEdit(e) {
  const sheet = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (row <= JOBS_HEADER_ROW) return;

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (col === recommendationCol) {
    const rawRecommendation = String(e.range.getValue() || "").trim();
    const recommendation = rawRecommendation.toLowerCase();

    if (recommendation === "apply anyway") {
      applyApplyAnywayOverride(sheet, row, headers);
    }

    if (recommendation === "discuss") {
      applyDiscussOverride(sheet, row, headers);
    }

    if (recommendation === "apply if easy - apply" || recommendation === "discuss - apply") {
      canonicalizeBorderlineRecommendation(sheet, row, recommendationCol, rawRecommendation);
      applyBorderlineResolutionOverride(sheet, row, headers, "Apply");
    }

    if (recommendation === "apply if easy - skip" || recommendation === "discuss - skip") {
      canonicalizeBorderlineRecommendation(sheet, row, recommendationCol, rawRecommendation);
      applyBorderlineResolutionOverride(sheet, row, headers, "Skip");
    }
  }

  if (col === statusCol) {
    const originalStatus = String(e.range.getValue() || "").trim();
    const formattedStatus = formatStatusForManualEntry(originalStatus);

    if (formattedStatus && formattedStatus !== originalStatus) {
      e.range.setValue(formattedStatus);
    }

    const nextStepForStatus = nextStepFromStatus(formattedStatus || originalStatus);
    if (nextStepForStatus && nextStepCol) {
      sheet.getRange(row, nextStepCol).setValue(nextStepForStatus);
      ensureDueDateForActionRow(sheet, row, headers);
    }

    if (lastUpdatedCol) {
      sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
    }

    try {
      const idCol = getColumnIndexByHeader(headers, "ID");
      const companyCol2 = getColumnIndexByHeader(headers, "Company");
      const roleCol2 = getColumnIndexByHeader(headers, "Role");
      const rowData2 = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const jobId = idCol ? String(rowData2[idCol - 1] || "") : "";
      const company2 = companyCol2 ? String(rowData2[companyCol2 - 1] || "") : "";
      const role2 = roleCol2 ? String(rowData2[roleCol2 - 1] || "") : "";
      logActivityChange(jobId, company2, role2, "Status", e.oldValue || "", formattedStatus || originalStatus);
    } catch(logErr) { console.warn("Activity log error: " + logErr.message); }
  }

  if (col === nextStepCol) {
    ensureDueDateForActionRow(sheet, row, headers);

    if (lastUpdatedCol) {
      sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
    }

    try {
      const idCol = getColumnIndexByHeader(headers, "ID");
      const companyCol2 = getColumnIndexByHeader(headers, "Company");
      const roleCol2 = getColumnIndexByHeader(headers, "Role");
      const rowData2 = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const jobId = idCol ? String(rowData2[idCol - 1] || "") : "";
      const company2 = companyCol2 ? String(rowData2[companyCol2 - 1] || "") : "";
      const role2 = roleCol2 ? String(rowData2[roleCol2 - 1] || "") : "";
      logActivityChange(jobId, company2, role2, "Next Step", e.oldValue || "", String(e.range.getValue() || ""));
    } catch(logErr) { console.warn("Activity log error: " + logErr.message); }
  }

  if (isPrioritySortColumn(headers, col)) {
    ensureDueDateForActionRow(sheet, row, headers);
  }

  const jobLinkCol2 = getColumnIndexByHeader(headers, "Job Link");
  if (col === jobLinkCol2) {
    const atsPlatformCol = getColumnIndexByHeader(headers, "ATS / Platform");
    if (atsPlatformCol) {
      const rowData3 = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const jobDescCol2 = getColumnIndexByHeader(headers, "Job Description");
      const jobDesc2 = jobDescCol2 ? String(rowData3[jobDescCol2 - 1] || "") : "";
      const newLink = String(e.range.getValue() || "");
      const platform = inferAtsPlatform(newLink, jobDesc2);
      if (platform) sheet.getRange(row, atsPlatformCol).setValue(platform);
    }
  }

  // Log all other manual edits not already captured above
  const SKIP_LOG_COLS = new Set(["ID", "Last Updated", "Auto Response", "Status", "Next Step"]);
  const headerName = String(headers[col - 1] || "").trim();
  if (headerName && !headerName.startsWith("__") && !SKIP_LOG_COLS.has(headerName)) {
    try {
      const idCol = getColumnIndexByHeader(headers, "ID");
      const companyCol3 = getColumnIndexByHeader(headers, "Company");
      const roleCol3 = getColumnIndexByHeader(headers, "Role");
      const rowDataLog = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
      const jobId = idCol ? String(rowDataLog[idCol - 1] || "") : "";
      const company3 = companyCol3 ? String(rowDataLog[companyCol3 - 1] || "") : "";
      const role3 = roleCol3 ? String(rowDataLog[roleCol3 - 1] || "") : "";
      if (jobId) {
        const oldVal = String(e.oldValue !== undefined ? e.oldValue : "").trim();
        const newVal = String(e.range.getValue() || "").trim();
        logActivityChange(jobId, company3, role3, headerName, oldVal, newVal);
      }
    } catch(logErr) { console.warn("Activity log error: " + logErr.message); }
  }
}

function formatStatusForManualEntry(statusValue) {
  const text = String(statusValue || "").trim();
  const lower = text.toLowerCase();

  if (!text) return text;

  // Only append date to bare status words — leave "Applied - <date>" patterns untouched
  if (lower === "applied") {
    return "Applied - " + formatDate(new Date());
  }

  if (lower === "rejected") {
    return "Rejected - " + formatDate(new Date());
  }

  if (lower === "closed") {
    return "Closed - " + formatDate(new Date());
  }

  return text;
}

function nextStepFromStatus(statusValue) {
  const lower = String(statusValue || "").toLowerCase().trim();
  if (lower.startsWith("applied")) return "Follow Up";
  if (lower.startsWith("rejected") || lower.startsWith("closed") || lower === "skip") return "Pass";
  if (lower === "networking") return "Find Contact";
  return null;
}

function isPrioritySortColumn(headers, editedColumn) {
  for (const header of SORT_TRIGGER_HEADERS) {
    const col = getColumnIndexByHeader(headers, header);

    if (col && col === editedColumn) {
      return true;
    }
  }

  return false;
}

function ensureDueDateForActionRow(sheet, row, headers) {
  const config = getConfigForRuntime();

  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (!nextStepCol || !dueDateCol) return;

  const nextStep = sheet.getRange(row, nextStepCol).getValue();
  const dueDate = sheet.getRange(row, dueDateCol).getValue();
  const lastUpdated = lastUpdatedCol ? sheet.getRange(row, lastUpdatedCol).getValue() : new Date();

  if (isActionNextStep(nextStep) && !dueDate) {
    sheet.getRange(row, dueDateCol).setValue(getDueDateForStep(nextStep, lastUpdated || new Date(), config));
  }

  if (!isActionNextStep(nextStep) && dueDate && normalizeNextStep(nextStep) === "Pass") {
    sheet.getRange(row, dueDateCol).clearContent();
  }
}

function advanceNextStepForRow(sheet, row, headers) {
  const config = getConfigForRuntime();

  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");
  const idCol = getColumnIndexByHeader(headers, "ID");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");

  if (!nextStepCol) return;

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const currentStep = normalizeNextStep(rowData[nextStepCol - 1]);
  const nextStep = getNextStep(currentStep);

  sheet.getRange(row, nextStepCol).setValue(nextStep);

  if (dueDateCol) {
    sheet.getRange(row, dueDateCol).setValue(getDueDateForStep(nextStep, new Date(), config));
  }

  if (lastUpdatedCol) {
    sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
  }

  try {
    const jobId = idCol ? String(rowData[idCol - 1] || "") : "";
    const company = companyCol ? String(rowData[companyCol - 1] || "") : "";
    const role = roleCol ? String(rowData[roleCol - 1] || "") : "";
    logActivityChange(jobId, company, role, "Next Step", currentStep, nextStep);
  } catch(e) { console.warn("Activity log error: " + e.message); }
}

function confirmRowAction(sheet, row, taskName) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const rowData = sheet.getRange(row, 1, 1, lastCol).getValues()[0];
  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role = roleCol ? String(rowData[roleCol - 1] || "").trim() : "";
  const label = [company, role].filter(Boolean).join(" — ") || "row " + row;
  const ui = SpreadsheetApp.getUi();
  return ui.alert(taskName, label, ui.ButtonSet.OK_CANCEL) === ui.Button.OK;
}

function recalculateAtsScoreForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }
  const row = sheet.getActiveCell().getRow();
  if (row <= JOBS_HEADER_ROW) { SpreadsheetApp.getUi().alert("Select a job row first."); return; }
  if (!confirmRowAction(sheet, row, "Recalculate ATS Score")) return;
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  SpreadsheetApp.getActiveSpreadsheet().toast("Recalculating ATS score...", "Job Assistant", 60);
  recalculateAtsScore(sheet, row, headers);
  SpreadsheetApp.getActiveSpreadsheet().toast("Done.", "Job Assistant", 3);
}

function analyzeResumeDiffForSelectedRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }
  const row = sheet.getActiveCell().getRow();
  if (row <= JOBS_HEADER_ROW) { SpreadsheetApp.getUi().alert("Select a job row first."); return; }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];
  const config = getConfigForRuntime();

  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");

  const tailoredUrl = tailoredDraftCol
    ? extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowValues[tailoredDraftCol - 1])
    : "";
  if (!tailoredUrl) {
    SpreadsheetApp.getUi().alert("No tailored resume found for this row. Generate a tailored resume first.");
    return;
  }

  const baseUrl = config["Base Resume URL"] || "";
  if (!baseUrl) {
    SpreadsheetApp.getUi().alert("No Base Resume URL in Config.");
    return;
  }

  const company = companyCol ? String(rowValues[companyCol - 1] || "").trim() : "Unknown";
  const role = roleCol ? String(rowValues[roleCol - 1] || "").trim() : "Unknown";

  ss.toast("Reading both resumes and analyzing differences — this may take 20–30 seconds...", "Job Assistant", 60);

  const baseText = getResumeText(baseUrl);
  const tailoredText = getResumeText(tailoredUrl);

  if (!baseText) { ss.toast("", "", 1); SpreadsheetApp.getUi().alert("Could not read base resume."); return; }
  if (!tailoredText) { ss.toast("", "", 1); SpreadsheetApp.getUi().alert("Could not read tailored resume."); return; }

  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const prompt = "You are comparing two versions of a resume — a base version and a tailored version created for " + company + " (" + role + ").\n\n" +
    "Provide an exhaustive, line-by-line diff. For every change:\n" +
    "- REMOVED: quote the exact original text that was deleted or replaced\n" +
    "- ADDED: quote the exact new text that replaced it or was inserted\n" +
    "- If a bullet was reworded, show both the old and new wording side by side\n" +
    "- Flag any section where nothing changed\n" +
    "- At the end, summarize: how many bullets were changed, how many were untouched, and your assessment of whether the changes meaningfully differentiate the resume for this role\n\n" +
    "BASE RESUME:\n" + baseText + "\n\n" +
    "TAILORED RESUME:\n" + tailoredText;

  const payload = {
    model: model,
    max_tokens: 4096,
    system: "You are a precise resume diff analyst. Be exhaustive — miss nothing.",
    messages: [{ role: "user", content: prompt }]
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", options);
  const json = JSON.parse(response.getContentText());
  const diffText = json.content && json.content[0] && json.content[0].text ? json.content[0].text : "";

  ss.toast("", "", 1);

  if (!diffText) {
    SpreadsheetApp.getUi().alert("Analysis failed. Check Apps Script execution log.");
    return;
  }

  const folderUrl = config["Application Docs Folder"] || "";
  let outputUrl = "";
  if (folderUrl) {
    const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) || folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) {
      const parentFolder = DriveApp.getFolderById(match[1]);
      const existing = parentFolder.getFoldersByName(company);
      const companyFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(company);
      const doc = DocumentApp.create(company + (role ? " — " + role.substring(0, 40) : "") + " — Resume Diff Analysis");
      DriveApp.getFileById(doc.getId()).moveTo(companyFolder);
      doc.getBody().setText(diffText);
      doc.saveAndClose();
      outputUrl = "https://docs.google.com/document/d/" + doc.getId() + "/edit";
    }
  }

  if (outputUrl) {
    SpreadsheetApp.getActiveSpreadsheet().toast("Resume diff saved to Drive: " + outputUrl, "Job Assistant", 12);
  } else {
    SpreadsheetApp.getUi().alert("Resume Diff Analysis:\n\n" + diffText.substring(0, 1500) + (diffText.length > 1500 ? "\n\n[truncated — add Application Docs Folder to Config to save full output to Drive]" : ""));
  }
}

function proofreadResumeForSelectedRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }
  const row = sheet.getActiveCell().getRow();
  if (row <= JOBS_HEADER_ROW) { SpreadsheetApp.getUi().alert("Select a job row first."); return; }
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];
  const rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const tailoredDraftColPR = getColumnIndexByHeader(headers, "Tailored Resume Draft");
  const tailoredUrl = tailoredDraftColPR
    ? extractUrlFromFormula(rowFormulas[tailoredDraftColPR - 1], rowValues[tailoredDraftColPR - 1])
    : "";
  const config = getConfigForRuntime();
  const resumeUrl = tailoredUrl || config["Base Resume URL"] || "";
  if (!resumeUrl) { SpreadsheetApp.getUi().alert("No resume found. Check Base Resume URL in Config."); return; }
  if (!confirmRowAction(sheet, row, "Proofread Resume")) return;
  proofreadResumeForRow(sheet, row, headers, resumeUrl);
}

function markNextStepComplete() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const row = sheet.getActiveCell().getRow();
  if (row <= JOBS_HEADER_ROW) { SpreadsheetApp.getUi().alert("Select a job row first."); return; }
  if (!confirmRowAction(sheet, row, "Mark Next Step Complete")) return;

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  advanceNextStepForRow(sheet, row, headers);

  // If advancing to Follow Up, auto-set status to Applied - date
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  if (nextStepCol && statusCol) {
    const newNextStep = String(sheet.getRange(row, nextStepCol).getValue() || "").trim();
    if (newNextStep === "Paste Domain" || newNextStep === "Follow Up") {
      const currentStatus = String(sheet.getRange(row, statusCol).getValue() || "").trim().toLowerCase();
      if (!currentStatus.startsWith("applied")) {
        const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
        sheet.getRange(row, statusCol).setValue("Applied - " + today);
      }
    }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Next step advanced.", "Job Assistant", 4);
}

function removeNextStepCompleteColumn() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = getColumnIndexByHeader(headers, "Next Step Complete");
  if (!col) { SpreadsheetApp.getUi().alert("No 'Next Step Complete' column found."); return; }

  sheet.deleteColumn(col);
  SpreadsheetApp.getActiveSpreadsheet().toast("Next Step Complete column removed.", "Job Assistant", 8);
}

function getNextStep(currentStep) {
  const normalized = normalizeNextStep(currentStep);

  if (!normalized || normalized === "Pass" || normalized === "Closed" || normalized === "Review if easy" || normalized === "Discuss") {
    return "Verify company site";
  }

  const index = NEXT_STEP_SEQUENCE.indexOf(normalized);

  if (index === -1) return "Verify company site";
  if (index >= NEXT_STEP_SEQUENCE.length - 1) return "Follow Up";

  return NEXT_STEP_SEQUENCE[index + 1];
}

// Rewrites the Recommendation cell to canonical case (e.g. typed "discuss - apply" -> "Discuss - Apply").
// handleJobsSheetEdit matches case-insensitively, but the sheet's native conditional-format
// regex ("- Apply$") is case-sensitive — without this, a non-canonical-case entry would be
// processed correctly (Status/Next Step) but never get the matching color highlighting.
function canonicalizeBorderlineRecommendation(sheet, row, recommendationCol, rawRecommendation) {
  const canonical = BORDERLINE_RESOLUTION_RECOMMENDATIONS.filter(function (v) {
    return v.toLowerCase() === rawRecommendation.toLowerCase();
  })[0];

  if (canonical && canonical !== rawRecommendation) {
    sheet.getRange(row, recommendationCol).setValue(canonical);
  }
}

// Shared column lookups for the manual Recommendation-override handlers below
// (Apply Anyway / Discuss / borderline resolution) — they all touch the same four columns.
function getOverrideColumns(headers) {
  return {
    nextStepCol: getColumnIndexByHeader(headers, "Next Step"),
    dueDateCol: getColumnIndexByHeader(headers, "Next Step Due Date"),
    statusCol: getColumnIndexByHeader(headers, "Status"),
    lastUpdatedCol: getColumnIndexByHeader(headers, "Last Updated")
  };
}

function applyApplyAnywayOverride(sheet, row, headers) {
  const config = getConfigForRuntime();

  const cols = getOverrideColumns(headers);
  const nextStepCol = cols.nextStepCol;
  const dueDateCol = cols.dueDateCol;
  const lastUpdatedCol = cols.lastUpdatedCol;
  const statusCol = cols.statusCol;

  if (!nextStepCol) return;

  const currentStep = normalizeNextStep(sheet.getRange(row, nextStepCol).getValue());

  if (!currentStep || currentStep === "Pass" || currentStep === "Closed") {
    sheet.getRange(row, nextStepCol).setValue("Verify company site");

    if (dueDateCol) {
      sheet.getRange(row, dueDateCol).setValue(getDueDateForStep("Verify company site", new Date(), config));
    }

    if (statusCol) {
      sheet.getRange(row, statusCol).setValue("Review");
    }

    if (lastUpdatedCol) {
      sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
    }
  }
}

function applyDiscussOverride(sheet, row, headers) {
  const config = getConfigForRuntime();

  const cols = getOverrideColumns(headers);
  const nextStepCol = cols.nextStepCol;
  const dueDateCol = cols.dueDateCol;
  const statusCol = cols.statusCol;
  const lastUpdatedCol = cols.lastUpdatedCol;

  if (nextStepCol) {
    sheet.getRange(row, nextStepCol).setValue("Discuss");
  }

  if (dueDateCol) {
    sheet.getRange(row, dueDateCol).setValue(getDueDateForStep("Discuss", new Date(), config));
  }

  if (statusCol) {
    sheet.getRange(row, statusCol).setValue("Review");
  }

  if (lastUpdatedCol) {
    sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
  }
}

function applyBorderlineResolutionOverride(sheet, row, headers, resolvedTo) {
  const cols = getOverrideColumns(headers);
  const nextStepCol = cols.nextStepCol;
  const dueDateCol = cols.dueDateCol;
  const statusCol = cols.statusCol;
  const lastUpdatedCol = cols.lastUpdatedCol;

  if (resolvedTo === "Skip") {
    if (nextStepCol) sheet.getRange(row, nextStepCol).setValue("Pass");
    if (statusCol) sheet.getRange(row, statusCol).setValue("Skip");
    if (lastUpdatedCol) sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
    return;
  }

  // resolvedTo === "Apply": only reset Next Step / Status if they're still sitting on
  // borderline placeholder values — don't stomp on progress already made (e.g. re-editing
  // the Recommendation cell on a row that's since been applied to or moved to Interview).
  const config = getConfigForRuntime();

  if (nextStepCol) {
    const currentStep = normalizeNextStep(sheet.getRange(row, nextStepCol).getValue());
    const unresolvedSteps = ["", "Review if easy", "Discuss", "Pass", "Closed"];

    if (unresolvedSteps.indexOf(currentStep) !== -1) {
      sheet.getRange(row, nextStepCol).setValue("Verify company site");

      if (dueDateCol) {
        sheet.getRange(row, dueDateCol).setValue(getDueDateForStep("Verify company site", new Date(), config));
      }
    }
  }

  if (statusCol) {
    const currentStatus = String(sheet.getRange(row, statusCol).getValue() || "").trim();
    const unresolvedStatuses = ["", "New", "Review"];

    if (unresolvedStatuses.indexOf(currentStatus) !== -1) {
      sheet.getRange(row, statusCol).setValue("New");
    }
  }

  if (lastUpdatedCol) sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
}

function sortJobsByPrioritySafe() {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    sortJobsByPriority();
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}

function insertMetricsDashboardRow() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const firstCell = sheet.getRange(1, 1).getValue();
  if (String(firstCell).trim() === "__metrics__") {
    SpreadsheetApp.getUi().alert("Metrics row already exists (row 1).");
    return;
  }

  sheet.insertRowBefore(1);
  sheet.getRange(1, 1).setValue("__metrics__");
  sheet.setFrozenRows(2);
  SpreadsheetApp.getActiveSpreadsheet().toast("Metrics row inserted as row 1. Row 2 is now the header row. Add your metrics formulas and Claude question cell to row 1.", "Job Assistant", 12);
}

function insertSecondDashboardRow() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const row2Cell = String(sheet.getRange(2, 1).getValue()).trim();
  if (row2Cell === "__metrics2__") {
    SpreadsheetApp.getUi().alert("Second dashboard row already exists (row 2).");
    return;
  }

  sheet.insertRowBefore(2);
  sheet.getRange(2, 1).setValue("__metrics2__");
  sheet.setFrozenRows(3);
  SpreadsheetApp.getActiveSpreadsheet().toast("Second dashboard row inserted as row 2. Row 3 is now the header row.", "Job Assistant", 8);
}

function purgeOrphanedSortColumns(sheet) {
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  for (let i = headers.length - 1; i >= 0; i--) {
    if (String(headers[i]).startsWith("__sort_")) {
      sheet.deleteColumn(i + 1);
    }
  }
}

function sortJobsByPriority() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);

  if (!sheet) throw new Error("Missing tab: Jobs");

  purgeOrphanedSortColumns(sheet);

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 2) return;

  let headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idCol = getColumnIndexByHeader(headers, "ID");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const fitCol = getColumnIndexByHeader(headers, "Fit Score");
  const atsCol = getColumnIndexByHeader(headers, "ATS Score");
  const stabilityCol = getColumnIndexByHeader(headers, "Stability");
  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const companyResponseCol = getColumnIndexByHeader(headers, "Auto Response");
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const responseProbCol = getColumnIndexByHeader(headers, "Response Probability");

  if (!companyCol || !fitCol || !atsCol || !recommendationCol || !statusCol || !nextStepCol || !dueDateCol) {
    throw new Error("Cannot sort. Required columns missing.");
  }

  normalizeStatusFormattingBeforeSort(sheet, headers);
  ensureDueDatesBeforeSort(sheet, headers);

  headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

  const refreshedLastCol = sheet.getLastColumn();
  const helperStartCol = refreshedLastCol + 1;
  const helperCount = 12;

  sheet.insertColumnsAfter(refreshedLastCol, helperCount);

  try {
    const helperHeaders = [
      "__sort_valid_row__",
      "__sort_color__",
      "__sort_lifecycle__",
      "__sort_recommendation__",
      "__sort_response_prob__",
      "__sort_due_urgency__",
      "__sort_next_step__",
      "__sort_fit__",
      "__sort_ats__",
      "__sort_stability__",
      "__sort_due_date__",
      "__sort_id__"
    ];

    sheet.getRange(JOBS_HEADER_ROW, helperStartCol, 1, helperCount).setValues([helperHeaders]);

    const dataRange = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, refreshedLastCol);
    const values = dataRange.getValues();

    const helperValues = values.map(row => {
      const company = String(row[companyCol - 1] || "").trim();
      const recommendation = recommendationCol ? String(row[recommendationCol - 1] || "").trim() : "";
      const status = statusCol ? String(row[statusCol - 1] || "").trim() : "";
      const response = companyResponseCol ? String(row[companyResponseCol - 1] || "").trim() : "";
      const nextStep = nextStepCol ? String(row[nextStepCol - 1] || "").trim() : "";
      const dueDate = dueDateCol ? row[dueDateCol - 1] : "";
      const stability = stabilityCol ? String(row[stabilityCol - 1] || "").trim() : "";
      const fitScore = fitCol ? Number(row[fitCol - 1]) || 0 : 0;
      const atsRaw = atsCol ? String(row[atsCol - 1] || "") : "";
      const atsScore = extractBaseAtsScore(atsRaw) || 0;
      const idValue = idCol ? String(row[idCol - 1] || "999999") : "999999";
      const responseProbRaw = responseProbCol ? String(row[responseProbCol - 1] || "") : "";
      // Only breaks ties within the resolved-Apply tier (Apply / Discuss - Apply / Apply if
      // Easy - Apply); 0 for every other row so it doesn't affect their existing ordering.
      const responseProbRank = isResolvedApplyRecommendation(recommendation)
        ? -(extractResponseProbabilityValue(responseProbRaw) || 0)
        : 0;

      return [
        company ? 0 : 9,
        getColorSortRank(recommendation, status, response),
        getLifecycleSortRank(status, response, recommendation, nextStep),
        getRecommendationSortRank(recommendation),
        responseProbRank,
        getDueUrgencySortRank(status, response, recommendation, nextStep, dueDate),
        getNextStepSortRank(nextStep),
        -fitScore,
        -atsScore,
        getStabilitySortRank(stability),
        getDateSortRank(dueDate),
        idValue
      ];
    });

    sheet.getRange(JOBS_DATA_START_ROW, helperStartCol, helperValues.length, helperCount).setValues(helperValues);

    const fullSortRange = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, refreshedLastCol + helperCount);

    fullSortRange.sort([
      { column: helperStartCol,      ascending: true },
      { column: helperStartCol + 1,  ascending: true },
      { column: helperStartCol + 2,  ascending: true },
      { column: helperStartCol + 3,  ascending: true },
      { column: helperStartCol + 4,  ascending: true },
      { column: helperStartCol + 5,  ascending: true },
      { column: helperStartCol + 6,  ascending: true },
      { column: helperStartCol + 7,  ascending: true },
      { column: helperStartCol + 8,  ascending: true },
      { column: helperStartCol + 9,  ascending: true },
      { column: helperStartCol + 10, ascending: true },
      { column: helperStartCol + 11, ascending: true }
    ]);
  } finally {
    sheet.deleteColumns(helperStartCol, helperCount);
    SpreadsheetApp.flush();
  }

  // Clear any banding ranges and individual cell backgrounds carried by the sort
  sheet.getBandings().forEach(function(b) { b.remove(); });
  const dataRows = sheet.getLastRow() - JOBS_HEADER_ROW;
  if (dataRows > 0) {
    sheet.getRange(JOBS_DATA_START_ROW, 1, dataRows, sheet.getLastColumn()).setBackground("#ffffff");
  }

  const validation = validateTrackerSilent(sheet);

  if (!validation.ok) {
    console.warn(validation.message);
  }
}

function normalizeStatusFormattingBeforeSort(sheet, headers) {
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (!statusCol) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const statusValues = sheet.getRange(JOBS_DATA_START_ROW, statusCol, lastRow - JOBS_HEADER_ROW, 1).getValues();
  const lastUpdatedValues = lastUpdatedCol
    ? sheet.getRange(JOBS_DATA_START_ROW, lastUpdatedCol, lastRow - JOBS_HEADER_ROW, 1).getValues()
    : [];

  let changed = false;

  for (let i = 0; i < statusValues.length; i++) {
    const raw = String(statusValues[i][0] || "").trim();
    const lower = raw.toLowerCase();

    // Only rewrite bare status words — never touch "Applied - <date>" already set
    if (lower === "applied") {
      const fallbackDate = lastUpdatedCol
        ? getDateTextFromValue(lastUpdatedValues[i][0])
        : "";

      statusValues[i][0] = "Applied - " + (fallbackDate || formatDate(new Date()));
      changed = true;
    }

    if (lower === "rejected") {
      const fallbackDate = lastUpdatedCol
        ? getDateTextFromValue(lastUpdatedValues[i][0])
        : "";

      statusValues[i][0] = "Rejected - " + (fallbackDate || formatDate(new Date()));
      changed = true;
    }

    if (lower === "closed") {
      const fallbackDate = lastUpdatedCol
        ? getDateTextFromValue(lastUpdatedValues[i][0])
        : "";

      statusValues[i][0] = "Closed - " + (fallbackDate || formatDate(new Date()));
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(JOBS_DATA_START_ROW, statusCol, lastRow - JOBS_HEADER_ROW, 1).setValues(statusValues);
  }
}

function ensureDueDatesBeforeSort(sheet, headers) {
  const config = getConfigForRuntime();

  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (!nextStepCol || !dueDateCol) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const nextStepValues = sheet.getRange(JOBS_DATA_START_ROW, nextStepCol, lastRow - JOBS_HEADER_ROW, 1).getValues();
  const dueDateValues = sheet.getRange(JOBS_DATA_START_ROW, dueDateCol, lastRow - JOBS_HEADER_ROW, 1).getValues();
  const lastUpdatedValues = lastUpdatedCol
    ? sheet.getRange(JOBS_DATA_START_ROW, lastUpdatedCol, lastRow - JOBS_HEADER_ROW, 1).getValues()
    : [];

  let changed = false;

  for (let i = 0; i < nextStepValues.length; i++) {
    const nextStep = nextStepValues[i][0];
    const dueDate = dueDateValues[i][0];
    const baseDate = lastUpdatedCol ? lastUpdatedValues[i][0] : new Date();

    if (isActionNextStep(nextStep) && !dueDate) {
      dueDateValues[i][0] = getDueDateForStep(nextStep, baseDate || new Date(), config);
      changed = true;
    }

    if (!isActionNextStep(nextStep) && normalizeNextStep(nextStep) === "Pass" && dueDate) {
      dueDateValues[i][0] = "";
      changed = true;
    }
  }

  if (changed) {
    sheet.getRange(JOBS_DATA_START_ROW, dueDateCol, lastRow - JOBS_HEADER_ROW, 1).setValues(dueDateValues);
  }
}

function resetInitialActionDueDatesFromDateAdded() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  const config = getConfigForRuntime();

  if (!sheet) throw new Error("Missing tab: Jobs");

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (!recommendationCol || !statusCol || !nextStepCol || !dueDateCol || !lastUpdatedCol) {
    SpreadsheetApp.getUi().alert("Missing required columns.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, sheet.getLastColumn()).getValues();

  let updated = 0;

  for (let i = 0; i < values.length; i++) {
    const rowNumber = i + 2;
    const row = values[i];

    const recommendation = String(row[recommendationCol - 1] || "").trim();
    const status = String(row[statusCol - 1] || "").trim().toLowerCase();
    const nextStep = normalizeNextStep(row[nextStepCol - 1]);
    const lastUpdated = row[lastUpdatedCol - 1];

    const isInitialActive =
      status === "new" &&
      (recommendation === "Apply" || recommendation === "Apply if Easy" || recommendation === "Discuss") &&
      (nextStep === "Verify company site" || nextStep === "Review if easy" || nextStep === "Discuss");

    if (isInitialActive) {
      const baseDate = lastUpdated || new Date();
      sheet.getRange(rowNumber, dueDateCol).setValue(getDueDateForStep(nextStep, baseDate, config));
      updated++;
    }
  }

  sortJobsByPriority();

  SpreadsheetApp.getActiveSpreadsheet().toast("Updated " + updated + " initial action due date(s).", "Job Assistant", 8);
}

function getDateTextFromValue(value) {
  if (!value) return "";

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return formatDate(value);
  }

  const text = String(value || "").trim();
  const match = text.match(/\d{4}-\d{2}-\d{2}/);

  if (match) return match[0];

  return "";
}

function isInterviewStatus(status) {
  return String(status || "").trim().toLowerCase().indexOf("interview -") === 0;
}

// A borderline recommendation (Discuss / Apply if Easy) resolved via manual edit
// to "<original> - Apply" or "<original> - Skip" should sort/color identically to
// a plain Apply or Skip recommendation.
function isResolvedToApply(recommendation) {
  return /- Apply$/i.test(String(recommendation || "").trim());
}

function getColorSortRank(recommendation, status, response) {
  const rec  = String(recommendation || "").trim();
  const stat = String(status || "").trim();

  if (isInterviewStatus(stat) || isPositiveCompanyResponse(String(response || "").trim().toLowerCase())) return 0; // red text: interview/positive response — always highest priority
  if ((rec === "Apply" || isResolvedToApply(rec)) && stat === "New") return 1; // green: hottest
  if (/^review$/i.test(stat))                     return 2; // blue: in review
  if (rec === "Apply if Easy" && stat === "New")  return 3; // yellow: warm
  if (/^applied/i.test(stat))                     return 5; // yellow-green: already applied
  if (/skip|rejected/i.test(stat))               return 6; // red: done
  return 4;                                                  // everything else active
}

function getLifecycleSortRank(status, response, recommendation, nextStep) {
  const statusText = String(status || "").trim().toLowerCase();
  const responseText = String(response || "").trim().toLowerCase();
  const recommendationText = String(recommendation || "").trim().toLowerCase();
  const nextStepText = String(nextStep || "").trim().toLowerCase();

  // Rejected/closed always goes to bottom regardless of Auto Response text
  if (isRejectedOrClosed(statusText, responseText, nextStepText)) return 90;

  if (isInterviewStatus(statusText) || isPositiveCompanyResponse(responseText)) return 0;

  if (
    statusText === "skip" ||
    recommendationText === "skip" ||
    nextStepText === "pass"
  ) {
    return 95;
  }

  // Applied rows always sort below active rows, regardless of next step
  if (statusText.indexOf("applied") === 0) return 70;

  if (statusText.indexOf("networking") === 0) return 60;

  if (
    statusText === "review" ||
    recommendationText === "discuss" ||
    nextStepText.indexOf("discuss") !== -1
  ) {
    return 20;
  }

  if (statusText === "new" || !statusText) return 10;

  return 50;
}

function getDueUrgencySortRank(status, response, recommendation, nextStep, dueDate) {
  const statusText = String(status || "").trim().toLowerCase();
  const responseText = String(response || "").trim().toLowerCase();
  const recommendationText = String(recommendation || "").trim().toLowerCase();
  const nextStepText = String(nextStep || "").trim().toLowerCase();

  if (isRejectedOrClosed(statusText, responseText, nextStepText)) return 9;

  if (
    statusText === "skip" ||
    recommendationText === "skip" ||
    nextStepText === "pass"
  ) {
    return 9;
  }

  if (isInterviewStatus(statusText) || isPositiveCompanyResponse(responseText)) return 0;

  // Applied/networking rows: deprioritize due urgency so lifecycle rank governs placement
  if (statusText.indexOf("applied") === 0 || statusText.indexOf("networking") === 0) return 7;

  if (!isActionNextStep(nextStep)) return 8;

  if (!dueDate) return 3;

  const todayRank = Number(Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMdd"));
  const dueRank = getDateSortRank(dueDate);

  if (dueRank <= todayRank) return 0;

  return 2;
}

function isRejectedOrClosed(statusText, responseText, nextStepText) {
  const combined = [statusText, responseText, nextStepText].join(" ");

  return (
    combined.indexOf("rejected") !== -1 ||
    combined.indexOf("declined") !== -1 ||
    combined.indexOf("closed") !== -1 ||
    combined.indexOf("not selected") !== -1 ||
    combined.indexOf("no longer") !== -1 ||
    combined.indexOf("passed") !== -1
  );
}

function isPositiveCompanyResponse(responseText) {
  const text = String(responseText || "").trim();

  if (!text) return false;

  // checkJobApplicationEmails() always writes entries as "DATE: <Classification> — summary",
  // where Classification is one of a known set. Matching the label itself (rather than loose
  // keywords like "next step" or "follow up") avoids false positives from routine "Application
  // Received" confirmations, which commonly contain that same boilerplate language.
  return /:\s*(Interview Request|Offer|Assessment)\b/i.test(text);
}

function getRecommendationSortRank(recommendation) {
  const text = String(recommendation || "").trim().toLowerCase();

  if (text === "apply" || /- apply$/.test(text)) return 1;
  if (text === "apply anyway") return 1.5;
  if (text === "discuss") return 2;
  if (text === "apply if easy") return 3;
  if (text === "review") return 3.5;
  if (text === "skip" || /- skip$/.test(text)) return 4;

  return 5;
}

// Matches "Apply", "Discuss - Apply", and "Apply if Easy - Apply" — every recommendation
// that resolves to an Apply action, whether original or manually resolved from a borderline
// call. Same predicate getRecommendationSortRank uses for its rank-1 tier.
function isResolvedApplyRecommendation(recommendation) {
  const text = String(recommendation || "").trim().toLowerCase();
  return text === "apply" || /- apply$/.test(text);
}

function getNextStepSortRank(nextStep) {
  const normalized = normalizeNextStep(nextStep);

  // Further along the workflow = lower rank = sorts first
  if (normalized === "Apply") return 1;
  if (normalized === "Paste Domain") return 2;
  if (normalized === "Create Resume/Letter") return 3;
  if (normalized === "Find Contact") return 4;
  if (normalized === "Verify company site") return 5;
  if (normalized === "Discuss") return 6;
  if (normalized === "Review if easy") return 7;
  if (normalized === "Follow Up") return 8;
  if (normalized === "Pass") return 9;
  if (normalized === "Closed") return 10;

  return 8;
}

function getStabilitySortRank(stability) {
  const text = String(stability || "").trim().toLowerCase();

  if (!text) return 5;
  if (text.includes("strong")) return 1;
  if (text.includes("stable")) return 1;
  if (text.includes("good")) return 2;
  if (text.includes("medium")) return 3;
  if (text.includes("mixed")) return 4;
  if (text.includes("risk")) return 6;
  if (text.includes("unknown")) return 7;

  return 5;
}

function getDateSortRank(value) {
  if (!value) return 99999999;

  if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
    return Number(Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyyMMdd"));
  }

  const text = String(value || "").trim();
  const match = text.match(/(\d{4})-(\d{2})-(\d{2})/);

  if (match) {
    return Number(match[1] + match[2] + match[3]);
  }

  return 99999999;
}

function validateTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);

  if (!sheet) throw new Error("Missing tab: Jobs");

  const result = validateTrackerSilent(sheet);

  SpreadsheetApp.getUi().alert(result.message);
}

function validateTrackerSilent(sheet) {
  const sortResult = validateSortOrderSilent(sheet);
  const fieldResult = validateRequiredFieldsSilent(sheet);

  const ok = sortResult.ok && fieldResult.ok;

  let message = "";

  if (ok) {
    message = "Tracker validation passed. Sort order and required fields look good.";
  } else {
    message =
      "Tracker validation needs attention.\n\n" +
      sortResult.message + "\n\n" +
      fieldResult.message;
  }

  return {
    ok: ok,
    message: message
  };
}

function validateSortOrderSilent(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 3) {
    return {
      ok: true,
      message: "Sort validation passed. Not enough rows to compare."
    };
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idCol = getColumnIndexByHeader(headers, "ID");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const fitCol = getColumnIndexByHeader(headers, "Fit Score");
  const atsCol = getColumnIndexByHeader(headers, "ATS Score");
  const stabilityCol = getColumnIndexByHeader(headers, "Stability");
  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const companyResponseCol = getColumnIndexByHeader(headers, "Auto Response");
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");
  const responseProbCol = getColumnIndexByHeader(headers, "Response Probability");

  if (!companyCol || !fitCol || !atsCol || !recommendationCol || !statusCol || !nextStepCol || !dueDateCol) {
    return {
      ok: false,
      message: "Sort validation failed. Required sort columns are missing."
    };
  }

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, lastCol).getValues();

  const ranks = values.map(row => {
    const company = String(row[companyCol - 1] || "").trim();
    const status = statusCol ? String(row[statusCol - 1] || "").trim() : "";
    const response = companyResponseCol ? String(row[companyResponseCol - 1] || "").trim() : "";
    const recommendation = recommendationCol ? String(row[recommendationCol - 1] || "").trim() : "";
    const nextStep = nextStepCol ? String(row[nextStepCol - 1] || "").trim() : "";
    const dueDate = dueDateCol ? row[dueDateCol - 1] : "";
    const fitScore = fitCol ? Number(row[fitCol - 1]) || 0 : 0;
    const atsScore = atsCol ? Number(row[atsCol - 1]) || 0 : 0;
    const stability = stabilityCol ? String(row[stabilityCol - 1] || "").trim() : "";
    const idValue = idCol ? Number(row[idCol - 1]) || 999999 : 999999;
    const responseProbRaw = responseProbCol ? String(row[responseProbCol - 1] || "") : "";
    const responseProbRank = isResolvedApplyRecommendation(recommendation)
      ? -(extractResponseProbabilityValue(responseProbRaw) || 0)
      : 0;

    return [
      company ? 0 : 9,
      getColorSortRank(recommendation, status, response),
      getLifecycleSortRank(status, response, recommendation, nextStep),
      getDueUrgencySortRank(status, response, recommendation, nextStep, dueDate),
      getRecommendationSortRank(recommendation),
      responseProbRank,
      getNextStepSortRank(nextStep),
      -fitScore,
      -atsScore,
      getStabilitySortRank(stability),
      getDateSortRank(dueDate),
      idValue
    ];
  });

  for (let i = 1; i < ranks.length; i++) {
    if (compareRankArrays(ranks[i - 1], ranks[i]) > 0) {
      return {
        ok: false,
        message: "Sort validation failed around sheet rows " + (i + 1) + " and " + (i + 2) + ". Run Job Assistant → Sort Jobs by Priority again."
      };
    }
  }

  return {
    ok: true,
    message: "Sort validation passed."
  };
}

function validateRequiredFieldsSilent(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 2) {
    return {
      ok: true,
      message: "Field validation passed. No job rows found."
    };
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, lastCol).getValues()[0];

  const idCol = getColumnIndexByHeader(headers, "ID");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const statusCol = getColumnIndexByHeader(headers, "Status");
  const nextStepCol = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol = getColumnIndexByHeader(headers, "Next Step Due Date");

  const values = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, lastCol).getValues();

  const issues = [];

  for (let i = 0; i < values.length; i++) {
    const sheetRow = i + JOBS_DATA_START_ROW;
    const row = values[i];

    const id = idCol ? String(row[idCol - 1] || "").trim() : "";
    const company = companyCol ? String(row[companyCol - 1] || "").trim() : "";
    const role = roleCol ? String(row[roleCol - 1] || "").trim() : "";
    const recommendation = recommendationCol ? String(row[recommendationCol - 1] || "").trim() : "";
    const status = statusCol ? String(row[statusCol - 1] || "").trim() : "";
    const nextStep = nextStepCol ? String(row[nextStepCol - 1] || "").trim() : "";
    const dueDate = dueDateCol ? row[dueDateCol - 1] : "";

    if (!company && !role) continue;

    if (!id) {
      issues.push("Row " + sheetRow + ": missing ID.");
    }

    if (!company) {
      issues.push("Row " + sheetRow + ": missing Company.");
    }

    if (!role) {
      issues.push("Row " + sheetRow + ": missing Role.");
    }

    if (recommendation &&
        VALID_RECOMMENDATIONS.indexOf(recommendation) === -1 &&
        recommendation.toLowerCase() !== "apply anyway" &&
        BORDERLINE_RESOLUTION_RECOMMENDATIONS.map(function (v) { return v.toLowerCase(); }).indexOf(recommendation.toLowerCase()) === -1) {
      issues.push("Row " + sheetRow + ": invalid Recommendation: " + recommendation);
    }

    if (status && !isAllowedStatusValue(status)) {
      issues.push("Row " + sheetRow + ": unusual Status: " + status);
    }

    if (nextStep && !isAllowedNextStepValue(nextStep)) {
      issues.push("Row " + sheetRow + ": invalid Next Step: " + nextStep);
    }

    if (isActionNextStep(nextStep) && !dueDate) {
      issues.push("Row " + sheetRow + ": action row missing Next Step Due Date.");
    }
  }

  if (issues.length === 0) {
    return {
      ok: true,
      message: "Field validation passed."
    };
  }

  return {
    ok: false,
    message: "Field validation found " + issues.length + " issue(s):\n" + issues.slice(0, 12).join("\n") + (issues.length > 12 ? "\n...and more." : "")
  };
}

function isAllowedStatusValue(value) {
  const text = String(value || "").trim();

  if (VALID_STATUSES.indexOf(text) !== -1) return true;

  const lower = text.toLowerCase();

  if (lower.indexOf("applied -") === 0) return true;
  if (lower.indexOf("interview -") === 0) return true;
  if (lower.indexOf("rejected -") === 0) return true;
  if (lower.indexOf("closed -") === 0) return true;

  return false;
}

function isAllowedNextStepValue(value) {
  const text = String(value || "").trim();

  if (VALID_NEXT_STEPS.indexOf(text) !== -1) return true;

  const normalized = normalizeNextStep(text);

  return normalized !== "";
}

function compareRankArrays(a, b) {
  for (let i = 0; i < Math.max(a.length, b.length); i++) {
    const av = a[i];
    const bv = b[i];

    if (av < bv) return -1;
    if (av > bv) return 1;
  }

  return 0;
}

function getResponseSortRank(response) {
  const text = String(response || "").trim().toLowerCase();

  if (!text) return 0;

  if (
    text.includes("rejected") ||
    text.includes("declined") ||
    text.includes("closed") ||
    text.includes("not selected") ||
    text.includes("no longer") ||
    text.includes("passed")
  ) {
    return 9;
  }

  if (
    text.includes("interview") ||
    text.includes("screen") ||
    text.includes("recruiter") ||
    text.includes("next step")
  ) {
    return 1;
  }

  return 2;
}

function getColumnIndexByHeader(headers, headerName) {
  const target = normalizeHeader(headerName);

  for (let i = 0; i < headers.length; i++) {
    if (normalizeHeader(headers[i]) === target) {
      return i + 1;
    }
  }

  return null;
}

function normalizeHeader(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s*\([^)]*\)/g, "")  // strip parenthetical annotations like "(Base / Tailored)"
    .replace(/\s+/g, " ")
    .trim();
}

// ---- Resume helpers ----

function extractGoogleDocId(url) {
  const s = String(url || "");
  const match = s.match(/\/document\/d\/([a-zA-Z0-9_-]+)/) ||
                s.match(/\/d\/([a-zA-Z0-9_-]+)/) ||
                s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  return match ? match[1] : "";
}

function extractUrlFromCellRange(range) {
  try {
    const rt = range.getRichTextValue();
    if (rt) {
      const link = rt.getLinkUrl();
      if (link) return link;
    }
  } catch (e) {}
  const formula = range.getFormula();
  const match = formula && formula.match(/=HYPERLINK\("([^"]+)"/i);
  if (match) return match[1];
  return String(range.getValue() || "").trim();
}

function extractUrlFromFormula(formula, fallback) {
  const match = formula && formula.match(/=HYPERLINK\("([^"]+)"/i);
  if (match) return match[1];
  return String(fallback || "").trim();
}

function resumeHyperlinkFormula(url, label) {
  if (!url) return "";
  return '=HYPERLINK("' + url + '","' + (label || "Resume") + '")';
}

function getResumeText(url) {
  if (!url) return "";
  try {
    const docId = extractGoogleDocId(url);
    if (!docId) return "";
    const doc = DocumentApp.openById(docId);
    return doc.getBody().getText();
  } catch (e) {
    console.warn("Could not read resume document: " + e.message);
    return "";
  }
}

// ---- ATS recalculation ----

function recalculateAtsScore(sheet, row, headers) {
  const config = getConfigForRuntime();

  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const atsScoreCol = getColumnIndexByHeader(headers, "ATS Score");
  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");
  const lastUpdatedCol = getColumnIndexByHeader(headers, "Last Updated");

  if (!atsScoreCol || !jobDescCol) return;

  const rowRange = sheet.getRange(row, 1, 1, sheet.getLastColumn());
  const rowData = rowRange.getValues()[0];
  const rowFormulas = rowRange.getFormulas()[0];
  const jobDescription = String(rowData[jobDescCol - 1] || "").trim();
  if (!jobDescription) return;

  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role = roleCol ? String(rowData[roleCol - 1] || "").trim() : "";
  const baseUrl = config["Base Resume URL"] || "";

  const baseText = getResumeText(baseUrl);
  const baseResult = callClaudeForAtsScore(jobDescription, company, role, baseText, config);
  if (baseResult === null) return;
  const baseScore = baseResult.score;

  const tailoredUrl = tailoredDraftCol
    ? extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowData[tailoredDraftCol - 1])
    : "";

  let displayValue;
  if (tailoredUrl) {
    const tailoredText = getResumeText(tailoredUrl);
    const tailoredResult = tailoredText ? callClaudeForAtsScore(jobDescription, company, role, tailoredText, config) : null;
    displayValue = tailoredResult !== null ? baseScore + " / " + tailoredResult.score : baseScore;
  } else {
    displayValue = baseScore;
  }

  sheet.getRange(row, atsScoreCol).setValue(displayValue);
  if (lastUpdatedCol) {
    sheet.getRange(row, lastUpdatedCol).setValue(formatDate(new Date()));
  }
}

function extractBaseAtsScore(cellValue) {
  const plain = String(cellValue || "").trim();
  // "77 / 88" format — base is first number
  const slashMatch = plain.match(/^(\d+)\s*\//);
  if (slashMatch) return Number(slashMatch[1]);
  // legacy "(base)" format
  const baseMatch = plain.match(/^(\d+)\s*\(base\)/);
  if (baseMatch) return Number(baseMatch[1]);
  const num = Number(plain);
  return isNaN(num) ? null : num;
}

function extractBaseResponseProbability(cellValue) {
  const plain = String(cellValue || "").trim();
  // "20-30% / 25-35%" format — base is everything before the slash
  const slashMatch = plain.match(/^(.+?)\s*\/\s*.+$/);
  if (slashMatch) return slashMatch[1].trim();
  return plain || null;
}

// Numeric sort value from a Response Probability cell (e.g. "20-30%" or "20-30% / 25-35%").
// Uses the last segment (post-tailoring, when present) and averages the range's two numbers.
function extractResponseProbabilityValue(cellValue) {
  const plain = String(cellValue || "").trim();
  if (!plain) return null;
  const parts = plain.split("/");
  const segment = parts[parts.length - 1].trim();
  const numbers = segment.match(/\d+(\.\d+)?/g);
  if (!numbers || !numbers.length) return null;
  const values = numbers.map(Number);
  return values.reduce(function(a, b) { return a + b; }, 0) / values.length;
}

function callClaudeForResponseProbability(jobDescription, company, role, resumeText, config, baseAtsScore, tailoredAtsScore) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return null;

  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const toolSchema = {
    name: "response_probability_result",
    description: "Estimated probability of getting a response/interview based on the tailored resume",
    input_schema: {
      type: "object",
      required: ["response_probability"],
      properties: {
        response_probability: {
          type: "string",
          description: "Probability range as a percentage, e.g. '25-35%'. Never include referral probability."
        }
      }
    }
  };

  const atsContext = (baseAtsScore !== null && tailoredAtsScore !== null)
    ? "ATS score: " + baseAtsScore + " (base) → " + tailoredAtsScore + " (tailored). "
    : "";

  const userName = config["User Display Name"] || config["User Full Name"] || "the user";
  const userPrompt =
    "Estimate the probability " + userName + " receives a response or interview invitation for this role, expressed as a range like '20-30%'.\n\n" +
    "Calibration: a cold application with no referral to a competitive SE role typically has a 10-20% response rate. " +
    "Only exceed 30% if keyword match is strong AND experience requirements are substantially met. " +
    "An ATS score below 75 indicates meaningful gaps — probability should reflect that. " +
    "Do not factor in referrals.\n\n" +
    atsContext +
    "Company: " + company + "\n" +
    "Role: " + role + "\n\n" +
    "Tailored resume:\n" + (resumeText ? resumeText.substring(0, 6000) : "") + "\n\n" +
    "Job description:\n" + jobDescription.substring(0, 4000);

  const payload = {
    model: model,
    max_tokens: 64,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "response_probability_result" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return null;

  const data = JSON.parse(response.getContentText());
  for (const block of (data.content || [])) {
    if (block.type === "tool_use" && block.name === "response_probability_result") {
      return block.input.response_probability || null;
    }
  }
  return null;
}

// Payload/response builders are split out from callClaudeForAtsScore so the parallel
// cover-letter+resume flow (generateCoverLetterAndResumeParallel) can build this exact request
// and fire it concurrently with the cover-letter request via UrlFetchApp.fetchAll, without
// duplicating the prompt itself.
function buildAtsScorePayload(jobDescription, company, role, resumeText, config) {
  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const toolSchema = {
    name: "ats_score_result",
    description: "ATS score result",
    input_schema: {
      type: "object",
      required: ["ats_score", "matched_keywords"],
      properties: {
        ats_score: { type: "integer" },
        matched_keywords: {
          type: "array",
          items: { type: "string" },
          description: "Every keyword or phrase from the job description that appears in the resume"
        }
      }
    }
  };

  const userPrompt =
    "Score this resume against this job description (1-100) and list every keyword or phrase from the job description that appears in the resume.\n\n" +
    "Company: " + company + "\n" +
    "Role: " + role + "\n\n" +
    "Resume:\n" + (resumeText ? resumeText.substring(0, 6000) : "Use the candidate's standard background from Config.") + "\n\n" +
    "Job description:\n" + jobDescription.substring(0, 6000);

  return {
    model: model,
    max_tokens: 512,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "ats_score_result" },
    messages: [{ role: "user", content: userPrompt }]
  };
}

function parseAtsScoreResponse(response) {
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return null;

  const data = JSON.parse(response.getContentText());
  for (const block of (data.content || [])) {
    if (block.type === "tool_use" && block.name === "ats_score_result") {
      return {
        score: clampNumber(block.input.ats_score, 1, 100),
        keywords: block.input.matched_keywords || []
      };
    }
  }

  return null;
}

function callClaudeForAtsScore(jobDescription, company, role, resumeText, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return null;

  const payload = buildAtsScorePayload(jobDescription, company, role, resumeText, config);

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return parseAtsScoreResponse(response);
}

function updateCalibrationFromNotes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const jobsSheet = ss.getSheetByName(JOBS_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  const ui = SpreadsheetApp.getUi();

  if (!jobsSheet || !configSheet) {
    ui.alert("Jobs or Config sheet not found.");
    return;
  }

  const headers = jobsSheet.getRange(JOBS_HEADER_ROW, 1, 1, jobsSheet.getLastColumn()).getValues()[0];
  const calibCol = getColumnIndexByHeader(headers, CALIBRATION_NOTE_HEADER);

  if (!calibCol) {
    ui.alert("Add a '" + CALIBRATION_NOTE_HEADER + "' column to the Jobs sheet first.");
    return;
  }

  const companyCol        = getColumnIndexByHeader(headers, "Company");
  const roleCol           = getColumnIndexByHeader(headers, "Role");
  const recommendationCol = getColumnIndexByHeader(headers, "Recommendation");
  const atsCol            = getColumnIndexByHeader(headers, "ATS Score");
  const responseProbCol   = getColumnIndexByHeader(headers, "Response Probability");
  const jobDescCol        = getColumnIndexByHeader(headers, "Job Description");

  const lastRow = jobsSheet.getLastRow();
  if (lastRow <= JOBS_HEADER_ROW) {
    ui.alert("No job rows to review.");
    return;
  }

  const numRows = lastRow - JOBS_HEADER_ROW;
  const data = jobsSheet.getRange(JOBS_DATA_START_ROW, 1, numRows, jobsSheet.getLastColumn()).getValues();
  const calibFontColors = jobsSheet.getRange(JOBS_DATA_START_ROW, calibCol, numRows, 1).getFontColors();

  const items = [];
  const pendingRows = [];

  data.forEach(function(row, i) {
    const note = String(row[calibCol - 1] || "").trim();
    if (!note) return;
    if (calibFontColors[i][0] === CALIBRATION_PROCESSED_FONT_COLOR) return; // already ingested

    pendingRows.push(i + JOBS_DATA_START_ROW);
    items.push({
      company: companyCol ? String(row[companyCol - 1] || "").trim() : "",
      role: roleCol ? String(row[roleCol - 1] || "").trim() : "",
      recommendation: recommendationCol ? String(row[recommendationCol - 1] || "").trim() : "",
      atsScore: atsCol ? String(row[atsCol - 1] || "").trim() : "",
      responseProbability: responseProbCol ? String(row[responseProbCol - 1] || "").trim() : "",
      note: note,
      jobDescription: jobDescCol ? String(row[jobDescCol - 1] || "").substring(0, 2000) : ""
    });
  });

  if (!items.length) {
    ui.alert("No new calibration notes to process.");
    return;
  }

  const config = getConfigValues(configSheet);

  let summaryRow = findConfigRowByExactKey(configSheet, CALIBRATION_SUMMARY_KEY);
  if (!summaryRow) {
    summaryRow = configSheet.getLastRow() + 1;
    configSheet.getRange(summaryRow, 1).setValue(CALIBRATION_SUMMARY_KEY);
  }

  const existingSummary = String(configSheet.getRange(summaryRow, 2).getValue() || "").trim();

  const result = callClaudeForCalibrationSummary(existingSummary, items, config);
  if (!result) {
    ui.alert("Claude did not return a calibration update — check the API key and try again.");
    return;
  }

  if (result.has_new_pattern && result.new_summary) {
    configSheet.getRange(summaryRow, 2).setValue(result.new_summary.trim());
  }

  // Mark reviewed either way — a note that adds nothing beyond standing Config
  // rules has still been evaluated and shouldn't be re-sent on the next run.
  pendingRows.forEach(function(sheetRow) {
    jobsSheet.getRange(sheetRow, calibCol).setFontColor(CALIBRATION_PROCESSED_FONT_COLOR);
  });

  ui.alert(result.has_new_pattern
    ? "Calibration Summary updated from " + items.length + " note(s)."
    : items.length + " note(s) reviewed — already covered by existing Config rules, summary unchanged.");
}

function findConfigRowByExactKey(configSheet, key) {
  const values = configSheet.getDataRange().getValues();
  const target = String(key).trim().toLowerCase();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0] || "").trim().toLowerCase() === target) return i + 1;
  }
  return null;
}

function callClaudeForCalibrationSummary(existingSummary, items, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return null;

  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const toolSchema = {
    name: "calibration_summary",
    description: "Whether the tagged notes reveal a genuinely new calibration pattern, and if so, the rewritten summary",
    input_schema: {
      type: "object",
      required: ["has_new_pattern", "new_summary"],
      properties: {
        has_new_pattern: {
          type: "boolean",
          description: "True only if the tagged note(s) reveal a pattern not already covered by the OTHER Config rows shown in the system prompt above (Fit Score, Response Probability, ATS Score, Compensation Fit, Stability, Recommendation, Next Step rules, etc.) and not already present in the current calibration summary. False if the note is just confirming or restating a rule that already exists elsewhere in Config."
        },
        new_summary: {
          type: "string",
          description: "Only meaningful when has_new_pattern is true: the complete calibration summary, rewritten from scratch to fold in the new pattern. Generalize rather than listing individual jobs. Keep to roughly 500 words — drop stale, superseded, or redundant points. Must NOT restate anything already covered by the other Config rules visible in the system prompt — this summary exists only to capture nuances those standing rules don't already handle. If has_new_pattern is false, return the existing summary unchanged (or an empty string if it was already empty)."
        }
      }
    }
  };

  const itemsText = items.map(function(item, i) {
    return (
      "Tagged job " + (i + 1) + ":\n" +
      "Company: " + item.company + "\n" +
      "Role: " + item.role + "\n" +
      "Recommendation given: " + item.recommendation + "\n" +
      "ATS Score: " + item.atsScore + "\n" +
      "Response Probability: " + item.responseProbability + "\n" +
      "Calibration note: " + item.note + "\n" +
      "Job description excerpt: " + item.jobDescription
    );
  }).join("\n\n");

  const userPrompt =
    "CURRENT calibration summary (used to guide job scoring, separate from the other Config rules already shown above in the system prompt):\n\n" +
    (existingSummary || "(none yet)") + "\n\n" +
    "The user tagged the job(s) below with a calibration note reflecting whether the analysis was good or bad, and why. " +
    "This summary must NEVER restate rules that already exist elsewhere in Config (Fit Score definitions/scale, Response " +
    "Probability calibration, ATS Score rules, Compensation Fit, Stability, Recommendation logic, Next Step logic, industry " +
    "targeting, etc. — all shown in the system prompt above) — those already apply on their own. Only add something here if " +
    "the tagged note reveals a genuinely new nuance, edge case, or correction those standing rules don't already cover. " +
    "If the note is simply an instance of a rule that already exists, set has_new_pattern to false and change nothing. " +
    "If it does reveal something new, rewrite the summary to fold it in, generalized rather than tied to this one job, " +
    "keeping the whole thing to roughly 500 words by dropping anything stale or superseded.\n\n" +
    itemsText;

  const payload = {
    model: model,
    max_tokens: 1536,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "calibration_summary" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) return null;

  const data = JSON.parse(response.getContentText());
  for (const block of (data.content || [])) {
    if (block.type === "tool_use" && block.name === "calibration_summary") {
      return block.input;
    }
  }

  return null;
}

// ---- Button wrappers (assigned to drawing buttons above columns) ----

function btnRunTailoring() {
  const check = getButtonRowCheck_();
  if (!check) return;
  const { headers, rowData, jobDescCol } = check;
  const jobDesc = jobDescCol ? String(rowData[jobDescCol - 1] || "").trim() : "";
  if (!jobDesc) {
    SpreadsheetApp.getUi().alert(
      "No job description found.\n\nCapture this job via the bookmarklet first, then run Tailoring."
    );
    return;
  }
  generateTailoringAndCoverLetter("tailoring");
}

function btnCreateDrafts() {
  const check = getButtonRowCheck_();
  if (!check) return;
  const { headers, rowData, rowFormulas } = check;
  const tailoringCol = getColumnIndexByHeader(headers, "Resume Tailoring Notes");
  const notesUrl = tailoringCol
    ? extractUrlFromFormula(rowFormulas[tailoringCol - 1], rowData[tailoringCol - 1])
    : "";
  if (!notesUrl) {
    SpreadsheetApp.getUi().alert(
      "No tailoring notes found.\n\nRun Tailoring first before creating letter and resume drafts."
    );
    return;
  }
  generateCoverLetterAndResumeParallel();
}

function btnProofreadResume() {
  const check = getButtonRowCheck_();
  if (!check) return;
  const { headers, rowData, rowFormulas } = check;
  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");
  const draftUrl = tailoredDraftCol
    ? extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowData[tailoredDraftCol - 1])
    : "";
  if (!draftUrl) {
    SpreadsheetApp.getUi().alert(
      "No tailored resume found.\n\nRun Create Drafts first before proofreading."
    );
    return;
  }
  proofreadResumeForSelectedRow();
}

function btnInterviewPrep() {
  const check = getButtonRowCheck_();
  if (!check) return;
  const { headers, rowData, jobDescCol } = check;
  const jobDesc = jobDescCol ? String(rowData[jobDescCol - 1] || "").trim() : "";
  if (!jobDesc) {
    SpreadsheetApp.getUi().alert(
      "No job description found.\n\nCapture this job via the bookmarklet first, then run Interview Prep."
    );
    return;
  }
  generateInterviewPrep();
}

function getButtonRowCheck_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== JOBS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Select a row in the Jobs sheet first.");
    return null;
  }
  const row = sheet.getActiveRange().getRow();
  if (row <= JOBS_HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a job row first (click a cell in the row you want to process).");
    return null;
  }
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];
  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  return { sheet, row, headers, rowData, rowFormulas, jobDescCol };
}

// ---- Tailoring & cover letter ----

function generateTailoringAndCoverLetter(mode) {
  // mode: "tailoring" = notes only, "cover_letter" = cover letter only, anything else = both
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== JOBS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Select a row in the Jobs sheet first.");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= JOBS_HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a job row, not the header row.");
    return;
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const config = getConfigForRuntime();

  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const tailoringCol = getColumnIndexByHeader(headers, "Resume Tailoring Notes");
  const coverLetterCol = getColumnIndexByHeader(headers, "Cover Letter Draft");
  const referralCol = getColumnIndexByHeader(headers, "Referral / Contact");
  const tailoredDraftColTCL = getColumnIndexByHeader(headers, "Tailored Resume Draft");

  if (!tailoringCol && !coverLetterCol) {
    SpreadsheetApp.getUi().alert("Add 'Resume Tailoring Notes' and 'Cover Letter Draft' column headers to the Jobs sheet first.");
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];

  const jobDescription = jobDescCol ? String(rowData[jobDescCol - 1] || "").trim() : "";
  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role = roleCol ? String(rowData[roleCol - 1] || "").trim() : "";
  const hiringManager = referralCol ? String(rowData[referralCol - 1] || "").trim() : "";
  const tailoredUrl = tailoredDraftColTCL
    ? extractUrlFromFormula(rowFormulas[tailoredDraftColTCL - 1], rowData[tailoredDraftColTCL - 1])
    : "";
  const resumeUrl = tailoredUrl || config["Base Resume URL"] || "";

  const existingTailoringNotesUrl = tailoringCol
    ? extractUrlFromFormula(rowFormulas[tailoringCol - 1], rowData[tailoringCol - 1])
    : "";
  const existingCoverLetterUrl = coverLetterCol
    ? extractUrlFromFormula(rowFormulas[coverLetterCol - 1], rowData[coverLetterCol - 1])
    : "";
  const existingTailoringNotesText = existingTailoringNotesUrl
    ? (getResumeText(existingTailoringNotesUrl) || "")
    : (tailoringCol ? String(rowData[tailoringCol - 1] || "") : "");

  // Re-running "Run Tailoring" replaces the AI-authored notes but must not wipe out anything
  // the user added under the Manual Notes marker below them.
  const existingManualNotes = extractManualTailoringNotes(existingTailoringNotesText);

  // "Create Drafts" (mode "cover_letter") reads the full reviewed notes doc as context.
  // "Run Tailoring" (mode "tailoring"/both) is a fresh AI pass — but it must still see any
  // carried-over Manual Notes, or it has no way to know facts only the user recorded (e.g.
  // hands-on tool experience) and can end up writing a "gap" note that directly contradicts the
  // manual note sitting right below it in the same doc.
  const existingTailoringNotes = mode === "cover_letter" ? existingTailoringNotesText : existingManualNotes;

  if (!jobDescription) {
    SpreadsheetApp.getUi().alert(
      "No job description stored for this row.\n\n" +
      "Job descriptions are stored automatically for jobs added after the v1.2 update. " +
      "For older rows, paste the job description into the Job Description column manually."
    );
    return;
  }

  const actionLabel = mode === "tailoring" ? "Generate Tailoring Notes"
    : mode === "cover_letter" ? "Generate Cover Letter"
    : "Generate Tailoring & Cover Letter";
  if (!confirmRowAction(sheet, row, actionLabel)) return;

  const toastMsg = mode === "tailoring"
    ? "Generating tailoring notes — this may take 15–30 seconds..."
    : mode === "cover_letter"
    ? "Generating cover letter — this may take 15–30 seconds..."
    : "Generating tailoring notes and cover letter — this may take 15–30 seconds...";
  ss.toast(toastMsg, "Job Assistant", 60);

  const resumeText = getResumeText(resumeUrl);
  const result = callClaudeForTailoringAndCoverLetter(jobDescription, company, role, resumeText, config, hiringManager, existingTailoringNotes);

  ss.toast("", "", 1);

  if (!result) {
    SpreadsheetApp.getUi().alert("Generation failed. Check Apps Script execution log for details.");
    return;
  }

  const writeTailoring = mode !== "cover_letter";
  const writeCoverLetter = mode !== "tailoring";

  const tailoringNotesToSave = writeTailoring
    ? buildTailoringNotesWithManualSection(result.tailoring_notes || "", existingManualNotes)
    : (result.tailoring_notes || "");

  const folderUrl = config["Application Docs Folder"] || "";
  if (folderUrl) {
    const docs = saveApplicationDocs(company, role, tailoringNotesToSave, result.cover_letter || "", folderUrl, mode, config);
    if (writeTailoring && tailoringCol && docs.tailoringUrl) {
      trashOldDocIfDifferent(existingTailoringNotesUrl, docs.tailoringUrl);
      sheet.getRange(row, tailoringCol).setFormula('=HYPERLINK("' + docs.tailoringUrl + '","Tailoring Notes")');
    }
    if (writeCoverLetter && coverLetterCol && docs.coverLetterUrl) {
      trashOldDocIfDifferent(existingCoverLetterUrl, docs.coverLetterUrl);
      sheet.getRange(row, coverLetterCol).setFormula('=HYPERLINK("' + docs.coverLetterUrl + '","Cover Letter")');
    }
    const doneMsg = mode === "tailoring"
      ? "Done! Tailoring Notes doc created in your Application Docs folder for " + company + "."
      : mode === "cover_letter"
      ? "Done! Cover Letter doc created in your Application Docs folder for " + company + "."
      : "Done! Google Docs created in your Application Docs folder for " + company + ".";
    SpreadsheetApp.getActiveSpreadsheet().toast(doneMsg, "Job Assistant", 10);
  } else {
    if (writeTailoring && tailoringCol && tailoringNotesToSave) {
      sheet.getRange(row, tailoringCol).setValue(tailoringNotesToSave);
    }
    if (writeCoverLetter && coverLetterCol && result.cover_letter) {
      sheet.getRange(row, coverLetterCol).setValue(result.cover_letter);
    }
    const doneMsg = mode === "tailoring"
      ? "Done! Check Resume Tailoring Notes for row " + row + "."
      : mode === "cover_letter"
      ? "Done! Check Cover Letter Draft for row " + row + "."
      : "Done! Check Resume Tailoring Notes and Cover Letter Draft for row " + row + ". Tip: add 'Application Docs Folder' to Config to save as Drive docs instead.";
    SpreadsheetApp.getActiveSpreadsheet().toast(doneMsg, "Job Assistant", 12);
  }
}

// Marks a manual-entry section at the bottom of the Tailoring Notes doc/cell — a place for
// the user to note things Claude should always factor in (e.g. specific tools/products they use)
// that survive re-running "Run Tailoring", which otherwise regenerates the AI-authored notes
// above this section from scratch.
const TAILORING_MANUAL_NOTES_MARKER = "MANUAL NOTES (preserved across regeneration)";
const TAILORING_MANUAL_NOTES_PLACEHOLDER = "— add anything Claude should always factor in here (e.g. specific tools/products you use):";

// Doesn't assume any particular formatting after the marker — the user might type inline
// (replacing the placeholder text entirely, no newline), on a new line below it, or leave
// it untouched. Strip the placeholder text if it's still there, then strip any leftover
// leading separator punctuation/whitespace; whatever remains is their content.
function extractManualTailoringNotes(existingText) {
  const text = String(existingText || "");
  const idx = text.indexOf(TAILORING_MANUAL_NOTES_MARKER);
  if (idx === -1) return "";

  let afterMarker = text.substring(idx + TAILORING_MANUAL_NOTES_MARKER.length);
  const placeholderIdx = afterMarker.indexOf(TAILORING_MANUAL_NOTES_PLACEHOLDER);
  if (placeholderIdx !== -1) {
    afterMarker = afterMarker.substring(placeholderIdx + TAILORING_MANUAL_NOTES_PLACEHOLDER.length);
  }
  return afterMarker.replace(/^[\s:—-]+/, "").trim();
}

function buildTailoringNotesWithManualSection(aiNotes, manualNotes) {
  return String(aiNotes || "").trim() +
    "\n\n---\n" + TAILORING_MANUAL_NOTES_MARKER + " " + TAILORING_MANUAL_NOTES_PLACEHOLDER + "\n\n" +
    (manualNotes || "");
}

// A plain substring(0, maxLength) truncates from the front, which can silently drop the
// Manual Notes section entirely when the AI-generated notes above it are long enough to push
// it past the cutoff — exactly the case this section exists to prevent (it's meant to survive
// regeneration and be treated as authoritative). Always keeps Manual Notes in full and only
// trims the AI-generated portion above it to fit within maxLength.
function truncateNotesPreservingManualSection(notesText, maxLength) {
  const text = String(notesText || "");
  const markerIndex = text.indexOf(TAILORING_MANUAL_NOTES_MARKER);

  if (markerIndex === -1) return text.substring(0, maxLength);

  const manualSection = text.substring(markerIndex);
  const aiSection = text.substring(0, markerIndex);
  const aiBudget = Math.max(0, maxLength - manualSection.length);

  return aiSection.substring(0, aiBudget) + manualSection;
}

// Regenerating a doc creates a brand-new Drive file rather than overwriting in place (needed
// so a fresh doc is always self-consistent rather than partially edited), which otherwise
// leaves the previous version orphaned in the folder. Trash it once the new one is linked.
function trashOldDocIfDifferent(oldUrl, newUrl) {
  if (!oldUrl || oldUrl === newUrl) return;
  try {
    const oldId = extractGoogleDocId(oldUrl);
    if (oldId) DriveApp.getFileById(oldId).setTrashed(true);
  } catch (e) {
    console.warn("Could not trash old doc (" + oldUrl + "): " + e.message);
  }
}

// One-time cleanup: trashes orphaned pre-auto-cleanup doc versions in the Atlassian folder.
// Safe to delete this function after running it once.
function trashOrphanedAtlassianDocs() {
  const ids = [
    "1RpXpgPuywe1YSkuXZuBcjMsyxvCBDZ21Prl5cU8jfA0",
    "1IIHolPcbF7fL7y5HJa0mQ8RX-XZhEkWIOsgb-XiIWrY",
    "1kNl8Y1ZC2GIcQr6yYwl_rJlDjFjIgp3KQNuJGbweiFE",
    "15y97PKcYVM7z7wNCH-32uZCv-zMBDUUg5Fgs-tvO1Pg"
  ];
  let trashed = 0;
  ids.forEach(function(id) {
    try {
      DriveApp.getFileById(id).setTrashed(true);
      trashed++;
    } catch (e) {
      console.warn("Could not trash " + id + ": " + e.message);
    }
  });
  SpreadsheetApp.getActiveSpreadsheet().toast("Trashed " + trashed + " of " + ids.length + " orphaned docs.", "Job Assistant", 10);
}

function saveApplicationDocs(company, role, tailoringNotes, coverLetter, parentFolderUrl, mode, config) {
  const match = parentFolderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) ||
               parentFolderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Invalid 'Application Docs Folder' URL in Config. Must be a Google Drive folder URL.");

  const parentFolder = DriveApp.getFolderById(match[1]);
  const existing = parentFolder.getFoldersByName(company);
  const companyFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(company);
  const prefix = company + (role ? " — " + role.substring(0, 40) : "");

  const result = {};

  if (mode !== "cover_letter") {
    const tailoringDoc = DocumentApp.create(prefix + " — Tailoring Notes");
    DriveApp.getFileById(tailoringDoc.getId()).moveTo(companyFolder);
    tailoringDoc.getBody().setText(tailoringNotes);
    tailoringDoc.saveAndClose();
    result.tailoringUrl = "https://docs.google.com/document/d/" + tailoringDoc.getId() + "/edit";
  }

  if (mode !== "tailoring") {
    const coverDoc = DocumentApp.create(buildDeliverableFileName(config || {}, company, "Letter"));
    DriveApp.getFileById(coverDoc.getId()).moveTo(companyFolder);
    coverDoc.getBody().setText(coverLetter);
    coverDoc.saveAndClose();
    result.coverLetterUrl = "https://docs.google.com/document/d/" + coverDoc.getId() + "/edit";
  }

  return result;
}

// ---- Interview prep ----

function generateInterviewPrep() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== JOBS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Select a row in the Jobs sheet first.");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= JOBS_HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a job row, not the header row.");
    return;
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const config = getConfigForRuntime();

  const jobDescCol       = getColumnIndexByHeader(headers, "Job Description");
  const companyCol       = getColumnIndexByHeader(headers, "Company");
  const roleCol          = getColumnIndexByHeader(headers, "Role");
  const autoResponseCol  = getColumnIndexByHeader(headers, "Auto Response");
  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");
  const interviewPrepCol = getColumnIndexByHeader(headers, "Interview Prep");

  const rowData     = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];

  const jobDescription = jobDescCol      ? String(rowData[jobDescCol - 1] || "").trim() : "";
  const company         = companyCol      ? String(rowData[companyCol - 1] || "").trim() : "";
  const role             = roleCol         ? String(rowData[roleCol - 1] || "").trim() : "";
  const autoResponse    = autoResponseCol ? String(rowData[autoResponseCol - 1] || "").trim() : "";
  const existingInterviewPrepUrl = interviewPrepCol
    ? extractUrlFromFormula(rowFormulas[interviewPrepCol - 1], rowData[interviewPrepCol - 1])
    : "";

  if (!jobDescription) {
    SpreadsheetApp.getUi().alert(
      "No job description stored for this row.\n\n" +
      "Job descriptions are stored automatically for jobs added after the v1.2 update. " +
      "For older rows, paste the job description into the Job Description column manually."
    );
    return;
  }

  if (!confirmRowAction(sheet, row, "Generate Interview Prep")) return;

  ss.toast("Generating interview prep — this may take 15–30 seconds...", "Job Assistant", 60);

  const tailoredUrl = tailoredDraftCol
    ? extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowData[tailoredDraftCol - 1])
    : "";
  const resumeUrl  = tailoredUrl || config["Base Resume URL"] || "";
  const resumeText = getResumeText(resumeUrl);

  const result = callClaudeForInterviewPrep(jobDescription, company, role, resumeText, autoResponse, config);

  ss.toast("", "", 1);

  if (!result) {
    SpreadsheetApp.getUi().alert("Generation failed. Check Apps Script execution log for details.");
    return;
  }

  const folderUrl = config["Application Docs Folder"] || "";
  if (folderUrl) {
    const docUrl = saveInterviewPrepDoc(company, role, result.interview_prep || "", folderUrl);
    trashOldDocIfDifferent(existingInterviewPrepUrl, docUrl);
    if (interviewPrepCol) {
      sheet.getRange(row, interviewPrepCol).setFormula('=HYPERLINK("' + docUrl + '","Interview Prep")');
    }
    SpreadsheetApp.getActiveSpreadsheet().toast("Done! Interview Prep doc created in your Application Docs folder for " + company + ".", "Job Assistant", 10);
  } else if (interviewPrepCol) {
    sheet.getRange(row, interviewPrepCol).setValue(result.interview_prep || "");
    SpreadsheetApp.getActiveSpreadsheet().toast(
      "Done! Check Interview Prep for row " + row + ". Tip: add 'Application Docs Folder' to Config to save as a Drive doc instead.",
      "Job Assistant", 12
    );
  } else {
    SpreadsheetApp.getUi().alert(
      "Add an 'Interview Prep' column to Jobs, or an 'Application Docs Folder' to Config, to save this output.\n\n" +
      (result.interview_prep || "").substring(0, 1500)
    );
  }
}

function saveInterviewPrepDoc(company, role, content, parentFolderUrl) {
  const match = parentFolderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) ||
               parentFolderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!match) throw new Error("Invalid 'Application Docs Folder' URL in Config. Must be a Google Drive folder URL.");

  const parentFolder  = DriveApp.getFolderById(match[1]);
  const existing      = parentFolder.getFoldersByName(company);
  const companyFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(company);
  const prefix         = company + (role ? " — " + role.substring(0, 40) : "");

  const doc = DocumentApp.create(prefix + " — Interview Prep");
  DriveApp.getFileById(doc.getId()).moveTo(companyFolder);
  doc.getBody().setText(content);
  doc.saveAndClose();
  return "https://docs.google.com/document/d/" + doc.getId() + "/edit";
}

function callClaudeForInterviewPrep(jobDescription, company, role, resumeText, autoResponse, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("Missing Claude API key.");

  const model = config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL;
  const userName = config["User Display Name"] || config["User Full Name"] || "the user";

  const toolSchema = {
    name: "interview_prep_output",
    description: "Interview preparation study sheet",
    input_schema: {
      type: "object",
      required: ["interview_prep"],
      properties: {
        interview_prep: { type: "string" }
      }
    }
  };

  const userPrompt =
    "Build an interview prep study sheet for " + userName + " ahead of an interview for this role.\n\n" +
    "Company: " + company + "\n" +
    "Role: " + role + "\n\n" +
    "Job description:\n" + jobDescription.substring(0, 8000) + "\n\n" +
    (autoResponse
      ? "Interview-related email activity on this application (may include the interview request itself, format, interviewer names, or scheduling details):\n" + autoResponse.substring(0, 3000) + "\n\n"
      : "No interview email content is available yet — base prep entirely on the job description.\n\n") +
    userName + "'s resume:\n" + (resumeText ? resumeText.substring(0, 6000) : "Use " + userName + "'s background from Config.") + "\n\n" +
    "Produce a study sheet with these sections, using clear headers and bullet points (plain text, no markdown symbols):\n\n" +
    fillConfigTemplate(config["Interview Prep Writing Rules"] || "", { USER_NAME: userName });

  const payload = {
    model: model,
    max_tokens: 8192,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "interview_prep_output" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error("Claude interview prep error: " + response.getContentText());
    return null;
  }

  const data = JSON.parse(response.getContentText());
  console.log("DEBUG interviewPrep: stop_reason=" + data.stop_reason + " content_types=" + (data.content || []).map(function(b){return b.type;}).join(",") );
  for (const block of (data.content || [])) {
    if (block.type === "tool_use" && block.name === "interview_prep_output") {
      console.log("DEBUG interviewPrep: input keys=" + JSON.stringify(Object.keys(block.input || {})) + " interview_prep length=" + String((block.input || {}).interview_prep || "").length);
      return block.input;
    }
  }

  return null;
}

// ---- Resume Reformatter ----

function reformatBaseResume() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfigForRuntime();

  const resumeUrl = config["Base Resume URL"];
  if (!resumeUrl) {
    SpreadsheetApp.getUi().alert("No Base Resume URL found in Config.");
    return;
  }

  ss.toast("Reading resume...", "Resume Reformatter", 60);
  const resumeText = getResumeText(resumeUrl);
  if (!resumeText) {
    SpreadsheetApp.getUi().alert("Could not read resume. Check Base Resume URL in Config.");
    return;
  }

  ss.toast("Reformatting with Claude Sonnet — this may take 20–30 seconds...", "Resume Reformatter", 60);
  let result;
  try {
    result = callClaudeForResumeReformat(resumeText, config);
  } catch (e) {
    SpreadsheetApp.getUi().alert("Claude API error:\n\n" + e.message);
    return;
  }
  if (!result) {
    SpreadsheetApp.getUi().alert("Claude returned no structured output. Check Apps Script logs (Extensions → Apps Script → Executions).");
    return;
  }

  ss.toast("Creating Google Doc...", "Resume Reformatter", 60);
  const docUrl = createFormattedResumeDoc(result, config);

  ss.toast(
    "Reformatted resume created: " + docUrl + ". Review it, make any final tweaks, then update Base Resume URL in Config when ready.",
    "Resume Reformatter", 15
  );
}

function callClaudeForResumeReformat(resumeText, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("Missing Claude API key.");

  const toolSchema = {
    name: "resume_structure",
    description: "Structured resume content reformatted for ATS compatibility",
    input_schema: {
      type: "object",
      required: ["name", "contact_line", "sections"],
      properties: {
        name: { type: "string" },
        contact_line: { type: "string", description: "Single line: City, ST | Phone | Email | LinkedIn URL" },
        sections: {
          type: "array",
          items: {
            type: "object",
            required: ["title", "type"],
            properties: {
              title: { type: "string" },
              type: { type: "string", enum: ["paragraph", "skills", "experience", "education"] },
              content: { type: "string", description: "For paragraph and skills types — plain text" },
              entries: {
                type: "array",
                description: "For experience and education types",
                items: {
                  type: "object",
                  properties: {
                    organization: { type: "string" },
                    title: { type: "string" },
                    dates: { type: "string" },
                    location: { type: "string" },
                    bullets: { type: "array", items: { type: "string" } }
                  }
                }
              }
            }
          }
        }
      }
    }
  };

  const userPrompt =
    "Extract and reformat this resume into a clean ATS-compatible structure.\n\n" +
    "Rules:\n" +
    "- Preserve ALL original wording exactly — do not paraphrase, rewrite, or improve any content\n" +
    "- Convert all column and table-based layouts to flat single-column structure\n" +
    "- Core Competencies or Skills: output as a single comma-separated line, type='skills'\n" +
    "- Each job role is a separate entry with organization, title, dates, location, and bullets array\n" +
    "- Preserve every bullet point verbatim\n" +
    "- Contact info: single pipe-separated line (City, ST | Phone | Email | LinkedIn)\n\n" +
    "Resume text (extracted from a table-based Google Doc — may have minor formatting artifacts):\n\n" +
    resumeText;

  const payload = {
    model: config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL,
    max_tokens: 4096,
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "resume_structure" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error("Claude reformat error: " + response.getContentText());
    throw new Error("Claude API error " + response.getResponseCode() + ": " + response.getContentText().substring(0, 200));
  }

  const data = JSON.parse(response.getContentText());
  if (!data.content) return null;
  for (const block of data.content) {
    if (block.type === "tool_use" && block.name === "resume_structure") return block.input;
  }
  console.error("No resume_structure tool block found: " + JSON.stringify(data).substring(0, 500));
  return null;
}

function createFormattedResumeDoc(data, config) {
  const doc = DocumentApp.create((data.name || "Resume") + " (Reformatted)");
  const file = DriveApp.getFileById(doc.getId());

  const folderUrl = config["Application Docs Folder"];
  if (folderUrl) {
    const match = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) || folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (match) {
      try {
        const folder = DriveApp.getFolderById(match[1]);
        folder.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      } catch(e) { console.warn("Could not move doc to folder: " + e.message); }
    }
  }

  const body = doc.getBody();
  body.clear();
  body.editAsText().setFontFamily("Arial").setFontSize(11);

  const namePara = body.appendParagraph(data.name || "");
  namePara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
  namePara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  namePara.editAsText().setBold(true).setFontSize(16);

  if (data.contact_line) {
    const contactPara = body.appendParagraph(data.contact_line);
    contactPara.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    contactPara.editAsText().setFontSize(10).setBold(false);
  }

  (data.sections || []).forEach(function(section) {
    body.appendParagraph("");

    const headerPara = body.appendParagraph(section.title.toUpperCase());
    headerPara.setHeading(DocumentApp.ParagraphHeading.NORMAL);
    headerPara.editAsText().setBold(true).setFontSize(11);

    if (section.type === "paragraph" || section.type === "skills") {
      const para = body.appendParagraph(section.content || "");
      para.editAsText().setBold(false).setFontSize(11);

    } else if (section.type === "experience" || section.type === "education") {
      (section.entries || []).forEach(function(entry) {
        const orgLine = [entry.organization, entry.dates].filter(Boolean).join("  |  ");
        const orgPara = body.appendParagraph(orgLine);
        orgPara.editAsText().setBold(true).setFontSize(11);

        if (entry.title || entry.location) {
          const titleLine = [entry.title, entry.location].filter(Boolean).join("  |  ");
          const titlePara = body.appendParagraph(titleLine);
          titlePara.editAsText().setItalic(true).setBold(false).setFontSize(11);
        }

        (entry.bullets || []).forEach(function(bullet) {
          const bulletPara = body.appendParagraph("•  " + bullet);
          bulletPara.editAsText().setItalic(false).setBold(false).setFontSize(11);
          bulletPara.setIndentStart(18);
        });
      });
    }
  });

  doc.saveAndClose();
  return "https://docs.google.com/document/d/" + doc.getId() + "/edit";
}

// Seeds every Config row that's generic scoring/formatting methodology rather than personal
// data — same idempotent append-only pattern as the other addXConfigRows() seed functions.
// Personal fields (name/contact/resume URL/career background/target role/compensation) are
// NOT seeded here; the Setup Wizard collects and writes those directly. Safe to run multiple
// times — only missing keys get added, existing values (including ones you've since edited)
// are never touched.
function seedGenericConfigDefaults() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    SpreadsheetApp.getUi().alert("Missing tab: Config");
    return;
  }

  const rows = [
    ["Config tab rule", "Use Config tab as source of truth for tracker logic, scoring definitions, preferences, resume baseline, and protected/manual columns."],
    ["Hardcoding rule", "Do not hardcode scoring rules unless they are mechanical/safety logic. Read scoring guidance from Config whenever possible."],
    ["Manual protected columns", "Never overwrite Notes or Auto Response (Auto Response is written only by the email monitor). Do not overwrite an Apply Anyway or Discuss recommendation unless the candidate changes it."],

    ["Company column", "Company name only. Example: Superhuman."],
    ["Role column", "Exact job title. Example: Sales Engineer, Superhuman for Education."],
    ["Job Link column", "Full job URL."],
    ["Industry Type column", "Informational classification from job/company. Examples: SaaS, Data & Analytics, BI, AI, Education Technology, Sales Enablement, Data Platform, Other."],
    ["Sector column", "Informational market/vertical. Examples: Higher Education, Enterprise Software, Productivity, Analytics, Data Infrastructure, Customer Success, Sales, Other."],
    ["ATS / Platform column", "Record the application system when visible, such as Ashby, Greenhouse, Lever, Workday, SmartRecruiters, LinkedIn Easy Apply, company site, or unknown."],
    ["Notes column", "Entered by user."],
    ["Referral / Contact column", "Blank unless known. Record referral/contact only for networking, notes, and next action. Do not use it to raise Response Probability."],
    ["Auto Response column", "Written only by the email monitor (checkJobApplicationEmails). Never generate, infer, or overwrite this field from job analysis — always leave it blank on initial scoring."],
    ["Auto Response classification values", "Application Received, Interview Request, Assessment, Offer, Rejected, No Response Needed, Other."],
    ["Status column", "Established values: New, Review, Applied - <date>, Networking, Skip, Rejected - <date>, Closed - <date>, Interview - <date>. Job analysis itself only ever outputs New, Review, Applied, Networking, or Skip (Skip when Recommendation is Skip; Review when Recommendation is Discuss). The date-stamped Rejected/Closed/Interview/Applied values are set later by the email monitor or manually."],
    ["Next Step column", "One clear action, not a paragraph. Allowed sequence: Verify company site, Find Contact, Create Resume/Letter, Apply, Paste Domain, Follow Up, Pass. Discuss and Review if easy are used instead of the main sequence when Recommendation is Discuss or Apply if Easy respectively."],
    ["Next Step decision rule", "If estimated Response Probability is greater than 25%, write the next step. If Response Probability is 25% or lower, write pass unless Recommendation is manually changed to Apply Anyway."],
    ["Next Step Due Date column", "Date next step should be completed by. Use yyyy-MM-dd. Default due dates: Find Contact = today + 2 days; Create Resume/Letter = today + 3 days; Apply = today + 4 days; Follow Up = today + 7 days; pass = blank."],
    ["Find Contact Due Days", "2"],
    ["Create Resume Letter Due Days", "3"],
    ["Apply Due Days", "4"],
    ["Follow Up Due Days", "7"],
    ["Initial Apply Due Days", "7"],
    ["Initial Apply If Easy Due Days", "7"],
    ["Initial Discuss Due Days", "7"],

    ["Duplicate posting rule", "If company, role, and job link duplicate an existing row, mark or skip duplicate rather than scoring it as a separate opportunity."],
    ["Recommendation column", "AI-output values: Apply, Discuss, Apply if Easy, Skip. Apply Anyway is a manual-only override the candidate sets themselves in the sheet — never output it from analysis (it is sanitized to Skip if it appears)."],
    ["Recommendation logic", "Recommend Apply for strong fit/time-worthiness and adequate response chance; Apply if Easy for mixed fit or uncertain time value; Discuss when the analysis surfaces genuine ambiguity or a judgment call worth the candidate's input before proceeding; Skip for weak fit/low response odds unless manually overridden."],
    ["Recommendation/status consistency rule", "Recommendation, Status, and Next Step must be internally consistent. If Next Step is pass, Recommendation should usually be Skip unless manually set to Apply Anyway."],
    ["Apply Anyway rule", "Apply Anyway is manual only. The candidate may set it when they want next steps despite low response probability."],
    ["Apply Anyway override rule", "If the candidate manually changes Recommendation to Apply Anyway, Next Step follows normal sequence even when Response Probability is 25% or lower."],

    ["Fit Score definition", "Fit Score = career/time-worthiness: role alignment, company quality, comp, stability, work arrangement, market fit, and whether applying is worth the candidate's time."],
    ["Fit Score scale", "1–10. Calibrate conservatively: strong fit usually 8; exceptional 9; use 10 only for near-perfect alignment."],
    ["Fit Score 8 rule", "Use 8 for strong alignment with the target role, good background/domain fit, and strong compensation, but without a perfect match or a confirmed hiring advantage."],
    ["Fit Score 9 rule", "Use 9 only when role, company, compensation, remote preference, background fit, and likely response path are all very strong."],
    ["Fit Score 10 rule", "Use 10 rarely. Only for near-perfect fit with outstanding alignment, compensation, company stability, and a strong direct hiring path."],
    ["Fit Score lowers when", "Low comp; weak alignment with target role/domain; unwanted work arrangement; limited relevance to the target role; unstable company; poor match to the candidate's background; weak must-have/domain match."],
    ["Must-have/domain gap cap", "If the candidate lacks a clear must-have domain or core technology, cap Fit Score at 6, ATS Score at 65, and Response Probability at 5–15% unless Training Notes or a known contact justify manual review."],
    ["Weak gap scoring rule", "For weak domain/must-have gaps, lower Fit, ATS, and Response Probability. Be conservative. Do not inflate weak matches."],
    ["Product-specific experience gap rule", "If the role requires hands-on experience with a specific product/platform the candidate lacks, lower Fit, ATS, and Response Probability accordingly."],
    ["Hard location mismatch rule", "If a posting requires a region the candidate cannot relocate to or work remotely from, mark Skip unless remote eligibility is explicit or the candidate manually chooses Apply Anyway."],
    ["Startup consideration rule", "Startups may be considered when compensation, growth potential, and target-role fit are strong, but mark Stability Unknown or Risky unless funding, traction, or customer base is clear."],

    ["ATS Score definition", "ATS Score = resume/job requirement match, not generic keyword vibes."],
    ["ATS Score scale", "1–100 estimate of how well the candidate's resume matches the job description for title, skills, keywords, industry/domain, tools, and stated requirements."],
    ["ATS Score high", "80–100 means strong keyword, skills, title, industry, and requirement match."],
    ["ATS Score medium/low", "60–79 means usable but likely needs tailoring; below 60 means weak ATS match or missing important requirements."],
    ["ATS Score source", "Use the Tailored Resume Draft when one exists for this job; otherwise use the candidate's Base Resume URL from Config."],

    ["Response Probability definition", "Response Probability = likelihood of company/recruiter response without referral."],
    ["Response Probability format", "Use one single percentage range only. Examples: 5–10%, 10–15%, 15–25%, 25–40%. Do not output High, Medium, Low, Unknown, or with-referral values unless truly impossible to estimate."],
    ["Response Probability lowers when", "Weak domain fit; missing must-have requirements; poor title alignment; low compensation fit; heavy location mismatch; crowded/generalist applicant pool; likely stronger competitor profiles."],
    ["Response Probability referral rule", "Response Probability ignores referrals. Never output \"with referral\" probability."],
    ["Cold application response calibration", "For cold applications without a referral or known contact (LinkedIn or company-site), keep Response Probability conservative: 5–10% for weak title/domain match or a crowded/generalist applicant pool; 10–15% for an okay/acceptable match; 15–25% for a strong match with clear domain fit and good resume alignment; 25–40% reserved only for exceptional fit with very strong non-referral evidence (exact title/domain/skill match, verified location fit, compensation meets target, visible/accessible hiring team). Do not use 25–40% just because the title looks like a match — it requires genuinely strong alignment across role, domain, compensation, and location, not title match alone."],
    ["Quick response calibration", "If an application is declined within 24–48 hours, treat it as an early ATS or initial screen signal unless there is evidence of human review. Lower future ATS Score and Response Probability for similar role patterns."],
    ["Company-site verification timing", "Do not penalize LinkedIn-sourced jobs for missing company-site verification during initial scoring. Initial scoring should use the pasted LinkedIn JD. Company-site verification only becomes a scoring/risk factor after the user discusses it here or reaches the apply step."],

    ["Compensation Fit column", "Show full actual compensation details from job posting, preserving every listed zone, band, base, OTE, commission, bonus, and currency. If not provided, estimate likely OTE/base range and prefix with \"Estimated:\"."],
    ["Compensation preservation rule", "When posting contains multiple zones, bands, bases, OTEs, commissions, or bonuses, include every listed range. Never select one range or omit another listed range."],
    ["Unknown compensation rule", "Unknown compensation should lower confidence. Do not assume compensation meets target unless market/title evidence is strong; use Estimated: only when there is a reasonable basis."],

    ["Stability column", "Established values: Strong, Neutral, Risky, Unknown."],
    ["Stability Strong", "Established/stable company or strong growth signals."],
    ["Stability Neutral", "No clear stability risk or strength."],
    ["Stability Risky", "Unclear business stability, weak signals, layoffs, funding risk, or concerning role/company risk."],
    ["Stability Unknown", "Not enough information."],

    ["Tailoring Notes Writing Rules",
      "Provide specific, actionable bullet points for tailoring the candidate's resume to this role. Include: exact keywords and phrases to add, skills and tools to highlight or reframe, achievements to lead with, sections to strengthen. Reference actual job requirements and the candidate's actual experience. No generic advice — be specific to this job.\n\n" +
      "PROFESSIONAL SUMMARY — CRITICAL RULES:\n" +
      "1. The summary must ALWAYS be rewritten to open with the candidate's most relevant differentiator for THIS specific role — not a generic opener.\n" +
      "2. If the job description strongly emphasizes a particular skill, technology, or domain the candidate has real experience in, the first sentence should lead with that experience rather than a generic credential.\n" +
      "3. Broad career achievement credentials (e.g. awards, tenure, past titles) are powerful but must never open the summary. They should appear mid-summary as proof of execution, after establishing primary relevance to the role.\n" +
      "4. Always recommend specific replacement text for the summary opener — do not leave it generic."],
    ["Cover Letter Writing Rules",
      "[Paragraph 1: Open strong — what drew the candidate to this specific role and company. Reference something specific about the company's product, mission, or approach. No filler openers like 'I am excited to apply' or 'I am writing to express my interest.']\n\n" +
      "[Paragraph 2: The candidate's most relevant background — specific achievements, tools, and experience that directly match the job requirements. Include numbers where possible. Must explicitly name-check the candidate's current role (the most recent position on the resume) — never skip straight to older roles even if they sound more relevant to the job description; the reader should know what the candidate is doing right now before hearing about past experience. If the job description specifies years of experience (e.g. '8+ years'), explicitly address how the candidate meets that threshold by adding up their full combination of relevant roles. Never open with a years figure lower than the JD's requirement.]\n\n" +
      "[Paragraph 3: Why this company specifically — something about their product, approach, or market position that connects to the candidate's experience or interests. Personal and specific, not generic.]\n\n" +
      "[Paragraph 4: Brief closing value statement — the combination of skills the candidate brings, and an invitation to discuss.]\n\n" +
      "Confident, specific, results-focused. Reference real achievements with numbers where possible. Address the company and role directly. Write in clear, direct prose — every sentence must be grammatically sound and natural when read aloud. Avoid overly clever, metaphorical, or try-hard constructions. If a phrase would sound odd or forced, rewrite it plainly. Do not borrow or adapt phrasing patterns from other cover letters or examples. LENGTH: The entire cover letter body (salutation through closing) must not exceed 350 words. One page maximum. Cut ruthlessly — every sentence must earn its place."],
    ["Resume Tailoring Change Policy", "Only change bullets that genuinely benefit from rewording for this role. Limit to 8–12 high-impact changes — focus on the most relevant roles."],
    ["Resume Professional Summary Opener Rule", "PROFESSIONAL SUMMARY OPENER (mandatory): Always include a replacement for the first sentence of the professional summary, tailored to what this specific job description emphasizes most. Broad career credentials belong in sentence 2 or later — never as the lead, unless they are themselves the most relevant differentiator for this particular role."],
    ["Resume Keyword Preservation Rule", "KEYWORD PRESERVATION (critical): The following keywords and phrases are currently matching between the base resume and the job description — every one of them MUST appear in the tailored resume: {{KEYWORDS}}. Before finalizing any 'replace' value, verify it contains all matching keywords from the 'find' value. You may rephrase around them but may not drop them. For new_skills_line: the updated Core Competencies line must contain every keyword from the matching list that currently appears there. You may add skills but may not remove existing JD-matching ones. If Core Competencies need adjusting, provide the full updated line in new_skills_line."],

    ["File Naming Pattern", "{{COMPANY}}_{{TYPE}}"],
    ["Interview Prep Writing Rules",
      "1. INTERVIEW SNAPSHOT — likely format/round (phone screen, technical, panel, etc.) and any named interviewers, inferred from the email activity provided. Note if this is unknown.\n\n" +
      "2. WHAT THIS ROLE IS REALLY TESTING FOR — read between the lines of the JD: the top 3-5 things the interviewer will be screening for.\n\n" +
      "3. TALKING POINTS FROM {{USER_NAME}}'s BACKGROUND — 4-6 STAR-format stories (Situation, Task, Action, Result) pulled from the actual resume that map directly to the JD's requirements. Use real numbers from the resume.\n\n" +
      "4. LIKELY QUESTIONS — 8-10 questions this interviewer will probably ask given the role and company, each with a one-line pointer on how {{USER_NAME}} should angle the answer using their real background.\n\n" +
      "5. QUESTIONS TO ASK THEM — 5-6 sharp, specific questions {{USER_NAME}} should ask, tied to something concrete in the job description or company.\n\n" +
      "6. WATCH-OUTS — anything in the JD that could be a gap or objection for {{USER_NAME}}, and how to preempt it honestly.\n\n" +
      "7. KEY TERMS & TECHNOLOGY GLOSSARY — every acronym, tool, technology, internal product name, methodology, or industry-specific term appearing anywhere in the job description or the interview email activity that isn't universally common knowledge. For each term, give a one-to-two sentence plain-language definition and, where relevant, why it matters for this interview. Err on the side of including a term rather than skipping it.\n\n" +
      "Be concrete and specific to this company/role/resume — no generic interview advice."]
  ];

  const lastRow = sheet.getLastRow();
  const existingKeys = lastRow > 0
    ? sheet.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return String(r[0]).trim(); })
    : [];

  let added = 0;
  rows.forEach(function(row) {
    if (existingKeys.indexOf(row[0]) === -1) {
      sheet.appendRow(row);
      added++;
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast(added > 0
    ? "Added " + added + " generic Config row(s). Run the Setup Wizard next to fill in your personal info."
    : "All generic Config rows already exist — no changes made.", "Job Assistant", 12);
}

// Entry point for a friend's first run: seeds the generic methodology rows (idempotent, safe
// to have already run) then opens the personal-info wizard.
// A fresh copy of this template (via File > Make a Copy, or `clasp create-script --type
// sheets`) is a genuinely blank spreadsheet — Jobs/Input/Config don't exist yet. Every other
// function in this file assumes they already do (getSheetByName(...) returning null just
// throws/alerts "Missing tab: ..."), so this has to run before anything else can work. Safe
// to re-run: only creates a tab if it's actually missing, never touches an existing one.
function initializeJobMaverickWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.getRange(1, 1, 1, 2).setValues([["Key", "Value"]]);
    configSheet.setFrozenRows(1);
  }

  let inputSheet = ss.getSheetByName(INPUT_SHEET_NAME);
  if (!inputSheet) {
    inputSheet = ss.insertSheet(INPUT_SHEET_NAME);
    inputSheet.getRange(1, 1, 1, 3).setValues([["Job Link", "Job Description", "Ready to Analyze"]]);
    inputSheet.getRange("C2").insertCheckboxes();
    inputSheet.setFrozenRows(1);
  }

  let jobsSheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!jobsSheet) {
    jobsSheet = ss.insertSheet(JOBS_SHEET_NAME);
    const headers = [
      "ID", "Company", "Role", "Job Link", "Job Description", "Calibration Note", "Email Domain",
      "Fit Score", "ATS Score (Base / Tailored)", "Response Probability (Proj / Final)",
      "Compensation Fit", "Stability", "Recommendation", "Status", "Referral / Contact", "Notes",
      "Resume Tailoring Notes", "Cover Letter Draft", "Tailored Resume Draft", "Proof Notes",
      "Next Step", "Next Step Due Date", "Industry Type", "Sector", "Last Updated",
      "Response Date", "Auto Response", "ATS / Platform"
    ];
    // Row 1-2 reserved for an optional metrics dashboard (see insertMetricsDashboardRow()) —
    // JOBS_HEADER_ROW is hardcoded to 3, so headers go there regardless of whether a dashboard
    // is ever added.
    jobsSheet.getRange(JOBS_HEADER_ROW, 1, 1, headers.length).setValues([headers]);
    jobsSheet.setFrozenRows(JOBS_HEADER_ROW);
  }

  // clasp create-script leaves a default "Sheet1" behind — remove it once real tabs exist.
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  // Defensive fallback for a copy that didn't come from File > Make a Copy on the master
  // template (which already carries this tab) — e.g. a fresh clasp create-script instance.
  // Never overwrites an existing tab, so a friend's own edits to their copy are untouched.
  if (!ss.getSheetByName(INSTRUCTIONS_SHEET_NAME)) {
    createInstructionsSheet();
  }
}

// Builds (or refreshes) the "Instructions" tab — the in-sheet setup/usage guide every friend
// sees as soon as they open their copy, since it's a normal tab and travels with File > Make a
// Copy like any other. Unlike the Config seed functions, this ALWAYS overwrites — it's Rick's
// own authored content, not something a friend's data could conflict with, and re-running it
// (after editing the source below) is how doc updates get refreshed on the master template
// before sharing further. Call directly any time the wording needs updating.
const INSTRUCTIONS_SHEET_NAME = "Instructions";

function createInstructionsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(INSTRUCTIONS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(INSTRUCTIONS_SHEET_NAME, 0);
  } else {
    sheet.clear();
  }

  const sections = [
    { type: "title", text: "Job Maverick — Setup & Usage" },
    { type: "body", text: "Everything below happens in your browser — Google Sheets and the Apps Script editor (which opens inside Google, not a separate program). No installs, no command line." },
    { type: "spacer" },
    { type: "heading", text: "What you'll need" },
    { type: "body", text: "• A Google account with Google Sheets.\n• A Claude API key from console.anthropic.com (Plans & Billing > API keys). This is pay-as-you-go — typical job-search usage costs a few dollars a month, not a subscription.\n• Your resume as a Google Doc (not a PDF/Word file — if you have a PDF, open it in Google Docs first: File > Open > Upload, or paste the text into a new Doc)." },
    { type: "spacer" },
    { type: "heading", text: "1. Get your own copy" },
    { type: "body", text: "If you're viewing someone else's template rather than your own copy, use File > Make a copy first. This gives you your own independent spreadsheet and code — nothing you do here affects the original or anyone else's copy." },
    { type: "spacer" },
    { type: "heading", text: "2. Deploy the web app (one-time, needed for job capture)" },
    { type: "body", text: "1. Extensions > Apps Script.\n2. Deploy (top right) > New deployment.\n3. Click the gear icon next to \"Select type\" > Web app.\n4. Set \"Execute as\" to Me, and \"Who has access\" to Anyone.\n5. Click Deploy. The first time, Google will ask you to authorize the script — click through the permission screens (a \"Google hasn't verified this app\" warning is expected for a personal script you just created yourself, not a real security warning; click Advanced > Go to [project name] (unsafe)).\n6. You don't need to copy the deployment URL shown here — the Setup Wizard fetches it automatically." },
    { type: "spacer" },
    { type: "heading", text: "3. Run the Setup Wizard" },
    { type: "body", text: "Job Assistant > Setup Wizard. Fill in personal info, your resume/docs links, target role, background, and work preferences, then your Claude API key. Nothing needs to be perfect on the first pass — every field is a normal Config cell you can edit later directly in the Config tab. Click Save Setup." },
    { type: "spacer" },
    { type: "heading", text: "4. Install your bookmarklet" },
    { type: "body", text: "Job Assistant > Get Bookmarklet. Drag the blue button into your browser's bookmarks bar — that's your job-capture button. Click it on any job posting page (LinkedIn, a company careers page, etc.) and the description gets sent to your sheet automatically, scored within about 30 seconds.\nIf dragging doesn't work in your browser, the same dialog gives you the code as text — create a new bookmark manually (right-click your bookmarks bar > Add page) and paste it as the URL." },
    { type: "spacer" },
    { type: "heading", text: "5. Turn on the email monitor (optional but recommended)" },
    { type: "body", text: "This checks your inbox for application responses and updates your tracker automatically. In Extensions > Apps Script, use the function dropdown near the top toolbar to run, one at a time: installEmailMonitorTrigger, installExtendedEmailCheckTrigger, and installNewModelCheckTrigger. For each job you apply to, add an Email Domain (e.g. greenhouse.io or the company's own domain) to that row so the monitor knows what to search for." },
    { type: "spacer" },
    { type: "heading", text: "Using the tracker day to day" },
    { type: "body", text: "• Click your bookmarklet on job postings you find.\n• Check the Jobs tab — new rows appear scored (Fit Score, ATS Score, Recommendation) within about 30 seconds of capture.\n• Use the Job Assistant menu for everything else: generating tailored resumes/cover letters, interview prep, sorting, checking email responses.\n• The Config tab is your scoring rulebook — edit any row any time to change how Claude scores or writes for you. Add a new row with a short label and a plain-English instruction to add your own rule; it's picked up automatically on the next analysis." },
    { type: "spacer" },
    { type: "heading", text: "Have a question? Just ask." },
    { type: "body", text: "Job Assistant > Open Claude Chat opens a chat sidebar right in this sheet — ask it anything about how the tracker works, why a job scored the way it did, or how to change a Config rule. It's the fastest way to get unstuck, any time." },
    { type: "spacer" },
    { type: "heading", text: "If something breaks" },
    { type: "body", text: "Check Extensions > Apps Script > Executions (left sidebar) for an error log of the most recent runs. You can also ask Job Assistant > Open Claude Chat, or reach out to whoever shared this template with you." }
  ];

  let row = 1;
  sections.forEach(function(s) {
    if (s.type === "spacer") { row++; return; }
    const cell = sheet.getRange(row, 1);
    cell.setValue(s.text).setWrap(true);
    if (s.type === "title") {
      cell.setFontSize(16).setFontWeight("bold");
    } else if (s.type === "heading") {
      cell.setFontSize(12).setFontWeight("bold").setFontColor("#1a73e8");
    } else {
      cell.setFontSize(11).setFontWeight("normal").setFontColor("#333333");
    }
    row++;
  });

  sheet.setColumnWidth(1, 700);
  sheet.autoResizeRows(1, row - 1);
  sheet.setFrozenRows(1);
  ss.setActiveSheet(sheet);

  SpreadsheetApp.getActiveSpreadsheet().toast("Instructions tab created/refreshed.", "Job Assistant", 8);
}

function viewInstructions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(INSTRUCTIONS_SHEET_NAME);
  if (sheet) {
    ss.setActiveSheet(sheet);
  } else {
    createInstructionsSheet();
  }
}

function showSetupWizard() {
  initializeJobMaverickWorkbook();
  seedGenericConfigDefaults();

  const html = HtmlService.createHtmlOutputFromFile("setup_wizard")
    .setTitle("Job Maverick Setup")
    .setWidth(420);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Writes the Setup Wizard's answers into Config (personal fields only — generic methodology
// rows come from seedGenericConfigDefaults) and into Script Properties (API key, webhook
// token). Auto-generates a webhook token if one doesn't already exist, rather than asking a
// non-technical friend to invent one themselves.
function saveSetupWizardAnswers(formData) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) throw new Error("Missing tab: Config");

  const data = formData || {};
  const set = function(key, value) {
    if (value) setConfigValue(sheet, key, value);
  };

  set("User Full Name", data.fullName);
  set("User Display Name", data.displayName);
  set("User Email", data.email);
  set("User Phone", data.phone);
  set("User Location", data.location);
  set("User LinkedIn", data.linkedin);
  set("Base Resume URL", data.resumeUrl);
  set("Application Docs Folder", data.docsFolderUrl);
  set("Primary role target", data.primaryRole);
  set("Secondary role targets", data.secondaryRoles);
  set("Industries / sectors", data.industries);
  set("High Fit Score earns", data.highFitEarns);
  set("User Background", data.background);
  set("Remote preference", data.remotePref);
  set("Hybrid preference", data.hybridPref);
  set("Onsite", data.onsitePref);
  set("Travel preference", data.travelPref);
  set("Job-search constraint", data.jobSearchConstraint);

  if (data.minComp) {
    set("Compensation target", data.targetComp
      ? (data.minComp + " acceptable; " + data.targetComp + " preferred")
      : (data.minComp + " acceptable"));
    set("Compensation Fit Weak", "Actual listed range is below " + data.minComp + ".");
    set("Compensation Fit Good", "Actual listed range is " + data.minComp + "+ acceptable range.");
    set("Compensation Fit Unknown", "Only use if compensation cannot be found or reasonably estimated.");
  }
  if (data.targetComp) {
    set("Compensation Fit Strong", "Actual listed range meets/exceeds " + data.targetComp + " preferred target.");
  }

  const props = PropertiesService.getScriptProperties();
  if (data.apiKey) props.setProperty("CLAUDE_API_KEY", data.apiKey);

  let token = props.getProperty("WEBHOOK_TOKEN");
  if (!token) {
    token = "job-" + Utilities.getUuid().split("-")[0];
    props.setProperty("WEBHOOK_TOKEN", token);
  }

  return { webhookToken: token };
}

// Payload/response builders are split out from callClaudeForTailoringAndCoverLetter so the
// parallel cover-letter+resume flow (generateCoverLetterAndResumeParallel) can build this exact
// request and fire it concurrently with the ATS-scoring request via UrlFetchApp.fetchAll,
// without duplicating the prompt itself.
function buildTailoringAndCoverLetterPayload(jobDescription, company, role, resumeText, config, hiringManager, existingTailoringNotes) {
  const model = config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL;

  const toolSchema = {
    name: "tailoring_output",
    description: "Resume tailoring notes and cover letter draft",
    input_schema: {
      type: "object",
      required: ["tailoring_notes", "cover_letter"],
      properties: {
        tailoring_notes: { type: "string" },
        cover_letter: { type: "string" }
      }
    }
  };

  const userName     = config["User Display Name"] || config["User Full Name"] || "the user";
  const userFullName = config["User Full Name"]    || config["User Display Name"] || "Your Name";
  const userPhone    = config["User Phone"]    || "";
  const userEmail    = config["User Email"]    || "";
  const userLocation = config["User Location"] || "";
  const userLinkedIn = config["User LinkedIn"] || "";

  const letterhead = [userFullName, userLocation, userPhone, userEmail, userLinkedIn]
    .filter(v => v).join(" | ");

  const userPrompt =
    "Generate resume tailoring notes and a cover letter for " + userName + " for this specific job.\n\n" +
    "Company: " + company + "\n" +
    "Role: " + role + "\n\n" +
    "Job description:\n" + jobDescription.substring(0, 8000) + "\n\n" +
    userName + "'s resume:\n" + (resumeText ? resumeText.substring(0, 6000) : "Use " + userName + "'s background from Config.") + "\n\n" +
    (existingTailoringNotes
      ? "Existing tailoring notes for this job, already reviewed and edited by " + userName + " — treat as authoritative, especially any specific tools, products, or details called out that may not be obvious from the job description alone:\n" + truncateNotesPreservingManualSection(existingTailoringNotes, 3000) + "\n\n"
      : "") +
    "tailoring_notes instructions: " + (config["Tailoring Notes Writing Rules"] || "") + "\n\n" +
    "cover_letter instructions: Write a professional, compelling cover letter in " + userName + "'s voice using EXACTLY this format:\n\n" +
    letterhead + "\n\n" +
    Utilities.formatDate(new Date(), "America/New_York", "MMMM d, yyyy") + "\n\n" +
    (hiringManager ? "[Hiring manager name from contact info below]\n[Their title — infer from job description if possible]\n" : "") +
    "[Company name]\n\n" +
    "RE: [Exact role title]\n\n" +
    (hiringManager
      ? "Dear [First name only extracted from: " + hiringManager + "],\n\n"
      : "Dear Hiring Manager,\n\n") +
    (config["Cover Letter Writing Rules"] || "") + "\n\n" +
    "Thank you for your time and consideration.\n\n" +
    "Sincerely,\n\n" + (config["User Display Name"] || userFullName) + "\n\n" +
    (hiringManager
      ? "The contact field contains: \"" + hiringManager + "\" — extract the person's first name for the salutation and full name for the recipient block. "
      : "");



  return {
    model: model,
    max_tokens: 4096,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "tailoring_output" },
    messages: [{ role: "user", content: userPrompt }]
  };
}

function parseTailoringAndCoverLetterResponse(response) {
  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error("Claude tailoring error: " + response.getContentText());
    return null;
  }

  const data = JSON.parse(response.getContentText());
  for (const block of (data.content || [])) {
    if (block.type === "tool_use" && block.name === "tailoring_output") {
      return block.input;
    }
  }

  return null;
}

function callClaudeForTailoringAndCoverLetter(jobDescription, company, role, resumeText, config, hiringManager, existingTailoringNotes) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("Missing Claude API key.");

  const payload = buildTailoringAndCoverLetterPayload(jobDescription, company, role, resumeText, config, hiringManager, existingTailoringNotes);

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  return parseTailoringAndCoverLetterResponse(response);
}

// ---- Tailored resume generation ----

function generateTailoredResume() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== JOBS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Select a row in the Jobs sheet first.");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= JOBS_HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a job row, not the header row.");
    return;
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const config = getConfigForRuntime();

  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const tailoringNotesCol = getColumnIndexByHeader(headers, "Resume Tailoring Notes");
  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");

  if (!tailoredDraftCol) {
    SpreadsheetApp.getUi().alert("Add a 'Tailored Resume Draft' column header to the Jobs sheet first.");
    return;
  }

  const baseResumeUrl = config["Base Resume URL"] || "";
  const baseDocId = extractGoogleDocId(baseResumeUrl);
  if (!baseDocId) {
    SpreadsheetApp.getUi().alert("Could not find base resume. Check 'Base Resume URL' in Config.");
    return;
  }

  if (!confirmRowAction(sheet, row, "Generate Tailored Resume")) return;

  const resumeText = getResumeText(baseResumeUrl);
  if (!resumeText) {
    SpreadsheetApp.getUi().alert("Could not read base resume. Check 'Base Resume URL' in Config.");
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];
  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role = roleCol ? String(rowData[roleCol - 1] || "").trim() : "";
  const jobDescription = jobDescCol ? String(rowData[jobDescCol - 1] || "").trim() : "";
  const existingTailoredResumeUrl = extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowData[tailoredDraftCol - 1]);

  let tailoringNotes = "";
  if (tailoringNotesCol) {
    const notesUrl = extractUrlFromCellRange(sheet.getRange(row, tailoringNotesCol));
    if (notesUrl) tailoringNotes = getResumeText(notesUrl) || "";
  }

  if (!tailoringNotes) {
    SpreadsheetApp.getUi().alert("No tailoring notes found for this row. Run Generate Tailoring & Cover Letter first for best results. Proceeding with job description only.");
  }

  const folderUrl = config["Application Docs Folder"] || "";
  const folderMatch = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) || folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!folderMatch) {
    SpreadsheetApp.getUi().alert("Set 'Application Docs Folder' in Config to save the tailored resume.");
    return;
  }

  const parentFolder = DriveApp.getFolderById(folderMatch[1]);
  const existing = parentFolder.getFoldersByName(company);
  const companyFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(company);

  const atsCol = getColumnIndexByHeader(headers, "ATS Score");
  const storedAtsRaw = atsCol ? String(rowData[atsCol - 1] || "") : "";
  const storedBaseScore = extractBaseAtsScore(storedAtsRaw);

  ss.toast("Scoring base resume for keywords...", "Job Assistant", 60);
  const baseAtsResult = callClaudeForAtsScore(jobDescription, company, role, resumeText, config);

  finishTailoredResumeGeneration(ss, sheet, row, headers, rowData, config, company, role, jobDescription,
    tailoringNotes, resumeText, baseDocId, companyFolder, existingTailoredResumeUrl, tailoredDraftCol,
    atsCol, baseAtsResult, storedBaseScore);
}

// Extracted from generateTailoredResume's inner buildAndScoreDraft closure so it can also be
// called from generateCoverLetterAndResumeParallel — same doc-build/replace/rescore mechanics,
// just as a standalone function taking its captured context as explicit params.
function buildAndScoreResumeDraft(result, label, baseDocId, companyFolder, config, company, role, jobDescription, resumeText) {
  const prefix = company + (role ? " — " + role.substring(0, 40) : "");
  const fileName = label === "Tailored Resume" ? buildDeliverableFileName(config, company, "Resume") : prefix + " — " + label;
  const copyFile = DriveApp.getFileById(baseDocId).makeCopy(fileName, companyFolder);
  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();
  (result.replacements || []).forEach(function(r) {
    if (r.find && r.replace) {
      try { body.replaceText(escapeRegex(r.find.trim()), r.replace.trim()); }
      catch(e) { console.warn("replaceText failed: " + r.find + " — " + e.message); }
    }
  });
  if (result.new_skills_line) {
    const currentSkills = extractSkillsLine(resumeText);
    if (currentSkills) {
      try { body.replaceText(escapeRegex(currentSkills.trim()), result.new_skills_line.trim()); }
      catch(e) { console.warn("Skills replace failed: " + e.message); }
    }
  }
  doc.saveAndClose();
  const url = "https://docs.google.com/document/d/" + copyFile.getId() + "/edit";
  const text = getResumeText(url);
  const atsResult = text ? callClaudeForAtsScore(jobDescription, company, role, text, config) : null;
  return { copyFile: copyFile, url: url, score: atsResult ? atsResult.score : null };
}

// Everything after the base ATS score is available: generate the tailored edits, build/score
// the draft, retry if the score dropped, save, and recalculate Response Probability. Shared by
// generateTailoredResume (sequential) and generateCoverLetterAndResumeParallel (fires the base
// ATS-scoring call concurrently with the cover-letter call, then hands off here) so the actual
// doc-generation logic exists in exactly one place.
function finishTailoredResumeGeneration(ss, sheet, row, headers, rowData, config, company, role, jobDescription,
    tailoringNotes, resumeText, baseDocId, companyFolder, existingTailoredResumeUrl, tailoredDraftCol,
    atsCol, baseAtsResult, storedBaseScore) {
  const baseKeywords = baseAtsResult ? baseAtsResult.keywords : [];
  // Always use the stored base score as the canonical display value — never overwrite it with a re-scored value
  const baseScore = storedBaseScore !== null ? storedBaseScore : (baseAtsResult ? baseAtsResult.score : null);

  ss.toast("Generating targeted edits — pass 1...", "Job Assistant", 120);

  let result;
  try {
    result = callClaudeForTailoredResume(jobDescription, tailoringNotes, resumeText, company, role, config, null, baseKeywords);
  } catch(e) {
    ss.toast("", "", 1);
    SpreadsheetApp.getUi().alert("Claude API error:\n\n" + e.message);
    return;
  }
  if (!result) {
    ss.toast("", "", 1);
    SpreadsheetApp.getUi().alert("Generation failed. Check Apps Script execution log.");
    return;
  }

  ss.toast("Scoring pass 1...", "Job Assistant", 60);
  let draft = buildAndScoreResumeDraft(result, "Tailored Resume", baseDocId, companyFolder, config, company, role, jobDescription, resumeText);
  let bestFile = draft.copyFile;
  let finalUrl = draft.url;
  let finalScore = draft.score;

  const MAX_TAILORING_RETRIES = 3;
  for (let attempt = 1; attempt < MAX_TAILORING_RETRIES; attempt++) {
    if (finalScore === null || baseScore === null || finalScore >= baseScore) break;

    ss.toast("Score dropped (" + baseScore + " → " + finalScore + ") — retry " + attempt + " of " + (MAX_TAILORING_RETRIES - 1) + "...", "Job Assistant", 120);

    const atsFeedback =
      "CRITICAL: The previous tailoring attempt reduced the ATS score from " + baseScore + " to " + finalScore + ". " +
      "This is not acceptable. The tailored resume must score AT LEAST " + baseScore + ". " +
      "Go through each replacement pair you are about to produce and verify: does the 'replace' value contain every word from 'find' that also appears in the job description? " +
      "If not, revise the replacement until it does. " +
      "The most common cause of score drops is removing matching keywords from Core Competencies or bullets — do not do this under any circumstances.";

    let retryResult;
    try {
      retryResult = callClaudeForTailoredResume(jobDescription, tailoringNotes, resumeText, company, role, config, atsFeedback, baseKeywords);
    } catch(e) {
      retryResult = null;
    }

    if (retryResult) {
      ss.toast("Scoring retry " + attempt + "...", "Job Assistant", 60);
      const retryDraft = buildAndScoreResumeDraft(retryResult, "Tailored Resume", baseDocId, companyFolder, config, company, role, jobDescription, resumeText);
      if (retryDraft.score !== null && (finalScore === null || retryDraft.score > finalScore)) {
        bestFile.setTrashed(true);
        bestFile = retryDraft.copyFile;
        finalUrl = retryDraft.url;
        finalScore = retryDraft.score;
      } else {
        retryDraft.copyFile.setTrashed(true);
      }
    }
  }

  if (finalScore !== null && baseScore !== null && finalScore < baseScore) {
    SpreadsheetApp.getUi().alert(
      "Warning: tailored resume scored " + finalScore + " vs base " + baseScore + " after " + (MAX_TAILORING_RETRIES - 1) + " retries.\n\n" +
      "The best result has been saved. Review the tailoring notes — the JD keywords may conflict with the current resume structure."
    );
  }

  trashOldDocIfDifferent(existingTailoredResumeUrl, finalUrl);
  sheet.getRange(row, tailoredDraftCol).setFormula('=HYPERLINK("' + finalUrl + '","Tailored Resume Draft")');

  if (atsCol && baseScore !== null && finalScore !== null) {
    sheet.getRange(row, atsCol).setValue(baseScore + " / " + finalScore);
  }

  // Recalculate Response Probability using tailored resume
  const respProbCol = getColumnIndexByHeader(headers, "Response Probability");
  if (respProbCol) {
    ss.toast("Updating response probability...", "Job Assistant", 30);
    const tailoredText = getResumeText(finalUrl);
    const tailoredProb = tailoredText ? callClaudeForResponseProbability(jobDescription, company, role, tailoredText, config, baseScore, finalScore) : null;
    if (tailoredProb) {
      const currentProb = String(rowData[respProbCol - 1] || "").trim();
      const baseProb = extractBaseResponseProbability(currentProb) || currentProb;
      if (baseProb && tailoredProb !== baseProb) {
        sheet.getRange(row, respProbCol).setValue(baseProb + " / " + tailoredProb);
      }
    }
  }

  ss.toast("", "", 1);

  const scoreNote = (finalScore !== null && baseScore !== null)
    ? " ATS: " + baseScore + " → " + finalScore + "."
    : "";
  ss.toast("Done!" + scoreNote + " Review the draft and run Proofread Resume when ready.", "Job Assistant", 12);
}

// Combined "Create Drafts" path: the cover letter call and the resume's base ATS-scoring call
// are independent of each other (neither's prompt depends on the other's result), so they're
// fired concurrently via UrlFetchApp.fetchAll instead of back-to-back — saves the wall-clock
// time of one full Claude round-trip. The tailored-resume EDIT call still has to wait for the
// ATS score (it needs the matched-keyword list), so it remains sequential after the fetchAll —
// same as the rest of the retry/rescore logic in finishTailoredResumeGeneration.
function generateCoverLetterAndResumeParallel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== JOBS_SHEET_NAME) {
    SpreadsheetApp.getUi().alert("Select a row in the Jobs sheet first.");
    return;
  }

  const row = sheet.getActiveRange().getRow();
  if (row <= JOBS_HEADER_ROW) {
    SpreadsheetApp.getUi().alert("Select a job row, not the header row.");
    return;
  }

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const config = getConfigForRuntime();

  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol = getColumnIndexByHeader(headers, "Role");
  const tailoringCol = getColumnIndexByHeader(headers, "Resume Tailoring Notes");
  const coverLetterCol = getColumnIndexByHeader(headers, "Cover Letter Draft");
  const referralCol = getColumnIndexByHeader(headers, "Referral / Contact");
  const tailoredDraftCol = getColumnIndexByHeader(headers, "Tailored Resume Draft");

  if (!coverLetterCol || !tailoredDraftCol) {
    SpreadsheetApp.getUi().alert("Add 'Cover Letter Draft' and 'Tailored Resume Draft' column headers to the Jobs sheet first.");
    return;
  }

  const baseResumeUrl = config["Base Resume URL"] || "";
  const baseDocId = extractGoogleDocId(baseResumeUrl);
  if (!baseDocId) {
    SpreadsheetApp.getUi().alert("Could not find base resume. Check 'Base Resume URL' in Config.");
    return;
  }

  const folderUrl = config["Application Docs Folder"] || "";
  const folderMatch = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) || folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (!folderMatch) {
    SpreadsheetApp.getUi().alert("Set 'Application Docs Folder' in Config first — this combined action saves both docs there.");
    return;
  }

  if (!confirmRowAction(sheet, row, "Create Cover Letter + Resume")) return;

  const resumeText = getResumeText(baseResumeUrl);
  if (!resumeText) {
    SpreadsheetApp.getUi().alert("Could not read base resume. Check 'Base Resume URL' in Config.");
    return;
  }

  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const rowFormulas = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getFormulas()[0];

  const jobDescription = jobDescCol ? String(rowData[jobDescCol - 1] || "").trim() : "";
  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role = roleCol ? String(rowData[roleCol - 1] || "").trim() : "";
  const hiringManager = referralCol ? String(rowData[referralCol - 1] || "").trim() : "";
  const existingCoverLetterUrl = coverLetterCol
    ? extractUrlFromFormula(rowFormulas[coverLetterCol - 1], rowData[coverLetterCol - 1])
    : "";
  const existingTailoredResumeUrl = extractUrlFromFormula(rowFormulas[tailoredDraftCol - 1], rowData[tailoredDraftCol - 1]);

  if (!jobDescription) {
    SpreadsheetApp.getUi().alert(
      "No job description stored for this row.\n\n" +
      "Job descriptions are stored automatically for jobs added after the v1.2 update. " +
      "For older rows, paste the job description into the Job Description column manually."
    );
    return;
  }

  const existingTailoringNotesUrl = tailoringCol
    ? extractUrlFromFormula(rowFormulas[tailoringCol - 1], rowData[tailoringCol - 1])
    : "";
  const tailoringNotesText = existingTailoringNotesUrl
    ? (getResumeText(existingTailoringNotesUrl) || "")
    : (tailoringCol ? String(rowData[tailoringCol - 1] || "") : "");

  if (!tailoringNotesText) {
    SpreadsheetApp.getUi().alert("No tailoring notes found for this row. Run Tailoring first before creating letter and resume drafts.");
    return;
  }

  const parentFolder = DriveApp.getFolderById(folderMatch[1]);
  const existingFolder = parentFolder.getFoldersByName(company);
  const companyFolder = existingFolder.hasNext() ? existingFolder.next() : parentFolder.createFolder(company);

  const atsCol = getColumnIndexByHeader(headers, "ATS Score");
  const storedAtsRaw = atsCol ? String(rowData[atsCol - 1] || "") : "";
  const storedBaseScore = extractBaseAtsScore(storedAtsRaw);

  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) {
    SpreadsheetApp.getUi().alert("Missing Claude API key.");
    return;
  }

  const coverLetterPayload = buildTailoringAndCoverLetterPayload(jobDescription, company, role, resumeText, config, hiringManager, tailoringNotesText);
  const atsPayload = buildAtsScorePayload(jobDescription, company, role, resumeText, config);

  const requestHeaders = {
    "x-api-key": apiKey,
    "anthropic-version": "2023-06-01",
    "anthropic-beta": "prompt-caching-2024-07-31"
  };

  ss.toast("Generating cover letter and scoring base resume in parallel...", "Job Assistant", 60);

  let responses;
  try {
    responses = UrlFetchApp.fetchAll([
      { url: "https://api.anthropic.com/v1/messages", method: "post", contentType: "application/json", headers: requestHeaders, payload: JSON.stringify(coverLetterPayload), muteHttpExceptions: true },
      { url: "https://api.anthropic.com/v1/messages", method: "post", contentType: "application/json", headers: requestHeaders, payload: JSON.stringify(atsPayload), muteHttpExceptions: true }
    ]);
  } catch (e) {
    ss.toast("", "", 1);
    SpreadsheetApp.getUi().alert("Claude API error:\n\n" + e.message);
    return;
  }

  ss.toast("", "", 1);

  const coverLetterResult = parseTailoringAndCoverLetterResponse(responses[0]);
  const baseAtsResult = parseAtsScoreResponse(responses[1]);

  if (!coverLetterResult) {
    SpreadsheetApp.getUi().alert("Cover letter generation failed. Check Apps Script execution log for details.");
  } else {
    const docs = saveApplicationDocs(company, role, "", coverLetterResult.cover_letter || "", folderUrl, "cover_letter", config);
    if (docs.coverLetterUrl) {
      trashOldDocIfDifferent(existingCoverLetterUrl, docs.coverLetterUrl);
      sheet.getRange(row, coverLetterCol).setFormula('=HYPERLINK("' + docs.coverLetterUrl + '","Cover Letter")');
    }
  }

  finishTailoredResumeGeneration(ss, sheet, row, headers, rowData, config, company, role, jobDescription,
    tailoringNotesText, resumeText, baseDocId, companyFolder, existingTailoredResumeUrl, tailoredDraftCol,
    atsCol, baseAtsResult, storedBaseScore);
}

function escapeRegex(str) {
  return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function extractSkillsLine(resumeText) {
  const lines = resumeText.split('\n');
  let afterHeader = false;
  for (const line of lines) {
    const trimmed = line.trim();
    if (/core competencies/i.test(trimmed)) { afterHeader = true; continue; }
    if (afterHeader && trimmed.length > 0) return trimmed;
  }
  return null;
}

function callClaudeForTailoredResume(jobDescription, tailoringNotes, resumeText, company, role, config, atsFeedback, matchedKeywords) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("Missing Claude API key.");

  const toolSchema = {
    name: "resume_edits",
    description: "Targeted text replacements to tailor the base resume for this specific role",
    input_schema: {
      type: "object",
      required: ["replacements"],
      properties: {
        replacements: {
          type: "array",
          items: {
            type: "object",
            required: ["find", "replace"],
            properties: {
              find: { type: "string", description: "Exact text from the resume to replace — no bullet character, verbatim match required" },
              replace: { type: "string", description: "New tailored text — no bullet character" }
            }
          }
        },
        new_skills_line: {
          type: "string",
          description: "Optional — full updated comma-separated Core Competencies line if skills need adjusting"
        }
      }
    }
  };

  const userName = config["User Display Name"] || config["User Full Name"] || "the user";
  const userPrompt =
    "Generate targeted find/replace edits to tailor " + userName + "'s base resume for this role.\n\n" +
    "Company: " + company + "\n" +
    "Role: " + role + "\n\n" +
    "Tailoring notes:\n" + truncateNotesPreservingManualSection(tailoringNotes, 4000) + "\n\n" +
    "Job description:\n" + (jobDescription || "").substring(0, 4000) + "\n\n" +
    "Base resume (exact text — your 'find' values must match this verbatim):\n" + resumeText.substring(0, 8000) + "\n\n" +
    "Rules:\n" +
    "- 'find' must be an exact substring of the resume text above — copy it character for character\n" +
    "- Do not include the bullet character (•) in find or replace values\n" +
    "- Do not invent experience, skills, or metrics " + userName + " does not have\n" +
    "- Preserve all dates, company names, titles, and real metrics\n" +
    "- " + (config["Resume Tailoring Change Policy"] || "") + "\n" +
    "- " + (config["Resume Professional Summary Opener Rule"] || "") + "\n" +
    "- " + fillConfigTemplate(config["Resume Keyword Preservation Rule"], {
      KEYWORDS: (matchedKeywords && matchedKeywords.length > 0) ? matchedKeywords.join(", ") : "(score base resume first to populate this list)"
    }) +
    (atsFeedback ? "\n\n" + atsFeedback : "");

  const payload = {
    model: config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL,
    max_tokens: 4096,
    system: [{ type: "text", text: buildConfigSystemPrompt(config), cache_control: { type: "ephemeral" } }],
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "resume_edits" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
      "anthropic-beta": "prompt-caching-2024-07-31"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    throw new Error("Claude API error " + response.getResponseCode() + ": " + response.getContentText().substring(0, 200));
  }

  const data = JSON.parse(response.getContentText());
  if (!data.content) return null;
  for (const block of data.content) {
    if (block.type === "tool_use" && block.name === "resume_edits") return block.input;
  }
  return null;
}

// ---- Resume proofreading ----

function proofreadResumeForRow(sheet, row, headers, resumeUrl) {
  const proofNotesCol = getColumnIndexByHeader(headers, "Proof Notes");
  if (!proofNotesCol) return;

  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol    = getColumnIndexByHeader(headers, "Role");
  const idCol      = getColumnIndexByHeader(headers, "ID");
  const rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  const company = companyCol ? String(rowData[companyCol - 1] || "").trim() : "";
  const role    = roleCol    ? String(rowData[roleCol - 1]    || "").trim() : "";
  const id      = idCol      ? String(rowData[idCol - 1]      || "").trim() : "";

  const config = getConfigForRuntime();
  const resumeText = getResumeText(resumeUrl);

  if (!resumeText) {
    sheet.getRange(row, proofNotesCol).setValue("Could not read resume doc at: " + resumeUrl);
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast("Proofreading resume...", "Job Assistant", 60);

  const issues = callClaudeForProofread(resumeText, company, role, config);
  const content = issues || "No issues found.";

  // Save as .txt in company folder and link in Proof Notes
  try {
    const folderUrl = config["Application Docs Folder"] || "";
    const folderMatch = folderUrl.match(/folders\/([a-zA-Z0-9_-]+)/) || folderUrl.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    if (folderMatch && company) {
      const parentFolder = DriveApp.getFolderById(folderMatch[1]);
      const existing = parentFolder.getFoldersByName(company);
      const companyFolder = existing.hasNext() ? existing.next() : parentFolder.createFolder(company);
      const fileName = company + (role ? " — " + role.substring(0, 40) : "") + " — Proof Notes.txt";
      const file = companyFolder.createFile(fileName, content, MimeType.PLAIN_TEXT);
      const fileUrl = "https://drive.google.com/file/d/" + file.getId() + "/view";
      sheet.getRange(row, proofNotesCol).setFormula('=HYPERLINK("' + fileUrl + '","Proof Notes")');
    } else {
      sheet.getRange(row, proofNotesCol).setValue(content);
    }
  } catch(e) {
    console.warn("Could not save proof notes to Drive: " + e.message);
    sheet.getRange(row, proofNotesCol).setValue(content);
  }

  openTrainingSidebar(id, company, role, content);
}

function callClaudeForProofread(resumeText, company, role, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return "Missing API key.";

  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const userPrompt =
    "Proofread this resume for " + (config["User Display Name"] || config["User Full Name"] || "the user") + "'s application to " + (role || "a role") + " at " + (company || "a company") + ".\n\n" +
    "Check for:\n" +
    "- Spelling errors\n" +
    "- Grammar mistakes\n" +
    "- Inconsistent formatting or punctuation\n" +
    "- Awkward or unclear phrasing\n" +
    "- Factual inconsistencies (overlapping dates, mismatched titles or company names). NOTE: consecutive roles at the same company where one ends in Month X and the next begins in Month X+1 (e.g. July 2022 → August 2022) are promotions — do NOT flag these as overlaps.\n\n" +
    "Return ONLY a concise bulleted list of specific issues with the exact text and suggested fix. " +
    "If no issues are found, return exactly: No issues found.\n\n" +
    "Resume:\n" + resumeText.substring(0, 12000);

  const payload = {
    model: model,
    max_tokens: 1024,
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    console.error("Proofread error: " + response.getContentText());
    return "Proofread failed — check execution log.";
  }

  const data = JSON.parse(response.getContentText());
  for (const block of (data.content || [])) {
    if (block.type === "text") return block.text;
  }

  return "No response from Claude.";
}

// ---- Training Tab ----

function getOrCreateTrainingSummarySheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(TRAINING_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(TRAINING_SHEET_NAME);
    sheet.getRange(1, 1, 1, 4).setValues([["ID / Company", "Letter Changes", "Resume Changes", "Proof Notes"]]);
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 220);
    sheet.setColumnWidth(2, 280);
    sheet.setColumnWidth(3, 280);
    sheet.setColumnWidth(4, 280);
  }
  return sheet;
}

function saveTrainingRow(id, company, role, letterChanges, resumeChanges, proofNotes) {
  const sheet = getOrCreateTrainingSummarySheet();
  const label = [id, company, role].filter(Boolean).join(" — ");
  sheet.appendRow([label, letterChanges || "", resumeChanges || "", proofNotes || ""]);
}

function openTrainingSidebar(id, company, role, proofContent) {
  const template = HtmlService.createTemplateFromFile("training_sidebar");
  template.id           = id           || "";
  template.company      = company      || "";
  template.role         = role         || "";
  template.proofContent = proofContent || "";
  const html = template.evaluate().setTitle("Training Notes").setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ---- Gmail Email Monitor ----

function openClaudeSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("sidebar")
    .setTitle("Claude Chat")
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

function askClaudeFromSidebar(history) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty("CLAUDE_API_KEY");
  if (!apiKey) throw new Error("Claude API key not set. Use Job Assistant → Set Claude API Key.");

  const config = getConfigValues();
  const model = config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL;

  const payload = {
    model: model,
    max_tokens: 2048,
    system: "You are a helpful job search assistant for " + (config["User Full Name"] || config["User Display Name"] || "the user") + ". " + (config["User Background"] || "Help with their job search.") + " Help with: analyzing job descriptions, career positioning, drafting content, and answering questions about the job tracker. Be concise and direct.",
    messages: history
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: {
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01"
    },
    payload: JSON.stringify(payload)
  });

  const data = JSON.parse(response.getContentText());
  const textBlock = (data.content || []).find(function(c) { return c.type === "text"; });
  if (!textBlock) throw new Error("No text response from Claude.");
  return textBlock.text;
}

function installEmailMonitorTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkJobApplicationEmails") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkJobApplicationEmails").timeBased().everyHours(1).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("Email monitor installed — checks Gmail every hour. Add an Email Domain column to Jobs (e.g. 'devrev.ai'). For multiple senders (company domain + ATS domain), separate with a comma (e.g. 'devrev.ai, greenhouse.io').", "Job Assistant", 15);
}

function checkJobApplicationEmailsExtended() {
  checkJobApplicationEmails(30);
}

// One-time setup: daily 30-day-lookback safety-net sweep, in addition to (not replacing)
// the hourly checkJobApplicationEmails trigger. Catches rows whose Email Domain was fixed/added
// after the hourly trigger's forward-only search window had already passed them by.
function installExtendedEmailCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkJobApplicationEmailsExtended") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkJobApplicationEmailsExtended").timeBased().everyDays(1).atHour(3).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("Daily extended email check installed — runs a 30-day lookback sweep around 3am every day, alongside the existing hourly check.", "Job Assistant", 12);
}

// Lock-guarded: the hourly check and the daily extended check are separate triggers that can
// land close together (Apps Script doesn't align trigger firing times to the exact minute), and
// both read/write the same Auto Response/Status/Next Step cells and the same dedup Script
// Properties. Without this, an overlap could double-append the same Auto Response entry or let
// one run's dedup-list write clobber the other's. Skipping an overlapping run is safe — the
// run already in progress covers the same ground.
function checkJobApplicationEmails(daysBackOverride) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    console.warn("checkJobApplicationEmails already running elsewhere — skipping this run.");
    return;
  }
  try {
    checkJobApplicationEmailsImpl(daysBackOverride);
  } catch (err) {
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}

function checkJobApplicationEmailsImpl(daysBackOverride) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) return;

  const config = getConfigForRuntime();
  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];

  const emailDomainCol  = getColumnIndexByHeader(headers, "Email Domain");
  const autoResponseCol = getColumnIndexByHeader(headers, "Auto Response");
  const companyCol      = getColumnIndexByHeader(headers, "Company");
  const roleCol         = getColumnIndexByHeader(headers, "Role");
  const statusCol       = getColumnIndexByHeader(headers, "Status");
  const nextStepCol     = getColumnIndexByHeader(headers, "Next Step");
  const dueDateCol      = getColumnIndexByHeader(headers, "Next Step Due Date");
  const lastUpdatedCol  = getColumnIndexByHeader(headers, "Last Updated");
  const idCol           = getColumnIndexByHeader(headers, "ID");

  if (!emailDomainCol || !autoResponseCol) {
    console.warn("Email Domain or Auto Response column not found — skipping email monitor.");
    return;
  }

  const props = PropertiesService.getScriptProperties();
  const lastCheckRaw = props.getProperty("email_monitor_last_check");
  const lastCheckDate = lastCheckRaw ? new Date(lastCheckRaw) : new Date(Date.now() - 7 * 24 * 60 * 60 * 1000);
  const searchSinceDate = daysBackOverride
    ? new Date(Date.now() - daysBackOverride * 24 * 60 * 60 * 1000)
    : lastCheckDate;
  const afterDate = Utilities.formatDate(searchSinceDate, "UTC", "yyyy/MM/dd");

  const processedRaw = props.getProperty("email_monitor_processed") || "[]";
  const processedIds = new Set(JSON.parse(processedRaw));
  const newProcessedIds = new Set(processedIds);

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const data = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, sheet.getLastColumn()).getValues();
  let anyUpdates = false;

  data.forEach(function(row, i) {
    const sheetRow = i + JOBS_DATA_START_ROW;
    const emailDomainRaw = String(row[emailDomainCol - 1] || "").trim();
    if (!emailDomainRaw) return;

    const company = companyCol ? String(row[companyCol - 1] || "").trim() : "";
    const role    = roleCol    ? String(row[roleCol - 1]    || "").trim() : "";
    const jobId   = idCol      ? String(row[idCol - 1]      || "").trim() : "";

    // Supports comma-separated list (e.g. company domain + ATS domain) so interview
    // requests from an ATS platform or a recruiter's personal company email both match.
    const domains = emailDomainRaw.split(",")
      .map(function(d) { return d.trim(); })
      .filter(function(d) { return d; })
      .map(function(d) { return d.includes("@") ? d.split("@")[1] : d; });
    if (!domains.length) return;

    const fromQuery = domains.length > 1
      ? "(" + domains.map(function(d) { return "from:" + d; }).join(" OR ") + ")"
      : "from:" + domains[0];

    let threads;
    try {
      threads = GmailApp.search(fromQuery + " after:" + afterDate, 0, 20);
    } catch(e) {
      console.warn("Gmail search failed for " + domains.join(", ") + ": " + e.message);
      return;
    }

    threads.forEach(function(thread) {
      thread.getMessages().forEach(function(message) {
        const msgId = message.getId();
        // Dedup is scoped per-job, not global: shared ATS domains (ashbyhq.com,
        // greenhouse-mail.io, etc.) mean the same message can legitimately match
        // multiple rows' searches. A global msgId set let whichever row processed
        // it first permanently claim it, silently blocking the correct row forever.
        const dedupKey = jobId + "|" + msgId;
        if (newProcessedIds.has(dedupKey)) return;

        const subject  = message.getSubject();
        const body     = message.getPlainBody().substring(0, 3000);

        // Cheap mechanical pre-filter, no API call: if the company name doesn't appear
        // anywhere in the message, this is essentially certainly a different company's
        // email on a shared ATS domain. Gmail's date search is day-granularity, so the
        // same day's messages resurface on every hourly run — dedup here too, or this
        // cheap check (harmless but pointless) repeats all day for the same message.
        if (company && (subject + " " + body).toLowerCase().indexOf(company.toLowerCase()) === -1) {
          newProcessedIds.add(dedupKey);
          return;
        }

        const result = classifyJobEmail(subject, body, company, role, config);
        if (!result) return;

        // Mark processed as soon as Claude gives a definitive answer, whether or not it's
        // a match — otherwise company_match:false is never deduped, and Gmail's
        // day-granularity search re-surfaces the same message every hourly run for the
        // rest of the day, re-billing Claude for a question it already answered.
        newProcessedIds.add(dedupKey);

        if (result.company_match === false) return;

        const dateStr = Utilities.formatDate(message.getDate(), Session.getScriptTimeZone(), "M/d/yyyy");
        const entry   = dateStr + ": " + result.classification + (result.summary ? " — " + result.summary : "");
        // Read live, not from the snapshot taken at the top of the function — if more than
        // one message for this row gets processed within the same run, a snapshot read would
        // go stale after the first write and the second write would clobber it instead of
        // appending.
        const current = String(sheet.getRange(sheetRow, autoResponseCol).getValue() || "").trim();

        // Guards against duplicate lines if a message gets re-evaluated for a row that
        // already recorded it (e.g. the one-time dedup-key migration below).
        if (current.indexOf(entry) !== -1) return;

        anyUpdates = true;
        sheet.getRange(sheetRow, autoResponseCol).setValue(current ? current + "\n" + entry : entry);

        try { logActivityChange(jobId, company, role, "Auto Response", current || "(none)", entry); } catch(e) {}

        if (result.classification === "Rejected") {
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
          const oldStatus = statusCol ? String(row[statusCol - 1] || "") : "";
          if (statusCol)   sheet.getRange(sheetRow, statusCol).setValue("Rejected - " + today);
          if (nextStepCol) sheet.getRange(sheetRow, nextStepCol).setValue("Pass");
          if (dueDateCol)  sheet.getRange(sheetRow, dueDateCol).setValue("");
          try { logActivityChange(jobId, company, role, "Status", oldStatus, "Rejected - " + today); } catch(e) {}
        } else if (result.classification === "Interview Request") {
          const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
          const oldStatus = statusCol ? String(row[statusCol - 1] || "") : "";
          if (statusCol && oldStatus.toLowerCase().indexOf("interview -") !== 0) {
            sheet.getRange(sheetRow, statusCol).setValue("Interview - " + today);
            try { logActivityChange(jobId, company, role, "Status", oldStatus, "Interview - " + today); } catch(e) {}
          }

          const oldNextStep = nextStepCol ? String(row[nextStepCol - 1] || "") : "";
          if (nextStepCol) sheet.getRange(sheetRow, nextStepCol).setValue("Follow Up");
          const due = new Date(); due.setDate(due.getDate() + 1);
          if (dueDateCol) sheet.getRange(sheetRow, dueDateCol).setValue(
            Utilities.formatDate(due, Session.getScriptTimeZone(), "M/d/yyyy")
          );
          try { logActivityChange(jobId, company, role, "Next Step", oldNextStep, "Follow Up"); } catch(e) {}
        }

        if (lastUpdatedCol) sheet.getRange(sheetRow, lastUpdatedCol).setValue(formatDate(new Date()));
      });
    });
  });

  props.setProperty("email_monitor_last_check", new Date().toISOString());
  props.setProperty("email_monitor_processed", JSON.stringify(Array.from(newProcessedIds).slice(-1000)));

  if (anyUpdates) sortJobsByPrioritySafe();
}

function classifyJobEmail(subject, body, company, role, config) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return null;

  const toolSchema = {
    name: "email_classification",
    description: "Classification of a job application email",
    input_schema: {
      type: "object",
      required: ["classification", "summary"],
      properties: {
        classification: {
          type: "string",
          enum: ["Application Received", "Interview Request", "Assessment", "Offer", "Rejected", "No Response Needed", "Other"]
        },
        summary: {
          type: "string",
          description: "One sentence: what the email says and any action needed"
        },
        company_match: {
          type: "boolean",
          description: "True if the email body clearly references the expected company. False if this appears to be an email about a different company (e.g. a shared ATS domain used by multiple employers)."
        }
      }
    }
  };

  const userPrompt =
    "Classify this job application email.\n\n" +
    "Expected company: " + company + "\nRole: " + role + "\n" +
    "Subject: " + subject + "\n\nBody:\n" + body + "\n\n" +
    "Set company_match to true only if the email body clearly references '" + company + "'. " +
    "If the email appears to be about a different company (shared ATS platform like Greenhouse, Lever, Workday used by multiple employers), set company_match to false.\n\n" +
    "Classifications:\n" +
    "- Application Received: automated confirmation the application was received\n" +
    "- Interview Request: they want to schedule a call or interview\n" +
    "- Assessment: skills test, coding challenge, or assignment sent\n" +
    "- Offer: job offer extended\n" +
    "- Rejected: application declined\n" +
    "- No Response Needed: marketing, newsletter, or unrelated email\n" +
    "- Other: anything else";

  const payload = {
    model: config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL,
    max_tokens: 256,
    tools: [toolSchema],
    tool_choice: { type: "tool", name: "email_classification" },
    messages: [{ role: "user", content: userPrompt }]
  };

  const response = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
    method: "post",
    contentType: "application/json",
    headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
    const responseText = response.getContentText();
    console.error("Email classify error: " + responseText);
    if (isLowCreditsError(responseText)) notifyLowCreditsOnce("hourly email monitor", config);
    return null;
  }

  const data = JSON.parse(response.getContentText());
  if (!data.content) return null;
  for (const block of data.content) {
    if (block.type === "tool_use" && block.name === "email_classification") return block.input;
  }
  return null;
}

function isLowCreditsError(responseText) {
  return /credit balance is too low/i.test(String(responseText || ""));
}

// Time-based triggers have no UI context — SpreadsheetApp.getUi() throws — so a real popup
// isn't possible from the hourly email monitor. Email is the reliable equivalent; a note on
// cell A1 gives a persistent visual marker for whenever the sheet is next opened. Capped to
// once per day (via a Script Property) so a sustained outage doesn't spam every failed call.
function notifyLowCreditsOnce(context, config) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  if (props.getProperty("low_credit_alert_date") === today) return;
  props.setProperty("low_credit_alert_date", today);

  const recipient = (config && config["User Email"]) || Session.getEffectiveUser().getEmail();
  if (recipient) {
    try {
      MailApp.sendEmail(
        recipient,
        "Job Tracker: Claude API credit balance is too low",
        "The Claude API rejected a request (" + context + ") because the account's credit balance is too low. " +
        "Job analysis, tailoring, cover letters, and the hourly email monitor will not run until credits are added.\n\n" +
        "Add credits at https://console.anthropic.com (Plans & Billing)."
      );
    } catch (e) {
      console.warn("Could not send low-credit alert email: " + e.message);
    }
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
    if (sheet) {
      sheet.getRange(1, 1).setNote(
        "⚠ Claude API credit balance too low as of " + today + " (" + context + "). " +
        "Add credits at console.anthropic.com, then clear this note."
      );
    }
  } catch (e) {}
}

// ---- Claude model catalog / upgrade detection ----

const CLAUDE_MODEL_FAMILIES = ["haiku", "fable", "mythos", "sonnet", "opus"];

function getClaudeModelFamily(modelId) {
  const id = String(modelId || "").toLowerCase();
  for (let i = 0; i < CLAUDE_MODEL_FAMILIES.length; i++) {
    if (id.indexOf(CLAUDE_MODEL_FAMILIES[i]) !== -1) return CLAUDE_MODEL_FAMILIES[i];
  }
  return null;
}

// Version-number comparison fallback for when the Models API entry has no created_at —
// e.g. "claude-haiku-4-5-20251001" -> [4, 5, 20251001], compared component-wise.
function getModelVersionTuple(modelId) {
  const matches = String(modelId || "").match(/\d+/g);
  return matches ? matches.map(Number) : [];
}

function compareModelVersionTuples(a, b) {
  const len = Math.max(a.length, b.length);
  for (let i = 0; i < len; i++) {
    const diff = (a[i] || 0) - (b[i] || 0);
    if (diff !== 0) return diff;
  }
  return 0;
}

function fetchClaudeModelCatalog() {
  const apiKey = PropertiesService.getScriptProperties().getProperty("CLAUDE_API_KEY");
  if (!apiKey) return [];

  const models = [];
  let afterId = "";
  let guard = 0;

  while (guard < 10) {
    guard++;
    const url = "https://api.anthropic.com/v1/models?limit=100" +
      (afterId ? "&after_id=" + encodeURIComponent(afterId) : "");
    const response = UrlFetchApp.fetch(url, {
      method: "get",
      headers: { "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      console.warn("Models API error: " + response.getContentText());
      break;
    }

    const data = JSON.parse(response.getContentText());
    models.push.apply(models, data.data || []);

    if (!data.has_more || !data.last_id) break;
    afterId = data.last_id;
  }

  return models;
}

function setConfigValue(sheet, key, value) {
  const lastRow = sheet.getLastRow();
  const keys = lastRow > 0
    ? sheet.getRange(1, 1, lastRow, 1).getValues().map(function(r) { return String(r[0]).trim(); })
    : [];
  const rowIndex = keys.indexOf(key);

  if (rowIndex !== -1) {
    sheet.getRange(rowIndex + 1, 2).setValue(value);
  } else {
    sheet.appendRow([key, value]);
  }
}

// Weekly background check (see installNewModelCheckTrigger): diffs the live Models API
// catalog against the last-seen snapshot, and emails/notes only when a model that's brand
// new to the API shows up in the same family as the currently configured Analysis or
// Tailoring model. The first-ever run just establishes the baseline silently, since
// everything would otherwise look "new."
function checkForNewClaudeModelsAutomatic() {
  const config = getConfigForRuntime();
  const catalog = fetchClaudeModelCatalog();
  if (!catalog.length) return;

  const props = PropertiesService.getScriptProperties();
  const previouslySeenRaw = props.getProperty("claude_model_catalog_seen");
  const currentIds = catalog.map(function(m) { return m.id; });
  props.setProperty("claude_model_catalog_seen", JSON.stringify(currentIds));

  if (!previouslySeenRaw) return;

  const previouslySeen = new Set(JSON.parse(previouslySeenRaw));
  const newModels = catalog.filter(function(m) { return !previouslySeen.has(m.id); });
  if (!newModels.length) return;

  const watched = [
    { label: "Claude Analysis Model", current: config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL },
    { label: "Claude Tailoring Model", current: config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL }
  ];

  const relevant = [];
  watched.forEach(function(w) {
    const family = getClaudeModelFamily(w.current);
    if (!family) return;
    newModels.forEach(function(m) {
      if (getClaudeModelFamily(m.id) === family && m.id !== w.current) {
        relevant.push(w.label + ": " + (m.display_name || m.id) + " (" + m.id + ")");
      }
    });
  });

  if (!relevant.length) return;

  const recipient = config["User Email"] || Session.getEffectiveUser().getEmail();
  const message =
    "A new Claude model in the same family as your current settings just showed up:\n\n" +
    relevant.join("\n") +
    "\n\nUpdate the Config sheet's model row(s) if you want to switch, or run Job Assistant → Check for New Claude Models for an interactive prompt that can update it for you.";

  if (recipient) {
    try {
      MailApp.sendEmail(recipient, "Job Tracker: newer Claude model available", message);
    } catch (e) {
      console.warn("Could not send new-model alert email: " + e.message);
    }
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
    if (sheet) {
      const marker = "ℹ New Claude model available — see email.";
      const existingNote = sheet.getRange(1, 1).getNote();
      if (!existingNote || existingNote.indexOf(marker) === -1) {
        sheet.getRange(1, 1).setNote(existingNote ? existingNote + "\n" + marker : marker);
      }
    }
  } catch (e) {}
}

function installNewModelCheckTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === "checkForNewClaudeModelsAutomatic") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("checkForNewClaudeModelsAutomatic").timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(6).create();
  SpreadsheetApp.getActiveSpreadsheet().toast("Weekly new-model check installed — emails you Monday mornings if a model brand-new to Claude's catalog shows up in the same family as your configured Analysis/Tailoring model.", "Job Assistant", 12);
}

// Menu item: on-demand interactive check. Unlike the automatic weekly check (which can only
// flag models brand-new to the API), this compares against the FULL live catalog, so it can
// also surface an existing-but-newer same-family model you haven't switched to yet — and offers
// a real Yes/No prompt (only possible when a user is actively in the sheet, not from a trigger).
function checkForNewClaudeModelsInteractive() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) { ui.alert("Missing tab: Config"); return; }

  const config = getConfigForRuntime();
  ss.toast("Checking Claude's model catalog...", "Job Assistant", 30);
  const catalog = fetchClaudeModelCatalog();
  ss.toast("", "", 1);

  if (!catalog.length) {
    ui.alert("Couldn't reach the Claude Models API — check CLAUDE_API_KEY in Script Properties and try again.");
    return;
  }

  const watched = [
    { key: "Claude Analysis Model", current: config["Claude Analysis Model"] || CLAUDE_DEFAULT_MODEL },
    { key: "Claude Tailoring Model", current: config["Claude Tailoring Model"] || CLAUDE_DEFAULT_TAILORING_MODEL }
  ];

  let anyFamilyMatched = false;
  let anyOffered = false;

  watched.forEach(function(w) {
    const family = getClaudeModelFamily(w.current);
    if (!family) return;
    anyFamilyMatched = true;

    const candidates = catalog.filter(function(m) {
      return getClaudeModelFamily(m.id) === family && m.id !== w.current;
    });
    if (!candidates.length) return;

    candidates.sort(function(a, b) {
      if (a.created_at && b.created_at) return new Date(b.created_at) - new Date(a.created_at);
      return compareModelVersionTuples(getModelVersionTuple(b.id), getModelVersionTuple(a.id));
    });

    const newest = candidates[0];
    const currentEntry = catalog.filter(function(m) { return m.id === w.current; })[0];

    const isNewer = (currentEntry && currentEntry.created_at && newest.created_at)
      ? new Date(newest.created_at) > new Date(currentEntry.created_at)
      : (currentEntry
          ? compareModelVersionTuples(getModelVersionTuple(newest.id), getModelVersionTuple(w.current)) > 0
          : true);

    if (!isNewer) return;

    anyOffered = true;
    const response = ui.alert(
      "Newer Claude model available",
      w.key + " is currently \"" + w.current + "\".\n\n" +
      "Found: " + (newest.display_name || newest.id) + " (" + newest.id + ")" +
      (newest.created_at ? "\nReleased: " + newest.created_at : "") +
      "\n\nUpdate Config to use this model?",
      ui.ButtonSet.YES_NO
    );

    if (response === ui.Button.YES) {
      setConfigValue(configSheet, w.key, newest.id);
      ss.toast(w.key + " updated to " + newest.id + ".", "Job Assistant", 8);
    }
  });

  if (!anyFamilyMatched) {
    ui.alert("Couldn't determine a model family for your current Config values — nothing to check.");
  } else if (!anyOffered) {
    ui.alert("You're already on the newest model in each configured family.");
  }
}

// ---- One-time backfills ----

function backfillBlankAtsScores() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(JOBS_SHEET_NAME);
  if (!sheet) { SpreadsheetApp.getUi().alert("Jobs sheet not found."); return; }

  const config = getConfigForRuntime();
  const baseResumeUrl = config["Base Resume URL"] || "";
  const baseResumeText = getResumeText(baseResumeUrl);

  const headers = sheet.getRange(JOBS_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  const atsCol     = getColumnIndexByHeader(headers, "ATS Score");
  const jobDescCol = getColumnIndexByHeader(headers, "Job Description");
  const companyCol = getColumnIndexByHeader(headers, "Company");
  const roleCol    = getColumnIndexByHeader(headers, "Role");

  if (!atsCol || !jobDescCol) {
    SpreadsheetApp.getUi().alert("ATS Score or Job Description column not found.");
    return;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) { SpreadsheetApp.getUi().alert("No data rows found."); return; }

  const data = sheet.getRange(JOBS_DATA_START_ROW, 1, lastRow - JOBS_HEADER_ROW, sheet.getLastColumn()).getValues();
  let filled = 0;
  let alreadyScored = 0;
  let noDescription = 0;

  for (let i = 0; i < data.length; i++) {
    const sheetRow = i + JOBS_DATA_START_ROW;
    const atsRaw  = String(data[i][atsCol - 1] || "").trim();
    const jobDesc = String(data[i][jobDescCol - 1] || "").trim();
    const company = companyCol ? String(data[i][companyCol - 1] || "").trim() : "";
    const role    = roleCol    ? String(data[i][roleCol - 1]    || "").trim() : "";

    if (atsRaw !== "") { alreadyScored++; continue; }
    if (!jobDesc)      { noDescription++; continue; }

    const atsResult = callClaudeForAtsScore(jobDesc, company, role, baseResumeText, config);
    if (atsResult !== null) {
      sheet.getRange(sheetRow, atsCol).setValue(atsResult.score);
      filled++;
    }

    Utilities.sleep(500);
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "Backfill complete. Filled: " + filled + ". Already had score: " + alreadyScored + ". No job description (can't score): " + noDescription,
    "Job Assistant", 15
  );
}

// One-time menu action: extends the green "Apply + New" conditional formatting
// rule so it also matches rows resolved via "Discuss - Apply" / "Apply if Easy - Apply".
// The yellow "Apply if Easy + New", blue "Review", and red "Skip" rules already work
// unchanged because applyBorderlineResolutionOverride() writes literal Status values
// ("New"/"Skip") that those rules already match.
function normalizeFormulaForComparison(formula) {
  return String(formula || "").replace(/\s+/g, "").toLowerCase();
}

function updateConditionalFormattingForBorderlineResolution() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  const rules = sheet.getConditionalFormatRules();

  const oldFormula = '=and($M4="Apply",$N4="New")';
  const newFormula = '=and(or($M4="Apply",regexmatch($M4,"- Apply$")),$N4="New")';
  const normalizedOldFormula = normalizeFormulaForComparison(oldFormula);

  let updated = false;
  const foundFormulas = [];

  const newRules = rules.map(function (rule) {
    const condition = rule.getBooleanCondition();

    if (condition && condition.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA) {
      const formula = String(condition.getCriteriaValues()[0] || "");
      foundFormulas.push(formula);

      if (normalizeFormulaForComparison(formula) === normalizedOldFormula) {
        updated = true;
        return rule.copy().whenFormulaSatisfied(newFormula).build();
      }
    }

    return rule;
  });

  sheet.setConditionalFormatRules(newRules);

  SpreadsheetApp.getUi().alert(updated
    ? "Conditional formatting updated: green (Apply + New) rule now also matches rows resolved to '- Apply'."
    : "No matching rule found (expected formula: " + oldFormula + "). Custom formulas found instead:\n" + (foundFormulas.join("\n") || "(none)") + "\nNo changes made.");
}

// Google Sheets splits a conditional format rule's range into fragments when a row is
// inserted/deleted in the middle of it (e.g. a manual row insert), leaving a gap that
// silently stops getting that rule's formatting. Consolidates every rule's ranges back into
// a single contiguous block per column-span, and extends it to the sheet's actual max row
// (not a hardcoded buffer) so this can't quietly run out of room again as rows are added.
function repairConditionalFormattingRanges() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(JOBS_SHEET_NAME);
  const rules = sheet.getConditionalFormatRules();
  const maxRow = sheet.getMaxRows();
  const numRows = maxRow - JOBS_DATA_START_ROW + 1;

  const report = [];
  let repaired = 0;

  const newRules = rules.map(function(rule) {
    const ranges = rule.getRanges();
    const fragmentNotations = ranges.map(function(r) { return r.getA1Notation(); }).join(", ");

    // A single rule can legitimately cover more than one column strip (e.g. this sheet
    // highlights Recommendation (M) and Status (N) together) — only merge fragments that
    // share the SAME column span; different column groups stay as separate ranges.
    const groups = {};
    const groupOrder = [];
    ranges.forEach(function(r) {
      const key = r.getColumn() + ":" + r.getNumColumns();
      if (!groups[key]) { groups[key] = []; groupOrder.push(key); }
      groups[key].push(r);
    });

    const consolidatedRanges = groupOrder.map(function(key) {
      const group = groups[key];
      return sheet.getRange(JOBS_DATA_START_ROW, group[0].getColumn(), numRows, group[0].getNumColumns());
    });

    const alreadyCorrect = ranges.length === consolidatedRanges.length &&
      ranges.every(function(r, i) {
        return r.getRow() === consolidatedRanges[i].getRow() &&
          r.getLastRow() === consolidatedRanges[i].getLastRow() &&
          r.getColumn() === consolidatedRanges[i].getColumn() &&
          r.getNumColumns() === consolidatedRanges[i].getNumColumns();
      });

    if (alreadyCorrect) return rule;

    const newNotations = consolidatedRanges.map(function(r) { return r.getA1Notation(); }).join(", ");
    report.push(fragmentNotations + " -> " + newNotations);
    repaired++;
    return rule.copy().setRanges(consolidatedRanges).build();
  });

  sheet.setConditionalFormatRules(newRules);

  SpreadsheetApp.getUi().alert(
    "Conditional formatting range repair:\n\n" +
    (repaired > 0 ? repaired + " rule(s) updated.\n\n" : "No changes needed.\n\n") +
    report.join("\n")
  );
}
