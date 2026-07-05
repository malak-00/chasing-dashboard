// ============================================================
// Chasing — PRODUCTIVITY DASHBOARD BACKEND
// Google Apps Script Web App
// Deploy as: Execute as ME, Anyone can access
// ============================================================

// Historical hardcoded roster -- kept only as a one-time seed for the
// "Chasers" tab the first time it's created (see getOrCreateChasersTab
// below), so existing chasers keep working the moment this deploys. From
// then on the "Chasers" tab is the source of truth: add a new chaser or
// repoint an existing one's tracker Sheet ID from the dashboard's
// Settings tab, no code change or redeploy needed. Do not add new
// chasers here -- add them from the dashboard instead.
const CHASER_SHEETS = {
  Alex:  "1hAEVtrDXllL91lRss6O5nnaOrklaeHhaI73EkEzYnmc",
  Hope:  "1s6wgiSQkWq6D5cx_fk8oDEvG__xiQ6eQ-UcoVout9K0",
  Rose:  "1ZPgtnYh6g8ObJIgzUZrx4atOvsq2I5UlvW-mBQQU6KQ",
  Frank: "1XHX1FJ_1S6IxHjDec3OeV2wiyViRd87XfeOChLrTtl0",
  Nova:  "1O2mQvVYpy6Se2kKmubSa9WA9_scQpItzo6m0E6gmimU"
};

// ============================================================
// CHASER ROSTER -- persisted, editable from the dashboard
// ============================================================
const CHASERS_TAB_HEADERS = ["Name", "SheetId", "Active"];

function getOrCreateChasersTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Chasers");
  if (!sheet) {
    sheet = ss.insertSheet("Chasers");
    sheet.appendRow(CHASERS_TAB_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CHASERS_TAB_HEADERS.length).setFontWeight("bold");
    // Seed from the legacy hardcoded roster so existing chasers are
    // unaffected the first time this tab is created.
    Object.entries(CHASER_SHEETS).forEach(([name, sheetId]) => {
      sheet.appendRow([name, sheetId, true]);
    });
    Logger.log("Created Chasers tab, seeded " + Object.keys(CHASER_SHEETS).length + " chasers from CHASER_SHEETS");
  }
  return sheet;
}

// Full roster (including inactive) for the dashboard's Settings UI.
// [{ name, sheetId, active }, ...]
function getChasersConfig() {
  const cached = CacheService.getScriptCache().get("chasers_config");
  if (cached) return JSON.parse(cached);

  const sheet = getOrCreateChasersTab();
  const data  = sheet.getDataRange().getValues();
  const roster = [];
  for (let r = 1; r < data.length; r++) {
    const name = String(data[r][0] || "").trim();
    if (!name) continue;
    roster.push({
      name,
      sheetId: String(data[r][1] || "").trim(),
      active:  data[r][2] === true || String(data[r][2]).trim().toUpperCase() === "TRUE",
    });
  }
  CacheService.getScriptCache().put("chasers_config", JSON.stringify(roster), 300);
  return roster;
}

// { name: sheetId } for active chasers only -- every day/week read uses
// this in place of the old hardcoded CHASER_SHEETS.
function getActiveChaserSheetMap() {
  const map = {};
  getChasersConfig().forEach(c => { if (c.active) map[c.name] = c.sheetId; });
  return map;
}

// Add a new chaser or update an existing one's Sheet ID / Active flag.
// Matched by exact Name (renaming an existing chaser isn't supported here --
// historical archive rows are keyed by name, so add a new one instead).
// Validates the Sheet ID actually opens before saving so a typo fails loudly
// here instead of silently producing zero data on every future sync.
function saveChaser(params) {
  const name    = String(params.name || "").trim();
  const sheetId = String(params.sheetId || "").trim();
  const active  = String(params.active) !== "false"; // default true unless explicitly "false"

  if (!name)    return { success: false, error: "Chaser name is required." };
  if (!sheetId) return { success: false, error: "Tracker Sheet ID is required." };

  try {
    SpreadsheetApp.openById(sheetId);
  } catch (e) {
    return { success: false, error: "Could not open that Sheet ID -- check it's correct and shared with this script: " + e.message };
  }

  const sheet = getOrCreateChasersTab();
  const data  = sheet.getDataRange().getValues();
  let rowNum = -1;
  for (let r = 1; r < data.length; r++) {
    if (String(data[r][0]).trim() === name) { rowNum = r + 1; break; }
  }

  if (rowNum > 0) {
    sheet.getRange(rowNum, 1, 1, 3).setValues([[name, sheetId, active]]);
  } else {
    sheet.appendRow([name, sheetId, active]);
  }

  CacheService.getScriptCache().remove("chasers_config");
  return { success: true, name, sheetId, active };
}

// ============================================================
// RESPONSE SOURCES
// Each tab now also carries a campaignKey so byCampaign is
// populated automatically — no code changes needed to add more.
// ============================================================
const RESPONSE_SOURCES = [
  {
    id: "1TGltg5YzNfYeDl7qvVvwbOIMmoGLCfcfTvMV0uKtbVc",
    tabs: [
      { name: "ORT RESPONSES", feedbackCol: "FAX FEEDBACK",    chaserCol: "CHASER NAME", campaignKey: "ort" },
      { name: "CGM Responses", feedbackCol: "FAX SENT ON EST", chaserCol: "CHASER NAME", campaignKey: "cgm" }
    ]
  },
  {
    id: "1GLTMtACC6eeytfVdlUXAFpqPlBWfMoeac354UnuVlU8",
    tabs: [
      { name: "Responses",   feedbackCol: "FAX FEEDBACK", chaserCol: "CHASER NAME", campaignKey: "lymphc" }
    ]
  },
  {
    id: "1R-MO93QX48mHCjoBCcLtVosAvOZ1kCTABV37cECSVqc",
    tabs: [
      { name: "Responses LY", feedbackCol: "Fax Feedback", chaserCol: "Chaser Name", campaignKey: "lymphw" }
    ]
  }
];

// Single source of truth for the campaign roster -- add a 5th campaign
// here and it's picked up everywhere a per-campaign totals object is
// built (see zeroCampaignTotals below), instead of updating a dozen
// hand-copied literals scattered through this file.
const CAMPAIGN_LABELS = { ort: "ORT", cgm: "CGM", lymphc: "LymphC", lymphw: "LymphW" };

// Returns a FRESH { ort:{approved,denied[,total]}, cgm:..., ... } object --
// fresh on every call since callers mutate these in place (c.total++ etc.),
// so a single shared instance would leak counts across dates/chasers.
function zeroCampaignTotals(withTotal) {
  const obj = {};
  Object.keys(CAMPAIGN_LABELS).forEach(key => {
    obj[key] = withTotal ? { approved: 0, denied: 0, total: 0 } : { approved: 0, denied: 0 };
  });
  return obj;
}

// ============================================================
// APPROVAL / DENIAL MATCHERS
// ============================================================
const APPROVAL_PATTERNS = ["APPROVED", "APPROVED+CN"];
const DENIAL_PATTERNS   = ["DENIAL", "DENIED", "RECEIVED DENIAL", "REQUEST DENIED"];

function isApproval(text) {
  const upper = text.toUpperCase();
  return APPROVAL_PATTERNS.some(p => upper.includes(p));
}

function isDenial(text) {
  const upper = text.toUpperCase();
  return DENIAL_PATTERNS.some(p => upper.includes(p));
}

// ============================================================
// ARCHIVE
// ============================================================
const ARCHIVE_SHEET_ID = "1eRjelUuuEt0JJCPWw6ohhYe9utWlWdQHlontpOZQnEw";

// ============================================================
// ARCHIVE HELPERS
// ============================================================

// Standard column headers used in every month tab
const ARCHIVE_HEADERS = [
  "Date","Chaser","Cases","Positive",
  "Approvals","Denials","TimeMins","Efficiency","Productivity",
  "TotalShift","TotalCalls","TotalDurationMins","ACWDuration","ProductiveTime",
  "ORT_Approved","ORT_Denied",
  "CGM_Approved","CGM_Denied",
  "LymphC_Approved","LymphC_Denied",
  "LymphW_Approved","LymphW_Denied"
];

// Returns the tab name for a given date: "Jun 2026"
function monthTabName(dateStr) {
  // dateStr can be "6/25", "6/25/2026", or "2026-06-25"
  let d;
  if (dateStr instanceof Date) {
    d = dateStr;
  } else if (String(dateStr).includes("-")) {
    d = new Date(dateStr);
  } else {
    const parts = String(dateStr).split("/");
    const yr    = parts[2] ? parseInt(parts[2]) : new Date().getFullYear();
    d = new Date(yr, parseInt(parts[0])-1, parseInt(parts[1]));
  }
  const months = ["Jan","Feb","Mar","Apr","May","Jun",
                  "Jul","Aug","Sep","Oct","Nov","Dec"];
  return months[d.getMonth()] + " " + d.getFullYear();
}

// Converts a dashboard dateTab ("M/D" or "M/D/YYYY") into a full "M/D/YYYY"
// string for comparison against Lead History log dates, which always carry
// an explicit year. Bare "M/D" is assumed to be the current year, same
// year-inference rule as monthTabName() above.
function dateTabToFullDate(dateTab) {
  const parts = String(dateTab).split("/");
  const yr    = parts[2] ? parseInt(parts[2], 10) : new Date().getFullYear();
  return parseInt(parts[0], 10) + "/" + parseInt(parts[1], 10) + "/" + yr;
}

// Normalize an archive Date-column cell to "M/D" so it matches a dateTab
// like "6/25" whether or not Google Sheets has auto-converted the cell from
// plain text into a real Date value (Sheets does this silently for any
// cell that looks like a date, and nothing here sets the column to Plain
// Text). Non-date cells (e.g. "WEEK") pass through unchanged so they never
// accidentally match a dateTab.
function normalizeDateCellToTab(cellValue) {
  if (cellValue instanceof Date && !isNaN(cellValue)) {
    return (cellValue.getMonth() + 1) + "/" + cellValue.getDate();
  }
  const s     = String(cellValue).trim();
  const parts = s.split("/");
  if (parts.length >= 2 && parseInt(parts[0], 10) > 0 && parseInt(parts[1], 10) > 0) {
    return parseInt(parts[0], 10) + "/" + parseInt(parts[1], 10);
  }
  return s;
}

// Get or create a month tab in the archive spreadsheet
function getOrCreateMonthTab(ss, tabName) {
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(ARCHIVE_HEADERS);
    // Freeze header row
    sheet.setFrozenRows(1);
    // Bold the header
    sheet.getRange(1, 1, 1, ARCHIVE_HEADERS.length).setFontWeight("bold");
    Logger.log("Created new tab: " + tabName);
  }
  return sheet;
}

// Build duplicate guard from a specific month tab
function buildExistingSet(sheet) {
  const existing = new Set();
  const data     = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const key = normalizeDateCellToTab(data[r][0]) + "|" + String(data[r][1]);
    existing.add(key);
  }
  return existing;
}

// ============================================================
// ARCHIVE A SINGLE DAY  (called by EOD trigger + manual runs)
// ============================================================
function archiveDayData(dateTab) {
  dateTab = dateTab || getTodayTab();
  const data = getDayData(dateTab);

  const ss       = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const tabName  = monthTabName(dateTab);
  const sheet    = getOrCreateMonthTab(ss, tabName);
  const existing = buildExistingSet(sheet);

  // Use the dedicated per-chaser campaign counter (correct daily counts,
  // names normalized, no week-level inflation).
  const chaserCamps = readChaserCampaignCountsForDate(dateTab);

  // byChaser from getDayData gives total approvals/denials across all campaigns
  const bc = data.responses.byChaser   || {};

  // Check if this date already has rows (skip week header if so)
  const alreadyHasRows = [...existing].some(k => k.startsWith(dateTab + "|"));
  const isMonday       = isFirstDayOfWeek(dateTab);

  // Write week header row if this is a Monday and no rows exist for this date yet
  if (isMonday && !alreadyHasRows) {
    const weekLabel = getWeekLabel(dateTab);
    // Write a styled week header — "WEEK" in Date col makes parser skip it
    sheet.appendRow(["WEEK", weekLabel].concat(new Array(ARCHIVE_HEADERS.length - 2).fill("")));

    // Style the week header row
    const lastRow = sheet.getLastRow();
    const headerRange = sheet.getRange(lastRow, 1, 1, ARCHIVE_HEADERS.length);
    headerRange.setBackground("#1A2C42");
    headerRange.setFontColor("#00C2A8");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(10);
  }

  // If this is the first chaser row for a new date (not Monday but new day),
  // write a subtle date separator row
  if (!alreadyHasRows && !isMonday) {
    // No separator needed — days flow naturally, week headers mark the groups
  }

  let written = 0;
  const missingTabChasers = [];
  for (const c of data.chasers) {
    const key = dateTab + "|" + c.name;
    if (existing.has(key)) {
      Logger.log("Skipping duplicate: " + key);
      continue;
    }
    // If this chaser's tracker has no tab named dateTab + "." for this date, do NOT
    // write a row — a chaser tracker can have a leftover tab from last year with the
    // same "M/D" name but no trailing dot; readChaserTab() only ever looks up the
    // dotted (current-year) tab and never falls back to it, so tabFound=false here
    // genuinely means "no data yet", not "check the other tab". Writing zeros would
    // be indistinguishable from a real zero-case day and would corrupt the archive
    // permanently. Skip and leave it for a later Sync once the tab exists.
    if (!c.tabFound) {
      missingTabChasers.push(c.name);
      continue;
    }
    // Total approvals/denials across all campaigns for this chaser
    let r = bc[c.name];
    if (!r) {
      const matchKey = Object.keys(bc).find(k => normalizeChaserName(k) === c.name);
      if (matchKey) r = bc[matchKey];
    }
    r = r || { approvals: 0, denials: 0 };

    const eff = c.totalCases ? (c.totalPositive / c.totalCases * 100).toFixed(1) : "";

    // Per-chaser per-campaign breakdown — daily counts from readChaserCampaignCountsForDate
    // Each number = how many leads this chaser was listed on for that campaign today
    const cc = chaserCamps[c.name] || zeroCampaignTotals();

    sheet.appendRow([
      dateTab, c.name, c.totalCases, c.totalPositive,
      r.approvals, r.denials, c.totalTimeMins, eff,  // no fax column
      "",   // Productivity
      "",   // TotalShift
      "",   // TotalCalls
      "",   // TotalDurationMins
      "",   // ACWDuration
      "",   // ProductiveTime
      cc.ort.approved,    cc.ort.denied,
      cc.cgm.approved,    cc.cgm.denied,
      cc.lymphc.approved, cc.lymphc.denied,
      cc.lymphw.approved, cc.lymphw.denied,
    ]);
    existing.add(key);
    written++;
  }

  if (missingTabChasers.length) {
    Logger.log("*** ARCHIVE INCOMPLETE for " + dateTab + " *** no tab named \"" + dateTab +
      ".\" found for: " + missingTabChasers.join(", ") +
      " — their tracker tab for today may not exist yet. No row was written for them " +
      "(zeros were NOT recorded). Re-run Sync for " + dateTab + " once the tab exists.");
  }

  // Add bottom border to the last chaser row for this date
  const lastDataRow = sheet.getLastRow();
  if (lastDataRow > 1 && written > 0) {
    const borderRange = sheet.getRange(lastDataRow, 1, 1, ARCHIVE_HEADERS.length);
    borderRange.setBorder(
      null, null, true, null, null, null,  // bottom border only
      "#00C2A8",                            // teal color to match dashboard theme
      SpreadsheetApp.BorderStyle.SOLID_MEDIUM
    );
  }

  // After writing all chasers for a Friday, also add a blank separator row
  if (isFriday(dateTab)) {
    sheet.appendRow(new Array(ARCHIVE_HEADERS.length).fill(""));
  }

  Logger.log("Archived " + dateTab + " → " + tabName + " | Written: " + written + " rows");
}

// ── Week label helpers ─────────────────────────────────────────

// Returns true if dateTab is a Monday
function isFirstDayOfWeek(dateTab) {
  const d = parseDateTab(dateTab);
  return d.getDay() === 1;
}

// Returns true if dateTab is a Friday
function isFriday(dateTab) {
  const d = parseDateTab(dateTab);
  return d.getDay() === 5;
}

// Returns "Week of 6/23 - 6/27/2026"
function getWeekLabel(dateTab) {
  const mon = parseDateTab(dateTab);
  const fri = new Date(mon);
  fri.setDate(mon.getDate() + 4);
  const fmt = d => (d.getMonth()+1) + "/" + d.getDate();
  return "Week of " + fmt(mon) + " - " + fmt(fri) + "/" + fri.getFullYear();
}

// Parse "6/24", "6/24/2026" into a Date object
function parseDateTab(dateTab) {
  const parts = String(dateTab).split("/");
  const yr    = parts[2] ? parseInt(parts[2]) : new Date().getFullYear();
  return new Date(yr, parseInt(parts[0])-1, parseInt(parts[1]));
}

// ============================================================
// EOD TRIGGER
// ============================================================
function setupEodTrigger() {
  // Delete any existing eodArchive triggers first
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "eodArchive")
    .forEach(t => ScriptApp.deleteTrigger(t));

  // Runs every day at 5 PM — archives YESTERDAY's data
  // So on Monday 5 PM it archives Friday, on Tuesday 5 PM it archives Monday, etc.
  ScriptApp.newTrigger("eodArchive")
    .timeBased()
    .everyDays(1)
    .atHour(17)   // 5 PM — change this if needed (uses script timezone)
    .create();

  Logger.log("EOD trigger set: runs daily at 5 PM, archives previous weekday");
}

function eodArchive() {
  const today    = new Date();
  const todayDay = today.getDay(); // 0=Sun, 1=Mon, ..., 6=Sat

  // Don't archive on Sundays (would archive Saturday, no data)
  if (todayDay === 0) { Logger.log("Sunday — skipping."); return; }

  // Figure out which date to archive (yesterday, skipping weekends)
  const archiveDate = new Date(today);

  if (todayDay === 1) {
    // Monday — archive Friday (3 days back)
    archiveDate.setDate(today.getDate() - 3);
  } else {
    // Tuesday–Saturday — archive yesterday
    archiveDate.setDate(today.getDate() - 1);
  }

  const archiveTab = (archiveDate.getMonth()+1) + "/" + archiveDate.getDate();
  Logger.log("Running EOD archive for: " + archiveTab + " (triggered on " + getTodayTab() + ")");

  archiveDayData(archiveTab);
  archiveCampaignResponses(archiveTab);

  // Lead History sync reads LIVE state from the source tabs right now —
  // it is not tied to archiveTab's date, it always reflects "as of today".
  // Wrapped so any Lead History issue never blocks the existing daily archive.
  try {
    runDailyLeadHistorySync();
  } catch(e) {
    Logger.log("Lead History sync failed (archive still completed normally): " + e.message);
  }

  Logger.log("EOD archive complete for " + archiveTab);
}

// Manual override — archive a specific date right now
function archiveSpecificDate() {
  const DATE = "6/29";  // ← change this to whatever date you want
  Logger.log("Manually archiving: " + DATE);
  archiveDayData(DATE);
  Logger.log("Done");
}

// Check what the trigger is currently set to
function checkTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    Logger.log("Function: " + t.getHandlerFunction() +
               " | Type: " + t.getEventType() +
               " | Source: " + t.getTriggerSource());
  });
}

// ============================================================
// READ ARCHIVE — all month tabs combined, returned as flat rows
// Skips: blank rows, week header rows (Date === "WEEK")
// ============================================================
// ============================================================
// CHASER NAME NORMALIZATION
// Maps full names from historical sheets to display names
// Add entries here if new name variants appear
// ============================================================
const CHASER_NAME_MAP = {
  "alex woods":       "Alex",
  "hope smith":       "Hope",
  "rose simon":       "Rose",
  "frank clarkson":   "Frank",
  "nova grace":       "Nova",
  // Former chasers — kept as-is for historical data
  "nora atkins":      "Nora",
  "jamie williams":   "Jamie",
  "rick nelson":      "Rick",
  "caroline richards":"Caroline",
};

function normalizeChaserName(fullName) {
  const key = fullName.trim().toLowerCase();
  return CHASER_NAME_MAP[key] || fullName.trim(); // fallback: return as-is
}

function getArchiveData() {
  const ss          = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const sheets      = ss.getSheets();
  const allRows     = [];
  const MONTH_NAMES = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];

  for (const sheet of sheets) {
    const rawName = sheet.getName();
    const tabName = rawName.trim().toLowerCase().replace(/\s+/g," ");
    const isMonth = MONTH_NAMES.some(m => tabName.startsWith(m + " ") && /\d{4}$/.test(tabName));
    if (!isMonth) continue;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers = data[0].map(h => String(h).trim());

    for (let r = 1; r < data.length; r++) {
      const rawDate   = data[r][0];
      const chaserVal = String(data[r][1]).trim();

      if (!rawDate && !chaserVal) continue;
      if (!chaserVal) continue;

      // Normalize date — Google Sheets returns Date objects for date-formatted cells
      let dateStr;
      if (rawDate instanceof Date && !isNaN(rawDate)) {
        dateStr = (rawDate.getMonth()+1) + "/" + rawDate.getDate() + "/" + rawDate.getFullYear();
      } else {
        dateStr = String(rawDate).trim();
      }

      if (!dateStr || dateStr.toUpperCase() === "WEEK") continue;

      // Skip summary rows like "Total Week", "Total", etc.
      const chaserUpper = String(data[r][1]).trim().toUpperCase();
      if (chaserUpper.startsWith("TOTAL") || chaserUpper === "") continue;

      const obj = {};
      headers.forEach((h, i) => { obj[h] = data[r][i]; });
      obj["Date"]   = dateStr;
      obj["Chaser"] = normalizeChaserName(String(data[r][1]).trim());
      allRows.push(obj);
    }
  }

  Logger.log("Total archive rows returned: " + allRows.length);
  return { rows: allRows };
}



// ============================================================
// SETTINGS & UTLATEL PERSISTENCE
// Stored in the archive spreadsheet:
//   "Settings" tab  — shift history per chaser
//   "Utlatel" tab   — call duration data per agent per date
// ============================================================

// ── SETTINGS TAB HELPERS ──────────────────────────────────────

function getOrCreateSettingsTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Settings");
  if (!sheet) {
    sheet = ss.insertSheet("Settings");
    sheet.appendRow(["Chaser","ShiftType","EffectiveDate","Notes"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1A2C42").setFontColor("#00C2A8");
    Logger.log("Created Settings tab");
  }
  return sheet;
}

function getOrCreateUtlatelTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Utlatel");
  if (!sheet) {
    sheet = ss.insertSheet("Utlatel");
    sheet.appendRow(["Date","Agent","DurationMins","Calls"]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1A2C42").setFontColor("#00C2A8");
    Logger.log("Created Utlatel tab");
  }
  return sheet;
}

// ── READ SETTINGS ─────────────────────────────────────────────

function getSettings() {
  const sheet = getOrCreateSettingsTab();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { shiftHistory: [], utlatel: [] };

  const headers = data[0].map(h => String(h).trim());
  const shiftHistory = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    // Normalize date
    if (obj.EffectiveDate instanceof Date && !isNaN(obj.EffectiveDate)) {
      const d = obj.EffectiveDate;
      obj.EffectiveDate = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getFullYear();
    } else {
      obj.EffectiveDate = String(obj.EffectiveDate || "").trim();
    }
    return obj;
  }).filter(r => r.Chaser && r.ShiftType);

  return { shiftHistory };
}

function getUtlatelData() {
  const sheet = getOrCreateUtlatelTab();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  const headers = data[0].map(h => String(h).trim());
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h,i) => { obj[h] = row[i]; });
    // Normalize date
    if (obj.Date instanceof Date && !isNaN(obj.Date)) {
      const d = obj.Date;
      obj.Date = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getFullYear();
    } else {
      obj.Date = String(obj.Date || "").trim();
    }
    return obj;
  }).filter(r => r.Date && r.Agent);
}

// ── SAVE SHIFT ────────────────────────────────────────────────

function saveShiftSetting(chaser, shiftType, effectiveDate, notes) {
  const sheet = getOrCreateSettingsTab();
  sheet.appendRow([chaser, shiftType, effectiveDate, notes || ""]);
  Logger.log("Saved shift: " + chaser + " = " + shiftType + " from " + effectiveDate);
  // Clear archive cache so dashboard reloads fresh settings
  CacheService.getScriptCache().remove("settings_data");
}

// ── SAVE UTLATEL ──────────────────────────────────────────────

function saveUtlatelData(rows) {
  // rows: [{ date, agent, durationMins, calls }]
  const sheet    = getOrCreateUtlatelTab();
  const existing = sheet.getDataRange().getValues();

  // Build set of existing date|agent combos
  const existingKeys = new Set();
  for (let r = 1; r < existing.length; r++) {
    const d = existing[r][0] instanceof Date
      ? (existing[r][0].getMonth()+1)+"/"+existing[r][0].getDate()+"/"+existing[r][0].getFullYear()
      : String(existing[r][0]).trim();
    existingKeys.add(d + "|" + String(existing[r][1]).trim());
  }

  let written = 0;
  let updated = 0;

  for (const row of rows) {
    const key = row.date + "|" + row.agent;
    if (existingKeys.has(key)) {
      // Update existing row
      for (let r = 1; r < existing.length; r++) {
        const d = existing[r][0] instanceof Date
          ? (existing[r][0].getMonth()+1)+"/"+existing[r][0].getDate()+"/"+existing[r][0].getFullYear()
          : String(existing[r][0]).trim();
        const existKey = d + "|" + String(existing[r][1]).trim();
        if (existKey === key) {
          sheet.getRange(r+1, 3).setValue(row.durationMins);
          sheet.getRange(r+1, 4).setValue(row.calls);
          updated++;
          break;
        }
      }
    } else {
      sheet.appendRow([row.date, row.agent, row.durationMins, row.calls]);
      existingKeys.add(key);
      written++;
    }
  }

  Logger.log("Utlatel saved: " + written + " new, " + updated + " updated");
  CacheService.getScriptCache().remove("settings_data");
  return { written, updated };
}

// ── doGet HANDLERS ────────────────────────────────────────────
// Add these cases to your doGet function:
//   mode=settings   → returns shift history + utlatel data
//   mode=savesetting → saves a shift change (POST-like via GET params)
//   mode=saveutlatel → saves utlatel rows (JSON in params.data)

function handleSettingsMode() {
  const settings = getSettings();
  const utlatel  = getUtlatelData();
  return getCachedOrFetch("settings_data", () => ({ ...settings, utlatel }), 300);
}

function handleSaveShift(params) {
  const chaser        = params.chaser        || "";
  const shiftType     = params.shiftType     || "400"; // default shift minutes
  const effectiveDate = params.effectiveDate || getTodayTab();
  const notes         = params.notes         || "";
  if (!chaser) return { error: "Missing chaser name" };
  saveShiftSetting(chaser, shiftType, effectiveDate, notes);
  return { success: true };
}

function handleSaveUtlatel(params) {
  try {
    const rows = JSON.parse(params.data || "[]");
    const result = saveUtlatelData(rows);
    return { success: true, ...result };
  } catch(err) {
    return { error: err.message };
  }
}

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  const params = e.parameter;
  const mode   = params.mode || "archive";
  const date   = params.date || getTodayTab();

  let payload;
  try {

    if (mode === "archive") {
      payload = getCachedOrFetch("archive_all", getArchiveData, 300);

    } else if (mode === "settings") {
      payload = handleSettingsMode();

    } else if (mode === "savesetting") {
      payload = handleSaveShift(params);

    } else if (mode === "saveutlatel") {
      payload = handleSaveUtlatel(params);

    } else if (mode === "sync") {
      // Called by "Sync Today" button — reads live sheets for today,
      // overwrites today's archive rows, then clears cache
      const syncDate = date;

      // Delete existing rows for today so we can rewrite them
      deleteArchiveRowsForDate(syncDate);

      // Re-archive from live sheets
      archiveDayData(syncDate);
      archiveCampaignResponses(syncDate);

      // Clear archive cache so next dashboard load gets fresh data
      CacheService.getScriptCache().remove("archive_all");
      CacheService.getScriptCache().remove("campaign_responses");

      payload = { success: true, synced: syncDate, message: "Synced " + syncDate + " to archive" };

    } else if (mode === "campaignresponses") {
      payload = getCachedOrFetch("campaign_responses", getArchiveCampaignResponses, 300);

    } else if (mode === "chasers") {
      // Full roster (including inactive) for the dashboard's Settings tab.
      payload = { chasers: getChasersConfig() };

    } else if (mode === "savechaser") {
      // Add a new chaser or update an existing one's Sheet ID / Active flag.
      // Params: name, sheetId, active ("true"/"false")
      payload = saveChaser(params);

    } else if (mode === "leadsnapshot") {
      // Deduplicated Lead History snapshot as of the chosen date: per-campaign
      // leads/inProcess/verbalDenial, per-chaser active-lead counts, and
      // that day's new/concluded leads. Defined in LeadHistory.gs.
      const mdY = dateTabToFullDate(date);
      payload = getCachedOrFetch("leadsnapshot_" + mdY, () => getLeadSnapshotForDate(mdY), 300);

    } else if (mode === "leadactivity") {
      // Per-day New Leads / Leads Concluded across the whole Lead History
      // log, for the History tab to filter/sum over any date range
      // client-side. Defined in LeadHistory.gs.
      payload = getCachedOrFetch("lead_activity_daily", getLeadActivityByDay, 300);

    } else if (mode === "leadconflicts") {
      // Returns all PENDING conflicts from Lead History Conflicts tab for
      // dashboard review. Defined in LeadHistory.gs.
      payload = getPendingLeadConflictsForDashboard();

    } else if (mode === "resolveconflict") {
      // Called when a human resolves a Lead History conflict.
      // Params: conflictId, action, resolvedBy — action is one of
      // "pick:State", "bothvalid:State1,State2", "dataerror", "readytoretry",
      // "dismiss", "acknowledge". Defined in LeadHistory.gs.
      payload = resolveLeadConflict(
        params.conflictId, params.action, params.resolvedBy
      );
      CacheService.getScriptCache().remove("lead_conflicts");

    } else if (mode === "day") {
      // Legacy support — still works if called directly
      payload = getDayData(date);

    } else if (mode === "week") {
      // Legacy support
      payload = getWeekData(date);

    } else {
      payload = { error: "Unrecognized mode: " + mode };
    }

  } catch(err) {
    payload = { error: err.message };
  }

  return buildResponse(payload);
}

// Delete all archive rows for a specific date so they can be rewritten
function deleteArchiveRowsForDate(dateTab) {
  const ss           = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const tabName      = monthTabName(dateTab);
  const sheet        = ss.getSheetByName(tabName);
  if (!sheet) return;

  const data = sheet.getDataRange().getValues();
  // Collect row indices to delete (go backwards to preserve indices)
  const toDelete = [];
  for (let r = 1; r < data.length; r++) {
    if (normalizeDateCellToTab(data[r][0]) === dateTab) {
      toDelete.push(r + 1); // 1-indexed for Sheets API
    }
  }
  // Delete from bottom up
  toDelete.reverse().forEach(rowNum => sheet.deleteRow(rowNum));
  Logger.log("Deleted " + toDelete.length + " rows for " + dateTab);
}

// Generic cache helper — fetches and caches if not already cached
function getCachedOrFetch(key, fetchFn, ttlSeconds) {
  const cache  = CacheService.getScriptCache();
  const cached = cache.get(key);
  if (cached) {
    Logger.log("Cache hit: " + key);
    return JSON.parse(cached);
  }
  Logger.log("Cache miss: " + key);
  const data = fetchFn();
  try {
    const json = JSON.stringify(data);
    if (json.length < 90000) cache.put(key, json, ttlSeconds);
  } catch(e) {
    Logger.log("Cache write failed: " + e.message);
  }
  return data;
}

function buildResponse(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function clearCache() {
  CacheService.getScriptCache().removeAll([
    "day_" + getTodayTab(),
    "week_" + getTodayTab(),
    "archive_all"
  ]);
  Logger.log("Cache cleared");
}

// getArchiveData defined above (monthly tab version)

// ============================================================
// TODAY'S TAB NAME HELPER
// ============================================================
function getTodayTab() {
  const now = new Date();
  return (now.getMonth() + 1) + "/" + now.getDate();
}

// ============================================================
// SINGLE DAY DATA
// ============================================================
function getDayData(dateTab) {
  const chasers    = [];
  const teamTotals = { totalCases: 0, totalPositive: 0, totalTimeMins: 0, totalFaxes: 0 };

  for (const [name, sheetId] of Object.entries(getActiveChaserSheetMap())) {
    const data = readChaserTab(sheetId, dateTab, name);
    chasers.push(data);
    teamTotals.totalCases    += data.totalCases;
    teamTotals.totalPositive += data.totalPositive;
    teamTotals.totalTimeMins += data.totalTimeMins;
    teamTotals.totalFaxes    += data.totalFaxes;
  }

  const responses = readResponsesForDate(dateTab);

  return { mode: "day", date: dateTab, chasers, teamTotals, responses };
}

// ============================================================
// WEEK DATA
// ============================================================
function getWeekData(dateTab) {
  const tabs     = getWeekTabs(dateTab);
  const weekDays = tabs.map(tab => getDayData(tab));

  const weekTotals = {
    totalCases: 0, totalPositive: 0, totalTimeMins: 0,
    totalFaxes: 0, totalApproved: 0, totalDenied: 0
  };

  const chaserWeekMap = {};
  for (const name of Object.keys(getActiveChaserSheetMap())) {
    chaserWeekMap[name] = { name, totalCases: 0, totalPositive: 0, totalTimeMins: 0, totalFaxes: 0 };
  }

  for (const day of weekDays) {
    weekTotals.totalCases    += day.teamTotals.totalCases;
    weekTotals.totalPositive += day.teamTotals.totalPositive;
    weekTotals.totalTimeMins += day.teamTotals.totalTimeMins;
    weekTotals.totalFaxes    += day.teamTotals.totalFaxes;
    weekTotals.totalApproved += day.responses.approved;
    weekTotals.totalDenied   += day.responses.denied;

    for (const c of day.chasers) {
      if (chaserWeekMap[c.name]) {
        chaserWeekMap[c.name].totalCases    += c.totalCases;
        chaserWeekMap[c.name].totalPositive += c.totalPositive;
        chaserWeekMap[c.name].totalTimeMins += c.totalTimeMins;
        chaserWeekMap[c.name].totalFaxes    += c.totalFaxes;
      }
    }
  }

  return {
    mode: "week",
    weekDays,
    weekTotals,
    chaserWeekTotals: Object.values(chaserWeekMap)
  };
}

// ============================================================
// WEEK TABS
// ============================================================
function getWeekTabs(dateTab) {
  const year  = new Date().getFullYear();
  const parts = dateTab.split("/");
  const ref   = new Date(year, parseInt(parts[0]) - 1, parseInt(parts[1]));

  const day       = ref.getDay();
  const diffToMon = (day === 0) ? -6 : 1 - day;
  const monday    = new Date(ref);
  monday.setDate(ref.getDate() + diffToMon);

  const tabs = [];
  for (let i = 0; i < 5; i++) {
    const d = new Date(monday);
    d.setDate(monday.getDate() + i);
    tabs.push((d.getMonth() + 1) + "/" + d.getDate());
  }
  return tabs;
}

// ============================================================
// READ ONE CHASER TAB
// Only reads summary metrics — patient rows are not needed
// by the dashboard and are intentionally skipped.
// ============================================================
function readChaserTab(sheetId, dateTab, chaserName) {
  const result = {
    name: chaserName,
    totalCases: 0, totalPositive: 0, totalNegative: 0,
    totalTimeMins: 0, totalFaxes: 0,
    tabFound: false
  };

  try {
    const ss    = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheetByName(dateTab + ".");

    if (!sheet) {
      Logger.log(chaserName + " — tab not found: " + dateTab + ".");
      return result;
    }

    result.tabFound = true;

    // Read first 8 rows — all summary metrics live here
    // Layout varies: some sheets put all metrics in row 0, others spread across rows
    const data = sheet.getRange(1, 1, 8, sheet.getLastColumn()).getValues();

    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        // Normalize cell: collapse ALL whitespace (spaces, newlines, tabs) to single space
        const raw  = String(data[r][c]);
        const cell = raw.replace(/\s+/g, " ").trim().toUpperCase();

        if (cell.includes("TOTAL CHASED CASES")) {
          // Value is in the next cell on the same row
          result.totalCases    = Number(data[r][c + 1]) || 0;
        }

        if (cell.includes("TOTAL POSITIVE") || cell === "TOTAL POSITIVE") {
          result.totalPositive = Number(data[r][c + 1]) || 0;
        }

        if (cell.includes("TOTAL TIME TAKEN")) {
          result.totalTimeMins = Number(data[r][c + 1]) || 0;
        }

        // Match "FAXES SENT" or "PATIENTS FAXES SENT" — value is the NEXT non-empty
        // cell to the right, or the number in the row below the same column
        if (cell === "FAXES SENT" || cell === "FAXES SENT".toUpperCase() || cell.endsWith("FAXES SENT")) {
          // Try: number immediately to the right on same row
          const rightVal = Number(data[r][c + 1]);
          if (rightVal > 0) {
            result.totalFaxes = rightVal;
          } else {
            // Try: number in the row below, same column
            const belowVal = data[r + 1] ? Number(data[r + 1][c]) : 0;
            result.totalFaxes = belowVal || 0;
          }
        }
      }
    }

    // Log what we found for debugging
    Logger.log(chaserName + " | Cases=" + result.totalCases +
               " | Positive=" + result.totalPositive +
               " | Time=" + result.totalTimeMins +
               " | Faxes=" + result.totalFaxes);

  } catch (err) {
    result.error = err.message;
  }

  return result;
}

// ============================================================
// READ RESPONSES — returns byChaser AND byCampaign
// ============================================================
function readResponsesForDate(dateTab) {
  const result = {
    approved: 0,
    denied: 0,
    byChaser: {},
    byCampaign: zeroCampaignTotals(true)
  };

  for (const source of RESPONSE_SOURCES) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(source.id);
    } catch (err) {
      Logger.log("Could not open spreadsheet " + source.id + ": " + err.message);
      continue;
    }

    for (const tabConfig of source.tabs) {
      parseResponseTab(ss, tabConfig, dateTab, result);
    }
  }

  return result;
}

// Extract "M/D" date token from a feedback string and compare exactly.
// Prevents "5/1" from matching "5/10", "5/11", etc.
// feedbackStr: e.g. "Approved+CN 5/1" or "Received Denial 5/10 reason"
// dateTab:     e.g. "5/1" or "5/10"
function feedbackMatchesDate(feedbackStr, dateTab) {
  // Normalise dateTab to "M/D" (strip year if present)
  const tp   = dateTab.split("/");
  const want = parseInt(tp[0]) + "/" + parseInt(tp[1]);   // e.g. "5/1"

  // Extract every M/D or M/D/YYYY token from the feedback string
  const tokens = feedbackStr.match(/\b(\d{1,2}\/\d{1,2})(?:\/\d{2,4})?\b/g) || [];
  return tokens.some(tok => {
    const p = tok.split("/");
    return parseInt(p[0]) + "/" + parseInt(p[1]) === want;
  });
}

// ============================================================
// PARSE ONE RESPONSE TAB
// ============================================================
function parseResponseTab(ss, tabConfig, dateTab, result) {
  const sheet = ss.getSheetByName(tabConfig.name);
  if (!sheet) { Logger.log("Tab not found: " + tabConfig.name); return; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return;

  const headers         = data[0].map(h => String(h).trim().toUpperCase());
  const feedbackColName = tabConfig.feedbackCol.toUpperCase();
  const chaserColName   = tabConfig.chaserCol.toUpperCase();
  const feedbackCol     = headers.indexOf(feedbackColName);
  const chaserCol       = headers.indexOf(chaserColName);
  const campaignKey     = tabConfig.campaignKey || null;

  if (feedbackCol < 0) { Logger.log("Feedback column not found: " + tabConfig.feedbackCol + " in " + tabConfig.name); return; }
  if (chaserCol   < 0) { Logger.log("Chaser column not found: "   + tabConfig.chaserCol   + " in " + tabConfig.name); return; }

  for (let r = 1; r < data.length; r++) {
    const feedback   = String(data[r][feedbackCol] || "").trim();
    const chaserText = String(data[r][chaserCol]   || "").trim();

    if (!feedback || !chaserText)    continue;
    if (!feedbackMatchesDate(feedback, dateTab)) continue;

    const approval = isApproval(feedback);
    const denial   = isDenial(feedback);
    if (!approval && !denial) continue;

    const chasers = chaserText.split("/").map(x => x.trim()).filter(Boolean);

    // ── per-chaser ──
    for (const chaser of chasers) {
      if (!result.byChaser[chaser]) {
        result.byChaser[chaser] = {
          approvals: 0, denials: 0,
          // per-campaign breakdown per chaser
          campaigns: zeroCampaignTotals()
        };
      }
      if (approval) { result.byChaser[chaser].approvals++; result.approved++; }
      if (denial)   { result.byChaser[chaser].denials++;   result.denied++;   }

      // per-chaser per-campaign
      if (campaignKey && result.byChaser[chaser].campaigns[campaignKey]) {
        if (approval) result.byChaser[chaser].campaigns[campaignKey].approved++;
        if (denial)   result.byChaser[chaser].campaigns[campaignKey].denied++;
      }
    }

    // ── per-campaign team totals ──
    if (campaignKey && result.byCampaign[campaignKey]) {
      const camp = result.byCampaign[campaignKey];
      camp.total++;
      if (approval) camp.approved++;
      if (denial)   camp.denied++;
    }
  }
}

// ============================================================
// UTILITY
// ============================================================
function getAvailableTabs() {
  const tabs = {};
  for (const [name, sheetId] of Object.entries(CHASER_SHEETS)) {
    try {
      const ss   = SpreadsheetApp.openById(sheetId);
      tabs[name] = ss.getSheets().map(s => s.getName());
    } catch (e) {
      tabs[name] = [];
    }
  }
  return tabs;
}

// ============================================================
// TEST FUNCTIONS
// ============================================================
function testDay() {
  Logger.log(JSON.stringify(getDayData("6/25"), null, 2));
}

function testResponses() {
  const data = readResponsesForDate("6/25");
  Logger.log("Approved: " + data.approved);
  Logger.log("Denied:   " + data.denied);
  Logger.log("byChaser: "   + JSON.stringify(data.byChaser,   null, 2));
  Logger.log("byCampaign: " + JSON.stringify(data.byCampaign, null, 2));
}

function testArchive() {
  const data = getArchiveData();
  Logger.log("Archive rows: " + data.rows.length);
  if (data.rows.length) Logger.log("First row: " + JSON.stringify(data.rows[0]));
}

function testAllChasers() {
  const dateTab = "6/25";
  let grandCases = 0, grandPositive = 0, grandTime = 0, grandFaxes = 0;
  Object.entries(CHASER_SHEETS).forEach(([name, sheetId]) => {
    const data = readChaserTab(sheetId, dateTab, name);
    Logger.log(name + " | Cases=" + data.totalCases + " | Positive=" + data.totalPositive +
               " | Time=" + data.totalTimeMins + " | Faxes=" + data.totalFaxes);
    grandCases    += data.totalCases    || 0;
    grandPositive += data.totalPositive || 0;
    grandTime     += data.totalTimeMins || 0;
    grandFaxes    += data.totalFaxes    || 0;
  });
  Logger.log("TOTALS | Cases=" + grandCases + " | Positive=" + grandPositive +
             " | Time=" + grandTime + " | Faxes=" + grandFaxes);
}

function debugAlex() {
  const ss = SpreadsheetApp.openById(CHASER_SHEETS.Alex);
  ss.getSheets().forEach(sheet => Logger.log("Tab: " + sheet.getName()));
}


// ============================================================
// BACKFILL — reads historical tabs and writes to archive
//
// HOW TO USE:
//   1. Set FROM_DATE and TO_DATE below to the range you want
//   2. Run backfillDateRange() from the Apps Script editor
//   3. Check the archive sheet — one row per chaser per day
//
// NOTES:
//   - Skips weekends automatically
//   - Skips dates that are already in the archive (no duplicates)
//   - If a chaser has no tab for that date it writes zeros for that row
//   - Response data (approvals/denials) is only available if your
//     response sheets still contain those dates — older entries may
//     have been deleted, in which case approvals/denials will be 0
// ============================================================

// DEPRECATED — writes to old single "Data Archive" tab, not monthly tabs.
// Use archiveDayData() or the eodArchive trigger instead.
// Kept here for reference only — do not run.
function backfillDateRange_DEPRECATED() {
  const FROM_DATE = "2026-01-01";   // ← change to your start date (YYYY-MM-DD)
  const TO_DATE   = "2026-06-24";   // ← change to your end date   (YYYY-MM-DD)

  const from = new Date(FROM_DATE);
  const to   = new Date(TO_DATE);

  // Load existing archive dates to avoid duplicates
  const ss        = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet     = ss.getSheetByName("Data Archive");
  const existing  = new Set();

  if (!sheet) {
    sheet = ss.insertSheet("Data Archive");
    sheet.appendRow([
      "Date","Chaser","Cases","Positive",
      "Approvals","Denials","TimeMins","Efficiency",
      "ORT_Approved","ORT_Denied",
      "CGM_Approved","CGM_Denied",
      "LymphC_Approved","LymphC_Denied",
      "LymphW_Approved","LymphW_Denied"
    ]);
  } else {
    // Build set of "date|chaser" combos already archived
    const existingData = sheet.getDataRange().getValues();
    for (let r = 1; r < existingData.length; r++) {
      const key = String(existingData[r][0]) + "|" + String(existingData[r][1]);
      existing.add(key);
    }
  }

  Logger.log("Starting backfill from " + FROM_DATE + " to " + TO_DATE);
  Logger.log("Existing archive rows: " + existing.size);

  let written = 0;
  let skipped = 0;
  let current = new Date(from);

  while (current <= to) {
    const dayOfWeek = current.getDay();

    // Skip weekends
    if (dayOfWeek === 0 || dayOfWeek === 6) {
      current.setDate(current.getDate() + 1);
      continue;
    }

    const dateTab = (current.getMonth() + 1) + "/" + current.getDate();

    // Read responses for this date (shared across all chasers)
    let responses;
    try {
      responses = readResponsesForDate(dateTab);
    } catch(err) {
      Logger.log("Could not read responses for " + dateTab + ": " + err.message);
      responses = { approved:0, denied:0, byChaser:{}, byCampaign: zeroCampaignTotals(true) };
    }

    const bc = responses.byChaser  || {};
    const cp = responses.byCampaign || {};

    for (const [name, sheetId] of Object.entries(CHASER_SHEETS)) {
      const key = dateTab + "|" + name;

      if (existing.has(key)) {
        skipped++;
        continue;
      }

      const c   = readChaserTab(sheetId, dateTab, name);
      const r   = bc[name] || { approvals: 0, denials: 0 };
      const eff = c.totalCases ? (c.totalPositive / c.totalCases * 100).toFixed(1) : "";

      sheet.appendRow([
        dateTab, name, c.totalCases, c.totalPositive,
        r.approvals, r.denials, c.totalTimeMins, eff,
        cp.ort?.approved    || 0, cp.ort?.denied    || 0,
        cp.cgm?.approved    || 0, cp.cgm?.denied    || 0,
        cp.lymphc?.approved || 0, cp.lymphc?.denied || 0,
        cp.lymphw?.approved || 0, cp.lymphw?.denied || 0,
      ]);

      existing.add(key);
      written++;
    }

    Logger.log("Done: " + dateTab + " | written so far: " + written);
    current.setDate(current.getDate() + 1);

    // Pause every 20 days to avoid hitting Apps Script time limits
    // (uncomment if your range is very large, e.g. a full year)
    // Utilities.sleep(500);
  }

  Logger.log("=== BACKFILL COMPLETE ===");
  Logger.log("Written: " + written + " rows");
  Logger.log("Skipped (already existed): " + skipped + " rows");
}

// Quick test — backfills just yesterday
function backfillYesterday() {
  const d = new Date();
  d.setDate(d.getDate() - 1);
  const dateTab = (d.getMonth() + 1) + "/" + d.getDate();
  Logger.log("Backfilling: " + dateTab);
  archiveDayData(dateTab);
  Logger.log("Done");
}

// Backfill a single specific date
function backfillSpecificDate() {
  const DATE = "6/24";   // ← change to the date you want
  archiveDayData(DATE);
}


// ============================================================
// MIGRATE EXISTING TEAM SHEETS → ARCHIVE
//
// Source sheets have weekly tabs named "WEEK 2", "WEEK 3", etc.
// Inside each tab, rows are grouped by date with a date header
// row (e.g. "6/23/2026") followed by chaser data rows, separated
// by blank rows.
//
// HOW TO USE:
//   Run migrateExistingSheets() once from the Apps Script editor.
//   Safe to re-run — duplicate date+chaser combos are skipped.
// ============================================================

const EXISTING_SHEETS = [
  "1gvphK7eoaVLxkP3KAtK7gECkSVF1JFZG37KWuCi3_bY",  // Team 1
  "1C-8cmILCoPvXqTgcJwe-mYFv1cNI_p5CU_FrdArzdBw"   // Team 2
];

// Column header → archive field mapping (case-insensitive matching)
const COL_MAP = {
  "chaser name":           "chaser",
  "total shift":           "totalShift",
  "total calls":           "calls",
  "total duration":        "totalDuration",
  "total duration (min)":  "totalDurationMins",
  "acw duration":          "acwDuration",
  "faxes\nsent\n(duration)": "faxes",
  "faxes sent (duration)": "faxes",
  "faxes sent":            "faxes",
  "productive time":       "productiveTime",
  "total chased cases":    "cases",
  "total positive":        "positive",
  "total time taken":      "timeMins",
  "approvals":             "approvals",
  "denials":               "denials",
  "total responses":       "totalResponses",
  "productivity target":   "productivityTarget",
  "efficiency target":     "efficiencyTarget",
  "productivity":          "productivity",
  "efficiency":            "efficiency",
};

function migrateExistingSheets() {
  // ── Open / prepare archive sheet ──────────────────────────
  const archiveSS = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);

  // Build a combined duplicate guard across ALL month tabs
  const existing    = new Set();
  const monthPattern = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i;
  for (const sheet of archiveSS.getSheets()) {
    if (!monthPattern.test(sheet.getName().trim())) continue;
    const d = sheet.getDataRange().getValues();
    for (let r = 1; r < d.length; r++) {
      existing.add(normalizeDateCellToTab(d[r][0]) + "|" + String(d[r][1]));
    }
  }
  Logger.log("Existing archive rows across all month tabs: " + existing.size);

  let totalWritten = 0;
  let totalSkipped = 0;

  // ── Loop both source spreadsheets ─────────────────────────
  for (const ssId of EXISTING_SHEETS) {
    let ss;
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch(err) {
      Logger.log("Could not open sheet " + ssId + ": " + err.message);
      continue;
    }

    Logger.log("Processing: " + ss.getName());

    // Get all tabs named "WEEK X" (any number)
    const weekTabs = ss.getSheets().filter(s => /^WEEK\s+\d+$/i.test(s.getName().trim()));
    Logger.log("Found " + weekTabs.length + " week tabs in " + ss.getName());

    for (const tab of weekTabs) {
      Logger.log("  Reading tab: " + tab.getName());
      const rows = parseWeekTab(tab, archiveSS, existing);
      totalWritten += rows.written;
      totalSkipped += rows.skipped;
    }
  }

  Logger.log("=== MIGRATION COMPLETE ===");
  Logger.log("Written: " + totalWritten + " rows");
  Logger.log("Skipped (duplicates): " + totalSkipped + " rows");
}

// ── Parse one weekly tab ───────────────────────────────────────
function parseWeekTab(tab, archiveSS, existing) {
  const data    = tab.getDataRange().getValues();
  let written   = 0;
  let skipped   = 0;
  let currentDate = null;
  let colIndex    = {};   // field name → column index, reset per date block

  for (let r = 0; r < data.length; r++) {
    const row      = data[r];
    const firstVal = String(row[0] || "").trim();

    // ── Blank row: reset date context ──────────────────────
    if (row.every(c => String(c).trim() === "")) {
      currentDate = null;
      colIndex    = {};
      continue;
    }

    // ── Date header row: "6/23/2026" or a Date object ──────
    const parsedDate = tryParseDate(firstVal, row[0]);
    if (parsedDate) {
      currentDate = parsedDate;
      colIndex    = {};
      Logger.log("    Date: " + currentDate);
      continue;
    }

    // ── Column header row: contains "Chaser Name" ──────────
    if (isHeaderRow(row)) {
      colIndex = buildColIndex(row);
      continue;
    }

    // ── Data row ────────────────────────────────────────────
    if (!currentDate || Object.keys(colIndex).length === 0) continue;

    const chaser = String(row[colIndex.chaser] !== undefined ? row[colIndex.chaser] : "").trim();
    if (!chaser || chaser.toUpperCase() === "CHASER NAME") continue;

    const key = currentDate + "|" + chaser;
    if (existing.has(key)) { skipped++; continue; }

    const get = field => colIndex[field] !== undefined ? row[colIndex[field]] : "";

    // Parse productivity/efficiency — strip % if stored as string
    const prodRaw  = String(get("productivity") || "").replace("%","").trim();
    const effRaw   = String(get("efficiency")   || "").replace("%","").trim();
    const prod     = parseFloat(prodRaw) || "";
    const eff      = parseFloat(effRaw)  || "";

    const rowTabName  = monthTabName(currentDate);
    const rowSheet    = getOrCreateMonthTab(archiveSS, rowTabName);

    // Write week header if this is the first row for a Monday in this sheet
    const isMonDate   = parseDateTab(currentDate).getDay() === 1;
    const weekKey     = "WEEKHEADER|" + currentDate;
    if (isMonDate && !existing.has(weekKey)) {
      const weekLabel = getWeekLabel(currentDate);
      rowSheet.appendRow(["WEEK", weekLabel].concat(new Array(ARCHIVE_HEADERS.length - 2).fill("")));
      const lastRow = rowSheet.getLastRow();
      const hRange  = rowSheet.getRange(lastRow, 1, 1, ARCHIVE_HEADERS.length);
      hRange.setBackground("#1A2C42");
      hRange.setFontColor("#00C2A8");
      hRange.setFontWeight("bold");
      existing.add(weekKey);
    }

    rowSheet.appendRow([
      currentDate,
      chaser,
      Number(get("cases"))            || 0,   // Cases
      Number(get("positive"))         || 0,   // Positive
      Number(get("approvals"))        || 0,   // Approvals (NO faxes col)
      Number(get("denials"))          || 0,   // Denials
      Number(get("timeMins"))         || 0,   // TimeMins
      eff,                                    // Efficiency
      prod,                                   // Productivity
      Number(get("totalShift"))       || 0,   // TotalShift
      Number(get("calls"))            || 0,   // TotalCalls
      Number(get("totalDurationMins"))|| 0,   // TotalDurationMins
      Number(get("acwDuration"))      || 0,   // ACWDuration
      Number(get("productiveTime"))   || 0,   // ProductiveTime
      // Campaign columns — not in source sheets, default to 0
      // (eodArchive fills these going forward)
      0, 0, 0, 0, 0, 0, 0, 0
    ]);

    existing.add(key);
    written++;
  }

  return { written, skipped };
}

// ── Helpers ────────────────────────────────────────────────────

// Try to detect a date header cell — handles string "6/23/2026" or Date objects
function tryParseDate(strVal, rawVal) {
  // Google Sheets may give us a Date object directly
  if (rawVal instanceof Date && !isNaN(rawVal)) {
    const d = rawVal;
    return `${d.getMonth()+1}/${d.getDate()}/${d.getFullYear()}`;
  }
  // Match "6/23/2026", "06/23/2026", "6/23/26"
  const m = strVal.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
  if (m) {
    const yr = m[3].length === 2 ? "20" + m[3] : m[3];
    return `${parseInt(m[1])}/${parseInt(m[2])}/${yr}`;
  }
  return null;
}

// A header row is one whose first non-empty cell matches known headers
function isHeaderRow(row) {
  const firstNonEmpty = row.map(c => String(c).trim().toLowerCase()).find(c => c !== "");
  return firstNonEmpty === "chaser name" || firstNonEmpty === "chaser";
}

// Build colIndex map: field → column number
function buildColIndex(row) {
  const idx = {};
  row.forEach((cell, c) => {
    const key = String(cell).trim().toLowerCase();
    const mapped = COL_MAP[key];
    if (mapped && idx[mapped] === undefined) {
      idx[mapped] = c;
    }
  });
  return idx;
}

// Test migration without writing — just logs what it would do
function dryRunMigration() {
  for (const ssId of EXISTING_SHEETS) {
    try {
      const ss       = SpreadsheetApp.openById(ssId);
      const weekTabs = ss.getSheets().filter(s => /^WEEK\s+\d+$/i.test(s.getName().trim()));
      Logger.log(ss.getName() + " → " + weekTabs.map(t=>t.getName()).join(", "));

      for (const tab of weekTabs) {
        const data = tab.getDataRange().getValues();
        let dateCount = 0;
        let rowCount  = 0;
        for (const row of data) {
          const first = String(row[0]||"").trim();
          if (tryParseDate(first, row[0])) dateCount++;
          else if (!isHeaderRow(row) && first && !row.every(c=>String(c).trim()==="")) rowCount++;
        }
        Logger.log("  " + tab.getName() + " → " + dateCount + " date blocks, ~" + rowCount + " data rows");
      }
    } catch(err) {
      Logger.log("Error: " + ssId + " — " + err.message);
    }
  }
}
// renameArchiveFaxColumn removed — faxes column deleted from archive
function renameArchiveFaxColumn_DEPRECATED() {
  const ss           = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const MONTH_NAMES  = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  let   tabsUpdated  = 0;

  for (const sheet of ss.getSheets()) {
    const tabName = sheet.getName().trim().toLowerCase().replace(/\s+/g," ");
    const isMonth = MONTH_NAMES.some(m => tabName.startsWith(m + " ") && /\d{4}$/.test(tabName));
    if (!isMonth) continue;

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const faxIdx  = headers.findIndex(h => String(h).trim() === "Faxes");

    if (faxIdx >= 0) {
      sheet.getRange(1, faxIdx + 1).setValue("FaxDurationMins");
      Logger.log("Updated: " + sheet.getName() + " col " + (faxIdx+1));
      tabsUpdated++;
    } else {
      Logger.log("Already updated or not found: " + sheet.getName());
    }
  }

  Logger.log("Done. Tabs updated: " + tabsUpdated);
}


// ============================================================
// ONE-TIME: Add bottom borders to last chaser row of each day
// Run once to apply borders to all existing archive data
// ============================================================
function applyDayBordersToArchive() {
  const ss          = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const MONTH_NAMES = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  let   totalBorders = 0;

  for (const sheet of ss.getSheets()) {
    const tabName = sheet.getName().trim().toLowerCase().replace(/\s+/g," ");
    const isMonth = MONTH_NAMES.some(m => tabName.startsWith(m + " ") && /\d{4}$/.test(tabName));
    if (!isMonth) continue;

    Logger.log("Processing: " + sheet.getName());
    const data     = sheet.getDataRange().getValues();
    const numCols  = ARCHIVE_HEADERS.length;

    // Find the last row for each date and apply border
    let currentDate = null;
    let lastRowForDate = -1;

    for (let r = 1; r < data.length; r++) {
      const dateVal   = String(data[r][0]).trim();
      const chaserVal = String(data[r][1]).trim();

      // Skip blank, WEEK header, and Total rows
      if (!dateVal && !chaserVal) continue;
      if (dateVal.toUpperCase() === "WEEK") continue;
      if (chaserVal.toUpperCase().startsWith("TOTAL")) continue;
      if (!chaserVal) continue;

      const rowDate = dateVal || currentDate;

      if (rowDate !== currentDate) {
        // New date — apply border to last row of previous date
        if (lastRowForDate > 0) {
          sheet.getRange(lastRowForDate, 1, 1, numCols).setBorder(
            null, null, true, null, null, null,
            "#00C2A8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
          );
          totalBorders++;
        }
        currentDate    = rowDate;
        lastRowForDate = r + 1; // 1-indexed
      } else {
        lastRowForDate = r + 1;
      }
    }

    // Apply border to the very last date in the sheet
    if (lastRowForDate > 0) {
      sheet.getRange(lastRowForDate, 1, 1, numCols).setBorder(
        null, null, true, null, null, null,
        "#00C2A8", SpreadsheetApp.BorderStyle.SOLID_MEDIUM
      );
      totalBorders++;
    }

    Logger.log("  Borders applied: " + totalBorders);
  }

  Logger.log("=== DONE === Total day borders applied: " + totalBorders);
}




// ============================================================
// BACKFILL CAMPAIGN COLUMNS IN EXISTING ARCHIVE ROWS
//
// Reads response sheets for every date in the archive and
// updates ORT/CGM/LymphC/LymphW approved/denied columns.
//
// HOW TO USE:
//   Run backfillCampaignColumns() once.
//   It processes one month tab at a time — if it times out,
//   just run it again (already-updated rows are skipped).
//
// PROGRESS KEY: "campaign_backfill_progress"
// ============================================================

function backfillCampaignColumns() {
  const PROGRESS_KEY  = "campaign_backfill_progress";
  const props         = PropertiesService.getScriptProperties();
  const doneDates     = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const ss          = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const MONTH_NAMES = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"];
  const monthPattern = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i;

  // Column indices in archive (0-based after header)
  // Headers: Date,Chaser,Cases,Positive,Approvals,Denials,TimeMins,Efficiency,Productivity,
  //          TotalShift,TotalCalls,TotalDurationMins,ACWDuration,ProductiveTime,
  //          ORT_Approved,ORT_Denied,CGM_Approved,CGM_Denied,
  //          LymphC_Approved,LymphC_Denied,LymphW_Approved,LymphW_Denied
  const COL = {
    date:          0,
    chaser:        1,
    ort_approved:  14,
    ort_denied:    15,
    cgm_approved:  16,
    cgm_denied:    17,
    lymphc_approved:18,
    lymphc_denied:  19,
    lymphw_approved:20,
    lymphw_denied:  21,
  };

  let totalUpdated = 0;
  let totalSkipped = 0;

  for (const sheet of ss.getSheets()) {
    const tabName = sheet.getName().trim();
    if (!monthPattern.test(tabName)) continue;

    Logger.log("Processing tab: " + tabName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    // Collect unique dates in this tab that haven't been processed
    const datesToProcess = new Set();
    for (let r = 1; r < data.length; r++) {
      const rawDate   = data[r][COL.date];
      const chaserVal = String(data[r][COL.chaser]).trim();
      if (!chaserVal || chaserVal.toUpperCase().startsWith("TOTAL")) continue;

      let dateStr;
      if (rawDate instanceof Date && !isNaN(rawDate)) {
        dateStr = (rawDate.getMonth()+1) + "/" + rawDate.getDate() + "/" + rawDate.getFullYear();
      } else {
        dateStr = String(rawDate).trim();
      }
      if (!dateStr || dateStr.toUpperCase() === "WEEK") continue;

      // Normalize to M/D for response sheet lookup
      const parts = dateStr.split("/");
      const shortDate = parseInt(parts[0]) + "/" + parseInt(parts[1]);

      if (!doneDates.includes(shortDate)) {
        datesToProcess.add(shortDate);
      } else {
        totalSkipped++;
      }
    }

    Logger.log("  Dates to process: " + datesToProcess.size + " | Already done: " + totalSkipped);

    // For each unique date, fetch per-chaser campaign counts and update archive rows
    for (const shortDate of datesToProcess) {
      Logger.log("  Fetching per-chaser campaign counts for: " + shortDate);

      let chaserCamps;
      try {
        // readChaserCampaignCountsForDate: reads response sheets for this day,
        // returns normalized chaser names mapped to per-campaign lead counts.
        // Each chaser gets +1 per lead they were listed on (daily, not week totals).
        chaserCamps = readChaserCampaignCountsForDate(shortDate);
      } catch(err) {
        Logger.log("  Error reading responses for " + shortDate + ": " + err.message);
        doneDates.push(shortDate);
        props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
        continue;
      }

      let rowsUpdated = 0;

      for (let r = 1; r < data.length; r++) {
        const rawDate   = data[r][COL.date];
        const chaserVal = String(data[r][COL.chaser]).trim();
        if (!chaserVal || chaserVal.toUpperCase().startsWith("TOTAL")) continue;
        if (chaserVal.toUpperCase() === "WEEK") continue;

        let dateStr;
        if (rawDate instanceof Date && !isNaN(rawDate)) {
          dateStr = (rawDate.getMonth()+1) + "/" + rawDate.getDate() + "/" + rawDate.getFullYear();
        } else {
          dateStr = String(rawDate).trim();
        }
        if (!dateStr || dateStr.toUpperCase() === "WEEK") continue;

        const parts = dateStr.split("/");
        const rowShortDate = parseInt(parts[0]) + "/" + parseInt(parts[1]);
        if (rowShortDate !== shortDate) continue;

        // Archive stores normalized short names; chaserCamps also uses normalized names
        const normalized = normalizeChaserName(chaserVal);
        const cc = chaserCamps[normalized] || chaserCamps[chaserVal] || zeroCampaignTotals();

        // Write all 8 campaign columns in one batch (faster than 8 individual setValue calls)
        sheet.getRange(r+1, COL.ort_approved+1, 1, 8).setValues([[
          cc.ort.approved,    cc.ort.denied,
          cc.cgm.approved,    cc.cgm.denied,
          cc.lymphc.approved, cc.lymphc.denied,
          cc.lymphw.approved, cc.lymphw.denied,
        ]]);
        rowsUpdated++;
      }

      Logger.log("  Updated " + rowsUpdated + " rows for " + shortDate);
      totalUpdated += rowsUpdated;

      // Save progress after each date
      doneDates.push(shortDate);
      props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
    }
  }

  Logger.log("=== DONE ===");
  Logger.log("Rows updated: " + totalUpdated);
  Logger.log("Dates skipped (already done): " + (doneDates.length - datesToProcess.size));
  Logger.log("Total dates processed so far: " + doneDates.length);

  // Check if all done
  Logger.log("Run backfillCampaignColumns() again if there are more dates to process.");
}

// Reset campaign backfill progress
function resetCampaignBackfill() {
  PropertiesService.getScriptProperties().deleteProperty("campaign_backfill_progress");
  Logger.log("Campaign backfill progress reset.");
}

// Check progress
function checkCampaignBackfillProgress() {
  const done = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("campaign_backfill_progress") || "[]"
  );
  Logger.log("Dates with campaign data backfilled: " + done.length);
}

// ============================================================
// CAMPAIGN RESPONSES TAB
// ============================================================
// Stores one aggregated row per date+campaign — no per-chaser
// rows, no per-lead rows. Each row is the daily total for that
// campaign: how many unique leads were approved/denied.
//
// Tab name: "Campaign Responses"
// Columns:  Date | Campaign | Approved | Denied | Approval%
//
// This is the ONLY correct source for campaign-level totals on
// the dashboard. The per-chaser ORT/CGM columns in monthly tabs
// store how many leads each chaser was listed on (for the
// individual chaser pages) — those numbers are correct per-chaser
// but must NOT be summed for campaign totals (would double-count).
//
// USAGE:
//   backfillCampaignResponses()  — one-time: reads all history
//   archiveCampaignResponses(dateTab) — daily (called by eodArchive + sync)
//   getArchiveCampaignResponses() — returns rows for dashboard
// ============================================================

const CAMPAIGN_RESPONSES_HEADERS = ["Date","Campaign","Approved","Denied","ApprovalPct"];

function getOrCreateCampaignResponsesTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Campaign Responses");
  if (!sheet) {
    sheet = ss.insertSheet("Campaign Responses");
    sheet.appendRow(CAMPAIGN_RESPONSES_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CAMPAIGN_RESPONSES_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#1A2C42")
         .setFontColor("#00C2A8");
    Logger.log("Created Campaign Responses tab");
  }
  return sheet;
}

// Read all leads for a date from response sheets.
// Returns aggregated counts per campaign: { ort:{approved,denied}, cgm:..., ... }
// Each lead row counts once regardless of how many chasers are listed.
function readCampaignTotalsForDate(dateTab) {
  const camps = zeroCampaignTotals();

  for (const source of RESPONSE_SOURCES) {
    let ss;
    try { ss = SpreadsheetApp.openById(source.id); }
    catch(e) { Logger.log("Cannot open " + source.id + ": " + e.message); continue; }

    for (const tabConfig of source.tabs) {
      const sheet = ss.getSheetByName(tabConfig.name);
      if (!sheet) continue;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      const headers     = data[0].map(h => String(h).trim().toUpperCase());
      const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
      const chaserCol   = headers.indexOf(tabConfig.chaserCol.toUpperCase());
      if (feedbackCol < 0 || chaserCol < 0) continue;

      const key = tabConfig.campaignKey;  // "ort", "cgm", "lymphc", "lymphw"

      for (let r = 1; r < data.length; r++) {
        const feedback = String(data[r][feedbackCol] || "").trim();
        if (!feedback || !feedbackMatchesDate(feedback, dateTab)) continue;

        const approval = isApproval(feedback);
        const denial   = isDenial(feedback);
        if (!approval && !denial) continue;

        // Count one per lead row — chasers column is irrelevant for campaign totals
        if (approval) camps[key].approved++;
        if (denial)   camps[key].denied++;
      }
    }
  }

  return camps;
}

// Read per-chaser per-campaign counts for a date.
// Returns byChaser[name].campaigns.{ort/cgm/...}.{approved/denied}
// Used to populate the ORT_Approved etc. columns in monthly archive rows.
function readChaserCampaignCountsForDate(dateTab) {
  const byChaser = {};

  for (const source of RESPONSE_SOURCES) {
    let ss;
    try { ss = SpreadsheetApp.openById(source.id); }
    catch(e) { continue; }

    for (const tabConfig of source.tabs) {
      const sheet = ss.getSheetByName(tabConfig.name);
      if (!sheet) continue;

      const data = sheet.getDataRange().getValues();
      if (data.length < 2) continue;

      const headers     = data[0].map(h => String(h).trim().toUpperCase());
      const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
      const chaserCol   = headers.indexOf(tabConfig.chaserCol.toUpperCase());
      if (feedbackCol < 0 || chaserCol < 0) continue;

      const campKey = tabConfig.campaignKey;

      for (let r = 1; r < data.length; r++) {
        const feedback   = String(data[r][feedbackCol] || "").trim();
        const chaserText = String(data[r][chaserCol]   || "").trim();
        if (!feedback || !chaserText || !feedbackMatchesDate(feedback, dateTab)) continue;

        const approval = isApproval(feedback);
        const denial   = isDenial(feedback);
        if (!approval && !denial) continue;

        // Each chaser listed on this lead gets +1 for this campaign
        const chasers = chaserText.split("/").map(x => normalizeChaserName(x.trim())).filter(Boolean);
        for (const name of chasers) {
          if (!byChaser[name]) {
            byChaser[name] = zeroCampaignTotals();
          }
          if (approval) byChaser[name][campKey].approved++;
          if (denial)   byChaser[name][campKey].denied++;
        }
      }
    }
  }

  return byChaser;
}

// Write or overwrite campaign totals for a date in the Campaign Responses tab.
// Format: one row per campaign per date.
// Deletes existing rows for this date first (so re-running is safe).
function archiveCampaignResponses(dateTab) {
  const sheet = getOrCreateCampaignResponsesTab();
  const data  = sheet.getDataRange().getValues();

  // Delete any existing rows for this date (work backwards to preserve indices)
  const toDelete = [];
  for (let r = 1; r < data.length; r++) {
    const rowDate = data[r][0] instanceof Date
      ? (data[r][0].getMonth()+1)+"/"+data[r][0].getDate()
      : String(data[r][0]).trim();
    // Normalize to M/D for comparison
    const rParts = rowDate.split("/");
    const rShort = parseInt(rParts[0]) + "/" + parseInt(rParts[1]);
    const tParts = dateTab.split("/");
    const tShort = parseInt(tParts[0]) + "/" + parseInt(tParts[1]);
    if (rShort === tShort) toDelete.push(r + 1);
  }
  toDelete.reverse().forEach(rowNum => sheet.deleteRow(rowNum));

  // Read fresh totals from response sheets
  const camps = readCampaignTotalsForDate(dateTab);

  let written = 0;
  for (const [key, label] of Object.entries(CAMPAIGN_LABELS)) {
    const c   = camps[key];
    const tot = c.approved + c.denied;
    const pct = tot ? parseFloat((c.approved / tot * 100).toFixed(1)) : 0;
    sheet.appendRow([dateTab, label, c.approved, c.denied, pct]);
    written++;
  }

  Logger.log("archiveCampaignResponses(" + dateTab + "): wrote " + written + " campaign rows");
  return written;
}

// One-time backfill: populate Campaign Responses tab for all dates in archive.
// Safe to re-run — existing rows for each date are deleted and rewritten.
// Progress stored in script properties so it can resume after a timeout.
function backfillCampaignResponses() {
  const PROGRESS_KEY = "campaign_responses_backfill2";
  const props        = PropertiesService.getScriptProperties();
  const doneDates    = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  // Collect every unique M/D date from all monthly archive tabs
  const ss           = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const monthPattern = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$/i;
  const allDates     = new Set();

  for (const sheet of ss.getSheets()) {
    if (!monthPattern.test(sheet.getName().trim())) continue;
    const d = sheet.getDataRange().getValues();
    for (let r = 1; r < d.length; r++) {
      const rawDate = d[r][0];
      const chaser  = String(d[r][1]).trim();
      if (!chaser || chaser.toUpperCase() === "WEEK") continue;

      let dateStr;
      if (rawDate instanceof Date && !isNaN(rawDate)) {
        dateStr = (rawDate.getMonth()+1) + "/" + rawDate.getDate();
      } else {
        const s = String(rawDate).trim();
        if (!s || s.toUpperCase() === "WEEK") continue;
        const parts = s.split("/");
        dateStr = parseInt(parts[0]) + "/" + parseInt(parts[1]);
      }
      if (dateStr && !doneDates.includes(dateStr)) allDates.add(dateStr);
    }
  }

  Logger.log("Dates to process: " + allDates.size + " | Already done: " + doneDates.length);

  let totalWritten = 0;
  for (const dateTab of allDates) {
    try {
      totalWritten += archiveCampaignResponses(dateTab);
    } catch(e) {
      Logger.log("Error for " + dateTab + ": " + e.message);
    }
    doneDates.push(dateTab);
    props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
  }

  Logger.log("=== DONE: " + totalWritten + " rows written across " + allDates.size + " dates ===");
  Logger.log("Run again if it timed out before finishing.");
}

function resetCampaignResponsesBackfill() {
  PropertiesService.getScriptProperties().deleteProperty("campaign_responses_backfill2");
  Logger.log("Campaign responses backfill progress reset.");
}

// Return all Campaign Responses rows for the dashboard.
// Returns { rows: [{ Date, Campaign, Approved, Denied, ApprovalPct }] }
function getArchiveCampaignResponses() {
  const sheet = getOrCreateCampaignResponsesTab();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { rows: [] };

  const headers = data[0].map(h => String(h).trim());
  const rows    = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    // Normalize date
    if (obj.Date instanceof Date && !isNaN(obj.Date)) {
      const d = obj.Date;
      obj.Date = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getFullYear();
    } else {
      obj.Date = String(obj.Date || "").trim();
    }
    // Ensure numeric types
    obj.Approved    = Number(obj.Approved)    || 0;
    obj.Denied      = Number(obj.Denied)      || 0;
    obj.ApprovalPct = Number(obj.ApprovalPct) || 0;
    return obj;
  }).filter(r => r.Date && r.Campaign);

  Logger.log("Campaign Responses rows returned: " + rows.length);
  return { rows };
}

// ============================================================
// ONE-TIME HISTORICAL BACKFILL — Combined Weekly Sheet
// ============================================================
// A separate sheet was found that holds responses for ALL 4 campaigns,
// updated weekly, in 4 tabs (one per campaign). This is used ONCE to
// backfill historical Campaign Responses + per-chaser campaign columns
// in the archive. It is completely separate from RESPONSE_SOURCES, which
// remains the daily source of truth — running this does not change
// day-to-day syncing in any way.
//
// Sheet ID: 1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU
// Each tab has its own Chaser Name column and a feedback/status column
// that contains the date embedded in the text, same as the daily sheets
// (e.g. "Approved+CN 5/1").
//
// USAGE (run once from the Apps Script editor):
//   backfillFromCombinedSheet()
// Resumable — saves progress after each date, safe to re-run if it times out.
// ============================================================

const COMBINED_HISTORICAL_SOURCE = {
  id: "1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU",
  tabs: [
    { name: "ORT Overall 2026",      feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", campaignKey: "ort"    },
    { name: "CGM Overall 2026",      feedbackCol: "FAX SENT ON EST", chaserCol: "Chaser",       campaignKey: "cgm"    },
    { name: "LY PUMP Overall 2026",  feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", campaignKey: "lymphc" },
    { name: "LY WRAP Overall 2026",  feedbackCol: "Fax Feedback",   chaserCol: "Chaser Name", campaignKey: "lymphw" },
  ]
};

// Scan every tab in the combined sheet and collect every unique "M/D" date
// found embedded in the feedback column text.
function collectDatesFromCombinedSheet() {
  const dates = new Set();
  let ss;
  try { ss = SpreadsheetApp.openById(COMBINED_HISTORICAL_SOURCE.id); }
  catch(e) { Logger.log("Cannot open combined sheet: " + e.message); return []; }

  for (const tabConfig of COMBINED_HISTORICAL_SOURCE.tabs) {
    const sheet = ss.getSheetByName(tabConfig.name);
    if (!sheet) { Logger.log("Tab not found: " + tabConfig.name); continue; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers     = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    if (feedbackCol < 0) { Logger.log("Feedback col not found in " + tabConfig.name + ": " + tabConfig.feedbackCol); continue; }

    for (let r = 1; r < data.length; r++) {
      const feedback = String(data[r][feedbackCol] || "").trim();
      if (!feedback) continue;
      const tokens = feedback.match(/\b(\d{1,2}\/\d{1,2})(?:\/\d{2,4})?\b/g) || [];
      tokens.forEach(tok => {
        const p = tok.split("/");
        dates.add(parseInt(p[0]) + "/" + parseInt(p[1]));
      });
    }
  }

  return [...dates];
}

// Read campaign totals (one count per lead row, chasers irrelevant) for ONE date
// from the combined historical sheet.
function readCombinedCampaignTotalsForDate(dateTab) {
  const camps = zeroCampaignTotals();

  let ss;
  try { ss = SpreadsheetApp.openById(COMBINED_HISTORICAL_SOURCE.id); }
  catch(e) { return camps; }

  for (const tabConfig of COMBINED_HISTORICAL_SOURCE.tabs) {
    const sheet = ss.getSheetByName(tabConfig.name);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers     = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    const chaserCol   = headers.indexOf(tabConfig.chaserCol.toUpperCase());
    if (feedbackCol < 0 || chaserCol < 0) continue;

    const key = tabConfig.campaignKey;

    for (let r = 1; r < data.length; r++) {
      const feedback = String(data[r][feedbackCol] || "").trim();
      if (!feedback || !feedbackMatchesDate(feedback, dateTab)) continue;

      const approval = isApproval(feedback);
      const denial   = isDenial(feedback);
      if (!approval && !denial) continue;

      if (approval) camps[key].approved++;
      if (denial)   camps[key].denied++;
    }
  }

  return camps;
}

// Read per-chaser per-campaign counts for ONE date from the combined historical sheet.
// Same shape as readChaserCampaignCountsForDate, used to populate ORT_Approved etc.
// columns in the monthly archive tabs.
function readCombinedChaserCampaignCountsForDate(dateTab) {
  const byChaser = {};

  let ss;
  try { ss = SpreadsheetApp.openById(COMBINED_HISTORICAL_SOURCE.id); }
  catch(e) { return byChaser; }

  for (const tabConfig of COMBINED_HISTORICAL_SOURCE.tabs) {
    const sheet = ss.getSheetByName(tabConfig.name);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers     = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    const chaserCol   = headers.indexOf(tabConfig.chaserCol.toUpperCase());
    if (feedbackCol < 0 || chaserCol < 0) continue;

    const campKey = tabConfig.campaignKey;

    for (let r = 1; r < data.length; r++) {
      const feedback   = String(data[r][feedbackCol] || "").trim();
      const chaserText = String(data[r][chaserCol]   || "").trim();
      if (!feedback || !chaserText || !feedbackMatchesDate(feedback, dateTab)) continue;

      const approval = isApproval(feedback);
      const denial   = isDenial(feedback);
      if (!approval && !denial) continue;

      const chasers = chaserText.split("/").map(x => normalizeChaserName(x.trim())).filter(Boolean);
      for (const name of chasers) {
        if (!byChaser[name]) {
          byChaser[name] = zeroCampaignTotals();
        }
        if (approval) byChaser[name][campKey].approved++;
        if (denial)   byChaser[name][campKey].denied++;
      }
    }
  }

  return byChaser;
}

// Write/overwrite Campaign Responses rows for ONE date using the combined sheet's data.
// Mirrors archiveCampaignResponses() but reads from COMBINED_HISTORICAL_SOURCE instead
// of RESPONSE_SOURCES.
function writeCombinedCampaignResponses(dateTab) {
  const sheet = getOrCreateCampaignResponsesTab();
  const data  = sheet.getDataRange().getValues();

  const toDelete = [];
  for (let r = 1; r < data.length; r++) {
    const rowDate = data[r][0] instanceof Date
      ? (data[r][0].getMonth()+1)+"/"+data[r][0].getDate()
      : String(data[r][0]).trim();
    const rParts = rowDate.split("/");
    const rShort = parseInt(rParts[0]) + "/" + parseInt(rParts[1]);
    const tParts = dateTab.split("/");
    const tShort = parseInt(tParts[0]) + "/" + parseInt(tParts[1]);
    if (rShort === tShort) toDelete.push(r + 1);
  }
  toDelete.reverse().forEach(rowNum => sheet.deleteRow(rowNum));

  const camps = readCombinedCampaignTotalsForDate(dateTab);

  let written = 0;
  for (const [key, label] of Object.entries(CAMPAIGN_LABELS)) {
    const c   = camps[key];
    const tot = c.approved + c.denied;
    const pct = tot ? parseFloat((c.approved / tot * 100).toFixed(1)) : 0;
    sheet.appendRow([dateTab, label, c.approved, c.denied, pct]);
    written++;
  }

  return written;
}

// Update per-chaser ORT/CGM/LymphC/LymphW columns in the monthly archive tab
// for ONE date using the combined sheet's data. Mirrors the inner loop of
// backfillCampaignColumns() but reads from COMBINED_HISTORICAL_SOURCE.
function writeCombinedChaserCampaignColumns(dateTab) {
  const ss = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);

  // Figure out which monthly tab this date belongs to (assume current year
  // unless dateTab already has one)
  const fullDate = dateTab.split("/").length >= 3 ? dateTab : dateTab + "/" + new Date().getFullYear();
  const tabName  = monthTabName(fullDate);
  const sheet    = ss.getSheetByName(tabName);
  if (!sheet) { Logger.log("Monthly tab not found for " + dateTab + " (" + tabName + ") — skipping, run a Sync for this date first."); return 0; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h).trim());
  const COL = {};
  headers.forEach((h, i) => { COL[h] = i; });
  const required = ["Date","Chaser","ORT_Approved","ORT_Denied","CGM_Approved","CGM_Denied","LymphC_Approved","LymphC_Denied","LymphW_Approved","LymphW_Denied"];
  if (!required.every(h => COL.hasOwnProperty(h))) { Logger.log("Monthly tab missing expected columns: " + tabName); return 0; }

  const chaserCamps = readCombinedChaserCampaignCountsForDate(dateTab);

  let rowsUpdated = 0;
  for (let r = 1; r < data.length; r++) {
    const rawDate   = data[r][COL.Date];
    const chaserVal = String(data[r][COL.Chaser]).trim();
    if (!chaserVal || chaserVal.toUpperCase().startsWith("TOTAL") || chaserVal.toUpperCase() === "WEEK") continue;

    let rowDateStr;
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      rowDateStr = (rawDate.getMonth()+1) + "/" + rawDate.getDate();
    } else {
      const s = String(rawDate).trim();
      if (!s || s.toUpperCase() === "WEEK") continue;
      const parts = s.split("/");
      rowDateStr = parseInt(parts[0]) + "/" + parseInt(parts[1]);
    }
    if (rowDateStr !== dateTab) continue;

    const normalized = normalizeChaserName(chaserVal);
    const cc = chaserCamps[normalized] || chaserCamps[chaserVal] || zeroCampaignTotals();

    sheet.getRange(r+1, COL.ORT_Approved+1, 1, 8).setValues([[
      cc.ort.approved,    cc.ort.denied,
      cc.cgm.approved,    cc.cgm.denied,
      cc.lymphc.approved, cc.lymphc.denied,
      cc.lymphw.approved, cc.lymphw.denied,
    ]]);
    rowsUpdated++;
  }

  return rowsUpdated;
}

// ============================================================
// MAIN ENTRY POINT — run this once from the Apps Script editor
// ============================================================
// 1. Finds every date mentioned in the combined sheet's feedback columns
// 2. For each date: writes Campaign Responses rows (deduplicated totals)
// 3. For each date: updates per-chaser ORT/CGM/LymphC/LymphW columns in
//    the matching monthly archive tab — but ONLY for chaser rows that
//    already exist there. If a date has no archive rows yet, run a
//    Sync for that date first (so Cases/Positive/TimeMins etc. exist),
//    then re-run this function.
// Resumable: progress is saved after each date.
function backfillFromCombinedSheet() {
  const PROGRESS_KEY = "combined_sheet_backfill";
  const props        = PropertiesService.getScriptProperties();
  const doneDates    = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const allDates = collectDatesFromCombinedSheet().filter(d => !doneDates.includes(d));
  Logger.log("Dates found in combined sheet: " + (allDates.length + doneDates.length) + " | Remaining: " + allDates.length);

  let totalCampaignRows = 0;
  let totalChaserRows   = 0;
  let skippedNoArchive  = [];

  for (const dateTab of allDates) {
    try {
      totalCampaignRows += writeCombinedCampaignResponses(dateTab);
      const updated = writeCombinedChaserCampaignColumns(dateTab);
      totalChaserRows += updated;
      if (updated === 0) skippedNoArchive.push(dateTab);
    } catch(e) {
      Logger.log("Error processing " + dateTab + ": " + e.message);
    }
    doneDates.push(dateTab);
    props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
  }

  Logger.log("=== DONE ===");
  Logger.log("Campaign Responses rows written: " + totalCampaignRows);
  Logger.log("Chaser rows updated: " + totalChaserRows);
  if (skippedNoArchive.length) {
    Logger.log("Dates with NO matching archive rows (Campaign Responses still written, but per-chaser columns skipped): " + skippedNoArchive.join(", "));
    Logger.log("Sync those dates first, then re-run backfillFromCombinedSheet() to fill in the per-chaser columns.");
  }
}

function resetCombinedSheetBackfill() {
  PropertiesService.getScriptProperties().deleteProperty("combined_sheet_backfill");
  Logger.log("Combined sheet backfill progress reset.");
}
