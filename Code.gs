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
  Alex:  "1byPJ-RjIQzA4IcwuMieDXHpVb3EyR5Q8kJBnAs-npCU",
  Hope:  "1LGKRvxveag0hdiVSuiPNWgbd_o6RDLDlYuWDMFNrlb0",
  Rose:  "1pk4UmN6sH4qZVnLo3L1UphMaIOwUlGydDkHwOS9smzY",
  Frank: "1-CuYnkkj8w9KO5RSjt6tyQ4l9xo4Pv5zZT_Gg1sfddA",
  Nova:  "1_KrQtNWg3L-QMv1CedqT_qfNPZ215Bv6Z31nWqbg0D0"
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
// RESPONSE SOURCES -- retired. Campaign approvals/denials now come
// exclusively from BACKFILL_RESPONSES_SHEET_ID/BACKFILL_RESPONSES_TABS via
// the weekly pullWeeklyCampaignData() pull, not from these live day-to-day
// fax sheets -- see updateChaserCampaignColumnsForDate() and
// writeCombinedCampaignResponses() below.
// ============================================================

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
const DENIAL_PATTERNS   = ["RECEIVED DENIAL"];

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

// Maps "date|chaser" -> { rowNum, values } for a specific month tab, so
// archiveDayData() can update an existing row in place (fresh Cases/
// Positive/etc. from a re-sync) instead of either skipping it (leaving
// stale numbers forever) or blindly appending a duplicate row.
function buildExistingRowMap(sheet) {
  const map  = new Map();
  const data = sheet.getDataRange().getValues();
  for (let r = 1; r < data.length; r++) {
    const key = normalizeDateCellToTab(data[r][0]) + "|" + String(data[r][1]);
    map.set(key, { rowNum: r + 1, values: data[r] });
  }
  return map;
}

// ============================================================
// ARCHIVE A SINGLE DAY  (called by EOD trigger + manual runs)
// ============================================================
// Utlatel duration/calls for a given date, factored out of archiveDayData()
// so backfillCampaignColumns() can reuse it when inserting a row for a
// chaser+date combo that never had one (e.g. missing tracker tab).
function buildUtlatelLookup(dateTab) {
  const utlatelForDate = getUtlatelData().filter(r => r.Date === dateTab);
  return function utlatelTotalsForChaser(chaserName) {
    let mins = 0, calls = 0;
    utlatelForDate.forEach(r => {
      if (String(r.Agent || "").toLowerCase().includes(chaserName.toLowerCase())) {
        mins  += Number(r.DurationMins) || 0;
        calls += Number(r.Calls)        || 0;
      }
    });
    return { mins, calls };
  };
}

// Shift minutes (TotalShift) and ACW multiplier for a given date, mirroring
// the dashboard's applyShiftHistory()/formulaSettings.acwMult exactly, so
// Sync/EOD (and the campaign-columns backfill, which reuses this) write the
// same numbers the dashboard would otherwise only compute live:
//   - TotalShift: most recent per-chaser Settings-tab entry with
//     EffectiveDate <= the date being archived, else DEFAULT_SHIFT_MINS.
//   - ACW multiplier: the "__formula__"/"acwMult" global override (last
//     one wins, same as the dashboard's own load-time reducer), else 2.
// Factored out of archiveDayData() so backfillCampaignColumns() can reuse it.
function buildShiftAndAcwContext(dateTab) {
  const DEFAULT_SHIFT_MINS = 400;
  const rawShiftHistory    = getSettings().shiftHistory || [];
  const archiveTargetDate  = parseDateTab(dateTab);

  let acwMult = 2;
  rawShiftHistory.forEach(s => {
    if (s.Chaser === "__formula__" && s.ShiftType === "acwMult") {
      const v = parseFloat(s.Notes);
      if (!isNaN(v)) acwMult = v;
    }
  });

  function parseMDY(s) {
    const p = String(s || "").trim().split("/");
    if (p.length < 3) return null;
    const m = parseInt(p[0], 10), d = parseInt(p[1], 10), y = parseInt(p[2], 10);
    return (isNaN(m) || isNaN(d) || isNaN(y)) ? null : new Date(y, m - 1, d);
  }

  function shiftMinsForChaser(chaserName) {
    const entries = rawShiftHistory
      .filter(s => s.Chaser === chaserName)
      .map(s => ({ mins: parseFloat(s.ShiftType), date: parseMDY(s.EffectiveDate) }))
      .filter(s => s.date && s.date <= archiveTargetDate && !isNaN(s.mins))
      .sort((a, b) => b.date - a.date);
    return entries.length ? entries[0].mins : DEFAULT_SHIFT_MINS;
  }

  return { acwMult, shiftMinsForChaser };
}

// archiveDayData() only ever touches Cases/Positive/TimeMins/Efficiency/
// Productivity/TotalShift/TotalCalls/TotalDurationMins/ACWDuration/
// ProductiveTime (indices 2,3,6-13) -- everything derived from tracker tabs
// + Utlatel. Approvals/Denials/campaign columns (4,5,14-21) belong to the
// weekly campaign pull (see updateChaserCampaignColumnsForDate() below) and
// are deliberately left untouched here: when updating an existing row,
// whatever the weekly pull already wrote into those columns is preserved
// as-is rather than being zeroed out by every daily re-run.
function archiveDayData(dateTab) {
  dateTab = dateTab || getTodayTab();
  const data = getDayData(dateTab);

  const ss       = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const tabName  = monthTabName(dateTab);
  const sheet    = getOrCreateMonthTab(ss, tabName);
  const existing = buildExistingRowMap(sheet);

  // Utlatel rows for this date, so TotalCalls/TotalDurationMins get written
  // into the archive row itself instead of staying permanently blank and
  // relying solely on the dashboard's client-side join at render time (the
  // dashboard's own "Data source priority" doc already treats this archive
  // column as a valid fallback -- it just never got populated until now).
  const utlatelTotalsForChaser = buildUtlatelLookup(dateTab);

  const { acwMult, shiftMinsForChaser } = buildShiftAndAcwContext(dateTab);

  // Check if this date already has rows (skip week header if so)
  const alreadyHasRows = [...existing.keys()].some(k => k.startsWith(dateTab + "|"));
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
  let updated = 0;
  const missingTabChasers = [];
  const readErrorChasers  = [];
  for (const c of data.chasers) {
    const key = dateTab + "|" + c.name;
    const existingRow = existing.get(key);

    // If this chaser's tracker has no tab named dateTab + "." for this date, do NOT
    // write Cases/Positive/TimeMins for them — a chaser tracker can have a leftover
    // tab from last year with the same "M/D" name but no trailing dot; readChaserTab()
    // only ever looks up the dotted (current-year) tab and never falls back to it, so
    // tabFound=false here genuinely means "no data yet", not "check the other tab".
    // Writing zeros for those fields would be indistinguishable from a real
    // zero-case day and would corrupt the archive permanently. A chaser with real
    // campaign credit but no tracker tab is handled by the weekly pull instead,
    // which inserts a tracker-columns-blank row for exactly that case.
    //
    // readChaserTab() also sets tabFound=false when openById/getSheetByName threw
    // (permissions, bad sheet ID, transient API error) — that is NOT "no data yet",
    // so it's tracked and logged separately instead of being lumped in with a
    // genuinely missing tab, which would send troubleshooting in the wrong direction.
    if (!c.tabFound) {
      if (c.error) {
        readErrorChasers.push(c.name + " (" + c.error + ")");
      } else {
        missingTabChasers.push(c.name);
      }
      continue;
    }

    const eff = c.totalCases ? (c.totalPositive / c.totalCases * 100).toFixed(1) : "";

    const utl       = utlatelTotalsForChaser(c.name);
    const shiftMins = shiftMinsForChaser(c.name);

    // ProductiveTime/ACWDuration/Productivity only mean something once we
    // have real call-duration data for this chaser+date -- mirrors the
    // dashboard's own "productivityStub" gate (buildRow: !utlMins), which
    // shows an "Upload Utlatel" stub rather than asserting a number derived
    // from zero duration.
    let acwDuration = "", productiveTime = "", productivity = "";
    if (utl.mins > 0) {
      acwDuration    = acwMult * utl.calls;
      productiveTime = utl.mins + acwDuration;
      productivity   = shiftMins > 0 ? (productiveTime / shiftMins * 100).toFixed(1) : "";
    }

    // Approvals(4)/Denials(5)/campaign columns(14-21) are owned by the weekly
    // campaign pull, never by this daily write. Start from the existing row's
    // values (so those columns are preserved untouched) if there is one, else
    // default the whole row to blank/0 for a brand-new date+chaser combo.
    const rowValues = existingRow
      ? existingRow.values.slice()
      : [dateTab, c.name, 0, 0, 0, 0, 0, "", "", "", "", "", "", "", 0,0,0,0,0,0,0,0];

    rowValues[0] = dateTab;
    rowValues[1] = c.name;
    rowValues[2] = c.totalCases;
    rowValues[3] = c.totalPositive;
    rowValues[6] = c.totalTimeMins;
    rowValues[7] = eff;
    rowValues[8] = productivity;
    rowValues[9] = shiftMins;
    rowValues[10] = utl.calls || "";
    rowValues[11] = utl.mins  || "";
    rowValues[12] = acwDuration;
    rowValues[13] = productiveTime;

    if (existingRow) {
      sheet.getRange(existingRow.rowNum, 1, 1, ARCHIVE_HEADERS.length).setValues([rowValues]);
      updated++;
    } else {
      sheet.appendRow(rowValues);
      existing.set(key, { rowNum: sheet.getLastRow(), values: rowValues });
      written++;
    }
  }

  if (missingTabChasers.length) {
    Logger.log("*** ARCHIVE INCOMPLETE for " + dateTab + " *** no tab named \"" + dateTab +
      ".\" found for: " + missingTabChasers.join(", ") +
      " — their tracker tab for today may not exist yet. Cases/Positive/TimeMins were " +
      "NOT recorded for them. Re-run the daily archive for " + dateTab + " once the tab exists.");
  }
  if (readErrorChasers.length) {
    Logger.log("*** ARCHIVE INCOMPLETE for " + dateTab + " *** could not read tracker sheet for: " +
      readErrorChasers.join(", ") + " — this is NOT a missing tab, the sheet read itself " +
      "threw an error (see parenthetical above). Cases/Positive/TimeMins were NOT recorded " +
      "for them. Re-run the daily archive for " + dateTab + " once the underlying error is fixed.");
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

  // After writing all chasers for a Friday, also add a blank separator row --
  // gated on written>0 so re-syncing an already-archived Friday (now updating
  // rows in place instead of skipping) doesn't append another separator
  // every time.
  if (isFriday(dateTab) && written > 0) {
    sheet.appendRow(new Array(ARCHIVE_HEADERS.length).fill(""));
  }

  Logger.log("Archived " + dateTab + " → " + tabName + " | Written: " + written + ", Updated: " + updated + " rows");
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
// EOD TRIGGER — retired. Daily archiving is now a manual button
// (mode=sync, see doGet) instead of an automatic time-based trigger.
// Run removeEodTrigger() once from the Apps Script editor after deploying
// this to delete whatever eodArchive trigger is currently installed live.
// ============================================================
function removeEodTrigger() {
  const removed = ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "eodArchive");
  removed.forEach(t => ScriptApp.deleteTrigger(t));
  Logger.log("Removed " + removed.length + " eodArchive trigger(s).");
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



// Force the given column names to Plain Text formatting on the whole column
// so Google Sheets never silently auto-converts a written date string (e.g.
// "6/20/2026") into a real Date value with a guessed year. Every getOrCreate*Tab
// call that writes dates runs this unconditionally, not just on first creation,
// since it must hold for as long as the tab exists.
function forcePlainTextColumns(sheet, headers, fieldNames) {
  fieldNames.forEach(function(field) {
    const idx = headers.indexOf(field);
    if (idx === -1) return;
    let letter = "", n = idx + 1;
    while (n > 0) {
      const rem = (n - 1) % 26;
      letter = String.fromCharCode(65 + rem) + letter;
      n = Math.floor((n - 1) / 26);
    }
    sheet.getRange(letter + ":" + letter).setNumberFormat("@");
  });
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

const UTLATEL_HEADERS = ["Date","Agent","DurationMins","Calls"];

function getOrCreateUtlatelTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Utlatel");
  if (!sheet) {
    sheet = ss.insertSheet("Utlatel");
    sheet.appendRow(UTLATEL_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,4).setFontWeight("bold").setBackground("#1A2C42").setFontColor("#00C2A8");
    Logger.log("Created Utlatel tab");
  }
  // Without this, Sheets auto-coerces a "6/25"-style Date string into a real
  // Date cell defaulted to some year -- getUtlatelData() then reads it back
  // as "6/25/<that year>", which never matches the year-less "M/D" dateTab
  // keys (getTodayTab/pickerToTab) the dashboard joins Utlatel data against,
  // so uploaded duration/calls silently fail to show up for that date.
  forcePlainTextColumns(sheet, UTLATEL_HEADERS, ["Date"]);
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
    // Normalize date to "M/D" (no year) -- matches the year-less dateTab
    // keys (getTodayTab/pickerToTab) the dashboard joins Utlatel data
    // against; a row that slipped through as a real Date cell before
    // forcePlainTextColumns was applied must still come back in this format.
    if (obj.Date instanceof Date && !isNaN(obj.Date)) {
      const d = obj.Date;
      obj.Date = (d.getMonth()+1) + "/" + d.getDate();
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

  // Build set of existing date|agent combos. Dates are compared as "M/D"
  // (no year) -- same reasoning as getUtlatelData() -- so a row already
  // sitting in the sheet as a real Date cell still matches this and future
  // uploads' year-less date strings instead of silently duplicating.
  const existingKeys = new Set();
  for (let r = 1; r < existing.length; r++) {
    const d = existing[r][0] instanceof Date
      ? (existing[r][0].getMonth()+1)+"/"+existing[r][0].getDate()
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
          ? (existing[r][0].getMonth()+1)+"/"+existing[r][0].getDate()
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
      // The manual "daily archive" button -- reads chaser tracker tabs for
      // the given date, overwrites that date's Cases/Positive/TimeMins/
      // Shift/Productivity archive columns, then clears cache. Replaces the
      // old automatic EOD trigger (see removeEodTrigger()) -- campaign
      // approvals/denials are no longer touched here at all; those come
      // exclusively from the weekly campaign pull (mode=weeklycampaignsync).
      const syncDate = date;

      // Delete existing rows for today so we can rewrite them
      deleteArchiveRowsForDate(syncDate);

      // Re-archive from live tracker sheets
      archiveDayData(syncDate);

      // Clear archive cache so next dashboard load gets fresh data
      CacheService.getScriptCache().remove("archive_all");

      payload = { success: true, synced: syncDate, message: "Archived " + syncDate + " from trackers" };

    } else if (mode === "weeklycampaignsync") {
      // The manual "weekly response pull" button -- reprocesses the
      // trailing WEEKLY_PULL_DAYS days against the backfill source for
      // every chaser, updating campaign columns + the Campaign Responses
      // tab. See pullWeeklyCampaignData().
      payload = pullWeeklyCampaignData();
      CacheService.getScriptCache().remove("archive_all");
      CacheService.getScriptCache().remove("campaign_responses");

    } else if (mode === "campaignresponses") {
      payload = getCachedOrFetch("campaign_responses", getArchiveCampaignResponses, 300);

    } else if (mode === "chasers") {
      // Full roster (including inactive) for the dashboard's Settings tab.
      payload = { chasers: getChasersConfig() };

    } else if (mode === "savechaser") {
      // Add a new chaser or update an existing one's Sheet ID / Active flag.
      // Params: name, sheetId, active ("true"/"false")
      payload = saveChaser(params);

    } else if (mode === "debug") {
      // Diagnostic view of exactly what a Sync would see for one date —
      // see getDebugInfoForDate() for what this exposes and why it exists
      // (no Execution Log access needed; visible via the Network tab).
      payload = getDebugInfoForDate(date);

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

  // Campaign approvals/denials no longer come from a live day-to-day read --
  // see updateChaserCampaignColumnsForDate()/pullWeeklyCampaignData() below.
  // Kept as a zeroed shape (not removed) so legacy mode=day/week callers
  // don't break on a missing field.
  const responses = { approved: 0, denied: 0, byChaser: {}, byCampaign: zeroCampaignTotals(true) };

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

// Chaser tracker tabs are named "M/D/YY" or "M/D/YYYY" (e.g. "6/25/26" or
// "6/25/2026"), with month/day either zero-padded or not depending on the
// chaser's own sheet -- tries every combination so it doesn't matter which
// one a given sheet actually uses. dateTab here is always a live/current
// date (this function is only ever used for daily archiving, never a
// stale historical backfill), so a missing year defaults to the current
// year, same reasoning as parseDateTab() elsewhere. Falls back to the old
// "M/D." dotted format last, for any tab not yet renamed.
function findChaserTrackerSheet(ss, dateTab) {
  const parts = String(dateTab).split("/");
  const m     = parseInt(parts[0], 10), d = parseInt(parts[1], 10);
  const year  = parts[2] ? parseInt(parts[2], 10) : new Date().getFullYear();
  const yy    = String(year).slice(-2);
  const mm    = String(m).padStart(2, "0");
  const dd    = String(d).padStart(2, "0");

  const candidates = [
    m + "/" + d + "/" + year,
    m + "/" + d + "/" + yy,
    mm + "/" + dd + "/" + year,
    mm + "/" + dd + "/" + yy,
    m + "/" + d + ".", // legacy dotted format, for any tab not yet renamed
  ];
  for (const name of candidates) {
    const sheet = ss.getSheetByName(name);
    if (sheet) return sheet;
  }
  return null;
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
    const sheet = findChaserTrackerSheet(ss, dateTab);

    if (!sheet) {
      Logger.log(chaserName + " — tab not found for: " + dateTab);
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
// DEBUG — surfaces exactly what the daily archive button and the weekly
// campaign pull would see for one date, without needing Apps Script
// Execution Log access (Editor > Executions). Hit ?mode=debug&date=M/D
// directly in the browser (or via the Network tab like the other modes)
// to inspect: whether each chaser's tracker tab was found (and the real
// error if openById/getSheetByName threw, instead of that being silently
// reported as "tab not found"); which chasers already have an archive row
// for this date; and, per BACKFILL_RESPONSES_TABS tab (the weekly
// campaign-data source), every row the code currently resolves to this
// date, regardless of whether it passed the approval/denial text check.
// ============================================================
function getDebugInfoForDate(dateTab) {
  const chaserSheets = getActiveChaserSheetMap();
  const chasers = Object.entries(chaserSheets).map(([name, sheetId]) => {
    const r = readChaserTab(sheetId, dateTab, name);
    return {
      name, sheetId, tabFound: r.tabFound, error: r.error || null,
      totalCases: r.totalCases, totalPositive: r.totalPositive, totalTimeMins: r.totalTimeMins
    };
  });

  // Chasers who already have a row for this date — archiveDayData skips
  // these as duplicates on the next Sync instead of refreshing them.
  const existingArchiveChasers = [];
  try {
    const ss      = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
    const tabName = monthTabName(dateTab);
    const sheet   = ss.getSheetByName(tabName);
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let r = 1; r < data.length; r++) {
        if (normalizeDateCellToTab(data[r][0]) === dateTab) {
          existingArchiveChasers.push(String(data[r][1]));
        }
      }
    }
  } catch (err) { /* leave existingArchiveChasers empty */ }

  const fullDateTab = dateTabToFullDate(dateTab);
  const responseSources = [];
  let ss;
  try { ss = SpreadsheetApp.openById(BACKFILL_RESPONSES_SHEET_ID); }
  catch (err) {
    for (const tabConfig of BACKFILL_RESPONSES_TABS) {
      responseSources.push({ tab: tabConfig.tabName, error: "Could not open spreadsheet: " + err.message });
    }
  }

  if (ss) {
    for (const tabConfig of BACKFILL_RESPONSES_TABS) {
      const sheet = ss.getSheetByName(tabConfig.tabName);
      if (!sheet) { responseSources.push({ tab: tabConfig.tabName, error: "Tab not found" }); continue; }

      const data        = sheet.getDataRange().getValues();
      const headers     = data[0] ? data[0].map(h => String(h).trim().toUpperCase()) : [];
      const feedbackCol   = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
      const chaserCol     = headers.indexOf(tabConfig.chaserCol.toUpperCase());
      const conclusionCol = tabConfig.conclusionCol ? headers.indexOf(tabConfig.conclusionCol.toUpperCase()) : -1;
      const tabYear       = resolveYearFromTabName(tabConfig.tabName);

      const entry = {
        tab: tabConfig.tabName,
        headersRaw: data[0] || [],
        feedbackColConfigured: tabConfig.feedbackCol,
        feedbackColFound: feedbackCol >= 0,
        chaserColConfigured: tabConfig.chaserCol,
        chaserColFound: chaserCol >= 0,
        conclusionColConfigured: tabConfig.conclusionCol,
        conclusionColFound: conclusionCol >= 0,
        totalDataRows: Math.max(0, data.length - 1),
        matchingDateRows: []
      };

      if (feedbackCol >= 0 && chaserCol >= 0) {
        for (let r = 1; r < data.length; r++) {
          const feedback = String(data[r][feedbackCol] || "").trim();
          const chaser   = String(data[r][chaserCol]   || "").trim();
          if (!feedback) continue;
          const resolved = resolveRowDate(feedback, conclusionCol >= 0 ? data[r][conclusionCol] : null, tabYear);
          if (resolved !== fullDateTab) continue;
          entry.matchingDateRows.push({
            row: r + 1, feedback, chaser,
            isApproval: isApproval(feedback), isDenial: isDenial(feedback)
          });
        }
      }

      responseSources.push(entry);
    }
  }

  return { date: dateTab, chasers, existingArchiveChasers, responseSources };
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
      // (the weekly campaign pull fills these going forward)
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
// Reads BACKFILL_RESPONSES_SHEET_ID/BACKFILL_RESPONSES_TABS below (a
// combined "Overall 2026" history spreadsheet) and:
//   - updates ORT/CGM/LymphC/LymphW approved/denied columns on existing rows
//   - INSERTS a new row (Cases/Positive/TimeMins at 0) for a chaser+date that
//     has real campaign credit but no existing archive row at all -- e.g. no
//     tracker tab existed for that chaser that day, so nothing was ever
//     written for them.
//
// This is now the ONLY source of campaign approvals/denials in the whole
// app -- the daily archive write (archiveDayData()) never touches these
// columns at all (see updateChaserCampaignColumnsForDate() below, shared
// by this one-time historical backfill AND the manual weekly pull,
// pullWeeklyCampaignData()).
//
// HOW TO USE (one-time historical catch-up only -- for ongoing data use
// the "Pull Weekly Responses" dashboard button instead):
//   Run backfillCampaignColumns() once.
//   It processes one month tab at a time — if it times out,
//   just run it again (already-processed dates are skipped).
//
// PROGRESS KEY: "campaign_backfill_progress5"
// ============================================================

// This spreadsheet holds a fuller combined historical record than any live
// day-to-day sheet ever did.
const BACKFILL_RESPONSES_SHEET_ID = "1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU";

// conclusionCol ("Date of conclusion") is only present on the ORT/CGM tabs
// -- it's the actual resolution date (with a real year), authoritative
// over anything parsed out of the feedback text. LY PUMP/LY WRAP don't
// have it, but they've also never carried anything but current-year data,
// so falling back to the tab's own year there is safe.
const BACKFILL_RESPONSES_TABS = [
  { tabName: "ORT Overall 2026",     campaignKey: "ort",    idCol: "MBI",                feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: "IDN", conclusionCol: "Date of conclusion" },
  { tabName: "CGM Overall 2026",     campaignKey: "cgm",    idCol: "MBI",                feedbackCol: "FAX SENT ON EST", chaserCol: "Chaser",      submissionCol: "Submission Date", idnCol: "IDN", conclusionCol: "Date of conclusion" },
  { tabName: "LY PUMP Overall 2026", campaignKey: "lymphc", idCol: "Insurance ID Number", feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: null,  conclusionCol: null },
  { tabName: "LY WRAP Overall 2026", campaignKey: "lymphw", idCol: "MBI",                feedbackCol: "Fax Feedback",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: null,  conclusionCol: null },
];

// Reads BACKFILL_RESPONSES_TABS for one date and returns per-chaser totals
// in the same shape the live path produces (readResponsesForDate +
// readChaserCampaignCountsForDate combined into one pass): overall
// approvals/denials summed across all 4 campaigns, plus a per-campaign
// breakdown -- so a chaser+date combo missing from the archive entirely can
// be inserted with real data, not just have an existing row updated.
function readChaserTotalsFromBackfillSource(dateTab) {
  const result = {};
  let ss;
  try {
    ss = SpreadsheetApp.openById(BACKFILL_RESPONSES_SHEET_ID);
  } catch (err) {
    Logger.log("Could not open backfill responses spreadsheet: " + err.message);
    return result;
  }

  for (const tabConfig of BACKFILL_RESPONSES_TABS) {
    const sheet = ss.getSheetByName(tabConfig.tabName);
    if (!sheet) { Logger.log("Backfill tab not found: " + tabConfig.tabName); continue; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers       = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol   = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    const chaserCol     = headers.indexOf(tabConfig.chaserCol.toUpperCase());
    const conclusionCol = tabConfig.conclusionCol ? headers.indexOf(tabConfig.conclusionCol.toUpperCase()) : -1;
    if (feedbackCol < 0) { Logger.log("Feedback column not found: " + tabConfig.feedbackCol + " in " + tabConfig.tabName); continue; }
    if (chaserCol   < 0) { Logger.log("Chaser column not found: "   + tabConfig.chaserCol   + " in " + tabConfig.tabName); continue; }

    const campaignKey = tabConfig.campaignKey;
    const tabYear      = resolveYearFromTabName(tabConfig.tabName);

    for (let r = 1; r < data.length; r++) {
      const feedback   = String(data[r][feedbackCol] || "").trim();
      const chaserText = String(data[r][chaserCol]   || "").trim();
      if (!feedback || !chaserText) continue;
      // ORT/CGM's "Date of conclusion" is authoritative when present (see
      // resolveRowDate); LY PUMP/LY WRAP fall back to the feedback text
      // with this tab's own year as default -- same reasoning as the
      // combined-sheet backfill functions above.
      const fullDate = resolveRowDate(feedback, conclusionCol >= 0 ? data[r][conclusionCol] : null, tabYear);
      if (fullDate !== dateTab) continue;

      const approval = isApproval(feedback);
      const denial   = isDenial(feedback);
      if (!approval && !denial) continue;

      const chasers = chaserText.split("/").map(x => normalizeChaserName(x.trim())).filter(Boolean);
      for (const chaser of chasers) {
        if (!result[chaser]) {
          result[chaser] = { approvals: 0, denials: 0, campaigns: zeroCampaignTotals() };
        }
        if (approval) result[chaser].approvals++;
        if (denial)   result[chaser].denials++;
        if (result[chaser].campaigns[campaignKey]) {
          if (approval) result[chaser].campaigns[campaignKey].approved++;
          if (denial)   result[chaser].campaigns[campaignKey].denied++;
        }
      }
    }
  }

  return result;
}

// Updates one date's archive rows with per-chaser campaign data (Approvals,
// Denials, and the 8 ORT/CGM/LymphC/LymphW columns) from the backfill
// source, and inserts a new row (tracker columns blank/0) for any chaser
// with real credit on this date but no existing row (no tracker tab that
// day). Only the campaign-owned columns (4,5,14-21) are ever touched on an
// existing row -- everything else (owned by the daily archive write) is
// read back unchanged from the row itself. Returns { updated, inserted }.
// Shared by backfillCampaignColumns() (full history, progress-tracked) and
// pullWeeklyCampaignData() (trailing window, no progress tracking).
function updateChaserCampaignColumnsForDate(fullDate) {
  const ss      = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const tabName = monthTabName(fullDate);
  const sheet   = getOrCreateMonthTab(ss, tabName);

  let chaserTotals;
  try {
    chaserTotals = readChaserTotalsFromBackfillSource(fullDate);
  } catch (err) {
    Logger.log("Error reading backfill source for " + fullDate + ": " + err.message);
    return { updated: 0, inserted: 0 };
  }

  const data = sheet.getDataRange().getValues();

  // The year ALWAYS comes from the tab's own name (e.g. "Jun 2026"), never
  // from the cell itself -- getOrCreateMonthTab() never forces its Date
  // column to plain text, so a bare "6/25" can get auto-coerced by Sheets
  // into a Date object with Sheets' own year guess.
  const tabYearMatch = tabName.match(/(\d{4})\s*$/);
  const tabYear      = tabYearMatch ? parseInt(tabYearMatch[1], 10) : new Date().getFullYear();

  function rowDate(rawDate) {
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      return (rawDate.getMonth()+1) + "/" + rawDate.getDate() + "/" + tabYear;
    }
    const s = String(rawDate || "").trim();
    if (!s || s.toUpperCase() === "WEEK") return null;
    const parts = s.split("/");
    const m = parseInt(parts[0], 10), d = parseInt(parts[1], 10);
    if (isNaN(m) || isNaN(d)) return null;
    return m + "/" + d + "/" + tabYear;
  }

  let rowsUpdated = 0;
  const matchedChasers = new Set();

  for (let r = 1; r < data.length; r++) {
    const chaserVal = String(data[r][1]).trim();
    if (!chaserVal || chaserVal.toUpperCase().startsWith("TOTAL")) continue;

    const rowFullDate = rowDate(data[r][0]);
    if (rowFullDate !== fullDate) continue;

    // Archive stores normalized short names; chaserTotals also uses normalized names
    const normalized = normalizeChaserName(chaserVal);
    matchedChasers.add(normalized);
    const totals = chaserTotals[normalized] || chaserTotals[chaserVal];
    const cc     = totals ? totals.campaigns : zeroCampaignTotals();

    // Patch only the campaign-owned columns into the row's existing values
    // -- Cases/Positive/TimeMins/etc (owned by the daily archive write)
    // pass through untouched.
    const rowValues = data[r].slice();
    rowValues[4] = totals ? totals.approvals : 0;
    rowValues[5] = totals ? totals.denials   : 0;
    rowValues[14] = cc.ort.approved;    rowValues[15] = cc.ort.denied;
    rowValues[16] = cc.cgm.approved;    rowValues[17] = cc.cgm.denied;
    rowValues[18] = cc.lymphc.approved; rowValues[19] = cc.lymphc.denied;
    rowValues[20] = cc.lymphw.approved; rowValues[21] = cc.lymphw.denied;
    sheet.getRange(r+1, 1, 1, ARCHIVE_HEADERS.length).setValues([rowValues]);
    rowsUpdated++;
  }

  // Insert a row for any chaser with real credit on this date that didn't
  // already have an archive row -- no tracker tab existed for them that day.
  let rowsInserted = 0;
  const { acwMult, shiftMinsForChaser } = buildShiftAndAcwContext(fullDate);
  const utlatelTotalsForChaser = buildUtlatelLookup(fullDate);

  for (const [chaserName, totals] of Object.entries(chaserTotals)) {
    if (matchedChasers.has(chaserName)) continue; // already has a row, handled above
    const hasCredit = totals.approvals || totals.denials ||
      Object.values(totals.campaigns).some(c => (c.approved||0) > 0 || (c.denied||0) > 0);
    if (!hasCredit) continue;

    const utl       = utlatelTotalsForChaser(chaserName);
    const shiftMins = shiftMinsForChaser(chaserName);
    let acwDuration = "", productiveTime = "", productivity = "";
    if (utl.mins > 0) {
      acwDuration    = acwMult * utl.calls;
      productiveTime = utl.mins + acwDuration;
      productivity   = shiftMins > 0 ? (productiveTime / shiftMins * 100).toFixed(1) : "";
    }

    const cc = totals.campaigns;
    sheet.appendRow([
      fullDate, chaserName, 0, 0,
      totals.approvals, totals.denials, 0, "",  // Cases,Positive,TimeMins,Efficiency
      productivity, shiftMins, utl.calls || "", utl.mins || "",
      acwDuration, productiveTime,
      cc.ort.approved,    cc.ort.denied,
      cc.cgm.approved,    cc.cgm.denied,
      cc.lymphc.approved, cc.lymphc.denied,
      cc.lymphw.approved, cc.lymphw.denied,
    ]);
    rowsInserted++;
    Logger.log("Inserted missing-tracker row for " + chaserName + " on " + fullDate +
      " (Approvals=" + totals.approvals + ", Denials=" + totals.denials + ")");
  }

  return { updated: rowsUpdated, inserted: rowsInserted };
}

function backfillCampaignColumns() {
  const PROGRESS_KEY  = "campaign_backfill_progress5";
  const props         = PropertiesService.getScriptProperties();
  const doneDates     = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const ss           = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  const monthPattern  = /^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{4})$/i;

  let totalUpdated  = 0;
  let totalInserted = 0;
  let totalSkipped  = 0;

  // rowDate: given a raw cell value, returns "M/D/YYYY" -- the year ALWAYS
  // comes from the tab's own name (e.g. "Jun 2026"), never from the cell
  // itself. getOrCreateMonthTab() never forces its Date column to plain
  // text either, so a bare "6/25" can get auto-coerced by Sheets into a
  // Date object with Sheets' own year guess -- trusting that guessed year
  // here would reintroduce the exact bug this fix exists to avoid. Only
  // month/day are ever read off the cell; the tab name is the one thing
  // that's never ambiguous.
  function rowDate(rawDate, tabYear) {
    if (rawDate instanceof Date && !isNaN(rawDate)) {
      return (rawDate.getMonth()+1) + "/" + rawDate.getDate() + "/" + tabYear;
    }
    const s = String(rawDate || "").trim();
    if (!s || s.toUpperCase() === "WEEK") return null;
    const parts = s.split("/");
    const m = parseInt(parts[0], 10), d = parseInt(parts[1], 10);
    if (isNaN(m) || isNaN(d)) return null;
    return m + "/" + d + "/" + tabYear;
  }

  for (const sheet of ss.getSheets()) {
    const tabName  = sheet.getName().trim();
    const tabMatch = tabName.match(monthPattern);
    if (!tabMatch) continue;
    const tabYear = parseInt(tabMatch[2], 10);

    Logger.log("Processing tab: " + tabName);
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    // Collect unique dates in this tab that haven't been processed. Note:
    // this only discovers dates that already have AT LEAST ONE chaser's row
    // in the archive -- a date where every single chaser had no tracker tab
    // (so archiveDayData() never wrote anything at all for that day) won't
    // be found here. That's a narrower, more extreme edge case than "one
    // chaser's credit went missing while others were archived normally."
    const datesToProcess = new Set();
    for (let r = 1; r < data.length; r++) {
      const chaserVal = String(data[r][1]).trim();
      if (!chaserVal || chaserVal.toUpperCase().startsWith("TOTAL")) continue;

      const fullDate = rowDate(data[r][0], tabYear);
      if (!fullDate) continue;

      if (!doneDates.includes(fullDate)) {
        datesToProcess.add(fullDate);
      } else {
        totalSkipped++;
      }
    }

    Logger.log("  Dates to process: " + datesToProcess.size + " | Already done: " + totalSkipped);

    for (const fullDate of datesToProcess) {
      Logger.log("  Fetching per-chaser campaign counts for: " + fullDate);
      const { updated, inserted } = updateChaserCampaignColumnsForDate(fullDate);
      Logger.log("  Updated " + updated + " rows, inserted " + inserted + " rows for " + fullDate);
      totalUpdated  += updated;
      totalInserted += inserted;

      doneDates.push(fullDate);
      props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
    }
  }

  Logger.log("=== DONE ===");
  Logger.log("Rows updated: " + totalUpdated);
  Logger.log("Rows inserted (missing tracker tab, real campaign credit): " + totalInserted);
  Logger.log("Dates skipped (already done): " + totalSkipped);
  Logger.log("Total dates processed so far: " + doneDates.length);
  Logger.log("Run backfillCampaignColumns() again if there are more dates to process.");
}

// Reset campaign backfill progress
function resetCampaignBackfill() {
  PropertiesService.getScriptProperties().deleteProperty("campaign_backfill_progress5");
  Logger.log("Campaign backfill progress reset.");
}

// Check progress
function checkCampaignBackfillProgress() {
  const done = JSON.parse(
    PropertiesService.getScriptProperties().getProperty("campaign_backfill_progress5") || "[]"
  );
  Logger.log("Dates with campaign data backfilled: " + done.length);
}

// ============================================================
// WEEKLY CAMPAIGN DATA PULL — manual button, replaces the old daily
// RESPONSE_SOURCES-based reads entirely. Every click reprocesses the
// trailing WEEKLY_PULL_DAYS calendar days against the backfill source
// (BACKFILL_RESPONSES_SHEET_ID/TABS) -- both the per-chaser campaign
// columns in the monthly archive tabs and the aggregated Campaign
// Responses tab. No progress tracking: the window is small enough to
// finish in one execution, and reprocessing it every time means a date
// whose data arrived late in the backfill sheet self-corrects on the next
// click instead of staying wrong forever.
// ============================================================
const WEEKLY_PULL_DAYS = 14;

function pullWeeklyCampaignData() {
  const today = new Date();
  let totalUpdated = 0, totalInserted = 0, totalCampaignRows = 0;
  const datesProcessed = [];

  for (let i = 0; i < WEEKLY_PULL_DAYS; i++) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const fullDate = (d.getMonth()+1) + "/" + d.getDate() + "/" + d.getFullYear();

    const { updated, inserted } = updateChaserCampaignColumnsForDate(fullDate);
    totalCampaignRows += writeCombinedCampaignResponses(fullDate);
    totalUpdated  += updated;
    totalInserted += inserted;
    datesProcessed.push(fullDate);
  }

  Logger.log("pullWeeklyCampaignData(): processed " + datesProcessed.length + " days (" +
    datesProcessed[datesProcessed.length-1] + " to " + datesProcessed[0] + ") | " +
    "Rows updated: " + totalUpdated + " | Rows inserted: " + totalInserted +
    " | Campaign Responses rows written: " + totalCampaignRows);

  return {
    success: true,
    daysProcessed: datesProcessed.length,
    fromDate: datesProcessed[datesProcessed.length-1],
    toDate: datesProcessed[0],
    rowsUpdated: totalUpdated,
    rowsInserted: totalInserted,
    campaignResponseRows: totalCampaignRows,
  };
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
//   pullWeeklyCampaignData() — manual weekly button, writes the trailing
//     WEEKLY_PULL_DAYS days via writeCombinedCampaignResponses(dateTab)
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
  // Without this, a bare "6/25" Date value gets auto-coerced by Sheets into
  // a real Date cell using Sheets' own year guess -- a guessed year silently
  // breaks the dashboard's Month/Quarter views.
  forcePlainTextColumns(sheet, CAMPAIGN_RESPONSES_HEADERS, ["Date"]);
  return sheet;
}

// readCampaignTotalsForDate/readChaserCampaignCountsForDate/archiveCampaignResponses/
// backfillCampaignResponses -- retired along with RESPONSE_SOURCES. The
// Campaign Responses tab and the per-chaser campaign columns are now written
// exclusively by the weekly pull (pullWeeklyCampaignData(), which calls
// writeCombinedCampaignResponses()/updateChaserCampaignColumnsForDate() against
// BACKFILL_RESPONSES_SHEET_ID/BACKFILL_RESPONSES_TABS) below.

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
// in the archive. For ongoing data, the "Pull Weekly Responses" dashboard
// button (pullWeeklyCampaignData()) covers the trailing weeks automatically
// — running this one-time backfill again is only needed to catch up on
// older history it hasn't processed yet.
//
// Reuses BACKFILL_RESPONSES_SHEET_ID / BACKFILL_RESPONSES_TABS (declared
// above, near backfillCampaignColumns()) instead of redeclaring the same
// spreadsheet ID and tab list a second time.
//
// USAGE (run once from the Apps Script editor):
//   backfillFromCombinedSheet()
// Resumable — saves progress after each date, safe to re-run if it times out.
// ============================================================

// Scan every tab in the combined sheet and collect every unique "M/D" date
// found embedded in the feedback column text.
// Resolves a year from a backfill tab's own name (e.g. "ORT Overall 2026"
// -> 2026) -- same pattern as the main archive's month tabs. Used as the
// last-resort default year, for tabs (LY PUMP/LY WRAP) that have no
// conclusionCol to fall back on.
function resolveYearFromTabName(tabName) {
  const m = String(tabName || "").match(/(\d{4})\s*$/);
  return m ? parseInt(m[1], 10) : new Date().getFullYear();
}

// Resolves a row's true full "M/D/YYYY" date. The ORT and CGM tabs carry a
// "Date of conclusion" column -- the actual resolution date, including a
// real year -- which is authoritative and used directly whenever present,
// bypassing feedback-text parsing (and its year ambiguity) entirely. LY
// PUMP/LY WRAP tabs have no such column, so they fall back to parsing the
// feedback text's own first date token, defaulting to the tab's own year
// when that token doesn't carry one -- safe there since neither of those
// campaigns has ever carried anything but current-year data.
function resolveRowDate(feedback, conclusionCellValue, tabYear) {
  if (conclusionCellValue !== null && conclusionCellValue !== undefined && conclusionCellValue !== "") {
    if (conclusionCellValue instanceof Date && !isNaN(conclusionCellValue)) {
      return (conclusionCellValue.getMonth()+1) + "/" + conclusionCellValue.getDate() + "/" + conclusionCellValue.getFullYear();
    }
    const parts = String(conclusionCellValue).trim().split("/");
    if (parts.length >= 3) {
      const m = parseInt(parts[0], 10), d = parseInt(parts[1], 10);
      let y = parseInt(parts[2], 10);
      if (!isNaN(m) && !isNaN(d) && !isNaN(y)) {
        if (y < 100) y += 2000;
        return m + "/" + d + "/" + y;
      }
    }
  }

  const match = feedback.match(/\b(\d{1,2})\/(\d{1,2})(?:\/(\d{2,4}))?\b/);
  if (!match) return null;
  const month = parseInt(match[1], 10), day = parseInt(match[2], 10);
  let year = match[3] ? parseInt(match[3], 10) : null;
  if (year !== null && year < 100) year += 2000;
  if (year === null) year = tabYear;
  return month + "/" + day + "/" + year;
}

// Scan every tab in the combined sheet and collect every unique "M/D/YYYY"
// date, resolved per row via resolveRowDate() above (conclusionCol when
// present, else the feedback text with the tab's year as default).
function collectDatesFromCombinedSheet() {
  const dates = new Set();
  let ss;
  try { ss = SpreadsheetApp.openById(BACKFILL_RESPONSES_SHEET_ID); }
  catch(e) { Logger.log("Cannot open combined sheet: " + e.message); return []; }

  for (const tabConfig of BACKFILL_RESPONSES_TABS) {
    const sheet = ss.getSheetByName(tabConfig.tabName);
    if (!sheet) { Logger.log("Tab not found: " + tabConfig.tabName); continue; }

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers       = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol   = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    const conclusionCol = tabConfig.conclusionCol ? headers.indexOf(tabConfig.conclusionCol.toUpperCase()) : -1;
    if (feedbackCol < 0) { Logger.log("Feedback col not found in " + tabConfig.tabName + ": " + tabConfig.feedbackCol); continue; }

    const tabYear = resolveYearFromTabName(tabConfig.tabName);

    for (let r = 1; r < data.length; r++) {
      const feedback = String(data[r][feedbackCol] || "").trim();
      if (!feedback) continue;
      const fullDate = resolveRowDate(feedback, conclusionCol >= 0 ? data[r][conclusionCol] : null, tabYear);
      if (fullDate) dates.add(fullDate);
    }
  }

  return [...dates];
}

// Read campaign totals (one count per lead row, chasers irrelevant) for ONE
// full "M/D/YYYY" date from the combined historical sheet. Matches each
// row's own resolveRowDate() result exactly against dateTab.
function readCombinedCampaignTotalsForDate(dateTab) {
  const camps = zeroCampaignTotals();

  let ss;
  try { ss = SpreadsheetApp.openById(BACKFILL_RESPONSES_SHEET_ID); }
  catch(e) { return camps; }

  for (const tabConfig of BACKFILL_RESPONSES_TABS) {
    const sheet = ss.getSheetByName(tabConfig.tabName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) continue;

    const headers       = data[0].map(h => String(h).trim().toUpperCase());
    const feedbackCol   = headers.indexOf(tabConfig.feedbackCol.toUpperCase());
    const chaserCol     = headers.indexOf(tabConfig.chaserCol.toUpperCase());
    const conclusionCol = tabConfig.conclusionCol ? headers.indexOf(tabConfig.conclusionCol.toUpperCase()) : -1;
    if (feedbackCol < 0 || chaserCol < 0) continue;

    const key     = tabConfig.campaignKey;
    const tabYear = resolveYearFromTabName(tabConfig.tabName);

    for (let r = 1; r < data.length; r++) {
      const feedback = String(data[r][feedbackCol] || "").trim();
      if (!feedback) continue;
      const fullDate = resolveRowDate(feedback, conclusionCol >= 0 ? data[r][conclusionCol] : null, tabYear);
      if (fullDate !== dateTab) continue;

      const approval = isApproval(feedback);
      const denial   = isDenial(feedback);
      if (!approval && !denial) continue;

      if (approval) camps[key].approved++;
      if (denial)   camps[key].denied++;
    }
  }

  return camps;
}

// Write/overwrite Campaign Responses rows for ONE full "M/D/YYYY" date using
// the combined sheet's data (BACKFILL_RESPONSES_TABS). Called by both
// backfillFromCombinedSheet() (one-time historical) and
// pullWeeklyCampaignData() (ongoing weekly pull).
function writeCombinedCampaignResponses(dateTab) {
  const sheet = getOrCreateCampaignResponsesTab();
  const data  = sheet.getDataRange().getValues();

  const tParts       = dateTab.split("/");
  const targetShort  = parseInt(tParts[0]) + "/" + parseInt(tParts[1]);

  // Matches by M/D alone (not the full date) on purpose: earlier versions of
  // this backfill mis-resolved the year for most rows (year-less entirely,
  // then wrongly defaulted to Submission Date's year), leaving stale rows
  // sitting under the wrong year for the same calendar day. This combined
  // source is a single continuous timeline with no genuine same-M/D
  // collisions across years, so an M/D match is safe here and guarantees
  // those stale mis-dated rows get cleared out the next time this M/D is
  // reprocessed, regardless of what year they were wrongly filed under.
  const toDelete = [];
  for (let r = 1; r < data.length; r++) {
    const cell = data[r][0];
    const parts = cell instanceof Date && !isNaN(cell)
      ? [cell.getMonth()+1, cell.getDate(), cell.getFullYear()]
      : String(cell).trim().split("/");
    if (parts.length < 2) continue;
    const short = parseInt(parts[0]) + "/" + parseInt(parts[1]);
    if (short === targetShort) toDelete.push(r + 1);
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

// ============================================================
// MAIN ENTRY POINT — run this once from the Apps Script editor
// ============================================================
// 1. Finds every date mentioned in the combined sheet's feedback columns
//    (independent of the archive -- works even if a month tab has been
//    deleted entirely, e.g. before re-running migrateExistingSheets()).
// 2. For each date: writes Campaign Responses rows (deduplicated totals).
// 3. For each date: updates per-chaser ORT/CGM/LymphC/LymphW/Approvals/
//    Denials columns via updateChaserCampaignColumnsForDate() -- this
//    creates the monthly archive tab if it doesn't exist yet, and inserts
//    a row (tracker columns blank/0) for any chaser with real campaign
//    credit but no existing row for that date.
// Resumable: progress is saved after each date. If you've deleted archive
// rows/tabs and want this to fully reprocess them, call
// resetCombinedSheetBackfill() first -- otherwise dates already marked
// done will be skipped even though the data underneath them is gone.
function backfillFromCombinedSheet() {
  const PROGRESS_KEY = "combined_sheet_backfill4";
  const props        = PropertiesService.getScriptProperties();
  const doneDates    = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const allDates = collectDatesFromCombinedSheet().filter(d => !doneDates.includes(d));
  Logger.log("Dates found in combined sheet: " + (allDates.length + doneDates.length) + " | Remaining: " + allDates.length);

  let totalCampaignRows = 0;
  let totalChaserUpdated = 0;
  let totalChaserInserted = 0;

  for (const dateTab of allDates) {
    try {
      totalCampaignRows += writeCombinedCampaignResponses(dateTab);
      const { updated, inserted } = updateChaserCampaignColumnsForDate(dateTab);
      totalChaserUpdated  += updated;
      totalChaserInserted += inserted;
    } catch(e) {
      Logger.log("Error processing " + dateTab + ": " + e.message);
    }
    doneDates.push(dateTab);
    props.setProperty(PROGRESS_KEY, JSON.stringify(doneDates));
  }

  Logger.log("=== DONE ===");
  Logger.log("Campaign Responses rows written: " + totalCampaignRows);
  Logger.log("Chaser rows updated: " + totalChaserUpdated + " | inserted: " + totalChaserInserted);
}

function resetCombinedSheetBackfill() {
  PropertiesService.getScriptProperties().deleteProperty("combined_sheet_backfill4");
  Logger.log("Combined sheet backfill progress reset.");
}
