// ============================================================
// LEAD HISTORY — STAGE 1
// Standalone file, lives in the same Apps Script project as Code.gs.
// Triggered daily by Code.gs (via runDailyLeadHistorySync), but all
// lead-state logic is self-contained here.
// ============================================================
//
// PURPOSE
// Tracks every individual lead (by MBI / Insurance ID Number) through
// its full lifecycle across states: InProcess, Yellow, Hold, Approved,
// Denied, VerbalDenial, BTO, Disregarded, Frozen.
//
// This is fundamentally different from Campaign Responses, which only
// stores daily APPROVED/DENIED totals. Lead History tracks every state
// a lead passes through, including non-terminal ones, as a full
// append-only EVENT LOG — one row per transition, never overwritten.
//
// STORAGE — TWO TABS, DIFFERENT JOBS
// "Lead History" — the permanent, append-only audit trail. Every
//   transition ever recorded, forever. Only ever WRITTEN to during the
//   daily sync, never read back — this keeps the sync fast indefinitely
//   no matter how large this tab grows over months/years.
//   Columns: MBI | Campaign | IDN | SubmissionDate | FromStatus | ToStatus | TransitionDate | Chasers
//
// "Current Lead State" — a small, fast lookup tab, UPDATED IN PLACE
//   (not appended). One row per active (MBI, Campaign, SubmissionDate)
//   lifecycle. This is what the daily sync actually reads to figure out
//   "what was true yesterday" — never the full Lead History log.
//   Non-terminal lifecycles stay here indefinitely. Terminal lifecycles
//   are pruned after TERMINAL_RETENTION_DAYS (30 days) by
//   cleanupAgedOutLeadStates() — pruned only from THIS tab, their full
//   record remains permanently in Lead History regardless.
//   Columns: MBI | Campaign | IDN | SubmissionDate | Status | StatusDate | Chasers
//
// Every write to Lead History (via appendLeadHistoryRows) automatically
// keeps Current Lead State in sync — both tabs are always written
// together from one call, so they never drift apart.
//
// STATE MODEL
//   Non-terminal (still being tracked): InProcess, Yellow, Hold
//   Terminal (DaysToResolution calculated once reached):
//     Approved, Denied, VerbalDenial, BTO, Disregarded, Frozen
//
// Every lead's very first appearance produces a "birth" row:
//   FromStatus = "" , ToStatus = "InProcess" (or whatever state it was
//   first seen in — most leads start in InProcess, but a lead could
//   theoretically first be seen in any state if tracking starts mid-life)
//
// SubmissionDate is read once from the source tab (it's a fixed
// attribute of the lead, never changes) and carried on every row for
// that lead so SubmissionDate vs TransitionDate of the birth row shows
// whether a lead was already old when tracking began, or just started.
//
// "DISAPPEARED" HANDLING
// If a lead was in a non-terminal state yesterday and is not found in
// ANY tab today (checked across every state for that campaign), it is
// marked:
//   ToStatus = "Unknown — last seen in <TabName> on <Date>"
// It stays in that state until it reappears somewhere, at which point a
// normal transition resolves it out of Unknown.
//
// SOURCES (daily, ongoing — NOT the one-time backfill sheets)
// Each campaign's 3 main sheets, read across every relevant tab:
//   InProcess, Yellow, Hold, Responses (Approved/Denied), VerbalDenial,
//   BTO, Disregarded, Frozen (LymphW combines the last 3 into one tab
//   disambiguated by its Status column)
//
// This file does NOT touch RESPONSE_SOURCES or any of Code.gs's daily
// productivity/campaign logic — fully additive, fully separate.
// ============================================================


// ============================================================
// LEAD STATE SOURCES — one entry per campaign
// ============================================================
// Each campaign lists every tab that represents a state, with the
// column names specific to that campaign's sheets (already confirmed
// against real headers).
//
// idCol: the column holding the patient identifier.
//   ORT / CGM / LymphW use "MBI". LymphC uses "Insurance ID Number".
// idnCol: the office/identity column, ORT and CGM only (null elsewhere).
// statusCol: the literal "STATUS" / "Status" column in that tab —
//   holds the reason text for non-terminal states, and for LymphW's
//   combined BTO/DIS/FROZEN tab, tells us which of the three it is.
// feedbackCol: the FAX FEEDBACK / FAX SENT ON EST column — used to
//   detect Approved/Denied in Responses tabs (same logic as Code.gs's
//   isApproval/isDenial), and to extract the response date when
//   relevant.
// terminal: true if this tab represents a terminal state.
// stateName: the canonical state name written into FromStatus/ToStatus.
// ============================================================

const LEAD_STATE_SOURCES = {

  ort: {
    sheetId: "1flemAA9Q5hEn78ZtCnGnhDlCVq18RJnjspirfH_uXlU",
    idCol: "MBI",
    idnCol: "IDN",
    submissionCol: "Submission Date",
    chaserCol: "Chaser Name",
    tabs: [
      { name: "ORT INPROCESS",      statusCol: "STATUS", stateName: "InProcess",    terminal: false },
      { name: "ORT YELLOW",         statusCol: "STATUS", stateName: "Yellow",       terminal: false },
      { name: "ORT ON HOLD",        statusCol: "STATUS", stateName: "Hold",         terminal: false },
      { name: "ORT RESPONSES",      statusCol: "STATUS", stateName: null,           terminal: true,  isResponseTab: true,    feedbackCol: "FAX FEEDBACK" },
      { name: "ORT BTO",            statusCol: "STATUS", stateName: "BTO",          terminal: true  },
      { name: "ORT DISREGARDED",    statusCol: "STATUS", stateName: "Disregarded",  terminal: true  },
      { name: "ORT FROZEN",         statusCol: "STATUS", stateName: "Frozen",       terminal: true  },
      { name: "ORT VD / PT CANCEL", statusCol: "STATUS", stateName: "VerbalDenial", terminal: true,  isVerbalDenialTab: true },
    ]
  },

  cgm: {
    sheetId: "1flemAA9Q5hEn78ZtCnGnhDlCVq18RJnjspirfH_uXlU",
    idCol: "MBI",
    idnCol: "IDN",
    submissionCol: "Submission Date",
    chaserCol: "Chaser Name",
    tabs: [
      { name: "CGM INPROCESS",      statusCol: "Status", stateName: "InProcess",    terminal: false },
      { name: "CGM YELLOW",         statusCol: "Status", stateName: "Yellow",       terminal: false },
      { name: "CGM ON HOLD",        statusCol: "Status", stateName: "Hold",         terminal: false },
      { name: "CGM Responses",      statusCol: "Status", stateName: null,           terminal: true,  isResponseTab: true,    feedbackCol: "FAX SENT ON EST" },
      { name: "CGM BTO",            statusCol: "Status", stateName: "BTO",          terminal: true  },
      { name: "CGM DISREGARDED",    statusCol: "Status", stateName: "Disregarded",  terminal: true  },
      { name: "CGM FROZEN",         statusCol: "Status", stateName: "Frozen",       terminal: true  },
      { name: "CGM VD / PT CANCEL", statusCol: "Status", stateName: "VerbalDenial", terminal: true,  isVerbalDenialTab: true },
    ]
  },

  lymphc: {
    sheetId: "1tuMofJVYSzv_Y_kcIkFXFG0PkGRIzLIvh23tXenrL4Y",
    idCol: "Insurance ID Number",
    idnCol: null,
    submissionCol: "Submission Date",
    chaserCol: "Chaser Name",
    tabs: [
      { name: "IN PROCESS",          statusCol: "STATUS", stateName: "InProcess",    terminal: false },
      { name: "YELLOW",              statusCol: "STATUS", stateName: "Yellow",       terminal: false },
      { name: "HOLD",                statusCol: "STATUS", stateName: "Hold",         terminal: false },
      { name: "Responses",           statusCol: "STATUS", stateName: null,           terminal: true,  isResponseTab: true,    feedbackCol: "FAX FEEDBACK" },
      { name: "BTO",                 statusCol: "STATUS", stateName: "BTO",          terminal: true  },
      { name: "Disregarded",         statusCol: "STATUS", stateName: "Disregarded",  terminal: true  },
      { name: "Frozen",              statusCol: "STATUS", stateName: "Frozen",       terminal: true  },
      { name: "VD/Patient Canceled", statusCol: "STATUS", stateName: "VerbalDenial", terminal: true,  isVerbalDenialTab: true },
    ]
  },

  lymphw: {
    sheetId: "1IgcuMvtQ9QAfQRPfh2PmuQWK1bW4jdoDlZITUto35XU",
    idCol: "MBI",
    idnCol: null,
    submissionCol: "Submission Date",
    chaserCol: "Chaser Name",
    tabs: [
      { name: "INPROCESS",       statusCol: "Status", stateName: "InProcess",    terminal: false },
      { name: "Yellow",          statusCol: "Status", stateName: "Yellow",       terminal: false },
      { name: "Hold",            statusCol: "Status", stateName: "Hold",         terminal: false },
      { name: "Responses",       statusCol: "Status", stateName: null,           terminal: true,  isResponseTab: true,    feedbackCol: "Fax Feedback" },
      { name: "VD/PT Cancel",    statusCol: "Status", stateName: "VerbalDenial", terminal: true,  isVerbalDenialTab: true },
      { name: "BTO/DIS/FROZEN",  statusCol: "Status", stateName: null,           terminal: true,  isCombinedTerminal: true },
    ]
  }

};

// ============================================================
// LEAD HISTORY TAB HELPERS
// ============================================================

// Force the given column names to Plain Text formatting on the whole column
// so Google Sheets never silently auto-converts a written date string (e.g.
// "6/20/2026") into a real Date value. Every composite key in this file
// (lifecycle lookups, conflict suppression, cleanup) depends on reading
// back the exact string that was written, so this must hold for as long as
// the tab exists — not just at creation, which is why every getOrCreate*Tab
// call runs this unconditionally rather than only inside the "just created"
// branch.
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

const LEAD_HISTORY_HEADERS = ["MBI","Campaign","IDN","SubmissionDate","Lifecycle","FromStatus","ToStatus","TransitionDate","Chasers"];


function getOrCreateLeadHistoryTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Lead History");
  if (!sheet) {
    sheet = ss.insertSheet("Lead History");
    sheet.appendRow(LEAD_HISTORY_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LEAD_HISTORY_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#1A2C42")
         .setFontColor("#00C2A8");
    Logger.log("Created Lead History tab");
  }
  forcePlainTextColumns(sheet, LEAD_HISTORY_HEADERS, ["SubmissionDate", "TransitionDate"]);
  return sheet;
}

// ============================================================
// LEAD HISTORY CONFLICTS TAB
// ============================================================
// A TRUE conflict is: same MBI + same Campaign + same SubmissionDate,
// found in more than one tab on the same day. That means the same exact
// lifecycle is sitting in two places at once, and the sheets disagree
// about where it actually is — this is NOT something the system should
// guess at. It gets logged here for a human to resolve from the
// dashboard; resolving it is what allows the actual Lead History
// transition row to finally be written.
//
// (This is different from the same MBI appearing twice with DIFFERENT
// Submission Dates — that's not a conflict, that's two genuinely
// separate lifecycles, both handled normally and independently.)
// ============================================================

const LEAD_CONFLICTS_HEADERS = [
  "ConflictID","MBI","Campaign","SubmissionDate","ConflictType","TabsFound","DateDetected",
  "Status","ResolvedBy","ResolvedTo","ResolvedDate"
];
// ConflictType values:
//   "Blocking" — same MBI, same Campaign, same SubmissionDate, found in
//     multiple tabs today. The sheets genuinely disagree about where this
//     ONE lifecycle currently is. Lead History is NOT written for this
//     lifecycle until a human resolves it.
//   "Review"   — same MBI, same Campaign, DIFFERENT SubmissionDates, but
//     more than one lifecycle showed up in tabs on the SAME day (e.g. an
//     Approved lifecycle and a BTO lifecycle both active today). This is
//     not necessarily a bug — could be a legitimate resubmission alongside
//     an old lifecycle that hasn't been cleaned up yet — so each lifecycle
//     IS still written to Lead History normally. The conflict row exists
//     purely so a human can glance at it and confirm nothing's actually
//     wrong, "just in case."

function getOrCreateLeadConflictsTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Lead History Conflicts");
  if (!sheet) {
    sheet = ss.insertSheet("Lead History Conflicts");
    sheet.appendRow(LEAD_CONFLICTS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, LEAD_CONFLICTS_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#1A2C42")
         .setFontColor("#E8B84B");
    Logger.log("Created Lead History Conflicts tab");
  }
  forcePlainTextColumns(sheet, LEAD_CONFLICTS_HEADERS, ["SubmissionDate", "DateDetected", "ResolvedDate"]);
  return sheet;
}

// Returns existing PENDING conflicts as a Set of "MBI|Campaign|SubmissionDate|ConflictType"
// keys, so the daily sync doesn't re-flag the same unresolved conflict every
// single day it remains unresolved. Blocking and Review conflicts are tracked
// separately even for the same MBI+Campaign+SubmissionDate, since they mean
// different things.
function getPendingConflictKeys() {
  const sheet = getOrCreateLeadConflictsTab();
  const data  = sheet.getDataRange().getValues();
  const keys  = new Set();
  if (data.length < 2) return keys;

  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });

  for (let r = 1; r < data.length; r++) {
    const rowStatus = String(data[r][col.Status]).trim();
    // Suppress both "Pending" and "DataError - Pending Fix" rows so the
    // daily sync doesn't re-flag conflicts already known to be in-progress
    if (rowStatus !== "Pending" && rowStatus !== "DataError - Pending Fix") continue;
    const key = String(data[r][col.MBI]).trim() + "|" + String(data[r][col.Campaign]).trim() + "|" + normalizeDateCell(data[r][col.SubmissionDate]) + "|" + String(data[r][col.ConflictType]).trim();
    keys.add(key);
  }
  return keys;
}

// Append new conflict rows in one batch write, auto-assigning sequential
// ConflictID values. Rows passed in should NOT include the ConflictID column —
// it is prepended here automatically.
// Input row shape: [MBI, Campaign, SubmissionDate, ConflictType, TabsFound,
//                  DateDetected, Status, ResolvedBy, ResolvedTo, ResolvedDate]
function appendLeadConflictRows(rows) {
  if (!rows.length) return;
  const sheet    = getOrCreateLeadConflictsTab();
  const lastRow  = sheet.getLastRow(); // includes header row
  // Next ConflictID = how many data rows already exist + 1
  let   nextId   = lastRow; // header is row 1, so lastRow - 1 data rows -> next id = lastRow

  const withIds = rows.map(r => [nextId++, ...r]);
  sheet.getRange(lastRow + 1, 1, withIds.length, LEAD_CONFLICTS_HEADERS.length).setValues(withIds);
  Logger.log("Lead History Conflicts: appended " + withIds.length + " conflict rows (IDs " + (nextId - withIds.length) + " to " + (nextId - 1) + ")");
}

// Returns all PENDING conflicts in a clean shape for the dashboard.
// { conflicts: [{ MBI, Campaign, SubmissionDate, ConflictType, TabsFound, DateDetected }] }
// ConflictType is "Blocking" (lifecycle held back, must be resolved to
// proceed) or "Review" (already processed normally, flagged for visibility
// only — can be acknowledged with no change, or corrected if something
// actually looks wrong).
// Returns all PENDING and DataError conflicts in a clean shape for the dashboard.
// Pending = needs human action (Blocking) or human review (Review).
// DataError - Pending Fix = source sheet has been flagged as wrong, human
//   will see a "Ready to Retry" button once the fix is complete.
// { conflicts: [{ ConflictID, MBI, Campaign, SubmissionDate, ConflictType,
//                TabsFound, DateDetected, Status, ... }] }
function getPendingLeadConflictsForDashboard() {
  const sheet = getOrCreateLeadConflictsTab();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return { conflicts: [] };

  const ACTIVE_STATUSES = new Set(["Pending", "DataError - Pending Fix"]);
  const headers = data[0].map(h => String(h).trim());
  const conflicts = data.slice(1)
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      // Ensure ConflictID is a number for reliable equality checks
      obj.ConflictID = Number(obj.ConflictID);
      return obj;
    })
    .filter(r => ACTIVE_STATUSES.has(String(r.Status).trim()));

  Logger.log("Active lead conflicts for dashboard: " + conflicts.length);
  return { conflicts };
}

// Called from the dashboard once a human takes action on a conflict.
// Looks up by ConflictID (auto-assigned integer in first column).
//
// ACTION values:
//   Blocking conflicts:
//     "pick:<StateName>"  — pick which tab's state is correct, write transition
//     "bothvalid"         — both occurrences valid (e.g. secondary Dr rechase),
//                           writes Lifecycle 1 and Lifecycle 2 rows
//     "dataerror"         — source sheet has a real error, human will fix it;
//                           Status -> "DataError - Pending Fix" (suppressed from
//                           daily re-flagging while fix is in progress)
//     "readytoretry"      — human fixed the source sheet, next daily sync should
//                           re-read it naturally; Status -> "ReadyToRetry"
//     "dismiss"           — known bad data, stop tracking; Status -> "Dismissed"
//   Review conflicts:
//     "acknowledge"       — everything is fine, mark Resolved, no new rows
//     "pick:<StateName>"  — something was actually wrong, write corrective row
//     "dismiss"           — permanently suppress
function resolveLeadConflict(conflictId, action, resolvedBy) {
  const sheet   = getOrCreateLeadConflictsTab();
  const data    = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).trim());
  const col     = {};
  headers.forEach((h, i) => { col[h] = i; });

  const id = Number(conflictId);
  const todayStr = (() => { const d = new Date(); return (d.getMonth()+1)+"/"+d.getDate()+"/"+d.getFullYear(); })();

  for (let r = 1; r < data.length; r++) {
    if (Number(data[r][col.ConflictID]) !== id) continue;

    const currentStatus = String(data[r][col.Status]).trim();
    if (currentStatus === "Resolved" || currentStatus === "Dismissed") {
      return { success: false, error: "Conflict #" + id + " is already " + currentStatus };
    }

    const mbi            = String(data[r][col.MBI]).trim();
    const campaign       = String(data[r][col.Campaign]).trim();
    const submissionDate = normalizeDateCell(data[r][col.SubmissionDate]);
    const conflictType   = String(data[r][col.ConflictType]).trim();
    const tabsFound      = String(data[r][col.TabsFound]).trim();
    const rowNum         = r + 1;
    const act            = String(action || "").toLowerCase().trim();

    // Status-only actions (no Lead History write)
    if (act === "acknowledge") {
      // Review conflicts are purely informational -- every lifecycle involved
      // was already written to Lead History normally, so "acknowledge" just
      // dismisses the flag. A Blocking conflict is different: its lifecycle
      // is deliberately held BACK from Lead History until resolved, so
      // acknowledging one without writing anything would silently and
      // permanently drop that day's transition from the record. Force
      // Blocking conflicts through 'pick' or 'bothvalid' instead, which
      // actually write the resolution.
      if (conflictType !== "Review") {
        return { success: false, error: "'acknowledge' only applies to Review conflicts -- resolve a Blocking conflict with 'pick:State' or 'bothvalid:State1,State2' so the held-back lifecycle actually gets recorded" };
      }
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["Resolved", resolvedBy||"", "Acknowledged", todayStr]]);
      return { success: true, action: "acknowledged" };
    }
    if (act === "dataerror") {
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["DataError - Pending Fix", resolvedBy||"", "", ""]]);
      return { success: true, action: "dataerror" };
    }
    if (act === "readytoretry") {
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["ReadyToRetry", resolvedBy||"", "", todayStr]]);
      return { success: true, action: "readytoretry" };
    }
    if (act === "dismiss") {
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["Dismissed", resolvedBy||"", "Dismissed", todayStr]]);
      return { success: true, action: "dismissed" };
    }

    // Both Valid -- write two parallel lifecycle rows.
    // Action format: "bothvalid:StateForL1,StateForL2"
    // where the human explicitly picks which occurrence is Lifecycle 1 and
    // which is Lifecycle 2 (since both share the same Submission Date, there
    // is no natural chronological ordering to use automatically).
    // Example: "bothvalid:BTO,Disregarded" means L1=BTO, L2=Disregarded.
    if (act.startsWith("bothvalid")) {
      if (conflictType !== "Blocking") {
        return { success: false, error: "'bothvalid' only applies to Blocking conflicts" };
      }
      // Parse the human-specified ordering from the action string, or fall
      // back to TabsFound order if none provided (e.g. plain "bothvalid").
      let states;
      if (action.includes(":")) {
        states = action.slice(action.indexOf(":") + 1).split(",").map(s => s.trim()).filter(Boolean);
      } else {
        states = tabsFound.split(",").map(s => s.trim()).filter(Boolean);
      }
      if (states.length < 2) {
        return { success: false, error: "Need at least two states for bothvalid. Format: bothvalid:State1,State2. TabsFound: " + tabsFound };
      }
      const allCurrent = getCurrentLeadStates();
      const prior1  = allCurrent.get(mbi + "|" + campaign + "|" + submissionDate + "|1");
      const idn     = prior1 ? prior1.idn : "";
      const chasers = prior1 ? prior1.chasers : "";
      appendLeadHistoryRows([
        [mbi, campaign, idn, submissionDate, 1, "", states[0], todayStr, chasers],
        [mbi, campaign, idn, submissionDate, 2, "", states[1], todayStr, chasers],
      ]);
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["Resolved", resolvedBy||"", "Both Valid: L1=" + states[0] + " L2=" + states[1], todayStr]]);
      Logger.log("Conflict #" + id + " Both Valid: Lifecycle 1=" + states[0] + ", Lifecycle 2=" + states[1]);
      return { success: true, action: "bothvalid", lifecycle1: states[0], lifecycle2: states[1] };
    }

    // Pick a specific state
    if (act.startsWith("pick:")) {
      // Only meaningful for Blocking conflicts -- a Review conflict's
      // lifecycle(s) were already recorded normally, so there's nothing
      // held back to "pick" a resolution for. Use 'acknowledge' instead.
      if (conflictType !== "Blocking") {
        return { success: false, error: "'pick' only applies to Blocking conflicts -- a Review conflict's lifecycle was already recorded normally, use 'acknowledge' instead" };
      }
      const chosenState = action.slice(5).trim();
      if (!chosenState) return { success: false, error: "No state specified in pick action" };
      const allCurrent   = getCurrentLeadStates();
      const lifecycleKey = mbi + "|" + campaign + "|" + submissionDate + "|1";
      const prior        = allCurrent.get(lifecycleKey);
      appendLeadHistoryRows([[
        mbi, campaign, prior ? prior.idn : "", submissionDate, 1,
        prior ? prior.status : "", chosenState, todayStr, prior ? prior.chasers : ""
      ]]);
      sheet.getRange(rowNum, col.Status+1, 1, 4).setValues([["Resolved", resolvedBy||"", chosenState, todayStr]]);
      Logger.log("Conflict #" + id + " picked state: " + chosenState);
      return { success: true, action: "pick", resolvedTo: chosenState };
    }

    return { success: false, error: "Unrecognized action: " + action };
  }

  return { success: false, error: "Conflict #" + conflictId + " not found" };
}


// Read the ENTIRE Lead History tab and derive current state per LIFECYCLE.
// A lifecycle's true identity is (MBI, Campaign, SubmissionDate) — NOT just
// (MBI, Campaign) — because the same MBI+Campaign can have multiple distinct
// ============================================================
// CURRENT LEAD STATE — fast lookup tab, NOT append-only
// ============================================================
// Lead History (above) is the permanent, append-only audit trail and is
// never read back during the daily sync — only ever written to. That
// keeps the sync fast indefinitely as Lead History grows.
//
// This tab is the live snapshot: one row per (MBI, Campaign,
// SubmissionDate) lifecycle, UPDATED IN PLACE rather than appended.
// It's what getCurrentLeadStates() actually reads, so the daily diff
// only ever scans a small, roughly-bounded table instead of the entire
// historical log.
//
// RETENTION: non-terminal lifecycles stay here indefinitely (they're
// actively being tracked). Terminal lifecycles (Approved, Denied,
// VerbalDenial, BTO, Disregarded, Frozen, and SUPERSEDED) are kept for
// 30 days after their TransitionDate so recent outcomes are still fast
// to look up, then pruned from this tab by cleanupAgedOutLeadStates().
// Nothing is ever lost — the full record remains permanently in Lead
// History regardless of what happens here.
// ============================================================

const CURRENT_LEAD_STATE_HEADERS = ["MBI","Campaign","IDN","SubmissionDate","Lifecycle","Status","StatusDate","Chasers"];
const TERMINAL_RETENTION_DAYS = 30;

function getOrCreateCurrentLeadStateTab() {
  const ss    = SpreadsheetApp.openById(ARCHIVE_SHEET_ID);
  let   sheet = ss.getSheetByName("Current Lead State");
  if (!sheet) {
    sheet = ss.insertSheet("Current Lead State");
    sheet.appendRow(CURRENT_LEAD_STATE_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, CURRENT_LEAD_STATE_HEADERS.length)
         .setFontWeight("bold")
         .setBackground("#1A2C42")
         .setFontColor("#00C2A8");
    Logger.log("Created Current Lead State tab");
  }
  forcePlainTextColumns(sheet, CURRENT_LEAD_STATE_HEADERS, ["SubmissionDate", "StatusDate"]);
  return sheet;
}

// Read Current Lead State (the small, fast table) and return it in the
// exact same shape getCurrentLeadStates() always returned, so every
// existing caller in syncLeadHistoryForCampaign / resolveLeadConflict
// keeps working unchanged.
//
// Returns a Map keyed by "MBI|Campaign|SubmissionDate" ->
//   { mbi, campaign, idn, submissionDate, status, statusDate, chasers }
function getCurrentLeadStates() {
  const sheet = getOrCreateCurrentLeadStateTab();
  const data  = sheet.getDataRange().getValues();
  const current = new Map();

  if (data.length < 2) return current;

  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    const mbi      = String(row[col.MBI] || "").trim();
    const campaign = String(row[col.Campaign] || "").trim();
    if (!mbi || !campaign) continue;

    const submissionDate = normalizeDateCell(row[col.SubmissionDate]);
    const lifecycle      = Number(row[col.Lifecycle] || 1);
    // Composite key includes Lifecycle so parallel attempts on the same lead
    // (same MBI+Campaign+SubmissionDate, different Lifecycle numbers) each
    // get their own independent entry and are never confused with each other.
    const key = mbi + "|" + campaign + "|" + submissionDate + "|" + lifecycle;
    current.set(key, {
      mbi, campaign, lifecycle,
      idn:            row[col.IDN] || "",
      submissionDate,
      status:         row[col.Status] || "",
      statusDate:     normalizeDateCell(row[col.StatusDate]),
      chasers:        row[col.Chasers] || "",
      _rowNum:        r + 1,
    });
  }

  return current;
}

// Build an index of Current Lead State row numbers keyed by full lifecycle key
// "MBI|Campaign|SubmissionDate|Lifecycle" for in-place updates.
function getCurrentLeadStateRowIndex() {
  const states = getCurrentLeadStates();
  const index = new Map();
  states.forEach((v, key) => index.set(key, v._rowNum));
  return index;
}

// Apply today's new transitions to Current Lead State: update existing
// lifecycle rows in place, append rows for brand-new lifecycles.
//
// NOTE on SUPERSEDED rows: these ARE kept in Current Lead State (not
// deleted), even though they represent "dead" lineage. This is necessary
// so future resubmissions can still compute the correct next Lifecycle
// number (computeLifecycleNumber needs to see the highest lifecycle ever
// assigned for this MBI+Campaign, including superseded ones, or a later
// resubmission could collide with an already-used number once its
// predecessor scrolls out of view).
//
// If SUPERSEDED rows feel cluttered when browsing this tab or querying it
// for the dashboard, filter them out at READ time instead:
//   rows.filter(r => String(r.Status).indexOf("SUPERSEDED") !== 0)
// getActiveLeadStatesForDashboard() below does exactly this — use that
// function (not getCurrentLeadStates directly) for any dashboard display
// that should only show currently-relevant leads.
//
// rows: [MBI, Campaign, IDN, SubmissionDate, Lifecycle, FromStatus, ToStatus, TransitionDate, Chasers]
function upsertCurrentLeadStates(rows) {
  if (!rows.length) return;
  const sheet = getOrCreateCurrentLeadStateTab();
  const rowIndex = getCurrentLeadStateRowIndex();

  const toAppend = [];

  rows.forEach(r => {
    const [mbi, campaign, idn, submissionDate, lifecycle, fromStatus, toStatus, transitionDate, chasers] = r;
    const lc  = Number(lifecycle) || 1;
    const key = mbi + "|" + campaign + "|" + submissionDate + "|" + lc;
    const existingRowNum = rowIndex.get(key);

    if (existingRowNum) {
      sheet.getRange(existingRowNum, 1, 1, CURRENT_LEAD_STATE_HEADERS.length).setValues([[
        mbi, campaign, idn, submissionDate, lc, toStatus, transitionDate, chasers
      ]]);
    } else {
      toAppend.push([mbi, campaign, idn, submissionDate, lc, toStatus, transitionDate, chasers]);
    }
  });

  if (toAppend.length) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, toAppend.length, CURRENT_LEAD_STATE_HEADERS.length).setValues(toAppend);
  }

  Logger.log("Current Lead State: " + (rows.length - toAppend.length) + " updated in place, " + toAppend.length + " new lifecycles added");
}

// Returns Current Lead State rows EXCLUDING SUPERSEDED lifecycles -- use
// this for any dashboard display of "what's currently happening" so dead
// lineage from resubmissions doesn't clutter the view. SUPERSEDED rows are
// still present in the underlying tab (needed for correct Lifecycle
// numbering on future resubmissions) but are filtered out here at read time.
function getActiveLeadStatesForDashboard() {
  const all = getCurrentLeadStates();
  const active = new Map();
  all.forEach((v, key) => {
    if (String(v.status).indexOf("SUPERSEDED") === 0) return;
    active.set(key, v);
  });
  return active;
}

// Remove terminal/superseded lifecycles from Current Lead State once they're
// older than TERMINAL_RETENTION_DAYS. Their full record remains permanently
// in Lead History — this only prunes the fast lookup table so it doesn't
// grow unboundedly. Non-terminal lifecycles (InProcess/Yellow/Hold) are
// never pruned, since they're actively being tracked.
//
// Safe to run daily (cheap — single read + targeted row deletions) or on a
// separate weekly trigger if preferred.
function cleanupAgedOutLeadStates() {
  const sheet = getOrCreateCurrentLeadStateTab();
  const data  = sheet.getDataRange().getValues();
  if (data.length < 2) return 0;

  const headers = data[0].map(h => String(h).trim());
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });

  const TERMINAL_PREFIXES = ["Approved","Denied","VerbalDenial","BTO","Disregarded","Frozen","SUPERSEDED"];
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - TERMINAL_RETENTION_DAYS);

  const rowsToDelete = [];

  for (let r = 1; r < data.length; r++) {
    const status     = String(data[r][col.Status] || "");
    const statusDate = parseLeadHistoryDate(normalizeDateCell(data[r][col.StatusDate]));
    const isTerminal = TERMINAL_PREFIXES.some(p => status.indexOf(p) === 0);

    if (isTerminal && statusDate && statusDate < cutoff) {
      rowsToDelete.push(r + 1); // 1-indexed sheet row
    }
  }

  // Delete from the bottom up so row numbers don't shift mid-deletion
  rowsToDelete.reverse().forEach(rowNum => sheet.deleteRow(rowNum));

  Logger.log("cleanupAgedOutLeadStates: removed " + rowsToDelete.length + " aged-out terminal lifecycles (older than " + TERMINAL_RETENTION_DAYS + " days). They remain permanently in Lead History.");
  return rowsToDelete.length;
}

// Parse a date string in "M/D/YYYY" format (as written by this file) into
// a Date object. Returns null if unparseable.
function parseLeadHistoryDate(str) {
  if (!str) return null;
  const parts = String(str).trim().split("/");
  if (parts.length !== 3) return null;
  const m = parseInt(parts[0]), d = parseInt(parts[1]), y = parseInt(parts[2]);
  if (isNaN(m) || isNaN(d) || isNaN(y)) return null;
  return new Date(y, m - 1, d);
}

// Append a batch of new transition rows to Lead History in one write.
// rows: array of [MBI, Campaign, IDN, SubmissionDate, FromStatus, ToStatus, TransitionDate, Chasers]
// ALWAYS pair this with a call to upsertCurrentLeadStates(rows) using the
// same rows, so the fast lookup table stays in sync with the permanent log.
function appendLeadHistoryRows(rows) {
  // rows: [MBI, Campaign, IDN, SubmissionDate, Lifecycle, FromStatus, ToStatus, TransitionDate, Chasers]
  // Lifecycle=1 for normal leads; 2/3 for parallel "Both Valid" lifecycles.
  if (!rows.length) return;
  const sheet = getOrCreateLeadHistoryTab();
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, LEAD_HISTORY_HEADERS.length).setValues(rows);
  Logger.log("Lead History: appended " + rows.length + " transition rows");
  upsertCurrentLeadStates(rows);
}


// ============================================================
// READING A SINGLE STATE TAB
// ============================================================
// Returns a Map keyed by MBI -> { idn, submissionDate, chasers, statusReason }
// for every row found in this tab. Used for all non-combined, non-response
// state tabs (InProcess, Yellow, Hold, BTO, Disregarded, Frozen except LymphW).
// Returns an ARRAY of { id, idn, submissionDate, chasers, statusReason } —
// one entry per row in the sheet. Deliberately NOT a Map keyed by id, because
// the same MBI can legitimately appear more than once in the SAME tab (e.g.
// two BTO rows for the same MBI with the same Submission Date — a real data
// conflict that needs to be visible to the cross-tab conflict detector in
// syncLeadHistoryForCampaign, not silently collapsed away here).
function readStateTab(campaignConfig, tabConfig) {
  const result = [];
  let ss;
  try { ss = SpreadsheetApp.openById(campaignConfig.sheetId); }
  catch(e) {
    Logger.log("Cannot open sheet for tab " + tabConfig.name + ": " + e.message);
    result.tabMissing = true; result.tabMissingReason = e.message;
    return result;
  }

  const sheet = ss.getSheetByName(tabConfig.name);
  if (!sheet) {
    Logger.log("Tab not found: " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "tab not found";
    return result;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const idCol      = headers.indexOf(campaignConfig.idCol.toUpperCase());
  const idnCol     = campaignConfig.idnCol ? headers.indexOf(campaignConfig.idnCol.toUpperCase()) : -1;
  const subCol     = headers.indexOf(campaignConfig.submissionCol.toUpperCase());
  const chaserCol  = headers.indexOf(campaignConfig.chaserCol.toUpperCase());
  const statusCol  = tabConfig.statusCol ? headers.indexOf(tabConfig.statusCol.toUpperCase()) : -1;

  if (idCol < 0) {
    Logger.log("ID column '" + campaignConfig.idCol + "' not found in " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "ID column '" + campaignConfig.idCol + "' not found";
    return result;
  }
  // SubmissionDate is load-bearing for every lifecycle key in this whole
  // file. A tab that's missing this column (e.g. a Verbal Denial tab whose
  // real layout doesn't match the shared InProcess/Yellow/Hold column
  // config it's borrowing) would otherwise silently produce a blank
  // submissionDate for every row, merging unrelated leads into bogus
  // lifecycle groups. Bail out loud instead of guessing.
  if (subCol < 0) {
    Logger.log("SubmissionDate column '" + campaignConfig.submissionCol + "' not found in " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "SubmissionDate column '" + campaignConfig.submissionCol + "' not found";
    return result;
  }

  for (let r = 1; r < data.length; r++) {
    const id = String(data[r][idCol] || "").trim();
    if (!id) continue;

    result.push({
      id,
      idn:            idnCol >= 0 ? String(data[r][idnCol] || "").trim() : "",
      submissionDate: normalizeDateCell(data[r][subCol]),
      chasers:        chaserCol >= 0 ? String(data[r][chaserCol] || "").trim() : "",
      statusReason:   statusCol >= 0 ? String(data[r][statusCol] || "").trim() : "",
    });
  }

  return result;
}

// Read a Responses tab — uses isApproval/isDenial (from Code.gs) on the
// feedback column to determine whether each row is Approved or Denied.
// Returns an ARRAY of { id, idn, submissionDate, chasers, resolvedAs }
// where resolvedAs is "Approved" or "Denied" — array, not a Map, so
// duplicate MBI rows within this single tab are preserved rather than
// silently collapsed (same reasoning as readStateTab above).
function readResponseTabForLeadHistory(campaignConfig, tabConfig) {
  const result = [];
  let ss;
  try { ss = SpreadsheetApp.openById(campaignConfig.sheetId); }
  catch(e) { result.tabMissing = true; result.tabMissingReason = e.message; return result; }

  const sheet = ss.getSheetByName(tabConfig.name);
  if (!sheet) {
    Logger.log("Tab not found: " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "tab not found";
    return result;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const idCol       = headers.indexOf(campaignConfig.idCol.toUpperCase());
  const idnCol      = campaignConfig.idnCol ? headers.indexOf(campaignConfig.idnCol.toUpperCase()) : -1;
  const subCol      = headers.indexOf(campaignConfig.submissionCol.toUpperCase());
  const chaserCol   = headers.indexOf(campaignConfig.chaserCol.toUpperCase());
  const feedbackCol = headers.indexOf(tabConfig.feedbackCol.toUpperCase());

  if (idCol < 0 || feedbackCol < 0 || subCol < 0) {
    Logger.log("Missing required column in " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "missing required column (ID, feedback, or SubmissionDate)";
    return result;
  }

  for (let r = 1; r < data.length; r++) {
    const id       = String(data[r][idCol] || "").trim();
    const feedback = String(data[r][feedbackCol] || "").trim();
    if (!id || !feedback) continue;

    let resolvedAs = null;
    if (isApproval(feedback)) resolvedAs = "Approved";
    else if (isDenial(feedback)) resolvedAs = "Denied";
    if (!resolvedAs) continue; // row exists but no terminal outcome yet — shouldn't normally happen in a Responses tab

    result.push({
      id,
      idn:            idnCol >= 0 ? String(data[r][idnCol] || "").trim() : "",
      submissionDate: subCol >= 0 ? normalizeDateCell(data[r][subCol]) : "",
      chasers:        chaserCol >= 0 ? String(data[r][chaserCol] || "").trim() : "",
      resolvedAs,
    });
  }

  return result;
}

// LymphW's combined BTO/Disregarded/Frozen tab.
//
// IMPORTANT: the Status column here is NOT a clean category label — it's
// free-text guidance, the same as STATUS in the InProcess/Yellow/Hold tabs
// (e.g. "PT NOT ACTIVE", "Dr Doesn't Sign", "Wrong Dr", "DISREGARD").
// BTO and Disregarded each have their own known vocabulary that never
// overlaps with each other or with Frozen. Frozen has NO distinct
// vocabulary of its own — a Frozen lead carries whatever status text it
// already had from Yellow/InProcess before being moved here after sitting
// too long. So Frozen is classified by ELIMINATION: if the text doesn't
// match a known BTO or Disregarded phrase, and the row is physically in
// this combined tab, it's Frozen.
//
// BTO_PATTERNS / DISREGARDED_PATTERNS are substring matches (case-
// insensitive) against known real phrasing, confirmed against actual
// data. As new phrasing is discovered (via a human resolving a row that
// genuinely doesn't fit), add it to these lists.
const BTO_PATTERNS = [
  "pt not active",
  "dr doesn't sign", "dr doesnt sign", "doctor doesn't sign", "doctor doesnt sign",
  "can't reach office", "cant reach office", "cannot reach office",
  "wrong dr", "wrong doctor",
];
const DISREGARDED_PATTERNS = [
  "disregard", // catches "DISREGARD" and "DISREGARDED"
];

function classifyCombinedTerminalStatus(rawStatus) {
  const text = String(rawStatus || "").trim().toLowerCase();
  if (!text) return null;

  if (BTO_PATTERNS.some(p => text.includes(p)))         return "BTO";
  if (DISREGARDED_PATTERNS.some(p => text.includes(p))) return "Disregarded";

  // No BTO/Disregarded match -> Frozen by elimination (see comment above).
  return "Frozen";
}

// Returns an ARRAY of { id, idn, submissionDate, chasers, statusReason, stateName }
// — array, not a Map, so duplicate MBI rows within this tab are preserved.
function readCombinedTerminalTab(campaignConfig, tabConfig) {
  const result = [];
  let ss;
  try { ss = SpreadsheetApp.openById(campaignConfig.sheetId); }
  catch(e) { result.tabMissing = true; result.tabMissingReason = e.message; return result; }

  const sheet = ss.getSheetByName(tabConfig.name);
  if (!sheet) {
    Logger.log("Tab not found: " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "tab not found";
    return result;
  }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers = data[0].map(h => String(h).trim().toUpperCase());
  const idCol     = headers.indexOf(campaignConfig.idCol.toUpperCase());
  const idnCol    = campaignConfig.idnCol ? headers.indexOf(campaignConfig.idnCol.toUpperCase()) : -1;
  const subCol    = headers.indexOf(campaignConfig.submissionCol.toUpperCase());
  const chaserCol = headers.indexOf(campaignConfig.chaserCol.toUpperCase());
  const statusCol = headers.indexOf(tabConfig.statusCol.toUpperCase());

  if (idCol < 0 || statusCol < 0 || subCol < 0) {
    Logger.log("Missing required column in " + tabConfig.name);
    result.tabMissing = true; result.tabMissingReason = "missing required column (ID, status, or SubmissionDate)";
    return result;
  }

  for (let r = 1; r < data.length; r++) {
    const id     = String(data[r][idCol] || "").trim();
    const rawSt  = String(data[r][statusCol] || "").trim();
    if (!id || !rawSt) continue;

    const stateName = classifyCombinedTerminalStatus(rawSt);
    // classifyCombinedTerminalStatus never returns null for non-empty input
    // (Frozen is the fallback), so this is just a defensive guard.
    if (!stateName) { Logger.log("Could not classify status '" + rawSt + "' in " + tabConfig.name + " for ID " + id); continue; }

    result.push({
      id,
      idn:            idnCol >= 0 ? String(data[r][idnCol] || "").trim() : "",
      submissionDate: subCol >= 0 ? normalizeDateCell(data[r][subCol]) : "",
      chasers:        chaserCol >= 0 ? String(data[r][chaserCol] || "").trim() : "",
      statusReason:   rawSt,
      stateName,
    });
  }

  return result;
}


// Normalize a date cell (Date object or string) to "M/D/YYYY"
// Normalize a date cell — handles real Date objects AND string-formatted
// dates (e.g. "6/20/2026", "06/20/2026", "6/20/26", with or without extra
// whitespace) — all collapse to the exact same "M/D/YYYY" string.
//
// This matters a lot for conflict detection: two rows representing the
// SAME Submission Date must produce identical strings here, or they'll
// be treated as two different lifecycles instead of a same-lifecycle
// conflict (which is exactly what happened before this fix — two BTO
// rows for the same lead, same actual date, but stored with slightly
// different text formatting, so they silently became separate lifecycles
// instead of triggering the conflict check).
function normalizeDateCell(cell) {
  if (cell instanceof Date && !isNaN(cell)) {
    return (cell.getMonth()+1) + "/" + cell.getDate() + "/" + cell.getFullYear();
  }

  const str = String(cell || "").trim();
  if (!str) return "";

  // Try to parse "M/D/YYYY", "M/D/YY", "MM/DD/YYYY" etc. — anything with
  // slash-separated numeric parts — and re-format consistently.
  const parts = str.split("/");
  if (parts.length === 3) {
    const m = parseInt(parts[0], 10);
    const d = parseInt(parts[1], 10);
    let   y = parseInt(parts[2], 10);
    if (!isNaN(m) && !isNaN(d) && !isNaN(y)) {
      if (y < 100) y += 2000; // handle 2-digit years (e.g. "26" -> 2026)
      return m + "/" + d + "/" + y;
    }
  }

  // Fallback: try the generic Date parser (handles ISO-ish strings etc.)
  const parsed = new Date(str);
  if (!isNaN(parsed) && parsed.getFullYear() > 1970) {
    return (parsed.getMonth()+1) + "/" + parsed.getDate() + "/" + parsed.getFullYear();
  }

  // Truly unparseable — return the trimmed original so it's at least
  // consistent for THIS exact string, rather than silently losing it.
  return str;
}


// ============================================================
// DAILY SYNC — THE CORE DIFF ALGORITHM
// ============================================================
// For ONE campaign:
//   1. Read every state tab (InProcess, Yellow, Hold, Responses, BTO,
//      Disregarded, Frozen, VerbalDenial -- or the combined tab for LymphW)
//   2. Build today's snapshot, grouped by MBI: each MBI may have MULTIPLE
//      occurrences today (e.g. seen in both BTO and Disregarded). This is
//      expected and handled correctly below -- it is NOT automatically a
//      data problem, because a lead's true identity is (MBI, Campaign,
//      SubmissionDate). Two occurrences with DIFFERENT Submission Dates are
//      two genuinely separate lifecycles. Two occurrences with the SAME
//      Submission Date in different tabs is a real conflict (the sheets
//      disagree about where this exact lead currently is) and gets logged
//      to the Lead History Conflicts tab for human resolution instead of
//      guessing.
//   3. Read known lifecycle states from Lead History itself, keyed by
//      (MBI, Campaign, SubmissionDate) -- see getCurrentLeadStates().
//   4. For each occurrence found today, resolved against its own
//      SubmissionDate-matched lifecycle:
//      - No prior lifecycle with this SubmissionDate -> birth row
//        (FromStatus = "", ToStatus = today's state). If an OLDER lifecycle
//        exists for the same MBI+Campaign with a DIFFERENT SubmissionDate,
//        that older one gets a "Lead resubmitted, see new lifecycle" note
//        appended so it can be checked and confirmed.
//      - Prior lifecycle exists with the SAME SubmissionDate, state
//        changed -> transition row
//      - Prior lifecycle exists with the SAME SubmissionDate, same state
//        -> nothing to record
//   5. Lifecycles known previously, non-terminal, with NO occurrence found
//      anywhere today -> "Unknown" row (skipped if already marked Unknown,
//      or if the lifecycle is terminal -- terminal leads disappearing from
//      active tabs is expected, not a problem).
//   6. Append all resulting rows (Lead History + Conflicts) in batch writes.
// ============================================================

// Compute the correct Lifecycle number for a new submission date, given all
// known existing lifecycles for the same MBI+Campaign.
//
// Rule: Lifecycle 1 = oldest Submission Date, Lifecycle 2 = next oldest, etc.
// This is assigned at birth (when the row is first written) so Lead History
// always carries the correct number without needing human input.
//
// Edge case: if the computed rank conflicts with an existing lifecycle number
// (rare -- can happen if a backfill introduces a lead with an older Submission
// Date than one already tracked), we assign the next available number above
// the highest existing one and log a warning. The date ordering may not be
// perfectly reflected in that case, but no existing rows are ever renumbered
// (preserves the append-only guarantee).
function computeLifecycleNumber(newSubDate, existingLifecycles) {
  // Parse "M/D/YYYY" into a comparable number (YYYYMMDD)
  const parseSub = s => {
    if (!s) return 99999999;
    const p = String(s).trim().split("/");
    if (p.length !== 3) return 99999999;
    const m = parseInt(p[0]), d = parseInt(p[1]), y = parseInt(p[2]);
    return isNaN(m) || isNaN(d) || isNaN(y) ? 99999999 : y * 10000 + m * 100 + d;
  };

  const newVal = parseSub(newSubDate);
  const existingVals = existingLifecycles.map(e => ({
    val: parseSub(e.submissionDate),
    lifecycle: Number(e.lifecycle) || 1,
  }));

  // Chronological rank among all dates (including the new one)
  const allVals = [...existingVals.map(e => e.val), newVal];
  const sorted  = [...new Set(allVals)].sort((a, b) => a - b);
  const rank    = sorted.indexOf(newVal) + 1; // 1-indexed

  // Check if this rank is already taken by an existing lifecycle
  const taken = new Set(existingVals.map(e => e.lifecycle));
  if (!taken.has(rank)) return rank;

  // Collision (rare backfill edge case) -- use next available above highest
  const maxExisting = Math.max(0, ...existingVals.map(e => e.lifecycle));
  const fallback = maxExisting + 1;
  Logger.log("computeLifecycleNumber: rank " + rank + " already taken for " + newSubDate + " -- using fallback " + fallback);
  return fallback;
}

function syncLeadHistoryForCampaign(campaignKey) {
  const config = LEAD_STATE_SOURCES[campaignKey];
  if (!config) { Logger.log("Unknown campaign: " + campaignKey); return { rows: 0, conflicts: 0 }; }

  const today = new Date();
  const todayStr = (today.getMonth()+1) + "/" + today.getDate() + "/" + today.getFullYear();

  // -- Step 1+2: build today's snapshot, grouped by MBI (array of occurrences) --
  const todayOccurrences = new Map(); // MBI -> [ { stateName, idn, submissionDate, chasers }, ... ]

  const addOccurrence = (id, stateName, info) => {
    if (!todayOccurrences.has(id)) todayOccurrences.set(id, []);
    todayOccurrences.get(id).push({
      stateName,
      idn:            info.idn || "",
      submissionDate: info.submissionDate || "",
      chasers:        info.chasers || "",
    });
  };

  // Tabs that came back missing/unreadable/wrong-shaped today -- a renamed
  // or restructured source tab otherwise fails completely silently (the
  // readers just log and return no rows), which then reads as "every
  // previously-active lead in that state vanished" and gets mass-relabeled
  // Unknown in Step 5 below with nothing pointing a human at the real cause.
  const missingTabs = [];

  for (const tabConfig of config.tabs) {
    if (tabConfig.isCombinedTerminal) {
      // LymphW's BTO/DIS/FROZEN tab
      const rows = readCombinedTerminalTab(config, tabConfig);
      if (rows.tabMissing) missingTabs.push(tabConfig.name + " (" + rows.tabMissingReason + ")");
      rows.forEach(info => addOccurrence(info.id, info.stateName, info));
    } else if (tabConfig.isResponseTab) {
      // Approved / Denied
      const rows = readResponseTabForLeadHistory(config, tabConfig);
      if (rows.tabMissing) missingTabs.push(tabConfig.name + " (" + rows.tabMissingReason + ")");
      rows.forEach(info => addOccurrence(info.id, info.resolvedAs, info));
    } else {
      // InProcess / Yellow / Hold / BTO / Disregarded / Frozen (non-combined)
      const rows = readStateTab(config, tabConfig);
      if (rows.tabMissing) missingTabs.push(tabConfig.name + " (" + rows.tabMissingReason + ")");
      rows.forEach(info => addOccurrence(info.id, tabConfig.stateName, info));
    }
  }

  // Verbal Denial is read above through the normal tabs loop -- VD tab
  // configs live in LEAD_STATE_SOURCES with stateName: "VerbalDenial", so
  // they already go through readStateTab() like every other state tab.

  // -- Step 3: read known lifecycle states for this campaign --
  const allCurrent = getCurrentLeadStates();
  const currentForCampaign = new Map(); // "MBI|SubmissionDate" -> lifecycle info
  const lifecyclesByMbi    = new Map(); // MBI -> [ lifecycle info, ... ]  (all known lifecycles, any SubmissionDate)
  allCurrent.forEach((v) => {
    if (v.campaign !== campaignKey) return;
    // Key now includes Lifecycle so parallel "Both Valid" attempts are tracked
    // independently and never confused with each other in the diff logic.
    const subKey = v.mbi + "|" + v.submissionDate + "|" + (v.lifecycle || 1);
    currentForCampaign.set(subKey, v);
    if (!lifecyclesByMbi.has(v.mbi)) lifecyclesByMbi.set(v.mbi, []);
    lifecyclesByMbi.get(v.mbi).push(v);
  });

  const TERMINAL_STATES = new Set(["Approved","Denied","VerbalDenial","BTO","Disregarded","Frozen"]);
  const pendingConflictKeys = getPendingConflictKeys();

  const newRows      = [];
  const conflictRows = [];
  const seenLifecycleKeysToday = new Set(); // "MBI|SubmissionDate" we successfully processed today

  // Surface missing/broken tabs the same way every other sync anomaly is
  // surfaced -- as a row on the existing Lead History Conflicts dashboard --
  // instead of only a Logger.log line nobody is likely to check. Deduped
  // per campaign so it doesn't re-flag every single day the tab stays missing.
  if (missingTabs.length) {
    const missingTabsStr = missingTabs.join(", ");
    const alertKey = "SYSTEM|" + campaignKey + "|" + missingTabsStr + "|Review";
    Logger.log("*** " + campaignKey + " SYNC INCOMPLETE *** unreadable tab(s): " + missingTabsStr);
    if (!pendingConflictKeys.has(alertKey)) {
      conflictRows.push([
        "SYSTEM", campaignKey, missingTabsStr, "Review",
        "Unreadable source tab(s): " + missingTabsStr, todayStr,
        "Pending", "", "", ""
      ]);
    }
  }

  // -- PRE-PASS: reserve Lifecycle numbers in true chronological order --
  // If the same MBI has multiple DISTINCT, brand-new Submission Dates all
  // appearing for the first time in THIS SAME sync run, we cannot assign
  // Lifecycle numbers as we encounter them in the main loop below, because
  // encounter order depends on tab-read order, not date order. Instead,
  // for each MBI with more than one new Submission Date this run, sort
  // them chronologically FIRST and pre-register them into lifecyclesByMbi
  // in the correct order, so the main loop's computeLifecycleNumber() call
  // always sees the full, correctly-ordered picture no matter which date
  // it happens to process first.
  // Shared chronological sort key for "M/D/YYYY" SubmissionDate strings —
  // used both to pre-order same-batch births below and to make sure the
  // SUPERSEDED check further down only ever fires against a genuinely
  // OLDER submission date, not just a "different" one.
  const dateSortKey = s => {
    if (!s) return 99999999;
    const p = String(s).trim().split("/");
    if (p.length !== 3) return 99999999;
    const m = parseInt(p[0]), d = parseInt(p[1]), y = parseInt(p[2]);
    return isNaN(m) || isNaN(d) || isNaN(y) ? 99999999 : y * 10000 + m * 100 + d;
  };

  const reservedLifecycles = new Map(); // "MBI|SubmissionDate" -> reserved lifecycle number
  todayOccurrences.forEach((occurrences, mbi) => {
    // Distinct Submission Dates for this MBI today that have NO existing
    // Lead History record at all (i.e. would be births)
    const distinctSubDates = [...new Set(occurrences.map(o => o.submissionDate || ""))];
    const brandNewDates = distinctSubDates.filter(sub => {
      // "brand new" = no existing lifecycle already tracked for this exact date
      const anyLifecycleExists = [...currentForCampaign.keys()].some(k => k.startsWith(mbi + "|" + sub + "|"));
      return !anyLifecycleExists;
    });
    if (brandNewDates.length < 2) return; // nothing to pre-order

    // Sort chronologically (same parser logic as computeLifecycleNumber)
    const sortedDates = [...brandNewDates].sort((a, b) => dateSortKey(a) - dateSortKey(b));

    // Reserve numbers starting after any already-existing lifecycle for this MBI
    const existingMax = Math.max(0, ...(lifecyclesByMbi.get(mbi) || []).map(l => Number(l.lifecycle) || 1));
    sortedDates.forEach((sub, i) => {
      const num = existingMax + i + 1;
      reservedLifecycles.set(mbi + "|" + sub, num);
    });
  });

  // -- Step 4: process each MBI's occurrence(s) today --
  todayOccurrences.forEach((occurrences, mbi) => {

    // Group today's occurrences for this MBI by SubmissionDate -- each
    // distinct SubmissionDate is a separate, independent lifecycle.
    const bySubmissionDate = new Map(); // submissionDate -> [ occurrence, ... ]
    occurrences.forEach(occ => {
      const sub = occ.submissionDate || "";
      if (!bySubmissionDate.has(sub)) bySubmissionDate.set(sub, []);
      bySubmissionDate.get(sub).push(occ);
    });

    bySubmissionDate.forEach((occsForThisSub, submissionDate) => {

      if (occsForThisSub.length > 1) {
        // BLOCKING CONFLICT -- same MBI, same Campaign, same SubmissionDate,
        // found in more than one tab today. The sheets disagree about
        // where this exact lifecycle currently is. Do not guess --
        // log it for human resolution, unless already pending. The
        // lifecycle stays at its last known state until resolved.
        const conflictKey = mbi + "|" + campaignKey + "|" + submissionDate + "|Blocking";
        if (!pendingConflictKeys.has(conflictKey)) {
          const tabsFound = occsForThisSub.map(o => o.stateName).join(", ");
          conflictRows.push([
            mbi, campaignKey, submissionDate, "Blocking", tabsFound, todayStr,
            "Pending", "", "", ""
          ]);
        }
        seenLifecycleKeysToday.add(mbi + "|" + submissionDate + "|1");
        return;
      }

      // Exactly one occurrence for this MBI+SubmissionDate today.
      // Check all known Lifecycle numbers for this key -- normally just
      // Lifecycle 1, but could be 1 and 2 if a human previously confirmed
      // "Both Valid" for a parallel-attempt scenario.
      const todayInfo = occsForThisSub[0];

      // Find all existing lifecycles for this MBI+Campaign+SubmissionDate
      const existingLifecycles = [];
      let lc = 1;
      while (true) {
        const lcKey = mbi + "|" + submissionDate + "|" + lc;
        const existing = currentForCampaign.get(lcKey);
        if (!existing) break;
        existingLifecycles.push({ lifecycle: lc, key: lcKey, prior: existing });
        lc++;
      }

      // Mark all found keys as seen so they are not flagged as Unknown
      existingLifecycles.forEach(({ key }) => seenLifecycleKeysToday.add(key));

      if (existingLifecycles.length === 0) {
        // Birth -- first time THIS lifecycle (this exact SubmissionDate) has been seen.
        // Lifecycle number: use the pre-pass reservation if this MBI had multiple
        // brand-new dates this run (guarantees correct chronological order
        // regardless of tab-read order), otherwise compute directly (safe for
        // the common case of a single new date).
        const allKnownLifecycles = lifecyclesByMbi.get(mbi) || [];
        const reservedKey = mbi + "|" + submissionDate;
        const assignedLifecycle = reservedLifecycles.has(reservedKey)
          ? reservedLifecycles.get(reservedKey)
          : computeLifecycleNumber(submissionDate, allKnownLifecycles);
        const lifecycleKey = mbi + "|" + submissionDate + "|" + assignedLifecycle;
        seenLifecycleKeysToday.add(lifecycleKey);
        newRows.push([
          mbi, campaignKey, todayInfo.idn, todayInfo.submissionDate, assignedLifecycle,
          "", todayInfo.stateName, todayStr, todayInfo.chasers
        ]);

        // CRITICAL: update lifecyclesByMbi in-memory immediately so that if
        // another row for the same MBI appears later in this same sync batch
        // (before anything has been written to Current Lead State), it sees
        // this newly-assigned lifecycle and computes its own number correctly.
        // Without this, all same-MBI births in one batch would get Lifecycle 1.
        if (!lifecyclesByMbi.has(mbi)) lifecyclesByMbi.set(mbi, []);
        lifecyclesByMbi.get(mbi).push({
          mbi, campaign: campaignKey,
          submissionDate, lifecycle: assignedLifecycle,
          status: todayInfo.stateName, statusDate: todayStr,
          idn: todayInfo.idn, chasers: todayInfo.chasers,
        });

        // If an OLDER, STILL-ACTIVE lifecycle exists for the same MBI+Campaign
        // with a DIFFERENT SubmissionDate, flag it as superseded -- this is
        // the resubmission case. Only lifecycles that are non-terminal
        // (InProcess/Yellow/Hold) get marked SUPERSEDED -- a lifecycle that
        // already reached a terminal state (Approved/Denied/BTO/etc.) simply
        // completed on its own; it was never preempted by anything, so
        // marking it "superseded" would be semantically wrong and just adds
        // noise. Terminal lifecycles are left exactly as they are.
        const NON_TERMINAL = new Set(["InProcess","Yellow","Hold"]);
        const otherLifecycles = allKnownLifecycles;
        otherLifecycles.forEach(older => {
          if (older.submissionDate === submissionDate) return;
          // Only a genuinely OLDER submission date can be superseded by this
          // birth. bySubmissionDate.forEach above iterates in tab-read order,
          // not date order, so within one batch a birth for an EARLIER date
          // can be processed after a birth for a LATER date for the same
          // MBI -- without this guard, the earlier-date birth would wrongly
          // supersede the later (more current) lifecycle it's processed
          // after, instead of the other way around.
          if (dateSortKey(older.submissionDate) >= dateSortKey(submissionDate)) return;
          if (String(older.status).indexOf("SUPERSEDED") === 0) return; // already marked
          if (!NON_TERMINAL.has(older.status)) return; // terminal -- completed on its own, leave it

          newRows.push([
            mbi, campaignKey, older.idn, older.submissionDate, older.lifecycle || 1,
            older.status, "SUPERSEDED -- Lead resubmitted, see new lifecycle (Submission Date " + submissionDate + ")",
            todayStr, older.chasers
          ]);

          // CRITICAL: mutate this object's status in-memory RIGHT NOW.
          // Since `older` is a direct reference into the array stored in
          // lifecyclesByMbi (JS objects/arrays are reference types), this
          // update is immediately visible to any LATER birth processed in
          // this same batch -- preventing the same lifecycle from being
          // marked SUPERSEDED multiple times when several new lifecycles
          // for the same MBI are all born together (e.g. on a first-ever
          // sync run where multiple historical lifecycles surface at once).
          older.status = "SUPERSEDED -- Lead resubmitted, see new lifecycle (Submission Date " + submissionDate + ")";
        });
        return;
      }

      if (existingLifecycles.length >= 2) {
        // Two parallel lifecycles already confirmed "Both Valid" by a human.
        // The single occurrence today belongs to whichever lifecycle it matches
        // by state -- update both, treating this as a normal (no-conflict) update.
        // For simplicity, update Lifecycle 1 with today's state (the occurrence
        // belongs to the active lifecycle, which is usually 1).
        const prior = existingLifecycles[0].prior;
        const priorIsUnknown = String(prior.status).indexOf("Unknown") === 0;
        if (prior.status !== todayInfo.stateName || priorIsUnknown) {
          newRows.push([
            mbi, campaignKey, todayInfo.idn || prior.idn, prior.submissionDate, 1,
            prior.status, todayInfo.stateName, todayStr, todayInfo.chasers
          ]);
        }
        return;
      }

      // Exactly one existing lifecycle (normal case) -- check for state change
      const { prior, lifecycle } = existingLifecycles[0];
      const lifecycleKey = mbi + "|" + submissionDate + "|" + lifecycle;
      seenLifecycleKeysToday.add(lifecycleKey);

      // Was previously "Unknown" and has now reappeared -- resolve it
      const priorIsUnknown = String(prior.status).indexOf("Unknown") === 0;

      if (prior.status !== todayInfo.stateName || priorIsUnknown) {
        newRows.push([
          mbi, campaignKey, todayInfo.idn || prior.idn, prior.submissionDate, lifecycle,
          prior.status, todayInfo.stateName, todayStr, todayInfo.chasers
        ]);
      }
      // else: same state as before, nothing to record
    });

    // REVIEW CONFLICT -- after processing every distinct SubmissionDate
    // group for this MBI, check whether more than one group was active
    // (i.e. more than one genuinely different lifecycle showed up in tabs
    // TODAY). This is NOT necessarily a bug -- could be a legitimate
    // resubmission alongside an old lifecycle that hasn't been cleaned up
    // -- so every lifecycle above was already written to Lead History
    // normally. This conflict row is purely informational, "just in
    // case," so a human can glance at it and confirm nothing's wrong.
    if (bySubmissionDate.size > 1) {
      const allSubDates = [...bySubmissionDate.keys()];
      const conflictKey = mbi + "|" + campaignKey + "|" + allSubDates.join(",") + "|Review";
      if (!pendingConflictKeys.has(conflictKey)) {
        const tabsFound = allSubDates.map(sub => {
          const occs = bySubmissionDate.get(sub);
          return "[" + sub + ": " + occs.map(o => o.stateName).join("+") + "]";
        }).join(" ");
        // SubmissionDate column holds a comma-joined list since this
        // conflict spans multiple lifecycles, not one specific date.
        conflictRows.push([
          mbi, campaignKey, allSubDates.join(", "), "Review", tabsFound, todayStr,
          "Pending", "", "", ""
        ]);
      }
    }
  });

  // -- Step 5: lifecycles known previously, non-terminal, missing today -> Unknown --
  currentForCampaign.forEach((prior, lifecycleKey) => {
    if (seenLifecycleKeysToday.has(lifecycleKey)) return; // accounted for above (including conflicts)
    if (TERMINAL_STATES.has(prior.status)) return; // terminal leads disappearing is expected, not an issue
    if (String(prior.status).indexOf("Unknown") === 0) return; // already marked unknown, don't re-mark every day
    if (String(prior.status).indexOf("SUPERSEDED") === 0) return; // superseded lifecycles are closed, leave them be

    const lastSeenTab = prior.status;
    newRows.push([
      prior.mbi, campaignKey, prior.idn, prior.submissionDate, prior.lifecycle || 1,
      prior.status, "Unknown -- last seen in " + lastSeenTab + " on " + prior.statusDate, todayStr, prior.chasers
    ]);
  });

  // -- Step 6: write --
  appendLeadHistoryRows(newRows);
  appendLeadConflictRows(conflictRows);

  Logger.log("syncLeadHistoryForCampaign(" + campaignKey + "): " + newRows.length + " transition rows, " + conflictRows.length + " new conflicts");
  return { rows: newRows.length, conflicts: conflictRows.length };
}

// Run the daily sync for all 4 campaigns. This is the function Code.gs's
// eodArchive trigger should call.
function runDailyLeadHistorySync() {
  const campaigns = Object.keys(LEAD_STATE_SOURCES);
  let totalRows = 0;
  let totalConflicts = 0;
  for (const key of campaigns) {
    try {
      const result = syncLeadHistoryForCampaign(key);
      totalRows += result.rows;
      totalConflicts += result.conflicts;
    } catch(e) {
      Logger.log("ERROR syncing Lead History for " + key + ": " + e.message);
    }
  }

  // Prune terminal lifecycles older than 30 days from the fast lookup table.
  // Their full record stays permanently in Lead History regardless.
  let agedOut = 0;
  try {
    agedOut = cleanupAgedOutLeadStates();
  } catch(e) {
    Logger.log("Cleanup of Current Lead State failed (sync still completed normally): " + e.message);
  }

  Logger.log("=== runDailyLeadHistorySync complete: " + totalRows + " transition rows, " + totalConflicts + " new conflicts to resolve, " + agedOut + " aged-out lifecycles pruned ===");
  return { rows: totalRows, conflicts: totalConflicts, agedOut };
}


// ============================================================
// MANUAL TEST / INSPECTION HELPERS
// ============================================================

// Run a sync for just one campaign and log the result — useful for
// verifying tab names and column mappings are correct before relying
// on this in the daily trigger.
function testLeadHistorySyncOneCampaign(campaignKey) {
  const effectiveKey = campaignKey || "ort";
  Logger.log("Testing Lead History sync for: " + effectiveKey + (campaignKey ? "" : " (no argument given, defaulted to ort)"));
  const result = syncLeadHistoryForCampaign(effectiveKey);
  Logger.log("Result: " + JSON.stringify(result));
}

// Dump current state counts per campaign — quick sanity check after a sync.
function logLeadHistorySummary() {
  const all = getCurrentLeadStates();
  const summary = {};
  all.forEach(v => {
    const key = v.campaign + " | " + v.status;
    summary[key] = (summary[key] || 0) + 1;
  });
  const lines = Object.entries(summary).sort().map(([k,v]) => k + ": " + v);
  Logger.log("=== Lead History current state summary ===\n" + lines.join("\n"));
}

// Manually run the 30-day terminal cleanup without a full daily sync —
// useful to test it in isolation, or run on demand if you don't want to
// wait for the next scheduled sync.
function testCleanupAgedOutLeadStates() {
  const removed = cleanupAgedOutLeadStates();
  Logger.log("Removed " + removed + " aged-out terminal lifecycles from Current Lead State.");
}

// Quick sanity check that the two-tab architecture is working as intended:
// Current Lead State should stay small and roughly bounded, while Lead
// History grows over time. If Current Lead State is ever anywhere close
// to Lead History's row count, something's wrong with the pairing.
function logTableSizeComparison() {
  const historySheet = getOrCreateLeadHistoryTab();
  const currentSheet = getOrCreateCurrentLeadStateTab();
  Logger.log("Lead History rows: " + historySheet.getLastRow() +
             " | Current Lead State rows: " + currentSheet.getLastRow());
}

// ============================================================
// STAGE 2 — ONE-TIME HISTORICAL BACKFILL
// ============================================================
// Two sources:
//   1. Combined Responses sheet (ID: 1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU)
//      ORT/CGM/LymphC/LymphW Overall 2026 tabs — Approved and Denied leads
//      for the full year. Has MBI/Insurance ID Number, Submission Date,
//      Chaser, and feedback text with date embedded.
//
//   2. Historical Verbal Denial sheet (ID: 1LPs0zAVmxU6RPh4dC8zZg-vrmf7x4KEyPpPs2M89fyY)
//      ORT 2026 and CGM 2026 tabs — VerbalDenial leads for the full year.
//      LymphC and LymphW are excluded (too new, no historical VD sheet).
//
// For each lead found in these sheets, the backfill writes:
//   - A birth row (FromStatus="", ToStatus=resolved state)
//   anchored to the lead's actual Submission Date.
//
// CONFLICT HANDLING (same rules as daily sync):
//   - Same MBI + Campaign + SubmissionDate already in Lead History
//     → Blocking conflict (holds back, doesn't overwrite history)
//   - Same MBI + Campaign + different SubmissionDate already in Lead History
//     → Review conflict (writes new lifecycle, flags for visibility)
//   - No existing rows → clean birth, write normally
//
// RESUMABLE: progress saved after each tab. Safe to re-run.
// Run resetBackfillProgress() if you need a completely clean restart.
// ============================================================

const BACKFILL_RESPONSES_SHEET_ID = "1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU";
const BACKFILL_VD_SHEET_ID        = "1LPs0zAVmxU6RPh4dC8zZg-vrmf7x4KEyPpPs2M89fyY";

const BACKFILL_RESPONSES_TABS = [
  { tabName: "ORT Overall 2026",     campaignKey: "ort",    idCol: "MBI",                feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: "IDN" },
  { tabName: "CGM Overall 2026",     campaignKey: "cgm",    idCol: "MBI",                feedbackCol: "FAX SENT ON EST", chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: "IDN" },
  { tabName: "LY PUMP Overall 2026", campaignKey: "lymphc", idCol: "Insurance ID Number", feedbackCol: "FAX FEEDBACK",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: null  },
  { tabName: "LY WRAP Overall 2026", campaignKey: "lymphw", idCol: "MBI",                feedbackCol: "Fax Feedback",   chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: null  },
];

const BACKFILL_VD_TABS = [
  { tabName: "ORT 2026", campaignKey: "ort", idCol: "MBI",  chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: "IDN" },
  { tabName: "CGM 2026", campaignKey: "cgm", idCol: "MBI",  chaserCol: "Chaser Name", submissionCol: "Submission Date", idnCol: "IDN" },
];

// Read all rows from one backfill source tab.
// Returns array of { mbi, campaign, idn, submissionDate, chasers, resolvedAs }
// where resolvedAs is "Approved", "Denied", or "VerbalDenial".
function readBackfillTab(sheetId, tabConfig, fixedState) {
  const result = [];
  let ss;
  try { ss = SpreadsheetApp.openById(sheetId); }
  catch(e) { Logger.log("Cannot open backfill sheet " + sheetId + ": " + e.message); return result; }

  const sheet = ss.getSheetByName(tabConfig.tabName);
  if (!sheet) { Logger.log("Backfill tab not found: " + tabConfig.tabName); return result; }

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return result;

  const headers   = data[0].map(h => String(h).trim().toUpperCase());
  const idCol     = headers.indexOf(tabConfig.idCol.toUpperCase());
  const subCol    = headers.indexOf(tabConfig.submissionCol.toUpperCase());
  const chaserCol = headers.indexOf(tabConfig.chaserCol.toUpperCase());
  const idnCol    = tabConfig.idnCol ? headers.indexOf(tabConfig.idnCol.toUpperCase()) : -1;
  const feedbackCol = tabConfig.feedbackCol ? headers.indexOf(tabConfig.feedbackCol.toUpperCase()) : -1;

  if (idCol < 0) { Logger.log("ID column not found in " + tabConfig.tabName); return result; }

  for (let r = 1; r < data.length; r++) {
    const id = String(data[r][idCol] || "").trim();
    if (!id) continue;

    let resolvedAs = fixedState || null;

    // For Responses tabs (Approved/Denied), determine from feedback text
    if (!fixedState && feedbackCol >= 0) {
      const feedback = String(data[r][feedbackCol] || "").trim();
      if (!feedback) continue;
      if (isApproval(feedback))     resolvedAs = "Approved";
      else if (isDenial(feedback))  resolvedAs = "Denied";
      else continue; // no usable outcome in this row
    }

    if (!resolvedAs) continue;

    result.push({
      mbi:            id,
      campaign:       tabConfig.campaignKey,
      idn:            idnCol >= 0 ? String(data[r][idnCol] || "").trim() : "",
      submissionDate: subCol >= 0 ? normalizeDateCell(data[r][subCol]) : "",
      chasers:        chaserCol >= 0 ? String(data[r][chaserCol] || "").trim() : "",
      resolvedAs,
    });
  }

  Logger.log("readBackfillTab(" + tabConfig.tabName + "): " + result.length + " rows");
  return result;
}

// Process one backfill tab's rows against the current Lead History state.
// Applies the same conflict logic as the daily sync:
//   - No prior history for this MBI+Campaign+SubmissionDate -> birth row
//   - Same SubmissionDate already exists -> Blocking conflict
//   - Different SubmissionDate already exists -> Review conflict + write new lifecycle
// Returns { written, blocking, review }
function processBackfillRows(rows, currentStates, todayStr) {
  const newRows      = [];
  const conflictRows = [];
  const pendingKeys  = getPendingConflictKeys();
  const TERMINAL_STATES = new Set(["Approved","Denied","VerbalDenial","BTO","Disregarded","Frozen"]);

  // Group rows by MBI+Campaign to detect within-source conflicts
  const byMbiCampaign = new Map();
  rows.forEach(row => {
    const key = row.mbi + "|" + row.campaign;
    if (!byMbiCampaign.has(key)) byMbiCampaign.set(key, []);
    byMbiCampaign.get(key).push(row);
  });

  byMbiCampaign.forEach((occurrences, mbiCampaignKey) => {
    // Group by SubmissionDate within this MBI+Campaign
    const bySubDate = new Map();
    occurrences.forEach(occ => {
      const sub = occ.submissionDate || "";
      if (!bySubDate.has(sub)) bySubDate.set(sub, []);
      bySubDate.get(sub).push(occ);
    });

    bySubDate.forEach((occs, submissionDate) => {
      const row = occs[0]; // representative row for this lifecycle

      // Same MBI+Campaign+SubmissionDate appearing multiple times in the
      // backfill source itself — treat as Blocking
      if (occs.length > 1) {
        const conflictKey = row.mbi + "|" + row.campaign + "|" + submissionDate + "|Blocking";
        if (!pendingKeys.has(conflictKey)) {
          const tabsFound = occs.map(o => o.resolvedAs).join(", ");
          conflictRows.push([
            row.mbi, row.campaign, submissionDate, "Blocking", tabsFound, todayStr,
            "Pending", "", "", ""
          ]);
        }
        return;
      }

      // Check existing Lead History state for this lifecycle
      const lifecycleKey = row.mbi + "|" + row.campaign + "|" + submissionDate;
      const existingLife  = currentStates.get(lifecycleKey);

      if (existingLife) {
        // Lifecycle already exists in Lead History with this exact SubmissionDate
        // -- Blocking conflict, don't overwrite
        const conflictKey = row.mbi + "|" + row.campaign + "|" + submissionDate + "|Blocking";
        if (!pendingKeys.has(conflictKey)) {
          conflictRows.push([
            row.mbi, row.campaign, submissionDate, "Blocking",
            "BackfillSource:" + row.resolvedAs + " vs LeadHistory:" + existingLife.status,
            todayStr, "Pending", "", "", ""
          ]);
        }
        return;
      }

      // Check if any OTHER lifecycle exists for same MBI+Campaign (different SubmissionDate)
      // That's a Review conflict -- but still write this one normally
      const otherLifecycles = [...currentStates.entries()]
        .filter(([k]) => k.startsWith(row.mbi + "|" + row.campaign + "|") && !k.endsWith("|" + submissionDate))
        .map(([, v]) => v);

      if (otherLifecycles.length > 0) {
        const allDates = [submissionDate, ...otherLifecycles.map(l => l.submissionDate)].sort().join(", ");
        const conflictKey = row.mbi + "|" + row.campaign + "|" + allDates + "|Review";
        if (!pendingKeys.has(conflictKey)) {
          const tabsFound = "[" + submissionDate + ": " + row.resolvedAs + "] " +
            otherLifecycles.map(l => "[" + l.submissionDate + ": " + l.status + "]").join(" ");
          conflictRows.push([
            row.mbi, row.campaign, allDates, "Review", tabsFound, todayStr,
            "Pending", "", "", ""
          ]);
        }
        // Fall through -- still write this lifecycle's birth row
      }

      // Write birth row for this lifecycle
      newRows.push([
        row.mbi, row.campaign, row.idn, row.submissionDate, 1,
        "", row.resolvedAs, todayStr, row.chasers
      ]);

      // Also update currentStates in memory so subsequent rows in the same
      // batch can detect Review conflicts against this newly-added lifecycle
      currentStates.set(lifecycleKey, {
        mbi: row.mbi, campaign: row.campaign,
        idn: row.idn, submissionDate: row.submissionDate,
        status: row.resolvedAs, statusDate: todayStr, chasers: row.chasers,
      });
    });

    // Review conflict for multiple DIFFERENT SubmissionDates from this source alone
    if (bySubDate.size > 1) {
      const allDates = [...bySubDate.keys()].sort().join(", ");
      const conflictKey = occurrences[0].mbi + "|" + occurrences[0].campaign + "|" + allDates + "|Review";
      if (!pendingKeys.has(conflictKey)) {
        const tabsFound = [...bySubDate.entries()].map(([sub, occs]) =>
          "[" + sub + ": " + occs.map(o => o.resolvedAs).join("+") + "]"
        ).join(" ");
        conflictRows.push([
          occurrences[0].mbi, occurrences[0].campaign, allDates, "Review", tabsFound, todayStr,
          "Pending", "", "", ""
        ]);
      }
    }
  });

  appendLeadHistoryRows(newRows);
  appendLeadConflictRows(conflictRows);

  return { written: newRows.length, blocking: conflictRows.filter(r => r[3] === "Blocking").length, review: conflictRows.filter(r => r[3] === "Review").length };
}

// ── MAIN BACKFILL ENTRY POINTS ────────────────────────────────

// One-time backfill from the combined Responses sheet (Approved + Denied)
function backfillFromResponsesSheet() {
  const PROGRESS_KEY = "lh_backfill_responses";
  const props        = PropertiesService.getScriptProperties();
  const done         = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const today    = new Date();
  const todayStr = (today.getMonth()+1) + "/" + today.getDate() + "/" + today.getFullYear();
  const current  = getCurrentLeadStates();

  let totalWritten = 0, totalBlocking = 0, totalReview = 0;

  for (const tabConfig of BACKFILL_RESPONSES_TABS) {
    if (done.includes(tabConfig.tabName)) {
      Logger.log("Skipping already-processed tab: " + tabConfig.tabName);
      continue;
    }

    Logger.log("Processing backfill tab: " + tabConfig.tabName);
    const rows   = readBackfillTab(BACKFILL_RESPONSES_SHEET_ID, tabConfig, null);
    const result = processBackfillRows(rows, current, todayStr);

    totalWritten  += result.written;
    totalBlocking += result.blocking;
    totalReview   += result.review;

    done.push(tabConfig.tabName);
    props.setProperty(PROGRESS_KEY, JSON.stringify(done));
    Logger.log("Tab complete: " + result.written + " written, " + result.blocking + " blocking, " + result.review + " review");
  }

  Logger.log("=== backfillFromResponsesSheet DONE: " + totalWritten + " rows, " + totalBlocking + " blocking conflicts, " + totalReview + " review conflicts ===");
}

// One-time backfill from the historical Verbal Denial sheet (ORT + CGM only)
function backfillFromVerbalDenialSheet() {
  const PROGRESS_KEY = "lh_backfill_vd";
  const props        = PropertiesService.getScriptProperties();
  const done         = JSON.parse(props.getProperty(PROGRESS_KEY) || "[]");

  const today    = new Date();
  const todayStr = (today.getMonth()+1) + "/" + today.getDate() + "/" + today.getFullYear();
  const current  = getCurrentLeadStates();

  let totalWritten = 0, totalBlocking = 0, totalReview = 0;

  for (const tabConfig of BACKFILL_VD_TABS) {
    if (done.includes(tabConfig.tabName)) {
      Logger.log("Skipping already-processed tab: " + tabConfig.tabName);
      continue;
    }

    Logger.log("Processing VD backfill tab: " + tabConfig.tabName);
    // fixedState="VerbalDenial" -- no feedback column needed, every row is a verbal denial
    const rows   = readBackfillTab(BACKFILL_VD_SHEET_ID, tabConfig, "VerbalDenial");
    const result = processBackfillRows(rows, current, todayStr);

    totalWritten  += result.written;
    totalBlocking += result.blocking;
    totalReview   += result.review;

    done.push(tabConfig.tabName);
    props.setProperty(PROGRESS_KEY, JSON.stringify(done));
    Logger.log("Tab complete: " + result.written + " written, " + result.blocking + " blocking, " + result.review + " review");
  }

  Logger.log("=== backfillFromVerbalDenialSheet DONE: " + totalWritten + " rows, " + totalBlocking + " blocking conflicts, " + totalReview + " review conflicts ===");
}

// Run both backfills in sequence. Safe to call once — each tab is
// individually tracked and skipped if already processed.
function runAllBackfills() {
  Logger.log("=== Starting Lead History backfill ===");
  backfillFromResponsesSheet();
  backfillFromVerbalDenialSheet();
  Logger.log("=== All backfills complete ===");
}

// Reset backfill progress so all tabs are reprocessed from scratch.
// WARNING: also clear Lead History and Current Lead State first if you
// want a truly clean restart, otherwise you will get Blocking conflicts
// for every row that already exists.
function resetBackfillProgress() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("lh_backfill_responses");
  props.deleteProperty("lh_backfill_vd");
  Logger.log("Backfill progress reset. Run runAllBackfills() to start fresh.");
}
