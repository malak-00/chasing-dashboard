# DME Desk Chasing Dashboard + Lead History — Handoff Notes

This document captures the architecture, decisions, and open items from the
conversation that built this system, so a fresh Claude Code session (or a
future you) doesn't have to rediscover any of it.

## Files in this package

- `Code.gs` — main Apps Script backend: chaser productivity archiving,
  Campaign Responses tracking, Settings/Utlatel handling, the web app API
  the dashboard talks to.
- `LeadHistory.gs` — separate Apps Script file, same project. Tracks
  individual leads (by MBI) through their full lifecycle across states.
  Triggered daily by Code.gs but fully self-contained.
- `dashboard.html` — single-file frontend hosted on GitHub Pages. Talks to
  Code.gs's web app deployment via `mode=` query params.

All three are currently syntax-valid (verified with `node --check` against
an ASCII-stripped copy, since the files contain em-dashes and box-drawing
characters in comments that Node's strict parser chokes on but Apps
Script's V8 runtime handles fine — this is a known non-issue, not a bug).

---

## Part 1 — Dashboard system (Code.gs + dashboard.html)

### What it does
Tracks 5 chasers (Alex, Hope, Rose, Frank, Nova) working DME leads across 4
campaigns (ORT, CGM, LymphC, LymphW). Archives daily productivity (cases,
time, efficiency) and campaign response totals (approvals/denials) into a
Google Sheets "archive" spreadsheet. Dashboard reads from that archive via
a `mode=` API.

### Key architecture decisions
- **Archive spreadsheet** (`ARCHIVE_SHEET_ID = "1eRjelUuuEt0JJCPWw6ohhYe9utWlWdQHlontpOZQnEw"`)
  has monthly tabs (one row per chaser per day) plus a `Campaign Responses`
  tab (one row per date+campaign, deduplicated counts — NOT summed from
  chaser rows, since a lead can have multiple chasers and summing
  double-counts).
- **Campaign Responses tab is the source of truth for campaign totals.**
  The per-chaser ORT/CGM/etc. columns in monthly tabs are for *individual
  chaser attribution* (how many leads a chaser touched) — these should
  NEVER be summed across chasers for team totals, that was a bug we fixed
  early in this project.
- **Sync button** (`syncSelectedDate()` in dashboard.html) re-reads
  whichever date is picked in the date picker — not hardcoded to "today"
  (that was also a bug, fixed).
- **Shift time** is entered as raw minutes via a number input (numpad-
  friendly, for meeting-adjusted shifts) — not a Full/Half/Quarter/Off
  dropdown. `shiftMap` stores minutes directly; `DEFAULT_SHIFT_MINS = 400`.
- **Formula settings** (Productivity/Efficiency targets) support per-chaser
  AND dated overrides via `formulaOverrides` array, resolved by
  `getGoalTarget(metricKey, chaserName, dateTab)` — most-specific-first
  (per-chaser+dated > all-chasers+dated > global base).
- Campaigns tab has Day/Week/Month/Quarter views (Month/Quarter added late,
  read directly from Campaign Responses via `buildCampaignFromLeadsDateRange`).

### Known-fixed bugs (don't reintroduce)
- Leaderboard sort arrows: CSS `.sort-asc`/`.sort-desc` must be separate
  rules, not both targeting `.sort-asc::after`.
- `filterByPeriod("today")` must compare parsed Date objects (via
  `parseArchiveDate`), not raw string equality — archive dates can be
  `M/D`, `M/D/YYYY`, or `YYYY-MM-DD`.
- Date range filters (History tab, Compare tab, Chart builder) must parse
  `<input type="date">` ISO strings as **local midnight**
  (`new Date(y, m-1, d)`), not `new Date(isoString)` — the latter parses as
  UTC and causes an off-by-one day in any timezone east of UTC.
- `saveUtlatelToArchive()` must dedupe by `date|agent` key before appending
  to `utlatelPersist`, or calling Save twice duplicates in-memory rows.
- A stray newline once broke a `"FAXES SENT"` string literal across two
  lines in Code.gs (search for it if `node --check` ever fails again).

### Not yet built
- Dashboard UI for browsing/resolving Lead History conflicts (backend is
  ready, see Part 2).
- "Leads" metric (deduplicated by MBI) to replace "Cases" in History,
  Overview, and Compare tabs — deferred until Lead History is proven
  reliable, since it's the data source for a true unique-lead count.
- Verbal Denials as a visible line item in the Campaigns tab, plus an "All
  Denied" (Denied + Verbal Denied) combined view — also deferred, depends
  on Lead History since Verbal Denial lives in its own tab per campaign,
  same as In Process/Yellow/etc.

---

## Part 2 — Lead History system (LeadHistory.gs)

### Why this exists
The user wanted historical comparison of lead pipeline states (In Process,
Yellow, Hold, Approved, Denied, Verbal Denial, BTO, Disregarded, Frozen) —
not just daily Approved/Denied counts. Since leads have a real identity
(MBI), true lifecycle tracking is possible: know exactly what happened to
every individual lead over time.

### The state model
| State | Terminal? | Notes |
|---|---|---|
| InProcess | No | active work |
| Yellow | No | needs PT outreach |
| Hold | No | paused for a period |
| Approved | Yes | |
| Denied | Yes | |
| VerbalDenial | Yes | separate from Denied; combinable as "All Denied" |
| BTO | Yes | sent back to sales team |
| Disregarded | Yes | lead no longer qualifies |
| Frozen | Yes | manually moved after sitting stagnant too long (NOT an automated timeout — chasers move it manually as a judgment call) |

All state transitions are **manual moves by chasers between tabs** — no
automated rules the system needs to enforce or predict.

### Sheet/tab configuration (confirmed against real data)
```
ORT:    sheetId 1flemAA9Q5hEn78ZtCnGnhDlCVq18RJnjspirfH_uXlU
        tabs: ORT INPROCESS, ORT YELLOW, ORT ON HOLD, ORT RESPONSES,
              ORT BTO, ORT DISREGARDED, ORT FROZEN, ORT VD / PT CANCEL
        idCol: MBI, idnCol: IDN

CGM:    same sheetId as ORT (different tabs)
        tabs: CGM INPROCESS, CGM YELLOW, CGM ON HOLD, CGM Responses,
              CGM BTO, CGM DISREGARDED, CGM FROZEN, CGM VD / PT CANCEL
        idCol: MBI, idnCol: IDN

LymphC: sheetId 1tuMofJVYSzv_Y_kcIkFXFG0PkGRIzLIvh23tXenrL4Y
        tabs: IN PROCESS, YELLOW, HOLD, Responses, BTO, Disregarded,
              Frozen, VD/Patient Canceled
        idCol: Insurance ID Number, idnCol: null (only one identity)

LymphW: sheetId 1IgcuMvtQ9QAfQRPfh2PmuQWK1bW4jdoDlZITUto35XU
        tabs: INPROCESS, Yellow, Hold, Responses, VD/PT Cancel,
              BTO/DIS/FROZEN (combined tab, see below)
        idCol: MBI, idnCol: null
```
IDN = "office identity" a lead is worked from, relevant only to ORT/CGM
(some patients need different offices depending on state licensing). It's
a stable attribute, never changes mid-lifecycle — just carried as context.

**LymphW's `BTO/DIS/FROZEN` tab** combines three terminal states into one,
disambiguated by free-text `Status` column values — NOT clean category
labels. Classified via `classifyCombinedTerminalStatus()` using known
phrase patterns:
```js
BTO_PATTERNS = ["pt not active", "dr doesn't sign", "can't reach office", "wrong dr", ...]
DISREGARDED_PATTERNS = ["disregard"]
// anything matching neither -> Frozen by elimination (Frozen has no
// distinct vocabulary of its own; it inherits whatever Yellow/InProcess
// text it had before being moved)
```
**This pattern list will need ongoing maintenance** as chasers type new
phrasing. Add new confirmed phrases to these arrays as they come up.

### Two-tab architecture (important — don't collapse back to one)
- **`Lead History`** — permanent, append-only event log. One row per
  transition, ever. **Never read back in bulk during sync** — this is
  what keeps daily sync fast indefinitely regardless of table size.
  Columns: `MBI | Campaign | IDN | SubmissionDate | Lifecycle | FromStatus
  | ToStatus | TransitionDate | Chasers`
- **`Current Lead State`** — small, fast lookup, UPDATED IN PLACE (not
  appended). One row per active `(MBI, Campaign, SubmissionDate,
  Lifecycle)` tuple. This is what `getCurrentLeadStates()` actually reads
  every day. Terminal lifecycles (including SUPERSEDED) are pruned after
  `TERMINAL_RETENTION_DAYS = 30` by `cleanupAgedOutLeadStates()` — pruned
  only from this tab, full record stays in Lead History forever.
- **Pairing is automatic**: `appendLeadHistoryRows()` always calls
  `upsertCurrentLeadStates()` internally. Every write path goes through
  this one function, so the two tabs can never drift apart. If you add a
  new code path that writes transitions, route it through
  `appendLeadHistoryRows()`, don't write Lead History directly.

### Lifecycle numbering (handles resubmissions / parallel attempts)
A lead can have multiple independent "lifecycles" under the same
MBI+Campaign:
- **Resubmission** (different Submission Date) → new lifecycle, older one
  gets marked `SUPERSEDED -- Lead resubmitted, see new lifecycle (...)`.
  **Only non-terminal (InProcess/Yellow/Hold) older lifecycles get
  superseded** — an already-terminal lifecycle (Approved etc.) simply
  completed on its own and is left untouched.
- **Same Submission Date, both valid** (e.g. rechased for a secondary Dr)
  → human resolves via dashboard as "Both Valid", gets Lifecycle 1 and 2
  as two independent parallel tracks.

**Lifecycle number = chronological rank by Submission Date** (oldest = 1),
assigned automatically at birth, no human input needed for the normal
resubmission case. This required a **pre-pass** (`computeLifecycleNumber`
+ the reservation block at the top of `syncLeadHistoryForCampaign`) because
naively assigning numbers as rows are encountered gives wrong results —
processing order depends on tab-read order, not date order, so the first
lifecycle *encountered* would always get rank 1 regardless of whether it's
actually the oldest. The pre-pass sorts all brand-new Submission Dates for
an MBI chronologically BEFORE the main loop runs, reserving correct numbers
upfront.

**Known subtlety**: when an object in `lifecyclesByMbi` gets marked
SUPERSEDED mid-batch, the code mutates the object in place (`older.status =
"SUPERSEDED..."`) rather than just pushing a row — because JS arrays/objects
are references, this makes the change visible to any later birth processed
in the same sync run, preventing the same lifecycle from being marked
SUPERSEDED multiple times in one batch. This was a real bug we hit and
fixed — if lifecycle numbering or superseding logic is ever touched again,
re-verify this still holds.

### Conflict detection & resolution
Two types, tracked in `Lead History Conflicts` tab
(`ConflictID | MBI | Campaign | SubmissionDate | ConflictType | TabsFound |
DateDetected | Status | ResolvedBy | ResolvedTo | ResolvedDate`):

- **Blocking** — same MBI+Campaign+SubmissionDate found in >1 tab today.
  Genuinely contradictory. Lifecycle is held back from Lead History until
  resolved.
- **Review** — same MBI+Campaign, different Submission Dates, multiple
  lifecycles active same day (e.g. old Approved lifecycle + new
  resubmission both showing up). NOT necessarily a bug — lifecycles are
  written normally, this is purely a visibility flag.

`resolveLeadConflict(conflictId, action, resolvedBy)` — actions:
- `"pick:StateName"` — write that state as the resolution
- `"bothvalid:State1,State2"` — human-ordered, writes Lifecycle 1=State1,
  Lifecycle 2=State2 (order matters, no automatic date-based ordering
  possible since both share the same Submission Date)
- `"dataerror"` → Status becomes `"DataError - Pending Fix"`, suppressed
  from re-flagging while human fixes the source sheet
- `"readytoretry"` → clears suppression, next daily sync re-reads
  naturally
- `"dismiss"` → permanently suppressed
- `"acknowledge"` (Review only) → mark Resolved, no Lead History write

`getPendingLeadConflictsForDashboard()` returns both `"Pending"` and
`"DataError - Pending Fix"` rows for dashboard display.
`getActiveLeadStatesForDashboard()` returns Current Lead State filtered to
exclude SUPERSEDED rows — use this (not `getCurrentLeadStates()` directly)
for any "what's happening right now" dashboard view.

### Verbal Denial — daily reads now wired in, NOT YET TESTED end-to-end
Verbal Denial tabs (`ORT VD / PT CANCEL`, `CGM VD / PT CANCEL`,
`VD/Patient Canceled`, `VD/PT Cancel`) are configured directly in
`LEAD_STATE_SOURCES` with `isVerbalDenialTab: true`, read via the standard
`readStateTab()` path (same column shape as every other state tab). This
was added late in the conversation and has NOT been run against real data
yet — test with `testLeadHistorySyncOneCampaign()` for each campaign and
watch for column-mismatch warnings before trusting it.

### Stage 2 — one-time historical backfill (built, NOT YET RUN)
Two sources, run once via `runAllBackfills()`:
- `backfillFromResponsesSheet()` — combined Responses sheet
  (`1QVnmXYRg-IbMi46Lvi752ZbIfL5NHd666DsH0WI7SmU`), tabs "ORT Overall
  2026" / "CGM Overall 2026" / "LY PUMP Overall 2026" / "LY WRAP Overall
  2026". Produces birth+terminal rows for historically Approved/Denied
  leads.
- `backfillFromVerbalDenialSheet()` — historical VD sheet
  (`1LPs0zAVmxU6RPh4dC8zZg-vrmf7x4KEyPpPs2M89fyY`), tabs "ORT 2026" / "CGM
  2026" only (LymphC/LymphW excluded — too new, no historical VD data for
  them).

Both resumable (progress tracked via `PropertiesService`, per-tab), apply
the same Blocking/Review conflict logic as daily sync. **A third historical
source exists but is intentionally excluded**: a numbers-only weekly
summary sheet (Submitted/InProcess/Approved/Denied counts, no MBI) — this
cannot feed Lead History since there's no lead identity to key rows to.
User wants this used for a separate "Weekly Summary" aggregate tab instead
(not yet built), backfilled once from that sheet, then kept current going
forward by deriving counts from Lead History itself.

Run order note: **clear the three Lead History tabs before running
backfill** if there's any leftover test data from earlier debugging
sessions — several column-shape bugs were found and fixed during testing
(see below), and old test data was written under earlier, incorrect
schemas.

### Bugs found and fixed during testing (context for why the code looks the way it does)
1. Readers (`readStateTab`, `readResponseTabForLeadHistory`,
   `readCombinedTerminalTab`) originally returned `Map<MBI, info>`, which
   silently collapsed duplicate MBI rows *within a single tab* before the
   conflict detector ever saw them. Fixed by returning arrays instead.
2. `normalizeDateCell()` only handled real Date objects; string-formatted
   dates (`"6/20/2026"` vs `"06/20/2026"` vs 2-digit years) weren't
   normalized to a consistent format, causing two rows with the "same"
   date to be treated as different lifecycles. Fixed to parse and
   re-format any date-like string consistently.
3. LymphW's `BTO/DIS/FROZEN` tab was originally assumed to have clean
   category labels in its Status column — it doesn't, it has free-text
   reasons like real chasing tabs. Required the keyword-classifier
   rewrite described above.
4. Lifecycle numbering: see "Known subtlety" above — same-batch multiple
   births for one MBI need in-memory reservation/mutation to avoid
   assigning wrong numbers or duplicate SUPERSEDED markings.
5. `getCurrentLeadStates()` / `getPendingConflictKeys()` etc. all needed
   updating in lockstep whenever the Lead History or Conflicts tab schema
   changed (Lifecycle column added, ConflictID column added) — if you add
   another column to either tab in the future, grep for every place that
   does `headers.forEach((h,i) => col[h]=i)` and hardcoded array
   destructuring (`const [mbi, campaign, ...] = row`) to make sure nothing
   is silently reading the wrong column.

### Manual test/entry-point functions (Apps Script editor, run directly)
```
testLeadHistorySyncOneCampaign("ort")   // run one campaign, inspect log
testLeadHistorySyncOneCampaign("cgm")
testLeadHistorySyncOneCampaign("lymphc")
testLeadHistorySyncOneCampaign("lymphw")
logLeadHistorySummary()                 // state counts per campaign
logTableSizeComparison()                // confirm two-tab arch is working
testCleanupAgedOutLeadStates()          // run 30-day prune manually
runAllBackfills()                       // Stage 2, NOT YET RUN
resetBackfillProgress()                 // clear backfill progress markers
resetCampaignBackfill()                 // (Code.gs) unrelated legacy backfill
```

`runDailyLeadHistorySync()` is what Code.gs's `eodArchive()` trigger calls
automatically — wrapped in try/catch so a Lead History failure never blocks
the existing daily chaser/campaign archive.

---

## Immediate next steps (in order)

1. Clear `Lead History`, `Current Lead State`, `Lead History Conflicts`
   tabs in the archive spreadsheet (leftover test data uses old schemas).
2. Run `testLeadHistorySyncOneCampaign("ort")`, verify output, especially:
   - No "Unrecognized status" warnings
   - No duplicate SUPERSEDED rows for the same lifecycle
   - Terminal lifecycles (Approved etc.) never marked SUPERSEDED
3. Repeat for `"cgm"`, `"lymphc"`, `"lymphw"` — LymphW is highest risk for
   new unclassified BTO/DIS/FROZEN phrases.
4. Verify Verbal Denial tabs are actually being read (check
   `logLeadHistorySummary()` shows nonzero VerbalDenial counts if any
   exist in the source sheets).
5. Once daily sync is verified solid across all 4 campaigns, run
   `runAllBackfills()` for Stage 2.
6. Build dashboard UI for Lead History Conflicts (backend functions ready:
   `getPendingLeadConflictsForDashboard()`, `resolveLeadConflict()`).
7. Build the Weekly Summary tab (backfill from the numbers-only historical
   sheet once access is granted — user was still requesting access as of
   this conversation — then derive going forward from Lead History).
8. Once Lead History is proven reliable, swap "Cases" for a deduplicated
   "Leads" metric in dashboard.html (History, Overview, Compare tabs) and
   add Verbal Denials as a visible Campaigns tab line item.
