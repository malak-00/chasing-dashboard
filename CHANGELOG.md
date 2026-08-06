# Changelog

Every update made to this app, most recent first. This is maintained after
every change — read it before starting new work instead of re-deriving
context from scratch, and add a new entry at the top after every change
going forward (code or otherwise).

Each entry: what changed, why, and what it touches. PR numbers refer to
`malak-00/chasing-dashboard` on GitHub.

---

## PR #26 — Rewrite campaign-response pipeline: Status+conclusion only, single-pass parse (2026-08-06)
Three changes to the campaign/backfill-response reading layer:
- **Name variants**: `normalizeChaserName()` only did an exact lowercase
  match, so `"ALEX WOODS"` matched fine but `"Alex Woods."` (trailing
  period) or irregular spacing didn't. Now strips periods and collapses
  whitespace before matching. Added `"tom" -> "Tom Walker"` to
  `CHASER_NAME_MAP` — first-name-only entries already worked for every
  other chaser since their canonical form already *is* the first name.
- **Status + Date of conclusion only, all 4 campaigns**: ORT, CGM, LY
  PUMP, and LY WRAP now all carry Status and Date of conclusion columns.
  Approval/denial comes exclusively from the Status column; the date
  exclusively from Date of conclusion. Fax-feedback text is no longer
  read anywhere in this pipeline — a row missing a conclusion date is
  skipped, never guessed at. Widened `DENIAL_PATTERNS` to catch a plain
  "Denied"/"Denial" (was tuned for free-text feedback phrasing like
  "RECEIVED DENIAL"). Dropped `feedbackCol`/`idCol`/`idnCol`/
  `submissionCol` from `BACKFILL_RESPONSES_TABS` — dead fields left over
  from the now-deleted `LeadHistory.gs`.
- **Efficiency**: `readChaserTotalsFromBackfillSource()`,
  `collectDatesFromCombinedSheet()`, and `readCombinedCampaignTotalsForDate()`
  each independently re-scanned all 4 backfill tabs — called once per
  date, the weekly pull's 14-day loop alone triggered ~28 full re-scans
  per click. Replaced with `parseBackfillResponses()`, which reads the 4
  tabs **once** and buckets every row by resolved date; every entry point
  (`pullWeeklyCampaignData`, `backfillCampaignColumns`,
  `backfillFromCombinedSheet`) now parses once and threads the bucket
  through instead of re-scanning per date.
- Bumped both backfill progress keys (`campaign_backfill_progress5->6`,
  `combined_sheet_backfill4->5`) since the underlying matching logic
  changed — old "done" dates need reprocessing under the new
  Status-column-based reading, not silent skipping.

## PR #25 — Fix migrateExistingSheets() duplicate-check key format mismatch (2026-07-21)
Root cause of `migrateExistingSheets()` looking "stuck" and writing
duplicates on re-run: its duplicate guard built keys as year-less `"M/D"`
(via `normalizeDateCellToTab()`), but the actual lookup in `parseWeekTab()`
checks full `"M/D/YYYY"` keys. The two formats never matched, so the dedup
check was silently a no-op — every re-run rewrote everything as fresh
duplicates. Fixed by deriving each existing row's year from its own month
tab's name when building the guard set, so both sides compare the same
format.

## PR #24 — Make backfillFromCombinedSheet() recreate deleted month tabs (2026-07-21)
Ahead of a planned full rebuild of the archive (delete month tabs, re-run
`migrateExistingSheets()` + this backfill): the per-chaser column step
previously used `writeCombinedChaserCampaignColumns()`, which required the
month tab to already exist and silently skipped it otherwise. Swapped in
`updateChaserCampaignColumnsForDate()` (already shared by
`backfillCampaignColumns()`/`pullWeeklyCampaignData()`), which creates the
tab via `getOrCreateMonthTab()` if missing and inserts rows for chasers
with credit but no existing row. Removed the now-superseded
`writeCombinedChaserCampaignColumns()`/`readCombinedChaserCampaignCountsForDate()`.

## PR #23 — Remove Lead History, Overview tab, Conflicts tab; support new tracker naming (2026-07-21)
Second half of a two-part architecture change:
- Deleted `LeadHistory.gs` entirely (moved `forcePlainTextColumns()` and
  `BACKFILL_RESPONSES_SHEET_ID`/`BACKFILL_RESPONSES_TABS` into `Code.gs`
  first, since those are part of Code.gs's own campaign pipeline despite
  living in that file). Removed the `leadsnapshot`/`leadactivity`/
  `leadconflicts`/`resolveconflict` doGet modes and the
  `runDailyLeadHistorySync()` call inside `mode=sync`.
- Removed the Overview tab (nav button, header's "Controls" toggle,
  legacy Utlatel upload handlers, `buildShiftGrid`/`setShift`,
  `toggleControls`/`renderGoalsControls`) and the Conflicts tab (nav
  button, panel, all "Lead History Conflicts" JS). Caught and fixed two
  regressions this surfaced: `buildAgentMapUI()` and
  `applyUtlatelPersist()` were unconditionally targeting Overview's DOM
  elements from the *shared* Utlatel pipeline (also used by the Controls
  tab) — added an equivalent agent-map section to the Controls tab
  instead of losing the feature. `campaigns` is now the default tab.
- `readChaserTab()` now looks up tracker tabs by `"M/D/YY"`/`"M/D/YYYY"`
  (padded or not) via a new `findChaserTrackerSheet()` helper, instead of
  the old `"M/D."` dotted format — falls back to the dotted format last,
  for any tab not yet renamed.

## PR #23 (first commit) — Restructure sync architecture: manual daily/weekly buttons, new Leaderboard (2026-07-14 → merged 2026-07-21)
Major architecture change: read chaser trackers daily (manual button, not
automatic), read campaign responses from the backfill sheet weekly for
every chaser (manual button), keep all existing Settings controls.
- **Retired `RESPONSE_SOURCES` entirely** (removed `readResponsesForDate`,
  `parseResponseTab`, `feedbackMatchesDate`, `readCampaignTotalsForDate`,
  `readChaserCampaignCountsForDate`, `archiveCampaignResponses`,
  `backfillCampaignResponses`, and the dead `backfillDateRange_DEPRECATED()`)
  — that live fax-sheet source had repeated column-mismatch/stale-data
  problems; campaign data now comes exclusively from the backfill sheet.
- `archiveDayData()` (daily button) now **only** writes tracker-derived
  columns (Cases/Positive/TimeMins/Efficiency/Productivity/Shift/Calls/
  Duration/ACW/ProductiveTime); Approvals/Denials/campaign columns are
  preserved untouched on existing rows.
- New `updateChaserCampaignColumnsForDate(fullDate)`, shared by
  `backfillCampaignColumns()` (historical) and the new
  `pullWeeklyCampaignData()` (manual weekly pull, trailing 14 days, no
  progress tracking — reprocessed fully every click).
- Removed the automatic `eodArchive` time trigger and `setupEodTrigger()`;
  added `removeEodTrigger()` for one-time live cleanup. `mode=sync` now
  also runs Lead History sync (previously trigger-only). New
  `mode=weeklycampaignsync` drives the weekly pull.
- **Leaderboard redesigned** to: Chaser Name, Total Shift, Total Calls,
  Total Duration, Total Duration (Min), ACW Duration, Productive Time,
  Total Chased Cases, Total Positive, Total Time Taken, Approvals,
  Denials, Total Responses, Productivity, Efficiency — dropping
  Approval%/Denial% (kept for History tab/CSV export). Table markup now
  driven generically by a column array.
- New "Pull Weekly Responses" button in the Controls tab.

## PR #22 — Use backfill tab's own year, not Submission Date, as the default year (2026-07-14)
Follow-up on #21: defaulting a year-less feedback date to Submission Date's
year was wrong — most leads submit in one year but resolve well into the
next, so Submission Date mislabeled most rows. Backfill tabs are literally
named for their year (`"ORT Overall 2026"`), same pattern as the archive's
month tabs, so that's the correct default now (`resolveYearFromTabName()`),
only overridden when a feedback token explicitly spells out a different
year. `readChaserTotalsFromBackfillSource()` got the same conclusion-date
matching CGM/ORT already had.

## PR #21 — Use ORT/CGM's "Date of conclusion" column as the authoritative resolution date (2026-07-14)
ORT/CGM carry a "Date of conclusion" column with the real resolution date
(including year) — introduced `resolveRowDate()` to use it directly when
present, bypassing feedback-text date parsing (and its year ambiguity)
entirely for those two campaigns. LY PUMP/LY WRAP (no such column) still
fell back to feedback-text parsing at this point. Also fixed year-less
dates written by `backfillFromCombinedSheet()`.

## PR #20 — Fix CGM chaser column name in backfill config, bump backfill progress key (2026-07-13)
CGM's backfill tab uses a different chaser column name than ORT; the
shared config had it wrong, silently dropping CGM per-chaser credit.
Bumped progress key so previously-"done" dates get reprocessed correctly.

## PR #19 — Remove BACKFILL_RESPONSES_SHEET_ID/TABS redeclaration, already in LeadHistory.gs (2026-07-13)
Cleanup — Code.gs and LeadHistory.gs shared one Apps Script global scope,
so a duplicate declaration was pure risk (whichever file's copy got out of
sync would silently override the other).

## PR #18 — Dedupe combined-sheet backfill config onto BACKFILL_RESPONSES_SHEET_ID (2026-07-13)
Further backfill-source consolidation, same theme as #17.

## PR #17 — Extend campaign backfill / don't lose credit when tracker tab missing (2026-07-13)
`backfillCampaignColumns()` extended to insert rows for chasers with real
campaign credit but no existing archive row (no tracker tab that day) —
previously that credit was silently dropped rather than just deferred.

## PR #16 — Fix backfill year-collision bug and a scoping crash (2026-07-13)
`backfillCampaignResponses()`/`backfillCampaignColumns()` were discarding
each date's real year in places, causing same-M/D-different-year
collisions. Established the pattern (still in use) of always trusting the
month tab's own name for the year, never a cell's own auto-coerced value.

## PR #15 — Fix feedbackMatchesDate double-counting rows with a stray date in notes (2026-07-13)
A feedback string mentioning a second, incidental date in free text (e.g.
`"Approved+CN 6/1 / ... {Logan 6/2}"`) was getting counted on both dates.
Fixed to only match the *first* date token in the string as the row's
actual status date. (This whole class of bug later became structurally
impossible for ORT/CGM once #21/#26 moved those off feedback-text parsing
entirely.)

## PR #14 — Fix Campaign Responses date-year bug breaking Month/Quarter views (2026-07-07)
A year-less date written to the Campaign Responses tab got auto-coerced by
Sheets into a Date with a guessed year; Month/Quarter views (which filter
by real Date ranges) silently miscategorized or dropped those rows. Day/
Week never showed the problem since they only ever compare year-less
`"M/D"`.

## PR #13 — Leaderboard reads shift/productivity from archive; team-relative Approval%/Denial% (2026-07-07)
`TotalShift`/`ProductiveTime`/`Productivity` now computed and written to
the archive at sync time rather than recomputed live from session sliders
— Leaderboard reads the persisted values so it matches Code.gs's own
formula and doesn't drift from an un-saved slider nudge. Approval%/Denial%
changed from "% of this chaser's own responses" to "this chaser's share of
the team's total responses for the range."

## PR #12 — archiveDayData updates existing rows in place; write Utlatel totals to archive (2026-07-07)
Re-syncing an already-archived date used to skip existing rows entirely
(stale numbers forever); now updates in place. `TotalCalls`/
`TotalDurationMins` from Utlatel now written into the archive row itself
instead of only ever existing as a client-side join at render time.

## PR #11 — Fix Utlatel date format mismatch so uploaded data joins correctly (2026-07-07)
Upload dates and archive dates were in incompatible formats, silently
failing to join.

## PR #10 — Utlatel upload improvements, PDF export fix (2026-07-06)
Date picker for Utlatel uploads without their own Date column; prefer the
Extension column over Agent Name for chaser matching (more reliable);
PDF export now prints whichever tab is currently active instead of always
forcing Overview.

## PR #9 — Remove History Total Cases card, add date range to Chaser tab (2026-07-06)
Also removed early "Leads" stat displays (later reintroduced properly via
Lead History, then removed again in PR #23 when Lead History was retired).

## PR #8 — 90s request timeout, shift-minutes default fix, drop Campaigns Leads stat (2026-07-06)

## PR #7 — Controls-tab Utlatel re-upload merges instead of replaces; fix ACW formula (2026-07-06)
ACW duration was multiplying by cases instead of calls.

## PR #6 — Fix Sync button stuck on "Syncing...", debug endpoint, sheet ID updates (2026-07-06)
Added `mode=debug` (now `getDebugInfoForDate()`) so sync issues could be
inspected without Apps Script Execution Log access.

## PR #5 — Fix broken sync timeout wiring (2026-07-06)

## PR #4 — New color palette; redefine Lead snapshot "Leads" count (2026-07-06)
Rich black/off-white base, dark teal + deep ruby accents. Leads counted by
SubmissionDate instead of TransitionDate.

## PR #3 — Add Conflicts tab, per-campaign Lead snapshots, Leads metric (2026-07-05)
First version of the Conflicts tab and the Lead History-backed "Leads"
metric (both later removed in PR #23 once Lead History was retired as
unneeded).

## PR #2 — Fix duplicate Current Lead State rows on same-batch birth+supersede (2026-07-05)

## PR #1 — Add editable chaser roster; harden dashboard saves (2026-07-05)
Chaser roster became data-driven (Chasers tab in Settings + `getChasersConfig()`/
`saveChaser()`) instead of the hardcoded `CHASER_SHEETS` map. Deduped
campaign/chaser lists that had been hand-copied in multiple places.
Also folded in an earlier security/correctness review: fixed a stored XSS
in `buildAgentMapUI`, a chronology-unsafe SUPERSEDED-marking bug in Lead
History, missing-tab alerting, and a date-coercion guard gap.

## Initial build (2026-06-25 → 2026-07-05, pre-review-cycle)
`dashboard.html` created and rapidly iterated (chart builder, campaign
data handling, sync functionality, date parsing, styling) before
`Code.gs`/`LeadHistory.gs` existed as separate reviewable PRs. `Code.gs`
and `LeadHistory.gs` added 2026-07-05, building the two-tab Lead History
architecture (permanent event log + fast-lookup current-state table),
conflict detection (Blocking/Review), and the daily chaser/campaign
archive this app was originally built around.
