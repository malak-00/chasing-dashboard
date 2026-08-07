# Changelog

Every update made to this app, most recent first. This is maintained after
every change — read it before starting new work instead of re-deriving
context from scratch, and add a new entry at the top after every change
going forward (code or otherwise).

Each entry: what changed, why, and what it touches. PR numbers refer to
`malak-00/chasing-dashboard` on GitHub.

---

## PR #26 — Campaign pipeline rewrite, presentable archive sheet, executive UI overhaul (2026-08-06 – 2026-08-07)

### Whole-number Total Duration (twelfth batch)
`durationMins` is a sum of per-day call-duration minutes, and the
underlying values can be fractional (Utlatel's own "H,M" duration text
parses to e.g. 27.13 mins), so an unrounded sum read as an ugly decimal.
`LB_COLS`' "Total Duration (Min)" leaderboard column had no format
function at all (showed the raw summed value) -- added a new
`formatWholeMinutes()` and wired it in. The Controls tab's Utlatel Archive
Summary table used `.toFixed(1)` explicitly; switched to `Math.round()`.
`applyMonthTabFormatting()` (Code.gs) now also sets a whole-number `'0'`
format on the archive sheet's `TotalDurationMins` column, matching the
dashboard's display. Run `reformatArchiveTabs()` once to apply the new
column format to existing month tabs.

### Fixed Productivity/Efficiency over 100% showing as e.g. "1.8%" in the archive (eleventh batch)
Follow-up to the previous batch's percent-scale fix, reported immediately
after: that fix guessed "a real percent-formatted cell's value is always
`<=1`" to tell it apart from a plain number already on the 0-100 scale --
which breaks the moment the true percentage exceeds 100%. A percent-
formatted cell showing "180%" has an underlying value of `1.8`, which is
`>1`, so the old logic treated it as "already correct" and left it as
`1.8` -- written into the archive and displayed as "1.8%" via its percent
number format. Value magnitude alone can never reliably disambiguate this
(`1.8` could mean "180%, percent-formatted" or "literally 1.8, a plain
number" -- both are valid raw `getValues()` outputs).

Switched to checking the cell's ACTUAL number format via
`getNumberFormats()` instead of guessing from magnitude: if the format
contains a literal "%", the value is a fraction and gets multiplied by
100 regardless of how large it is; otherwise it's taken as-is. This is
unambiguous in every case. `collectRowsFromWeekTab()` now reads
`getNumberFormats()` alongside `getValues()` (same shape/position) and
threads the relevant cell's format through to `parsePercentField()` for
both Efficiency and Productivity.

Verified with 8 unit cases (including the exact reported 180%-showing-as-
1.8% scenario) and a full `migrateExistingSheets()` integration test
confirming a 180%-percent-formatted source cell writes `180` (not `1.8`)
into the real archive row.

### Fixed Productivity/Efficiency % scale bug + stopped capping the Overview ranking bar at 100% (tenth batch)
Two data-display bugs reported directly:
- **Productivity/Efficiency showing e.g. "0.4%" instead of "38.4%"**:
  `collectRowsFromWeekTab()` (the historical migration path) only ever
  stripped a literal "%" suffix from text cells. A source cell that's a
  real Sheets-percentage-formatted number (Format > Number > Percent)
  comes back from `getValues()` as the raw underlying fraction (0.384),
  not the string "38.4%" -- that fraction was written straight into the
  archive on the wrong scale, and once the archive's own percent number
  format was applied it displayed as "0.4%". New `parsePercentField()`
  handles both cases (verified against 7 cases including 0%/100% edges).
  Only affected `migrateExistingSheets()`-written data -- the daily Sync
  path already multiplies by 100 itself and was never affected.
- **Overview ranking bar capped at 100%**: Productivity can legitimately
  exceed 100% (e.g. ACW-inflated productive time vs. nominal shift
  minutes), but the bar's width was hard-capped via `Math.min(pct,100)`
  while the adjacent number was never capped -- a standout performer at
  150% correctly showed "150.0%" but their bar looked identical to someone
  at exactly 100%. Fixed by scaling every bar to the highest value actually
  present today (floored at 100, so the normal everyone-under-100% case is
  visually unchanged) instead of a hard cap -- simply removing the
  `Math.min()` alone wouldn't have worked, since `.ov-rank-bar-bg` has
  `overflow:hidden` and a >100% width would just get silently clipped at
  the container edge, visually indistinguishable from the original bug.
  Verified via Playwright: a chaser at 150% fills the bar (today's max)
  while one at 80% renders proportionally shorter (53.3% width) instead of
  both looking identically capped.

### Full names as the canonical chaser identity everywhere (ninth batch)
Requested after noticing the archive only ever showed first names.
`CHASER_SHEETS`/`CHASER_NAME_MAP` (Code.gs) and `CHASERS`/`CHASER_COLORS`
(dashboard.html) switched from short names to full names ("Alex Woods"
instead of "Alex") for every chaser, active and former -- matching the
"Tom Walker" convention already used for a newer chaser, now applied
consistently everywhere instead of as a one-off exception.

Short names turned out to be a real matching KEY in a few places, not
just display text, so this needed more than a find-and-replace:
- `buildUtlatelLookup()` matched Utlatel's raw "Agent" text as a substring
  of the chaser name -- a full-name substring match would silently stop
  matching an Agent value that's just "Alex" (shorter than "Alex Woods"),
  breaking duration/calls credit for every renamed chaser going forward.
  Now matches on just the first-name token, robust to both old short-name
  Utlatel exports and any future full-name text.
- The exact same substring-match issue existed twice on the frontend
  (`buildAgentMapUI`, `applyUtlatelPersist`, both auto-detecting which
  chaser a raw Utlatel agent string refers to) -- consolidated both into
  one `autoMapAgentToChaser()` helper with the same first-name-token fix,
  removing a pre-existing duplication in the process.
- Found and deleted a dead, already-out-of-sync duplicate of
  `CHASER_NAME_MAP`/`normalizeName()` sitting in dashboard.html -- never
  called anywhere, and missing the "tom" entry Code.gs's real copy already
  had (a drift that had already happened once with only 2 copies of the
  same map).
- Found and fixed a stale hardcoded 5-name `<option>` list in the Custom
  Chart Builder's chaser multi-select (never included Tom Walker even
  before this change, would now be actively wrong) -- populated
  dynamically from `CHASERS` like the app's other chaser dropdowns.
- Found and fixed a real width-clipping bug via a Playwright screenshot:
  `.ov-rank-name` (Overview tab's ranking rows) was sized for short names
  and truncated "Frank Clarkson" to "Frank Clar…" -- widened 90px -> 130px.

Added a new one-time `renameChasersToFullNames()` (Code.gs) to bring
already-written data in line: renames the "Chasers" roster tab's Name
column and every archive month tab's Chaser column (exact match against a
`CHASER_RENAME_MAP`, batched per-tab, idempotent -- safe to re-run),
invalidates the roster/archive caches afterward. **Needs to be run once**
after this deploys, or the roster/archive will still show short names
until it's run.

Verified via a Node `vm`-sandboxed test of `renameChasersToFullNames()`
(roster + archive rename, WEEK/TOTAL rows correctly left alone, idempotent
re-run, cache invalidation) and Playwright (`CHASER_COLORS`/
`autoMapAgentToChaser` resolve correctly for full names given short-name
and short-name-with-extension inputs, `cbChasers` populates dynamically,
zero console errors, no truncated names on the Overview tab after the
width fix).

### Day-separator borders on the historical rebuild path + a denials-flag rule (eighth batch)
Prompted by rebuilding the archive from scratch (delete month tabs +
Campaign Responses, re-run `migrateExistingSheets()` +
`backfillFromCombinedSheet()`) and finding the rebuilt sheet had no visual
separation between days:
- `migrateExistingSheets()` now adds the same bottom teal border under each
  day's last row that `archiveDayData()` (the daily Sync write) already
  did -- it was the one write path missing this, so a full historical
  rebuild left every day's block looking identical to the next.
- Fixed a real bug in `applyDayBordersToArchive()` (the retroactive
  one-time border utility, for anyone with existing un-bordered data):
  it compared raw `String(dateCell)` values, which mis-groups a day
  whenever Sheets silently auto-coerced only some of that day's rows into
  real Date objects -- switched to `normalizeDateCellToTab()`, the same
  helper every other date comparison in this file already uses for exactly
  this reason.
- `applyMonthTabFormatting()` now also adds a conditional-format rule that
  lightly highlights any row where Denials > Approvals, so a day/chaser
  needing attention is visible while scanning. Deliberately a formatting
  rule and not a sortable Filter -- a Filter would let someone sort a
  column and scramble the manually-built week-header/day-border structure
  this whole feature exists to create.
- Extracted `columnLetter()` out of `forcePlainTextColumns()` into its own
  reusable helper.
- Considered and deliberately did NOT add: a basic Filter/sort control
  (risks destroying the manual row structure, see above); a per-day
  subtotal row (bigger scope, would need a decision on what should be
  summed -- worth a follow-up if wanted); per-week alternating shading
  instead of per-row banding (marginal benefit once day borders exist).

Verified via 3 Node `vm`-sandboxed tests: `migrateExistingSheets()`
borders land on the correct last-row-of-each-day (including the
end-of-run flush for the final date), `applyDayBordersToArchive()`
correctly groups a day whose rows are a mix of Date-object and string
date cells (previously mis-grouped this), and the conditional format rule
is added once and not duplicated on repeat `applyMonthTabFormatting()`
calls (important since it's re-applied by `reformatArchiveTabs()`).

### Added 3 new campaigns: PPO (ORT), PPO (LY), UTI (seventh batch)
Extended the campaign roster from 4 to 7 by adding `PPO (ORT)`, `PPO (LY)`,
and `UTI` to `CAMPAIGN_LABELS`, plus 3 matching entries in
`BACKFILL_RESPONSES_TABS` pointing at their real tabs in the "Overall 2026"
spreadsheet (`"PPO (ORT)"`, `"PPO (LY)"`, `"UTI Overall 2026"` — the first
two have no `" Overall 2026"` suffix, unlike the other 5 tabs), all reading
Chaser Name/Status/Date of conclusion exactly like the existing 4.

This required refactoring the archive schema itself rather than just adding
config, since the old code hardcoded "8 campaign columns at indices 14-21"
in 3 different places:
- `ARCHIVE_HEADERS` is now built by concatenating a fixed tracker-owned
  block with one Approved/Denied column pair per entry in the new
  `CAMPAIGN_KEYS` (derived from `CAMPAIGN_LABELS`), via a new
  `CAMPAIGN_COL_PREFIX` map that keeps column header text clean
  (`PPOORT_Approved` etc.) independent of the display label used in the UI
  (`"PPO (ORT)"`, which has spaces/parens). The first 4 campaigns' column
  text is unchanged from before this refactor.
- `updateChaserCampaignColumnsForDate()`, `archiveDayData()`'s new-row
  default, and `migrateExistingSheets()`'s campaign-columns-default now all
  loop `CAMPAIGN_KEYS` instead of writing 8 literal array slots/zeros, so
  adding another campaign in the future is one line in `CAMPAIGN_LABELS` +
  `CAMPAIGN_COL_PREFIX` + `BACKFILL_RESPONSES_TABS`, nothing else.
- New `padRowToArchiveWidth()` extends any row shorter than the current
  `ARCHIVE_HEADERS.length` (i.e. every row written before this change, at
  22 columns) out to the new 28-column width with zeros before it's reused
  in a `setValues()` rewrite -- without this, the very next Sync or weekly
  pull touching an old row would throw a range-width-mismatch error the
  first time it tried to write 28 values into what Sheets still measured as
  a 22-column range.
- New one-time `addNewCampaignColumnsToExistingTabs()` (companion to the
  existing `reformatArchiveTabs()`) backfills just the 6 missing header
  cells + column widths/formatting onto every month tab that predates this
  change, so the header row matches the data going forward. Data rows need
  no separate migration -- they're padded automatically the next time
  anything touches them.

Frontend: added the 3 campaigns to `CAMPAIGNS` and `CAMPAIGN_ARCHIVE_COLS`
(dashboard.html), plus matching entries in `campLabelToKey()`. Added 3 new
CSS color tokens (`--violet`, `--orange`, `--slate`, with light-mode
variants) distinct from the existing campaign colors and from
`--green`/`--red` (reserved for status). Found and fixed a genuine
redundancy while doing this: `renderChaserCampaigns()` (chaser detail tab)
had its own third independently-hardcoded 4-campaign list instead of
reusing the shared `CAMPAIGNS`/`CAMPAIGN_ARCHIVE_COLS` -- consolidated so
it picks up new campaigns automatically. Every other campaign-rendering
spot (Overview snapshot, Compare tab totals, CSV export, weekly trend
chart) already looped `CAMPAIGNS`/`CAMPAIGN_KEYS` generically and needed no
changes.

Also found and fixed a real layout bug introduced mid-change: widening
`.campaign-grid`/`.ov-camp-grid` from a fixed 2-column layout to a
responsive `auto-fill` (needed so 7 cards don't leave an awkward orphaned
card in a 4-row 2-column layout) shrank card width enough that the
existing 4-stats-in-a-flex-row `.campaign-stats` layout started clipping/
overlapping text. Changed `.campaign-stats` to a 2x2 grid instead, which
stays legible at any card width instead of only the one width the old
fixed 2-column grid happened to produce.

Verified end-to-end: a Node `vm`-sandboxed mock of `SpreadsheetApp`/
`PropertiesService` confirmed `ARCHIVE_HEADERS` is exactly 28 columns,
`updateChaserCampaignColumnsForDate()` correctly pads and rewrites a
legacy 22-column row without a range-mismatch error, new campaign values
land in the correct columns, and both the update and insert code paths
work correctly together. Playwright confirmed all 7 campaign cards render
without overlap on the Overview tab, Campaigns tab, and a chaser's
Campaign Breakdown panel (the last one specifically exercising the
consolidated `renderChaserCampaigns()`), with zero console errors.
`node --check` clean on both files; grepped for and fixed every remaining
comment that hardcoded "4 campaigns"/"~4 tabs" from before this change.

### Dead-code cleanup from the Code.gs review (sixth batch)
Removed everything flagged as genuinely redundant/unreachable in a full
read-through of Code.gs — see the findings list from that review (in
conversation, not repeated here since none of it changed behavior):
`clearCache()` fixed to actually clear the 4 real cache keys instead of 2
that nothing ever wrote to; deleted `renameArchiveFaxColumn_DEPRECATED()`
(dead, referenced a column that no longer exists); simplified
`APPROVAL_PATTERNS`/`DENIAL_PATTERNS` (2 entries were fully subsumed by
another entry via substring matching, e.g. `"DENIAL"` already catches
`"RECEIVED DENIAL"`); pruned 7 unused `COL_MAP` entries left over from a
since-removed Faxes archive column; removed a no-op empty `if` block in
`archiveDayData()`; and pointed `getAvailableTabs()`/`testAllChasers()`/
the renamed `debugChaser(name)` (was `debugAlex()`) at the current active
roster instead of the frozen `CHASER_SHEETS` seed, so they reflect chasers
added or repointed purely through the dashboard's Settings tab. Verified
with `node --check` and by re-running the mocked-Sheets test from the
`dryRunCampaignBackfill()` commit — identical output before/after.

### `dryRunCampaignBackfill()` — read-only preview before running the real backfills (fifth batch)
Added a new dry-run test function alongside the existing `dryRunMigration()`,
for the two functions that actually write data (`backfillCampaignColumns()`,
`backfillFromCombinedSheet()`): calls the same `parseBackfillResponses()`
every real backfill uses, then reports the date range found, per-campaign
approved/denied totals, every distinct chaser name credited (flagging any
that didn't resolve to a known canonical name via `normalizeChaserName()`
-- catches a typo'd/unmapped name before it silently creates a brand-new
"chaser"), and a simulated updated-vs-inserted row count by re-running
`updateChaserCampaignColumnsForDate()`'s own matching logic against the
archive read-only (never calling `setValues()`/`appendRow()`). Also reports
both real functions' `PropertiesService` progress-key state so you know
what a real run would actually still touch. Verified with a mocked
Sheets/PropertiesService environment in Node (`vm` sandbox) against a small
synthetic dataset -- confirmed it correctly distinguishes an already-
archived date (counted as "would update") from one with no existing row
("would insert") and correctly flags an unmapped chaser name.

### Card depth, fonts, header decluttering, bar animation (fourth batch)
Follow-up visual pass on top of the Overview/status-color work below, based
on the same "make it look like a modern exec dashboard" research:
- **Fonts**: only static weights 300/400/500/600/700 were loaded from
  Google Fonts, so the 7 existing `font-weight:650` rules (a non-loaded
  weight) would snap unpredictably to whichever static weight the browser
  picked as "nearest" — inconsistent across browsers. Switched to loading
  Inter as a variable font (`wght@300..800` range syntax) so every weight
  in between renders as requested.
- **Real shadows/depth**: `--shadow-sm`/`--shadow-md` were tuned for a
  light background and read as almost invisible against the dark theme's
  near-black surfaces. Strengthened both, added a new `--shadow-lg` for
  hover-lifted cards, and a `--card-elevated` surface-color step (a shade
  lighter than `--card`) so hover/elevation reads as an actual lift, not
  just a barely-there outline change. Same treatment mirrored for light
  mode's `body.light` block.
- **Card interactivity**: `.metric`, `.campaign-card`, `.ov-camp-card`, and
  `.ov-rank-row` now lift (`translateY(-2px)` + `--shadow-md`/`--card`) on
  hover. Deliberately scoped to card-level elements only — large panel
  containers and table rows keep their existing (non-lift) treatment,
  where a hover-lift would look wrong.
- **Header decluttering**: the header was accumulating buttons (Load,
  Sync, Pull Weekly, CSV, PDF, theme toggle) competing for space. Moved
  the theme toggle and CSV/PDF export — all occasional-use actions — into
  a single `⋮` overflow menu (closes on outside click or Escape,
  `aria-expanded`/`role="menu"` for accessibility), leaving the persistent
  toolbar down to what's used every day: date picker, Load, and the
  Sync/Pull-Weekly group. Also removed a now-stale `.theme-btn` reference
  in the `@media print` block (that button no longer exists in the
  header) and pointed it at `.header-menu-panel` instead, so PDF export
  still correctly hides the menu button/panel from print output.
- **Overview bar animation**: the ranking bars' width was baked directly
  into the initial `innerHTML`, so the existing `transition:width .4s
  ease` CSS never actually animated on first render (only on later
  re-renders after data changed). Bars now render at `width:0%` with the
  real value in a `data-target-width` attribute, then a new
  `animateRankBars()` sets the real width one double-`requestAnimationFrame`
  later so the transition has something to animate from. Verified via
  Playwright that the bars are synchronously at 0% immediately after
  `renderOverview()` and settle at their correct target widths shortly
  after.
- Verified via Playwright: header menu opens/closes/closes-on-outside-
  click/closes-on-Escape, theme toggle updates the menu's icon+label and
  auto-closes the menu, hover-elevation visible on a metric card in both
  themes, bars animate from 0% as described above. `node --check` clean
  on both `Code.gs` and the extracted `dashboard.html` script contents;
  grepped for stray `.theme-btn`/`themeBtn` references (none left).

### Executive UI/UX overhaul (third batch of commits on this PR)
Researched 2026 KPI/executive dashboard best practices (decision-first
layout, color reserved for status only, no-scroll primary view,
progressive disclosure) and applied them:
- **New Overview tab**, now the default landing page — status-colored KPI
  cards against configured Productivity/Efficiency targets (not bare
  numbers), a ranked team-comparison bar strip for today's Productivity,
  a campaign approval-rate snapshot for the week, and drill-down links
  into Leaderboard/Compare/History. Reuses `buildRowFromTotals()` (the
  Leaderboard's own row-building logic) so the numbers never disagree
  between the two. Verified via Playwright with injected mock data —
  status colors, sort order, and the "no Utlatel data" stub all render
  correctly, and it fits one 1440x900 screen with no scroll.
- **Consolidated sync controls** — the header's Sync button and the
  Controls tab's separate "Pull Weekly Responses" button are now one
  grouped header control: `[date picker] [Sync] | [week picker] [Pull
  Weekly]`. The week picker is functional: `pullWeeklyCampaignData()`
  takes an optional anchor date, and `mode=weeklycampaignsync` accepts an
  ISO week param (`"YYYY-Www"`, the native `<input type="week">` value
  format) that anchors the same self-correcting 14-day trailing window to
  that week's Sunday instead of today. Verified the ISO week math
  round-trips correctly against real calendar dates.
- **Status-colored Productivity/Efficiency everywhere** — new
  `statusColorForMetric()`/`statusPill()` helpers applied to Leaderboard
  and History (previously bare numbers or an always-teal pill regardless
  of value). Consolidated the Chaser detail tab's pre-existing
  `getGoalStatus()` (same met/close/miss thresholds, independently
  duplicated) to delegate to the shared helper instead.
- **Two real bugs found and fixed during verification**: the Compare
  tab's per-chaser change table showed Denials increasing as a green "up"
  change (the sibling panel right next to it already had an
  invert-for-denials rule, this table just never got it); and
  `getCachedOrFetch()` silently stopped caching once a payload crossed
  ~90KB (CacheService's per-key limit) with no fallback, meaning every
  dashboard load would eventually re-scan every month tab from scratch as
  the archive grows, with no visible sign caching had stopped working —
  now chunks large payloads across multiple keys, with a new
  `invalidateCache()` so the 6 existing cache-bust call sites clear
  chunked entries too instead of leaving stale ones being served.

### Presentable archive sheet + migration fixes (second batch)

### Campaign-response pipeline (Status+conclusion only, single-pass parse)
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

### Presentable archive sheet + migration fixes
People other than dashboard users look directly at the raw archive
spreadsheet, so:
- **Neater sheet**: month tabs now get real header styling (matching the
  dark navy/teal look already used for week-separator rows), Date/Chaser
  columns frozen, sized column widths, alternating row banding that
  survives future appended rows, and Efficiency/Productivity displayed
  with a "%" suffix (display-only, stored values unchanged). Shared via
  `applyMonthTabFormatting()` between `getOrCreateMonthTab()` (new tabs)
  and a new `reformatArchiveTabs()` (one-time pass over existing tabs, no
  data touched).
- **Migration write order**: `migrateExistingSheets()` processed Team 1's
  entire spreadsheet before starting Team 2 at all, and always
  `appendRow()`'d — every Team 2 row landed after every Team 1 row in
  each shared month tab, requiring a manual re-sort afterward.
  `parseWeekTab()` is now `collectRowsFromWeekTab()`, a pure parser with
  no archive writes; `migrateExistingSheets()` collects every row from
  both spreadsheets first, sorts chronologically, then writes in that
  order — both teams now interleave correctly by date regardless of
  which spreadsheet gets read first. Also normalizes the chaser name at
  collection time, so the archive shows clean canonical names instead of
  raw tracker text.
- **"Total" rows**: the source week tabs' own weekly "Total" summary row
  had no skip-check (unlike the archive-reading side, which already skips
  names starting with "TOTAL") — it was getting written into the archive
  as a fake chaser row, permanently stuck at 0 for Approvals/Denials/
  campaign columns since the campaign backfills explicitly skip TOTAL
  rows. Now skipped at collection time.

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
