# Prioritization

## Goal

Every quarter, collect data from six (6) centres, submitted by each centre's
Centre Lead. The data must capture:

1. What priorities each centre is working on.
2. How the centre ranks those priorities.
3. How many resources are working on each priority. Resources are associated
   with Teams at a percentage, and a Team indicates what percentage it is
   dedicated to each Priority.

## Domain rules

1. Centres have Teams that work on Priorities.
2. Centres can create Teams as they wish.
3. Centres can assign Resources to one or more Teams.
4. A Team can only work on one Priority Type (Tactical / Initiative /
   Assistance).
5. A Team can work on more than one Priority, provided all are of the Type
   the Team is assigned to.

## Technology constraints

1. Must work fully disconnected from the internet.
2. May use Excel (Office 2021).
3. May use SharePoint 2013.
4. May use data connections where permitted.

This is an **Excel + SharePoint solution**, not a hosted web app — "code"
here means Excel structure/formulas/data validation, VBA macros, and Power
Query, all of which must run without an internet connection.

5. Python is not available on the work machine — the `management/scripts/`
   and `templates/scripts/` build/wiring scripts (and Python generally)
   only run on the home machine (Microsoft 365). Regenerate or rewire a
   workbook there, then carry the finished `.xlsx` over to work; don't
   plan on running Python at work.

## Data model

Canonical schema: [`data-model/priorities.dbml`](data-model/priorities.dbml)
(DBML format). Enhance/change it as the design evolves — keep this file in
sync with whatever the Excel workbooks actually implement.

## People

1. There is a central management person.  Call this stakehold "Management"
2. Each centre has a Centre Lead.  The Centre Lead is responsible for the data entry. 

## Process

1. Preparation Phase: As part of the preparation, 
   1. Management populates the Priorities including their attributes like Type, Rating, Actor, etc from the Tactical, Assistance and Initiative tables. 
   2. Management populates Resources, their Position is known/populated.
      1. The Position table is 
   3. Management populates other tables such as Centres, 
   4. RatingLookup is also populated so all scoring can be associated a label like "Very High", etc. (A VLOOKUP in excel of Category, MIN:MAX)
2. Send Data Collection Excel Files Phase:  Management creates & populates an Excel for each Centre Lead
   1. Management creates a Excel for each Centre (consistent filenaming convention).  This file is cloned from the "templates".
   2. Management sets the Centre name in each file to be the Centre-name for the file (or ideally this is auto-done based on the filename or automation)
   3. Management ensures the data is refreshed (data connection to the preparation phase Excel file)
3. Data Collection Phase.  Management sends a request (and link to Excel) to each Centre Lead asking to:
   1. Identify Teams : populate Teams table
   2. Indicate which Resources are working on which Teams and Percentage 
   3. Rank each Priority and whether it is Resourced (Resourced indicates the Team is working on it)
   4. Indicate what Priorities Teams are working on and Percentage

## Synthesized Data

For helping to assess the workbooks, populate the data from the Preparation phase as follows:

1. Resources: 100 synthesized Resources.  Resources are at various Position Levels 
2. Priorities: 50 Priorities, split between Tactical, Assistance and Initiative types
3. Centres: Six (6) centres; one Cyber, one Economic, fake others.

## Workflow

The workflow for a Centre Lead is:

1. Enter Team Names
2. Associate Resources to Teams and their Percentage
3. Rank Priorates
4. Indicate Team allocation to Priorities and if they are resourced. 

## Measures of success

1. Ease of data entry for Centre Leads — ideally in Excel, stored in
   SharePoint.
2. Ease of consolidating the six centres' data.
3. Correct sums/rollups of resources by Centre, Team, and Priority.
4. Visualizations in Excel are nice-to-have; the data is likely fed onward
   to Power BI or Tableau, where visualization matters more.
5. I believe this should be a Data Model in Excel so Power Pivot and Measures are possible.  The Data Model should use the relationships as shown in the DBML file.
6. For Conditional Format ensure the formulas are as clean as possible for maintainability and reading.  For example, apply to a column as opposed to a range if that is easier. 
7. Simple Excel skills for data entry.  The simpler that better as the Centre Leads are not IT experts.
8. Ability to work with these Excel files post-Claude.  The data-model, formulas, and everything should be easy to maintain, and update while disconnected from Claude and the internet. 

## Approach

1. Model the input data in Excel using the Data Model above.
2. One Excel workbook per Centre, with tabs/tables for data entry.
3. Protect reference/lookup data on those workbooks so Centre Leads can't
   corrupt it.
4. Teams Entry
   1. Maximum 15-teams - sheet should look clean.

# User Interface

### Ranking

- ideally sorted as Rank is entered

### Input Validation

- Percentages are in 5% increments. 
- Rankings should be in order, 1 being highest. 
- Duplicate rankings should be highlighted.
- A Teams allocation to Priorities should total 100%
- A Resource allocation to Teams should not be greater than 100% (it does not have to total 100%)

### Excel Tables

- Excel tables in an Excel tab should be choose a design pattern that provides for visual clues as to the size of the table. 

## Outputs

Write these as Markdown in `/docs`:

1. How the Excel files are designed..  Include tools used (for example Python to build data model and links for online help)
   1. Include ERD diagrams based off the priorities.dbml file.  Ensure these are sized for easy reading in Letter mode.  Separating into more than one diagram maybe desirable. 
2. A User Guide for Centre Leads.
3. A User Guide for the Management
4. A User Guide for the Management technical prime; including a visualization of the DMBL data model (multiple pages if required for legibility)

## References

[Claude Code cheatsheet | Claude Help Center](https://support.claude.com/en/articles/14553413-claude-code-cheatsheet) 

## Status

- `management/preparation.xlsx` — Management's Preparation-phase workbook:
  reference tables (Centres, Position, ProblemSet, InitiativeType,
  AssistanceType, RatingLookup) plus synthesized test data (100 Resources,
  50 Priorities with full Tactical/Initiative/Assistance scoring), wired
  into a real Power Pivot Data Model with relationships matching
  data-model/priorities.dbml's `Ref:` lines.
- FIXED (2026-08-17), project-wide, previously silent: every conditional-
  format fill colour in this project (Teams %/Person % red-amber-green,
  Tactical/Initiative/Assistance risk/value band colours) had never
  actually rendered — `PatternFill("solid", fgColor=colour)` alone doesn't
  work in a `dxf` (conditional format) context, only `fgColor` +
  `bgColor` together does. No error anywhere; caught only by exporting a
  sheet to PDF via `ExportAsFixedFormat` and looking at the real render
  (COM's `Interior.Color`/`DisplayFormat.Interior.Color` were themselves
  unreliable for checking this). Fixed at all 7 call sites in
  build_centre_template.py/build_preparation.py; both preparation.xlsx and
  centre-template.xlsx regenerated and re-verified. See
  excel-file-design.md for the full writeup.
- `templates/centre-template.xlsx` — Centre Lead data-entry workbook
  (Teams, Select Priorities, Priorities & Ranking, Resource Allocation —
  4 Steps as of 2026-09-03, see the Priority Selection entry below),
  native Excel validation only, no VBA. Built in 2 stages (2026-09-01) —
  build_centre_template.py (shell + RatingLookup + Steps 1-4) then
  templates/scripts/wire_reference_data.py (the other 10 reference
  tables, Power-Query-refreshable) — see that RESOLVED entry below for
  the full story. Shares preparation.xlsx's data model — same table
  names/columns, except Priority is split 3 ways by Type
  (PriorityTactical/PriorityInitiative/PriorityAssistance) and Resources
  is filtered to just this centre's associated people (2026-09-03, via
  the new ResourceCentres join table — see its own entry below). Wired
  into its own Power Pivot Data Model — 14/16 relationships live for a
  real centre file (13/16 for the generic master template specifically,
  since its "EX" placeholder CentreCode has no real ResourceCentres
  associations — see the Resources-scoping entry below), permanently:
  the 2 involving Teams can never be created, by design
  (see the RESOLVED entry below on Step 1 - Teams becoming a real
  15-row Table). Power Pivot rejects a relationship whose "one" side key
  column contains *any* blank, and Teams always has 13+ blank rows for
  any centre with fewer than 15 teams — confirmed via wire_data_model.py
  against both the unfilled template and filled dev fixtures
  (management/scripts/dev_fixtures/make_test_centres.py), both landing
  at 14/16. (Earlier versions of this note claimed first 17/17, then
  15/17, were reachable — both measured against since-abandoned
  designs; see the RESOLVED entries below for why each was replaced.
  The relationship COUNT also changed, 17 -> 16: splitting Priority 3
  ways removes the single Priorities-Team's-input-table -> Priority
  relationship, since a Priorities row could belong to any of the 3
  split tables and Power Pivot can't express an either-or FK target —
  nothing was using that relationship, so not a functional loss.)
  Nothing shipped depends on the 2 still-missing Teams relationships
  today; see the Resource x Team matrix PivotTable entry further down
  for the feature that eventually will, and how it'll need to work
  around this (likely the same Power-Query-filters-blanks approach the
  Consolidation workbook already uses). Centre Name/Code on
  the Instructions sheet are now stamped
  and locked at generation time (build_centre_template.py's `main()`
  takes an optional centre_name/centre_code, validated against
  preparation.xlsx's Centres table; `--all` generates all six in one
  run) instead of an editable "Example Centre"/"EX" placeholder with
  nothing keeping it in sync with the filename.)
  - Row 2 is no longer prefilled with example data (was confusing on a
    recurring-use workbook); worked example moved into the Instructions
    tab as text.
  - RESOLVED (2026-09-01), previously silent data-loss bug: Step 1 - Teams'
    "official Table = row 2 only, backed by a styled/unlocked 14-row
    buffer that Excel auto-extends to absorb" design (shipped in e36ca92)
    did not work. Typing into the buffer rows was real — the cells were
    unlocked and took input fine — but the Table never actually grew to
    include them, verified against a live Excel session (Table.Range
    stayed `$A$1:$E$2` after setting A3, both by direct user report and by
    a COM-driven reproduction). Since `TeamNameList`, the Step 2 team
    dropdown, and every `SUMIFS`/lookup against `Teams[...]` all key off
    the Table's actual range, a Centre Lead's 2nd through 15th team were
    silently invisible everywhere downstream — a real data-loss bug, not
    just a cosmetic one (it was reported as "the table's visual size
    doesn't match the shading" before the underlying cause was found).
    Root cause: no protected-sheet mechanism exists to grow an Excel
    Table at all (see the resize/protection finding elsewhere in this
    Status section) — the buffer-row design assumed typing-triggered
    auto-extend was an exception to that, which was never actually
    verified against live interactive typing (only against a dev-fixture
    script that force-set `Table.ref` directly via openpyxl, bypassing
    the question entirely). Fix: Step 1 - Teams is now a real, full
    15-row Table (`A1:E16`) from generation time, the same pattern
    Step 2/3 already used successfully with their 200-row Tables — no
    resize ever needed. Verified fixed via the same live-Excel
    reproduction (Table.Range now `$A$1:$E$16` immediately, and
    `TeamNameList` correctly picks up every row typed into). Trade-off
    accepted: this reopens the Teams-relationship Power Pivot blocker
    covered above, permanently rather than just pre-fill — judged worth
    it since nothing shipped depends on that relationship today.
  - NEW (2026-08-17): "Step 2 - Priorities & Ranking" (renamed
    "Step 3 - Priorities & Ranking" on 2026-09-03 — see the Priority
    Selection entry below) has a `Value/Risk
    (auto)` column — Management's Risk (Tactical) or Value (Initiative/
    Assistance) band for the row's Priority, looked up from the already-
    embedded scoring tables and colour-coded, shown at ranking time so a
    Centre Lead can see e.g. a "Minimal"-risk item they've ranked #1.
    Handles the Risk/Value "Minimal" band-name collision (opposite
    colour meaning) via Type-qualified conditional-format rules — see
    excel-file-design.md. Verified against the dev fixtures via PDF
    export (Low/Minimal/Moderate render with the correct distinct
    colours).
  - Steps 1–3 now have sheet protection: computed columns locked,
    input columns unlocked.
  - Both Allocation % columns are dropdowns constrained to 5% steps
    (list validation) rather than a custom-formula 100%-cap block —
    Excel can't combine both in one rule; the Total % columns still
    flag an overage visually. See excel-file-design.md for the
    trade-off and two more structured-reference gotchas hit while
    building this (apostrophe in a column name silently breaking
    `Table[Column]`; structured refs not working inside Data
    Validation custom formulas).
- `management/consolidation/consolidation.xlsx` — Management's Consolidation workbook:
  combines the six returned centre files into consolidated views (which
  priorities each centre works, rankings, sum of FTEs). Built by
  `management/consolidation/build_consolidation.py` (shell + static
  reference tables), `wire_power_query.py` (authors the 3 Folder-connector
  Power Query combine queries — Teams_All/Priorities_All/
  ResourceAllocation_All — and loads each into the Data Model), then
  `wire_data_model.py` (5/5 relationships, via a composite TeamKey =
  CentreCode | Team Name, since Team Name alone collides across centres)
  and `add_measures.py` (3 DAX measures: Total FTE, FTE Delivered to
  Priority — a two-hop calculation, see excel-file-design.md — and
  Priority Rank (Min)). All of this turned out fully scriptable via COM,
  including Power Query authoring and DAX measures, neither of which had
  any prior precedent in this codebase — see excel-file-design.md's
  "Consolidation workbook" section for the exact recipe and the dead ends
  it avoids. Verified end-to-end (2026-08-17) against 3 hand-checkable dev
  fixtures (management/scripts/dev_fixtures/make_test_centres.py,
  deliberately planting a same-named "Ops Team" in two different centres
  to exercise the TeamKey collision case): all 3 measures matched
  hand-computed expected values exactly via CUBEVALUE, including the
  collision case never cross-contaminating.
  The 3 requested PivotTable views (Consolidated Priorities, Sum of FTEs
  by Centre/Team, Sum of FTEs by Priority) are built by
  `create_pivots.py` (2026-08-17) — the long-standing PARKED
  PivotTable-via-COM failure below turned out to be a wrong enum
  constant (SourceType=5 instead of the real xlExternal=2), not a
  genuine COM limitation; see below, this resolves that PARKED item's
  sub-blocker 2 for real. All 3 verified against the dev fixtures:
  correct centre attribution, correct FTE sums, and the ECO/INF "Ops
  Team" collision case confirmed never merging. "Rankings" doesn't need
  a PivotTable at all — the Priorities_All (data) worksheet already is
  that view.
- `management/scripts/` and `templates/scripts/` — regenerate each
  workbook (`build_preparation.py` / `build_centre_template.py`) and wire
  it into its Data Model (`wire_data_model.py`, COM automation against a
  running Excel — see each file's docstring for the undocumented
  connection recipe it uses).
- Known gotcha (see build_centre_template.py's comment): a Table's
  displayName and a workbook defined name can't be identical — collided
  once (Tactical/Initiative/Assistance vs. the dependent-dropdown named
  ranges), and Excel silently deleted the Table on open with no error
  beyond a generic "repaired records" log. Those three reference tables
  are now named TacticalScores/InitiativeScores/AssistanceScores to
  avoid it.
- `/docs` — [`excel-file-design.md`](docs/excel-file-design.md) (design,
  tooling, known gotchas/gaps),
  [`centre-lead-user-guide.md`](docs/centre-lead-user-guide.md), and, as of
  2026-09-01,
  [`management-user-guide.md`](docs/management-user-guide.md) (the
  practical Excel-only-vs-needs-the-Build-person workflow, with 2 Mermaid
  flowcharts) and
  [`management-technical-guide.md`](docs/management-technical-guide.md)
  (3 Mermaid ERD diagrams derived from priorities.dbml, split by theme —
  scoring engine, people/positions, Centre Lead data entry — plus an
  entity-to-physical-table map across all 3 workbooks) all written. This
  satisfies CLAUDE.md's Outputs #2-4; #1's ERD sub-requirement was
  previously unmet (no diagrams existed anywhere) and is now covered by
  management-technical-guide.md rather than duplicated into
  excel-file-design.md.
- RESOLVED (2026-09-01): the process gap above (CLAUDE.md's Process
  section described a live data-connection refresh from preparation.xlsx
  when Management clones a centre file; what was actually built was a
  static snapshot baked in at generation time, requiring Python +
  redistribution for every reference-data change) is now built — see
  `templates/scripts/wire_reference_data.py`. Driven by a portability
  requirement: Management will run these files with no Python access at
  all, so Python can only ever be a build-time tool on a machine Claude
  Code has access to (see CLAUDE.md's Technology constraints #5) —
  routine reference-data changes needed an Excel-only path.
  `centre-template.xlsx` build is now 2 stages, same split the
  Consolidation workbook already used: build_centre_template.py (shell:
  RatingLookup + Lookups + Instructions + Steps 1-3) then
  wire_reference_data.py (COM, authors Power Query connections from the
  other 10 reference tables back to preparation.xlsx, loads each to a
  worksheet Table, restyles it, wires the dependent-dropdown/picker
  defined names). Once built, updating reference data is Excel-only end
  to end: Management edits preparation.xlsx directly (always was
  possible) and hits Data > Refresh All on the distributed file — no
  Python, no redistribution, for any change that's just new/edited rows.
  Structural changes (new column, new Priority Type, new dropdown) still
  need a Python regeneration + redistribution, same as before — that
  residual scope is much smaller than routine data updates in practice.
  RatingLookup is the deliberate exception, kept as a static Python-
  authored copy — see wire_reference_data.py's docstring for why (it's a
  structural constant its own VLOOKUP formulas depend on staying in a
  fixed row order, converting it would reopen exactly the fragility this
  whole effort exists to avoid, for a table that in practice never
  changes).
  Two non-obvious things had to be redesigned, not just converted, to
  make this refresh-safe:
  1. The Priority dependent-dropdown (Step 2's Type -> Priority list) and
     ResourceNameList (Step 3's resource picker) both depended on
     defined names pointing at row ranges computed at generation time —
     refresh-unsafe by construction, since a row count change would
     silently desync them. Priority is now 3 separate Power Queries
     split by Type (PriorityTactical/PriorityInitiative/
     PriorityAssistance, M-filtered), and all the affected defined names
     are now plain structured references (`PriorityTactical[Title]`,
     `Resources[FullName]`) that auto-size with their table on every
     refresh — no row math, anywhere.
  2. Sheet protection blocks a Table from growing or shrinking on
     refresh, exactly like the already-documented "Table can't resize
     while its sheet is protected" constraint for Step 1 - Teams (below)
     — confirmed by direct testing, including with
     AllowInsertingRows/AllowDeletingRows/AllowFormattingCells all
     explicitly granted (still failed identically). So these 10-turned-
     12 reference sheets (Priority's split makes it 12) are deliberately
     left unprotected — matching the precedent the Consolidation
     workbook's own Power-Query-loaded sheets already set
     (wire_power_query.py never protected them either). No password was
     ever set on these sheets; protection was only ever a guard against
     an accidental Centre Lead typo, and any accidental edit gets
     overwritten by the next refresh regardless.
     Verified (2026-09-01): a live add-a-row-to-preparation.xlsx-then-
     Refresh-All test on a wired centre file showed the new row appearing
     in the right split Priority table and its defined name within seconds,
     with zero formula errors across the whole workbook afterward; the same
     test against the *protected* design failed silently (RefreshAll
     swallowed the per-connection error) until sheets were made
     unprotected. Re-ran the dev-fixture + Consolidation regression
     end-to-end against fixtures rebuilt through the full 3-stage pipeline
     (make_test_centres.py now runs wire_reference_data.py + wire_data_model.py
     on each fixture too, not just stage 1) — all 5 hand-computed FTE
     values and the ECO/INF collision case still matched exactly.
     All 6 real per-centre files now live in `management/centres/`
     (Centre-CYB/ECO/INF/HLT/ENV/DIP.xlsx), built and wired through the
     full pipeline, 14/16 relationships each, zero formula errors.
     Gotcha for next time: `ListObjects.Add(SourceType=0, ...)` for a new
     Power-Query connection intermittently failed with a blank, contentless
     COM error after many rapid successive query-author attempts in one
     long-running Excel session (reproduced even for a trivial literal
     query with no external source at all, and even in the
     already-proven-working Consolidation workbook) — resolved every time
     by closing all open workbooks in that Excel session and retrying, no
     code change involved. Looked exactly like an environment regression at
     first (a completely unmodified, previously-working script failed
     identically); turned out to be Excel/Mashup-engine session staleness
     from heavy iteration, not a real bug. If this recurs, close all
     workbooks (not just the one being wired) before assuming the recipe
     itself is broken.
- RESOLVED (2026-09-01), found via user report ("Step 3 - Resource
  Allocation Position Title (auto) field has #REF! errors") the very
  next session after the live-refresh work above shipped: real bug,
  latent in every workbook built by that work (the master template, all
  6 real centre files, all 3 dev fixtures). Step 2's `Value/Risk (auto)`
  and Step 3's `Position Title (auto)` columns had their formulas written
  by build_centre_template.py (stage 1) referencing TacticalScores/
  InitiativeScores/AssistanceScores/Resources — tables that don't exist
  until stage 2 (wire_reference_data.py) creates them. The instant Excel
  first opened that stage-1-only file — including wire_reference_data.py's
  own first open of it — it couldn't resolve those table names and
  silently, PERMANENTLY rewrote the structured references to literal
  `#REF!` in the formula text. Creating the tables afterward doesn't
  un-corrupt an already-`#REF!`'d formula. Escaped detection because (a)
  cells with a blank input column short-circuit past the broken branch,
  so an unfilled workbook shows zero visible errors, and (b) the
  Consolidation workbook's DAX measures don't read either of these
  "(auto)" columns at all, so even the filled dev-fixture regression
  passed clean. Fixed by having build_centre_template.py leave these two
  columns styled but formula-less, and wire_reference_data.py setting the
  real formula (one bulk `Range.Formula` assignment per column, after
  temporarily unprotecting each sheet — Steps 2/3 stay protected
  otherwise, unlike the reference sheets) once the tables it references
  genuinely exist. Verified by actually typing a Team/Priority/Resource
  into a fresh workbook and confirming all three auto-columns compute
  correctly (not just checking for absence of errors on blank rows, the
  gap that let this ship in the first place) — then rebuilt and
  re-verified the master template, all 6 real centre files, and all 3 dev
  fixtures the same way, plus re-ran the Consolidation regression via its
  actual PivotTable (FTEByPriority) rather than ad-hoc CUBEVALUE formulas,
  which turned out to have their own unrelated `#GETTING_DATA`
  (Excel error 2043) timing flakiness under heavy COM automation load —
  a verification-harness quirk, not a product defect; the real PivotTable
  matched all 5 hand-computed values exactly.
- RESOLVED (2026-09-01): the RatingLookup-driven band formulas
  (`band_formulas()` in build_preparation.py) used INDEX/MATCH rather than
  VLOOKUP, which cut against maintainability by an intermediate Excel
  user. Fixed by reordering RatingLookup to `minValue, id, RatingType,
  maxValue, BandName, ColourCode` (Min now precedes Id) and switching all
  three lookups to VLOOKUP against two named table ranges per rating type.
  Both preparation.xlsx and centre-template.xlsx regenerated, recalculated
  in Excel with zero formula errors, and rewired (13/13 and 15/17
  relationships respectively — unchanged from before, since the Data
  Model wiring addresses RatingLookup.id by name, not position). See
  excel-file-design.md.
  Re-verified (2026-09-01) that this reorder didn't regress the
  Consolidation workbook: regenerated the 3 dev fixtures
  (management/scripts/dev_fixtures/make_test_centres.py) against the
  updated preparation.xlsx (unaffected by the reorder, since
  random.seed(42) keeps the synthesized Priority titles deterministic),
  then refreshed management/consolidation/consolidation.xlsx's Power
  Query + Data Model against them (read-only check, not saved). All 5
  hand-computed FTE Delivered to Priority values matched exactly, Total
  FTE = 3.00 as expected, and the ECO/INF "Ops Team" name-collision case
  stayed correctly separate (0.75 vs 0.25, never merged).
- UNBLOCKED, not yet built (2026-08-17, see below for the long PARKED
  history this closes): Resource x Team matrix PivotTable, sourced from
  the Power Pivot Data Model, in templates/centre-template.xlsx. Both
  sub-blockers that parked it are closed and the working recipe is
  proven (management/consolidation/create_pivots.py) — what's left is a
  short, mechanical application of that recipe to centre-template.xlsx
  specifically (a new script, or a manual few-click pivot), not design
  work.
  1. Refresh UX — `PivotCache.RefreshOnFileOpen = True` is an ordinary
     settable COM property once a real PivotCache exists (no VBA, no
     Ribbon-XML backstop needed after all); `create_pivots.py` sets it
     on every PivotTable it creates.
  2. Creating the PivotTable itself via COM
     (`PivotCaches().Create(xlExternal, wb.Connections("ThisWorkbookDataModel"))`
     then `.CreatePivotTable(...)`) — this was believed to be a genuine,
     unfixable COM/Excel-version limitation ("Reference isn't valid",
     tried explicit Version, refreshing the connection first, activating
     the sheet, different destinations, even retried once more against a
     real fully-wired workbook — same failure every time). It was
     actually a **parameter bug**: every attempt passed `SourceType=5`
     for `xlExternal`, but the real enum value is **2** — `5` is
     `xlPivotTableVersion15`, an unrelated constant. Found and fixed
     while building the Consolidation workbook's 3 PivotTables (below);
     with `SourceType=2` this is completely reliable, ordinary COM. No
     more reason to prefer manual creation over scripting it — see
     `management/consolidation/create_pivots.py` for the working pattern
     to reuse here.
     Separately: Excel Tables categorically cannot be resized while their
     sheet is protected — confirmed via two independent native mechanisms
     (typing into the row below, and `ListRows.Add()` / Insert Table Row),
     the latter failing identically even with `AllowInsertingRows=True`
     explicitly granted on `Worksheet.Protect`. An earlier fix attempt (shipped
     in e36ca92, superseded 2026-09-01 — see the Step 1 - Teams RESOLVED
     entry below) tried backing a single-row official Table with a styled,
     unlocked "buffer" below it, on the theory that typing into the row
     immediately below a Table auto-extends it even under protection. That
     theory was never actually true — see below — so Step 1 - Teams is now
     a real, full-size Table from generation time instead, sidestepping the
     resize question entirely (nothing needs to resize under protection
     because nothing needs to grow).
- NEW (2026-09-03): Priority Selection. `templates/centre-template.xlsx`
  gained a new sheet, "Step 2 - Select Priorities" (Team sheets stay
  Step 1; Priorities & Ranking and Resource Allocation renumbered to
  Steps 3 and 4). A Centre Lead now picks every priority under
  consideration for their centre from Management's full master list on
  this new sheet first; Step 3's Ranking dropdown then only offers
  *that* shortlist, further narrowed to the team's own Type — not the
  full master list directly, as before. Two real dead ends found before
  landing on the working design (both fully written up in
  excel-file-design.md's "Priority Selection (Step 2)" section, worth
  reading before reaching for `FILTER()` in a defined name anywhere else
  in this project): a `FILTER()`-based defined name needs an
  `_xlfn._xlws.` prefix in its raw formula text or Excel refuses to open
  the file at all (openpyxl has no awareness of this), and even fixed,
  Excel's Data Validation "List" type cannot consume a dynamic-array
  defined name as its source under any circumstance — confirmed
  decisively via `Validation.Add` throwing outright with no `INDIRECT`
  involved. Shipped design: three hidden plain-`IF()` helper columns on
  the new sheet, one per Type, with `SelectedTactical`/
  `SelectedInitiative`/`SelectedAssistance` as ordinary structured-
  reference defined names over them (`PrioritySelection[Tactical
  Helper]`, etc.) — the same already-proven pattern `TeamNameList` uses,
  no dynamic arrays anywhere. Verified via `Range.Validation.Value`
  (`True`/`False` for selected/unselected, both in isolation and on the
  real shipped `Centre-CYB.xlsx`) and a full formula-error sweep (zero)
  after a real Save+reopen cycle.
- NEW (2026-09-03): Resources scoped to Centres. `Resources` previously
  had no Centre affiliation — every centre file's Step 4 dropdown showed
  the same 100-person roster regardless of who actually works where. A
  new `preparation.xlsx` table, `ResourceCentres` (`Resource`,
  `CentreCode`), is a many-to-many join — a person can be associated
  with more than one Centre — and each centre file's own `Resources`
  copy is now filtered (via `wire_reference_data.py`) to just that
  centre's associated people, reading the file's own `CentreCode`
  defined name from inside the Power Query the same way `PrepFilePath`
  already was. `preparation.xlsx`'s own `Resources` table gained a real
  `FullName` formula column (`=A{r}&" "&B{r}`) it didn't have before, as
  the join key for `ResourceCentres` in Management's own Data Model (now
  15/15 relationships there, up from 13/13) — which broke the centre
  files' own `resources_formula()` in a small, findable way (`Table.AddColumn`
  tried to add a `FullName` column that already existed in the source
  data) until that redundant step was removed. `consolidation.xlsx`'s own
  `Resources` copy deliberately stays global/unfiltered — Management
  needs to see everyone. Full write-up, including the master-template
  13/16-vs-14/16 edge case, in excel-file-design.md's "Resources scoped
  to Centres" section. Verified: the CYB dev fixture's filtered Resources
  table matched an independently-computed ground truth exactly as a set;
  all 6 real centre files landed on different row counts
  (28/32/26/22/23/33), confirming the filter is genuinely per-centre; the
  Consolidation regression (5 hand-computed FTE values + Total FTE) still
  matched exactly after the full regeneration, since none of it touches
  Resources scoping.
- FIXED (2026-09-04), reported by a Centre Lead: the Team Name (Steps 3/4)
  and Priority Title (Step 3) dropdowns in `templates/centre-template.xlsx`
  were showing several blank entries below the real choices, and were
  listed in whatever order they were typed rather than alphabetically.
  Root cause: `TeamNameList`/`Selected{Type}` pointed straight at the
  growable-in-place `Teams`/`PrioritySelection` input Tables (15 and 200
  rows respectively, mostly blank until filled — see the Step 1 - Teams
  entry above for why they're pre-sized), and a plain Data Validation
  List shows one dropdown entry per cell including blanks; a comment on
  the original helper columns claiming Excel silently skips blank cells
  in a list source was never actually verified and was wrong. Fixed with
  a new computed helper column per dropdown
  (`=IFERROR(INDEX(_xlfn.SORT(_xlfn.FILTER(...))),ROW()-1),"")`) that
  compacts out blanks and sorts what's left, with the `_xlfn.` prefix on
  SORT/FILTER confirmed required for a plain cell formula written via
  openpyxl — omitting it doesn't corrupt the formula, it makes Excel
  refuse to open the file at all (a genuinely new failure mode for this
  codebase, not the familiar #REF! one). The small literal Lookups enums
  (Priority Type, Resource Status, Yes/No) are now written pre-sorted;
  the Power Query-loaded reference tables (Resources, the three
  Priority split queries) got a one-line `Table.Sort` M step each — no
  blank-compaction needed there since Power Query tables are always
  exact-sized. Regenerated and verified zero formula errors on the
  master template, all 6 real centre files, and all 3 dev fixtures;
  relationship counts unchanged (13/16 / 14/16). Full write-up, including
  a live typed-out-of-order verification and the `_xlfn.` isolation test,
  in excel-file-design.md's "Dropdowns sorted alphabetically and
  blank-free" section. That same section also documents an unrelated,
  pre-existing PivotTable staleness quirk found (not caused) while
  re-running the Consolidation regression — `consolidation.xlsx`'s own
  PivotTables stopped showing Infrastructure Centre's row after the dev
  fixtures were regenerated, despite `CUBEVALUE` confirming the Data
  Model itself has the correct data throughout; left as a known gap for
  whoever next touches that workbook's PivotTables, not fixed here.
