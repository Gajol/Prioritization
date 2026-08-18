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
   4. RatingLookup is also populated so all scoring can be associated a label like "Very High", etc. (A VLOOKUP in excel of Cateogry, MIN:MAX)
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

## Outputs

Write these as Markdown in `/docs`:

1. How the Excel files are designed..  Include tools used (for example Python to build data model and links for online help)
   1. Include ERD diagrams based off the priorities.dbml file.  Ensure these are sized for easy reading in Letter mode.  Separating into more than one diagram maybe desirable. 
2. A User Guide for Centre Leads.

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
  (Teams, Priorities & Ranking, Resource Allocation), native Excel
  validation only, no VBA. Embeds read-only copies of all 11 reference
  tables from preparation.xlsx and shares its data model — same table
  names/columns. Wired into its own Power Pivot Data Model (15/17
  relationships live; the 2 involving Teams can't be created until a
  Centre Lead fills in real team data. Confirmed by re-running
  wire_data_model.py against the current template (2026-08-17): Power
  Pivot rejects the relationship even with Teams shrunk to a single
  blank data row — the real rule is that the "one" side's key column
  can't contain *any* blank, not just duplicate blanks as originally
  assumed. So there's no template-side trick that unblocks this before
  real data exists; re-run wire_data_model.py, which is idempotent,
  once a Centre Lead has filled in Step 1 - Teams. CONFIRMED WORKING
  (2026-08-17) against 3 filled dev fixtures
  (management/scripts/dev_fixtures/make_test_centres.py) — all three
  wired to 17/17 relationships once real, non-blank Team Name data
  existed. Centre Name/Code on the Instructions sheet are now stamped
  and locked at generation time (build_centre_template.py's `main()`
  takes an optional centre_name/centre_code, validated against
  preparation.xlsx's Centres table; `--all` generates all six in one
  run) instead of an editable "Example Centre"/"EX" placeholder with
  nothing keeping it in sync with the filename.)
  - Row 2 is no longer prefilled with example data (was confusing on a
    recurring-use workbook); worked example moved into the Instructions
    tab as text.
  - NEW (2026-08-17): "Step 2 - Priorities & Ranking" has a `Value/Risk
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
  tooling, known gotchas/gaps) and
  [`centre-lead-user-guide.md`](docs/centre-lead-user-guide.md) written.
- Known process gap (flagged in excel-file-design.md, not yet built):
  CLAUDE.md's Process section describes a live data-connection refresh
  from preparation.xlsx when Management clones a centre file; what's
  actually built is a static snapshot copy baked in at generation time.
  Now a lower-risk follow-on than it looked: the Consolidation workbook's
  Power Query combine (wire_power_query.py) proved out the M-authoring-
  via-COM technique this would need, including the worksheet-Table
  intermediate step that makes the resulting table nameable/wireable —
  same recipe should transfer directly.
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
     Separately, RESOLVED (shipped in e36ca92): Excel Tables categorically
     cannot be resized while their sheet is protected — confirmed via two
     independent native mechanisms (typing into the row below, and
     `ListRows.Add()` / Insert Table Row), the latter failing identically
     even with `AllowInsertingRows=True` explicitly granted on
     `Worksheet.Protect`. Fix: Step 1 - Teams' official Table range is a
     single data row, backed by a 14-row buffer (rows 3-16, teams #2-15)
     that's pre-styled/validated/unlocked but sits outside the Table —
     typing into it works under protection, and the Table auto-extends to
     absorb the row. Sheet protection stays on throughout, so no need to
     leave Step 1 unprotected. This fix was originally motivated by the
     Teams relationship problem above but doesn't solve it (see above —
     the real blocker is any blank on the one-side key, not duplicate
     blanks); it's kept anyway since it's a real, independently-useful fix
     for the protection/resize conflict.
