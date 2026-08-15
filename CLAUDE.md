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
2. Data Collection Phase.  Management sends a request (and link to Excel) to each Centre Lead asking to:
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

## Approach

1. Model the input data in Excel using the Data Model above.
2. One Excel workbook per Centre, with tabs/tables for data entry.
3. Protect reference/lookup data on those workbooks so Centre Leads can't
   corrupt it.

## Outputs

Write these as Markdown in `/docs`:

1. How the Excel files are designed..  Include tools used (for example Python to build data model and links for online help)
2. A User Guide for Centre Leads.

## Status

- `management/preparation.xlsx` — Management's Preparation-phase workbook:
  reference tables (Centres, Position, ProblemSet, InitiativeType,
  AssistanceType, RatingLookup) plus synthesized test data (100 Resources,
  50 Priorities with full Tactical/Initiative/Assistance scoring), wired
  into a real Power Pivot Data Model with relationships matching
  data-model/priorities.dbml's `Ref:` lines.
- `templates/centre-template.xlsx` — Centre Lead data-entry workbook
  (Teams, Priorities & Ranking, Resource Allocation), native Excel
  validation only, no VBA. Embeds read-only copies of all 11 reference
  tables from preparation.xlsx and shares its data model — same table
  names/columns. Wired into its own Power Pivot Data Model (15/17
  relationships live; the 2 involving Teams can't be created until a
  Centre Lead fills in real team data — Teams is currently a blank
  template, and Power Pivot rejects a relationship whose "one" side is
  mostly blank cells. Re-run wire_data_model.py, which is idempotent,
  once real data exists.)
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
- Not yet started: `/docs` outputs (Excel design doc — including the
  Python/COM tooling used and links to relevant docs — and the Centre
  Lead user guide).
