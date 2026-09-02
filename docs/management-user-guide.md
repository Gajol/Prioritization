# User Guide for Management

## What this covers

Your role across the quarterly cycle, in plain Excel terms — what you can do
yourself (no Python, no IT ticket) versus what needs to come from whoever
maintains the build scripts (called "the Build person" below; that's Claude
Code / Doug on the home machine today).

The short version: **routine reference-data changes are Excel-only now** — a
new Priority, a new Resource, a new Centre. Edit `preparation.xlsx` directly
and hit **Data > Refresh All** on the distributed centre files. **Structural
changes** — a new column, a new Priority Type, a brand-new centre's
data-entry file, or anything about `RatingLookup` — still need the Build
person. The rest of this guide tells you which is which.

## The quarterly cycle

```mermaid
flowchart LR
    classDef excel fill:#d4edda,stroke:#28a745,color:#155724
    classDef build fill:#fff3cd,stroke:#e0a800,color:#856404
    classDef lead fill:#d1ecf1,stroke:#17a2b8,color:#0c5460

    A["1. Preparation<br/>edit preparation.xlsx"]:::excel --> B["2. Send Files<br/>Build person generates,<br/>you distribute"]:::build
    B --> C["3. Data Collection<br/>Centre Leads fill in<br/>Steps 1-3, send back"]:::lead
    C --> D["4. Consolidation<br/>Refresh consolidation.xlsx,<br/>review the 3 PivotTables"]:::excel
    D -. next quarter .-> A
```

<span style="background:#d4edda">Green</span> = you, in Excel, on your own.
<span style="background:#fff3cd">Amber</span> = needs the Build person.
<span style="background:#d1ecf1">Blue</span> = the Centre Lead's part (see
[`centre-lead-user-guide.md`](centre-lead-user-guide.md)).

## 1. Preparation — editing `preparation.xlsx`

This workbook holds every piece of reference data the six centre files and
the Consolidation workbook are built from. You edit it directly in Excel —
always could, this part was never Python-dependent.

| Sheet | Holds | Columns |
|---|---|---|
| Centres | The six centres | Centre, CentreCode |
| Position | Job titles/levels | Title, PositionArea, Position, Level |
| Resources | People | FirstName, LastName, PositionTitle |
| Priority | Every priority, all three types | Title, Type, Impact, Resources |
| Tactical | Risk scoring for Tactical priorities | PriorityReference, Actor, ProblemSet, Intent, Capability, Consequence (Likelihood/Risk/band columns compute automatically) |
| Initiative | Value scoring for Initiative priorities | PriorityReference, Actor, InitiativeType, Dividend, Feasibility, Cost (Value/band columns compute automatically) |
| Assistance | Value scoring for Assistance priorities | PriorityReference, Actor, AssistanceType, Alignment, Contribution, Capacity (Value/band columns compute automatically) |
| ProblemSet / InitiativeType / AssistanceType | Small category lookups used by the scoring sheets | Title + a couple of descriptive columns |
| RatingLookup | Band definitions (thresholds, names, colours) for Likelihood/Risk/Value | minValue, id, RatingType, maxValue, BandName, ColourCode |

### Adding a new Priority

A Priority is really two entries, on two different sheets — both needed:

1. On **Priority**, add a row: Title (this is the name everything else
   refers back to — spell it exactly the same everywhere), Type (Tactical /
   Initiative / Assistance), Impact, Resources.
2. On the matching scoring sheet (**Tactical**, **Initiative**, or
   **Assistance** — whichever Type you picked), add a row: PriorityReference
   (must match the Title exactly), Actor, the relevant category column, and
   the raw score inputs (e.g. Intent/Capability/Consequence for Tactical).
   The Likelihood/Risk/Value and band-label/colour columns to the right are
   **formulas** — they fill in on their own once the raw inputs are there.

To add the row itself: click into the last cell of the last row and press
**Tab** (or right-click a cell in the table and choose **Insert > Table Rows
Above**) — either way Excel extends the Table properly and copies the
formula columns down. Don't just type into the blank row below a Table by
itself; on an unprotected sheet it often works, but the two methods above
are the ones Excel guarantees will behave correctly.

### Adding a new Resource

Add a row to **Resources**: FirstName, LastName, PositionTitle (spell it
exactly as it appears on the **Position** sheet).

### Adding a new Centre

Add a row to **Centres**: Centre name, CentreCode. This makes the centre
*known* to the reference data — dropdowns, the Consolidation workbook, etc.
will recognize it. It does **not** create that centre's actual data-entry
file; for that, see [Phase 2](#2-send-data-collection-files--needs-the-build-person)
below.

### Changing RatingLookup

You can edit `RatingLookup` directly here — it's a normal Excel sheet. But
**this is the one exception to the "just refresh" story**: the six
distributed centre files do *not* pull `RatingLookup` live. It's baked in
when the Build person generates them, on purpose (see
[excel-file-design.md](excel-file-design.md#resolved-static-snapshot-vs-live-refresh-2026-09-01)
for the technical reason). If you change a band's threshold, name, or
colour, ask the Build person to regenerate and redistribute — a plain
Refresh All won't pick it up.

## 2. Send Data Collection Files — needs the Build person

Creating or regenerating the six centre files still requires running the
Python build scripts — you can't do this step yourself. Ask the Build
person for:

- A brand-new centre's file (after adding it to Centres, above).
- Fresh copies of all six, if something structural changed.

What comes back is `Centre-<CentreCode>.xlsx` for each centre (filenames
already match the Centres sheet — nothing for you to rename). Your part:
distribute them to the right Centre Leads, typically via SharePoint.

## 3. Data Collection — waiting on Centre Leads

Nothing to build here. Centre Leads fill in Steps 1-3 following
[`centre-lead-user-guide.md`](centre-lead-user-guide.md) and send the file
back. Once you have all six back, collect them into **one folder** — the
Consolidation workbook reads whatever's in that folder.

## 4. Consolidation — `consolidation.xlsx`, Excel only

1. Make sure all six returned files are in one folder.
2. Open `consolidation.xlsx`.
3. On the **Instructions** tab, check the **SourceFolder** cell points at
   that folder — it's a normal editable cell, update it if the folder moved.
4. **Data > Refresh All.**
5. Review the three PivotTables:
   - **Consolidated Priorities** — which priorities each centre is working.
   - **Sum of FTEs by Centre-Team** — resourcing by centre and team.
   - **Sum of FTEs by Priority** — total FTE landing on each priority across
     all six centres.

If a cell briefly shows `#GETTING_DATA` right after a refresh, that's normal
— it's Excel still populating the Data Model in the background. Give it a
few seconds and refresh again if it doesn't clear on its own.

## Routine reference-data updates, after files are already distributed

This is the workflow the live-refresh work exists for. Once a centre file
has already gone out to a Centre Lead, you can still get new Priorities or
Resources in front of them without resending the file:

```mermaid
flowchart TD
    classDef excel fill:#d4edda,stroke:#28a745,color:#155724
    classDef build fill:#fff3cd,stroke:#e0a800,color:#856404

    Q{What are you changing?}
    Q -->|New Priority| R1["Add a row to Priority AND the<br/>matching Tactical/Initiative/<br/>Assistance sheet"]:::excel
    Q -->|New Resource| R2["Add a row to Resources"]:::excel
    Q -->|New Centre| R3["Add a row to Centres<br/>(the file itself still needs<br/>the Build person)"]:::excel
    Q -->|RatingLookup band/colour| R4["Ask the Build person —<br/>the one exception, stays static"]:::build
    Q -->|New column, new Priority Type,<br/>new dropdown, any structure change| R5["Ask the Build person<br/>(needs Python)"]:::build

    R1 --> F["Data > Refresh All on each<br/>already-distributed centre file"]:::excel
    R2 --> F
```

That's it for the Excel-only path: edit `preparation.xlsx`, then whoever has
the distributed file (you or the Centre Lead) hits **Data > Refresh All**.
No Python, no new file.

## When you need the Build person

- Creating a brand-new centre's data-entry file, or regenerating all six.
- Any structural change: a new column, a new Priority Type, a new dropdown,
  changing how Teams/Priorities/Resource Allocation are laid out.
- Any change to `RatingLookup` (thresholds, names, colours).
- Anything that errors in a way not covered in this guide's Troubleshooting
  section below.

## Troubleshooting

- **Refresh doesn't pick up a change I made** — check the **PrepFilePath**
  cell on the centre file's Instructions tab (Config table) actually points
  at the current location of `preparation.xlsx`. If the file moved, update
  that cell, then Refresh All again.
- **A PivotTable or measure looks wrong right after refreshing** — try
  Refresh All once more; the Data Model can take a few seconds to fully
  populate (see the `#GETTING_DATA` note above).
- **Power Pivot features (PivotTables, DAX measures) don't show up at all
  in your copy of Excel** — Power Pivot ships with Office 2021 Professional
  Plus and most volume-license editions, but not retail Home & Student /
  Home & Business. Check your Office edition if this happens; raw data
  entry and everything not Data-Model-driven still works regardless.
- **A reference sheet in a distributed centre file looks stale or wrong** —
  those sheets are read-only for Centre Leads by design; if the data itself
  is wrong, fix it in `preparation.xlsx` and refresh (or ask the Build
  person, if it's a `RatingLookup` or structural issue).

## Related docs

- [`centre-lead-user-guide.md`](centre-lead-user-guide.md) — what Centre
  Leads see and do.
- [`excel-file-design.md`](excel-file-design.md) — the technical design,
  including the live-refresh architecture referenced throughout this guide.
