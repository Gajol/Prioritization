# TODO

## Goal this serves

The core Centre Lead workflow is: rank priorities well (`CLAUDE.md`'s Workflow
step 3). Value/Risk + colour banding should be visible wherever that ranking
actually happens — ideally at data-entry time, not just after the fact in
Management's Consolidation workbook — so a Centre Lead can see "this one's Very
High risk" while they're still deciding what to rank #1.

Both items below need the same underlying lookup (given a Priority Title + Type,
return its Score/Label/ColourCode from Management's Tactical/Initiative/
Assistance scoring) — worth factoring that into one shared helper rather than
writing it twice.

## 1. Centre-template: show Value/Risk + colour band during ranking (Step 2) — DONE (2026-08-17)

- [x] Added a computed `Value/Risk (auto)` column to "Step 2 - Priorities &
      Ranking", looked up by Priority Title + Type — same
      `INDEX(...,MATCH(...))` pattern as the existing `Type (auto)` column.
- [x] Colour-coded via Type-qualified conditional formatting (10 rules: 5 Risk
      bands + 5 Value bands, each gated on Type — handles the Risk/Value
      "Minimal" band-name collision).
- [x] Regenerated `templates/centre-template.xlsx`, verified with
      `scripts/check_errors_excel.py`, re-ran `wire_data_model.py` (15/17,
      unaffected).
- [x] Updated `docs/centre-lead-user-guide.md` and `docs/excel-file-design.md`.
- [x] **Bonus find while verifying this**: discovered and fixed a real,
      previously-silent bug affecting every conditional-format fill in the
      whole project (openpyxl's `PatternFill(fgColor=...)` alone doesn't
      render in a `dxf`/conditional-format context — needs `bgColor` too).
      Fixed at all 7 call sites, `preparation.xlsx` and `centre-template.xlsx`
      both regenerated and re-verified via real PDF export
      (`ExportAsFixedFormat`, since COM's `.Interior.Color` properties turned
      out unreliable for checking this). See `excel-file-design.md`.

## 2. Consolidation workbook: Value/Risk + colour band in the Rankings view

As discussed — lets Management visually check whether Centre Leads' rankings
line up with Management's own Value/Risk scoring, across all six centres at
once.

- [ ] `build_consolidation.py`: add a normalized `PriorityValueRisk` reference
      table (Priority Title, Type, Score, Label, ColourCode), built from
      `preparation.xlsx`'s Tactical/Initiative/Assistance sheets.
- [ ] `wire_power_query.py`: merge `PriorityValueRisk` into the `Priorities_All`
      combine query (keyed on Priority Title) so Label/Colour sit directly next
      to Rank on the flat `Priorities_All (data)` sheet — that sheet is already
      the de facto Rankings view.
- [ ] `wire_power_query.py`: add the same Type-qualified conditional formatting
      to the new Label column (10 rules: 5 Risk bands + 5 Value bands, each
      gated on Type). **Remember**: `PatternFill("solid", fgColor=X, bgColor=X)`
      — both colours, not just `fgColor` — or the fill silently won't render
      (see item 1's bonus find below).
- [ ] `wire_data_model.py`: relate `PriorityValueRisk` into the Data Model
      (`Priority Title -> Priority.Title`) so it's also usable as a filter/slicer
      on the existing PivotTables.
- [ ] Verify against the dev fixtures — known expected values:
      Harden network perimeter defenses (Risk 6, "Low"), Close critical
      patching gap (Risk 2, "Minimal"), Reduce third-party vendor risk exposure
      (Risk 6, "Low"), Launch predictive analytics pilot (Value 7, "Moderate"),
      Provide cross-centre threat-briefing support (Value 5, "Limited").
- [ ] Update `docs/excel-file-design.md`'s "Consolidation workbook" section.

## 3. Carried over from before, not yet done

- [ ] Resource x Team matrix PivotTable in `centre-template.xlsx` — unblocked
      (the `xlExternal=2` fix applies here too), just needs the same
      `create_pivots.py`-style recipe applied to this workbook. Not yet built.
- [ ] Live data-connection refresh from `preparation.xlsx` into
      `centre-template.xlsx` (currently a static snapshot baked in at
      generation time) — the Power-Query-authoring-via-COM technique proven
      for the Consolidation workbook makes this lower-risk than it looked
      before that work, per `CLAUDE.md`'s Status section.
