"""
Author the Consolidation workbook's DAX measures — the first DAX measures
anywhere in this project (management/scripts/wire_data_model.py and
templates/scripts/wire_data_model.py only ever wire tables/relationships,
never measures).

Gotcha found by direct experimentation: Model.ModelMeasures.Add's
FormatInformation parameter looks optional (win32com's generated signature
shows it defaulting to an empty PyOleEmpty placeholder) but passing None
fails every time with a bare "Exception occurred" / E_INVALIDARG and no
useful message. It needs a real ModelFormat* object — e.g.
model.ModelFormatDecimalNumber or model.ModelFormatWholeNumber — not None
and not a plain string/omitted argument.

FTE Delivered to Priority is the one genuinely tricky measure here: a
resource's contribution to a specific Priority is a two-hop calculation
(their % allocation to a Team, times that Team's % allocation to the
Priority), and ResourceAllocation_All/Priorities_All are two fact tables
hanging off the same Teams_All dimension — not directly related to each
other. Default DAX relationship propagation (dimension -> fact only) won't
produce the right number by itself, so the measure explicitly correlates
the two fact tables per TeamKey inside SUMX/CALCULATE/FILTER rather than
leaning on ambient filter context. See docs/excel-file-design.md for the
full reasoning and the hand-verified expected values this was checked
against (5/5 exact matches via CUBEVALUE against the dev fixtures in
management/scripts/dev_fixtures/).

Usage:
    1. Open the target workbook in Excel yourself, with wire_power_query.py
       and wire_data_model.py (this directory) already run against it.
    2. python add_measures.py <workbook-filename-as-shown-in-Excel>

Idempotent: a measure already present (by name) is left alone.
"""
import sys

import pywintypes
import win32com.client as win32

# (measure_name, associated_table, dax_formula, format_property_name)
MEASURES = [
    (
        "Total FTE",
        "ResourceAllocation_All",
        "SUM(ResourceAllocation_All[Allocation % of Person Effort])",
        "ModelFormatDecimalNumber",
    ),
    (
        "FTE Delivered to Priority",
        "Priorities_All",
        """SUMX(
    Priorities_All,
    VAR CurrentTeamKey = Priorities_All[TeamKey]
    VAR CurrentAllocToPriority = Priorities_All[Allocation % of Team Effort]
    VAR TeamFTE =
        CALCULATE(
            SUM(ResourceAllocation_All[Allocation % of Person Effort]),
            FILTER(ALL(ResourceAllocation_All), ResourceAllocation_All[TeamKey] = CurrentTeamKey)
        )
    RETURN CurrentAllocToPriority * TeamFTE
)""",
        "ModelFormatDecimalNumber",
    ),
    (
        "Priority Rank (Min)",
        "Priorities_All",
        "MIN(Priorities_All[Rank])",
        "ModelFormatWholeNumber",
    ),
]


def main(workbook_name):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    model = wb.Model
    existing = {m.Name for m in model.ModelMeasures}

    for name, table_name, formula, format_prop in MEASURES:
        if name in existing:
            print(f"SKIP (already exists): {name}")
            continue
        try:
            table = model.ModelTables(table_name)
            fmt = getattr(model, format_prop)
            # A decimal format with no DecimalPlaces set defaults to 0 --
            # e.g. 0.75 would display as "1" in any PivotTable that uses
            # this measure (Grand Totals still sum correctly; only the
            # per-cell display rounds). Not relevant for the whole-number
            # format (Priority Rank).
            if format_prop == "ModelFormatDecimalNumber":
                fmt.DecimalPlaces = 2
            model.ModelMeasures.Add(name, table, formula, fmt)
            print(f"OK measure added: {name}")
        except pywintypes.com_error as e:
            print(f"FAIL adding {name}: {e}")

    wb.Save()
    print("SAVED")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python add_measures.py <workbook-filename-as-shown-in-excel>")
    main(sys.argv[1])
