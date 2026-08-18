"""
Create the 3 Data-Model-sourced PivotTables for the Consolidation workbook:
Consolidated Priorities, Sum of FTEs by Centre/Team, Sum of FTEs by
Priority. ("Rankings" doesn't get a PivotTable — the Priorities_All (data)
worksheet already is that view, sortable/filterable as-is; see
docs/excel-file-design.md.)

This was believed unfixable via COM (see the PARKED item's history in
CLAUDE.md's Status section: "Reference isn't valid", "confirmed NOT a
parameter mistake", manual creation recommended instead). It turned out to
BE a parameter mistake: `wb.PivotCaches().Create(SourceType:=xlExternal,
SourceData:=<ThisWorkbookDataModel connection>)` needs the real xlExternal
enum value, which is **2** — every earlier attempt (including one retried
in this project on 2026-08-17, see excel-file-design.md) used 5, which is
actually `xlPivotTableVersion15`, an unrelated constant. With the correct
value this is completely reliable, ordinary COM — no manual UI step
needed after all.

Fields are addressed by their cube-field name, `[TableName].[ColumnName]`
for a column, `[Measures].[MeasureName]` for a DAX measure or Excel's
auto-created implicit count measures (`[Measures].[__XL_Count <Table>]`).

Gotcha: a measure's format (set via add_measures.py's FormatInformation
argument) directly controls how PivotTable cells display it. A
ModelFormatDecimalNumber with no DecimalPlaces set defaults to 0 decimal
places — meaning e.g. 0.75 displays as "1" and 0.25 as "0" (Grand Totals
still sum correctly under the hood; only the per-cell *display* rounds).
Set DecimalPlaces explicitly before assigning the format.

Usage:
    1. Open the target workbook in Excel yourself, with wire_power_query.py,
       wire_data_model.py, and add_measures.py already run against it.
    2. python create_pivots.py <workbook-filename-as-shown-in-Excel>

Idempotent: an existing PivotTable of the same name is left alone.
"""
import sys

import pywintypes
import win32com.client as win32

XL_EXTERNAL = 2  # NOT 5 -- see docstring above
XL_ROW_FIELD = 1
XL_COLUMN_FIELD = 2
XL_DATA_FIELD = 4

PIVOTS = [
    (
        "ConsolidatedPriorities",
        "Consolidated Priorities",
        [("[Priority].[Title]", XL_ROW_FIELD)],
        [("[Centres].[Centre]", XL_COLUMN_FIELD)],
        [("[Measures].[__XL_Count Priorities_All]", XL_DATA_FIELD)],
    ),
    (
        "FTEByCentreTeam",
        "Sum of FTEs by Centre-Team",
        [("[Centres].[Centre]", XL_ROW_FIELD), ("[Teams_All].[Team Name]", XL_ROW_FIELD)],
        [],
        [("[Measures].[Total FTE]", XL_DATA_FIELD)],
    ),
    (
        "FTEByPriority",
        "Sum of FTEs by Priority",
        [("[Priority].[Title]", XL_ROW_FIELD)],
        [],
        [("[Measures].[FTE Delivered to Priority]", XL_DATA_FIELD)],
    ),
]

SHEET_ORDER = [
    "Instructions", "Consolidated Priorities", "Sum of FTEs by Centre-Team",
    "Sum of FTEs by Priority", "Priorities_All (data)", "Teams_All (data)",
    "ResourceAllocation_All (data)", "Centres", "Priority", "Resources",
]


def main(workbook_name):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    xl.DisplayAlerts = False
    existing_sheets = {s.Name for s in wb.Worksheets}
    existing_pivots = set()
    for sheet in wb.Worksheets:
        for pt in sheet.PivotTables():
            existing_pivots.add(pt.Name)

    conn = wb.Connections("ThisWorkbookDataModel")

    for pt_name, sheet_name, rows, cols, values in PIVOTS:
        if pt_name in existing_pivots:
            print(f"SKIP (already exists): {pt_name}")
            continue
        try:
            if sheet_name not in existing_sheets:
                ws = wb.Worksheets.Add()
                ws.Name = sheet_name
            else:
                ws = wb.Worksheets(sheet_name)
            pc = wb.PivotCaches().Create(SourceType=XL_EXTERNAL, SourceData=conn)
            pt = pc.CreatePivotTable(TableDestination=ws.Range("A3"), TableName=pt_name)
            for field_name, orientation in rows + cols + values:
                pt.CubeFields(field_name).Orientation = orientation
            pt.PivotCache().RefreshOnFileOpen = True
            print(f"OK pivot created: {pt_name} on {sheet_name!r}")
        except pywintypes.com_error as e:
            print(f"FAIL creating {pt_name}: {e}")

    for name in reversed(SHEET_ORDER):
        if name in {s.Name for s in wb.Worksheets}:
            wb.Worksheets(name).Move(Before=wb.Worksheets(1))
    wb.Worksheets(1).Activate()

    wb.Save()
    print("SAVED")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python create_pivots.py <workbook-filename-as-shown-in-excel>")
    main(sys.argv[1])
