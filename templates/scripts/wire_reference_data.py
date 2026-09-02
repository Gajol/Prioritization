"""
Stage 2 of the Centre Lead workbook build (after build_centre_template.py):
authors Power Query connections back to preparation.xlsx for every
reference table EXCEPT RatingLookup, loads each into a worksheet Table,
restyles it to match the rest of the workbook, and wires the defined
names Step 2/3's dropdowns depend on. This is what makes routine reference
-data updates (a new Priority, a new Resource, a new Centre) something
Management can push out with Data > Refresh All in plain Excel — no
Python needed — instead of re-running build_centre_template.py and
redistributing the file. See CLAUDE.md's Status entry (2026-09-01) for
the portability motivation.

Same COM recipe as management/consolidation/wire_power_query.py (see that
file's docstring for the fuller "why this exact chain of COM calls"
story): author the M query via wb.Queries.Add, load it to a worksheet
Table via ListObjects.Add with an OLEDB Mashup connection, then rename the
resulting ListObject to the real target name (a ModelTable can't be
renamed after the fact, but a ListObject can). Unlike the Consolidation
workbook's Folder-connector queries, these read a single file's Tables
directly via Excel.Workbook(File.Contents(...)) rather than combining a
folder — same underlying function Power Query's own "From Workbook"
import uses.

The source file path is NOT hardcoded into the queries: it's read from a
one-row "Config" Table on the Instructions sheet (PrepFilePath column,
stamped with an absolute path by build_centre_template.py at generation
time) via Excel.CurrentWorkbook(){[Name="Config"]}[Content]{0}[PrepFilePath]
-- the same pattern build_consolidation.py already uses for its
SourceFolder. If preparation.xlsx moves, Management edits that one cell
and hits Data > Refresh All; no query needs re-authoring.

Two tables get non-trivial treatment, both to remove a dependency on
fixed row positions that a refresh could invalidate:

- Priority is split into 3 Power Queries (PriorityTactical/
  PriorityInitiative/PriorityAssistance, M-filtered by Type), each loaded
  to its own small Table. Step 2's dependent Priority dropdown then reads
  a defined name (Tactical/Initiative/Assistance, as before -- required
  by the sheet's INDIRECT($C2) formula, which passes the literal Type
  string) that's now a structured reference (e.g. PriorityTactical[Title])
  instead of a hardcoded row range -- it auto-sizes with the table on
  every refresh, with no row-bound math to keep in sync. Before this,
  build_centre_template.py computed fixed A$2:A$18-style ranges at
  generation time; those would silently go stale the moment a refresh
  added or removed a Priority of that Type.
- ResourceNameList (Step 3's resource picker) had the identical
  fixed-range problem; fixed the same way, now =Resources[FullName].

The Tactical/Initiative/Assistance scoring tables' band-colour
conditional formatting is applied to each Table's label column via
ListColumns(...).DataBodyRange rather than a fixed row range, for the
same reason -- a genuine Table-column CF rule (as opposed to a CF rule
that merely happens to cover a table's current extent) is tracked by
Excel and auto-extends when the table grows or shrinks on refresh.

RatingLookup is the one reference table deliberately NOT converted here
-- still built as a static copy by build_centre_template.py. It's a
structural constant (3 RatingTypes x 5 bands, in that exact grouped
order) that this workbook's VLOOKUP band-lookup formulas
(Likelihood_MinTable/Risk_MinTable/Value_MinTable named ranges,
see management/scripts/build_preparation.py's band_formulas()) depend on
staying in that fixed order and row count. Converting it to a live query
would reopen that fragility for a table that, in practice, essentially
never changes -- not worth it. If RatingLookup's bands genuinely need to
change, that still goes through build_centre_template.py/
build_preparation.py and a redistributed file, same as any structural
change.

These 12 sheets are deliberately NOT sheet-protected, unlike the earlier
static-copy version and unlike Steps 1-3. Confirmed by direct testing
(2026-09-01) that Excel blocks a Table from growing or shrinking on
refresh under sheet protection -- the exact same "Table can't resize
while its sheet is protected" constraint already documented for Step 1 -
Teams (see CLAUDE.md), just hit via a different code path (a query
refresh, rather than manual typing or ListRows.Add()) -- confirmed to
fail identically even with AllowInsertingRows/AllowDeletingRows/
AllowFormattingCells all explicitly granted. This matches the existing
precedent already set by the Consolidation workbook: its Power-Query-
loaded "*_All" sheets (wire_power_query.py) were never protected either,
for the same underlying reason. No password was ever set on these
sheets -- the protection was only ever a guard against an accidental
Centre Lead typo, not a security boundary -- and any accidental edit
would just get overwritten by the next refresh anyway, so leaving them
unprotected trades a soft guardrail for the refresh actually working,
which is the entire point of this script.

Usage:
    1. Open the target workbook (build_centre_template.py's stage-1
       output) in Excel yourself.
    2. python wire_reference_data.py <workbook-filename-as-shown-in-Excel>

Idempotent: a query/table/defined-name already present is left alone.
Run templates/scripts/wire_data_model.py after this (it expects these
tables to already exist).
"""
import sys
import time
from pathlib import Path

import pywintypes
import win32com.client as win32
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parent))
from build_centre_template import REF_HEADER_FILL, HEADER_FONT_COLOR, RISK_BANDS, VALUE_BANDS  # noqa: E402

PREP_PATH_M = 'Excel.CurrentWorkbook(){[Name="Config"]}[Content]{0}[PrepFilePath]'

# (query_name, sheet_name, table_name, source_table_in_prep, col_widths, extra_m)
SIMPLE_TABLES = [
    ("Centres", "Centres", "Centres", "Centres", [6, 26, 12], None),
    ("Position", "Position", "Position", "Position", [26, 14, 10, 8, 12], None),
    ("ProblemSet", "ProblemSet", "ProblemSet", "ProblemSet", [6, 28, 14, 18], None),
    ("InitiativeType", "InitiativeType", "InitiativeType", "InitiativeType", [6, 32, 14, 18, 18], None),
    ("AssistanceType", "AssistanceType", "AssistanceType", "AssistanceType", [6, 28, 14, 18, 18], None),
    ("Resources", "Resources", "Resources", "Resources", [16, 16, 26, 26],
     'Table.AddColumn(Data, "FullName", each [FirstName] & " " & [LastName])'),
]

# (query_name, sheet_name, table_name, priority_type, col_widths)
PRIORITY_SPLIT = [
    ("PriorityTactical", "Priority - Tactical", "PriorityTactical", "Tactical", [42, 12, 10, 12]),
    ("PriorityInitiative", "Priority - Initiative", "PriorityInitiative", "Initiative", [42, 12, 10, 12]),
    ("PriorityAssistance", "Priority - Assistance", "PriorityAssistance", "Assistance", [42, 12, 10, 12]),
]

# (query_name, sheet_name, table_name, source_table_in_prep, col_widths, bands, label_col_name)
SCORES_TABLES = [
    ("TacticalScores", "Tactical", "TacticalScores", "Tactical",
     [42, 10, 26, 8, 10, 11, 11, 13, 9, 11, 12, 10], RISK_BANDS, "RiskLabel"),
    ("InitiativeScores", "Initiative", "InitiativeScores", "Initiative",
     [34, 10, 30, 16, 9, 10, 7, 8, 11, 13, 10], VALUE_BANDS, "ValueLabel"),
    ("AssistanceScores", "Assistance", "AssistanceScores", "Assistance",
     [34, 10, 26, 16, 10, 12, 9, 8, 11, 13, 10], VALUE_BANDS, "ValueLabel"),
]

SHEET_ORDER = [
    "Instructions", "Centres", "Position", "Resources",
    "Priority - Tactical", "ProblemSet", "Tactical",
    "Priority - Initiative", "InitiativeType", "Initiative",
    "Priority - Assistance", "AssistanceType", "Assistance",
    "RatingLookup", "Lookups",
    "Step 1 - Teams", "Step 2 - Priorities & Ranking", "Step 3 - Resource Allocation",
]


def passthrough_formula(source_table, extra_m=None):
    body = (
        f'let\n'
        f'    PrepFilePath = {PREP_PATH_M},\n'
        f'    Source = Excel.Workbook(File.Contents(PrepFilePath), null, true),\n'
        f'    Data = Source{{[Item="{source_table}",Kind="Table"]}}[Data]'
    )
    if extra_m:
        body += f',\n    Result = {extra_m}\nin\n    Result'
    else:
        body += '\nin\n    Data'
    return body


def priority_split_formula(priority_type):
    return (
        f'let\n'
        f'    PrepFilePath = {PREP_PATH_M},\n'
        f'    Source = Excel.Workbook(File.Contents(PrepFilePath), null, true),\n'
        f'    Data = Source{{[Item="Priority",Kind="Table"]}}[Data],\n'
        f'    Filtered = Table.SelectRows(Data, each [Type] = "{priority_type}")\n'
        f'in\n'
        f'    Filtered'
    )


def bgr(hex_argb):
    hex_rgb = hex_argb[-6:]
    r, g, b = int(hex_rgb[0:2], 16), int(hex_rgb[2:4], 16), int(hex_rgb[4:6], 16)
    return r + (g << 8) + (b << 16)


def load_query_to_table(wb, query_name, formula, sheet_name, table_name,
                         existing_queries, existing_sheets, existing_tables):
    if query_name not in existing_queries:
        wb.Queries.Add(Name=query_name, Formula=formula)
        print(f"OK query authored: {query_name}")
    else:
        print(f"SKIP (query already exists): {query_name}")

    if table_name in existing_tables:
        print(f"SKIP (already loaded): {table_name}")
        return None

    if sheet_name not in existing_sheets:
        ws = wb.Worksheets.Add()
        ws.Name = sheet_name
    else:
        ws = wb.Worksheets(sheet_name)

    conn_string = (
        f'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
        f'Location={query_name};Extended Properties=""'
    )
    try:
        lo = ws.ListObjects.Add(SourceType=0, Source=conn_string, Destination=ws.Range("A1"))
        lo.QueryTable.CommandType = 2  # xlCmdSql
        lo.QueryTable.CommandText = f"SELECT * FROM [{query_name}]"
        lo.QueryTable.Refresh()
        lo.Name = table_name
        print(f"OK loaded to table: {table_name} ({lo.Range.Rows.Count - 1} rows)")
        return lo
    except pywintypes.com_error as e:
        print(f"FAIL loading {table_name}: {e}")
        return None


def style_table(ws, lo, col_widths):
    lo.TableStyle = "TableStyleMedium3"
    header = lo.HeaderRowRange
    header.Interior.Color = bgr(REF_HEADER_FILL)
    header.Font.Color = bgr(HEADER_FONT_COLOR)
    header.Font.Bold = True
    header.HorizontalAlignment = -4108  # xlCenter
    for i, w in enumerate(col_widths, start=1):
        ws.Columns(i).ColumnWidth = w
    ws.Activate()
    ws.Application.ActiveWindow.DisplayGridlines = False


def main(workbook_name):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    xl.DisplayAlerts = False
    existing_queries = {q.Name for q in wb.Queries}
    existing_sheets = {s.Name for s in wb.Worksheets}
    existing_tables = set()
    for sheet in wb.Worksheets:
        for lo in sheet.ListObjects:
            existing_tables.add(lo.Name)

    for query_name, sheet_name, table_name, source_table, widths, extra_m in SIMPLE_TABLES:
        lo = load_query_to_table(
            wb, query_name, passthrough_formula(source_table, extra_m), sheet_name, table_name,
            existing_queries, existing_sheets, existing_tables,
        )
        if lo is not None:
            style_table(wb.Worksheets(sheet_name), lo, widths)

    for query_name, sheet_name, table_name, ptype, widths in PRIORITY_SPLIT:
        lo = load_query_to_table(
            wb, query_name, priority_split_formula(ptype), sheet_name, table_name,
            existing_queries, existing_sheets, existing_tables,
        )
        if lo is not None:
            style_table(wb.Worksheets(sheet_name), lo, widths)

    for query_name, sheet_name, table_name, source_table, widths, bands, label_col in SCORES_TABLES:
        lo = load_query_to_table(
            wb, query_name, passthrough_formula(source_table), sheet_name, table_name,
            existing_queries, existing_sheets, existing_tables,
        )
        if lo is not None:
            ws = wb.Worksheets(sheet_name)
            style_table(ws, lo, widths)
            col_range = lo.ListColumns(label_col).DataBodyRange
            col_letter = get_column_letter(col_range.Column)
            first_row = col_range.Row
            # FormulaRule (xlExpression), not CellIsRule: a CellIsRule
            # text-equality rule made Excel flag these for repair on open
            # in the earlier static-copy version (build_centre_template.py
            # has the full note) -- same avoidance applies here.
            for band_name, colour in bands:
                fc = col_range.FormatConditions.Add(
                    Type=2,  # xlExpression
                    Formula1=f'={col_letter}{first_row}="{band_name}"',
                )
                fc.Interior.Color = bgr(colour)

    # Defined names last, in one batch, after every table load/style
    # above has fully settled -- creating a structured-reference name (e.g.
    # =Resources[FullName]) immediately after the table it references was
    # just loaded was observed to intermittently fail here
    # ("There's a problem with this formula", despite the table genuinely
    # existing) when interleaved with the loop; batching it at the end
    # after a recalculation was reliable in testing.
    xl.CalculateFullRebuild()
    time.sleep(1)
    existing_names = {n.Name for n in wb.Names}
    if "ResourceNameList" not in existing_names:
        wb.Names.Add(Name="ResourceNameList", RefersTo="=Resources[FullName]")
        print("OK defined name: ResourceNameList")
    else:
        print("SKIP (already exists): ResourceNameList")
    for _, _, table_name, ptype, _ in PRIORITY_SPLIT:
        if ptype not in existing_names:
            wb.Names.Add(Name=ptype, RefersTo=f"={table_name}[Title]")
            print(f"OK defined name: {ptype}")
        else:
            print(f"SKIP (already exists): {ptype}")

    for name in reversed(SHEET_ORDER):
        if name in {s.Name for s in wb.Worksheets}:
            wb.Worksheets(name).Move(Before=wb.Worksheets(1))
    wb.Worksheets(1).Activate()

    wb.Save()
    print("SAVED")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python wire_reference_data.py <workbook-name-as-shown-in-excel>")
    main(sys.argv[1])
