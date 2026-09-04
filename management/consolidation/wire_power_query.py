"""
Author the 3 Folder-connector Power Query combine queries (Teams_All,
Priorities_All, ResourceAllocation_All) in the Consolidation workbook, and
wire each into the Power Pivot Data Model.

Recipe (worked out by direct experimentation against a real Excel session —
see docs/excel-file-design.md's "Consolidation workbook" section for the
full story of why this specific chain of COM calls, not a more obvious one):

  1. wb.Queries.Add(Name, Formula) authors the M query itself — plain,
     reliable COM, no UI needed.
  2. ws.ListObjects.Add(SourceType=0, Source=<OLEDB Mashup connection
     string>, Destination=...) loads that query's output onto a worksheet
     as a real Excel Table, refreshable via Data > Refresh All like any
     other query. Directly connecting a Power-Query-authored table straight
     into the Data Model via Connections.Add2 with an OLEDB Mashup
     connection string technically works too, but Excel names the
     resulting Model table something generic like "Query"/"Query1" with no
     way to rename it after the fact (ModelTable.Name has no setter) — a
     dead end for anything that needs to reference the table by name (every
     relationship and DAX measure here does).
  3. The resulting ListObject CAN be renamed (unlike a ModelTable), so
     rename it to the real target name (Teams_All, etc.) before...
  4. ...adding it to the Data Model with the exact same WORKSHEET;
     connection recipe already proven throughout this codebase
     (wire_data_model.py in both management/scripts and templates/scripts)
     — which then correctly picks up the renamed Table's name.

Each query reads a CentreCode/CentreName from the source file's own
CentreCode/CentreName defined names (stamped by build_centre_template.py --
see that script and the Centre Name/Code fix in CLAUDE.md's Status), and
computes a composite TeamKey (CentreCode & "|" & [Team Name]) so that two
different centres naming a team the same thing don't collide when combined
— see docs/excel-file-design.md for why Team Name alone isn't safe to key
on across centres.

Usage:
    1. Set the "SourceFolder" cell on the Instructions tab (or pass one
       here) to a folder containing the completed centre-template.xlsx
       files.
    2. Open the Consolidation workbook in Excel.
    3. python wire_power_query.py <workbook-name-as-shown-in-excel>

Idempotent: an existing query/table/connection of the same name is left
alone, not recreated.
"""
import sys

import pywintypes
import win32com.client as win32

QUERIES = ["Teams_All", "Priorities_All", "ResourceAllocation_All"]
SOURCE_TABLE = {"Teams_All": "Teams", "Priorities_All": "Priorities",
                 "ResourceAllocation_All": "ResourceAllocation"}


def combine_formula(source_table):
    return f'''let
    SourceFolder = Excel.CurrentWorkbook(){{[Name="Config"]}}[Content]{{0}}[SourceFolder],
    Source = Folder.Files(SourceFolder),
    Filtered = Table.SelectRows(Source, each Text.EndsWith([Name], ".xlsx") and not Text.StartsWith([Name], "~$")),
    Extracted = Table.AddColumn(Filtered, "TableData", each
        let
            Wbk = Excel.Workbook([Content], null, true),
            CentreCodeRow = Wbk{{[Item="CentreCode", Kind="DefinedName"]}}[Data],
            CentreCode = Text.From(CentreCodeRow{{0}}[Column1]),
            CentreNameRow = Wbk{{[Item="CentreName", Kind="DefinedName"]}}[Data],
            CentreName = Text.From(CentreNameRow{{0}}[Column1]),
            RawTable = Wbk{{[Item="{source_table}", Kind="Table"]}}[Data],
            // Priorities/ResourceAllocation are pre-built to 200 rows so Centre
            // Leads have room to type without extending the Table — most rows
            // are blank on any real submission. Team Name is the first field
            // filled on every real row (see centre-lead-user-guide.md), so a
            // blank/whitespace Team Name reliably means "unused template row".
            DataTable = Table.SelectRows(RawTable, each [Team Name] <> null and Text.Trim([Team Name]) <> ""),
            WithCentreCode = Table.AddColumn(DataTable, "CentreCode", each CentreCode),
            WithCentreName = Table.AddColumn(WithCentreCode, "CentreName", each CentreName),
            WithKey = Table.AddColumn(WithCentreName, "TeamKey", each CentreCode & "|" & [Team Name])
        in
            WithKey
    ),
    Combined = Table.Combine(Extracted[TableData])
in
    Combined'''


def refresh_status_formula():
    """One row per source file found, plus when it was last changed.

    Two questions this answers that nothing else could: "did the refresh
    actually pick up all six centres?" and "how current is what I'm
    looking at?". Previously a centre file that was missing, misnamed, or
    saved in the wrong folder just silently contributed nothing -- the
    PivotTables looked perfectly healthy with five centres in them.

    Reads the same Folder.Files() step the combine queries already use, so
    it sees exactly what they see. DateTime.LocalNow() is evaluated at
    refresh time, which is what makes it a "data as of" stamp rather than
    a formula that would re-evaluate on every recalculation (NOW() would,
    and would therefore always show the current time even if the data were
    a week stale -- exactly the wrong behaviour here).
    """
    return '''let
    SourceFolder = Excel.CurrentWorkbook(){[Name="Config"]}[Content]{0}[SourceFolder],
    RefreshedAt = DateTime.LocalNow(),
    Source = Folder.Files(SourceFolder),
    Filtered = Table.SelectRows(Source, each Text.EndsWith([Name], ".xlsx") and not Text.StartsWith([Name], "~$")),
    // Centre must be derived BEFORE any column selection: it reads the
    // file's [Content] blob, and dropping that column first makes every
    // row fall through to the "otherwise" branch (which is exactly what
    // happened on the first attempt -- it reported three files but named
    // none of them).
    WithCentre = Table.AddColumn(Filtered, "Centre", each
        let
            Wbk = Excel.Workbook([Content], null, true)
        in
            try Text.From(Wbk{[Item="CentreName", Kind="DefinedName"]}[Data]{0}[Column1])
            otherwise "(could not read - is this a centre file?)"
    , type text),
    WithStamp = Table.AddColumn(WithCentre, "Data as of", each RefreshedAt, type datetime),
    Ordered = Table.SelectColumns(WithStamp, {"Centre", "Name", "Date modified", "Data as of"}),
    // (SelectColumns is safe from here on -- Centre is already materialised.)
    Renamed = Table.RenameColumns(Ordered, {{"Name", "File"}, {"Date modified", "File last saved"}}),
    Sorted = Table.Sort(Renamed, {{"Centre", Order.Ascending}})
in
    Sorted'''


def main(workbook_name):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    xl.DisplayAlerts = False
    existing_queries = {q.Name for q in wb.Queries}
    existing_sheets = {s.Name for s in wb.Worksheets}
    model = wb.Model
    existing_tables = {mt.Name for mt in model.ModelTables}

    for query_name in QUERIES:
        if query_name not in existing_queries:
            wb.Queries.Add(Name=query_name, Formula=combine_formula(SOURCE_TABLE[query_name]))
            print(f"OK query authored: {query_name}")
        else:
            print(f"SKIP (query already exists): {query_name}")

        if query_name in existing_tables:
            print(f"SKIP (already in model): {query_name}")
            continue

        sheet_name = f"{query_name} (data)"
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
            lo.Name = query_name
            print(f"OK loaded to worksheet: {query_name} ({lo.Range.Rows.Count - 1} rows)")
        except pywintypes.com_error as e:
            print(f"FAIL loading {query_name} to worksheet: {e}")
            continue

        path, filename = wb.FullName, wb.Name
        try:
            wb.Connections.Add2(
                Name=f"WorksheetConnection_{filename}!{query_name}",
                Description="",
                ConnectionString=f"WORKSHEET;{path}",
                CommandText=f"{filename}!{query_name}",
                lCmdtype=7,  # xlCmdExcel
                CreateModelConnection=True,
                ImportRelationships=False,
            )
            print(f"OK added to model: {query_name}")
        except pywintypes.com_error as e:
            print(f"FAIL adding {query_name} to model: {e}")

    # Refresh Status: worksheet-only, deliberately NOT added to the Data
    # Model. It's an operational check for Management ("did all six files
    # actually get picked up, and how fresh are they?"), not something any
    # measure or PivotTable should be able to join to.
    if "RefreshStatus" not in existing_queries:
        wb.Queries.Add(Name="RefreshStatus", Formula=refresh_status_formula())
        print("OK query authored: RefreshStatus")
    else:
        print("SKIP (query already exists): RefreshStatus")

    if "Refresh Status" not in {s.Name for s in wb.Worksheets}:
        ws = wb.Worksheets.Add()
        ws.Name = "Refresh Status"
        ws.Range("A1").Value = "Source files picked up by the last Data > Refresh All"
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").Font.Size = 12
        ws.Range("A2").Value = (
            "One row per .xlsx found in the source folder. If a centre is "
            "missing here, it is missing from every PivotTable too.")
        ws.Range("A2").Font.Italic = True
        try:
            lo = ws.ListObjects.Add(SourceType=0, Source=(
                'OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
                'Location=RefreshStatus;Extended Properties=""'
            ), Destination=ws.Range("A4"))
            lo.QueryTable.CommandType = 2
            lo.QueryTable.CommandText = "SELECT * FROM [RefreshStatus]"
            lo.QueryTable.Refresh()
            lo.Name = "RefreshStatus"
            ws.Range("A5").Value = None  # let the query own the body
            for col, width in (("A", 28), ("B", 30), ("C", 22), ("D", 22)):
                ws.Columns(col).ColumnWidth = width
            # Count sits above the table so it reads before the detail.
            ws.Range("C1").Formula = '=COUNTA(RefreshStatus[File])&" file(s) found"'
            ws.Range("C1").Font.Bold = True
            print(f"OK loaded to worksheet: RefreshStatus "
                  f"({lo.Range.Rows.Count - 1} rows)")
        except pywintypes.com_error as e:
            print(f"FAIL loading RefreshStatus to worksheet: {e}")
    else:
        print("SKIP (sheet already exists): Refresh Status")

    for conn in wb.Connections:
        try:
            conn.OLEDBConnection.BackgroundQuery = False
        except Exception:
            pass

    wb.Save()
    print("SAVED")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python wire_power_query.py <workbook-name-as-shown-in-excel>")
    main(sys.argv[1])
