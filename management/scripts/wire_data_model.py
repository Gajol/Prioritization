"""
Wire a workbook's Excel Tables into its Power Pivot Data Model, with
relationships matching data-model/priorities.dbml's `Ref:` lines.

Why COM automation, not openpyxl: the Data Model is a proprietary embedded
format Excel itself owns (xl/model/item1.data, an internal Analysis
Services binary) — no pure-Python library can write it. This drives your
real, already-running Excel over COM instead.

The exact connection shape below (WORKSHEET; connection string,
CommandText="<file>!<table>", lCmdtype=7/xlCmdExcel) is what Excel's own
"Add to Data Model" button produces — reverse-engineered by doing that
click once and reading back WorksheetDataConnection.Connection /
.CommandText / .CommandType, since Microsoft doesn't document it. The
more "obvious" OLEDB Microsoft.Mashup.OleDb.1 connection string (used for
Power-Query-authored tables) does NOT work for a plain native Excel Table
and fails with a provider-not-registered error regardless of environment.

Usage:
    1. Open the target workbook in Excel yourself (regular open, so
       Power Pivot's COM add-in is loaded in that session).
    2. python wire_data_model.py <workbook-filename-as-shown-in-Excel>

Idempotent: tables/relationships already in the model are skipped, not
duplicated or errored on.
"""
import sys

import pywintypes
import win32com.client as win32

TABLES = ["Centres", "Position", "ProblemSet", "InitiativeType", "AssistanceType",
          "RatingLookup", "Resources", "Priority", "Tactical", "Initiative", "Assistance"]

RELATIONSHIPS = [
    ("Resources", "PositionTitle", "Position", "Title"),
    ("Tactical", "PriorityReference", "Priority", "Title"),
    ("Tactical", "ProblemSet", "ProblemSet", "Title"),
    ("Tactical", "Actor", "Centres", "CentreCode"),
    ("Tactical", "RiskLabelId", "RatingLookup", "id"),
    ("Initiative", "PriorityReference", "Priority", "Title"),
    ("Initiative", "InitiativeType", "InitiativeType", "Title"),
    ("Initiative", "Actor", "Centres", "CentreCode"),
    ("Initiative", "ValueLabelId", "RatingLookup", "id"),
    ("Assistance", "PriorityReference", "Priority", "Title"),
    ("Assistance", "AssistanceType", "AssistanceType", "Title"),
    ("Assistance", "Actor", "Centres", "CentreCode"),
    ("Assistance", "ValueLabelId", "RatingLookup", "id"),
]


def main(workbook_name):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    path, filename, model = wb.FullName, wb.Name, wb.Model
    existing_tables = {mt.Name for mt in model.ModelTables}

    for t in TABLES:
        if t in existing_tables:
            print(f"SKIP (already in model): {t}")
            continue
        try:
            wb.Connections.Add2(
                Name=f"WorksheetConnection_{filename}!{t}",
                Description="",
                ConnectionString=f"WORKSHEET;{path}",
                CommandText=f"{filename}!{t}",
                lCmdtype=7,  # xlCmdExcel
                CreateModelConnection=True,
                ImportRelationships=False,
            )
            print(f"OK added to model: {t}")
        except pywintypes.com_error as e:
            print(f"FAIL adding {t}: {e}")

    existing_rels = {
        (r.ForeignKeyTable.Name, r.ForeignKeyColumn.Name,
         r.PrimaryKeyTable.Name, r.PrimaryKeyColumn.Name)
        for r in model.ModelRelationships
    }
    rel_ok = 0
    for fk_table, fk_col, pk_table, pk_col in RELATIONSHIPS:
        if (fk_table, fk_col, pk_table, pk_col) in existing_rels:
            print(f"SKIP (already related): {fk_table}.{fk_col} -> {pk_table}.{pk_col}")
            rel_ok += 1
            continue
        try:
            fk = model.ModelTables(fk_table).ModelTableColumns(fk_col)
            pk = model.ModelTables(pk_table).ModelTableColumns(pk_col)
            model.ModelRelationships.Add(ForeignKeyColumn=fk, PrimaryKeyColumn=pk)
            rel_ok += 1
            print(f"OK relationship: {fk_table}.{fk_col} -> {pk_table}.{pk_col}")
        except pywintypes.com_error as e:
            print(f"FAIL relationship {fk_table}.{fk_col} -> {pk_table}.{pk_col}: {e}")

    print(f"{rel_ok}/{len(RELATIONSHIPS)} relationships present")
    print("Model tables:", [mt.Name for mt in model.ModelTables])
    wb.Save()
    print("SAVED")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python wire_data_model.py <workbook-name-as-shown-in-excel>")
    main(sys.argv[1])
