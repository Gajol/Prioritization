"""
Wire the Consolidation workbook's Data Model: adds the static reference
tables (Centres, Priority, Resources) via the same WORKSHEET; connection
recipe used throughout this codebase (see management/scripts/wire_data_model.py
and templates/scripts/wire_data_model.py), then wires the 5 relationships
that tie the combined per-centre tables (Teams_All, Priorities_All,
ResourceAllocation_All — already added to the model by wire_power_query.py)
back to those reference tables and to each other via the composite TeamKey
column (CentreCode | Team Name — see wire_power_query.py's docstring for
why Team Name alone isn't safe to relate on across centres).

Usage:
    1. Open the target workbook in Excel yourself.
    2. Run wire_power_query.py first (it must run before this script, since
       it's what creates the Teams_All/Priorities_All/ResourceAllocation_All
       tables this script relates).
    3. python wire_data_model.py <workbook-filename-as-shown-in-Excel>

Idempotent: tables/relationships already in the model are skipped.
"""
import sys

import pywintypes
import win32com.client as win32

TABLES = ["Centres", "Priority", "Resources"]

RELATIONSHIPS = [
    ("Priorities_All", "TeamKey", "Teams_All", "TeamKey"),
    ("ResourceAllocation_All", "TeamKey", "Teams_All", "TeamKey"),
    ("Priorities_All", "Priority Title", "Priority", "Title"),
    ("ResourceAllocation_All", "Resource", "Resources", "FullName"),
    ("Teams_All", "CentreCode", "Centres", "CentreCode"),
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
