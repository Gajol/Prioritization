"""
Wire the Centre Lead workbook's Excel Tables into its Power Pivot Data
Model — the reference-table copies (relating to each other exactly as in
management/scripts/wire_data_model.py) plus the centre's own input tables
(Teams, Priorities, ResourceAllocation) relating back into those
reference tables.

Same COM-automation approach and connection recipe as
management/scripts/wire_data_model.py — see that file's docstring for why
this can't be done with openpyxl and how the WORKSHEET/xlCmdExcel
connection shape was reverse-engineered.

Note: the Tactical/Initiative/Assistance reference sheets' Tables are
named TacticalScores/InitiativeScores/AssistanceScores here (not
Tactical/Initiative/Assistance) — those plain names are already taken by
this workbook's dependent-dropdown defined names, and a Table can't share
a name with a defined name (see build_centre_template.py's comment on
that bug).

Usage:
    1. Open the target workbook in Excel yourself.
    2. python wire_data_model.py <workbook-filename-as-shown-in-Excel>

Idempotent: tables/relationships already in the model are skipped.
"""
import sys

import pywintypes
import win32com.client as win32

TABLES = ["Centres", "Position", "ProblemSet", "InitiativeType", "AssistanceType",
          "RatingLookup", "Resources", "PriorityTactical", "PriorityInitiative",
          "PriorityAssistance", "TacticalScores", "InitiativeScores",
          "AssistanceScores", "Teams", "Priorities", "ResourceAllocation"]

# Priority is split into 3 Power-Query tables by Type (PriorityTactical/
# PriorityInitiative/PriorityAssistance — see wire_reference_data.py),
# so each Scores table's PriorityReference relates to its own matching
# split table, not a single shared "Priority" table. Priorities (the
# Centre Lead's Step 2 input, unrelated to the reference split above)
# still needs a single Priority-title dimension to relate to regardless
# of which Type a row is — that's PriorityTactical UNION'd conceptually
# with the other two, which Power Pivot can't express directly, so it
# relates to whichever of the 3 split tables actually contains that row's
# title; Power Pivot only requires the relationship to resolve for rows
# that have a match, and a given Priority Title exists in exactly one
# split table by construction (Priority is filtered by Type, and Type is
# fixed per Priority).
RELATIONSHIPS = [
    ("Resources", "PositionTitle", "Position", "Title"),
    ("TacticalScores", "PriorityReference", "PriorityTactical", "Title"),
    ("TacticalScores", "ProblemSet", "ProblemSet", "Title"),
    ("TacticalScores", "Actor", "Centres", "CentreCode"),
    ("TacticalScores", "RiskLabelId", "RatingLookup", "id"),
    ("InitiativeScores", "PriorityReference", "PriorityInitiative", "Title"),
    ("InitiativeScores", "InitiativeType", "InitiativeType", "Title"),
    ("InitiativeScores", "Actor", "Centres", "CentreCode"),
    ("InitiativeScores", "ValueLabelId", "RatingLookup", "id"),
    ("AssistanceScores", "PriorityReference", "PriorityAssistance", "Title"),
    ("AssistanceScores", "AssistanceType", "AssistanceType", "Title"),
    ("AssistanceScores", "Actor", "Centres", "CentreCode"),
    ("AssistanceScores", "ValueLabelId", "RatingLookup", "id"),
    ("Priorities", "Team Name", "Teams", "Team Name"),
    ("ResourceAllocation", "Team Name", "Teams", "Team Name"),
    ("ResourceAllocation", "Resource", "Resources", "FullName"),
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
