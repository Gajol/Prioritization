"""
Open a workbook in real Excel, force a full recalculation, and save it.

WHY THIS IS A REQUIRED BUILD STEP, not a convenience:

openpyxl writes formulas as *text*. It does not evaluate them, so a
freshly-generated .xlsx has formula cells with no cached value. Excel
itself doesn't care -- it recalculates on open. But **Power Query does
not open the file in Excel**: `Excel.Workbook(File.Contents(path))` reads
the stored cached values straight out of the XML. A formula cell that has
never been calculated reads as blank.

That means: after any openpyxl rebuild of `management/preparation.xlsx`,
every *computed* column in it (RiskLabel, ValueLabel, Likelihood, FullName,
the band colours -- anything that is a formula rather than a literal) is
blank as far as downstream consumers are concerned, until Excel has opened
and saved it once.

The failure is silent and looks like something else entirely. It surfaced
here (2026-09-04) as the centre template's "Value/Risk (auto)" column
returning empty for every priority. MATCH() found the right row; INDEX()
returned nothing, because `TacticalScores[RiskLabel]` -- Power-Query-copied
out of a freshly openpyxl-written preparation.xlsx -- was an entirely blank
column. Nothing errored anywhere.

So the correct build order is always:

    python management/scripts/build_preparation.py management/preparation.xlsx
    python scripts/recalc_and_save.py management/preparation.xlsx   <-- this
    ... then build/wire any centre file or the Consolidation workbook

Usage:
    python recalc_and_save.py <path.xlsx> [<path.xlsx> ...]
"""
import os
import sys
import time

import win32com.client as win32


def recalc(path):
    full = os.path.abspath(path)
    try:
        xl = win32.GetActiveObject("Excel.Application")
    except Exception:
        xl = win32.Dispatch("Excel.Application")
    xl.Visible = True
    xl.DisplayAlerts = False
    time.sleep(1)

    already_open = next((w for w in xl.Workbooks if w.FullName == full), None)
    wb = already_open or xl.Workbooks.Open(full)
    xl.CalculateFullRebuild()
    time.sleep(3)

    errors = []
    for sheet in wb.Worksheets:
        used = sheet.UsedRange
        for r in range(used.Row, used.Row + used.Rows.Count):
            for c in range(used.Column, used.Column + used.Columns.Count):
                v = sheet.Cells(r, c).Value
                if isinstance(v, str) and v.startswith("#"):
                    errors.append(f"{sheet.Name}!{sheet.Cells(r, c).Address} = {v}")

    wb.Save()
    print(f"recalculated + saved: {path}")
    print(f"  formula errors: {len(errors)}")
    for e in errors[:20]:
        print("   ", e)
    if not already_open:
        wb.Close(SaveChanges=False)
    return len(errors)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit("Usage: python recalc_and_save.py <path.xlsx> [...]")
    total = sum(recalc(p) for p in sys.argv[1:])
    sys.exit(1 if total else 0)
