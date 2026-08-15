"""
Scan every cell of the named sheets (or all sheets, if none named) for
Excel formula errors, using a real, already-open Excel session over COM —
not LibreOffice.

Why not LibreOffice: this project hit three separate cases where
LibreOffice reported a workbook as fully error-free (recalc.py, zero
errors) while real Excel either repaired it on open or silently replaced
a broken formula with #REF! (which LibreOffice evaluated differently and
without complaint). LibreOffice is not a reliable proxy for "does this
open cleanly in Excel" for a workbook using Tables/structured references
this heavily — treat it as a formula-syntax smoke test only, and use this
script for anything you actually need to trust.

Reads cell.Text (the as-displayed string), not cell.Value — Value returns
error cells as a raw negative integer via COM, not a "#REF!"-style
string, which silently defeats a naive `isinstance(v, str)` error check
(this bit us once already).

Usage:
    1. Open the target workbook in Excel yourself.
    2. python check_errors_excel.py <workbook-filename-as-shown-in-Excel> [sheet1 sheet2 ...]
"""
import sys

import win32com.client as win32

ERROR_PREFIXES = ("#REF!", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!", "#N/A")


def main(workbook_name, sheet_names=None):
    xl = win32.GetActiveObject("Excel.Application")
    wb = next((w for w in xl.Workbooks if w.Name == workbook_name), None)
    if wb is None:
        open_names = [w.Name for w in xl.Workbooks]
        raise SystemExit(f"{workbook_name!r} not open in Excel. Open workbooks: {open_names}")

    xl.CalculateFullRebuild()

    names = sheet_names or [ws.Name for ws in wb.Worksheets]
    total = 0
    for name in names:
        ws = wb.Worksheets(name)
        found = []
        for cell in ws.UsedRange:
            text = cell.Text
            if isinstance(text, str) and text.startswith(ERROR_PREFIXES):
                found.append((cell.Address(), text, cell.Formula))
        total += len(found)
        status = "OK" if not found else f"{len(found)} error(s)"
        print(f"{name}: {status}")
        for addr, text, formula in found[:20]:
            print(f"    {addr}  {text}  <-  {formula}")

    print(f"\nTotal error cells across {len(names)} sheet(s): {total}")


if __name__ == "__main__":
    if len(sys.argv) < 2:
        raise SystemExit(
            "Usage: python check_errors_excel.py <workbook-name-as-shown-in-excel> [sheet1 sheet2 ...]"
        )
    main(sys.argv[1], sys.argv[2:] or None)
