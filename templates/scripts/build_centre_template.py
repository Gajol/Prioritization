"""
Generate the Centre Lead data-entry workbook, sharing the same data model
as management/preparation.xlsx, and adds the Centre Lead's own input
tables (Teams, Priorities & Ranking, Resource Allocation).

This is stage 1 of a 2-stage build (mirroring how the Consolidation
workbook is built): this script only creates RatingLookup (a static copy
— see wire_reference_data.py's docstring for why it alone stays static),
the small hardcoded Lookups enum sheet, and the Instructions/Step 1-3
sheets. Every other reference sheet (Centres, Position, ProblemSet,
InitiativeType, AssistanceType, Resources, Priority split 3 ways by Type,
and the Tactical/Initiative/Assistance scoring tables) is created by
running templates/scripts/wire_reference_data.py against this script's
output next, via COM against a running Excel — it authors Power Query
connections back to preparation.xlsx so Management can update reference
data and refresh (Data > Refresh All) without needing Python at all. The
output of stage 1 alone is not a complete, distributable workbook.

Usage:
    python build_centre_template.py <preparation.xlsx> <output.xlsx> [centre-code]
    python build_centre_template.py <preparation.xlsx> --all <output-dir>

The optional centre-code (must match a CentreCode in preparation.xlsx's Centres
table) stamps and locks the Instructions sheet's Centre Name/Code cells at
generation time instead of leaving them as editable placeholders. --all
generates all six centre files in one run, named Centre-<CentreCode>.xlsx.

After running, verify formulas with ../../scripts/recalc_windows.py — but
point it at a COPY of the output, never the file itself. LibreOffice's
store() rewrites the whole file, and its OOXML export is not fully
Excel-conformant for Tables combined with shared conditional-formatting
styles: a LibreOffice-resaved copy of this workbook triggered Excel's
"repaired records" dialog on open even though LibreOffice itself reported
zero formula errors. Ship only the openpyxl-original; use LibreOffice
solely as a disposable formula-correctness check.

Then, with the output file open in Excel, run (in this order):
    1. wire_reference_data.py <workbook-name> <preparation.xlsx-path> —
       authors the Power-Query-refreshable reference sheets.
    2. ../../management/scripts/wire_data_model.py — wires the tables
       into the Power Pivot Data Model.
Both resaves are safe, since they're done by real Excel.
"""
import sys
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule, Rule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.styles.numbers import NumberFormat
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

sys.path.insert(0, str(Path(__file__).resolve().parents[2] / "config"))
from theme import load_theme  # noqa: E402
from buildstamp import build_stamp  # noqa: E402

# Every colour below comes from config/theme.json -- the single source of
# truth shared with build_preparation.py and build_consolidation.py. Paste
# corporate brand hex codes in there, not here. load_theme() contrast-checks
# every fill/text pair against WCAG AA on import and refuses to build if one
# fails, so a brand colour can't silently ship an unreadable workbook (run
# `python config/theme.py` for the full report).
THEME = load_theme()

FONT = THEME.font
N_ROWS = 200
# How many rows of each 200-row entry Table stay visible by default. The
# Tables themselves are still full-size (they must be -- see the Step 1 -
# Teams note below on why they can't grow under sheet protection); this
# only hides the empty tail so a small centre isn't scrolling past ~170
# blank rows. Centre Leads can unhide normally if they ever need more.
VISIBLE_ROWS = 30
# Rows on the read-only "Ranked View" sheet. Deliberately smaller than
# N_ROWS: that sheet recomputes its sort once per cell (see the note where
# it's built), so this bounds the cost. It reports overflow rather than
# truncating silently.
RANKED_ROWS = 100

YELLOW, _ = THEME.cell("editable")
GREY, _ = THEME.cell("computed")
HEADER_FILL, HEADER_FONT_COLOR = THEME.cell("entry_header")
REF_HEADER_FILL, _ = THEME.cell("reference_header")
GREEN, GREEN_TEXT = THEME.status("ok")
AMBER, AMBER_TEXT = THEME.status("under")
RED, RED_TEXT = THEME.status("over")
TAB_INSTRUCTIONS = THEME.tab("instructions")
TAB_ENTRY = THEME.tab("entry")
TAB_REFERENCE = THEME.tab("reference")

thin = Side(style="thin", color=THEME.border)
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

# (band name, fill, text) -- text is auto-picked for best contrast against
# each fill by theme.py unless overridden in theme.json.
RISK_BANDS = THEME.bands("Risk")
VALUE_BANDS = THEME.bands("Value")


def style_header(cell, ref=False):
    cell.font = Font(name=FONT, bold=True, color=HEADER_FONT_COLOR)
    cell.fill = PatternFill("solid", fgColor=REF_HEADER_FILL if ref else HEADER_FILL)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = BORDER


def style_body(cell, editable=False, bold=False):
    cell.font = Font(name=FONT, bold=bold)
    cell.border = BORDER
    if editable:
        cell.fill = PatternFill("solid", fgColor=YELLOW)
        cell.protection = Protection(locked=False)


def style_computed(cell):
    """A grey, formula-driven cell: locked (sheet protection must be on
    for that to matter — see input-sheet setup below)."""
    cell.fill = PatternFill("solid", fgColor=GREY)
    cell.font = Font(name=FONT)
    cell.border = BORDER


def _status_rule(fill, text, suffix, formula):
    """A conditional-format rule that carries BOTH colour and a text suffix.

    The text suffix is the accessibility half: red/amber/green alone puts
    all of the meaning in hue, which fails for the ~8% of men with a colour
    vision deficiency (and in greyscale print, which these workbooks do get
    printed in). Excel's built-in icon sets can't express this particular
    scale, because "exactly 100%" is the GOOD value with bad values on
    both sides -- icon sets are strictly monotonic, so a middle-is-best
    scale can't be mapped onto one. Overriding the number format from the
    rule gets the same job done without an extra column: the cell still
    holds a real number (so it still sums, sorts and compares), it just
    renders as e.g. `105% over`.
    """
    dxf = DifferentialStyle(
        font=Font(name=FONT, color=text),
        fill=PatternFill("solid", fgColor=fill, bgColor=fill),
        numFmt=NumberFormat(numFmtId=200, formatCode=f'0%" {suffix}"'),
    )
    return Rule(type="expression", formula=[formula], dxf=dxf)


def status_format(ws, cell_range, guard_ref, value_ref, include_under):
    """Red/amber/green + text suffix on a 'should total 100%' column.

    guard_ref is the cell that says "this row is in use" (blank rows stay
    unformatted); value_ref is the total being judged. include_under=False
    for a person's time, which legitimately need not reach 100%.
    """
    rules = [(RED, RED_TEXT, "over", f'AND({guard_ref}<>"",{value_ref}>1)')]
    if include_under:
        rules.append((AMBER, AMBER_TEXT, "under",
                      f'AND({guard_ref}<>"",{value_ref}<1,{value_ref}>0)'))
    rules.append((GREEN, GREEN_TEXT, "ok", f'AND({guard_ref}<>"",{value_ref}=1)'))
    for fill, text, suffix, formula in rules:
        ws.conditional_formatting.add(cell_range, _status_rule(fill, text, suffix, formula))


def paste_guard(ws, row_range, formula):
    """Highlight a whole row that breaks a uniqueness rule.

    Data Validation only fires on interactive entry -- pasting a block of
    rows silently bypasses every duplicate check in this workbook. The
    check columns still recalculate, but a single tinted cell is easy to
    scroll past. This tints the entire row instead, so a bad paste is
    obvious without hunting for it.
    """
    ws.conditional_formatting.add(
        row_range,
        FormulaRule(formula=[formula],
                    fill=PatternFill("solid", fgColor=RED, bgColor=RED),
                    font=Font(name=FONT, color=RED_TEXT, bold=True))
    )


def read_table(src_wb, sheet_name, header_row=1):
    ws = src_wb[sheet_name]
    headers = [c.value for c in ws[header_row]]
    while headers and headers[-1] is None:
        headers.pop()
    n = len(headers)
    rows = []
    for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
        if row[0] is None:
            continue
        rows.append(row[:n])
    return headers, rows


def write_reference_table(ws, name, top_row, headers, rows, col_widths=None, start_col=1,
                           protect=True):
    for i, h in enumerate(headers):
        style_header(ws.cell(row=top_row, column=start_col + i, value=h), ref=True)
    for r_off, row in enumerate(rows):
        for c_off, val in enumerate(row):
            style_body(ws.cell(row=top_row + 1 + r_off, column=start_col + c_off, value=val))
    last_row = top_row + len(rows)
    last_col_letter = get_column_letter(start_col + len(headers) - 1)
    first_col_letter = get_column_letter(start_col)
    ref = f"{first_col_letter}{top_row}:{last_col_letter}{last_row}"
    tab = Table(displayName=name, ref=ref)
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium3", showRowStripes=True)
    ws.add_table(tab)
    if col_widths:
        for i, w in enumerate(col_widths):
            ws.column_dimensions[get_column_letter(start_col + i)].width = w
    ws.sheet_view.showGridLines = False
    if protect:
        ws.protection.sheet = True
    return last_row


def main(prep_path, out_path, centre_name=None, centre_code=None):
    src = openpyxl.load_workbook(prep_path, data_only=True)

    if centre_code is not None:
        _, centres_rows = read_table(src, "Centres")
        by_code = {row[2]: row[1] for row in centres_rows}  # CentreCode -> Centre
        if centre_code not in by_code:
            raise ValueError(
                f"Unknown centre code {centre_code!r}; must be one of {sorted(by_code)}"
            )
        if centre_name is None:
            centre_name = by_code[centre_code]

    wb = openpyxl.Workbook()

    # ================================================================= Instructions
    ws = wb.active
    ws.title = "Instructions"
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 100

    text_blocks = [
        ("Prioritization — Centre Data Entry Template", 14, True),
        ("", None, False),
        ("Purpose", 12, True),
        ("One workbook per centre. Enter this centre's teams, rank the "
         "priorities each team is working on, mark whether each is "
         "resourced, and record how resources split their time across "
         "teams.", None, False),
        ("", None, False),
        ("This workbook shares its data model with Management's Preparation "
         "workbook (data-model/priorities.dbml). The grey-header sheets "
         "(Centres, Position, ProblemSet, InitiativeType, AssistanceType, "
         "RatingLookup, Resources, Priority - Tactical/Initiative/"
         "Assistance, Tactical, Initiative, Assistance) are Management's "
         "reference data, protected against editing. The blue-header "
         "sheets are yours to fill in. Most grey-header sheets refresh "
         "from Management's data (Data > Refresh All) rather than being "
         "sent out fresh each time — as a Centre Lead you won't normally "
         "need to do this yourself.", None, False),
        ("", None, False),
        ("How to use this workbook", 12, True),
        ("1. 'Step 1 - Teams' — list every team this centre has, whether it "
         "has dedicated resources, and which Priority Type it works on.", None, False),
        ("2. 'Step 2 - Select Priorities' — from Management's master list, "
         "pick every priority your centre is considering, across all three "
         "Types. This becomes the shortlist Step 3's ranking dropdown "
         "offers — nothing else.", None, False),
        ("3. 'Step 3 - Priorities & Ranking' — for each team, pick "
         "priorities from your Step 2 shortlist (only priorities matching "
         "the team's Priority Type will be offered), rank them, mark "
         "whether the team is actually resourcing it, and what share of "
         "the team's effort it gets.", None, False),
        ("4. 'Step 4 - Resource Allocation' — pick each person from "
         "Management's Resources list (only people associated with your "
         "centre will be offered), which team they're on, and what share "
         "of their time goes to that team. Their position fills in "
         "automatically.", None, False),
        ("", None, False),
        ("Worked example", 12, True),
        ("Say your centre has a team called 'Analytics Squad', dedicated "
         "(Yes) and working Tactical priorities. On 'Step 1 - Teams' you'd "
         "enter: Team Name = Analytics Squad, Resources Dedicated = Yes, "
         "Priority Type = Tactical. On 'Step 2', you'd add a row: Priority "
         "Type = Tactical, Priority Title = (pick one of the Tactical "
         "priorities offered) — do this for every priority your centre is "
         "considering, of any Type. On 'Step 3', a row for that team might "
         "then read: Priority Title = (pick one of the Tactical priorities "
         "*you selected in Step 2*), Rank = 1, Resourced = Yes, Allocation "
         "% = 60%. On 'Step 4', a person on that team might read: "
         "Resource = (pick them), Team Name = Analytics Squad, "
         "Allocation % = 50%.", None, False),
        ("", None, False),
        ("Legend", 12, True),
        ("Yellow cells = type or pick your data here.", None, False),
        ("Grey 'check' columns = calculated automatically. Don't type in them.", None, False),
        ("Grey-header sheets = read-only reference data from Management.", None, False),
        ("", None, False),
        ("Validation built into this workbook (no macros required)", 12, True),
        ("- Dropdowns keep Team, Type, Priority, Resource, and Position "
         "entries consistent with the master lists, and constrain "
         "Allocation % entries to 5% steps.", None, False),
        ("- The Priority dropdown is filtered to the team's Priority Type "
         "automatically.", None, False),
        ("- Excel will refuse a duplicate team name or a duplicate rank "
         "within a team as you type.", None, False),
        ("- Going over 100% (a team's priorities, or a person's time) isn't "
         "blocked — the dropdown can't check that and still offer a clean "
         "list of 5% steps — but the Total % columns turn red immediately, "
         "so check them before sending this file back.", None, False),
        ("- The duplicate-name/rank checks fire on manual typing. Pasting "
         "many rows at once can bypass them — glance at the check columns "
         "before sending this file back.", None, False),
        ("", None, False),
        ("Centre Info (auto-filled by Management when this workbook was "
         "generated — not something you need to enter)", 12, True),
    ]
    r = 1
    for text, size, bold in text_blocks:
        c = ws.cell(row=r, column=1, value=text)
        c.font = Font(name=FONT, bold=bool(bold), size=size or 11)
        if text and not bold and size is None:
            c.alignment = Alignment(wrap_text=True)
            ws.row_dimensions[r].height = 30
        r += 1
    # ---- Before you send this back: a live problem count ----------------
    # Previously the only way to know whether the workbook was clean was to
    # scan three sheets row by row looking for red cells. These are the
    # same rules the conditional formatting applies, counted. All four
    # should read 0. Formulas reference Tables created later in this same
    # script -- that's safe, because they all exist by the time the file is
    # saved; the #REF!-corruption trap only bites when a formula references
    # a table that another *stage* creates (see the Step 3/4 "(auto)"
    # columns, deliberately deferred to wire_reference_data.py).
    ws.cell(row=r, column=1, value="Before you send this back").font = Font(
        name=FONT, bold=True, size=12)
    r += 1
    ws.cell(row=r, column=1,
            value=("Every count below should be 0. Anything else is flagged "
                   "in red on the sheet named beside it.")).font = Font(
        name=FONT, italic=True)
    r += 1
    checks = [
        ("Teams whose priorities don't total 100%", "Step 1 / Step 3",
         'SUMPRODUCT((Teams[Team Name]<>"")*(Teams[Priority Allocation Total]<>1))'),
        ("Duplicate team names", "Step 1",
         'SUMPRODUCT((Teams[Team Name]<>"")*'
         '(COUNTIF(Teams[Team Name],Teams[Team Name]&"")>1))'),
        ("Rows sharing a rank within the same team", "Step 3",
         'SUMPRODUCT((Priorities[Team Name]<>"")*(Priorities[Rank]<>"")*'
         '(COUNTIFS(Priorities[Team Name],Priorities[Team Name]&"",'
         'Priorities[Rank],Priorities[Rank]&"")>1))'),
        ("People allocated more than 100% of their time", "Step 4",
         'SUMPRODUCT((ResourceAllocation[Resource]<>"")*'
         '(ResourceAllocation[Person Total %]>1))'),
    ]
    for label, where, formula in checks:
        ws.cell(row=r, column=1, value=label).font = Font(name=FONT)
        c = ws.cell(row=r, column=2, value=f"={formula}")
        style_computed(c)
        c.font = Font(name=FONT, bold=True)
        ws.cell(row=r, column=3, value=where).font = Font(name=FONT, italic=True)
        # Green at 0, red otherwise -- with the same text suffix treatment
        # used on the entry sheets so the status doesn't rely on hue alone.
        ws.conditional_formatting.add(
            f"B{r}",
            Rule(type="expression", formula=[f"$B{r}=0"],
                 dxf=DifferentialStyle(
                     font=Font(name=FONT, bold=True, color=GREEN_TEXT),
                     fill=PatternFill("solid", fgColor=GREEN, bgColor=GREEN),
                     numFmt=NumberFormat(numFmtId=201, formatCode='0" — all clear"'))))
        ws.conditional_formatting.add(
            f"B{r}",
            Rule(type="expression", formula=[f"$B{r}>0"],
                 dxf=DifferentialStyle(
                     font=Font(name=FONT, bold=True, color=RED_TEXT),
                     fill=PatternFill("solid", fgColor=RED, bgColor=RED),
                     numFmt=NumberFormat(numFmtId=202, formatCode='0" to fix"'))))
        r += 1
    ws.column_dimensions["B"].width = 22
    ws.column_dimensions["C"].width = 18
    r += 1

    centre_name_row = r
    ws.cell(row=r, column=1, value="Centre Name:").font = Font(name=FONT, bold=True)
    style_computed(ws.cell(row=r, column=2, value=centre_name or "Example Centre"))
    r += 1
    centre_code_row = r
    ws.cell(row=r, column=1, value="Centre Code:").font = Font(name=FONT, bold=True)
    style_computed(ws.cell(row=r, column=2, value=centre_code or "EX"))

    # Build stamp: which generation of the template this file came from.
    # These workbooks travel to a machine with no git, no network and no
    # repo access, so without this there is no way to answer "which
    # version is that?" about a file someone is holding. Exposed as a
    # defined name as well as a visible cell, so the Consolidation
    # workbook can read it back out of every returned file (see
    # wire_power_query.py's Refresh Status query) and Management can spot
    # a centre still working in a stale copy.
    r += 1
    build_stamp_row = r
    ws.cell(row=r, column=1, value="Workbook build:").font = Font(name=FONT, bold=True)
    style_computed(ws.cell(row=r, column=2, value=build_stamp()))
    ws.cell(row=r, column=3,
            value="Quote this if you report a problem with this workbook.").font = Font(
        name=FONT, italic=True)

    wb.defined_names["CentreName"] = DefinedName(
        "CentreName", attr_text=f"Instructions!$B${centre_name_row}")
    wb.defined_names["CentreCode"] = DefinedName(
        "CentreCode", attr_text=f"Instructions!$B${centre_code_row}")
    wb.defined_names["BuildStamp"] = DefinedName(
        "BuildStamp", attr_text=f"Instructions!$B${build_stamp_row}")

    # A one-row Table, not a plain cell: Power Query's standard parameter
    # pattern is Excel.CurrentWorkbook(){[Name="Config"]}[Content]{0}[PrepFilePath]
    # — a named Table survives being moved around the sheet; a hardcoded
    # cell address doesn't. Read by wire_reference_data.py's queries (run
    # as stage 2, after this file is saved) so Management can repoint this
    # workbook at wherever preparation.xlsx currently lives — e.g. after
    # moving it to a shared drive — by editing this one cell and hitting
    # Data > Refresh All, with no Python involved. Same pattern already
    # used by management/consolidation/build_consolidation.py's
    # SourceFolder config.
    r += 1
    config_header_row = r
    style_header(ws.cell(row=config_header_row, column=1, value="PrepFilePath"))
    style_body(ws.cell(row=config_header_row + 1, column=1,
                        value=str(Path(prep_path).resolve())), editable=True)
    tab = Table(displayName="Config", ref=f"A{config_header_row}:A{config_header_row + 1}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    ws.protection.sheet = True

    # ================================================================= Reference sheets (read-only)
    # Centres, Position, ProblemSet, InitiativeType, AssistanceType,
    # Resources, Priority (split 3 ways), and the Tactical/Initiative/
    # Assistance scoring tables are NOT built here as static openpyxl
    # copies any more (2026-09-01) — that was a "known process gap"
    # (see CLAUDE.md) against the Process section's original intent of a
    # live refresh. They're now Power-Query-authored against
    # preparation.xlsx by templates/scripts/wire_reference_data.py (COM,
    # same recipe as wire_power_query.py's Consolidation combine queries),
    # run as a second stage after this script — see that file's docstring
    # for the full recipe, the Priority-split rationale, and why
    # RatingLookup deliberately stays a static copy here instead. This
    # main() call sequence assumes wire_reference_data.py runs before the
    # workbook is handed to anyone: it creates those sheets and their
    # dependent named ranges (ResourceNameList; Tactical/Initiative/
    # Assistance for the Step 2 dependent dropdown), none of which exist
    # yet at the point this function returns.
    rl_headers, rl_rows = read_table(src, "RatingLookup", header_row=3)
    ws = wb.create_sheet("RatingLookup")
    write_reference_table(ws, "RatingLookup", 1, rl_headers, rl_rows,
                           [6, 14, 10, 10, 14, 12])
    for i, row in enumerate(rl_rows):
        ws.cell(row=2 + i, column=7).fill = PatternFill("solid", fgColor=row[5])
    ws.column_dimensions["G"].width = 4
    # RatingLookup is the one reference sheet that does NOT update on
    # Data > Refresh All -- it's a deliberate static copy (see
    # wire_reference_data.py's docstring for why). It looks identical to
    # every sheet that does refresh, so without this note the failure mode
    # is silent: Management edits a band in preparation.xlsx, refreshes,
    # and this sheet quietly keeps the old values.
    warn = ws.cell(row=1, column=9,
                   value=("Note: unlike the other reference sheets, this one does NOT "
                          "update on Data > Refresh All. It is a fixed copy taken when "
                          "this workbook was generated. Changing a band's thresholds, "
                          "name or colour needs a regenerated workbook — ask whoever "
                          "generated this file."))
    warn.font = Font(name=FONT, bold=True, color=AMBER_TEXT)
    warn.fill = PatternFill("solid", fgColor=AMBER)
    warn.alignment = Alignment(wrap_text=True, vertical="top")
    warn.border = BORDER
    ws.merge_cells(start_row=1, start_column=9, end_row=4, end_column=13)
    for col in "IJKLM":
        ws.column_dimensions[col].width = 22

    # ================================================================= Lookups (small enums, editable-free)
    ws = wb.create_sheet("Lookups")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Small enum lists — used to power dropdowns."
    ws["A1"].font = Font(name=FONT, bold=True)
    ws["A3"] = "Priority Type"
    ws["A3"].font = Font(name=FONT, bold=True)
    # Every literal enum list below is written in alphabetical order —
    # since these are the actual dropdown source cells (not a formula
    # pointing elsewhere), sorting them is just sorting the list once here.
    for i, val in enumerate(sorted(["Tactical", "Initiative", "Assistance"])):
        ws.cell(row=4 + i, column=1, value=val).font = Font(name=FONT)
    wb.defined_names["PriorityTypeList"] = DefinedName(
        "PriorityTypeList", attr_text="Lookups!$A$4:$A$6")
    ws["C3"] = "Resource Status"
    ws["C3"].font = Font(name=FONT, bold=True)
    for i, val in enumerate(sorted(["Yes", "No", "Temp"])):
        ws.cell(row=4 + i, column=3, value=val).font = Font(name=FONT)
    wb.defined_names["ResourceStatusList"] = DefinedName(
        "ResourceStatusList", attr_text="Lookups!$C$4:$C$6")
    ws["E3"] = "Yes/No"
    ws["E3"].font = Font(name=FONT, bold=True)
    for i, val in enumerate(sorted(["Yes", "No"])):
        ws.cell(row=4 + i, column=5, value=val).font = Font(name=FONT)
    wb.defined_names["YesNoList"] = DefinedName("YesNoList", attr_text="Lookups!$E$4:$E$5")
    ws["G3"] = "Allocation % (5% steps)"
    ws["G3"].font = Font(name=FONT, bold=True)
    for i in range(21):  # 0%, 5%, ..., 100%
        cell = ws.cell(row=4 + i, column=7, value=i * 0.05)
        cell.font = Font(name=FONT)
        cell.number_format = "0%"
    wb.defined_names["PercentIncrementsList"] = DefinedName(
        "PercentIncrementsList", attr_text="Lookups!$G$4:$G$24")
    for col, w in [("A", 16), ("C", 16), ("E", 10), ("G", 20)]:
        ws.column_dimensions[col].width = w

    last_row = N_ROWS + 1

    # ================================================================= Step 1 - Teams
    ws = wb.create_sheet("Step 1 - Teams")
    ws.sheet_view.showGridLines = False
    headers = ["Team Name", "Resources Dedicated", "Priority Type",
               "Priority Allocation Total", "Resource Effort Total (FTE)",
               "Team Name Sorted"]
    widths = [22, 20, 18, 24, 26, 22]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        style_header(ws.cell(row=1, column=i, value=h))
    ws.column_dimensions["F"].hidden = True

    # Teams is a real Table spanning all 15 team slots (rows 2-16, CLAUDE.md
    # caps this sheet at 15 teams so it "stays clean") from generation time
    # — NOT a single official row backed by a styled "buffer" that Excel
    # auto-extends to absorb. That smaller design was tried first and
    # confirmed broken (2026-09-01): typing into the row below a Table does
    # NOT auto-extend it while the sheet is protected — verified against a
    # live Excel session (Table.Range stayed $A$1:$E$2 after setting A3,
    # and TeamNameList never picked up the second row) — on top of the
    # already-documented finding that the native "Insert Table Row"/
    # ListRows.Add() mechanism also fails under protection. With no
    # protected-sheet mechanism to grow a Table at all, a Centre Lead's
    # data in any row past the Table's official range was silently invisible
    # to TeamNameList, the Step 3 team dropdown, and every SUMIFS/lookup
    # against Teams[...]. Building the full 15 rows in up front (same
    # pattern as Steps 2-4, which use the real N_ROWS-row Table already)
    # fixes that outright, at the cost of reopening the Data Model
    # relationship blocker below for Teams specifically: Power Pivot
    # rejects a relationship whose "one" side contains any blank, and with
    # 15 rows always in the Table, any unused team slot stays blank
    # forever (a real per-centre team count is very unlikely to hit 15).
    # Accepted: nothing shipped depends on that relationship today — see
    # CLAUDE.md's Status entry on the Resource x Team matrix PivotTable.
    teams_last_row = 16
    for r in range(2, teams_last_row + 1):
        style_body(ws.cell(row=r, column=1), editable=True)
        style_body(ws.cell(row=r, column=2), editable=True)
        style_body(ws.cell(row=r, column=3), editable=True)
        for col in (4, 5):
            cell = ws.cell(row=r, column=col)
            style_computed(cell)
            cell.number_format = "0%"
        ws.cell(row=r, column=4).value = (
            f'=IF($A{r}="","",SUMIFS(Priorities[Allocation % of Team Effort],'
            f'Priorities[Team Name],$A{r}))'
        )
        ws.cell(row=r, column=5).value = (
            f'=IF($A{r}="","",SUMIFS(ResourceAllocation[Allocation % of Person Effort],'
            f'ResourceAllocation[Team Name],$A{r}))'
        )
        # Hidden helper: the up-to-15 team names, alphabetised and with the
        # unused blank slots pushed out entirely, one real team name per
        # row starting at row 2. FILTER()/SORT() are dynamic-array
        # functions, but wrapping them in INDEX() to pull a single element
        # keeps the *result* a plain scalar in each cell — no spill, so
        # this is just an ordinary (if unusually powerful) cell formula,
        # not the dynamic-array-as-defined-name approach that was already
        # tried and confirmed unusable as a direct Data Validation source
        # (see the note on Selected{Type} below). TeamNameList points at
        # this column, not at Team Name directly.
        #
        # The `_xlfn.` prefix on SORT/FILTER below is required — confirmed
        # by testing (2026-09-04): a formula openpyxl writes with the bare
        # function names looks fine in the XML (well-formed, no typo) but
        # Excel refuses to even Open() the file at all (COM error, no
        # repair dialog, not the more familiar "silently rewrites to
        # #REF!" failure mode this codebase hit before). Opening it once
        # with CorruptLoad/xlRepairFile confirmed Excel's own repair step
        # was simply deleting these formulas outright. `_xlfn.` alone
        # (without the `_xlws.` half also needed for a *defined name*
        # formula — see the Priority Selection section of
        # excel-file-design.md) was sufficient here since these are plain
        # cell formulas, not defined names.
        style_computed(ws.cell(row=r, column=6))
        ws.cell(row=r, column=6).value = (
            f'=IFERROR(INDEX(_xlfn.SORT(_xlfn.FILTER($A$2:$A${teams_last_row},'
            f'$A$2:$A${teams_last_row}<>"")),ROW()-1),"")'
        )

    tab = Table(displayName="Teams", ref=f"A1:F{teams_last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)
    wb.defined_names["TeamNameList"] = DefinedName(
        "TeamNameList", attr_text="Teams[Team Name Sorted]")

    dv_resources = DataValidation(type="list", formula1="=ResourceStatusList", allow_blank=True)
    dv_resources.error = "Choose Yes, No, or Temp."
    dv_resources.errorTitle = "Invalid entry"
    dv_resources.promptTitle = "Dedicated resources?"
    dv_resources.prompt = (
        "Yes = people are assigned to this team. Temp = borrowed or short-term. No = the team exists but has nobody on it yet.")
    ws.add_data_validation(dv_resources)
    dv_resources.add(f"B2:B{teams_last_row}")

    dv_type = DataValidation(type="list", formula1="=PriorityTypeList", allow_blank=True)
    dv_type.error = "Choose Tactical, Initiative, or Assistance."
    dv_type.errorTitle = "Invalid entry"
    dv_type.promptTitle = "Priority Type"
    dv_type.prompt = (
        "A team works on only ONE Priority Type. If a team really spans two, list it twice under different names.")
    ws.add_data_validation(dv_type)
    dv_type.add(f"C2:C{teams_last_row}")

    # Plain range, not Teams[Team Name]: a structured reference here (valid
    # syntax, verified by hand) made Excel strip Data Validation from this
    # sheet on open — Data Validation formulas apparently don't support
    # Table structured references the way regular cell formulas do.
    dv_unique_team = DataValidation(
        type="custom",
        formula1=f'=AND($A2<>"",COUNTIF($A$2:$A${teams_last_row},$A2)=1)',
        allow_blank=True,
    )
    dv_unique_team.error = "Team names must be unique within this workbook."
    dv_unique_team.errorTitle = "Duplicate team name"
    dv_unique_team.promptTitle = "Team name"
    dv_unique_team.prompt = (
        "Type this team's name. It must be unique within this workbook, and it's what Steps 3 and 4 will offer you in their Team dropdowns.")
    ws.add_data_validation(dv_unique_team)
    dv_unique_team.add(f"A2:A{teams_last_row}")

    ws.protection.sheet = True

    status_format(ws, f"D2:D{teams_last_row}", "$A2", "$D2", include_under=True)
    paste_guard(ws, f"A2:F{teams_last_row}",
                f'AND($A2<>"",COUNTIF($A$2:$A${teams_last_row},$A2)>1)')
    ws.freeze_panes = "A2"

    # ================================================================= Step 2 - Select Priorities
    ws = wb.create_sheet("Step 2 - Select Priorities")
    ws.sheet_view.showGridLines = False
    # Columns C-E are helper columns, not something a Centre Lead fills in
    # — see the note below on why they exist instead of a FILTER() formula
    # — hidden so they don't clutter the sheet.
    headers = ["Priority Type", "Priority Title",
               "Tactical Helper", "Initiative Helper", "Assistance Helper"]
    widths = [16, 50, 12, 12, 12]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        style_header(ws.cell(row=1, column=i, value=h))
    for col_letter in ("C", "D", "E"):
        ws.column_dimensions[col_letter].hidden = True

    for r in range(2, last_row + 1):
        style_body(ws.cell(row=r, column=1), editable=True)
        style_body(ws.cell(row=r, column=2), editable=True)
        # One column per Type: this Type's selected Priority Titles,
        # alphabetised, with unselected/wrong-Type/not-yet-used rows
        # pushed out entirely rather than left as blanks in place. Confirmed
        # by testing (2026-09-03) that Excel's Data Validation "List" type
        # cannot consume a dynamic-array (FILTER()-based) defined name at
        # all when the name itself resolves to a spilling array — fails
        # even referenced directly with no INDIRECT involved (Validation.Add
        # itself throws). The fix isn't to avoid FILTER()/SORT() though —
        # it's to wrap them in INDEX(), which pulls a single element back
        # out to a plain scalar and never spills, so each cell holds an
        # ordinary formula result like any other. An earlier version of
        # this column used a plain IF() per row and left real blanks in
        # place for non-matching rows, on the theory that Excel's dropdown
        # silently skips blank cells in a list source — it does not; a
        # Centre Lead reported the dropdown showing a run of several blank
        # entries below the real ones, which is exactly what that IF()
        # produced (an untouched blank for every one of the 200 rows that
        # wasn't this Type). This INDEX/SORT/FILTER version compacts those
        # away instead of just hiding them individually. See the matching
        # `_xlfn.` note on Step 1 - Teams' own sorted helper column above —
        # same requirement applies here, for the same reason.
        for col, ptype in ((3, "Tactical"), (4, "Initiative"), (5, "Assistance")):
            cell = ws.cell(row=r, column=col)
            style_computed(cell)
            cell.value = (
                f'=IFERROR(INDEX(_xlfn.SORT(_xlfn.FILTER($B$2:$B${last_row},'
                f'($A$2:$A${last_row}="{ptype}")*($B$2:$B${last_row}<>""))),'
                f'ROW()-1),"")'
            )

    tab = Table(displayName="PrioritySelection", ref=f"A1:E{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    dv_type_select = DataValidation(type="list", formula1="=PriorityTypeList", allow_blank=True)
    dv_type_select.error = "Choose Tactical, Initiative, or Assistance."
    dv_type_select.errorTitle = "Invalid entry"
    dv_type_select.promptTitle = "Pick the Type first"
    dv_type_select.prompt = (
        "Choose the Type here first — the Priority Title dropdown next to it then narrows to just that Type.")
    ws.add_data_validation(dv_type_select)
    dv_type_select.add(f"A2:A{last_row}")

    # Same dependent-dropdown trick as the Ranking sheet below: pick Type
    # first, Priority Title is then filtered to that Type's full master
    # list (the plain Tactical/Initiative/Assistance defined names, still
    # the complete list here — this sheet is where a Centre Lead narrows
    # it down for the first time).
    dv_priority_select = DataValidation(type="list", formula1="=INDIRECT($A2)", allow_blank=True)
    dv_priority_select.error = "Pick a priority from Management's master list, matching the Type."
    dv_priority_select.errorTitle = "Unknown priority"
    dv_priority_select.promptTitle = "Priority Title"
    dv_priority_select.prompt = (
        "Pick from Management's master list for the Type you chose. Everything you pick here becomes the shortlist Step 3 offers.")
    ws.add_data_validation(dv_priority_select)
    dv_priority_select.add(f"B2:B{last_row}")

    # Plain range, not PrioritySelection[Priority Title] — see the note on
    # dv_unique_team above (structured refs break Data Validation).
    dv_unique_priority = DataValidation(
        type="custom",
        formula1=f'=AND($B2<>"",COUNTIF($B$2:$B${last_row},$B2)=1)',
        allow_blank=True,
    )
    dv_unique_priority.error = "Each priority can only be selected once."
    dv_unique_priority.errorTitle = "Duplicate priority"
    dv_unique_priority.promptTitle = "Priority Title"
    dv_unique_priority.prompt = (
        "Pick from Management's master list for the Type you chose. Each priority can only appear once on this sheet.")
    ws.add_data_validation(dv_unique_priority)
    dv_unique_priority.add(f"B2:B{last_row}")

    ws.protection.sheet = True
    ws.freeze_panes = "A2"

    # Selected-and-Type-filtered lists for Step 3's Ranking dropdown below.
    # Structured references, same proven pattern as TeamNameList
    # (`Teams[Team Name Sorted]`) — a defined name built on a structured
    # reference works fine as a Data Validation list source; it's only a
    # structured reference typed *directly* into a DataValidation formula
    # that breaks (see the note on dv_unique_team). The helper columns
    # themselves are already alphabetised and blank-compacted (see the
    # INDEX/SORT/FILTER note on the loop above) — a raw list-type range
    # does NOT skip blank cells, it shows one dropdown entry per cell
    # including the blank ones, which is what the original IF()-only
    # version of this column actually did in practice.
    for ptype, col_name in (("Tactical", "Tactical Helper"),
                             ("Initiative", "Initiative Helper"),
                             ("Assistance", "Assistance Helper")):
        wb.defined_names[f"Selected{ptype}"] = DefinedName(
            f"Selected{ptype}", attr_text=f"PrioritySelection[{col_name}]",
        )

    # ================================================================= Step 3 - Priorities & Ranking
    ws = wb.create_sheet("Step 3 - Priorities & Ranking")
    ws.sheet_view.showGridLines = False
    headers = ["Team Name", "Priority Title", "Type (auto)", "Rank", "Resourced",
               "Allocation % of Team Effort", "Value/Risk (auto)",
               "Team Total % (auto)"]
    widths = [22, 42, 16, 10, 12, 22, 16, 18]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        style_header(ws.cell(row=1, column=i, value=h))

    for r in range(2, last_row + 1):
        style_body(ws.cell(row=r, column=1), editable=True)
        style_body(ws.cell(row=r, column=2), editable=True)
        c3 = ws.cell(row=r, column=3)
        style_computed(c3)
        c3.value = f'=IF($A{r}="","",IFERROR(INDEX(Teams[Priority Type],MATCH($A{r},Teams[Team Name],0)),"unknown team"))'
        style_body(ws.cell(row=r, column=4), editable=True)
        style_body(ws.cell(row=r, column=5), editable=True)
        style_body(ws.cell(row=r, column=6), editable=True)
        ws.cell(row=r, column=6).number_format = "0%"
        # Running total of THIS row's team across all its priorities, shown
        # in-row. The same number is also on Step 1 (Teams' "Priority
        # Allocation Total"), but the 100%-per-team rule is the main thing
        # a Centre Lead is trying to satisfy while typing here, and having
        # to flip to another sheet to find out whether they'd hit it was
        # the single most-repeated bit of friction in this workbook. Same
        # SUMIFS + red/amber/green treatment Step 4's "Person Total %"
        # already used; deliberately repeats rather than cross-references
        # so each row is self-explanatory.
        c8 = ws.cell(row=r, column=8)
        style_computed(c8)
        c8.number_format = "0%"
        c8.value = (
            f'=IF($A{r}="","",SUMIFS(Priorities[Allocation % of Team Effort],'
            f'Priorities[Team Name],$A{r}))'
        )
        # Value/Risk band for the selected priority — looked up from
        # whichever of the three embedded scoring tables matches this row's
        # own Type (auto), so Centre Leads see it right where they're
        # deciding how to rank, not just after the fact.
        # NOT written here: this formula references TacticalScores/
        # InitiativeScores/AssistanceScores, which don't exist until stage
        # 2 (wire_reference_data.py) creates them. Writing it here (as an
        # earlier version did) means the very first time Excel opens this
        # stage-1-only file, it can't resolve those table names and
        # PERMANENTLY rewrites the structured references to #REF! in the
        # formula text itself -- creating the tables afterward does not
        # un-corrupt an already-#REF!'d formula. wire_reference_data.py
        # sets this formula (same text, via a single bulk Range.Formula
        # assignment) once those 3 tables genuinely exist.
        style_computed(ws.cell(row=r, column=7))

    tab = Table(displayName="Priorities", ref=f"A1:H{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    dv_team2 = DataValidation(type="list", formula1="=TeamNameList", allow_blank=True)
    dv_team2.error = "Pick a team defined on 'Step 1 - Teams'."
    dv_team2.errorTitle = "Unknown team"
    dv_team2.promptTitle = "Team name"
    dv_team2.prompt = (
        "Pick a team you listed in Step 1. Its Priority Type decides which priorities you'll be offered next.")
    ws.add_data_validation(dv_team2)
    dv_team2.add(f"A2:A{last_row}")

    # Dependent dropdown: the list source is INDIRECT("Selected" & this
    # row's auto Type), resolving to the FILTER()-based Selected{Type}
    # defined name from Step 2 — Step 2's shortlist, not Management's full
    # master list, and further narrowed to the team's own Priority Type.
    # $C2 itself still holds the plain "Tactical"/"Initiative"/"Assistance"
    # text (other formulas/conditional formatting on this sheet compare
    # against it directly), so the "Selected" prefix is added here rather
    # than changing what $C2 stores.
    dv_priority = DataValidation(type="list", formula1='=INDIRECT("Selected"&$C2)', allow_blank=True)
    dv_priority.error = "Pick a priority from your Step 2 selection, matching the team's Priority Type."
    dv_priority.errorTitle = "Unknown priority"
    dv_priority.promptTitle = "Priority Title"
    dv_priority.prompt = (
        "Only priorities you selected in Step 2 that match this team's Priority Type are offered here.")
    ws.add_data_validation(dv_priority)
    dv_priority.add(f"B2:B{last_row}")

    dv_resourced = DataValidation(type="list", formula1="=YesNoList", allow_blank=True)
    dv_resourced.error = "Choose Yes or No."
    dv_resourced.errorTitle = "Invalid entry"
    dv_resourced.promptTitle = "Resourced?"
    dv_resourced.prompt = (
        "Is the team actually putting people against this priority right now? Yes or No.")
    ws.add_data_validation(dv_resourced)
    dv_resourced.add(f"E2:E{last_row}")

    # Plain ranges here too — see the note on dv_unique_team above.
    dv_rank = DataValidation(
        type="custom",
        formula1=(f'=AND($D2<>"",$D2=INT($D2),$D2>0,'
                  f'COUNTIFS($A$2:$A${last_row},$A2,$D$2:$D${last_row},$D2)=1)'),
        allow_blank=True,
    )
    dv_rank.error = "Rank must be a positive whole number, unique within the team."
    dv_rank.errorTitle = "Invalid rank"
    dv_rank.promptTitle = "Rank"
    dv_rank.prompt = (
        "1 = this team's highest priority. Whole numbers only, and each rank can be used once per team.")
    ws.add_data_validation(dv_rank)
    dv_rank.add(f"D2:D{last_row}")

    # A dropdown (list) and a hard "team total <=100%" block can't coexist
    # in one native Data Validation rule — this picks the dropdown; the
    # Priority Allocation Total column on Step 1 still flags an overage
    # visually (red), just not as a typing-time block.
    dv_alloc2 = DataValidation(type="list", formula1="=PercentIncrementsList", allow_blank=True)
    dv_alloc2.error = "Pick a value from the dropdown (5% steps)."
    dv_alloc2.errorTitle = "Invalid entry"
    dv_alloc2.promptTitle = "Share of team effort"
    dv_alloc2.prompt = (
        "What share of THIS TEAM's total effort goes to this priority. A team's priorities should add up to exactly 100% — watch the Team Total % column at the end of the row.")
    ws.add_data_validation(dv_alloc2)
    dv_alloc2.add(f"F2:F{last_row}")

    # Risk bands and Value bands share band names ("Minimal") with opposite
    # meaning — Risk Minimal is green (good, low risk), Value Minimal is red
    # (bad, low value) — so each rule must check Type (column C), not just
    # match the Label text in column G, or a Tactical and an
    # Initiative/Assistance row with the same band name would get each
    # other's colour.
    for band_name, colour, text_colour in RISK_BANDS:
        ws.conditional_formatting.add(
            f"G2:G{last_row}",
            FormulaRule(formula=[f'AND($C2="Tactical",$G2="{band_name}")'],
                        fill=PatternFill("solid", fgColor=colour, bgColor=colour),
                        font=Font(name=FONT, color=text_colour))
        )
    for band_name, colour, text_colour in VALUE_BANDS:
        ws.conditional_formatting.add(
            f"G2:G{last_row}",
            FormulaRule(formula=[f'AND($C2<>"Tactical",$C2<>"",$G2="{band_name}")'],
                        fill=PatternFill("solid", fgColor=colour, bgColor=colour),
                        font=Font(name=FONT, color=text_colour))
        )
    status_format(ws, f"H2:H{last_row}", "$A2", "$H2", include_under=True)
    # A team's priorities must total exactly 100%, so "under" is a real
    # amber warning here (unlike Step 4, where a person under 100% is fine).
    paste_guard(ws, f"A2:H{last_row}",
                'AND($A2<>"",$D2<>"",'
                f'COUNTIFS($A$2:$A${last_row},$A2,$D$2:$D${last_row},$D2)>1)')

    ws.protection.sheet = True
    ws.freeze_panes = "A2"

    # ================================================================= Step 4 - Resource Allocation
    ws = wb.create_sheet("Step 4 - Resource Allocation")
    ws.sheet_view.showGridLines = False
    headers = ["Resource", "Position Title (auto)", "Team Name",
               "Allocation % of Person Effort", "Person Total %"]
    widths = [26, 26, 22, 24, 16]
    for i, (h, w) in enumerate(zip(headers, widths), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        style_header(ws.cell(row=1, column=i, value=h))

    for r in range(2, last_row + 1):
        style_body(ws.cell(row=r, column=1), editable=True)
        # NOT written here: this formula references Resources, which
        # doesn't exist until stage 2 (wire_reference_data.py) creates it
        # -- see the matching note on Step 3's Value/Risk (auto) column
        # above for why writing a cross-stage table reference in stage 1
        # would permanently corrupt it to #REF!.
        style_computed(ws.cell(row=r, column=2))
        style_body(ws.cell(row=r, column=3), editable=True)
        style_body(ws.cell(row=r, column=4), editable=True)
        ws.cell(row=r, column=4).number_format = "0%"
        c5 = ws.cell(row=r, column=5)
        style_computed(c5)
        c5.number_format = "0%"
        c5.value = (
            f'=IF($A{r}="","",SUMIFS(ResourceAllocation[Allocation % of Person Effort],'
            f'ResourceAllocation[Resource],$A{r}))'
        )

    tab = Table(displayName="ResourceAllocation", ref=f"A1:E{last_row}")
    tab.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tab)

    dv_res_name = DataValidation(type="list", formula1="=ResourceNameList", allow_blank=True)
    dv_res_name.error = "Pick a person from Management's Resources list."
    dv_res_name.errorTitle = "Unknown resource"
    dv_res_name.promptTitle = "Resource"
    dv_res_name.prompt = (
        "Only people associated with your centre are listed. If someone's missing, ask Management to associate them — don't type a name in by hand.")
    ws.add_data_validation(dv_res_name)
    dv_res_name.add(f"A2:A{last_row}")

    dv_team3 = DataValidation(type="list", formula1="=TeamNameList", allow_blank=True)
    dv_team3.error = "Pick a team defined on 'Step 1 - Teams'."
    dv_team3.errorTitle = "Unknown team"
    dv_team3.promptTitle = "Team name"
    dv_team3.prompt = (
        "Which of your Step 1 teams is this person working on?")
    ws.add_data_validation(dv_team3)
    dv_team3.add(f"C2:C{last_row}")

    # Dropdown, not a hard block — see the note on dv_alloc2 in Step 3.
    # Person Total % still flags an overage visually (red).
    dv_alloc3 = DataValidation(type="list", formula1="=PercentIncrementsList", allow_blank=True)
    dv_alloc3.error = "Pick a value from the dropdown (5% steps)."
    dv_alloc3.errorTitle = "Invalid entry"
    dv_alloc3.promptTitle = "Share of this person's time"
    dv_alloc3.prompt = (
        "What share of THIS PERSON's time goes to this team. It doesn't have to reach 100%, but it must not exceed it — watch the Person Total % column.")
    ws.add_data_validation(dv_alloc3)
    dv_alloc3.add(f"D2:D{last_row}")

    ws.protection.sheet = True

    # include_under=False: a person's time across teams doesn't have to
    # reach 100% (they may be part-allocated), so under-100% isn't a
    # warning here — only over-100% is. Same rule as before this became a
    # shared helper.
    status_format(ws, f"E2:E{last_row}", "$A2", "$E2", include_under=False)
    paste_guard(ws, f"A2:E{last_row}",
                f'AND($A2<>"",$C2<>"",'
                f'COUNTIFS($A$2:$A${last_row},$A2,$C$2:$C${last_row},$C2)>1)')
    ws.freeze_panes = "A2"

    # ================================================================= Ranked View (read-only)
    # CLAUDE.md's UI requirements ask for the ranking "ideally sorted as
    # Rank is entered". Sorting the entry Table itself as someone types is
    # not possible without VBA (and would fight the fixed-size, protected
    # Table design anyway), so this is the read-only companion instead: a
    # live, always-sorted view of whatever Step 3 currently holds, grouped
    # by team then rank.
    #
    # Per-cell INDEX over the sorted array, NOT one spilling formula. A
    # spilling `=SORT(FILTER(...))` was tried first and does not survive
    # being written by openpyxl: Excel applies implicit intersection to it
    # on open and the cell returns just the top-left value instead of
    # spilling (verified — the sheet showed one team name and nothing
    # else). Marking it as a legacy CSE array over a fixed range would
    # spill, but pads every unused cell with #N/A, which this project
    # treats as a defect. Wrapping in INDEX(...,row,col) returns a plain
    # scalar per cell, so nothing needs to spill at all — the same
    # already-proven trick the Step 1/Step 2 dropdown helper columns use.
    ws = wb.create_sheet("Ranked View")
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Your Step 3 priorities, sorted by team then rank"
    ws["A1"].font = Font(name=FONT, bold=True, size=12)
    ws["A2"] = ("Read-only — this updates itself from 'Step 3 - Priorities & "
                "Ranking'. Nothing to fill in here.")
    ws["A2"].font = Font(name=FONT, italic=True)
    ranked_headers = ["Team Name", "Priority Title", "Type", "Rank", "Resourced",
                      "Allocation %", "Value/Risk"]
    for i, (h, w) in enumerate(zip(ranked_headers,
                                   [22, 42, 16, 10, 12, 14, 16]), start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
        style_header(ws.cell(row=4, column=i, value=h))

    # SORT's sort_index takes an array ({1,4} = Team Name, then Rank), so a
    # single call covers both sort levels.
    sorted_array = (
        '_xlfn.SORT(_xlfn._xlws.FILTER(Priorities[[Team Name]:[Value/Risk (auto)]],'
        '(Priorities[Team Name]<>"")*(Priorities[Rank]<>"")),{1,4},{1,1})'
    )
    for i in range(RANKED_ROWS):
        r = 5 + i
        for c in range(1, len(ranked_headers) + 1):
            cell = ws.cell(row=r, column=c)
            style_computed(cell)
            cell.value = f'=IFERROR(INDEX({sorted_array},ROW()-4,{c}),"")'
        ws.cell(row=r, column=4).number_format = "0"
        ws.cell(row=r, column=6).number_format = "0%"

    # RANKED_ROWS is smaller than the 200-row Step 3 Table (recomputing the
    # sort per cell is the cost of not spilling, so this stays bounded).
    # 15 teams is the documented cap and ~50 priorities exist in total, so
    # overflowing this is implausible — but say so rather than truncating
    # silently if it ever happens.
    ws.cell(row=3, column=1, value=(
        f'=IF(COUNTA(Priorities[Team Name])>{RANKED_ROWS},'
        f'"More than {RANKED_ROWS} ranked rows exist — only the first '
        f'{RANKED_ROWS} are shown here. Step 3 still holds them all.","")'
    )).font = Font(name=FONT, bold=True, color=RED_TEXT)
    ws.protection.sheet = True
    ws.freeze_panes = "A5"

    # Final tab order (interleaving the Power-Query-authored reference
    # sheets from wire_reference_data.py with these) is set there, once
    # those sheets exist — same split as create_pivots.py owning
    # Consolidation's final sheet order. Here, just put what this stage
    # actually built in a sane order for anyone opening the file between
    # the two stages.
    SHEET_ORDER = ["Instructions", "RatingLookup", "Lookups",
                   "Step 1 - Teams", "Step 2 - Select Priorities",
                   "Step 3 - Priorities & Ranking", "Step 4 - Resource Allocation",
                   "Ranked View"]
    wb._sheets = [wb[name] for name in SHEET_ORDER]

    # Tab colours: the user guides have always described "blue tabs are
    # yours, grey tabs are Management's reference data", but the tabs
    # themselves were all default-white — the one place a Centre Lead
    # actually navigates from carried none of that distinction. Reference
    # sheets built in stage 2 get the same treatment in
    # wire_reference_data.py.
    wb["Instructions"].sheet_properties.tabColor = TAB_INSTRUCTIONS
    for name in ("Step 1 - Teams", "Step 2 - Select Priorities",
                 "Step 3 - Priorities & Ranking", "Step 4 - Resource Allocation"):
        wb[name].sheet_properties.tabColor = TAB_ENTRY
    for name in ("RatingLookup", "Lookups", "Ranked View"):
        wb[name].sheet_properties.tabColor = TAB_REFERENCE

    # Hide the empty tail of each 200-row entry sheet. The Tables stay
    # full-size (they have to — a Table can't grow under sheet protection,
    # see the Step 1 - Teams note above), this only stops a small centre
    # scrolling past ~170 blank rows to reach the end. Unhiding is the
    # normal Excel gesture and nothing breaks if a Centre Lead does it.
    for name in ("Step 2 - Select Priorities", "Step 3 - Priorities & Ranking",
                 "Step 4 - Resource Allocation"):
        ws_hide = wb[name]
        for r in range(VISIBLE_ROWS + 2, last_row + 1):
            ws_hide.row_dimensions[r].hidden = True

    wb.active = 0
    wb.save(out_path)
    print("saved", out_path)


if __name__ == "__main__":
    if len(sys.argv) < 3:
        raise SystemExit(
            "Usage:\n"
            "  python build_centre_template.py <preparation.xlsx> <output.xlsx> [centre-code]\n"
            "  python build_centre_template.py <preparation.xlsx> --all <output-dir>"
        )
    prep_arg = sys.argv[1]
    if sys.argv[2] == "--all":
        if len(sys.argv) != 4:
            raise SystemExit("Usage: python build_centre_template.py <preparation.xlsx> --all <output-dir>")
        out_dir = Path(sys.argv[3])
        out_dir.mkdir(parents=True, exist_ok=True)
        centres_src = openpyxl.load_workbook(prep_arg, data_only=True)
        _, all_centres = read_table(centres_src, "Centres")
        for _id, c_name, c_code in all_centres:
            main(prep_arg, str(out_dir / f"Centre-{c_code}.xlsx"), centre_name=c_name, centre_code=c_code)
    else:
        main(prep_arg, sys.argv[2], centre_code=sys.argv[3] if len(sys.argv) > 3 else None)
