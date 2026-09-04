"""
Dev-only fixture generator: produces a handful of small, hand-checkable
"completed" centre workbooks (Teams/Priorities/ResourceAllocation filled in)
for testing the Consolidation workbook against — nothing like this existed
anywhere in the repo, since real Centre Lead submissions don't exist yet.

Output goes to dev_fixtures/output/ (gitignored) — these are throwaway,
regenerate any time by re-running this script. Not a shipped deliverable.

Reuses build_centre_template.main() (so fixtures go through the real,
current template-generation code path, including the Centre Name/Code
stamping) and read_table() to pull real Priority/Resource names out of
preparation.xlsx rather than inventing fake ones.

Deliberately plants a Team Name collision across two centres (both ECO and
INF have a team called "Ops Team") — this is the exact cross-centre
collision the Consolidation workbook's TeamKey (CentreCode | Team Name)
composite key exists to handle. If TeamKey wiring is ever wrong, ECO's and
INF's "Ops Team" rows will get merged and their FTE sums will be wrong (see
expected values below).

Usage:
    1. Have Excel running (any workbook, or none) -- wire_fixture() drives
       it via COM (GetActiveObject) to open each fixture itself.
    2. python make_test_centres.py <preparation.xlsx>

Runs the full real pipeline for each fixture: build_centre_template.main()
(stage 1 shell) -> fill_fixture() (openpyxl, writes the Step 1-4 data
directly rather than simulating UI entry) -> wire_reference_data.py
(stage 2, authors the Power-Query reference sheets) -> wire_data_model.py
(Power Pivot Data Model). A fixture isn't realistic without stages 2/3,
since build_centre_template.py alone no longer creates the reference
sheets Step 3's formulas depend on (TacticalScores/InitiativeScores/
AssistanceScores, etc.) -- see templates/scripts/wire_reference_data.py's
docstring.

Expected hand-computed values (Allocation % of Person Effort := SUM per team;
FTE Delivered to Priority := that sum x Allocation % of Team Effort):

    CYB  Firewall Team FTE = 0.50 + 0.50 = 1.00
         -> Harden network perimeter defenses:  1.00 x 0.60 = 0.60
         -> Close critical patching gap:        1.00 x 0.40 = 0.40
    CYB  Threat Intel FTE  = 1.00
         -> Launch predictive analytics pilot:  1.00 x 1.00 = 1.00
    ECO  Ops Team FTE = 0.75
         -> Provide cross-centre threat-briefing support: 0.75 x 1.00 = 0.75
    INF  Ops Team FTE = 0.25   (a DIFFERENT team from ECO's "Ops Team" —
         this is the collision case)
         -> Reduce third-party vendor risk exposure: 0.25 x 1.00 = 0.25

    Total FTE across all three fixture centres = 3.00
    (If ECO/INF's "Ops Team" rows ever get merged by Team Name instead of
    TeamKey, you'd wrongly see a single 1.00-FTE "Ops Team" instead of the
    correct separate 0.75 and 0.25 — that's the regression this fixture
    set is designed to catch.)
"""
import sys
import time
from pathlib import Path

import openpyxl
import win32com.client as win32

sys.path.insert(0, str(Path(__file__).resolve().parents[3] / "templates" / "scripts"))
from build_centre_template import main as build_centre, read_table  # noqa: E402
import wire_reference_data  # noqa: E402
import wire_data_model  # noqa: E402

OUT_DIR = Path(__file__).resolve().parent / "output"

# (centre_code, teams, priorities, resource_allocations)
#   teams:        [(team_name, resources_dedicated, priority_type), ...]
#   priorities:   [(team_name, priority_title, rank, resourced, alloc_pct), ...]
#   allocations:  [(resource_pool_index, team_name, alloc_pct), ...] -- an
#     INDEX into that centre's own ResourceCentres-associated pool, not a
#     literal name: which specific people are associated with which centre
#     is itself randomly synthesized (build_preparation.py), so a
#     hardcoded name here could easily land on someone NOT actually
#     associated with this fixture's centre once Resources became
#     per-centre-filtered (2026-09-03) -- resolved to a real name in
#     main(), after reading each centre's actual pool back out of the
#     freshly-generated preparation.xlsx.
FIXTURES = [
    (
        "CYB",
        [
            ("Firewall Team", "Yes", "Tactical"),
            ("Threat Intel", "Yes", "Initiative"),
        ],
        [
            ("Firewall Team", "Harden network perimeter defenses", 1, "Yes", 0.60),
            ("Firewall Team", "Close critical patching gap", 2, "Yes", 0.40),
            ("Threat Intel", "Launch predictive analytics pilot", 1, "Yes", 1.00),
        ],
        [
            (0, "Firewall Team", 0.50),
            (1, "Firewall Team", 0.50),
            (2, "Threat Intel", 1.00),
        ],
    ),
    (
        "ECO",
        [
            ("Ops Team", "Yes", "Assistance"),
        ],
        [
            ("Ops Team", "Provide cross-centre threat-briefing support", 1, "Yes", 1.00),
        ],
        [
            (0, "Ops Team", 0.75),
        ],
    ),
    (
        "INF",
        [
            ("Ops Team", "Yes", "Tactical"),  # same name as ECO's team, different centre
        ],
        [
            ("Ops Team", "Reduce third-party vendor risk exposure", 1, "Yes", 1.00),
        ],
        [
            (0, "Ops Team", 0.25),
        ],
    ),
]


def fill_fixture(out_path, teams, priorities, allocations, resource_pool):
    wb = openpyxl.load_workbook(out_path)

    ws = wb["Step 1 - Teams"]
    for i, (name, dedicated, ptype) in enumerate(teams):
        row = 2 + i
        ws.cell(row=row, column=1, value=name)
        ws.cell(row=row, column=2, value=dedicated)
        ws.cell(row=row, column=3, value=ptype)
    # Teams is now a real 15-row Table from generation time (see
    # build_centre_template.py's Step 1 comment) — no resize needed here.

    # Step 2 - Select Priorities: derived from `priorities` + each row's
    # team's own Type, deduplicated, in first-seen order -- a real Centre
    # Lead would select these first, before ranking them, so a fixture
    # that skipped Step 2 wouldn't be a realistic exercise of the
    # two-stage flow (see build_centre_template.py's Step 2/3 split).
    team_type = {name: ptype for name, _dedicated, ptype in teams}
    seen = set()
    selections = []
    for team, title, _rank, _resourced, _pct in priorities:
        ptype = team_type[team]
        if title not in seen:
            seen.add(title)
            selections.append((ptype, title))
    ws = wb["Step 2 - Select Priorities"]
    for i, (ptype, title) in enumerate(selections):
        row = 2 + i
        ws.cell(row=row, column=1, value=ptype)
        ws.cell(row=row, column=2, value=title)
        # Columns C-E (Tactical/Initiative/Assistance Helper) are formulas
        # already baked into the template — nothing to fill here.

    ws = wb["Step 3 - Priorities & Ranking"]
    for i, (team, title, rank, resourced, pct) in enumerate(priorities):
        row = 2 + i
        ws.cell(row=row, column=1, value=team)
        ws.cell(row=row, column=2, value=title)
        ws.cell(row=row, column=4, value=rank)
        ws.cell(row=row, column=5, value=resourced)
        cell = ws.cell(row=row, column=6, value=pct)
        cell.number_format = "0%"

    ws = wb["Step 4 - Resource Allocation"]
    for i, (pool_idx, team, pct) in enumerate(allocations):
        row = 2 + i
        ws.cell(row=row, column=1, value=resource_pool[pool_idx])
        ws.cell(row=row, column=3, value=team)
        cell = ws.cell(row=row, column=4, value=pct)
        cell.number_format = "0%"

    wb.save(out_path)


def wire_fixture(out_path):
    """Stages 2/3 via COM: Power-Query reference sheets + Data Model."""
    xl = win32.GetActiveObject("Excel.Application")
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Open(str(out_path))
    wire_reference_data.main(wb.Name)
    xl.CalculateFullRebuild()
    time.sleep(1)
    wire_data_model.main(wb.Name)
    wb.Close(SaveChanges=False)


def main(prep_path):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    prep_wb = openpyxl.load_workbook(prep_path, data_only=True)
    _, centres_rows = read_table(prep_wb, "Centres")
    names_by_code = {row[2]: row[1] for row in centres_rows}
    _, rc_rows = read_table(prep_wb, "ResourceCentres")
    pool_by_code = {}
    for resource, code in rc_rows:
        pool_by_code.setdefault(code, []).append(resource)

    for code, teams, priorities, allocations in FIXTURES:
        out_path = OUT_DIR / f"Centre-{code}.xlsx"
        build_centre(prep_path, str(out_path), centre_name=names_by_code[code], centre_code=code)
        resource_pool = pool_by_code[code]
        fill_fixture(out_path, teams, priorities, allocations, resource_pool)
        wire_fixture(out_path)
        print(f"fixture ready: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python make_test_centres.py <preparation.xlsx>")
    main(sys.argv[1])
