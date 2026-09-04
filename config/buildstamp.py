"""
Build stamp for generated workbooks.

Why this exists: the shipped artefacts are binary .xlsx files that get
carried to a machine with no Python, no git and (often) no network. Git
can't diff them, so nothing about a workbook in someone's hands says which
build it came from. When a Centre Lead reports odd behaviour, "which
version is that?" needs to be answerable from the file itself.

The stamp is written into each workbook at generation time, locked, on the
Instructions sheet, and exposed as a `BuildStamp` defined name so the
Consolidation workbook can read it back out of every returned file (see
wire_power_query.py's Refresh Status query) -- which is what lets
Management spot a centre still working in a stale copy of the template.

Format:  2026-09-04 (7ae039c)        (clean tree)
         2026-09-04 (7ae039c+)       (uncommitted changes at build time)
         2026-09-04 (no-git)         (built outside a git checkout)

Deliberately plain ASCII: this value gets read aloud down a phone line and
retyped into emails, so no en-dashes or middots.

The trailing "+" matters: a stamp naming a commit is a lie if the tree had
local edits when the file was generated, and these workbooks are generated
during development all the time.
"""
import subprocess
from datetime import date
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]


def _git(*args):
    return subprocess.run(
        ["git", *args],
        cwd=REPO_ROOT,
        capture_output=True,
        text=True,
        timeout=10,
    )


def build_stamp(on_date=None):
    """'<YYYY-MM-DD> (<short-sha>[+])', or '(no-git)' outside a checkout."""
    stamp_date = (on_date or date.today()).isoformat()
    try:
        rev = _git("rev-parse", "--short", "HEAD")
        if rev.returncode != 0:
            return f"{stamp_date} (no-git)"
        sha = rev.stdout.strip()
        dirty = _git("status", "--porcelain")
        suffix = "+" if (dirty.returncode == 0 and dirty.stdout.strip()) else ""
        return f"{stamp_date} ({sha}{suffix})"
    except (OSError, subprocess.SubprocessError):
        return f"{stamp_date} (no-git)"


if __name__ == "__main__":
    print(build_stamp())
