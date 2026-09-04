"""
Loader for config/theme.json -- the single source of truth for every colour
used by all three workbooks (preparation, centre template, consolidation).

Why this exists: colour used to be hardcoded as module constants in three
separate build scripts, with the Risk/Value band palette duplicated in all
of build_preparation.py (twice -- once as RatingLookup's ColourCode data,
once as conditional-formatting constants) and build_centre_template.py.
Changing a corporate colour meant finding every copy. Now they all read
this file.

Two things this does beyond plain JSON loading:

1. **Normalises colour form.** Excel/openpyxl wants "AARRGGBB"; humans have
   "#RRGGBB" from a brand guide. Both (and bare "RRGGBB") are accepted and
   normalised on load, so brand hex codes paste in directly.

2. **Enforces WCAG contrast at build time.** Every fill/text pair is checked
   against accessibility.min_contrast_ratio (default 4.5, WCAG 2.1 AA for
   normal text). A failing pair raises ThemeContrastError and stops the
   build, naming the pair and its actual ratio -- the point being that a
   brand colour dropped in here can't silently produce an unreadable
   workbook for a Centre Lead. Set accessibility.enforce to false to
   downgrade this to a printed warning.

   Band text colours are not configured by default: pick_text_for() chooses
   black or white per band fill, whichever scores higher contrast, so the
   band palette stays readable no matter what fills are chosen. Set "text"
   explicitly on a band to override.

Usage:
    from theme import load_theme
    THEME = load_theme()
    fill, text = THEME.status("over")
    for name, fill, text in THEME.bands("Risk"):
        ...
"""
import json
from pathlib import Path

DEFAULT_PATH = Path(__file__).resolve().parent / "theme.json"

BLACK = "FF000000"
WHITE = "FFFFFFFF"


class ThemeContrastError(Exception):
    """Raised when a configured fill/text pair fails the contrast gate."""


def normalise(colour):
    """'#RRGGBB' / 'RRGGBB' / 'AARRGGBB' -> 'AARRGGBB' (openpyxl's form)."""
    c = str(colour).strip().lstrip("#").upper()
    if len(c) == 6:
        c = "FF" + c
    if len(c) != 8 or any(ch not in "0123456789ABCDEF" for ch in c):
        raise ValueError(f"not a valid colour: {colour!r}")
    return c


def _channel_luminance(v):
    v = v / 255.0
    return v / 12.92 if v <= 0.03928 else ((v + 0.055) / 1.055) ** 2.4


def relative_luminance(colour):
    """WCAG 2.1 relative luminance; alpha is ignored (Excel fills are opaque)."""
    c = normalise(colour)
    r, g, b = int(c[2:4], 16), int(c[4:6], 16), int(c[6:8], 16)
    return (0.2126 * _channel_luminance(r)
            + 0.7152 * _channel_luminance(g)
            + 0.0722 * _channel_luminance(b))


def contrast_ratio(colour_a, colour_b):
    """WCAG 2.1 contrast ratio, 1.0 (identical) .. 21.0 (black on white)."""
    la, lb = relative_luminance(colour_a), relative_luminance(colour_b)
    lighter, darker = max(la, lb), min(la, lb)
    return (lighter + 0.05) / (darker + 0.05)


def pick_text_for(fill):
    """Black or white against `fill`, whichever has the better contrast."""
    return BLACK if contrast_ratio(fill, BLACK) >= contrast_ratio(fill, WHITE) else WHITE


class Theme:
    def __init__(self, raw):
        self.font = raw.get("font", "Arial")
        acc = raw.get("accessibility", {})
        self.min_contrast = float(acc.get("min_contrast_ratio", 4.5))
        self.enforce = bool(acc.get("enforce", True))

        self._cells = {}
        for key, spec in raw["cells"].items():
            if key == "border":
                self.border = normalise(spec)
                continue
            self._cells[key] = (normalise(spec["fill"]), normalise(spec["text"]))

        self._status = {
            key: (normalise(spec["fill"]), normalise(spec["text"]))
            for key, spec in raw["status"].items()
        }
        self._tabs = {key: normalise(val) for key, val in raw["tabs"].items()}

        self._bands = {}
        for rating_type, entries in raw["bands"].items():
            resolved = []
            for e in entries:
                fill = normalise(e["fill"])
                text = normalise(e["text"]) if "text" in e else pick_text_for(fill)
                resolved.append((e["name"], fill, text))
            self._bands[rating_type] = resolved

    # -- accessors -----------------------------------------------------
    def cell(self, key):
        """(fill, text) for 'editable' / 'computed' / '*_header'."""
        return self._cells[key]

    def status(self, key):
        """(fill, text) for 'ok' / 'under' / 'over'."""
        return self._status[key]

    def tab(self, key):
        """Tab colour for 'instructions' / 'entry' / 'reference'."""
        return self._tabs[key]

    def bands(self, rating_type):
        """[(name, fill, text), ...] for 'Likelihood' / 'Risk' / 'Value'."""
        return list(self._bands[rating_type])

    def band_colour_map(self, rating_type):
        """{name: fill} -- for writing RatingLookup's ColourCode column."""
        return {name: fill for name, fill, _text in self._bands[rating_type]}

    # -- validation ----------------------------------------------------
    def contrast_report(self):
        """[(label, fill, text, ratio, passes), ...] for every colour pair."""
        rows = []
        for key, (fill, text) in sorted(self._cells.items()):
            rows.append((f"cells.{key}", fill, text))
        for key, (fill, text) in sorted(self._status.items()):
            rows.append((f"status.{key}", fill, text))
        for rating_type, entries in sorted(self._bands.items()):
            for name, fill, text in entries:
                rows.append((f"bands.{rating_type}.{name}", fill, text))
        out = []
        for label, fill, text in rows:
            ratio = contrast_ratio(fill, text)
            out.append((label, fill, text, ratio, ratio >= self.min_contrast))
        return out

    def validate(self, verbose=False):
        report = self.contrast_report()
        failures = [r for r in report if not r[4]]
        if verbose:
            for label, fill, text, ratio, ok in report:
                print(f"  {'OK ' if ok else 'FAIL'} {label:<34} "
                      f"{fill} on {text}  {ratio:.2f}:1")
        if failures:
            detail = "\n".join(
                f"  {label}: fill {fill} vs text {text} = {ratio:.2f}:1 "
                f"(needs >= {self.min_contrast}:1)"
                for label, fill, text, ratio, _ in failures
            )
            msg = (f"{len(failures)} colour pair(s) in theme.json fail the "
                   f"contrast gate:\n{detail}\n"
                   f"Fix the colours, or lower/disable accessibility.enforce.")
            if self.enforce:
                raise ThemeContrastError(msg)
            print("WARNING: " + msg)
        return failures


def load_theme(path=None, validate=True):
    with open(path or DEFAULT_PATH, encoding="utf-8") as fh:
        raw = json.load(fh)
    theme = Theme(raw)
    if validate:
        theme.validate()
    return theme


if __name__ == "__main__":
    # Report mode: always prints the full table and exits non-zero on
    # failure, rather than raising -- this is the "check my brand colours"
    # entry point, so the report matters more than the traceback.
    import sys

    t = load_theme(validate=False)
    print(f"theme.json - contrast gate >= {t.min_contrast}:1 "
          f"(enforce={t.enforce})\n")
    report = t.contrast_report()
    for label, fill, text, ratio, ok in report:
        print(f"  {'OK  ' if ok else 'FAIL'} {label:<34} "
              f"{fill} on {text}  {ratio:.2f}:1")
    fails = [r for r in report if not r[4]]
    print(f"\n{'FAILED' if fails else 'PASSED'}: {len(fails)} failing pair(s)")
    sys.exit(1 if fails else 0)
