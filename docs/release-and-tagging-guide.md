# Releases and Tagging — a Guide for the Build Person

## Who this is for

Whoever generates and distributes the workbooks. It assumes you've used
`git commit` and `git push` but never cut a tag or written a release before.

## The one-paragraph version

Every quarter, when the six centre files actually go out to Centre Leads,
you mark that moment in GitHub with a **tag** and write a short **release
note** saying what changed. That gives you a permanent, named point you can
return to — "this is exactly what Q3's files were built from" — and a
plain-English record of what changed for the people filling them in.
You do this **only at distribution time**, not on every commit.

---

## 1. Three words, in plain terms

| Word | What it is | Analogy |
|---|---|---|
| **Commit** | One saved change to the project. You make lots of these. | A page in a notebook |
| **Tag** | A permanent name pinned to one specific commit. | A sticky bookmark on one page |
| **Release** | A tag plus human-readable notes (and optionally attached files), shown on GitHub's Releases page. | The bookmark, plus a note explaining why that page matters |

A tag doesn't change anything in the project. It just makes one commit
findable by name forever, instead of by a code like `4b09b7f`.

## 2. Why this matters *here* specifically

More than on a typical software project, for two reasons:

**The deliverables are binary.** `.xlsx` files can't be meaningfully
compared by git. If you ask "what changed in the workbook between last
quarter and this one?", git genuinely cannot tell you — it only knows the
bytes differ. **Your release notes are the only changelog that will ever
exist** for these files. That's not a formality; it's the actual record.

**The files travel somewhere git can't follow.** They're carried to a
machine with no Python, no network and no repo access. So when a Centre
Lead reports a problem, nothing about their copy is traceable — unless you
can match it to a known build.

That's what the **build stamp** is for (see
[`excel-file-design.md`](excel-file-design.md#build-stamp-and-release-tagging)).
Every workbook carries one on its Instructions tab:

```
Workbook build:   2026-09-04 (4b09b7f)
```

The tag is how *the repository* remembers. The stamp is how *the file*
remembers. You need both — only the second one travels to work.

## 3. When to tag

**Tag when files actually go out to Centre Leads.** That's it.

- ✅ You've generated the six centre files and you're about to distribute them
- ❌ Not after every commit
- ❌ Not for work-in-progress, even significant work

Tagging every commit would leave you with dozens of meaningless bookmarks.
A tag should mean *"real people received these exact files."*

**Naming scheme:** `dist-YYYY-QN` — e.g. `dist-2026-Q3`. Sorts correctly,
matches the collection cycle, and is obvious a year later.

## 4. Doing it — step by step

### Step 1: Make sure everything is committed

```bash
git status
```

You want to see `nothing to commit, working tree clean`. If not, commit
your work first. **This step is not optional** — see the warning below.

### Step 2: Regenerate the files from the clean tree

```bash
python management/scripts/build_preparation.py management/preparation.xlsx
python scripts/recalc_and_save.py management/preparation.xlsx
python templates/scripts/build_centre_template.py management/preparation.xlsx --all management/centres
```

Then run the stage-2/3 wiring as described in
[`excel-file-design.md`](excel-file-design.md#regenerating-a-workbook).

> **⚠️ Why the clean tree matters.** The build stamp adds a `+` when it's
> built from a tree with uncommitted changes:
>
> | Stamp | Meaning |
> |---|---|
> | `2026-09-04 (4b09b7f)` | Built from exactly commit `4b09b7f`. Traceable. |
> | `2026-09-04 (4b09b7f+)` | Built from `4b09b7f` **plus unsaved edits**. Not traceable. |
>
> A `+` in a file that went out to a Centre Lead means you can never
> reconstruct what they actually have. Files for distribution should carry
> **no `+`**.

### Step 3: Commit the regenerated files

```bash
git add management/centres/ templates/ management/preparation.xlsx
git commit -m "Generate Q3 2026 distribution files"
git push
```

### Step 4: Create the tag

```bash
git tag -a dist-2026-Q3 -m "Q3 2026 distribution to Centre Leads"
git push origin dist-2026-Q3
```

`-a` makes an *annotated* tag — it records who made it and when. Always use
`-a`; a bare `git tag name` creates a lightweight tag with no author or
date, which is worth less later.

Note that `git push` alone does **not** push tags. You must push the tag
explicitly, as above.

### Step 5: Write the release on GitHub

1. Go to the repository → **Releases** (right-hand side) → **Draft a new release**
2. **Choose a tag** → pick `dist-2026-Q3` (already exists from Step 4)
3. **Release title:** `Q3 2026 distribution`
4. Write the notes (template below)
5. Optionally **attach the six centre files** — see §6
6. **Publish release**

## 5. What to write in the notes

Write for **Management and Centre Leads**, not for developers. The test:
*would a Centre Lead reading this know whether anything they do changes?*

Skip commit messages and file names. Describe what's different to use.

### Template

```markdown
## What's new for Centre Leads

- (Anything they'll see or do differently. If nothing, say "No changes to
  how you fill in the workbook.")

## What's new for Management

- (Changes to preparation.xlsx, the Consolidation workbook, or the process.)

## Reference data in this build

- 50 priorities, 100 resources, 6 centres.
- (Note anything notable that changed since last quarter.)

## Build stamp

Files in this distribution show: `2026-09-04 (4b09b7f)`
Ask a Centre Lead to quote this if they report a problem.
```

### Worked example

```markdown
## What's new for Centre Leads

- New **Step 2 - Select Priorities**: pick your shortlist first, then rank
  from it in Step 3. The ranking dropdown now only offers priorities you
  selected.
- Step 3 shows a running **Team Total %** in the row you're typing in, so
  you no longer have to switch sheets to check you've hit 100%.
- New read-only **Ranked View** tab showing your priorities sorted by team
  and rank.
- The **Resources** list now only shows people associated with your centre.
- Dropdowns are alphabetical with no blank entries.

## What's new for Management

- New **Refresh Status** tab in the Consolidation workbook: confirms which
  centre files a refresh picked up, and flags any centre working in an
  older template.
- New **ResourceCentres** sheet in preparation.xlsx — controls which people
  each centre can pick from.

## Build stamp

Files in this distribution show: `2026-09-04 (4b09b7f)`
```

## 6. Optionally: attach the files to the release

GitHub lets you attach files to a release (drag them onto the draft). Worth
doing here, because it gives you one dated, permanent bundle of exactly
what was sent — handy if someone loses a file mid-quarter, and it doesn't
depend on anyone being able to run the build scripts.

Attach the six `Centre-*.xlsx` files. Don't attach `preparation.xlsx` if it
contains anything you wouldn't want visible to everyone with repo
access — check the repository's visibility first.

## 7. Fixing mistakes

**Tagged the wrong commit, and haven't pushed yet:**
```bash
git tag -d dist-2026-Q3
```
Then re-tag correctly.

**Already pushed the tag.** Prefer *not* to move it — someone may already
be relying on it. Cut a new one instead:
```bash
git tag -a dist-2026-Q3b -m "Q3 2026 distribution (corrected)"
git push origin dist-2026-Q3b
```
and explain why in the release notes. Moving a published tag is the one
genuinely confusing thing you can do here — it makes the same name mean two
different things for different people.

**Typo in published release notes:** just edit them on GitHub. Notes are
free to change; tags are not.

## 8. Useful commands

```bash
git tag                      # list all tags
git show dist-2026-Q3        # what commit does this tag point at?
git checkout dist-2026-Q3    # look at the project exactly as it was
git checkout main            # ...and come back
```

`git checkout <tag>` is the payoff for all of this: it puts the whole
project back exactly as it was when those files were built, so you can
regenerate a byte-identical workbook a year later.

## Related docs

- [`excel-file-design.md`](excel-file-design.md) — build stamp mechanics
  and the full regeneration order.
- [`management-user-guide.md`](management-user-guide.md) — what Management
  does with the distributed files.
