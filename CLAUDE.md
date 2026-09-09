# Wilson Hall Directory — working notes

A phone-first room directory for visitors, reached by a QR code at the building
entrance. **It is already built and deployed** — do not recreate it.

**Live:** https://oplabvanderbilt.github.io/wilson-hall-directory/
**Repo:** OPLabVanderbilt/wilson-hall-directory (public, Pages from `main`)

`README.md` explains how everything works. This file is the state of play.

## The one thing to know

**The directory is the source of truth**, not the spreadsheet. Corrections are made
here; `make_updated_sheet.py` then produces a spreadsheet for staff to use
elsewhere. That sheet is an *output*, never edited by hand.

The original `2025_2026 …xlsx` is years out of date and **89 corrections** sit on
top of it in `build.py` — every departure, room move, title and lab rename came
from the Vice Chair or the correction form. **Never "fix" the directory back to
match that spreadsheet.**

### Pending: retire the old spreadsheet as the input

The layered design is backwards now, and `EXCLUDE_PEOPLE` grows forever (32 names
purely to suppress people the stale sheet keeps re-adding). The fix is to feed
`build.py` the *generated* sheet instead, since it is written in the sheet's own
conventions.

**This was tested on 2026-09-09 and works**: building from the generated sheet with
every correction map emptied reproduced 129 of 141 people exactly — nobody wrong,
nobody extra. The 12 missing are precisely the TBD people, who have no room and so
sit on the "Awaiting Room" sheet that `build.py` does not read.

To do it:

1. Make the generated sheet the input (`XLSX` / `SRC` in both scripts).
2. Teach `build.py` to read the "Awaiting Room" sheet into `EXTRA_PEOPLE`.
3. Empty the maps now absorbed: `EXCLUDE_PEOPLE`, `EXCLUDE_ROOMS`,
   `REMOVE_PLACEMENTS`, `EXTRA_ROOMS`, `ROLE_OVERRIDES`, `LAB_OVERRIDES`,
   `SECOND_LAB`, `ROOM_LAB_OVERRIDES`, most of `SPELLING_*`.
4. **Pass condition: still 141 people, same names.** Keep the old sheet in the
   folder as history.

Keep `SECOND_LAB` and the guards — a shared-room label still cannot express two
labs, and the stale-key check still matters.

Those corrections are keyed by person name. When someone's spelling changes, their
key silently stops matching and the correction reverts with no error. `review.txt`
guards this:

    !! CORRECTIONS KEYED ON A NAME THAT MATCHES NOBODY (0):

**If that count is not zero, a correction has stopped working.** Check it after
every build, before pushing.

## Workflow

```bash
python3 build.py          # spreadsheet + corrections -> data.js, review.txt
python3 make_updated_sheet.py   # optional: refresh the .xlsx for the department
```

Then read `review.txt`, commit `build.py data.js index.html`, and push. Pages takes
about a minute; `data.js?v=<hash>` in `index.html` changes each build, so verify the
live stamp matches before saying it is live.

**Add files explicitly — never `git add .`** The spreadsheet, `review.txt`, and
Office documents are gitignored, but new working files may not be.

## Where corrections go in build.py

| Map | For |
|---|---|
| `EXCLUDE_PEOPLE` | departed / retired |
| `EXCLUDE_ROOMS` | space withheld entirely |
| `REMOVE_PLACEMENTS` | room a person no longer occupies |
| `EXTRA_ROOMS` | room the sheet does not record |
| `EXTRA_PEOPLE` | person absent from the sheet (no room ⇒ shows TBD) |
| `ROLE_OVERRIDES`, `LAB_OVERRIDES`, `SECOND_LAB` | role and lab |
| `SPELLING_FIRST`, `SPELLING_LAST` | authoritative spellings |
| `TITLES`, `STAFF_TITLES` | from the department pages |
| `LAB_DISPLAY` / `LAB_PI` | lab naming — see below |

## Two audiences, two names

The **directory** shows the name a lab goes by (OPlab, CATlab, BRAINS Lab). The
**spreadsheet** names the faculty member instead (Gauthier, Palmeri, Kaczkurkin),
because staff reading it may not know the acronyms. `LAB_PI` maps between them.

## Open items

- **12 people have no room** and show TBD — the incoming graduate students. Whoever
  allocates offices could clear all twelve at once.
- **4 correction-form entries could not be applied**: two report a departed post-doc
  and two a graduated student, but the form did not capture *which person*. The form
  has since been fixed; those four need resubmitting.
- **A staff member listed from a 2021 entry** appears as Staff, but the department
  roster lists them as a graduate student. Unresolved.
- **15 students on the roster have no advisor listed**, so it is unclear whether they
  are in this building. Not added.
- **Three staff in the shared 043 rooms** took their lab from the room label, which
  names Hoffman before Womelsdorf. `review.txt` lists them under "LAB GUESSED FROM
  A SHARED ROOM". Unverified.
- **Room 301A** is labelled "Break Room" but was the Grants Specialist's office, now
  vacant and being refilled.
- **Rooms 205, 511/513/514, 611\*, 317/318** are withheld pending reassignment. When
  they are reassigned, delete them from `EXCLUDE_ROOMS`.

## Deliberately withheld from the public page

Animal-facility rooms, storage rooms, empty offices, the entire Notes column, and
anyone on `EXCLUDE_PEOPLE`. See README for the reasoning — the animal-facility
exclusion in particular is a security decision, not tidiness.
