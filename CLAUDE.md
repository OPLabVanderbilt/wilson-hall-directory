# Wilson Hall Directory — working notes

A phone-first room directory for visitors, reached by a QR code at the building
entrance. **It is already built and deployed** — do not recreate it.

**Live:** https://oplabvanderbilt.github.io/wilson-hall-directory/
**Repo:** OPLabVanderbilt/wilson-hall-directory (public, Pages from `main`)

`README.md` explains how everything works. This file is the state of play.

## The one thing to know

The spreadsheet is **not** the source of truth for who is here. It is years out of
date. Roughly 40 corrections live on top of it in `build.py` — every departure,
room move, title and lab rename came from the Vice Chair or the correction form,
never from the sheet. **Never "fix" the directory back to match the spreadsheet.**

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
- **Room 301A** is labelled "Break Room" but was the Grants Specialist's office, now
  vacant and being refilled.
- **Rooms 205, 511/513/514, 611\*, 317/318** are withheld pending reassignment. When
  they are reassigned, delete them from `EXCLUDE_ROOMS`.

## Deliberately withheld from the public page

Animal-facility rooms, storage rooms, empty offices, the entire Notes column, and
anyone on `EXCLUDE_PEOPLE`. See README for the reasoning — the animal-facility
exclusion in particular is a security decision, not tidiness.
