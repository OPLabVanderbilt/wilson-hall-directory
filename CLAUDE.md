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

The original `2025_2026 …xlsx` is years out of date and **102 corrections** sit on
top of it in `build.py` — every departure, room move, title and lab rename came
from the Vice Chair or the correction form. **Never "fix" the directory back to
match that spreadsheet.**

### Pending: retire the old spreadsheet as the input

The layered design is backwards now, and `EXCLUDE_PEOPLE` grows forever (39 names
purely to suppress people the stale sheet keeps re-adding). The fix is to feed
`build.py` the *generated* sheet instead, since it is written in the sheet's own
conventions.

**This was tested on 2026-09-09 and works**: building from the generated sheet with
every correction map emptied reproduced 129 of the 141 people then published — nobody
wrong, nobody extra. The 12 missing were precisely the TBD people, who have no room and
so sit on the "Awaiting Room" sheet that `build.py` does not read.

That test predates the corrections applied later the same day, so its raw numbers are
stale; the method is what it establishes. **The roster is now 136 people with 10 on
"Awaiting Room".**

To do it:

1. Make the generated sheet the input (`XLSX` / `SRC` in both scripts).
2. Teach `build.py` to read the "Awaiting Room" sheet into `EXTRA_PEOPLE`.
3. Empty the maps now absorbed: `EXCLUDE_PEOPLE`, `EXCLUDE_ROOMS`,
   `REMOVE_PLACEMENTS`, `EXTRA_ROOMS`, `ROLE_OVERRIDES`, `LAB_OVERRIDES`,
   `SECOND_LAB`, `ROOM_LAB_OVERRIDES`, most of `SPELLING_*`.
4. **Pass condition: the same people out as went in — 136 at the time of writing,
   so re-read the current count from `review.txt` before trusting that number.**
   Keep the old sheet in the folder as history.

Keep `SECOND_LAB` and the guards — a shared-room label still cannot express two
labs, and the stale-key check still matters.

Those corrections are keyed by person name. When someone's spelling changes, their
key silently stops matching and the correction reverts with no error. `review.txt`
guards this:

    !! CORRECTIONS KEYED ON A NAME THAT MATCHES NOBODY (0):

**If that count is not zero, a correction has stopped working.** Check it after
every build, before pushing.

A second guard sits beside it:

    !! PUBLISHED ROOMS LABELLED LIKE ANIMAL-FACILITY SPACE (0):

`ANIMAL_ROOM_KINDS` matches a room label exactly, so a renamed or newly added label
("Necropsy", "Housing Room 2") would reach the public page with nothing to notice it.
The `ANIMAL_SMELL` pattern catches anything that still reads like animal-facility
space and reports it. **It only warns.** To actually withhold the room, add its label
to `ANIMAL_ROOM_KINDS`.

## Workflow

**Ask the Vice Chair what today's date is before starting a revision, and set `TODAY`
in `build.py` to it.** Standing request, made 2026-09-10. Nothing in the build can
detect a stale date — it is pinned deliberately, so that a rebuild of an old checkout
reproduces that day's output — and a wrong one silently mis-stamps `data.js`,
`review.txt`, and the filename of the spreadsheet the department works from. It also
decides who is dropped for a past last day. Ask; do not assume the session's own clock
is what the department is working to.

`make_updated_sheet.py` takes its date and output filename from `build.TODAY`, so
there is only the one place to change. Each run writes a new dated file; **move the
previous one into `Old sheets/`** so the folder holds exactly one current sheet and
nobody has to guess which is live. The whole folder is gitignored, like every `.xlsx`.

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

Since 2026-09-09 the directory shows **both**: every lab row carries the faculty
member's full name under the lab name. `build.py` resolves it into a `fac` field on
each lab, matching `LAB_PI`'s surname against the published faculty. Doing it in the
build rather than in the page matters — three labs carry an acronym rather than a
surname, and Chen and Kaas each have more than one faculty member in the lab, so
surname-matching in JS would pick the wrong person.

**There are no shared labs.** Where the sheet labels a room with two PIs it means
shared space, and the directory says so: `Hoffman/Womelsdorf shared space`, not a
joint lab. `lab_name()` does this for any lab key containing `/`. Those entries have
rooms but no members, so their row shows no people count.

A lab row shows a **room count, not room numbers, and no floor.** Both were actively
misleading: the floor was that of `rooms[0]` alone, wrong for the 7 of 28 labs whose
rooms span more than one floor. Dropping the chips is also what made room for the
faculty name — they cost about 190px of a 375px row.

## Open items

- **10 people have no room** and show TBD — the incoming graduate students. Whoever
  allocates offices could clear all ten at once.
- **A staff member listed from a 2021 entry** appears as Staff, but the department
  roster lists them as a graduate student. Unresolved.
- **15 students on the roster have no advisor listed**, so it is unclear whether they
  are in this building. Not added.
- **Three staff in the shared 043 rooms** took their lab from the room label, which
  names Hoffman before Womelsdorf. `review.txt` lists them under "LAB GUESSED FROM
  A SHARED ROOM". Unverified.
- **Room 301A** is labelled "Break Room" but was the Grants Specialist's office, now
  vacant and being refilled.
- **Kris Clifft's role is provisional.** She took 013 after Chrissy Suell left and is
  recorded as Staff pending confirmation. The comment in `EXTRA_PEOPLE` says so.
- **Ziqi Wang keeps 213A as well as 402.** Only 402 was reported; the Park lab room was
  left in place because an office plus a lab room is the normal pattern here. Unverified.
- **The animal-facility exclusion hides function, not footprint.** Withheld rooms are
  inferable from the gaps: the basement lists 002–069 near-continuously, so the absent
  runs (031A–F, 036A–E, 039A–F, 044A–F, 041/042) mark where the housing and surgery
  rooms are. Raised 2026-09-09; the decision was that labels are the criterion and this
  is acceptable. Revisit only if the exposure concern changes.
- **Rooms 205, 511/513/514, 611\*, 317/318** are withheld pending reassignment. When
  they are reassigned, delete them from `EXCLUDE_ROOMS`.

## Deliberately withheld from the public page

Animal-facility rooms, storage rooms, empty offices, the entire Notes column, and
anyone on `EXCLUDE_PEOPLE`. See README for the reasoning — the animal-facility
exclusion in particular is a security decision, not tidiness. 000CB, the corridor
inside the animal facility, is withheld through `ANIMAL_ROOM_KINDS` rather than
`EXCLUDE_ROOMS`: it is not awaiting reassignment, and matching on the label catches
any future room named the same way.

**Empty offices go; empty lab rooms stay.** Dropping every empty room was tried on
2026-09-09 and reversed the same day — an empty lab room still answers "whose lab is
023?". The comment on `HIDE_WHEN_EMPTY` records this. Do not widen it to lab rooms.
