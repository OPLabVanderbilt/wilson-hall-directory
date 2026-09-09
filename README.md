# Wilson Hall Directory

A phone-first web directory for visitors entering Wilson Hall. Scan a QR code at the
entrance, type a name, get a room number.

**Live at <https://oplabvanderbilt.github.io/wilson-hall-directory/>.**
See `CLAUDE.md` for current state and open items.

## Files

| File | Role |
|---|---|
| `index.html` | The whole site. No frameworks, no build step, no network calls. |
| `data.js` | Generated. People, rooms and labs, as one JSON blob. |
| `build.py` | Regenerates `data.js` from the spreadsheet. |
| `review.txt` | Generated. Data-quality report **for the department, not for the web**. |
| `2025_2026 Wilson Hall Space Assignments.xlsx` | Source of truth. |

## Updating after the spreadsheet changes

```bash
python3 build.py
```

Then read `review.txt` before publishing — it lists every person dropped for a past
last-day, every fuzzy name merge the script made, and every conflicting lab
affiliation it could not resolve.

## What is deliberately withheld

- **The entire Notes column.** It holds termination dates, start dates, and internal
  space-negotiation history. None of it is read.
- **Animal-facility rooms.** Housing rooms, the surgery suite, food storage, the
  laundry/PPE room, freezer rooms, the veterinary office, cage wash, and lab-service
  rooms are excluded. Publishing exact NHP housing locations on a public URL is an
  avoidable exposure. To change this, set `HIDE_ANIMAL_FACILITY = False` in `build.py`.
- **Storage rooms.** Any room whose spreadsheet label is `(Storage)` — 107, 108, 109,
  518, and 314. Controlled by `HIDE_ROOM_KINDS` in `build.py`, matched on the label
  before friendly relabelling, so a new storage room is caught automatically. Note 314
  is the department supply room; its room label says Storage, so the rule catches it.
- **Empty offices, vacant rooms, and unlabelled rooms.** A room with nobody in it
  is dropped if its label matches `HIDE_WHEN_EMPTY` in `build.py` (`office` or
  `vacant`) or has no label at all. Empty *lab* rooms are kept, because "whose lab
  is 023?" is still a real question; the page groups those instead. Facilities —
  conference rooms, lounges, the kitchen — are kept and simply show no occupant.
- **Anyone whose recorded last day has passed.**
- **Anyone on the `EXCLUDE_PEOPLE` list in `build.py`** — people the spreadsheet still
  lists but who have retired or left with no last-day date to catch them. Add a
  lowercase `"first last"` line per person; `review.txt` reports who was dropped.
- **Internal annotations inside the Assignee cell** — "per Tom", "used to be",
  "temporarily", and any embedded dates.

Basement lab rooms that are *not* animal-facility rooms are still listed, because a
visitor looking for a person assigned there needs the number.

## Home-screen groupings

The home screen runs **People and Labs**, then **Common spaces**, then
**Browse by floor**.

**Common spaces** shows only what someone at the front door most likely wants —
Main Office, Chair, Vice Chair, Restrooms, Lactation Room — and folds the rest into
an expandable "Other common spaces". The `PRIMARY` regex in `viewHome()` decides
which rooms stay visible; everything else in the section collapses. Search is
unaffected: a collapsed room still appears in results.

`INFO_CARDS` in `build.py` holds standing guidance that is not a room — currently
the restrooms. Each card has a `title`, a one-line `sub`, a `body` (one string per
paragraph) and `keywords` that make it searchable; "bathroom", "toilet", "loo" and
"wc" all find the restrooms card. Add a card there, not in `index.html`. Its membership is a regex over the room's own label in `viewHome()`
in `index.html` — currently Main Office, Chair, Mail Room, Lactation, Conference Room,
Lounge, Copier.

**People and Labs** offers All / Faculty / Graduate students / Post-docs / Staff plus
All labs. Two structures control it, both in `index.html`: the `GROUPS` array in
`viewHome()` sets the buttons, and `ROLE_SETS` says which recorded roles each button
covers. A button may cover several roles — **Faculty** covers `Faculty` and `Lecturer`,
while each person still displays their own precise role.

18 undergraduate/RA entries (role `Student`) and 10 with no role recorded in the
spreadsheet are reachable only through **All**.

## Faculty titles

`TITLES` in `build.py` maps a person to the title shown in place of their generic
role. Titles were taken from <https://as.vanderbilt.edu/psychology/faculty/> on
2026-09-09 and reduced by one rule:

> Keep the academic rank and any administrative role. Drop endowed or named chairs,
> and drop appointments in other departments.

So "David K. Wilson Chair of Psychology; Vice Chair Department of Psychology;
Professor of Radiology and Radiological Sciences" becomes **Professor · Vice Chair**.

`STAFF_TITLES` does the same for departmental staff, from
<https://as.vanderbilt.edu/psychology/faculty/?group=staff>. Only the five
departmental staff appear on that page; lab personnel are not listed there and keep
the generic "Staff".

`EXTRA_PEOPLE` adds someone who belongs in the directory but has no room in the
spreadsheet yet; they render with a **TBD** chip instead of a room number. Delete the
entry once the spreadsheet carries them — the build skips it if the name already
appears, so a duplicate cannot slip through. `review.txt` lists everyone showing TBD.

`EXTRA_ROOMS` adds a room placement the spreadsheet does not record — currently the
Senior Administrative Officer, who works out of the main office (301) as well as her
own 301D.

Neither map is regenerated automatically — the department page has to be re-checked
by hand when titles change. `review.txt` reports any `TITLES` key that no longer
matches a person, and any faculty member with no title.

## Grouped lab rooms

Empty rooms belonging to one lab on one floor collapse into a single row — the
basement's ten empty Constantinidis rooms become one "Constantinidis Lab · 10 rooms"
entry linking to the lab. This happens in `groupRooms()` in `index.html`, at display
time only: every room stays individually addressable at `#/r/<num>`, and a room whose
number the reader actually typed is never folded away.

Room labels for lab rooms are rewritten to the canonical lab name during the build,
so the sheet's "Herculano Lab" and "Constantindis Lab" both display correctly.

## Cache

GitHub Pages serves everything with `max-age=600`, and caches `index.html` and
`data.js` independently — so without care a reader can get a new page with a
ten-minute-old dataset. `build.py` therefore stamps a content hash onto the script
URL (`data.js?v=<hash>`), which changes whenever the data does. **Do not remove that
stamp or hand-edit it**; the build rewrites it and warns if it cannot.

The page itself can still be up to ten minutes stale after a push. That is a Pages
setting and cannot be changed from the repo.

## Corrections

Every screen carries a **Request a change or correction** button pointing at
`https://redcap.link/ul4x2a5q`. Change `FORM_URL` near the top of the script block in
`index.html` if that link ever moves.

Corrections arrive in REDCap; they do **not** flow back automatically. The workflow is:
edit the spreadsheet, re-run `build.py`, commit.

## Publishing to GitHub Pages

```bash
git init && git add . && git commit -m "Wilson Hall directory"
```

Do not commit the spreadsheet or `review.txt` if the repository is public — a
`.gitignore` is included that excludes both. Push to a repo, then in
**Settings → Pages** set Source to `Deploy from a branch`, branch `main`, folder `/ (root)`.

The site will be at `https://<org-or-user>.github.io/<repo>/`.

## Making the QR code

Once the URL exists, any of these produce a print-ready code:

```bash
brew install qrencode && qrencode -o wilson-hall-qr.png -s 12 -m 2 "https://YOUR-URL-HERE"
```

Print it at least 3 cm square for reliable scanning at arm's length, and put the URL
in readable text underneath — some visitors will not scan.

## Local preview

```bash
python3 -m http.server 8731
```

Then open `http://localhost:8731/`.
