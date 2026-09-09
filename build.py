#!/usr/bin/env python3
"""
Build the Wilson Hall public directory from the space-assignment spreadsheet.

    python3 build.py

Reads  : 2025_2026 Wilson Hall Space Assignments.xlsx
Writes : data.js          (consumed by index.html)
         review.txt       (data-quality report for the Vice Chair -- NOT published)

Everything in the Notes column is discarded. Only Room Number + Assignee are read,
and both are scrubbed before they reach the public file.
"""

import json, re, unicodedata, datetime, sys, hashlib
from pathlib import Path
from collections import defaultdict

import openpyxl

HERE = Path(__file__).parent
XLSX = HERE / "2025_2026 Wilson Hall Space Assignments.xlsx"
TODAY = datetime.date(2026, 9, 9)

# --- Policy switches ---------------------------------------------------------

# Animal-facility rooms are withheld from the public page. Publishing exact
# locations of non-human-primate housing is a security exposure. Set to False
# only after a deliberate decision to publish them.
HIDE_ANIMAL_FACILITY = True

# Room types withheld from the public page regardless of floor. Matched against
# the label in the spreadsheet's Room Number cell, before any friendly relabelling,
# so a new "(Storage)" room is caught automatically.
HIDE_ROOM_KINDS = {
    "storage",
}

# An office with nobody in it tells a visitor nothing, so empty rooms of these
# kinds are dropped after occupancy is worked out. Lab rooms are NOT included:
# an empty lab room still answers "whose lab is 023?", and those are grouped by
# floor in the page instead.
#
# Revisited 2026-09-09: dropping every empty room was tried and reversed the same
# day. Empty labs stay, empty offices go. Do not widen this to lab rooms.
HIDE_WHEN_EMPTY = re.compile(r"office|vacant", re.I)

# ANIMAL_ROOM_KINDS below is matched exactly, so a label that is renamed or newly
# added in the spreadsheet -- "Necropsy", "Housing Room 2", "Quarantine" -- would be
# published with nothing to notice it. This pattern is the safety net: any label that
# survives to the public page and still reads like animal-facility space is reported
# in review.txt, the same way a stale correction key is. It only warns; widen
# ANIMAL_ROOM_KINDS to actually withhold the room.
ANIMAL_SMELL = re.compile(
    r"animal|whaf|\bdac\b|vivar|primate|\bnhp\b|monkey|housing|surger|"
    r"necrops|quarantin|autoclav|laundry|freezer|veterinar|food storage|"
    r"lab service|perfusion|husbandry|cage wash", re.I)

ANIMAL_ROOM_KINDS = {
    "housing room", "surgery suite", "autoclaves", "food storage",
    "laundry room", "freezer", "veterinary office", "dac breakroom",
    "lab service", "kaas personnel/microscope",
    # Circulation space inside the animal facility; its own cell reads "Corridor
    # within WHAF". Withheld 2026-09-09 for the same reason as the rooms it serves.
    "whaf corridor",
}

# --- Text scrubbing ----------------------------------------------------------

# Date annotations embedded in the Assignee cell.
RE_LASTDAY = re.compile(r"(?:last\s*day|term(?:\s*eff\.?)?)\s*:?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.I)
RE_DATEISH = re.compile(
    r"[,;]?\s*(?:start(?:\s*date)?|eff\.?|effective|adjoint\s*eff\.?|last\s*day|term(?:\s*eff\.?)?)?"
    r"\s*\d{1,2}/\d{1,2}/\d{2,4}", re.I)

# Internal chatter that must never reach a public page.
RE_INTERNAL = re.compile(
    r"\((?:per\s+tom|tom\s+said)[^)]*\)|per\s+tom[^,;)]*|tom\s+said[^,;)]*"
    r"|use\s*d?\s*to\s*be[^,;)]*|formerly[^,;)]*|currently[^,;)]*", re.I)

# Rows that describe a fixture, a pointer, or an empty seat -- not a person.
RE_POINTER = re.compile(r"^\s*see\s+(?:room\s+)?\d", re.I)
NON_PERSON_EXACT = {
    "", "?", "vacant", "storage", "lounge", "kitchen", "conference room",
    "copier room", "shop", "supply room", "storage room", "emeritus office",
    "lactation room", "mail room & dept. lounge", "dac breakroom",
    "research lab service", "non-human primate housing room", "veterinary office",
    "cage washing/autoclaves", "laundry room, ppe & chemical storage",
    "food storage for all animals in whaf", "teba server room",
    "research lab service (freezer room)", "whaf surgery suite used by all investigators",
    "corridor within whaf", "desk 1", "desk 2", "desk 3",
    "preference to schaefer lab members", "formerly wallace lab",
    "sohee park's office", "schaefer lab interview room",
}
RE_NON_PERSON = re.compile(
    r"^\s*(?:main\s*ent(?:e)?ry\s*door|vacant\s*desk|only\s+one\s+desk|desk\s*#?\s*\d"
    r"|microscope\s+room|non-human|ram's\s+ra|.*computer\s+space)", re.I)

# People still listed in the spreadsheet who should not appear in the public
# directory -- retired, departed, or otherwise no longer in the building, where
# the spreadsheet carries no last-day date to catch them automatically.
# Key on "first last", lowercase, no accents.
EXCLUDE_PEOPLE = {
    "jo anne bachorowski",   # retired
    "joe lappin",            # emeritus
    "rankin mcgugin",        # retired
    "erin duran",            # left the department
    "randolph blake",        # retired
    "sydnie rathert",        # left 2026-07-02 for DAR; Grants Specialist post being refilled
    "david schlundt",        # retired
    # Reported by Adam Tiesman 2026-09-09: nobody but Ram works in the
    # Ramachandran lab any more.
    "adriana schoenhaut",
    "alejandro tarabillo",
    "alexander mcleod",
    "amy stahl",
    "catherine alek",
    "chase mackey",
    "jackson mayfield",
    "jane burton",
    "karina jirik",
    # Same report: still in the Wallace lab, but not housed in Wilson Hall.
    "henry ong",
    "marcus watson",
    "mckenzie king",

    "saman abbaspoor",       # left, reported 2026-09-09
    # Departures reported 2026-09-09.
    "simon lilburn",
    "jason chow",
    "jordan gunn",
    "seth marx",
    # Schlundt lab members, following his retirement. All were listed only in 317.
    "ashley sellers",
    "he xia",
    "isabela arcila",
    "jiya patel",
    "kemberlee bonnet",
    "madeline stapp",
    "maya levinson",
    "seohyun choi",

    "chrissy suell",         # no longer at Vanderbilt, reported 2026-09-09

    # Reported 2026-09-09 as out of the Constantinidis lab; confirmed gone.
    "will banks",
    "russell jaffe",         # the sheet spells him both ways -- 013 has
    "rye jaffe",             # "Jaffe, Russell (Rye)", 413 has "Jaffe, Rye".
    # Reported 2026-09-09: moved on from the Park lab.
    "olivia jelsma",
    "hyeonseung lee",
    # Same person as Ziqi Wang below, who keeps 213A.
    "joanna wang",
}

# Standing information that is not a room. These appear in Common spaces and are
# searchable by their keywords. Edit the text here, not in index.html.
INFO_CARDS = [
    {
        "id": "restrooms",
        "title": "Restrooms",
        "sub": "Basement, 1st, 2nd, 4th and 6th floors",
        "body": [
            "Restrooms are on the basement, 2nd, 4th and 6th floors \u2014 "
            "at end of the hall to the left when facing the street.",
            "On the 1st floor, they are near the elevator.",
        ],
        "keywords": "restroom restrooms bathroom bathrooms toilet toilets "
                    "washroom wc lavatory loo",
    },
]

# Rooms the sheet labels wrongly or not at all. Setting the lab also fixes the
# room's displayed name, which is derived from it.
ROOM_LAB_OVERRIDES = {
    "614A": "Wallace",   # sheet says "Vacant Lab / Formerly Wallace Lab"
}

# Rooms to withhold: space vacated by a departure and not yet reassigned, where
# the spreadsheet still carries the former occupant. Anyone left with no room at
# all after this is dropped too, and reported in review.txt.
EXCLUDE_ROOMS = {
    # Randolph Blake, retired: offices, lab, and the 6th-floor Blake Lab suite.
    "511", "513", "514",
    "611", "611A", "611AA", "611AB", "611AC", "611AD",
    "611B", "611BA", "611BB", "611BC",
    # Schlundt lab, closed on his retirement.
    "317", "318",
    # Kaczkurkin reports 2026-09-09 that her lab is 210, not 205. What 205 is
    # now is unknown, so it is withheld rather than shown as her lab.
    "205",
}

# Titles from https://as.vanderbilt.edu/psychology/faculty/ (retrieved 2026-09-09).
# Rule, per the Vice Chair: keep the academic rank and any administrative role;
# drop endowed/named chairs and secondary departmental appointments. So
# "David K. Wilson Chair of Psychology; Vice Chair Department of Psychology;
# Professor of Radiology and Radiological Sciences" -> "Professor - Vice Chair".
# Keys are lowercase "first last" as the name appears in the spreadsheet.
TITLES = {
    "ryan balch":              "Senior Lecturer",
    "andre bastos":            "Assistant Professor",
    "devin burns":             "Senior Lecturer",
    "isabel gauthier":         "Professor \u00b7 Vice Chair",
    "kirsten haman":           "Assistant Professor of the Practice \u00b7 "
                               "Assistant Co-Director, Clinical Training",
    "suzana herculano-houzel": "Associate Professor",
    "kari hoffman":            "Associate Professor",
    "steve hollon":            "Professor",
    "jon kaas":                "Professor",
    "antonia kaczkurkin":      "Assistant Professor",
    "gordon logan":            "Professor",
    "alex maier":              "Associate Professor",
    "rene marois":             "Professor \u00b7 Director of Graduate Studies",
    "ashleigh maxcey":         "Senior Lecturer \u00b7 Research Assistant Professor",
    "tim mcnamara":            "Professor",
    "bunmi olatunji":          "Professor \u00b7 Director of Clinical Training",
    "tom palmeri":             "Professor \u00b7 Department Chair",
    "sohee park":              "Professor",
    "sean polyn":              "Associate Professor",
    "elisabeth sandberg":      "Senior Lecturer",
    "jon schaefer":            "Assistant Professor",
    "david schlundt":          "Associate Professor",
    "adriane seiffert":        "Principal Senior Lecturer \u00b7 "
                               "Director of Undergraduate Studies",
    "frank tong":              "Professor",
    "mark wallace":            "Professor",
    "ashley watts":            "Assistant Professor",
    "thilo womelsdorf":        "Professor",
    "geoff woodman":           "Professor",
}

# Departmental staff titles from
# https://as.vanderbilt.edu/psychology/faculty/?group=staff (retrieved 2026-09-09).
# Only the five departmental staff appear there; lab personnel are not listed and
# keep the generic "Staff".
STAFF_TITLES = {
    "bianca castellon":     "Administrative Manager \u00b7 Financial Unit Manager",
    "faith clark":          "Program Specialist",
    "savannah crutchfield": "Senior Administrative Officer",
    "ashley lowther":       "Administrative Specialist",
    "aubrey smith":         "Grants Manager",
}

# Corrections reported through the REDCap form, applied on top of the spreadsheet.
# Each entry records who asked and when, so a later spreadsheet update can be
# checked against them.
# Advisors the department roster records but the space sheet does not.
LAB_OVERRIDES = {
    "yinuo peng":    "Palmeri",   # department roster, 2026-09-09
    "ginni strehle": "Tong",
    # Both the space sheet ("Ikhwan Jeon (Tong Grad Student)") and the department
    # roster still say Tong. That is out of date -- confirmed 2026-09-09. Do not
    # "correct" this back from either source.
    "ikhwan jeon":   "OPlab",
    # Confirmed 2026-09-09. The sheet put him in 221C, a Woodman room, which is
    # what made him look like a Marois/Woodman conflict; he has left that room.
    "zengbo xie":    "Marois",
}

SECOND_LAB = {
    "adam tiesman": "Ramachandran",   # self-reported 2026-09-09
    "ikhwan jeon":  "CATlab",         # confirmed 2026-09-09; see LAB_OVERRIDES
}

ROLE_OVERRIDES = {
    "conor smithson": "Post-doc",      # self-reported 2026-09-09
}

# Placements in the spreadsheet that the occupant says are no longer true.
REMOVE_PLACEMENTS = {
    "conor smithson":  ["309"],        # self-reported 2026-09-09, now in the Gauthier lab
    # Reported by Adrian Wong 2026-09-09 (as "Salad", the only occupant of 429).
    # He keeps 043A.
    "sajad ahmadi nebi": ["429"],
    # Self-reported 2026-09-09. He keeps 405; 627A replaces 221C below.
    "zengbo xie":  ["221C"],
    # Moved to 627B, reported 2026-09-09. 517 keeps Jinhyeok Jeong.
    "lanting qiu": ["517"],
}

# Extra room placements not in the spreadsheet: people who work out of a room the
# sheet assigns to someone else, and whom a visitor would look for there.
EXTRA_ROOMS = {
    "savannah crutchfield": ["301"],   # Senior Administrative Officer, in the main office
    "sohee park":          ["525"],    # self-reported 2026-09-09: office is 525
    "antonia kaczkurkin":  ["210"],    # self-reported 2026-09-09: lab is 210, not 205
    # Both moved into the Marois lab rooms on 6, reported 2026-09-09. This also
    # settles Xie's Marois/Woodman lab conflict: 221C was the Woodman room.
    "zengbo xie":          ["627A"],
    "lanting qiu":         ["627B"],
}

# People who belong in the directory but have no room in the spreadsheet yet.
# An empty "m" renders as a TBD chip instead of a room number. Delete the entry
# once the spreadsheet carries them, or the person will appear twice.
EXTRA_PEOPLE = [
    {"first": "Shaina", "last": "Munin", "role": "Lecturer",
     "title": "Senior Lecturer", "lab": None, "rooms": ["503"]},
    # Her only spreadsheet room was 205, withheld now that Kaczkurkin has said
    # the lab is 210. Without this entry she would disappear from the directory.
    {"first": "Leighton", "last": "Durham", "role": "Grad student",
     "lab": "BRAINS", "rooms": ["210"]},

    # Graduate students on the department roster with no room in the space sheet,
    # added 2026-09-09 from
    # https://www.vanderbilt.edu/psychological_sciences/people/?group=graduate_student
    # Only students advised by Wilson Hall faculty; that roster also covers Peabody.
    {"first": "Elton",     "last": "Cross",        "nick": "Ellie",
     "role": "Grad student", "lab": "Gauthier"},
    {"first": "Adrian",    "last": "Wong",         "role": "Grad student", "lab": "Gauthier",
     "rooms": ["429"]},   # self-reported 2026-09-09, taking over from Sajad AhmadNabi
    # 402 assigned 2026-09-09. It is shared with David Ricci, who stays -- confirmed,
    # so do not read the sheet's lone VACANT slot there as room for only one more.
    {"first": "Francesca", "last": "de Marneffe",  "role": "Grad student", "lab": "Park",
     "rooms": ["402"]},
    {"first": "Alenka",    "last": "Doyle",        "role": "Grad student", "lab": "Woodman"},
    {"first": "Daniel",    "last": "Garcia-Barnett", "role": "Grad student", "lab": "Marois"},
    {"first": "Isabella",  "last": "Jackson",      "role": "Grad student", "lab": "Watts"},
    {"first": "Justin",    "last": "Jaraczewski",  "role": "Grad student", "lab": "Womelsdorf"},
    {"first": "Ashna",     "last": "Ramiah",       "role": "Grad student", "lab": "BRAINS"},
    {"first": "Yuerou",    "last": "Tang",         "role": "Grad student", "lab": "Tong"},
    {"first": "Andrew",    "last": "Tornatore",    "role": "Grad student", "lab": "Polyn"},
    {"first": "Ella",      "last": "Weeks",        "role": "Grad student", "lab": "Woodman"},
    {"first": "Minghua",   "last": "Zhang",        "role": "Grad student", "lab": "Polyn"},
    # The sheet's "Wang, Joanna (Staff)" in 213A is this same person; she goes by
    # Ziqi. Reported 2026-09-09, superseding the earlier reading of the 2026-08-06
    # orientation list as naming a second Wang. "joanna wang" is excluded above,
    # so this entry carries her room. The sheet called Joanna Wang staff; she is a
    # grad student -- clarified 2026-09-09, so do not restore the staff role.
    {"first": "Ziqi",      "last": "Wang",         "role": "Grad student", "lab": "Park",
     "rooms": ["213A", "402"]},   # 402 assigned 2026-09-09, alongside the lab room

    # Self-reported 2026-09-09. He works across the Wallace and Ramachandran
    # labs; a person carries one lab here, so this records where his desk is.
    {"first": "Adam", "last": "Tiesman", "role": "Grad student",
     "lab": "Wallace", "rooms": ["614A"]},

    # Reported 2026-09-09, in 013 following Chrissy Suell's departure. Role is
    # provisional -- recorded as Staff pending confirmation from the Vice Chair.
    {"first": "Kris", "last": "Clifft", "role": "Staff",
     "lab": "Constantinidis", "rooms": ["013"]},
]

# Confirmed spellings. The spreadsheet holds both variants for these people;
# without an entry here the merge picks one arbitrarily. Add a line whenever
# review.txt flags a fuzzy merge and you know which spelling is correct.
SPELLING_LAST = {
    "meuller": "Mueller",          # confirmed: Melina Mueller
    "mueller": "Mueller",
    "herculano": "Herculano-Houzel",   # spreadsheet truncates the surname
    "rbiez": "Rbeiz",                  # department roster spelling
    "rbeiz": "Rbeiz",
    # Confirmed 2026-09-09; the sheet holds a different spelling in each room.
    "abbspoor":   "Abbaspoor",
    "ahmadnabi":  "Ahmadi Nebi",
    "ahmadinabi": "Ahmadi Nebi",
    "yildrim":    "Yildirim",
    "yildrem":    "Yildirim",
}
SPELLING_FIRST = {
    "jinkyeok": "Jinhyeok",   # department roster spelling
    "toni": "Antonia",        # Antonia Kaczkurkin, self-reported 2026-09-09
    "anikita": "Ankita",      # Ankita Mohan, confirmed 2026-09-09
}

# Short forms that will not merge on edit distance alone.
NICKNAMES = {
    "allie": "alexandra", "alexandra": "alexandra",
    "penny": "zheyue", "zheyue": "zheyue",
    "mia": "maria", "maria": "maria",
    "rye": "russell", "russell": "russell",
    "sally": "seohyun", "seohyun": "seohyun",
    "sophy": "yihan", "yihan": "yihan",
    "ali": "seyed", "seyed": "seyed",
    "toni": "antonia", "tom": "thomas", "ginni": "virginia",
}

# Rooms everyone actually asks for at the front door.
FACILITY_LABELS = {
    "301":  "Department Main Office",
    "301F": "Department Chair",
    "308A": "Vice Chair",
    "313":  "Grad Student Office",   # the sheet says just "Grad Student" here
    "301AA":"Kitchen",
    "315":  "Mail Room & Department Lounge",
    "117":  "Lactation Room",
    "215":  "Copier Room",
}

ROLE_PATTERNS = [
    ("Faculty",     re.compile(r"\bfaculty\b|\bemeritus\b", re.I)),
    ("Lecturer",    re.compile(r"\blecturer\b", re.I)),
    ("Post-doc",    re.compile(r"post[\s\-]?doc", re.I)),
    ("Grad student",re.compile(r"grad(?:uate)?\s*student|\bneuro\s*student\b", re.I)),
    ("Staff",       re.compile(r"\bstaff\b|\bmanager\b|\bofficer\b|\bspecialist\b"
                               r"|\bcoordinator\b|research\s*assistant\b|\bemployee\b"
                               r"|\bgrants?\b|administrative", re.I)),
    ("Student",     re.compile(r"\bstudent\b|\bundergrad\b", re.I)),
]

def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")

def norm(s):
    return re.sub(r"\s+", " ", strip_accents(str(s or "")).lower()).strip()

def parse_mdy(s):
    m, d, y = (int(x) for x in s.split("/"))
    if y < 100:
        y += 2000
    return datetime.date(y, m, d)

# --- Room parsing ------------------------------------------------------------

RE_ROOM = re.compile(r"^\s*([0-9]{3}[A-Z]{0,2}|000[A-Z]{0,2})\s*Wilson\s*Hall\s*(?:\((.*)\))?", re.I)

FLOOR_ORDER = ["Basement", "1st Floor", "2nd Floor", "3rd Floor",
               "4th Floor", "5th Floor", "6th Floor"]

def parse_room(raw):
    m = RE_ROOM.match(raw)
    if not m:
        return None, None
    return m.group(1).upper(), (m.group(2) or "").strip()

# How each lab is written out. Most read "<PI> Lab"; these do not.
LAB_DISPLAY = {
    "OPlab": "OPlab",
    "CATlab": "CATlab",
    "BRAINS": "BRAINS Lab",
}
def lab_parts(pi):
    """The two PIs in a shared-room label, or None for an ordinary lab."""
    return [x.strip() for x in pi.split("/")] if "/" in pi else None

def lab_name(pi):
    # A label naming two PIs is not a joint lab -- there are none in the department.
    # It is one space the two share, so it reads "Hoffman/Womelsdorf shared space"
    # rather than inventing a "Hoffman / Womelsdorf Lab" that nobody belongs to and
    # that showed "0 listed" because every member is recorded under one PI or other.
    parts = lab_parts(pi)
    if parts:
        return "/".join(parts) + " shared space"
    return LAB_DISPLAY.get(pi, f"{pi} Lab")

# The space spreadsheet is read by staff who will not all know the lab acronyms,
# so it names the faculty member instead. The public directory still uses the
# name the lab goes by.
LAB_PI = {
    "BRAINS": "Kaczkurkin",
    "OPlab":  "Gauthier",
    "CATlab": "Palmeri",
}
def lab_pi(pi):
    return LAB_PI.get(pi, pi)

def canon_lab(name):
    """Fold the spreadsheet's PI-name typos into one spelling."""
    fixes = {
        "constantindis": "Constantinidis", "constantinidis": "Constantinidis",
        "ramchandran": "Ramachandran",     "ramachandran": "Ramachandran",
        "shaefer": "Schaefer",             "schaefer": "Schaefer",
        "herculano": "Herculano-Houzel",
        "kaczkurkin": "BRAINS",        # the lab goes by BRAINS Lab
        "gauthier": "OPlab",
        "palmeri": "CATlab",
    }
    return fixes.get(norm(name), name.strip())

# --- Person parsing ----------------------------------------------------------

def split_parens(s):
    """Return (text outside parens, [contents of each paren group])."""
    groups, depth, buf, out = [], 0, "", ""
    for ch in s:
        if ch == "(":
            depth += 1
            if depth == 1:
                continue
        elif ch == ")":
            depth -= 1
            if depth == 0:
                groups.append(buf); buf = ""
                continue
            if depth < 0:
                depth = 0
                continue
        if depth:
            buf += ch
        else:
            out += ch
    if buf:
        groups.append(buf)
    return out, groups

def title_name(part):
    """Capitalise a name fragment without wrecking McNamara / Kekes-Szabo / O'Brien."""
    out = []
    for word in part.split():
        if re.search(r"[a-z][A-Z]", word) or word.isupper() and len(word) <= 3:
            out.append(word)
        else:
            w = word[:1].upper() + word[1:]
            w = re.sub(r"([-'])(\w)", lambda m: m.group(1) + m.group(2).upper(), w)
            out.append(w)
    return " ".join(out)

def parse_person(raw_assignee, room_lab):
    """
    -> dict(name, first, last, nick, role, lab) or None if the cell is not a person.
    Raises SkipRow (via returning ('expired', date)) when a past last-day is present.
    """
    s = re.sub(r"\s+", " ", str(raw_assignee or "")).strip()
    if not s:
        return None
    if RE_POINTER.match(s) or RE_NON_PERSON.match(s):
        return None

    m = RE_LASTDAY.search(s)
    if m:
        try:
            if parse_mdy(m.group(1)) <= TODAY:
                return ("expired", s)
        except ValueError:
            pass

    s = RE_INTERNAL.sub(" ", s)
    s = RE_DATEISH.sub(" ", s)
    s = s.replace("*", " ")
    s = re.sub(r"\s+", " ", s).strip(" ,;-")

    if norm(s) in NON_PERSON_EXACT or not s:
        return None

    outside, groups = split_parens(s)

    role, lab, nick = None, None, None
    for g in groups:
        g = g.strip(" ,;")
        if not g:
            continue
        matched = False
        for label, pat in ROLE_PATTERNS:
            if pat.search(g):
                if role is None:
                    role = label
                matched = True
                break
        # "Gauthier Grad Student", "Womelsdorf Lab", "Kaczkurkin Lab"
        lm = re.match(r"^\s*([A-Z][A-Za-z\-]+)\s+(?:Lab\b|Grad|Post|Neuro|Personnel)", g)
        if lm and lab is None:
            lab = canon_lab(lm.group(1))
            matched = True
        elif re.fullmatch(r"[A-Z][A-Za-z\-]+\s+Lab", g.strip()) and lab is None:
            lab = canon_lab(g.split()[0]); matched = True
        if not matched and re.fullmatch(r"[A-Z][a-zA-Z]{1,12}", g.strip()):
            nick = g.strip()

    outside = re.sub(r"\s*[|/].*$", "", outside)          # "Faculty | Biomedical ENG"
    outside = re.sub(r"\s*-\s*work space\s*$", "", outside, flags=re.I)
    outside = re.sub(r"\b(faculty|staff|student|unfunded|adjoint|emeritus)\b", " ",
                     outside, flags=re.I)
    outside = re.sub(r"\s+", " ", outside).strip(" ,;-")

    if not outside or len(outside) < 3:
        return None
    if re.search(r"\b(room|office|lab|space|door|desk)\b", outside, re.I) and "," not in outside:
        return None

    if "," in outside:
        last, first = (p.strip() for p in outside.split(",", 1))
    else:
        bits = outside.split()
        if len(bits) < 2:
            return None
        first, last = " ".join(bits[:-1]), bits[-1]

    first, last = title_name(first), title_name(last)
    last  = SPELLING_LAST.get(norm(last), last)
    first = SPELLING_FIRST.get(norm(first), first)
    if not first or not last:
        return None

    from_room = False
    if lab is None and room_lab:
        lm = re.match(r"^([A-Z][A-Za-z\-]+)(?:/([A-Z][A-Za-z\-]+))?\s+Lab", room_lab)
        if lm:
            lab = canon_lab(lm.group(1))
            from_room = True
            if lm.group(2):
                # The room is shared; naming the first lab is an arbitrary choice.
                from_room = "shared"

    return {"first": first, "last": last, "nick": nick, "role": role,
            "lab": lab, "from_room": from_room}

# --- Duplicate merging -------------------------------------------------------

def edit_le1(a, b):
    if a == b:
        return True
    if abs(len(a) - len(b)) > 1:
        return False
    if len(a) > len(b):
        a, b = b, a
    i = j = 0
    slack = True
    while i < len(a) and j < len(b):
        if a[i] == b[j]:
            i += 1; j += 1
        elif slack:
            slack = False
            j += 1
            if len(a) == len(b):
                i += 1
        else:
            return False
    return True

def near(a, b):
    """Equal, one edit apart, or one adjacent transposition apart."""
    if edit_le1(a, b):
        return True
    if len(a) == len(b):
        diff = [i for i in range(len(a)) if a[i] != b[i]]
        if len(diff) == 2 and diff[1] == diff[0] + 1:
            i, j = diff
            return a[i] == b[j] and a[j] == b[i]
    return False


class DSU:
    def __init__(self): self.p = {}
    def find(self, x):
        self.p.setdefault(x, x)
        while self.p[x] != x:
            self.p[x] = self.p[self.p[x]]; x = self.p[x]
        return x
    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb: self.p[rb] = ra

# --- Main --------------------------------------------------------------------

def main():
    wb = openpyxl.load_workbook(XLSX, data_only=True)

    rooms = {}          # room number -> {num, floor, kind, lab, sort}
    excluded_names = set()
    all_room_nums = set()
    placements = []     # (person dict, room number)
    expired, dropped = [], []

    for ws in wb.worksheets:
        floor = ws.title.strip()
        for raw_room, raw_who, *_ in ws.iter_rows(min_row=2, values_only=True):
            if raw_room is None:
                continue
            num, kind = parse_room(str(raw_room))
            if num is None:
                continue

            kind_clean = RE_INTERNAL.sub("", kind).strip(" ,-")
            if HIDE_ANIMAL_FACILITY and norm(kind_clean) in ANIMAL_ROOM_KINDS:
                continue
            if norm(kind_clean) in HIDE_ROOM_KINDS:
                continue
            all_room_nums.add(num)

            # Parse the person up front, purely so an EXCLUDE_PEOPLE name is audited
            # even when the room it sits in is itself excluded. The room is still
            # registered below, so vacating an office leaves the office listed.
            person = parse_person(raw_who, kind_clean)
            person_excluded = (isinstance(person, dict) and
                               norm(f"{person['first']} {person['last']}") in EXCLUDE_PEOPLE)
            if person_excluded:
                excluded_names.add(f"{person['first']} {person['last']}")

            if num in EXCLUDE_ROOMS:
                continue

            lab = labs2 = None
            lm = re.match(r"^([A-Za-z\-]+)(?:\s*/\s*([A-Za-z\-]+))?\s+Lab", kind_clean)
            if lm and norm(lm.group(1)) not in {"vacant", "temp", "research"}:
                lab = canon_lab(lm.group(1))
                if lm.group(2) and norm(lm.group(2)) not in {"vacant"}:
                    labs2 = canon_lab(lm.group(2))

            r = rooms.setdefault(num, {
                "num": num, "floor": floor,
                "kind": FACILITY_LABELS.get(num, kind_clean),
                "lab": lab, "lab2": labs2,
                "sort": (FLOOR_ORDER.index(floor) if floor in FLOOR_ORDER else 9,
                         int(re.match(r"\d+", num).group()), num),
            })
            if num in ROOM_LAB_OVERRIDES:
                r["lab"] = ROOM_LAB_OVERRIDES[num]
            if lab and not r["lab"]:
                r["lab"] = lab
            if labs2 and not r.get("lab2"):
                r["lab2"] = labs2

            if person_excluded:
                continue
            p = person
            if p is None:
                continue
            if isinstance(p, tuple):
                expired.append((num, p[1])); continue
            if p["role"] is None:
                kl = norm(kind_clean)
                if "grad student" in kl:      p["role"] = "Grad student"
                elif "post-doc" in kl:        p["role"] = "Post-doc"
                elif num.startswith("301"):   p["role"] = "Staff"
            placements.append((p, num))

    # --- merge people ---------------------------------------------------------
    dsu = DSU()
    for i, (p, _) in enumerate(placements):
        dsu.find(i)
    def canon_given(x):
        x = norm(x).split()[0] if norm(x) else ""
        return NICKNAMES.get(x, x)
    keys = [(norm(p["last"]), canon_given(p["first"]), canon_given(p["nick"] or ""))
            for p, _ in placements]
    fuzzy = []
    for i in range(len(keys)):
        for j in range(i + 1, len(keys)):
            li, fi, ni = keys[i]
            lj, fj, nj = keys[j]
            same_first = fi == fj or (ni and ni == fj) or (nj and nj == fi)
            if li == lj and same_first:
                dsu.union(i, j)
            elif li == lj and near(fi, fj):
                dsu.union(i, j); fuzzy.append((keys[i], keys[j], "first-name variant"))
            elif fi == fj and near(li, lj):
                dsu.union(i, j); fuzzy.append((keys[i], keys[j], "last-name variant"))

    groups = defaultdict(list)
    for i, (p, room) in enumerate(placements):
        groups[dsu.find(i)].append((p, room))

    ROLE_RANK = {"Faculty": 0, "Lecturer": 1, "Post-doc": 2, "Grad student": 3,
                 "Staff": 4, "Student": 5, None: 9}

    # A lab is named for its PI. Anyone whose surname names a lab, and who has no
    # role recorded in the spreadsheet, is that PI -- e.g. "Schaefer, Jon".
    pi_surnames = {norm(l) for r in rooms.values() for l in
                   re.split(r"\s*/\s*", r["lab"] or "") if l}

    # Surname -> lab key, so a PI can be put in their own lab. Rooms shared by two
    # labs are labelled "(A/B Lab)", and anyone in them whose own cell names no lab
    # would otherwise inherit whichever name comes first -- which put Womelsdorf in
    # the Hoffman lab. Faculty only: a student sharing a PI's surname is not the PI.
    pi_by_surname = {}
    for r in rooms.values():
        for k in re.split(r"\s*/\s*", r["lab"] or ""):
            if k:
                pi_by_surname[norm(lab_pi(k))] = k

    people, conflicts, roomless, joint_guess = [], [], [], []
    for members in groups.values():
        for pp, _ in members:
            if pp["role"] is None and norm(pp["last"]) in pi_surnames:
                pp["role"] = "Faculty"
        best = min((p for p, _ in members), key=lambda p: ROLE_RANK.get(p["role"], 9))
        role = ROLE_OVERRIDES.get(norm(f"{best['first']} {best['last']}"), best["role"])
        name_counts = defaultdict(int)
        for p, _ in members:
            name_counts[(p["first"], p["last"])] += 1
        (first, last), _ = max(name_counts.items(), key=lambda kv: kv[1])
        nick = next((p["nick"] for p, _ in members if p["nick"]), None)
        extra = EXTRA_ROOMS.get(norm(f"{first} {last}"), [])
        rms = sorted({r for _, r in members if r in rooms} | {r for r in extra if r in rooms},
                     key=lambda r: rooms[r]["sort"])
        drop = set(REMOVE_PLACEMENTS.get(norm(f"{first} {last}"), []))
        rms = [r for r in rms if r not in drop]
        if not rms:
            roomless.append(f"{first} {last}")
            continue

        labs = sorted({p["lab"] for p, _ in members if p["lab"]})
        if not labs:
            # A lab attribution can be lost when the room it came from is
            # withheld, so fall back to the labs of the rooms actually occupied.
            labs = sorted({rooms[r]["lab"] for r in rms if rooms.get(r, {}).get("lab")})
        # Flag a lab attribution that came only from a room shared by two labs:
        # the sheet names one of them first, and that choice is arbitrary.
        if role != "Faculty" and all(pp.get("from_room") == "shared"
                                     for pp, _ in members if pp["lab"]):
            if any(pp["lab"] for pp, _ in members):
                joint_guess.append((f"{first} {last}", role or "no role",
                                    ", ".join(rms)))

        if role == "Faculty" and norm(last) in pi_by_surname:
            labs = [pi_by_surname[norm(last)]]

        override = LAB_OVERRIDES.get(norm(f"{first} {last}"))
        if override:
            labs = [canon_lab(override)]
        elif len(labs) > 1:
            conflicts.append((f"{first} {last}", "labs", labs))
        people.append({
            "n": f"{first} {last}",
            "t": TITLES.get(norm(f"{first} {last}"))
                 or STAFF_TITLES.get(norm(f"{first} {last}")),
            "s": last, "f": first,
            "k": nick,
            "r": role,
            "l": labs[0] if labs else None,
            "l2": canon_lab(SECOND_LAB[norm(f"{first} {last}")])
                  if norm(f"{first} {last}") in SECOND_LAB else None,
            "m": rms,
        })

    for e in EXTRA_PEOPLE:
        name = f"{e['first']} {e['last']}"
        if norm(name) in {norm(p["n"]) for p in people}:
            continue          # the spreadsheet caught up; the entry is redundant
        people.append({
            "n": name, "s": e["last"], "f": e["first"], "k": e.get("nick"),
            "t": e.get("title"), "r": e.get("role"),
            "l": canon_lab(e["lab"]) if e.get("lab") else None,
            "l2": canon_lab(SECOND_LAB[norm(name)]) if norm(name) in SECOND_LAB else None,
            "m": [r for r in e.get("rooms", []) if r in rooms],
            "tbd": not e.get("rooms"),
        })

    people.sort(key=lambda p: (norm(p["s"]), norm(p["f"])))

    # --- room occupancy -------------------------------------------------------
    occ = defaultdict(list)
    for p in people:
        for r in p["m"]:
            occ[r].append(p["n"])

    room_list = sorted(rooms.values(), key=lambda r: r["sort"])
    for r in room_list:
        r["people"] = sorted(occ.get(r["num"], []))
        if r.get("lab2"):
            r["lab"] = (r["lab"] or "") + " / " + r["lab2"]
        # Show the canonical lab name, not whatever the sheet typed: the file has
        # "Herculano Lab" and "Constantindis Lab" for labs named elsewhere in full.
        if r["lab"] and r["num"] not in FACILITY_LABELS:
            r["kind"] = lab_name(r["lab"])
        r.pop("lab2", None)
        del r["sort"]

    def hide_empty(r):
        # An unlabelled empty room says nothing at all, so it goes too.
        return not r["people"] and (not (r["kind"] or "").strip()
                                    or HIDE_WHEN_EMPTY.search(r["kind"]))
    dropped_empty = [r["num"] for r in room_list if hide_empty(r)]
    room_list = [r for r in room_list if not hide_empty(r)]

    # --- labs -----------------------------------------------------------------
    labs = defaultdict(lambda: {"rooms": [], "people": []})
    for r in room_list:
        for L in (r["lab"], r.get("lab2")):
            if L:
                labs[L]["rooms"].append(r["num"])
    for p in people:
        for L in (p.get("l"), p.get("l2")):
            if L:
                labs[L]["people"].append(p["n"])
    # Each lab row names the faculty member as well as the lab. The directory shows
    # the name a lab goes by -- OPlab, CATlab, BRAINS Lab -- and a visitor may only
    # know the professor. LAB_PI already maps those acronyms to a surname; a shared
    # room label like "Hoffman / Womelsdorf" names both.
    faculty_by_surname = defaultdict(list)
    for p_ in people:
        if p_["r"] in ("Faculty", "Lecturer"):
            faculty_by_surname[norm(p_["s"])].append(p_["n"])

    def lab_faculty(key):
        # Resolved here rather than in the page: the acronym-to-surname mapping
        # lives in this file, and matching on surname in JS would pick up the
        # wrong person in a lab that has more than one faculty member.
        names = []
        for part in key.split("/"):
            for n in faculty_by_surname.get(norm(lab_pi(part.strip())), []):
                if n not in names:
                    names.append(n)
        return " & ".join(names)

    lab_list = [{"n": lab_name(k), "pi": k, "fac": lab_faculty(k) or None,
                 "rooms": v["rooms"], "people": sorted(set(v["people"]))}
                for k, v in sorted(labs.items())]

    payload = {
        "updated": TODAY.isoformat(),
        "floors": [f for f in FLOOR_ORDER if any(r["floor"] == f for r in room_list)],
        "people": people,
        "rooms": room_list,
        "labs": lab_list,
        "labNames": {k: lab_name(k) for k in labs},
        "info": INFO_CARDS,
    }

    blob = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    (HERE / "data.js").write_text(
        "// Generated by build.py -- do not edit by hand.\n"
        "window.WH_DATA = " + blob + ";\n", encoding="utf-8")

    # index.html and data.js are cached independently by the CDN, so without a
    # content-stamped URL a reader can get a new page with a stale dataset.
    stamp = hashlib.sha1(blob.encode("utf-8")).hexdigest()[:10]
    idx = HERE / "index.html"
    html = idx.read_text(encoding="utf-8")
    html2 = re.sub(r'<script src="data\.js(?:\?v=[0-9a-f]+)?"></script>',
                   f'<script src="data.js?v={stamp}"></script>', html)
    if html2 != html:
        idx.write_text(html2, encoding="utf-8")
    if 'data.js?v=' not in html2:
        print("WARNING: could not stamp the data.js URL in index.html", file=sys.stderr)

    # --- reviewer report (local only) ----------------------------------------
    lines = [f"Wilson Hall directory build -- {TODAY}", ""]
    lines.append(f"people published : {len(people)}")
    lines.append(f"rooms published  : {len(room_list)}")
    lines.append(f"labs published   : {len(lab_list)}")
    lines.append("")
    lines.append(f"DROPPED, on the EXCLUDE_PEOPLE list in build.py ({len(excluded_names)}):")
    for n in sorted(excluded_names):
        lines.append(f"           {n}")
    lines.append("")
    lines.append(f"DROPPED, every room excluded so nothing left to list ({len(roomless)}):")
    for n in sorted(roomless):
        lines.append(f"           {n}")
    lines.append("")
    known = ({norm(p["n"]) for p in people}
             | {norm(n) for n in excluded_names}
             | {norm(n) for n in roomless})
    dead = []
    for label, keys in (("EXCLUDE_PEOPLE", EXCLUDE_PEOPLE),
                        ("REMOVE_PLACEMENTS", REMOVE_PLACEMENTS),
                        ("EXTRA_ROOMS", EXTRA_ROOMS),
                        ("ROLE_OVERRIDES", ROLE_OVERRIDES),
                        ("LAB_OVERRIDES", LAB_OVERRIDES),
                        ("SECOND_LAB", SECOND_LAB),
                        ("TITLES", TITLES),
                        ("STAFF_TITLES", STAFF_TITLES)):
        for k in sorted(keys):
            if k not in known:
                dead.append(f"{label}: {k}")
    lines.append(f"!! CORRECTIONS KEYED ON A NAME THAT MATCHES NOBODY ({len(dead)}):")
    for x in dead:
        lines.append(f"   {x}")
    if dead:
        lines.append("   Usually a rename: the person is still listed under a new")
        lines.append("   spelling, so the correction silently stopped applying.")
    lines.append("")
    smell = [r for r in room_list if ANIMAL_SMELL.search(r["kind"] or "")]
    lines.append(f"!! PUBLISHED ROOMS LABELLED LIKE ANIMAL-FACILITY SPACE ({len(smell)}):")
    for r in smell:
        lines.append(f"  {r['num']:8} {r['kind']}")
    if smell:
        lines.append("   These are on the public page. If the label names an animal")
        lines.append("   facility function, add it to ANIMAL_ROOM_KINDS in build.py.")
    lines.append("")
    missing_rooms = sorted(r for r in EXCLUDE_ROOMS if r not in all_room_nums)
    if missing_rooms:
        lines.append(f"  ! EXCLUDE_ROOMS entries matching no room: {', '.join(missing_rooms)}")
        lines.append("")
    lines.append(f"ROOMS withheld via EXCLUDE_ROOMS ({len(EXCLUDE_ROOMS)}):")
    lines.append("           " + ", ".join(sorted(EXCLUDE_ROOMS)))
    lines.append("")
    lines.append(f"DROPPED, past last-day ({len(expired)}):")
    for num, s in expired:
        lines.append(f"  {num:8s} {s}")
    lines.append("")
    lines.append(f"FUZZY MERGES -- confirm these are the same person ({len(fuzzy)}):")
    for a, b, why in fuzzy:
        lines.append(f"  {a[1]} {a[0]}  <->  {b[1]} {b[0]}   ({why})")
    if fuzzy:
        lines.append("  ^ the published spelling for each of these was chosen arbitrarily.")
        lines.append("    Add the correct one to SPELLING_LAST / SPELLING_FIRST in build.py.")
    lines.append("")
    lines.append(f"LAB GUESSED FROM A SHARED ROOM -- verify ({len(joint_guess)}):")
    for n, r, rms in sorted(joint_guess):
        lines.append(f"  {n} ({r}) in {rms}")
    if joint_guess:
        lines.append("  Their own cell names no lab, so they took whichever lab the")
        lines.append("  room label happens to list first.")
    lines.append("")
    lines.append(f"CONFLICTING LAB AFFILIATION ({len(conflicts)}):")
    for name, _, labs_ in conflicts:
        lines.append(f"  {name}: {', '.join(labs_)}")
    lines.append("")
    used = {norm(p["n"]) for p in people}
    unmatched = sorted(k for k in list(TITLES) + list(STAFF_TITLES) if k not in used)
    lines.append(f"TITLES/STAFF_TITLES in build.py matching nobody ({len(unmatched)}):")
    for k in unmatched:
        lines.append(f"  {k}")
    lines.append("")
    untitled = sorted(p["n"] for p in people
                      if p["r"] in ("Faculty", "Lecturer") and not p["t"])
    lines.append(f"Faculty/Lecturer with no title on the department page ({len(untitled)}):")
    for n in untitled:
        lines.append(f"  {n}")
    lines.append("")
    noroom = [p["n"] for p in people if not p["m"]]
    lines.append(f"AWAITING A ROOM, shown as TBD ({len(noroom)}):")
    for n in sorted(noroom):
        lines.append(f"           {n}")
    lines.append("")
    lines.append(f"animal-facility rooms withheld: {HIDE_ANIMAL_FACILITY}")
    lines.append(f"room types withheld: {', '.join(sorted(HIDE_ROOM_KINDS))}")
    lines.append("")
    lines.append(f"EMPTY OFFICES dropped ({len(dropped_empty)}):")
    lines.append("           " + ", ".join(dropped_empty))
    (HERE / "review.txt").write_text("\n".join(lines), encoding="utf-8")

    print("\n".join(lines[:6]))
    print(f"\nwrote data.js ({(HERE/'data.js').stat().st_size:,} bytes) and review.txt")

if __name__ == "__main__":
    main()
