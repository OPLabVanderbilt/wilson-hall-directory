#!/usr/bin/env python3
"""
Write an updated copy of the space spreadsheet.

    python3 make_updated_sheet.py

Every room in the original survives, in the original order -- including rooms the
public directory withholds (animal facility, storage, vacated offices). Only the
Assignee column changes: departed people are gone and the current occupants, as the
directory now records them, take their place.

Original Notes are preserved untouched. A new "2026-09 update" column records what
changed for each room, so nothing is lost silently.
"""

import json, re, datetime, shutil
from pathlib import Path
from collections import OrderedDict, defaultdict

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

import build   # reuse the parsing rules, so "person" means the same thing here

HERE = Path(__file__).parent
SRC = HERE / "2025_2026 Wilson Hall Space Assignments.xlsx"
OUT = HERE / "Wilson Hall Space Assignments - updated 2026-09-09.xlsx"
TODAY = "2026-09-09"

ROLE_OUT = {
    "Faculty": "Faculty", "Lecturer": "Lecturer", "Post-doc": "Post-Doc",
    "Grad student": "Grad Student", "Staff": "Staff", "Student": "Student",
}

def load_directory():
    txt = (HERE / "data.js").read_text(encoding="utf-8")
    return json.loads(txt.split("= ", 1)[1].rstrip().rstrip(";\n"))

def assignee_line(person, lab_names):
    """Render a person the way the sheet has always rendered them."""
    last, first = person["s"], person["f"]
    role = ROLE_OUT.get(person.get("r") or "")
    labs = [l for l in (person.get("l"), person.get("l2")) if l]
    # "Wallace, Mark (Wallace Faculty)" is noise; the sheet says "(Faculty)".
    if len(labs) == 1 and build.norm(build.lab_pi(labs[0])) == build.norm(last):
        labs = []
    # Faculty surname, not the lab acronym: staff reading this sheet may not
    # know that BRAINS is Kaczkurkin or CATlab is Palmeri.
    shown = "/".join(build.lab_pi(l) for l in labs)
    if shown and role:
        tail = f"{shown} {role}"
    elif shown:
        tail = f"{shown} Lab"
    else:
        tail = role
    return f"{last}, {first}" + (f" ({tail})" if tail else "")


RENAMED = {"kaczkurkin": "BRAINS Lab", "gauthier": "OPlab", "palmeri": "CATlab",
           "herculano": "Herculano-Houzel Lab", "constantindis": "Constantinidis Lab",
           "ramchandran": "Ramachandran Lab", "shaefer": "Schaefer Lab"}

def label_warning(kind, here):
    """The Room Number column is the key, so it is never rewritten -- but say when
    it no longer matches what is in the room."""
    k = (kind or "").lower()
    if here and "vacant" in k:
        return "label says Vacant; room is occupied"
    for old, new in RENAMED.items():
        if old in k:
            return f"lab now called {new}"
    return ""

def main():
    d = load_directory()
    lab_names = d.get("labNames", {})
    people_by_name = {p["n"]: p for p in d["people"]}
    by_norm = {build.norm(p["n"]): p for p in d["people"]}
    occupants = defaultdict(list)
    for p in d["people"]:
        for rm in p["m"]:
            occupants[rm].append(p)

    src = openpyxl.load_workbook(SRC)
    out = openpyxl.Workbook()
    out.remove(out.active)

    head_font = Font(name="Adobe Devanagari", size=16, bold=True)
    head_fill = PatternFill("solid", fgColor="FFF7E9F0")
    body_font = Font(name="Adobe Devanagari", size=12)
    note_font = Font(name="Adobe Devanagari", size=11, italic=True, color="FF808080")

    stats = defaultdict(int)
    changes = []

    for sname in src.sheetnames:
        ss = src[sname]
        ws = out.create_sheet(sname)

        # room number -> (kind, [original assignee strings], first note seen)
        rooms = OrderedDict()
        for raw_room, raw_who, *rest in ss.iter_rows(min_row=2, values_only=True):
            if raw_room is None:
                continue
            num, kind = build.parse_room(str(raw_room))
            if num is None:
                continue
            note = (rest[0] if rest else None) or ""
            entry = rooms.setdefault(str(raw_room).strip(),
                                     {"num": num, "kind": kind, "rows": [], "note": ""})
            entry["rows"].append("" if raw_who is None else str(raw_who).strip())
            if note and not entry["note"]:
                entry["note"] = str(note).strip()

        for col, title in enumerate(["Room Number", "Assignee", "Notes", "2026-09 update"], 1):
            c = ws.cell(row=1, column=col, value=title)
            c.font, c.fill = head_font, head_fill
        for col, w in zip("ABCD", [40.4, 70.6, 110.6, 46]):
            ws.column_dimensions[col].width = w
        ws.freeze_panes = "A2"

        r = 2
        for label, info in rooms.items():
            num, kind = info["num"], info["kind"]
            here = sorted(occupants.get(num, []), key=lambda p: (p["s"].lower(), p["f"].lower()))

            # Rows that never described a person: pointers, fixtures, door labels.
            kept = []
            for raw in info["rows"]:
                if not raw:
                    continue
                parsed = build.parse_person(raw, kind)
                if parsed is None and not re.match(r"^\s*vacant", raw, re.I):
                    kept.append(raw)

            had_people = any(isinstance(build.parse_person(x, kind), dict)
                             for x in info["rows"] if x)

            if here:
                lines = [assignee_line(p, lab_names) for p in here]
                note = "current occupants per directory" if had_people else "occupants added"
                warn = label_warning(kind, here)
                if warn:
                    note = f"{note}; {warn}"
                stats["rooms with people"] += 1
            elif kept:
                lines = kept
                note = ""
                stats["descriptive rooms kept"] += 1
            elif had_people:
                lines = ["Vacant"]
                # Someone can vacate a room without leaving the building; say which.
                moved, gone = [], []
                for raw in info["rows"]:
                    if not raw:
                        continue
                    parsed = build.parse_person(raw, kind)
                    if not isinstance(parsed, dict):
                        continue
                    nm = f"{parsed['first']} {parsed['last']}"
                    match = by_norm.get(build.norm(nm))
                    if match:
                        moved.append(f"{match['n']} -> {', '.join(match['m']) or 'no room yet'}")
                    else:
                        gone.append(nm)
                bits = []
                if moved:
                    bits.append("moved: " + "; ".join(sorted(set(moved))))
                if gone:
                    bits.append("no longer listed: " + ", ".join(sorted(set(gone))))
                note = f"vacated {TODAY} - " + " | ".join(bits) if bits else f"vacated {TODAY}"
                stats["rooms vacated"] += 1
                changes.append((sname, label, [x for x in info["rows"] if x]))
            else:
                lines = ["Vacant"]
                note = ""
                stats["already vacant"] += 1

            for i, line in enumerate(lines):
                ws.cell(row=r, column=1, value=label).font = body_font
                ws.cell(row=r, column=2, value=line).font = body_font
                if i == 0:
                    if info["note"]:
                        ws.cell(row=r, column=3, value=info["note"]).font = body_font
                    if note:
                        ws.cell(row=r, column=4, value=note).font = note_font
                r += 1

    # People the directory carries with no room yet.
    ws = out.create_sheet("Awaiting Room")
    for col, title in enumerate(["Name", "Role", "Lab", "Note"], 1):
        c = ws.cell(row=1, column=col, value=title)
        c.font, c.fill = head_font, head_fill
    for col, w in zip("ABCD", [34, 20, 24, 52]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "A2"
    r = 2
    for p in sorted((p for p in d["people"] if not p["m"]),
                    key=lambda p: (p["s"].lower(), p["f"].lower())):
        ws.cell(row=r, column=1, value=p["n"]).font = body_font
        ws.cell(row=r, column=2, value=ROLE_OUT.get(p.get("r") or "", "")).font = body_font
        lab = p.get("l")
        ws.cell(row=r, column=3,
                value=(build.lab_pi(lab) + " Lab") if lab else "").font = body_font
        ws.cell(row=r, column=4, value="no room recorded").font = note_font
        r += 1
    stats["people awaiting a room"] = r - 2

    out.save(OUT)

    print(f"wrote {OUT.name}")
    print()
    for k, v in stats.items():
        print(f"  {v:4d}  {k}")
    print()
    print(f"rooms vacated ({len(changes)}):")
    for sheet, label, was in changes:
        print(f"  {label}")
        for w in was:
            print(f"        was: {w}")

if __name__ == "__main__":
    main()
