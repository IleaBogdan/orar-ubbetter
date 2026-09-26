#!/usr/bin/env python3
"""
Parse UBB Cluj class schedule HTML into an .xlsx schedule like Tiberiu Popoviciu Highschool has.
Usage:
    python3 parse_orar.py <url> [<url> ...] [-o output.xlsx]

Each URL becomes one sheet (named after the file stem, e.g. "IA1").
"""

import re
import sys
import urllib.parse

import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

DISC = re.compile(r"/disc/")
CADRE = re.compile(r"/cadre/")
SALI = re.compile(r"/sali/")
SAPT = re.compile(r"sapt\.\s*([12])", re.I)
GR = re.compile(r"\bgr\.\s*(\d+)", re.I)
SUBGROUP = re.compile(r"^\d+/\d+$")

DAYS = {
    "luni": "Luni",
    "marti": "Marţi",
    "miercuri": "Miercuri",
    "joi": "Joi",
    "vineri": "Vineri",
    "sambata": "Sâmbătă",
    "duminica": "Duminică",
}


def clean(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def fetch(url: str) -> str:
    r = requests.get(url, timeout=60)
    r.raise_for_status()
    return r.text


def expand_table(trs):
    """Expand a table into a row/col grid, honouring rowspan/colspan."""
    grid = []
    for r, tr in enumerate(trs):
        if len(grid) <= r:
            grid.append([])
        row = grid[r]
        c = 0
        for cell in tr.find_all(["td", "th"], recursive=False):
            while c < len(row) and row[c] is not None:
                c += 1
            colspan = int(cell.get("colspan", 1))
            rowspan = int(cell.get("rowspan", 1))
            while len(row) < c + colspan:
                row.append(None)
            for cc in range(c, c + colspan):
                row[cc] = cell
            for rr in range(r + 1, r + rowspan):
                while len(grid) <= rr:
                    grid.append([])
                grow = grid[rr]
                while len(grow) < c + colspan:
                    grow.append(None)
                for cc in range(c, c + colspan):
                    grow[cc] = cell
            c += colspan
    width = max((len(row) for row in grid), default=0)
    for row in grid:
        while len(row) < width:
            row.append(None)
    return grid


def type_from_class(cell) -> str:
    classes = cell.get("class") or []
    for c in classes:
        if c.startswith("tip"):
            return c[3:].upper()
    return None


def extract_activity(elem):
    full = clean(elem.get_text(" ", strip=True))
    sapt = None
    m = SAPT.search(full)
    if m:
        sapt = int(m.group(1))

    subj = elem.find("a", href=DISC)
    teacher = elem.find("a", href=CADRE)
    room = elem.find("a", href=SALI)
    return {
        "subject": clean(subj.get_text(" ", strip=True)) if subj else None,
        "teacher": clean(teacher.get_text(" ", strip=True)) if teacher else None,
        "room": clean(room.get_text(" ", strip=True)) if room else None,
        "sapt": sapt,
        "type": type_from_class(elem),
    }


def parse_cell(cell):
    """Return [primary_activity, *group_overrides] for a data cell."""
    primary = extract_activity(cell)
    primary["groups"] = None
    overrides = []
    for span in cell.find_all(True, recursive=True):
        txt = clean(span.get_text(" ", strip=True))
        m = GR.search(txt)
        if m and span.find("a", href=DISC):
            act = extract_activity(span)
            act["groups"] = [m.group(1)]
            overrides.append(act)
    return [primary] + overrides


def parse_page(html: str):
    soup = BeautifulSoup(html, "lxml")
    title = clean(soup.title.get_text()) if soup.title else "Orar"
    table = soup.find("table")
    if table is None:
        raise ValueError("No <table> found in the HTML")

    trs = table.find_all("tr")
    grid = expand_table(trs)
    ncols = len(grid[0]) if grid else 0

    # subgroup columns: header <th> like "1011/1"
    subgroup_cols = []
    subgroup_names = {}
    for r, row in enumerate(grid):
        for c, cell in enumerate(row):
            if cell is not None and cell.name == "th":
                txt = clean(cell.get_text(" ", strip=True))
                if SUBGROUP.match(txt) and c not in subgroup_cols:
                    subgroup_cols.append(c)
                    subgroup_names[c] = txt
    subgroup_cols.sort()

    # day blocks: <th rowspan> at column 0 with a weekday name
    day_blocks = []  # (start_row, day, span)
    prev_day_cell = None
    for r, row in enumerate(grid):
        cell = row[0] if row else None
        if cell is not None and cell.name == "th" and cell is not prev_day_cell:
            prev_day_cell = cell
            txt = re.sub(r"\s+", "", cell.get_text("", strip=True)).lower()
            if txt in DAYS:
                span = int(cell.get("rowspan", 1))
                day_blocks.append((r, DAYS[txt], span))

    hours_order = []
    schedule = {}  # group -> day -> hour -> activity

    for start, day, span in day_blocks:
        for rr in range(start, start + span):
            hour_cell = grid[rr][1] if len(grid[rr]) > 1 else None
            hour = clean(hour_cell.get_text(" ", strip=True)) if hour_cell else ""
            if not hour or hour in ("Ore",):
                continue
            if hour not in hours_order:
                hours_order.append(hour)

            seen = set()
            for col in subgroup_cols:
                cell = grid[rr][col]
                if cell is None or id(cell) in seen:
                    continue
                seen.add(id(cell))
                activities = parse_cell(cell)
                covered = [c2 for c2 in subgroup_cols if grid[rr][c2] is cell]
                for c2 in covered:
                    name = subgroup_names[c2]
                    base = name.split("/")[0]
                    chosen = None
                    for act in activities:
                        if act["groups"] and base in act["groups"]:
                            chosen = act
                            break
                    if chosen is None:
                        chosen = activities[0]
                    if chosen.get("subject"):
                        schedule.setdefault(name, {}).setdefault(day, {})[hour] = chosen

    groups = [subgroup_names[c] for c in subgroup_cols]
    all_days = [d for _, d, _ in day_blocks]
    days = [d for d in all_days if any(schedule.get(g, {}).get(d) for g in groups)]

    return {
        "title": title,
        "groups": groups,
        "days": days,
        "hours": hours_order,
        "schedule": schedule,
    }


def fmt_subject(act) -> str:
    parts = []
    if act.get("sapt"):
        parts.append(f"sapt. {act['sapt']}:")
    parts.append(act["subject"])
    if act.get("type"):
        parts.append(f"[{act['type']}]")
    if act.get("teacher"):
        parts.append(f"({act['teacher']})")
    return " ".join(parts)


def write_sheet(ws, parsed):
    days = parsed["days"]
    hours = parsed["hours"]
    groups = parsed["groups"]
    schedule = parsed["schedule"]
    nhours = len(hours)

    ws.cell(row=1, column=1, value=parsed["title"])

    # day headers (row 3) and hour headers (row 4)
    for i, day in enumerate(days):
        ws.cell(row=3, column=2 + i * nhours, value=day)
        for j, hour in enumerate(hours):
            ws.cell(row=4, column=2 + i * nhours + j, value=hour)
    ws.cell(row=3, column=1, value="Grupa")
    ws.cell(row=4, column=1, value="Grupa")

    row = 5
    for g in groups:
        ws.cell(row=row, column=1, value=g)
        for i, day in enumerate(days):
            for j, hour in enumerate(hours):
                act = schedule.get(g, {}).get(day, {}).get(hour)
                if act:
                    ws.cell(row=row, column=2 + i * nhours + j, value=fmt_subject(act))
                    ws.cell(row=row + 1, column=2 + i * nhours + j, value=act["room"])
        row += 2

    # light formatting
    bold = Font(bold=True)
    fill = PatternFill("solid", fgColor="DDDDDD")
    for col in range(1, 2 + len(days) * nhours):
        ws.cell(row=3, column=col).font = bold
        ws.cell(row=3, column=col).fill = fill
        ws.cell(row=4, column=col).font = bold


def main(argv):
    urls = []
    output = "orar.xlsx"
    i = 0
    while i < len(argv):
        a = argv[i]
        if a in ("-o", "--output"):
            output = argv[i + 1]
            i += 2
        else:
            urls.append(a)
            i += 1

    if not urls:
        print("Url for an orar to be parsed needed!")
        sys.exit(1)

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for url in urls:
        stem = urllib.parse.urlparse(url).path.rsplit("/", 1)[-1].rsplit(".", 1)[0] or "orar"
        print(f"Fetching {url} ...")
        parsed = parse_page(fetch(url))
        ws = wb.create_sheet(title=stem[:31])
        write_sheet(ws, parsed)
        print(f"  -> {stem}: {len(parsed['groups'])} groups, "
              f"{len(parsed['days'])} days, {len(parsed['hours'])} hour slots")

    wb.save(output)
    print(f"Saved {output}")


if __name__ == "__main__":
    main(sys.argv[1:])
