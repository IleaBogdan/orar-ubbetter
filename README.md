# orar-ubbetter

Parse UBB Cluj Faculty of Mathematics and Computer Science schedule pages
(e.g. `https://www.cs.ubbcluj.ro/files/orar/2026-1/grafic/IA1.html`) and turn them
into a clean Excel schedule, one sheet per year/programme, laid out like `exemplu.xlsx`.

## What it does

- Downloads the HTML schedule page(s).
- Expands the timetable grid (handling `rowspan`/`colspan`).
- Extracts, for every subgroup (e.g. `1011/1`, `1011/2`, `1012/1` ...):
  - subject name
  - teacher
  - room
  - type (course / seminar / lab)
  - odd/even week (`sapt. 1` / `sapt. 2`) and group overrides (`gr.XXXX`)
- Writes an `.xlsx` where each subgroup is a block of two rows (subjects on top,
  rooms below), with day headers and hour slots, mirroring `exemplu.xlsx`.

## Requirements

- Python 3.8+
- `requests`, `beautifulsoup4`, `lxml`, `openpyxl`

```bash
pip install requests beautifulsoup4 lxml openpyxl
```

## Usage

```bash
# single schedule
python3 parse_orar.py "https://www.cs.ubbcluj.ro/files/orar/2026-1/grafic/IA1.html"

# several schedules -> one sheet each
python3 parse_orar.py \
  "https://www.cs.ubbcluj.ro/files/orar/2026-1/grafic/IA1.html" \
  "https://www.cs.ubbcluj.ro/files/orar/2026-1/grafic/IA2.html" \
  -o orar.xlsx
```

Options:

- `-o, --output FILE` — output `.xlsx` path (default: `orar.xlsx`).

With no URL argument, it falls back to the example `IA1.html` link.

## Output layout

```
      |  Luni   ...                        |  Marţi   ...
Grupa | 8-9 | 9-10 | ... | 19-20           | 8-9 | ...
1011/1| Logica computationala [S] (...)   | ...
      | C512                               | ...
1011/2| ...
```

Each day is a block of 12 hour slots (`8-9` … `19-20`); empty days are omitted.
