# ical2sheet

Fetch a public iCalendar (`.ics`) feed and export events for one academic year into an `.xlsx` sheet — bucketed by semester and week.

## Install

```sh
cd ical2sheet
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

```sh
python ical2sheet.py <ics-url> [--year YYYY] [-o output.xlsx]
```

Example — Kappa Theta Tau, current academic year:

```sh
python ical2sheet.py \
  'https://calendar.google.com/calendar/ical/kappa.thetatau%40gmail.com/public/basic.ics'
```

Flags:

- `--year YYYY` — starting year of the academic year. `--year 2026` means AY 2026-27 (Summer 2026 through Spring 2027). Defaults to the current AY based on today's date.
- `-o path` — output file. Defaults to `calendar_AY<year>-<year+1>.xlsx`.
- `--exclude PATTERN` — drop events whose summary contains PATTERN (case-insensitive). Repeatable. Useful for cutting out noise, e.g. `--exclude birthday`.

## Output columns

`Summary | Semester | Day | Week | Date | Start Time | End Time | All-Day | Location | Description | Status | Recurring? | Meet Link | UID`

- **Semester** — `SUMMER`, `FALL`, `WINTER`, or `SPRING`
- **Day** — `SUN`, `MON`, …, `SAT`
- **Week** — 0 for the partial week containing the semester cutoff; 1+ for subsequent Sun–Sat weeks

## Academic year definition

An academic year runs from Theta Tau exec change-over (end of Spring) to the next. AY `Y` starts May 16 of year `Y`.

| Semester | Range |
|---|---|
| Summer `Y` | May 16 – Aug 20 |
| Fall `Y` | Aug 21 – Dec 14 |
| Winter `Y`-`Y+1` | Dec 15 – Jan 9 |
| Spring `Y+1` | Jan 10 – May 15 |

## Week numbering

Weeks are Sunday–Saturday. Week 1 begins on the **first Sunday strictly after** the semester's cutoff date. Anything between the cutoff and that Sunday is Week 0.

Example — Fall 2026: cutoff Fri Aug 21 → Week 0 is Aug 21–22, Week 1 starts Sun Aug 23.
