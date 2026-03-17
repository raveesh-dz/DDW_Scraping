# DDW Conference Schedule Scraper

A Python tool that scrapes the **DDWLITE26** conference schedule from the official event planner and exports it to Excel with one sheet per day, keyword tagging, and ECCO-style formatting.

## Features

- **Full schedule capture** — Sessions for all conference days (May 1–5, 2026) from the Schedule tab
- **Session details** — Track, format, society, science type, and expandable sub-sessions/presentations
- **Keyword matching** — Populates Indication, Company, Product Name, and MOA from `DDW Keywords.xlsx`
- **Excel export** — One sheet per day, styled like the ECCO 2026 CRDI Planner (navy header, light blue rows, merged session columns)
- **Checkpointing** — Optional checkpoint save after each day; partial save on errors
- **Debug options** — Save iframe HTML or detail body text for selector tuning

## Project structure

```
.
├── scrape_ddw_schedule.py   # Main scraper (full run → Excel)
├── count_sessions.py        # Session counts per day (no Excel)
├── test_run.py              # Quick connectivity / iframe test
├── requirements.txt
├── DDW Keywords.xlsx        # Keyword lists for tagging (required)
├── FEASIBILITY.md           # Scraping approach and constraints
└── ddw_scraper/
    ├── config.py            # URLs, paths, flags, column definitions
    ├── browser.py           # Schedule tab and iframe handling
    ├── scraper.py           # Schedule and session-detail scraping
    ├── parser.py            # Date/time and presentation parsing
    ├── keywords.py          # Load and match keywords from Excel
    ├── exporter.py          # Excel write (ECCO-style)
    └── logger.py            # Timestamped logging
```

## Prerequisites

- **Python 3.10+** (uses `list[dict]` and similar type hints)
- **Chromium** (installed via Playwright)

## Installation

1. Clone the repository and go to the project directory:

   ```bash
   cd path/to/DDW6-Conference-Scraping
   ```

2. Create and activate a virtual environment:

   ```bash
   python -m venv .venv
   .venv\Scripts\activate   # Windows
   # source .venv/bin/activate   # macOS / Linux
   ```

3. Install dependencies and Playwright browser:

   ```bash
   pip install -r requirements.txt
   playwright install chromium
   ```

4. Ensure **DDW Keywords.xlsx** is in the project root (same folder as `scrape_ddw_schedule.py`). The scraper will run without it but keyword columns will be empty.

## Usage

### Full scrape (recommended)

```bash
python scrape_ddw_schedule.py
```

- Opens the planner in Chromium, clicks the Schedule tab, scrapes all days and sessions, and writes Excel to `results/`.
- Logs go to the console and to `scrape_ddw_schedule.log` (if `config.LOG_TO_FILE` is `True`).

### Session count per day

```bash
python count_sessions.py
```

Prints how many sessions the site shows for each day (Fri–Tue). Useful to verify the site layout before a full scrape.

### Connectivity test

```bash
python test_run.py
```

Checks that Playwright starts, the planner page loads, and the Schedule iframe appears. Use this to confirm the environment before a long run.

## Output

- **Excel file** — `results/ddw-conference-schedule-2026-YYYYMMDD-HHMMSS.xlsx`
  - One sheet per day: *May 1* … *May 5*
  - Columns: Priority, Coverage, Indication, Company, Product Name, Mechanism/Targets, Track, Science Type, Session Name/Type/Start/End/Location, Abstract #, Sub-Session Start/End/Title, Presenter, Sponsor/Collaboration, Affiliations
- **Checkpoint** (if enabled) — `results/ddw-conference-schedule-2026-YYYYMMDD-HHMMSS_checkpoint.xlsx`
- **Log** — `scrape_ddw_schedule.log` in the project root (when file logging is on)

The `results/` directory is created automatically on first run.

## Configuration

All settings are in **`ddw_scraper/config.py`**:

| Setting | Purpose |
|--------|---------|
| `HEADLESS` | `True` = no browser window; `False` = show browser (useful for first run / debugging) |
| `DEBUG_SAVE_HTML` | Save schedule iframe HTML to `schedule_iframe_debug.html` for selector tuning |
| `DEBUG_SAVE_ZERO_PRES_BODY` | Save detail body text for sessions with 0 presentations to `debug_body_text/` |
| `LOG_TO_FILE` | Append logs to `scrape_ddw_schedule.log` |
| `VERBOSE_SESSION_LOGS` | Log each session title and detail summary as it’s scraped |
| `CHECKPOINT_AFTER_EACH_DAY` | Write checkpoint Excel after each day |
| `ONLY_DATES` | Restrict to specific dates, e.g. `{"2026-05-01"}`; `None` = all days |

## Debugging and selector tuning

The schedule is inside an iframe and built with JavaScript. If sessions or days are missing:

1. Set `HEADLESS = False` and `DEBUG_SAVE_HTML = True` in `ddw_scraper/config.py`.
2. Run `python scrape_ddw_schedule.py` and let it save the iframe HTML.
3. Open `schedule_iframe_debug.html` in a browser and inspect the DOM (e.g. day sections, session blocks, class names).
4. Update the selectors in `ddw_scraper/scraper.py` and `ddw_scraper/browser.py` to match the current site structure.

## Documentation

- **[FEASIBILITY.md](FEASIBILITY.md)** — How the planner site works, why browser automation is required, constraints, and recommended approach.

## License

This project is provided as-is for personal or internal use. Ensure your use complies with the event planner’s terms of service and any applicable data policies.
