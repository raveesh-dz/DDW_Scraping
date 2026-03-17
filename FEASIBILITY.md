# DDWLITE26 Schedule Scraping – Feasibility Summary

## Source

- **URL:** https://eppro01.ativ.me/web/planner.php?id=DDWLITE26  
- **Schedule:** Loaded in the **Schedule** tab (content comes from an iframe).

## How the site works

1. **Main page** is mostly a shell: it shows “Loading…” until JavaScript runs.
2. **Tabs:** Home, Schedule, Help. Each tab’s content is loaded in an **iframe** when you click the tab.
3. **Schedule tab** iframe URL:
   `https://eppro01.ativ.me/web/plantab.php?id=DDWLITE26&plannersession=true&tab=urn%3Aeventpilot%3Aplanner%3Aagenda%3A%3A`
4. **Data loading:** The actual schedule (days, sessions, sub-sessions) is **built by JavaScript** inside that iframe. There is no obvious public JSON/XML API for the agenda in the page source; the app uses `api.php` for tracking/login, not for raw schedule data.
5. **Session details:** Expanding a session (e.g. “AI Literacy for Gastroenterologists…”) to see sub-sessions is done in the same iframe (expand/collapse in the DOM).

## Feasibility: **Yes**

Scraping is feasible using **browser automation** (e.g. **Playwright** or Selenium):

| Requirement | Feasible? | How |
|-------------|-----------|-----|
| Go to Schedule page | Yes | Open planner URL, click “Schedule” tab so the iframe loads. |
| Scrape by day (Fri–Tue, May 1–5) | Yes | Once inside the schedule iframe, the DOM is expected to have day sections; we can group sessions by day. |
| Conference name, time, date, region | Yes | These fields should be present in the rendered HTML (e.g. in list/card elements); we parse them from the iframe DOM. |
| Sub-sessions (e.g. “Welcome and course introduction” 12:00–12:05) | Yes | For sessions that have sub-sessions, we can click to expand (or detect expandable elements) and then scrape the revealed rows. |
| One sheet per day (Fri May 1 … Tue May 5) | Yes | Export to Excel (or CSV) with one sheet per day. |

## Constraints / risks

- **No direct API:** We depend on the rendered HTML inside the schedule iframe. If the site changes its DOM structure or class names, the scraper may need updates.
- **Iframe + JS:** We must use a real browser (Playwright) to load the page, switch to the Schedule tab, wait for the iframe to load, then parse (and optionally click to expand sub-sessions). Simple `requests` + BeautifulSoup is not enough.
- **Rate limiting / blocking:** The site may throttle or block aggressive requests. The scraper should use a single browser session, reasonable waits, and avoid unnecessary reloads.
- **Exact selectors:** The exact CSS selectors for “day”, “session”, “sub-session”, “time”, “region” are not visible from the main page source. They will be determined by inspecting the **loaded** schedule iframe (e.g. “Inspect” in the browser after opening Schedule). The provided script uses flexible, exploratory selectors and can be tuned after one run.

## Recommended approach

1. **Tool:** Python + **Playwright** (with Chromium).
2. **Flow:**  
   - Open planner URL → click Schedule → wait for iframe → switch to iframe → wait for agenda to render.  
   - Find day sections (e.g. Friday May 1 … Tuesday May 5).  
   - For each day, find all session blocks; extract name, time, date, region (and any other visible fields).  
   - For each session that has sub-sessions: trigger expand (click/aria) and scrape sub-session title, start/end time, date.  
3. **Output:** Excel file with one sheet per day (Friday May 1, Saturday May 2, …, Tuesday May 5), columns such as: Conference/Session name, Start time, End time, Date, Region, Sub-session name, Sub-session start, Sub-session end (and any extra fields you want).

## Next step

Run the provided Playwright scraper once with a visible browser (`headless=False`) so you can see the schedule structure. Then adjust the selectors in the script to match the actual DOM (e.g. class names for day headers and session blocks). After that, you can run it headless and rely on the Excel export.
