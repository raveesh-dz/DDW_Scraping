"""
Check the 10 affected May 4 sessions on the live DDW website to see if
their sub-sessions still exist or have been removed.

Uses the same iframe/frame approach as browser.py / scraper.py.
"""
import re
import time

from playwright.sync_api import sync_playwright

TARGET_SESSIONS = [
    "Pancreatic Surgery",
    "Beyond the Support Role",
    "Colorectal Surgery Quick Shots: I",
    "Endoscopy ePoster Showcase",
    "Primary Sclerosing Cholangitis",
    "Probiotics/Prebiotics/Synbiotics",
    "Surgical Complications",
    "Pediatric Small Intestinal Disease",
    "Training and Education for Surgery",
    "Clinical Considerations for Overt Hepatic Encephalopathy",
]

URL = "https://eppro01.ativ.me/web/planner.php?id=DDWLITE26"
MAY4_DATE = "2026-05-04"


def collect_text(locator, timeout=5000):
    try:
        return locator.inner_text(timeout=timeout).strip()
    except Exception:
        return ""


def close_detail_view(page, frame_locator):
    """Dismiss any open detail overlay."""
    try:
        btn = frame_locator.locator(".closeViewButton").first
        btn.click(timeout=2000)
        time.sleep(0.5)
    except Exception:
        pass
    # Also try via page evaluation
    try:
        page.evaluate("""
            const btn = document.querySelector('iframe#link2_iframe')
                ?.contentDocument?.querySelector('.closeViewButton');
            if (btn) btn.click();
        """)
    except Exception:
        pass


def main():
    results = {t: {"found_cell": False, "sub_count": 0, "sub_names": [], "error": None}
               for t in TARGET_SESSIONS}

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)  # visible for debugging
        page = browser.new_page(viewport={"width": 1400, "height": 900})

        print("Navigating to DDW planner...")
        page.goto(URL, timeout=60000)
        time.sleep(5)

        # Dismiss sponsor splash using JS (bypass click interception)
        print("Dismissing sponsor splash via JS...")
        page.evaluate("""
            const splash = document.getElementById('planner_sponsor_splash');
            if (splash) { splash.style.display = 'none'; }
        """)
        time.sleep(1)

        # Click the Schedule tab
        print("Clicking Schedule tab...")
        try:
            tab = page.get_by_role("link", name=re.compile(r"Schedule", re.I))
            tab.click(timeout=10000)
        except Exception:
            try:
                # Force click even if something overlaps
                page.locator("a").filter(has_text="Schedule").first.click(force=True, timeout=5000)
            except Exception as e:
                print(f"Schedule tab failed: {e}")
        time.sleep(3)

        # Wait for and find the schedule iframe as a real Frame object
        print("Waiting for schedule iframe...")
        try:
            page.wait_for_selector("iframe#link2_iframe", state="attached", timeout=20000)
        except Exception as e:
            print(f"iframe selector wait: {e}")
        time.sleep(2)

        # Get the actual Frame object (not FrameLocator) — needed for .frames[] access
        frame_locator = page.frame_locator("iframe#link2_iframe").first
        try:
            frame_locator.locator("body").wait_for(state="visible", timeout=15000)
        except Exception:
            pass
        time.sleep(4)
        print("Schedule frame ready.")

        # Click May 4 tab inside the frame
        print("Clicking May 4 tab...")
        try:
            frame_locator.locator(
                f'.agenda_multi_day_selector[data-date="{MAY4_DATE}"]'
            ).click(timeout=10000)
            time.sleep(2)
        except Exception as e:
            print(f"May 4 tab error: {e}")

        # Container for May 4
        container = frame_locator.locator(
            f'.date_container.agenda_sublist[data-listdate="{MAY4_DATE}"]'
        ).first

        # Scroll to load all cells
        print("Scrolling to load all May 4 sessions...")
        cells_sel = ".agenda_cell.list_cell"
        prev = 0
        stable_rounds = 0
        for rnd in range(40):
            try:
                cnt = container.locator(cells_sel).count()
            except Exception:
                cnt = 0
            if cnt == prev:
                stable_rounds += 1
                if stable_rounds >= 2 and rnd > 0:
                    break
            else:
                stable_rounds = 0
                print(f"  Round {rnd}: {cnt} cells")
            try:
                if cnt > 0:
                    container.locator(cells_sel).nth(cnt - 1).scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            time.sleep(1.5)
            prev = cnt
        print(f"Total cells loaded: {prev}")

        cells = container.locator(cells_sel)
        cell_count = cells.count()
        print(f"Cell count for matching: {cell_count}")

        # First pass: find which cells match our targets
        matched = {}
        for j in range(cell_count):
            try:
                cell = cells.nth(j)
                title = collect_text(cell.locator(".session_title.list_cell_title span"), 500)
                if not title:
                    title = collect_text(cell.locator(".session_title"), 500)
                for target in TARGET_SESSIONS:
                    if target not in matched:
                        if target.lower() in title.lower() or title.lower() in target.lower():
                            matched[target] = j
                            print(f"  Matched '{target}' at cell {j}: {title[:80]}")
            except Exception:
                pass

        print(f"\nMatched {len(matched)}/{len(TARGET_SESSIONS)} targets.")
        not_found = [t for t in TARGET_SESSIONS if t not in matched]
        if not_found:
            print(f"Not matched: {not_found}")

        # Second pass: click each matched cell and count sub-sessions
        for target, cell_idx in matched.items():
            print(f"\nChecking: {target}")
            try:
                close_detail_view(page, frame_locator)

                cell = cells.nth(cell_idx)
                cell.scroll_into_view_if_needed(timeout=3000)
                time.sleep(0.3)
                cell.click()
                time.sleep(2.5)

                # Wait for detail iframe
                frame_locator.locator("#planner_agenda_detail_iframe").wait_for(
                    state="visible", timeout=15000
                )
                time.sleep(1.5)

                detail_fl = frame_locator.frame_locator("#planner_agenda_detail_iframe").first
                body_text = collect_text(detail_fl.locator("body"), timeout=8000)

                lines = [l.strip() for l in body_text.split("\n") if l.strip()]
                pres_lines = [l for l in lines if re.match(r"^(Mo|Sp|Su|Sa|Tu|We|Fr)\d{3,4}", l)]

                results[target]["found_cell"] = True
                results[target]["sub_count"] = len(pres_lines)
                results[target]["sub_names"] = pres_lines[:10]

                print(f"  Presentations on website: {len(pres_lines)}")
                for pl in pres_lines[:5]:
                    print(f"    {pl[:100]}")
                if not pres_lines and body_text:
                    print(f"  Body preview: {body_text[:200]}")

                close_detail_view(page, frame_locator)
                time.sleep(0.5)

            except Exception as e:
                results[target]["error"] = str(e)
                print(f"  ERROR: {e}")

        browser.close()

    # Summary
    print("\n" + "=" * 72)
    print("SUMMARY — May 4 sessions: website vs Apr 9 scrape")
    print("=" * 72)
    apr9_counts = {
        "Pancreatic Surgery": 9,
        "Beyond the Support Role": 6,
        "Colorectal Surgery Quick Shots: I": 6,
        "Endoscopy ePoster Showcase": 5,
        "Primary Sclerosing Cholangitis": 5,
        "Probiotics/Prebiotics/Synbiotics": 3,
        "Surgical Complications": 3,
        "Pediatric Small Intestinal Disease": 2,
        "Training and Education for Surgery": 1,
        "Clinical Considerations for Overt Hepatic Encephalopathy": 1,
    }
    for target in TARGET_SESSIONS:
        r = results[target]
        status = "FOUND" if r["found_cell"] else "NOT IN CELL LIST"
        site_count = r.get("sub_count", 0)
        apr9 = apr9_counts.get(target, "?")
        verdict = ""
        if r["found_cell"]:
            if site_count >= apr9:
                verdict = "-> SCRAPER MISSED THEM (still on site)"
            elif site_count == 0:
                verdict = "-> REMOVED or 0 parsed (check body)"
            else:
                verdict = f"-> PARTIAL: {site_count} remain (was {apr9})"
        print(f"\n{target}")
        print(f"  Cell status: {status}")
        print(f"  Apr 9 count: {apr9}")
        print(f"  Website now: {site_count}  {verdict}")
        if r.get("error"):
            print(f"  Error      : {r['error'][:120]}")


if __name__ == "__main__":
    main()
