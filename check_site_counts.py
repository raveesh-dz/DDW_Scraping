"""
Quick site count check: load May 4 tab and count all free sessions.
Reusable for any day — set DAY_DATE below.
"""
import re, time
from playwright.sync_api import sync_playwright

URL      = "https://eppro01.ativ.me/web/planner.php?id=DDWLITE26"
DAYS     = {
    "2026-05-01": "May 1",
    "2026-05-02": "May 2",
    "2026-05-03": "May 3",
    "2026-05-04": "May 4",
    "2026-05-05": "May 5",
}

def scroll_load_all(frame_loc, container, label):
    cells_sel = ".agenda_cell.list_cell"
    prev = 0
    stable = 0
    for rnd in range(40):
        try: cnt = container.locator(cells_sel).count()
        except: cnt = 0
        if cnt == prev:
            stable += 1
            if stable >= 2 and rnd > 0: break
        else:
            stable = 0
            print(f"    {label} scroll {rnd}: {cnt} cells")
        try:
            if cnt > 0:
                container.locator(cells_sel).nth(cnt-1).scroll_into_view_if_needed(timeout=3000)
        except: pass
        time.sleep(1.5)
        prev = cnt
    return prev

def main():
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(viewport={"width":1400,"height":900})
        print("Navigating...")
        page.goto(URL, timeout=60000)
        time.sleep(5)
        page.evaluate("const s=document.getElementById('planner_sponsor_splash');if(s)s.style.display='none';")
        time.sleep(1)
        try:
            page.get_by_role("link", name=re.compile(r"Schedule", re.I)).click(timeout=10000)
        except: pass
        time.sleep(3)
        try:
            page.wait_for_selector("iframe#link2_iframe", state="attached", timeout=20000)
        except: pass
        time.sleep(2)
        fl = page.frame_locator("iframe#link2_iframe").first
        try: fl.locator("body").wait_for(state="visible", timeout=15000)
        except: pass
        time.sleep(4)

        results = {}
        for date_str, label in DAYS.items():
            print(f"\n  Clicking {label}...")
            try:
                fl.locator(f'.agenda_multi_day_selector[data-date="{date_str}"]').click(timeout=10000)
                time.sleep(2)
            except Exception as e:
                print(f"    Tab click failed: {e}")
                results[label] = {"total":0,"paid":0,"free":0}
                continue

            container = fl.locator(f'.date_container.agenda_sublist[data-listdate="{date_str}"]').first
            total = scroll_load_all(fl, container, label)
            print(f"    {label}: {total} total cells loaded")

            # Count paid vs free
            paid = 0
            cells = container.locator(".agenda_cell.list_cell")
            for j in range(total):
                try:
                    cell = cells.nth(j)
                    title_el = cell.locator(".session_title.list_cell_title span")
                    try: title = title_el.inner_text(timeout=300).strip()
                    except: title = ""
                    if not title:
                        try: title = cell.locator(".session_title").inner_text(timeout=300).strip()
                        except: title = ""
                    if "$$$" in title:
                        paid += 1
                except: pass

            results[label] = {"total": total, "paid": paid, "free": total - paid}

        browser.close()

    print("\n" + "="*55)
    print(f"{'Day':<8} {'Total':>7} {'Paid':>6} {'Free':>6}")
    print("="*55)
    grand_total = grand_free = 0
    for label, r in results.items():
        print(f"{label:<8} {r['total']:>7} {r['paid']:>6} {r['free']:>6}")
        grand_total += r["total"]
        grand_free  += r["free"]
    print("="*55)
    print(f"{'TOTAL':<8} {grand_total:>7} {'':>6} {grand_free:>6}")

if __name__ == "__main__":
    main()
