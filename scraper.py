"""
Phase 1 - Collect RAW Google Maps data (free, no API key).

Uses a real Chromium browser via Playwright. For each (category x city) it
searches Google Maps, scrolls the full result list, opens every listing and
records what Google shows. NOTHING is filtered here - every listing found is
saved to output/Raw_Data.xlsx. Missing fields stay blank (nothing is guessed).

Progress is saved after every listing to output/raw_progress.csv, so if the
script stops you can run it again and it resumes where it left off.

Usage:
    python scraper.py              # uses config.json
    python scraper.py my.json      # custom config
"""
import csv
import json
import os
import random
import re
import sys
import time
from datetime import datetime
from urllib.parse import quote_plus

import pandas as pd
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "output")
PROGRESS_CSV = os.path.join(OUT_DIR, "raw_progress.csv")
RAW_XLSX = os.path.join(OUT_DIR, "Raw_Data.xlsx")

RAW_COLUMNS = [
    "Business Name", "Phone Number", "Website URL", "Google Maps Listing URL",
    "Business Category", "City", "State / Province", "Country", "Full Address",
    "Google Rating", "Total Google Reviews", "Business Hours", "Business Status",
    "Additional Services", "Service Areas",
    # traceability columns (not filters)
    "Search Category", "Search City", "Search State", "Search Country",
    "Search Query", "Scraped At",
]

# JavaScript run on a place page. Reads only what is on screen; blanks if absent.
EXTRACT_JS = r"""
() => {
  const q = (s, r=document) => r.querySelector(s);
  const txt = el => el ? el.textContent.replace(/\s+/g,' ').trim() : '';
  const lab = el => el ? (el.getAttribute('aria-label') || '').replace(/\s+/g,' ').trim() : '';
  const main = q('div[role="main"]') || document.body;

  const name = txt(q('h1.DUwDvf', main)) || txt(q('h1', main));
  const category = txt(q('button.DkEaL', main));

  const address = lab(q('button[data-item-id="address"]', main)).replace(/^Address:\s*/i,'');
  let phone = lab(q('button[data-item-id^="phone:tel:"]', main)).replace(/^Phone:\s*/i,'');
  if (!phone) {
    const pb = q('[data-item-id^="phone:tel:"]', main);
    if (pb) phone = pb.getAttribute('data-item-id').replace('phone:tel:','');
  }
  const web = q('a[data-item-id="authority"]', main);
  const website = web ? (web.getAttribute('href') || '') : '';

  let rating = '', reviews = '';
  const f7 = q('div.F7nice', main);
  if (f7) {
    const r = q('span[aria-hidden="true"]', f7);
    rating = txt(r);
    const rv = q('span[aria-label*="review" i]', f7);
    if (rv) { const m = lab(rv).match(/([\d,.]+)/); if (m) reviews = m[1]; }
    if (!reviews) { const m = txt(f7).match(/\(([\d,.]+)\)/); if (m) reviews = m[1]; }
  }

  // Hours: aria-label of hours block lists the whole week
  let hours = '';
  const hEl = q('div.t39EBf[aria-label]', main) || q('[aria-label*="Hide open hours" i]', main);
  if (hEl) hours = lab(hEl).replace(/\.?\s*Hide open hours for the week\.?/i,'').replace(/\.?\s*Suggest new hours\.?/i,'');
  if (!hours) {
    const rows = main.querySelectorAll('table.eK4R0e tr, table tr.y0skZc');
    hours = Array.from(rows).map(tr => Array.from(tr.querySelectorAll('td')).map(td => txt(td)).join(' ')).filter(Boolean).join('; ');
  }

  // Status
  const mainText = main.innerText || '';
  let status = '';
  if (/Permanently closed/i.test(mainText)) status = 'Permanently closed';
  else if (/Temporarily closed/i.test(mainText)) status = 'Temporarily closed';
  else {
    const s = q('.ZDu9vd', main) || q('span.o0Svhf', main);
    const t = txt(s);
    if (/^Open/i.test(t)) status = 'Open';
    else if (/^Closed/i.test(t)) status = 'Closed now';
    else if (/^Opens/i.test(t)) status = 'Closed now';
  }

  // Service area (service-area businesses)
  let serviceAreas = '';
  const sa = Array.from(main.querySelectorAll('[aria-label]')).map(e => lab(e))
               .find(l => /^Service area/i.test(l) || /^Serves\b/i.test(l));
  if (sa) serviceAreas = sa.replace(/^Service area:?\s*/i,'').replace(/^Serves\s*/i,'');

  return {name, category, address, phone, website, rating, reviews, hours, status, serviceAreas};
}
"""

ABOUT_JS = r"""
() => {
  const main = document.querySelector('div[role="main"]') || document.body;
  const groups = [];
  main.querySelectorAll('div.iP2t7d, div[role="region"] > div').forEach(g => {
    const h = g.querySelector('h2');
    if (!h) return;
    const items = Array.from(g.querySelectorAll('li')).map(li =>
      (li.querySelector('[aria-label]')?.getAttribute('aria-label') || li.textContent || '').replace(/\s+/g,' ').trim()
    ).filter(Boolean);
    if (items.length) groups.push(h.textContent.trim() + ': ' + items.join(', '));
  });
  return Array.from(new Set(groups)).join(' | ');
}
"""

# Only these About-tab groups describe services; others (Accessibility, Amenities...) ignored.
SERVICE_GROUPS = ("service options", "offerings", "services", "highlights", "service area", "planning", "payments")


def load_config(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def sleep(cfg):
    time.sleep(random.uniform(cfg.get("min_delay_seconds", 1.5), cfg.get("max_delay_seconds", 3.5)))


def parse_address(addr):
    """Best-effort split of a Google address. Leaves blanks when unsure."""
    city = state = country = ""
    if not addr:
        return city, state, country
    parts = [p.strip() for p in addr.split(",") if p.strip()]
    if len(parts) >= 2 and not re.search(r"\d", parts[-1]) and len(parts[-1]) > 3:
        country = parts.pop()
    if len(parts) >= 2:
        last = parts[-1]
        m = re.match(r"^([A-Z]{2,3})\s+[A-Z0-9][A-Z0-9 \-]{2,9}$", last)  # "TX 77002", "ON M5V 2T6"
        if m:
            state = m.group(1)
            city = parts[-2] if len(parts) >= 2 else ""
        elif re.fullmatch(r"[A-Z]{2}", last):
            state = last
            city = parts[-2]
        else:
            # e.g. "London SW1A 1AA" or "75001 Paris" - strip postcode tokens
            c = re.sub(r"\b[A-Z]{0,2}\d[\w]*\b", "", last).strip()
            city = c if c and len(parts) >= 2 else ""
    return city, state, country


def handle_consent(page):
    if "consent.google" in page.url:
        for label in ("Accept all", "Reject all", "I agree"):
            try:
                page.get_by_role("button", name=label).first.click(timeout=3000)
                page.wait_for_load_state("domcontentloaded")
                return
            except Exception:
                pass


def check_captcha(page, cfg):
    if "/sorry/" in page.url or "unusual traffic" in (page.content()[:5000]).lower():
        if cfg.get("headless"):
            print("\n[!] Google is showing a CAPTCHA. Set \"headless\": false in config.json and run again.")
            sys.exit(1)
        input("\n[!] Google is showing a CAPTCHA. Solve it in the browser window, then press Enter here...")


def collect_listing_urls(page, query, cfg):
    url = f"https://www.google.com/maps/search/{quote_plus(query)}?hl=en"
    page.goto(url, wait_until="domcontentloaded", timeout=60000)
    handle_consent(page)
    check_captcha(page, cfg)
    max_n = int(cfg.get("max_results_per_search", 120))

    # Single result -> Google opens the place directly
    try:
        page.wait_for_selector('div[role="feed"], h1.DUwDvf', timeout=20000)
    except PWTimeout:
        return []
    if not page.query_selector('div[role="feed"]'):
        return [page.url] if "/maps/place/" in page.url else []

    feed = page.query_selector('div[role="feed"]')
    seen, stale = [], 0
    while True:
        links = page.eval_on_selector_all('div[role="feed"] a.hfpxzc', "els => els.map(e => e.href)")
        new = [l for l in links if l not in seen]
        seen.extend(new)
        if len(seen) >= max_n:
            break
        if page.query_selector('span.HlvSq') or "You've reached the end of the list" in feed.inner_text():
            break
        stale = stale + 1 if not new else 0
        if stale >= 6:
            break
        feed.evaluate("el => el.scrollBy(0, el.scrollHeight)")
        time.sleep(random.uniform(1.2, 2.2))
    return seen[:max_n]


def scrape_place(page, url, cfg):
    page.goto(url if "hl=en" in url else url + ("&" if "?" in url else "?") + "hl=en",
              wait_until="domcontentloaded", timeout=60000)
    check_captcha(page, cfg)
    try:
        page.wait_for_selector("h1", timeout=15000)
    except PWTimeout:
        pass
    time.sleep(1.0)
    data = page.evaluate(EXTRACT_JS)

    extra = ""
    if cfg.get("scrape_about_tab", True):
        try:
            tab = page.query_selector('button[role="tab"][aria-label^="About"]')
            if tab:
                tab.click()
                time.sleep(1.2)
                about = page.evaluate(ABOUT_JS)
                keep = [g for g in about.split(" | ") if g.split(":")[0].strip().lower().startswith(SERVICE_GROUPS)]
                extra = " | ".join(keep)
                if not data["serviceAreas"]:
                    sa = [g for g in about.split(" | ") if g.lower().startswith("service area")]
                    if sa:
                        data["serviceAreas"] = sa[0].split(":", 1)[1].strip()
        except Exception:
            pass
    data["extra"] = extra
    return data


def load_done():
    done = set()
    if os.path.exists(PROGRESS_CSV):
        with open(PROGRESS_CSV, encoding="utf-8-sig", newline="") as f:
            for row in csv.DictReader(f):
                done.add((row["Search Query"], row["Google Maps Listing URL"]))
    return done


def append_row(row):
    new_file = not os.path.exists(PROGRESS_CSV)
    with open(PROGRESS_CSV, "a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=RAW_COLUMNS)
        if new_file:
            w.writeheader()
        w.writerow(row)


def export_raw_xlsx():
    if not os.path.exists(PROGRESS_CSV):
        print("No data collected.")
        return
    df = pd.read_csv(PROGRESS_CSV, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    with pd.ExcelWriter(RAW_XLSX, engine="openpyxl") as xw:
        df.to_excel(xw, index=False, sheet_name="Raw_Data")
        ws = xw.sheets["Raw_Data"]
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = min(60, max(12, max(len(str(c.value or "")) for c in col[:200]) + 2))
            ws.column_dimensions[col[0].column_letter].width = width
    print(f"\n[OK] Raw data: {len(df)} rows -> {RAW_XLSX}")


def main():
    cfg_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "config.json")
    cfg = load_config(cfg_path)
    os.makedirs(OUT_DIR, exist_ok=True)
    done = load_done()

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=bool(cfg.get("headless", False)))
        ctx = browser.new_context(locale="en-US", viewport={"width": 1300, "height": 900})
        page = ctx.new_page()
        try:
            for loc in cfg["cities"]:
                place = ", ".join(x for x in (loc.get("city"), loc.get("state"), loc.get("country")) if x)
                for cat in cfg["categories"]:
                    query = f"{cat} in {place}"
                    print(f"\n=== {query} ===")
                    try:
                        urls = collect_listing_urls(page, query, cfg)
                    except Exception as e:
                        print(f"  search failed: {e}")
                        continue
                    print(f"  {len(urls)} listings found")
                    for i, url in enumerate(urls, 1):
                        if (query, url) in done:
                            continue
                        try:
                            d = scrape_place(page, url, cfg)
                        except Exception as e:
                            print(f"  [{i}] failed: {e}")
                            continue
                        city, state, country = parse_address(d["address"])
                        row = {
                            "Business Name": d["name"], "Phone Number": d["phone"],
                            "Website URL": d["website"], "Google Maps Listing URL": url,
                            "Business Category": d["category"], "City": city,
                            "State / Province": state, "Country": country,
                            "Full Address": d["address"], "Google Rating": d["rating"],
                            "Total Google Reviews": d["reviews"], "Business Hours": d["hours"],
                            "Business Status": d["status"], "Additional Services": d["extra"],
                            "Service Areas": d["serviceAreas"],
                            "Search Category": cat, "Search City": loc.get("city", ""),
                            "Search State": loc.get("state", ""), "Search Country": loc.get("country", ""),
                            "Search Query": query, "Scraped At": datetime.now().isoformat(timespec="seconds"),
                        }
                        append_row(row)
                        done.add((query, url))
                        print(f"  [{i}/{len(urls)}] {d['name']} | {d['phone'] or '-'} | "
                              f"{'web' if d['website'] else 'NO WEB'} | {d['reviews'] or 0} reviews")
                        sleep(cfg)
        except KeyboardInterrupt:
            print("\nStopped by user - progress saved.")
        finally:
            browser.close()
    export_raw_xlsx()


if __name__ == "__main__":
    main()
