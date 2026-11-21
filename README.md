# Google Maps Lead Scraper (Free)

Finds local service businesses on Google Maps that have **no website**, a valid phone number and **10+ reviews**. The result is a list of leads for cold calling about website development, website redesign and local SEO.

It is 100% free: no API key, no paid tool. It runs a real Chrome browser on **your own computer**.

## Setup (one time)
1. Install **Python 3.10+** from python.org. On Windows, tick "Add Python to PATH" during install.
2. Download this folder.

## Usage
1. Open `config.json` and set your target **cities** and **categories**:
   ```json
   "cities": [ {"city": "Houston", "state": "TX", "country": "United States"} ],
   "categories": [ "plumber", "electrician", "roofing contractor" ]
   ```
   Every category is searched in every city. For another country, change `phone_region`, e.g. `"GB"`, `"CA"` or `"AU"`.
2. Run it:
   - **Windows:** double-click `run_windows.bat`
   - **Mac/Linux:** `./run_mac_linux.sh`

   The first run takes a few minutes to install everything.
3. A Chrome window opens and collects the data. If Google shows a CAPTCHA, solve it in the browser and press Enter in the black window.
4. When it finishes, the `output/` folder holds:

| File | What it contains |
|---|---|
| `Raw_Data.xlsx` | **Every** listing collected, with no filtering (Phase 1) |
| `Clean_Qualified_Leads.xlsx` | Only qualified leads, ready for cold calling (Phases 2–5) |
| `Cleaning_Report.xlsx` | Every row that was removed and why, so you can verify the filtering |
| `raw_progress.csv` | Saved after every listing. If the run stops, run it again and it resumes from there |

To re-clean only (no new scraping), run: `python cleaner.py`

## Cleaning rules (cleaner.py)
A lead is removed if any of these is true:
- It is a duplicate listing (same Google place).
- It has a website.
- Its phone number is missing or invalid.
- It has fewer than 10 reviews (`min_reviews` in config).
- It is permanently or temporarily closed.
- It is a known chain, franchise or national brand (`CHAIN_NAMES` list), or the same name appears in 3+ cities.
- Its category is not a repair, maintenance, installation, contracting or other local service (`SERVICE_KEYWORDS` list).
- It shares a phone number with another lead (the listing with the most reviews is kept).

Standardization: phones become `(713) 555-0101` for US/CA and international format elsewhere. City, state and country names and business names are cleaned up.

Sorting: leads are sorted by priority: multiple services, multiple service areas, residential/commercial, high-demand categories and 50+ reviews. The reason is written in the Notes column.

Lead Status has a dropdown: New / Called / Interested / Not Interested.

## Important
- **Nothing is invented.** If Google does not show a field, it stays blank.
- "Services Offered" and "Service Areas" are only filled when Google shows them (About tab / service area). Many listings will leave them blank.
- Chain detection works from a name list, so it is not perfect. Before calling, open a lead's Google Maps link to check it.
- Google sometimes changes its page layout. If a field comes back empty for every listing, the selectors in `scraper.py` need updating.
- Use slow delays (the defaults are fine). Scraping too fast gets you blocked by Google. Make sure your use follows Google's Terms of Service and your local cold-calling laws (e.g. TCPA/DNC in the US).
