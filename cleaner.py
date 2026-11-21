"""
Phases 2-5 - Clean and filter Raw_Data.xlsx -> Clean_Qualified_Leads.xlsx

Raw_Data.xlsx is only READ, never changed.
Also writes Cleaning_Report.xlsx: every removed row and the reason, so you can
check that nothing good was dropped.

Usage:
    python cleaner.py                       # output/Raw_Data.xlsx
    python cleaner.py path/to/Raw_Data.xlsx
"""
import json
import os
import re
import sys

import pandas as pd
import phonenumbers
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(HERE, "output")

FINAL_COLUMNS = [
    "Business Name", "Phone Number", "Business Category", "City", "State / Province",
    "Rating", "Review Count", "Services Offered", "Service Areas", "Address",
    "Website Status", "Google Maps Link", "Lead Status", "Notes",
]

# Local service industries (matched against Google category, case-insensitive)
SERVICE_KEYWORDS = [
    "repair", "plumb", "electric", "hvac", "heating", "air condition", "furnace", "roof",
    "contractor", "construction", "remodel", "renovat", "handyman", "landscap", "lawn",
    "tree service", "arborist", "pest", "exterminat", "termite", "cleaning", "cleaner",
    "carpet", "janitorial", "painter", "painting", "floor", "tile", "concrete", "mason",
    "fence", "deck", "garage door", "locksmith", "appliance", "pool", "gutter", "siding",
    "window", "glass", "glazier", "drywall", "insulation", "septic", "drain", "sewer",
    "water heater", "solar", "install", "moving", "mover", "junk", "pressure wash",
    "power wash", "auto", "mechanic", "body shop", "collision", "towing", "tire", "brake",
    "transmission", "welding", "cabinet", "countertop", "kitchen", "bathroom", "foundation",
    "waterproof", "restoration", "mold", "chimney", "irrigation", "sprinkler", "paving",
    "asphalt", "excavat", "demolition", "carpent", "upholster", "computer", "maintenance",
    "service", "builder", "well drilling", "generator", "security system", "alarm",
    "door", "blind", "shutter", "awning", "stucco", "plaster", "epoxy", "waterproofing",
    "duct", "dryer vent", "gas", "boiler", "refrigeration", "elevator", "pump", "engineer",
]

# Known national chains / franchises (name contains, case-insensitive). Add more as you see them.
CHAIN_NAMES = [
    "home depot", "lowe's", "lowes", "roto-rooter", "roto rooter", "mr. rooter", "mr rooter",
    "mr. handyman", "mr handyman", "mr. electric", "mr electric", "mr. appliance", "mr appliance",
    "mister sparky", "one hour heating", "benjamin franklin plumbing", "ars rescue rooter",
    "rescue rooter", "sears", "servpro", "servicemaster", "stanley steemer", "chem-dry",
    "chem dry", "terminix", "orkin", "trugreen", "molly maid", "merry maids", "the maids",
    "two men and a truck", "ace hardware", "ace handyman", "precision garage door",
    "aire serv", "glass doctor", "window genie", "dryer vent wizard", "mosquito joe",
    "puroclean", "paul davis", "belfor", "1-800-got-junk", "got junk", "college hunks",
    "budget blinds", "bath fitter", "re-bath", "kitchen tune-up", "five star painting",
    "certapro", "jiffy lube", "midas", "meineke", "firestone", "pep boys", "safelite",
    "monro", "christian brothers automotive", "valvoline", "goodyear", "precision tune",
    "maaco", "grease monkey", "take 5", "brakes plus", "ntb", "discount tire", "les schwab",
    "big o tires", "aamco", "cottman", "lawn doctor", "weed man", "handyman connection",
    "rooter-man", "rooterman", "dreammaker", "renewal by andersen", "window world",
    "leaffilter", "leaf filter", "aptive", "rentokil", "abc home", "neighborly", "anago",
    "jan-pro", "coverall", "u-haul", "sherwin-williams", "best buy", "geek squad", "ubreakifix",
    "asurion", "cpr cell phone", "batteries plus", "caliber collision", "gerber collision",
    "service experts", "ars ", "one hour air", "frontdoor", "american home shield",
    "sunrun", "tesla", "vivint", "adt", "brinks", "pod", "pods", "public storage",
]

HIGH_DEMAND = ["plumb", "hvac", "heating", "air condition", "electric", "roof", "garage door",
               "water heater", "drain", "sewer", "appliance", "pest", "foundation",
               "restoration", "septic", "locksmith", "tree service", "remodel"]

COUNTRY_MAP = {"usa": "United States", "us": "United States", "u.s.": "United States",
               "u.s.a.": "United States", "united states of america": "United States",
               "uk": "United Kingdom", "u.k.": "United Kingdom", "england": "United Kingdom",
               "uae": "United Arab Emirates"}


def blank(v):
    return v is None or (isinstance(v, float) and pd.isna(v)) or str(v).strip() == ""


def clean_text(v):
    return "" if blank(v) else re.sub(r"\s+", " ", str(v)).strip()


def clean_name(n):
    n = clean_text(n)
    if n.isupper() and len(n) > 4:      # "ACME PLUMBING LLC" -> "Acme Plumbing LLC"
        n = " ".join(w if w in {"LLC", "INC", "LLP", "HVAC", "USA", "AC", "II", "III"} else w.capitalize()
                     for w in n.split())
    return n.strip(" -|,")


def std_phone(p, region):
    p = clean_text(p)
    if not p:
        return ""
    try:
        num = phonenumbers.parse(p, region)
        if not phonenumbers.is_valid_number(num):
            return ""
        if phonenumbers.region_code_for_number(num) in ("US", "CA") and region in ("US", "CA"):
            n = phonenumbers.national_significant_number(num)
            return f"({n[:3]}) {n[3:6]}-{n[6:]}"
        return phonenumbers.format_number(num, phonenumbers.PhoneNumberFormat.INTERNATIONAL)
    except phonenumbers.NumberParseException:
        return ""


def std_city(c):
    c = clean_text(c)
    return c.title() if c and (c.isupper() or c.islower()) else c


def std_state(s):
    s = clean_text(s)
    return s.upper() if len(s) <= 3 else s.title() if (s.isupper() or s.islower()) else s


def std_country(c):
    c = clean_text(c)
    return COUNTRY_MAP.get(c.lower(), c)


def to_int(v):
    s = re.sub(r"[^\d]", "", clean_text(v))
    return int(s) if s else None


def to_float(v):
    try:
        return float(clean_text(v).replace(",", "."))
    except ValueError:
        return None


def place_key(url):
    url = clean_text(url)
    m = re.search(r"!1s(0x[0-9a-f]+:0x[0-9a-f]+)", url) or re.search(r"[?&]cid=(\d+)", url)
    if m:
        return m.group(1)
    m = re.search(r"/maps/place/([^/]+)/@", url)
    return m.group(1).lower() if m else url.split("?")[0]


def is_chain(name):
    n = " " + name.lower() + " "
    return any(f" {c}" in n or n.strip().startswith(c) for c in CHAIN_NAMES)


def priority(row):
    """Score only from collected data. Higher = call first."""
    score, why = 0, []
    services = [s for s in re.split(r"[,|]", row["Services Offered"]) if s.strip()]
    areas = [a for a in re.split(r"[,;|]", row["Service Areas"]) if a.strip()]
    if len(services) >= 3:
        score += 2; why.append("multiple services")
    if len(areas) >= 2:
        score += 2; why.append("multiple service areas")
    blob = (row["Services Offered"] + " " + row["Business Category"] + " " + row["_name"]).lower()
    if "residential" in blob or "commercial" in blob:
        score += 1; why.append("residential/commercial")
    if any(k in row["Business Category"].lower() for k in HIGH_DEMAND):
        score += 2; why.append("high-demand service")
    if (row["Review Count"] or 0) >= 50:
        score += 1; why.append("50+ reviews")
    return score, why


def main():
    raw_path = sys.argv[1] if len(sys.argv) > 1 else os.path.join(OUT_DIR, "Raw_Data.xlsx")
    cfg_path = os.path.join(HERE, "config.json")
    cfg = json.load(open(cfg_path, encoding="utf-8")) if os.path.exists(cfg_path) else {}
    region = cfg.get("phone_region", "US")
    min_reviews = int(cfg.get("min_reviews", 10))
    chain_thr = int(cfg.get("chain_same_name_city_threshold", 3))

    raw = pd.read_excel(raw_path, dtype=str).fillna("")   # read-only copy
    print(f"Raw rows: {len(raw)}")
    df = raw.copy()
    df["_raw_row"] = df.index + 2   # Excel row number in Raw_Data.xlsx

    df["_name"] = df["Business Name"].map(clean_name)
    df["_phone"] = df["Phone Number"].map(lambda p: std_phone(p, region))
    df["_reviews"] = df["Total Google Reviews"].map(to_int)
    df["_rating"] = df["Google Rating"].map(to_float)
    df["_key"] = df["Google Maps Listing URL"].map(place_key)
    df["_cat"] = df["Business Category"].map(clean_text)

    # Same business name seen in many different cities -> likely chain/franchise
    city_col = df["City"].where(df["City"].str.strip() != "", df.get("Search City", ""))
    name_cities = df.assign(_c=city_col.str.lower()).groupby(df["_name"].str.lower())["_c"].nunique()
    multi_city = set(name_cities[name_cities >= chain_thr].index)

    removed = []

    def drop(mask, reason):
        nonlocal df
        r = df[mask].copy()
        r["Removal Reason"] = reason
        removed.append(r)
        df = df[~mask]

    drop(df["_name"] == "", "Missing business name")
    drop(df["_key"].duplicated(keep="first"), "Duplicate listing (same Google Maps place)")
    drop(df["Website URL"].str.strip() != "", "Has a website")
    drop(df["_phone"] == "", "Missing or invalid phone number")
    drop(df["_reviews"].isna() | (df["_reviews"] < min_reviews), f"Fewer than {min_reviews} reviews")
    drop(df["Business Status"].str.contains("permanently closed", case=False), "Permanently closed")
    drop(df["Business Status"].str.contains("temporarily closed", case=False), "Temporarily closed (not active)")
    drop(df["_name"].map(is_chain), "Known chain / franchise / national company")
    drop(df["_name"].str.lower().isin(multi_city), f"Same name in {chain_thr}+ cities (likely chain/franchise)")
    drop(~df["_cat"].str.lower().map(lambda c: any(k in c for k in SERVICE_KEYWORDS)),
         "Category not a local service industry (or category missing)")

    # Same phone = same business: keep the listing with most reviews
    df = df.sort_values("_reviews", ascending=False)
    drop(df["_phone"].duplicated(keep="first"), "Duplicate phone number (same business)")

    out = pd.DataFrame({
        "Business Name": df["_name"],
        "Phone Number": df["_phone"],
        "Business Category": df["_cat"],
        "City": df["City"].map(std_city),
        "State / Province": df["State / Province"].map(std_state),
        "Rating": df["_rating"],
        "Review Count": df["_reviews"].astype("Int64"),
        "Services Offered": df["Additional Services"].map(clean_text),
        "Service Areas": df["Service Areas"].map(clean_text),
        "Address": df["Full Address"].map(clean_text),
        "Website Status": "No Website",
        "Google Maps Link": df["Google Maps Listing URL"].map(clean_text),
        "Lead Status": "New",
        "_name": df["_name"],
        "_query": df.get("Search Query", ""),
    })
    pr = out.apply(priority, axis=1)
    out["_score"] = [s for s, _ in pr]
    out["Notes"] = [
        ("Priority: " + ("High" if s >= 4 else "Medium" if s >= 2 else "Normal")
         + (" (" + ", ".join(w) + ")" if w else "")
         + (f". Found via search: {q}" if q else ""))
        for (s, w), q in zip(pr, out["_query"])
    ]
    out = out.sort_values(["_score", "Review Count"], ascending=[False, False])
    out = out[FINAL_COLUMNS]

    clean_path = os.path.join(OUT_DIR, "Clean_Qualified_Leads.xlsx")
    write_clean(out, clean_path)

    rep = pd.concat(removed) if removed else pd.DataFrame()
    rep_path = os.path.join(OUT_DIR, "Cleaning_Report.xlsx")
    summary = (rep["Removal Reason"].value_counts().rename_axis("Reason").reset_index(name="Rows removed")
               if len(rep) else pd.DataFrame(columns=["Reason", "Rows removed"]))
    with pd.ExcelWriter(rep_path, engine="openpyxl") as xw:
        pd.concat([summary, pd.DataFrame([{"Reason": "QUALIFIED LEADS KEPT", "Rows removed": len(out)},
                                          {"Reason": "TOTAL RAW ROWS", "Rows removed": len(raw)}])]
                  ).to_excel(xw, index=False, sheet_name="Summary")
        if len(rep):
            cols = ["Removal Reason", "_raw_row"] + [c for c in raw.columns]
            rep[cols].rename(columns={"_raw_row": "Raw_Data Row"}).to_excel(xw, index=False, sheet_name="Removed_Rows")

    print(summary.to_string(index=False))
    print(f"\n[OK] Qualified leads: {len(out)} -> {clean_path}")
    print(f"[OK] Removal log -> {rep_path}")


def write_clean(out, path):
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        out.to_excel(xw, index=False, sheet_name="Qualified_Leads")
        ws = xw.sheets["Qualified_Leads"]
        ws.freeze_panes = "B2"
        ws.auto_filter.ref = ws.dimensions
        head = PatternFill("solid", fgColor="1F4E78")
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head
            c.alignment = Alignment(vertical="center", wrap_text=True)
        widths = {"A": 32, "B": 17, "C": 24, "D": 16, "E": 10, "F": 8, "G": 9, "H": 45,
                  "I": 30, "J": 40, "K": 14, "L": 40, "M": 15, "N": 50}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w
        n = ws.max_row
        for r in range(2, n + 1):
            link = ws[f"L{r}"]
            if link.value:
                link.hyperlink = link.value
                link.style = "Hyperlink"
        if n >= 2:
            dv = DataValidation(type="list", formula1='"New,Called,Interested,Not Interested"', allow_blank=True)
            ws.add_data_validation(dv)
            dv.add(f"M2:M{n}")


if __name__ == "__main__":
    main()
