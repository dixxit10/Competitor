#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_data.py — competitor_scoring_output.xlsx  ->  data.js   (one-click)

Usage
-----
    python build_data.py                       # reads ./competitor_scoring_output.xlsx
    python build_data.py --input other.xlsx    # explicit workbook
    python build_data.py --out data.js         # explicit output path
    python build_data.py --mock-min 20         # pad to 20 stores when real < 20 (default)
    python build_data.py --no-mock             # never generate mock stores
    python build_data.py --force-mock 50       # always pad to 50 stores
    python build_data.py --validate-only       # check the contract, write nothing

What it does
------------
1. Validates the workbook against the FROZEN data contract (v1). A missing or
   renamed column is a HARD ERROR - the script never emits a half-broken file.
2. Normalises the store score to 0-100 on a FIXED base (400), keeping the raw
   value alongside it so it stays traceable.
3. Decides AUTOMATICALLY whether mock stores are needed: real >= mock-min means
   NO mock at all; below that it pads up to mock-min so the layout can still be
   judged. Mock rows carry mock:true, and the front end hides every mock notice
   when none exist.
4. Writes window.SHARETEA_DATA = {...} to data.js.
5. Prints a summary: stores read, competitor rows, mock count, output size.

No external API is called. Deterministic: same workbook -> identical output.
"""

import argparse
import datetime
import json
import math
import os
import random
import sys

try:
    import openpyxl
except ImportError:
    sys.exit("ERROR: openpyxl is required.  Install it with:  pip install openpyxl")


# ============================================================================
# FROZEN CONSTANTS - must match the front end (index.html)
# ============================================================================
STORE_REF = 400      # fixed normalisation base for the store score
COMP_REF = 30        # fixed reference for the competitor score bar
THREAT_MIN = 10      # direct competitor influence_score that counts as a threat
INDIRECT_W = 0.5     # pipeline weight applied to indirect competitors

# tier -> (radius_mi, layer, density band, income band, competitor-density factor)
TIERS = [
    ("Core Urban",                   0.5, "L1", (20000, 60000), (90000, 260000), 1.00),
    ("Dense Urban",                  1.0, "L2", (5000, 20000),  (55000, 240000), 0.85),
    ("Urban",                        2.0, "L3", (3000, 5000),   (48000, 130000), 0.55),
    ("Suburban / Mixed Residential", 4.0, "L4", (1000, 3000),   (60000, 150000), 0.30),
    ("Sparse / Rural",               8.0, "L5", (100, 1000),    (45000, 110000), 0.10),
]
TIER_BY_NAME = {t[0]: t for t in TIERS}

# The pipeline emits area_type in Chinese. The front end is English-only, so the
# translation lives here - in the data layer - not in the browser.
AREA_TYPE_EN = {
    "學校": "School",
    "醫院": "Hospital",
    "公家機關": "Government office",
    "交通樞紐": "Transit hub",
    "購物中心": "Shopping mall",
    "大型量販": "Big-box retail",
    "加油站/主幹道商圈": "Gas station / arterial corridor",
    "一般零售街區": "General retail street",
}

# ============================================================================
# DATA CONTRACT v1 (frozen) - required columns per sheet
# ============================================================================
SHEET_STORES = "門市明細（依壓力分排序）"
SHEET_TIERS = "分層彙總"
SHEET_COMPS = "競業店名明細"

CONTRACT = {
    SHEET_STORES: [
        "store_id", "store_name", "region_code", "supervisor", "address",
        "lat", "lng", "census_tract_geoid", "population", "land_area_sqmi",
        "density_per_sqmi", "tier", "radius_mi", "area_type",
        "median_household_income", "median_age",
        "pct_white", "pct_black", "pct_asian", "pct_hispanic",
        "pct_under_18", "pct_18_34", "pct_35_54", "pct_55_plus",
        "direct_competitor_count", "direct_competitor_avg_rating",
        "direct_competitor_total_reviews", "direct_competitor_threat_score",
        "indirect_competitor_count", "indirect_competitor_avg_rating",
        "indirect_competitor_total_reviews", "indirect_competitor_threat_score",
        "total_competitive_pressure_score",
        "own_rating", "own_review_count", "own_matched_google_name",
    ],
    SHEET_TIERS: [
        "tier", "store_count", "avg_direct_competitors", "avg_indirect_competitors",
        "avg_direct_threat_score", "avg_indirect_threat_score",
        "avg_total_pressure_score", "avg_own_rating", "avg_median_income",
    ],
    SHEET_COMPS: [
        "store_id", "store_name", "competitor_type", "competitor_name",
        "google_type", "rating", "review_count", "distance_mi", "influence_score",
    ],
}

# Columns the front end actually renders. Everything else is carried for
# traceability but is not required to be displayed.
FRONTEND_DISPLAY = {
    "store": ["store_id", "store_name", "region_code", "supervisor", "address",
              "tier", "radius_mi", "area_type", "median_household_income",
              "median_age", "pct_white", "pct_black", "pct_asian", "pct_hispanic",
              "pct_under_18", "pct_18_34", "pct_35_54", "pct_55_plus",
              "total_competitive_pressure_score", "own_rating", "own_review_count",
              "own_matched_google_name"],
    "competitor": ["competitor_name", "google_type", "rating", "review_count",
                   "distance_mi", "influence_score", "competitor_type"],
}


# ============================================================================
# VALIDATION
# ============================================================================
class ContractError(Exception):
    pass


def _header_of(ws):
    return [c.value for c in ws[1]]


def validate(wb):
    """Check the workbook against the frozen contract. Raises ContractError."""
    problems = []

    for sheet, required in CONTRACT.items():
        if sheet not in wb.sheetnames:
            problems.append("missing sheet: %r" % sheet)
            continue
        have = _header_of(wb[sheet])
        missing = [c for c in required if c not in have]
        if missing:
            problems.append("sheet %r is missing %d column(s): %s"
                            % (sheet, len(missing), ", ".join(missing)))
        extra = [c for c in have if c and c not in required]
        if extra:
            # Extra columns are tolerated (the contract is a minimum), but the
            # user should know the workbook has drifted from what was frozen.
            print("  note: sheet %r has %d extra column(s) not in the contract: %s"
                  % (sheet, len(extra), ", ".join(str(e) for e in extra)))

    if problems:
        raise ContractError(
            "The workbook does not match the frozen data contract (v1):\n  - "
            + "\n  - ".join(problems)
            + "\n\nNothing was written. Fix the workbook (or update CONTRACT in "
              "build_data.py) and run again."
        )

    # tier values must be known, otherwise the Layer mapping is undefined
    ws = wb[SHEET_STORES]
    hdr = _header_of(ws)
    ti = hdr.index("tier")
    unknown = sorted({r[ti] for r in ws.iter_rows(min_row=2, values_only=True)
                      if r[ti] is not None and r[ti] not in TIER_BY_NAME})
    if unknown:
        raise ContractError(
            "Unknown tier value(s) in %r: %s\n"
            "The Layer mapping is defined in TIERS. Add the new tier there "
            "(with its radius_mi and layer code) before running again."
            % (SHEET_STORES, ", ".join(unknown))
        )

    # competitor_type must be one of the two known labels
    ws2 = wb[SHEET_COMPS]
    h2 = _header_of(ws2)
    ci = h2.index("competitor_type")
    bad = sorted({r[ci] for r in ws2.iter_rows(min_row=2, values_only=True)
                  if r[ci] not in ("直接", "間接")})
    if bad:
        raise ContractError(
            "Unexpected competitor_type value(s) in %r: %s (expected 直接 / 間接)"
            % (SHEET_COMPS, ", ".join(str(b) for b in bad))
        )


# ============================================================================
# READ
# ============================================================================
def _num(v, default=0.0):
    return default if v is None else float(v)


def _int(v, default=0):
    return default if v is None else int(v)


def translate_area_type(raw):
    """'交通樞紐, 學校' -> 'Transit hub, School'. Unknown tokens pass through."""
    if not raw:
        return "General retail street"
    parts = [p.strip() for p in str(raw).split(",") if p.strip()]
    return ", ".join(AREA_TYPE_EN.get(p, p) for p in parts)


def read_real_stores(wb):
    ws = wb[SHEET_STORES]
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[0])
    ix = {h: i for i, h in enumerate(hdr)}

    stores = []
    for r in rows[1:]:
        if r[ix["store_id"]] is None:
            continue
        g = lambda k: r[ix[k]]
        tier = g("tier")
        stores.append({
            "id": g("store_id"),
            "name": g("store_name"),
            "region": g("region_code"),
            "sup": g("supervisor"),
            "addr": g("address"),
            "tier": tier,
            "layer": TIER_BY_NAME[tier][2],
            "radius": _num(g("radius_mi")),
            "areaType": translate_area_type(g("area_type")),
            "lat": _num(g("lat")),
            "lng": _num(g("lng")),
            "pop": _int(g("population")),
            "dens": _num(g("density_per_sqmi")),
            "income": _int(g("median_household_income")),
            "mage": _num(g("median_age")),
            "eth": [_num(g("pct_white")), _num(g("pct_black")),
                    _num(g("pct_asian")), _num(g("pct_hispanic")), 0.0],
            "age": [_num(g("pct_under_18")), _num(g("pct_18_34")),
                    _num(g("pct_35_54")), _num(g("pct_55_plus"))],
            "own": {"rating": _num(g("own_rating")),
                    "reviews": _int(g("own_review_count")),
                    "name": g("own_matched_google_name") or "Sharetea"},
            "rep": {"d": _int(g("direct_competitor_count")),
                    "i": _int(g("indirect_competitor_count")),
                    "dt": _num(g("direct_competitor_threat_score")),
                    "it": _num(g("indirect_competitor_threat_score")),
                    "tot": _num(g("total_competitive_pressure_score")),
                    "dar": _num(g("direct_competitor_avg_rating")),
                    "dtr": _int(g("direct_competitor_total_reviews")),
                    "iar": _num(g("indirect_competitor_avg_rating")),
                    "itr": _int(g("indirect_competitor_total_reviews"))},
            "comps": [],
            "mock": False,
        })

    # "Other" is whatever the four reported groups do not cover
    for s in stores:
        s["eth"][4] = round(100.0 - sum(s["eth"][:4]), 1)

    # competitor rows
    ws2 = wb[SHEET_COMPS]
    rows2 = list(ws2.iter_rows(values_only=True))
    h2 = list(rows2[0])
    i2 = {h: i for i, h in enumerate(h2)}
    by_id = {s["id"]: s for s in stores}
    orphan = 0
    for r in rows2[1:]:
        sid = r[i2["store_id"]]
        if sid not in by_id:
            orphan += 1
            continue
        by_id[sid]["comps"].append({
            "n": r[i2["competitor_name"]],
            "t": r[i2["google_type"]],
            "r": _num(r[i2["rating"]]),
            "rv": _int(r[i2["review_count"]]),
            "d": _num(r[i2["distance_mi"]]),
            "s": _num(r[i2["influence_score"]]),
            "c": "d" if r[i2["competitor_type"]] == "直接" else "i",
        })
    if orphan:
        print("  note: %d competitor row(s) referenced a store_id not in the "
              "store sheet and were skipped" % orphan)

    return stores


def read_tier_summary(wb):
    ws = wb[SHEET_TIERS]
    rows = list(ws.iter_rows(values_only=True))
    hdr = list(rows[0])
    out = []
    for r in rows[1:]:
        if r[0] is None:
            continue
        out.append({h: (None if v is None else v) for h, v in zip(hdr, r)})
    return out


# ============================================================================
# MOCK (only generated when the real roster is too small to judge the layout)
# ============================================================================
POOL = [
    ("Teaspoon", "tea_house", 4.6, 180, True),
    ("Gong Cha", "tea_house", 4.1, 320, True),
    ("Happy Lemon", "tea_house", 4.3, 240, True),
    ("Chatime", "tea_house", 4.2, 410, True),
    ("CoCo Fresh Tea & Juice", "tea_house", 4.0, 290, True),
    ("Sharetea", "tea_house", 4.4, 350, True),
    ("Kung Fu Tea", "tea_house", 4.2, 520, True),
    ("Tiger Sugar", "tea_house", 4.3, 260, True),
    ("Yi Fang Taiwan Fruit Tea", "tea_house", 4.5, 190, True),
    ("Xing Fu Tang", "tea_house", 4.4, 210, True),
    ("Boba Guys", "tea_house", 4.3, 900, True),
    ("Plentea", "tea_house", 4.4, 480, True),
    ("Feng Cha Teahouse", "tea_house", 4.5, 300, True),
    ("Machi Machi", "tea_house", 4.6, 150, True),
    ("TP Tea", "tea_house", 4.4, 130, True),
    ("Chicha San Chen", "tea_house", 4.7, 220, True),
    ("T4 Tea For U", "hawaiian_restaurant", 4.3, 300, True),
    ("Bliss Sandwiches & Boba Cafe", "restaurant", 4.4, 90, True),
    ("Asian Boba Bistro", "asian_restaurant", 4.5, 60, True),
    ("Metro Hong Kong Dessert and Milk Tea", "chinese_restaurant", 4.2, 280, True),
    ("Hanabi Sushi & Boba", "meal_takeaway", 4.6, 110, True),
    ("Starbucks Coffee Company", "coffee_shop", 4.0, 700, False),
    ("Peet's Coffee", "coffee_shop", 4.2, 480, False),
    ("Blue Bottle Coffee", "coffee_shop", 4.1, 420, False),
    ("Dutch Bros Coffee", "coffee_shop", 4.3, 600, False),
    ("Sightglass Coffee", "coffee_shop", 4.4, 1200, False),
    ("Philz Coffee", "coffee_shop", 4.5, 900, False),
    ("The Coffee Bean & Tea Leaf", "coffee_shop", 4.1, 380, False),
    ("Dunkin'", "coffee_shop", 3.9, 450, False),
    ("A'Roma Roasters Coffee & Tea", "coffee_shop", 4.5, 700, False),
    ("Panera Bread", "cafe", 4.1, 400, False),
    ("Chai & Co.", "cafe", 4.5, 160, False),
    ("Junbi Matcha & Tea", "cafe", 4.7, 230, False),
    ("Quickly", "cafe", 4.0, 280, False),
    ("Brew", "cafe", 4.5, 600, False),
    ("85C Bakery Cafe", "bakery", 4.3, 340, False),
    ("Arsicault Bakery", "bakery", 4.6, 800, False),
    ("Paris Baguette", "bakery", 4.2, 520, False),
    ("Tous Les Jours", "bakery", 4.3, 380, False),
    ("Golden Gate Fortune Cookie Factory", "bakery", 4.5, 1500, False),
]

CITIES = [
    "Sunnyvale", "Fremont", "Hayward", "Milpitas", "Cupertino", "Santa Clara",
    "San Mateo", "Redwood City", "Palo Alto", "Mountain View", "Daly City",
    "Union City", "Newark", "San Leandro", "Walnut Creek", "Pleasanton",
    "Livermore", "Dublin", "San Ramon", "Danville", "Berkeley", "Oakland",
    "Alameda", "Richmond", "Vallejo", "Fairfield", "Vacaville", "Napa",
    "Petaluma", "Novato", "San Rafael", "Corte Madera", "Millbrae",
    "Burlingame", "San Bruno", "South San Francisco", "Pacifica", "Half Moon Bay",
    "Los Gatos", "Campbell", "Saratoga", "Morgan Hill", "Gilroy", "Watsonville",
    "Salinas", "Monterey", "Seaside", "Marina", "Stockton", "Tracy", "Manteca",
    "Modesto", "Turlock", "Merced", "Sacramento", "Elk Grove", "Roseville",
    "Folsom", "Rocklin", "Citrus Heights", "Davis", "Woodland", "Yuba City",
    "Chico", "Redding", "Santa Cruz", "Capitola", "Scotts Valley", "Los Altos",
    "Menlo Park", "Belmont", "Foster City", "San Carlos", "Los Banos", "Lodi",
    "Brentwood", "Antioch", "Pittsburg", "Concord", "Clayton", "Martinez",
    "Hercules", "Pinole", "El Cerrito", "Albany", "Emeryville", "Piedmont",
    "Orinda", "Moraga", "Lafayette", "Alamo", "Blackhawk", "Discovery Bay",
    "Oakley", "Ripon", "Escalon", "Lathrop", "Ceres", "Patterson", "Newman",
    "Gustine", "Atwater", "Livingston", "Delhi", "Hilmar", "Winton", "Planada",
    "Le Grand", "Chowchilla", "Madera", "Fresno", "Clovis", "Sanger", "Selma",
    "Kingsburg", "Reedley", "Dinuba", "Tulare", "Visalia", "Porterville",
    "Delano", "Wasco", "Bakersfield", "Tehachapi", "Lancaster", "Palmdale",
    "Santa Clarita", "Pasadena", "Glendale", "Burbank", "Torrance", "Long Beach",
]
SUFFIX = ["Plaza", "Mall", "Town Center", "Marketplace", "Commons", "Square",
          "Crossing", "Village", "Station", "Promenade", "Center", "Court"]
REGIONS = ["NCA", "SCA", "TXN", "TXC", "TXS", "TXW"]
REGION_W = [0.34, 0.30, 0.14, 0.10, 0.08, 0.04]
SUPERVISORS = ["Mills", "Okafor", "Reyes", "Tanaka", "Whitfield", "Delgado",
               "Nguyen", "Brennan", "Castillo", "Hollis"]
AREA_TYPES = ["Transit hub, School", "Retail corridor", "University district",
              "Office park", "Residential, School", "Entertainment district",
              "Suburban strip mall", "Downtown core"]


def influence(rating, reviews, dist, radius):
    """The pipeline's own formula, reproduced so mock rows are self-consistent."""
    return rating * math.log1p(reviews) * max(0.0, 1.0 - dist / radius)


def make_mock_stores(n, seed=20260922):
    rng = random.Random(seed)
    used, stores = set(), []
    tier_plan = (["Core Urban"] * 11 + ["Dense Urban"] * 41 + ["Urban"] * 34 +
                 ["Suburban / Mixed Residential"] * 34 + ["Sparse / Rural"] * 13)
    rng.shuffle(tier_plan)
    # if more mocks are requested than the plan holds, cycle through it
    while len(tier_plan) < n:
        tier_plan += tier_plan[:n - len(tier_plan)]

    for k in range(n):
        tier = tier_plan[k]
        _, radius, layer, dens_band, inc_band, dens_f = TIER_BY_NAME[tier]

        while True:
            nm = "%s %s" % (rng.choice(CITIES), rng.choice(SUFFIX))
            if nm not in used:
                used.add(nm)
                break

        region = rng.choices(REGIONS, weights=REGION_W, k=1)[0]
        dens = round(rng.uniform(*dens_band), 1)
        income = int(round(rng.uniform(*inc_band), -2))
        mage = round(rng.uniform(31.0, 49.0), 1)
        pop = int(rng.uniform(900, 9000))

        w = round(rng.uniform(18, 62), 1)
        b = round(rng.uniform(0.5, 14), 1)
        a = round(rng.uniform(3, 46), 1)
        h = round(rng.uniform(4, 40), 1)
        tot = w + b + a + h
        if tot > 96:
            sc = 96.0 / tot
            w, b, a, h = (round(x * sc, 1) for x in (w, b, a, h))
        eth = [w, b, a, h, round(100.0 - (w + b + a + h), 1)]

        u18 = round(rng.uniform(4, 24), 1)
        y1834 = round(rng.uniform(12, 34), 1)
        y3554 = round(rng.uniform(20, 34), 1)
        age = [u18, y1834, y3554, round(100.0 - u18 - y1834 - y3554, 1)]

        n_direct = max(1, int(round(rng.randint(2, 24) * dens_f)))
        n_indirect = max(2, int(round(rng.randint(4, 20) * dens_f)))
        comps = []
        for want_direct, count in ((True, n_direct), (False, n_indirect)):
            for _ in range(count):
                cnm, ty, br, brv, _ = rng.choice([p for p in POOL if p[4] == want_direct])
                rating = round(min(5.0, max(3.0, br + rng.uniform(-0.5, 0.4))), 1)
                reviews = max(3, int(brv * rng.uniform(0.15, 1.6) * (0.45 + 0.55 * dens_f)))
                dist = round(rng.uniform(0.02, radius * 0.95), 3)
                comps.append({"n": cnm, "t": ty, "r": rating, "rv": reviews, "d": dist,
                              "s": round(influence(rating, reviews, dist, radius), 2),
                              "c": "d" if want_direct else "i"})
        comps.sort(key=lambda c: -c["s"])

        d = [c for c in comps if c["c"] == "d"]
        i = [c for c in comps if c["c"] == "i"]
        dt = round(sum(c["s"] for c in d), 2)
        it = round(sum(c["s"] for c in i), 2)
        rep = {
            "d": len(d), "i": len(i), "dt": dt, "it": it,
            "tot": round(dt + INDIRECT_W * it, 2),
            "dar": round(sum(c["r"] for c in d) / len(d), 2) if d else 0.0,
            "dtr": sum(c["rv"] for c in d),
            "iar": round(sum(c["r"] for c in i) / len(i), 2) if i else 0.0,
            "itr": sum(c["rv"] for c in i),
        }
        stores.append({
            "id": "MOCK%04d" % (k + 1), "name": nm, "region": region,
            "sup": rng.choice(SUPERVISORS),
            "addr": "%d %s, %s, CA" % (rng.randint(100, 9800),
                                       rng.choice(["Main St", "El Camino Real",
                                                   "Market St", "Broadway",
                                                   "First St", "Valley Blvd"]),
                                       nm.split(" ")[0]),
            "tier": tier, "layer": layer, "radius": radius,
            "areaType": rng.choice(AREA_TYPES),
            "lat": round(rng.uniform(32.5, 40.5), 6),
            "lng": round(rng.uniform(-124.0, -117.0), 6),
            "pop": pop, "dens": dens, "income": income, "mage": mage,
            "eth": eth, "age": age,
            "own": {"rating": round(rng.uniform(3.6, 4.8), 1),
                    "reviews": rng.randint(20, 900), "name": "Sharetea"},
            "rep": rep, "comps": comps, "mock": True,
        })
    return stores


# ============================================================================
# MAIN
# ============================================================================
def main():
    ap = argparse.ArgumentParser(
        description="Build data.js for the Sharetea Competitor Monitor from "
                    "competitor_scoring_output.xlsx (one-click).")
    ap.add_argument("--input", default="competitor_scoring_output.xlsx",
                    help="workbook to read (default: competitor_scoring_output.xlsx)")
    ap.add_argument("--out", default="data.js",
                    help="output file (default: data.js)")
    ap.add_argument("--mock-min", type=int, default=20,
                    help="pad the roster to this many stores when the workbook "
                         "has fewer (default: 20). Real >= this means NO mock.")
    ap.add_argument("--no-mock", action="store_true",
                    help="never generate mock stores, whatever the real count")
    ap.add_argument("--force-mock", type=int, default=None, metavar="N",
                    help="always pad the roster to N stores")
    ap.add_argument("--validate-only", action="store_true",
                    help="check the contract and exit without writing anything")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        sys.exit("ERROR: workbook not found: %s\n"
                 "Put competitor_scoring_output.xlsx next to this script, or pass "
                 "--input <path>." % args.input)

    print("Reading %s" % args.input)
    wb = openpyxl.load_workbook(args.input, data_only=True)

    print("Validating against data contract v1 ...")
    try:
        validate(wb)
    except ContractError as e:
        sys.exit("\n" + str(e))
    print("  contract OK (%d sheets, %d store columns)"
          % (len(CONTRACT), len(CONTRACT[SHEET_STORES])))

    real = read_real_stores(wb)
    tier_summary = read_tier_summary(wb)
    if not real:
        sys.exit("ERROR: the store sheet has no data rows.")

    # ---- how many mock stores, if any -------------------------------------
    if args.no_mock:
        n_mock = 0
        why = "--no-mock"
    elif args.force_mock is not None:
        n_mock = max(0, args.force_mock - len(real))
        why = "--force-mock %d" % args.force_mock
    elif len(real) >= args.mock_min:
        n_mock = 0
        why = "real %d >= mock-min %d" % (len(real), args.mock_min)
    else:
        n_mock = args.mock_min - len(real)
        why = "real %d < mock-min %d" % (len(real), args.mock_min)

    mock = make_mock_stores(n_mock) if n_mock else []
    stores = real + mock
    stores.sort(key=lambda s: -s["rep"]["tot"])

    meta = {
        "generated": datetime.date.today().isoformat(),
        "source": os.path.basename(args.input),
        "scanDate": None,          # the workbook does not carry a scan date
        "realStores": len(real),
        "mockStores": len(mock),
        "totalStores": len(stores),
        "realCompetitorRows": sum(len(s["comps"]) for s in real),
        "storeRef": STORE_REF, "compRef": COMP_REF,
        "threatMin": THREAT_MIN, "indirectWeight": INDIRECT_W,
        "normBase": STORE_REF,
        "bands": [
            {"name": "Severe", "min": 90, "max": 100, "color": "#E5484D"},
            {"name": "High", "min": 70, "max": 90, "color": "#E5484D"},
            {"name": "Moderate", "min": 40, "max": 70, "color": "#F5A623"},
            {"name": "Low", "min": 0, "max": 40, "color": "#999999"},
        ],
        "tiers": [{"tier": t[0], "radius": t[1], "layer": t[2]} for t in TIERS],
        "tierSummary": tier_summary,
    }

    payload = {"meta": meta, "stores": stores}
    js = ("/* Sharetea Competitor Monitor - data contract v1 (frozen)\n"
          "   Generated by build_data.py from %s\n"
          "   %d real store(s) (verbatim) + %d mock store(s) (mock:true).\n"
          "   No external API was called. */\n"
          "window.SHARETEA_DATA = %s;\n"
          % (os.path.basename(args.input), len(real), len(mock),
             json.dumps(payload, ensure_ascii=False, separators=(",", ":"))))

    if args.validate_only:
        print("\n--validate-only: contract OK, nothing written.")
        return

    with open(args.out, "w", encoding="utf-8") as f:
        f.write(js)

    # ---- summary ----------------------------------------------------------
    print("")
    print("=" * 62)
    print("  BUILD SUMMARY")
    print("=" * 62)
    print("  workbook            : %s" % args.input)
    print("  stores read         : %d" % len(real))
    print("  competitor rows     : %d" % meta["realCompetitorRows"])
    print("  mock stores         : %d   (%s)" % (len(mock), why))
    print("  total stores        : %d" % len(stores))
    print("  output              : %s  (%.1f KB)"
          % (args.out, len(js.encode("utf-8")) / 1024))
    print("  normalisation       : min(100, raw / %d * 100)" % STORE_REF)
    print("-" * 62)
    print("  real stores, score check (front end recomputes from competitor rows):")
    for s in real:
        d = [c for c in s["comps"] if c["c"] == "d"]
        i = [c for c in s["comps"] if c["c"] == "i"]
        calc = round(sum(c["s"] for c in d) + INDIRECT_W * sum(c["s"] for c in i), 2)
        ok = "OK" if abs(calc - s["rep"]["tot"]) < 0.02 else "MISMATCH"
        print("    %-12s %-16s calc=%8.2f  excel=%8.2f  norm=%5.1f  %s"
              % (s["id"], s["name"][:16], calc, s["rep"]["tot"],
                 min(100.0, calc / STORE_REF * 100), ok))
    print("=" * 62)
    if len(mock):
        print("  NOTE: %d mock store(s) were generated because the workbook holds "
              "only %d." % (len(mock), len(real)))
        print("        They are tagged mock:true and the front end labels them.")
    else:
        print("  No mock stores: the workbook alone fills the roster.")
    print("")


if __name__ == "__main__":
    main()
