# ============================================================
# generate_data.py
# Multi-Year Sales Consolidation (2013–2024)
# ============================================================
# PURPOSE : Generate 12 realistic mock Excel sales files,
#           one per year from 2013 to 2024.
# RUN     : python generate_data.py
# OUTPUT  : data/raw/sales_YYYY.xlsx  (12 files)
# ============================================================

import os
import random
import numpy as np
import pandas as pd
from faker import Faker
from datetime import datetime, timedelta

# Seed for reproducibility (remove to get different data each run)
random.seed(42)
np.random.seed(42)
fake = Faker()
Faker.seed(42)

# ── Output folder ────────────────────────────────────────────
OUTPUT_DIR = os.path.join("data", "raw")
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ── Master reference lists ───────────────────────────────────
CATEGORIES = {
    "Technology": {
        "sub_categories": ["Phones", "Computers", "Accessories", "Copiers"],
        "products": {
            "Phones":       ["iPhone 14", "Samsung Galaxy S23", "Google Pixel 7", "OnePlus 11", "Motorola Edge"],
            "Computers":    ["Dell XPS 15", "MacBook Pro 14", "HP Pavilion", "Lenovo ThinkPad X1", "ASUS ZenBook"],
            "Accessories":  ["Logitech Webcam", "USB-C Hub", "Wireless Mouse", "Mechanical Keyboard", "Monitor Stand"],
            "Copiers":      ["Canon ImageRunner", "Xerox VersaLink", "HP LaserJet Pro", "Brother MFC-L8900", "Ricoh SP"],
        },
        "price_range": (100, 2500),
    },
    "Furniture": {
        "sub_categories": ["Chairs", "Tables", "Bookcases", "Furnishings"],
        "products": {
            "Chairs":       ["Ergonomic Office Chair", "Executive Chair", "Task Chair", "Conference Chair", "Mesh Chair"],
            "Tables":       ["Standing Desk", "Meeting Table", "Coffee Table", "Adjustable Desk", "Folding Table"],
            "Bookcases":    ["5-Shelf Bookcase", "3-Shelf Bookcase", "Glass Bookcase", "Corner Bookshelf", "Metal Rack"],
            "Furnishings":  ["Desk Lamp", "Filing Cabinet", "Whiteboard", "Cubicle Divider", "Magazine Rack"],
        },
        "price_range": (50, 1800),
    },
    "Office Supplies": {
        "sub_categories": ["Paper", "Binders", "Art", "Envelopes", "Fasteners", "Labels", "Storage"],
        "products": {
            "Paper":        ["A4 Ream 500 sheets", "Letter Size Paper", "Card Stock", "Photo Paper", "Graph Paper"],
            "Binders":      ["3-Ring Binder 2in", "1-Ring Binder", "Report Cover", "Hanging Folder", "Expanding File"],
            "Art":          ["Staedtler Markers", "Watercolor Set", "Colored Pencils", "Sketch Pad", "Acrylic Paint"],
            "Envelopes":    ["Box of 50 Envelopes", "Padded Mailers", "Bubble Wrap Envelopes", "Security Envelopes", "Window Envelopes"],
            "Fasteners":    ["Paper Clips Box", "Binder Clips Set", "Stapler", "Staples Box", "Rubber Bands"],
            "Labels":       ["Avery Address Labels", "Shipping Labels", "File Folder Labels", "Color Coding Labels", "Name Badge Labels"],
            "Storage":      ["Storage Box 12pk", "Plastic Bins Set", "Drawer Organizer", "Desktop Organizer", "Cable Box"],
        },
        "price_range": (5, 150),
    },
}

REGIONS   = ["East", "West", "Central", "South"]
COUNTRIES = {
    "East":    ["United States"],
    "West":    ["United States"],
    "Central": ["United States"],
    "South":   ["United States"],
}
SEGMENTS   = ["Consumer", "Corporate", "Home Office"]
SHIP_MODES = ["Standard Class", "Second Class", "First Class", "Same Day"]

# ── Revenue growth multipliers per year ──────────────────────
# Slight dips in 2016, 2019, 2020 (COVID)
YEAR_MULTIPLIERS = {
    2013: 1.00,
    2014: 1.08,
    2015: 1.15,
    2016: 1.10,   # slight dip
    2017: 1.20,
    2018: 1.30,
    2019: 1.25,   # slight dip
    2020: 1.18,   # COVID dip
    2021: 1.38,
    2022: 1.50,
    2023: 1.62,
    2024: 1.75,
}

# ── Helper functions ─────────────────────────────────────────

def random_order_date(year: int) -> str:
    """Return a random date within the given year."""
    start = datetime(year, 1, 1)
    end   = datetime(year, 12, 31)
    delta = end - start
    return (start + timedelta(days=random.randint(0, delta.days))).strftime("%Y-%m-%d")


def pick_category_and_product():
    """Return category, sub_category, product_name, and base unit price."""
    category = random.choice(list(CATEGORIES.keys()))
    info     = CATEGORIES[category]
    sub_cat  = random.choice(info["sub_categories"])
    product  = random.choice(info["products"][sub_cat])
    price    = round(random.uniform(*info["price_range"]), 2)
    return category, sub_cat, product, price


def generate_year_data(year: int, n_rows: int) -> pd.DataFrame:
    """Generate n_rows rows of sales data for a given year."""
    mult = YEAR_MULTIPLIERS[year]
    records = []

    for i in range(n_rows):
        order_id_num       = f"ORD-{year}-{str(i + 1).zfill(5)}"
        customer_name      = fake.name()
        segment            = random.choices(SEGMENTS, weights=[50, 35, 15])[0]
        region             = random.choices(REGIONS,  weights=[30, 30, 25, 15])[0]
        country            = "United States"
        category, sub_cat, product, unit_price = pick_category_and_product()

        # Apply year multiplier to price
        unit_price = round(unit_price * mult, 2)
        quantity   = random.randint(1, 10)
        discount   = round(random.choice([0, 0, 0, 0.05, 0.10, 0.15, 0.20, 0.30, 0.40]), 2)
        sales      = round(quantity * unit_price * (1 - discount), 2)

        # Profit: mostly positive, ~15% chance of loss (realistic)
        profit_pct = random.uniform(-0.20, 0.45)
        profit     = round(sales * profit_pct, 2)

        ship_mode     = random.choices(SHIP_MODES, weights=[55, 25, 15, 5])[0]
        shipping_cost = round(random.uniform(3, 60), 2)

        records.append({
            "Order_ID":        order_id_num,
            "Order_Date":      random_order_date(year),
            "Customer_Name":   customer_name,
            "Customer_Segment":segment,
            "Region":          region,
            "Country":         country,
            "Category":        category,
            "Sub_Category":    sub_cat,
            "Product_Name":    product,
            "Quantity":        quantity,
            "Unit_Price":      unit_price,
            "Discount":        discount,
            "Sales":           sales,
            "Profit":          profit,
            "Shipping_Cost":   shipping_cost,
            "Ship_Mode":       ship_mode,
        })

    return pd.DataFrame(records)


def introduce_data_quality_issues(df: pd.DataFrame, year: int) -> pd.DataFrame:
    """
    Intentionally inject data quality issues so the pipeline has
    something real to clean:
      1. Missing values in Discount, Shipping_Cost, Region
      2. Inconsistent text casing in Category
      3. Some Order_Date values as Python date objects (not strings)
      4. 2–3 duplicate rows
    """
    n = len(df)

    # 1a. Missing Discount (~3 % of rows)
    missing_discount_idx = df.sample(frac=0.03, random_state=year).index
    df.loc[missing_discount_idx, "Discount"] = np.nan

    # 1b. Missing Shipping_Cost (~4 % of rows)
    missing_ship_idx = df.sample(frac=0.04, random_state=year + 1).index
    df.loc[missing_ship_idx, "Shipping_Cost"] = np.nan

    # 1c. Missing Region (~2 % of rows)
    missing_region_idx = df.sample(frac=0.02, random_state=year + 2).index
    df.loc[missing_region_idx, "Region"] = np.nan

    # 2. Inconsistent Category casing
    cat_idx = df.sample(frac=0.15, random_state=year + 3).index
    df.loc[cat_idx, "Category"] = df.loc[cat_idx, "Category"].apply(
        lambda x: random.choice([x.upper(), x.lower(), x.title()]) if isinstance(x, str) else x
    )

    # 3. Convert ~5 % of Order_Date to a different string format (MM/DD/YYYY)
    #    Simulates inconsistent date formats found in real data.
    #    (pandas 3.0+ enforces string dtype — datetime objects cannot be stored
    #    in a string column, so we use an alternate string format instead.)
    date_obj_idx = df.sample(frac=0.05, random_state=year + 4).index
    df.loc[date_obj_idx, "Order_Date"] = pd.to_datetime(
        df.loc[date_obj_idx, "Order_Date"]
    ).dt.strftime("%m/%d/%Y")

    # 4. Duplicate rows (2–3 exact copies)
    n_dupes = random.randint(2, 3)
    dupe_rows = df.sample(n=n_dupes, random_state=year + 5)
    df = pd.concat([df, dupe_rows], ignore_index=True)

    return df


# ── Main ─────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  Multi-Year Sales Data Generator")
    print("  Generating files for years 2013 – 2024")
    print("=" * 60)

    for year in range(2013, 2025):
        n_rows = random.randint(800, 1200)
        print(f"\n[{year}]  Generating {n_rows} rows...", end=" ", flush=True)

        df = generate_year_data(year, n_rows)
        df = introduce_data_quality_issues(df, year)

        filepath = os.path.join(OUTPUT_DIR, f"sales_{year}.xlsx")
        df.to_excel(filepath, index=False, engine="openpyxl")
        print(f"Saved → {filepath}  ({len(df)} rows after duplicates)")

    print("\n" + "=" * 60)
    print("  All 12 files generated successfully.")
    print(f"  Location: {os.path.abspath(OUTPUT_DIR)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
