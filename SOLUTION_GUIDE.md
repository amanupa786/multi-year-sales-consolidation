# COMPLETE SOLUTION GUIDE
## Multi-Year Sales Consolidation (2013–2024)

> **Who this is for:** Anyone. Whether you are a beginner, a student, or someone who just ran
> the project but has no idea what happened — this document explains EVERYTHING.
> You do not need to Google a single thing. Every concept, every word, every line of logic
> is explained here in plain simple English.

---

# TABLE OF CONTENTS

1. [The Story — What Problem Does This Project Solve?](#1-the-story)
2. [The Folder Structure — What Is Where and Why](#2-folder-structure)
3. [Python Basics You Need to Know First](#3-python-basics)
4. [The Libraries — What They Are and Why We Use Them](#4-libraries)
5. [How generate_data.py Works — Line by Line](#5-generate-data)
6. [How pipeline.py Works — Every Step Explained](#6-pipeline)
7. [SQL From Scratch — Every Concept Used in This Project](#7-sql)
8. [The Charts — What They Show and What They Mean](#8-charts)
9. [The Excel Reports — What Is in Them](#9-excel-reports)
10. [The Jupyter Notebook — How to Use It](#10-jupyter)
11. [Common Errors and How to Fix Them](#11-errors)
12. [Learning Checkpoint Answers](#12-checkpoint-answers)
13. [How to Extend This Project](#13-extend)
14. [Interview Questions and Strong Answers](#14-interview)
15. [How to Present This on GitHub and in Interviews](#15-present)

---

# 1. THE STORY

## What Was the Problem?

Imagine you work as a data analyst at a retail company. Every year, the sales team saves that
year's sales records into an Excel file. After 12 years, you now have 12 files sitting in a folder:

```
sales_2013.xlsx
sales_2014.xlsx
sales_2015.xlsx
...
sales_2024.xlsx
```

Every time your manager asks "How did our revenue grow over the last 10 years?", you have to:

1. Open all 12 files one by one
2. Copy the rows from each file
3. Paste them into one big master Excel sheet
4. Realise that some files spell "Technology" as "TECHNOLOGY" or "technology"
5. Manually find and fix all the spelling differences
6. Remove rows that appear twice (duplicates)
7. Fix dates that are in different formats
8. Build charts from scratch
9. Create a summary table

This takes **3 full working days** every single quarter. And if you paste one file in the wrong
place, your revenue numbers are wrong — and you might not even notice.

## What Does This Project Do Instead?

This project replaces those 3 days with a single command:

```
python pipeline.py
```

You type that, press Enter, and in under 10 minutes you get:
- One clean master file with all 12 years of data combined
- A Year-over-Year growth report in Excel
- A 6-sheet summary report
- 3 professional charts saved as image files
- All data loaded into a MySQL database for SQL analysis

**That is the entire point of this project.** Everything else is just the technical steps to achieve it.

---

# 2. FOLDER STRUCTURE

```
multi-year-sales-consolidation/
│
├── data/
│   ├── raw/                 ← The original 12 Excel files. NEVER edit these directly.
│   └── processed/           ← The cleaned, combined CSV file goes here.
│
├── notebooks/               ← The Jupyter notebook for exploring data visually.
│
├── outputs/
│   ├── charts/              ← The 3 chart images (PNG files) go here.
│   └── ...                  ← The Excel report files go here.
│
├── sql/                     ← All SQL query files go here.
│
├── config.py                ← MySQL credentials. Edit ONLY this file.
├── generate_data.py         ← Creates the 12 fake Excel files.
├── pipeline.py              ← The main script. This is the real project.
├── requirements.txt         ← List of Python libraries needed.
└── README.md                ← The GitHub landing page for the project.
```

### Why "raw" and "processed" folders?

The `raw` folder holds the original files exactly as received. You never touch them directly.
The `processed` folder holds the cleaned output. This separation is a professional standard
called **non-destructive processing** — if something goes wrong in your cleaning code, you
still have the original data safe in `raw` and can re-run everything.

Think of it like a kitchen: the raw ingredients go in one drawer, the cooked meal goes on
the plate. You don't put leftover soup back in the original can.

---

# 3. PYTHON BASICS YOU NEED TO KNOW FIRST

Before understanding the code, you need to know a few fundamental Python concepts.
If you already know these, skip to Section 4.

## What Is a Variable?

A variable is a box that stores a value with a name.

```python
year = 2024          # The box "year" holds the number 2024
name = "Technology"  # The box "name" holds the text "Technology"
price = 199.99       # The box "price" holds a decimal number
```

## What Is a List?

A list is a collection of items in a specific order.

```python
years = [2013, 2014, 2015, 2016]
categories = ["Technology", "Furniture", "Office Supplies"]
```

You access items by position (starting from 0):
```python
categories[0]  # → "Technology"
categories[1]  # → "Furniture"
```

## What Is a Dictionary?

A dictionary stores information in key-value pairs — like a real dictionary where the word
is the key and the definition is the value.

```python
product_info = {
    "name":  "MacBook Pro",
    "price": 1999.99,
    "category": "Technology"
}

product_info["name"]   # → "MacBook Pro"
product_info["price"]  # → 1999.99
```

## What Is a Loop?

A loop repeats a block of code multiple times.

```python
# This runs 3 times
for year in [2013, 2014, 2015]:
    print(f"Processing year {year}")

# Output:
# Processing year 2013
# Processing year 2014
# Processing year 2015
```

## What Is a Function?

A function is a reusable block of code that does one specific job.
You define it once and call it whenever needed.

```python
# Define the function
def calculate_profit_margin(profit, sales):
    return (profit / sales) * 100

# Call the function
margin = calculate_profit_margin(500, 2000)
print(margin)  # → 25.0
```

## What Is a DataFrame?

A DataFrame is the most important concept in this project. Think of it as a Python
version of an Excel table — rows and columns, but programmable.

```
Order_ID   | Order_Date  | Category   | Sales  | Profit
-----------|-------------|------------|--------|-------
ORD-001    | 2020-01-15  | Technology | 1200   | 240
ORD-002    | 2020-01-18  | Furniture  | 350    | -50
ORD-003    | 2020-02-03  | Technology | 899    | 180
```

This whole table is a DataFrame. Each column is a Series (a single column of data).
The pandas library creates and manages DataFrames.

## What Is NaN?

NaN stands for "Not a Number." It is Python's way of representing a missing value — a cell
that is blank or empty.

In Excel, an empty cell is just blank. In pandas, that empty cell becomes NaN.
You cannot do math with NaN (5 + NaN = NaN), so you must fill or remove NaN values
before analysing data. That is what Step 2 of the pipeline does.

---

# 4. THE LIBRARIES

A library is a collection of pre-written code that you can use in your own programs.
Instead of writing a function to read Excel files from scratch (which would take weeks),
you just install a library that already does it perfectly.

## pandas

**What it is:** The backbone of all data work in Python.

**What it does:** Reads Excel/CSV files, creates DataFrames, filters rows, sorts data,
groups and aggregates (like Pivot Tables), merges tables, and saves outputs.

**Real-world analogy:** Pandas is like Excel, but instead of clicking menus, you type
instructions. And it handles millions of rows without freezing.

**Key functions used in this project:**

| Function | What it does | Excel equivalent |
|---|---|---|
| `pd.read_excel("file.xlsx")` | Opens an Excel file | File → Open |
| `pd.read_csv("file.csv")` | Opens a CSV file | File → Open |
| `df.groupby("Year")["Sales"].sum()` | Totals Sales per Year | Pivot Table |
| `pd.concat([df1, df2, df3])` | Stacks tables on top of each other | Copy-Paste rows |
| `df.merge(df2, on="Year")` | Joins two tables on a common column | VLOOKUP |
| `df.drop_duplicates()` | Removes duplicate rows | Remove Duplicates |
| `df["Col"].fillna(0)` | Fills blank cells with 0 | Find & Replace blank→0 |
| `df.to_csv("output.csv")` | Saves as CSV file | Save As CSV |
| `df.to_excel("output.xlsx")` | Saves as Excel file | Save As Excel |

## numpy

**What it is:** A math library for Python. Pandas is built on top of numpy.

**What it does:** Fast calculations on entire columns at once, special values like `np.nan`
(missing value) and `np.inf` (infinity), and array operations.

**Why we use it here:**
- `np.nan` → represents a missing value
- `np.where(condition, value_if_true, value_if_false)` → like Excel's IF formula
- `np.zeros(n)` → creates an array of n zeros (used for stacking chart bars)

**Real-world analogy:** If pandas is Excel, numpy is the calculator engine inside Excel
that handles all the math.

## openpyxl

**What it is:** A library specifically for reading and writing `.xlsx` Excel files.

**What it does:** Pandas uses openpyxl "under the hood" to open and save Excel files.
We also use it directly to apply formatting — bold headers, coloured cells, column widths.

**Why not just use pandas for everything?** Pandas writes the data. openpyxl controls
the appearance — the colours, fonts, borders, and frozen rows. Pandas cannot do that styling
on its own.

**Real-world analogy:** Pandas fills in the spreadsheet data. openpyxl is the formatting
paintbrush that makes it look professional.

## faker

**What it is:** A library that generates realistic fake data.

**What it does:** `fake.name()` returns a random realistic name like "Sarah Johnson" or
"Michael Chen". We use it to generate customer names.

**Why we need it:** Our project needs customer names that look real but are not real.
We could write `["Customer 1", "Customer 2", ...]` but that looks terrible in a portfolio.
Faker makes the data look like it came from a real business.

**Important:** This library is ONLY used in `generate_data.py`. In a real job, you would
never use faker — your real data already has real customer names.

## matplotlib

**What it is:** Python's main chart-drawing library.

**What it does:** Creates every type of chart — line, bar, scatter, pie, histogram —
and saves them as image files (PNG, PDF, etc.).

**How it works conceptually:**
```python
fig, ax = plt.subplots()  # Create a blank canvas (fig) with a chart area (ax)
ax.plot(x_values, y_values)  # Draw a line on the chart area
ax.set_title("My Chart")  # Add a title
plt.savefig("chart.png")  # Save to a file
```

`fig` is the entire image (like the picture frame).
`ax` is the chart inside the frame (the actual drawing).

## seaborn

**What it is:** A library built on top of matplotlib that makes beautiful statistical charts
with less code.

**What it does:** Creates heatmaps, box plots, violin plots, and applies professional
colour themes with one line of code.

**Why use both matplotlib and seaborn?** Seaborn is for quick, beautiful charts.
Matplotlib is for full control and customisation. We use seaborn for the heatmap and
matplotlib for the line and bar charts where we need precise control.

**Real-world analogy:** Matplotlib is painting with a brush (full control, more effort).
Seaborn is using a paint roller (faster, good results, less control).

## mysql-connector-python

**What it is:** The official driver that lets Python talk to a MySQL database.

**What it does:** Opens a connection to MySQL, sends SQL commands (CREATE TABLE, INSERT, SELECT),
and retrieves results.

**Real-world analogy:** Think of MySQL as a building with a locked door. mysql-connector is
the key that lets your Python script enter, put data inside, and query it.

---

# 5. HOW generate_data.py WORKS

## Why This File Exists

In a real job, you receive data from your company's sales system or database.
You would never need to create data from scratch.

This file exists ONLY because we need something to practise with. It creates 12 realistic
fake Excel files so the pipeline has something to clean and analyse.

**You would never write this file in a real job. Think of it as setup scaffolding.**

## Step-by-Step Walkthrough

### Part 1 — Setting Up

```python
import random
import numpy as np
import pandas as pd
from faker import Faker
```

These `import` statements load the libraries into memory. Until you import a library,
Python does not know it exists. It is like opening a toolbox — you need to open it
before you can use the tools inside.

```python
random.seed(42)
np.random.seed(42)
fake = Faker()
Faker.seed(42)
```

**What is a seed?** A seed makes "random" numbers reproducible. Without a seed, every time
you run the script, you get different random data. With seed 42, you always get the exact
same data. This is important for testing — if results change every run, you cannot compare them.

The number 42 is not special. It is just a convention (from the book "The Hitchhiker's
Guide to the Galaxy"). You can use any number.

### Part 2 — The Product Catalogue

```python
CATEGORIES = {
    "Technology": {
        "sub_categories": ["Phones", "Computers", "Accessories", "Copiers"],
        "products": {
            "Phones": ["iPhone 14", "Samsung Galaxy S23", ...],
            ...
        },
        "price_range": (100, 2500),
    },
    ...
}
```

This is a **nested dictionary** — a dictionary inside a dictionary. Think of it as a
filing cabinet:
- Drawer 1: "Technology"
  - Folder A: sub_categories
  - Folder B: products (another set of folders inside)
  - Folder C: price_range

We use this structure so that when we randomly pick "Technology" as a category, we can
immediately look up what sub-categories and products belong to it, and what a realistic
price range is.

### Part 3 — Year Multipliers

```python
YEAR_MULTIPLIERS = {
    2013: 1.00,
    2014: 1.08,
    2015: 1.15,
    2016: 1.10,   # dip
    2020: 1.18,   # COVID dip
    2024: 1.75,
}
```

This controls the business trend. When generating data for 2024, every unit price is
multiplied by 1.75 compared to 2013. So a product that cost $100 in 2013 costs $175 in 2024.

This is why the revenue trend chart shows growth — not because we generate more rows, but
because prices increase each year, simulating real business inflation and growth.

The dips in 2016 and 2020 simulate real events (market slowdowns and COVID-19).

### Part 4 — Generating One Year of Data

```python
def generate_year_data(year, n_rows):
    records = []
    for i in range(n_rows):
        # Pick a random category, sub-category, product
        category, sub_cat, product, unit_price = pick_category_and_product()
        
        # Apply growth multiplier
        unit_price = round(unit_price * mult, 2)
        
        quantity = random.randint(1, 10)
        discount = random.choice([0, 0, 0, 0.05, 0.10, 0.15, 0.20, 0.30, 0.40])
        sales = round(quantity * unit_price * (1 - discount), 2)
        
        records.append({...})   # Add all values as a dictionary
    
    return pd.DataFrame(records)   # Convert list of dicts to a DataFrame
```

**How `pd.DataFrame(records)` works:**
`records` is a Python list of dictionaries. Each dictionary is one row. Pandas looks at
the keys of the first dictionary ("Order_ID", "Sales", "Profit", etc.) and uses them as
column names. Then it fills each row from each dictionary.

This is the standard pattern for building a DataFrame row-by-row.

**Why is `discount` chosen from a list and not truly random?**

```python
discount = random.choice([0, 0, 0, 0.05, 0.10, 0.15, 0.20, 0.30, 0.40])
```

The list has three 0s and one each of the other values. So the chance of getting 0 discount
is 3/9 = 33%. This reflects reality — most orders have no discount. A truly random number
between 0 and 0.40 would give too many discounted orders.

### Part 5 — Introducing Data Quality Issues

This is the most important conceptual part of `generate_data.py`.

```python
def introduce_data_quality_issues(df, year):
```

This function deliberately breaks the data in 4 ways:

**Break 1 — Missing Values**
```python
missing_discount_idx = df.sample(frac=0.03).index
df.loc[missing_discount_idx, "Discount"] = np.nan
```

`df.sample(frac=0.03)` picks 3% of rows randomly.
`.index` gets the row numbers of those chosen rows.
`df.loc[those_row_numbers, "Discount"] = np.nan` sets those cells to blank (NaN).

This simulates what happens in real databases — sometimes a field was never entered.

**Break 2 — Inconsistent Casing**
```python
df.loc[cat_idx, "Category"] = df.loc[cat_idx, "Category"].apply(
    lambda x: random.choice([x.upper(), x.lower(), x.title()])
)
```

**What is a lambda?** A lambda is a tiny one-line function without a name.
`lambda x: x.upper()` means: "take x, return x in uppercase."

`random.choice([x.upper(), x.lower(), x.title()])` randomly picks one of:
- "TECHNOLOGY" (upper)
- "technology" (lower)
- "Technology" (title)

This simulates different people entering data in different ways.

**Break 3 — Wrong Date Format**
```python
df.loc[date_obj_idx, "Order_Date"] = pd.to_datetime(
    df.loc[date_obj_idx, "Order_Date"]
).dt.strftime("%m/%d/%Y")
```

Most dates are stored as `"2020-03-15"` (YYYY-MM-DD format).
This converts 5% of them to `"03/15/2020"` (MM/DD/YYYY format).

Now the column has two date formats mixed together. The pipeline's cleaning step must
detect and standardise all of them.

**Break 4 — Duplicate Rows**
```python
n_dupes = random.randint(2, 3)
dupe_rows = df.sample(n=n_dupes)
df = pd.concat([df, dupe_rows], ignore_index=True)
```

`pd.concat` stacks the original DataFrame and the duplicate rows on top of each other.
`ignore_index=True` resets row numbers from 0 after stacking.

This simulates what happens when data is exported twice or entered twice by mistake.

---

# 6. HOW pipeline.py WORKS

## The Big Picture

`pipeline.py` is the main script. It runs 7 steps in sequence. Each step takes the
output of the previous one. Here is the complete flow:

```
12 messy Excel files
       ↓ Step 1: Load
12 DataFrames in memory (one per year)
       ↓ Step 2: Clean
12 clean DataFrames in memory
       ↓ Step 3: Consolidate
1 master DataFrame (11,593 rows)
       ↓ Step 4: YoY Analysis
1 summary DataFrame (12 rows, one per year)
       ↓ Step 5: Summary Report
outputs/summary_report.xlsx (6 sheets)
       ↓ Step 6: Charts
outputs/charts/*.png (3 image files)
       ↓ Step 7: MySQL Load
Database table with all 11,593 rows
```

## Step 1 — Load All Files

### What glob does

```python
import glob
files = sorted(glob.glob("data/raw/sales_*.xlsx"))
```

`glob.glob("data/raw/sales_*.xlsx")` looks inside the `data/raw` folder and returns
a list of every file whose name matches the pattern `sales_*.xlsx`.

The `*` is called a **wildcard** — it means "any text here." So it matches:
- `sales_2013.xlsx` ✓
- `sales_2024.xlsx` ✓
- `sales_notes.txt` ✗ (does not end with .xlsx)
- `revenue_2013.xlsx` ✗ (does not start with sales_)

`sorted()` puts the list in alphabetical order, so files are processed 2013 → 2024.

### How the year is extracted from the filename

```python
filename = "sales_2013.xlsx"
year = int(filename.replace("sales_", "").replace(".xlsx", ""))
# Step 1: "sales_2013.xlsx" → "2013.xlsx"   (remove "sales_")
# Step 2: "2013.xlsx"       → "2013"         (remove ".xlsx")
# Step 3: int("2013")       → 2013           (convert text "2013" to number 2013)
```

### Why add a Year column?

When we stack all 12 DataFrames together, rows from 2013 and 2014 look identical in
structure. Without a Year column, we could never tell which year a row belongs to.

```python
df["Year"] = year
```

This adds a new column called "Year" and fills every row in that DataFrame with the year
number. After stacking, we can filter by year, group by year, or calculate year-on-year changes.

## Step 2 — Data Cleaning

This is the most important step. It fixes all the problems that `generate_data.py` introduced
(and that real-world data always has).

### 2a — Standardise Column Names

```python
df.columns = [c.strip().title().replace(" ", "_") for c in df.columns]
```

This is a **list comprehension** — a compact way to apply an operation to every item in a list.

It does three things to every column name:
1. `.strip()` → removes spaces from both ends: `" Order ID "` → `"Order ID"`
2. `.title()` → capitalises first letter of each word: `"order id"` → `"Order Id"`
3. `.replace(" ", "_")` → replaces spaces with underscores: `"Order Id"` → `"Order_Id"`

**Why does this matter?** If one file has a column called `"order_id"` and another has
`"Order ID"`, pandas treats them as completely different columns. After stacking, you'd
have both columns with most values missing. Standardising names ensures they match perfectly.

### 2b — Fix Text Casing

```python
df["Category"] = df["Category"].astype(str).str.strip().str.title()
df["Category"] = df["Category"].replace("Nan", np.nan)
```

**Line 1 breakdown:**
- `.astype(str)` → converts every value in the column to text (including any numbers or NaN values)
- `.str.strip()` → removes extra spaces from both ends
- `.str.title()` → capitalises first letter of each word

So "TECHNOLOGY", "technology", and " technology " all become "Technology".

**Line 2 breakdown:**
When NaN (a missing value) is converted to string using `.astype(str)`, it becomes the
text "nan". After `.str.title()` it becomes "Nan". This line converts "Nan" back to a
real missing value (np.nan) so it is treated correctly in future operations.

### 2c — Fix Dates

```python
df["Order_Date"] = pd.to_datetime(df["Order_Date"], errors="coerce")
```

`pd.to_datetime()` is incredibly smart. It can read:
- `"2020-03-15"` → converts to a proper date
- `"03/15/2020"` → converts to a proper date
- `"March 15, 2020"` → converts to a proper date
- `"not a date"` → would crash... unless you use `errors="coerce"`

`errors="coerce"` means: "if you cannot convert this value to a date, just put NaN there
instead of crashing." This is essential for real-world data that might have garbage values.

### 2d — Fill Missing Discount

```python
df["Discount"] = df["Discount"].fillna(0)
```

**Why fill with 0?** Business logic. A missing discount almost certainly means no discount
was applied. It does not mean the discount is unknown — it means it was 0%. So replacing
with 0 is a safe, accurate assumption.

**What if we left the NaN?** Any calculation involving NaN produces NaN. So `Sales = Quantity
× Price × (1 - NaN)` would produce NaN for that row, and it would be excluded from revenue
totals. We would undercount revenue.

### 2e — Fill Missing Shipping Cost with the Median

```python
shipping_median = df["Shipping_Cost"].median()
df["Shipping_Cost"] = df["Shipping_Cost"].fillna(shipping_median)
```

**Why the median and not the mean?**

The **mean** (average) adds all values and divides by count.
The **median** is the middle value when all values are sorted.

Imagine shipping costs: $5, $8, $10, $12, $15, $400

- Mean = (5 + 8 + 10 + 12 + 15 + 400) / 6 = **$75** → That $400 outlier destroys the mean
- Median = middle value = **$11** → A fair representation of typical shipping cost

In datasets with extreme outliers, the median gives a much more representative "typical"
value. For shipping costs (which can occasionally be very high for large items), the median
is always the safer choice.

### 2f — Fill Missing Region with Mode

```python
region_mode = df["Region"].mode()[0]
df["Region"] = df["Region"].fillna(region_mode)
```

**What is mode?** Mode is the most frequently occurring value. If East appears 400 times,
West 350, Central 250, South 200 — the mode is "East."

For a missing region, we do not know the true value. The best guess is the most common
region in that year's data.

`.mode()` returns a Series (list of modes — there could be multiple if tied).
`[0]` takes the first one.

### 2g — Remove Duplicates

```python
df.drop_duplicates(inplace=True)
```

This removes any row that is **completely identical** to another row (same values in every
single column). The first occurrence is kept; subsequent duplicates are deleted.

**Important distinction:** This removes rows where ALL columns match. It does NOT remove
rows that share just an Order_ID — because the same customer might order multiple products
on the same order (different rows, same Order_ID is legitimate).

### 2h — Remove Zero/Null Sales

```python
df = df[df["Sales"].notna() & (df["Sales"] != 0)]
```

`df["Sales"].notna()` → True for rows where Sales is not NaN
`(df["Sales"] != 0)` → True for rows where Sales is not zero
`&` → both conditions must be True

Combined: keep only rows where Sales has a real non-zero value. A row with $0 in sales
is not a real transaction — it's a data entry error.

### 2i — Validate the Sales Formula

```python
df["Sales_Calculated"] = (df["Quantity"] * df["Unit_Price"] * (1 - df["Discount"])).round(2)
mismatch_mask = (df["Sales_Calculated"] - df["Sales"]).abs() / df["Sales"].abs() > 0.01
```

**What this does:** Checks whether the Sales column matches what you would expect from
the formula: `Sales = Quantity × Unit_Price × (1 − Discount)`

`(calculated - actual).abs()` → absolute difference (always positive)
`/ actual.abs()` → relative difference (as a percentage of the actual value)
`> 0.01` → flag if the difference is more than 1%

**Why 1% tolerance?** Floating point numbers in computers are not perfectly precise.
0.1 + 0.2 in Python is not exactly 0.3 — it is 0.30000000000000004. Small rounding errors
are normal. We only flag differences bigger than 1% as actual data problems.

We do NOT delete these rows. We just count them and log them in the Data Quality report.
The Sales column is the official record — we trust it over a recalculated value.

### 2j — Add Profit Margin Column

```python
df["Profit_Margin_%"] = np.where(
    df["Sales"] != 0,
    (df["Profit"] / df["Sales"] * 100).round(2),
    0
)
```

**What `np.where` does:**
`np.where(condition, value_if_true, value_if_false)` is Python's version of Excel's IF formula.

- If Sales is not zero → calculate `(Profit / Sales) * 100`
- If Sales IS zero → put 0 (to avoid dividing by zero, which would crash)

**Why is this column useful?** Raw profit numbers cannot be compared fairly across orders
of different sizes. A $200 profit on a $10,000 order (2% margin) is much worse than a
$200 profit on a $600 order (33% margin). The margin percentage makes the comparison fair.

### 2k — Add Revenue Band Column

```python
df["Revenue_Band"] = pd.cut(
    df["Sales"],
    bins=[-np.inf, 500, 2000, np.inf],
    labels=["Low", "Medium", "High"]
)
```

**What `pd.cut` does:** Divides a continuous numeric column into labelled buckets (bins).

`bins=[-np.inf, 500, 2000, np.inf]` creates three ranges:
- From negative infinity to 500 → "Low"
- From 500 to 2000 → "Medium"
- From 2000 to positive infinity → "High"

`-np.inf` means "any number, no matter how small."
`np.inf` means "any number, no matter how large."

**What this is used for:** You can now quickly filter or group by deal size instead of
having to write `df[df["Sales"] > 2000]` every time. It is a business-friendly category label.

## Step 3 — Consolidate

```python
master_df = pd.concat(cleaned_dfs, ignore_index=True)
master_df.sort_values(["Year", "Order_Date"], inplace=True)
master_df.reset_index(drop=True, inplace=True)
master_df.to_csv("data/processed/consolidated_sales.csv", index=False)
```

**`pd.concat(cleaned_dfs)`** → Takes a list of 12 DataFrames and stacks them vertically
(one on top of the other). Exactly like doing 12 copy-pastes in Excel, but instant.

**`ignore_index=True`** → After stacking, row numbers restart from 0. Without this, you
would have row numbers 0–999 from the first file, then 0–850 from the second file — duplicate
index numbers cause all kinds of problems.

**`sort_values(["Year", "Order_Date"])`** → Sorts by Year first, then by date within each year.
So all 2013 rows appear first in date order, then all 2014 rows, etc. Sorted data is easier
to read, faster to query, and produces correct time-series charts.

**`to_csv(index=False)`** → Saves as CSV. `index=False` means do not save the row numbers as
a column — they are internal to Python and not part of the actual data.

## Step 4 — Year-over-Year Analysis

```python
yoy = (
    master_df.groupby("Year")
    .agg(
        Total_Revenue = ("Sales",    "sum"),
        Total_Profit  = ("Profit",   "sum"),
        Total_Orders  = ("Order_ID", "nunique"),
    )
    .reset_index()
)
```

**What `groupby` does:**
`groupby("Year")` splits the 11,593-row master DataFrame into 12 groups — one per year.
Then `.agg()` applies aggregation functions to each group.

Think of it exactly like an Excel Pivot Table:
- Rows: Year
- Values: Sum of Sales, Sum of Profit, Count of unique Order IDs

`"nunique"` means "count of unique values" — it counts distinct Order_IDs (not total rows,
since one order can appear on multiple rows if it has multiple products).

### How YoY Growth is Calculated

```python
yoy["YoY_Revenue_Growth_%"] = (
    (yoy["Total_Revenue"] - yoy["Total_Revenue"].shift(1)) /
    yoy["Total_Revenue"].shift(1) * 100
).round(2)
```

**What `.shift(1)` does:**

Imagine your Total_Revenue column looks like this:
```
Year    Total_Revenue
2013    1,200,000
2014    1,296,000
2015    1,440,000
```

After `.shift(1)`, a new shifted version looks like:
```
Year    Total_Revenue    .shift(1)
2013    1,200,000        NaN          ← no previous year
2014    1,296,000        1,200,000    ← previous year's value
2015    1,440,000        1,296,000    ← previous year's value
```

Now the subtraction `Total_Revenue - shift(1)` gives the change vs. the previous year.
Dividing by `shift(1)` and multiplying by 100 gives the percentage change.

For 2013 (first year), the result is NaN because there is no 2012 to compare against.

### Finding the Best Category Per Year

```python
best_cat = (
    master_df.groupby(["Year", "Category"])["Sales"]
    .sum()
    .reset_index()
    .sort_values(["Year", "Sales"], ascending=[True, False])
    .drop_duplicates("Year")
)
```

**Step by step:**
1. `groupby(["Year", "Category"])["Sales"].sum()` → Total Sales per Year per Category
   Result: 2013-Technology, 2013-Furniture, 2013-Office Supplies... repeated for all years
2. `.sort_values(["Year", "Sales"], ascending=[True, False])` → Sort year ascending,
   sales descending. So for each year, the highest-revenue category appears first.
3. `.drop_duplicates("Year")` → Keep only the first row per year. Since highest revenue
   appears first, this keeps the best-performing category per year.

## Step 5 — Summary Report

This step builds a 6-sheet Excel workbook. The code logic is the same concept repeated:
group the master DataFrame by different dimensions, save each result as a new sheet.

```python
by_cat = master_df.groupby(["Year", "Category"]).agg(
    Revenue = ("Sales",  "sum"),
    Profit  = ("Profit", "sum")
).reset_index()
```

Then in the Excel writer:
```python
with pd.ExcelWriter("outputs/summary_report.xlsx", engine="openpyxl") as writer:
    yoy_df.to_excel(writer, sheet_name="YoY Trend")
    by_cat.to_excel(writer, sheet_name="By Category")
    ...
```

**What `with` means:** The `with` statement ensures the Excel file is properly saved and
closed when the block finishes — even if an error occurs. Without `with`, the file might
be left open and locked, preventing you from opening it.

### How Excel Formatting Works

```python
def apply_excel_formatting(ws, header_fill_hex="1F4E79"):
    header_fill = PatternFill(fill_type="solid", fgColor=header_fill_hex)
    header_font = Font(bold=True, color="FFFFFF", size=11)
    
    for cell in ws[1]:  # ws[1] = first row (header row)
        cell.fill = header_fill
        cell.font = header_font
```

`ws[1]` means "the first row of the worksheet."
`PatternFill(fgColor="1F4E79")` sets a solid dark navy blue background.
`Font(bold=True, color="FFFFFF")` sets bold white text.

Colours are in hex format — "1F4E79" is a specific shade of dark blue.
"FFFFFF" is pure white. "C00000" is dark red. You can look up any colour's hex code online.

`ws.freeze_panes = "A2"` freezes everything above row 2. So when you scroll down, row 1
(the header) stays visible. This is the same as View → Freeze Panes in Excel.

## Step 6 — Charts

### Chart 1: Revenue Trend (Line Chart)

```python
fig, ax = plt.subplots(figsize=(12, 6))
ax.plot(yoy_df["Year"], yoy_df["Total_Revenue"] / 1e6, marker="o")
```

`figsize=(12, 6)` → Canvas size: 12 inches wide, 6 inches tall.
`Total_Revenue / 1e6` → Divides by 1,000,000 to show values in millions.
`1e6` is scientific notation for 1,000,000 (1 × 10^6).
`marker="o"` → Places a circle (dot) at each data point on the line.

```python
ax.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x, _: f"${x:.1f}M"))
```

**What this does:** Controls how the Y-axis numbers are displayed.
Without this, the Y-axis would show `3.5` — with it, it shows `$3.5M`.

`lambda x, _:` → a mini function that takes x (the number) and _ (ignored position).
`f"${x:.1f}M"` → f-string that formats x to 1 decimal place with $ and M around it.

### Chart 2: YoY Growth Bar Chart

```python
colors = ["#C00000" if v < 0 else "#1F4E79" for v in yoy_clean["YoY_Revenue_Growth_%"]]
```

This is a list comprehension that creates a colour for each bar:
- If the growth value is negative → dark red (#C00000)
- If positive → dark blue (#1F4E79)

This is why decline years appear in red automatically. The colour array is then passed
to `ax.bar(..., color=colors)` and each bar gets its corresponding colour.

### Chart 3: Category Stacked Bar

```python
cat_pivot = (
    master_df.groupby(["Year", "Category"])["Sales"]
    .sum()
    .unstack(fill_value=0)
    / 1e6
)
```

**What `.unstack()` does:**
Before unstack, the data looks like:
```
Year    Category          Sales
2013    Furniture         450,000
2013    Office Supplies   380,000
2013    Technology        720,000
2014    Furniture         486,000
...
```

After `.unstack()`, Category values become columns:
```
Category    Furniture    Office Supplies    Technology
Year
2013        450,000      380,000            720,000
2014        486,000      ...                ...
```

This "wide" format is exactly what matplotlib needs to draw a stacked bar chart —
each column becomes one coloured segment in the stack.

## Step 7 — MySQL Load

```python
conn = mysql.connector.connect(**DB_CONFIG)
cursor = conn.cursor()
```

`mysql.connector.connect(**DB_CONFIG)` opens a connection to MySQL.
`**DB_CONFIG` unpacks the dictionary as keyword arguments. It is equivalent to writing:
`connect(host="localhost", port=3306, user="root", password="...", database="...")`

`cursor` is like a pen that can write SQL commands to the database.

```python
cursor.execute("TRUNCATE TABLE consolidated_sales")
```

`TRUNCATE` deletes all rows from the table but keeps the table structure.
This ensures re-running the pipeline does not duplicate data.

```python
cursor.executemany(insert_sql, data_tuples)
conn.commit()
```

`executemany` sends thousands of INSERT commands in one batch — much faster than a loop.
`conn.commit()` finalises (saves) the inserted data to disk. Without commit, changes are
held in memory and lost when the connection closes.

---

# 7. SQL FROM SCRATCH

SQL (Structured Query Language) is a language for asking questions to a database.
Every SQL query is a question, and the database returns a table of answers.

## The Basic Structure

```sql
SELECT  what_you_want_to_see
FROM    which_table
WHERE   condition_to_filter_rows
GROUP BY column_to_group_by
ORDER BY column_to_sort_by;
```

## SELECT and FROM

```sql
-- Show me all sales and profits from the consolidated_sales table
SELECT Sales, Profit
FROM consolidated_sales;
```

`--` is how you write a comment in SQL. The database ignores anything after `--`.
`SELECT *` means "select all columns."

## WHERE — Filtering Rows

```sql
-- Show only orders from 2020
SELECT Order_ID, Sales, Profit
FROM consolidated_sales
WHERE Year = 2020;

-- Show only Technology orders with Sales over $500
SELECT Order_ID, Sales
FROM consolidated_sales
WHERE Category = 'Technology' AND Sales > 500;
```

## GROUP BY — Aggregating (Like Pivot Tables)

```sql
-- Total revenue per year
SELECT Year, SUM(Sales) AS Total_Revenue
FROM consolidated_sales
GROUP BY Year
ORDER BY Year;
```

`SUM(Sales)` adds up all Sales values within each group.
`AS Total_Revenue` gives the result column a friendly name (alias).
`GROUP BY Year` creates one result row per year.

Common aggregation functions:
| Function | What it does |
|---|---|
| `SUM(col)` | Adds up all values |
| `AVG(col)` | Calculates the average |
| `COUNT(*)` | Counts rows |
| `COUNT(DISTINCT col)` | Counts unique values |
| `MAX(col)` | Finds the highest value |
| `MIN(col)` | Finds the lowest value |

## CTEs — Common Table Expressions

A CTE creates a temporary named result that you can reference in the same query.
Think of it as giving a subquery a name so you can use it like a table.

```sql
-- WITHOUT CTE (hard to read):
SELECT Year, Revenue, Revenue - (
    SELECT Revenue FROM (SELECT Year, SUM(Sales) AS Revenue FROM consolidated_sales GROUP BY Year) sub
    WHERE sub.Year = main.Year - 1
) AS Change
FROM (SELECT Year, SUM(Sales) AS Revenue FROM consolidated_sales GROUP BY Year) main;

-- WITH CTE (clean and readable):
WITH yearly_revenue AS (
    SELECT Year, SUM(Sales) AS Revenue
    FROM consolidated_sales
    GROUP BY Year
)
SELECT Year, Revenue
FROM yearly_revenue;
```

Same result, completely different readability. CTEs are not faster — they are for you,
the human, to read and maintain the code more easily.

You can chain multiple CTEs:
```sql
WITH
first_step AS (
    SELECT Year, SUM(Sales) AS Revenue FROM consolidated_sales GROUP BY Year
),
second_step AS (
    SELECT Year, Revenue, LAG(Revenue) OVER (ORDER BY Year) AS Prev_Revenue
    FROM first_step
)
SELECT * FROM second_step;
```

The second CTE can reference the first one. This is called "chained CTEs."

## Window Functions — The Most Powerful SQL Feature

Regular GROUP BY collapses many rows into one summary row.
Window functions add a calculated column to each row WITHOUT collapsing them.

**GROUP BY example:**
```sql
SELECT Year, SUM(Sales) AS Revenue
FROM consolidated_sales
GROUP BY Year;
-- Result: 12 rows (one per year)
```

**Window function example:**
```sql
SELECT Year, Order_ID, Sales,
       SUM(Sales) OVER (PARTITION BY Year) AS Year_Total
FROM consolidated_sales;
-- Result: 11,593 rows (all original rows PLUS the year total added to each)
```

Every individual row keeps its own data AND gets the year total added alongside it.

### OVER() — The Window Definition

`OVER()` defines the "window" — which rows the function should consider.

```sql
SUM(Sales) OVER ()                    -- considers ALL rows in the table
SUM(Sales) OVER (PARTITION BY Year)   -- considers only rows in the same year
SUM(Sales) OVER (ORDER BY Year)       -- running total ordered by year
```

`PARTITION BY` is like GROUP BY but for window functions — it divides rows into groups
without collapsing them.

### LAG() — Look at the Previous Row

```sql
SELECT Year,
       SUM(Sales) AS Revenue,
       LAG(SUM(Sales)) OVER (ORDER BY Year) AS Prev_Year_Revenue
FROM consolidated_sales
GROUP BY Year;
```

`LAG(SUM(Sales)) OVER (ORDER BY Year)` looks at the revenue of the previous year's row.

Result:
```
Year    Revenue      Prev_Year_Revenue
2013    1,200,000    NULL               ← no previous year
2014    1,296,000    1,200,000          ← 2013's revenue
2015    1,440,000    1,296,000          ← 2014's revenue
```

**Why is 2013 NULL?** There is no row before 2013 in the window, so LAG returns nothing.

**How to calculate YoY growth:**
```sql
ROUND(
    (Revenue - LAG(Revenue) OVER (ORDER BY Year))
    / LAG(Revenue) OVER (ORDER BY Year) * 100,
2) AS YoY_Growth_Pct
```

This is the same formula as in mathematics:
`Growth% = (New - Old) / Old × 100`

### RANK() — Rank Items Within Groups

```sql
SELECT Category, Product_Name, SUM(Sales) AS Revenue,
       RANK() OVER (PARTITION BY Category ORDER BY SUM(Sales) DESC) AS Rank_In_Category
FROM consolidated_sales
GROUP BY Category, Product_Name;
```

`PARTITION BY Category` → rank separately within each category.
`ORDER BY SUM(Sales) DESC` → highest revenue = rank 1.

Result:
```
Category      Product_Name    Revenue    Rank_In_Category
Technology    iPhone 14       250,000    1
Technology    MacBook Pro     220,000    2
Technology    Dell XPS 15     180,000    3
Furniture     Standing Desk   150,000    1
Furniture     Ergonomic Chair 130,000    2
```

**RANK vs ROW_NUMBER vs DENSE_RANK:**
Suppose three products have revenues: $100k, $100k, $80k

```
RANK():        1, 1, 3   ← Ties share same rank, next rank skips
DENSE_RANK():  1, 1, 2   ← Ties share same rank, next rank does NOT skip
ROW_NUMBER():  1, 2, 3   ← Every row gets unique number regardless of ties
```

Use RANK() for competitions (gold-gold-bronze). Use ROW_NUMBER() when you need unique IDs.

### Running Total — SUM() OVER with Frame

```sql
SELECT Year,
       SUM(Sales) AS Annual_Revenue,
       SUM(SUM(Sales)) OVER (ORDER BY Year ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
           AS Cumulative_Revenue
FROM consolidated_sales
GROUP BY Year;
```

`ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW` means:
"From the very first row up to and including the current row."

Result:
```
Year    Annual_Revenue    Cumulative_Revenue
2013    1,200,000         1,200,000
2014    1,296,000         2,496,000
2015    1,440,000         3,936,000
```

The cumulative column shows how much revenue the company has generated in total up to that year.

### FIRST_VALUE() and LAST_VALUE()

```sql
FIRST_VALUE(Year) OVER (ORDER BY Revenue DESC
                         ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING)
```

`FIRST_VALUE()` returns the value from the first row of the window.
`LAST_VALUE()` returns the value from the last row.

When ordered by Revenue descending, `FIRST_VALUE(Year)` gives the year with highest revenue.
`ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING` means "consider all rows"
(the entire result set). Without this, LAST_VALUE would only look at rows seen so far.

### The CASE WHEN Statement

```sql
CASE
    WHEN Growth_Pct IS NULL THEN 'Baseline'
    WHEN Growth_Pct >= 5    THEN 'Growth'
    WHEN Growth_Pct <= -5   THEN 'Decline'
    ELSE                         'Stable'
END AS Year_Classification
```

This is SQL's version of an IF-ELIF-ELSE chain. It checks conditions from top to bottom
and returns the first matching result. If no condition matches, `ELSE` is returned.

The result is a new text column called "Year_Classification" with one of four values.

---

# 8. THE CHARTS

## Chart 1 — revenue_trend.png (Line Chart)

**What it shows:** Total annual revenue from 2013 to 2024 plotted as a connected line
with a filled area below it.

**Why a line chart?** Line charts are the standard for time-series data because the
connecting line implies continuity — each year flows into the next. A bar chart would
suggest each year is independent.

**What to look for:**
- Overall upward slope = long-term business growth
- Visible dips in 2016 and 2020 = market slowdowns
- Steepness of recovery after 2020 = how quickly the business bounced back from COVID

**What you would say to a manager:** "Revenue grew consistently over 12 years. We saw
a temporary dip in 2020 due to COVID-19, but recovered strongly in 2021 with the highest
year-on-year growth in the entire period."

## Chart 2 — yoy_growth.png (Bar Chart)

**What it shows:** The percentage change in revenue compared to the previous year.
Positive years are dark blue, negative years are red.

**Why this is more useful than the trend line:** The trend line shows ABSOLUTE numbers
(how much revenue). This chart shows RATE of growth (how fast it is growing). A company
can have high absolute revenue but be slowing down. This chart reveals that instantly.

**What to look for:**
- Red bars → years where revenue actually declined
- Bars getting shorter → growth is slowing (even if still positive)
- The 2021 bar → should be the tallest, showing the COVID recovery

**What you would say to a manager:** "Our growth rate peaked in the post-COVID recovery.
However, the 2016 and 2020 bars show revenue actually declined in those years, which
aligns with the market conditions we experienced."

## Chart 3 — category_breakdown.png (Stacked Bar Chart)

**What it shows:** Annual revenue broken into three coloured segments — one per category.
The total bar height = total company revenue. Each segment = one category's contribution.

**Why stacked bars?** This chart answers TWO questions at once:
1. Is total revenue growing? (look at bar heights)
2. Is the product mix changing? (look at segment proportions)

**What to look for:**
- Which colour segment is growing fastest
- Whether Technology's share is increasing relative to Furniture
- Whether any category is shrinking as a proportion

**What you would say to a manager:** "Technology consistently contributes the most revenue
and its share has grown. Office Supplies remains stable as a proportion. This suggests our
technology product line is our primary growth driver."

---

# 9. THE EXCEL REPORTS

## yoy_trend_analysis.xlsx

One sheet with 12 rows (one per year) showing:
- Total Revenue
- Total Profit
- Profit Margin %
- Total Orders
- Average Order Value
- YoY Revenue Growth %
- YoY Profit Growth %
- Best Performing Category
- Best Performing Region

This is the executive summary — what a CEO or director would look at first.

## summary_report.xlsx (6 Sheets)

**Sheet 1 — YoY Trend:** Same as above. Year-by-year KPIs.

**Sheet 2 — By Category:** Revenue and profit for Technology, Furniture, Office Supplies
broken down by year. Answers "Which category grew the most?"

**Sheet 3 — By Region:** Revenue broken down by East, West, Central, South per year.
Answers "Are some regions growing faster than others?"

**Sheet 4 — By Segment:** Revenue by Consumer, Corporate, Home Office per year.
Answers "Which customer type is most valuable?"

**Sheet 5 — Top Products:** The 10 highest-revenue products across all 12 years.
Answers "What are our best-selling products?"

**Sheet 6 — Data Quality Log:** A record of every data issue found and fixed per file.
Answers "How messy was the raw data, and what did we fix?" This sheet is important for
trust — stakeholders want to know the data was cleaned properly.

---

# 10. THE JUPYTER NOTEBOOK

## What Is Jupyter?

Jupyter Notebook is an interactive coding environment that runs in your web browser.
Instead of running an entire script at once, you run code in small blocks called "cells"
and see the output immediately below each cell.

It is perfect for exploration and learning because you can:
- Run one block of code at a time
- See a chart appear immediately below the code that created it
- Write notes and explanations alongside your code (Markdown cells)
- Change a value and re-run just that cell

## How to Run It

```bash
jupyter notebook notebooks/exploration.ipynb
```

This opens a tab in your browser. Each cell shows code. Press **Shift + Enter** to run
the selected cell and move to the next one.

## The 8 Sections

**Section 1:** Load the consolidated CSV and preview the first 10 rows. Verify the data
loaded correctly before doing any analysis.

**Section 2:** Check the shape (rows × columns), data types, and which columns have
missing values. This is always the second step in any data analysis — understand what
you have before you touch it.

**Section 3:** Draw the revenue trend line chart. Same as the pipeline's chart but
inside the notebook where you can see the code alongside the chart.

**Section 4:** Bar charts showing revenue and profit margin by category. Also a stacked
bar showing category mix per year.

**Section 5:** A heatmap — a colour-coded grid showing revenue by Region (rows) and Year
(columns). Dark blue = high revenue. Light blue = low revenue. You can instantly spot
which region is strongest and whether any region is declining.

**Section 6:** A histogram showing how profit margins are distributed across all orders.
Also a box plot showing the distribution by category. You can see whether losses are common,
and whether any category has a wide spread of margins.

**Section 7:** A correlation matrix — a grid showing how strongly each pair of numeric
columns is related. A correlation of +1 means they move together perfectly. -1 means one
goes up when the other goes down. 0 means no relationship.

**Section 8:** A plain text summary of all key findings — written as if presenting to a
manager. Numbers pulled automatically from the data.

---

# 11. COMMON ERRORS AND HOW TO FIX THEM

## Error: "is not recognized as an internal or external command"

**What it means:** You typed a file path into the terminal instead of a command.

**What happened:** You typed the path to the folder instead of using `cd` first.

**Fix:**
```bash
# Wrong:
C:\Users\aman9\OneDrive\Desktop\Cowork For Data\...

# Correct:
cd "C:\Users\aman9\OneDrive\Desktop\Cowork For Data\Projects\Multi-Year Sales Consolidation (2013-2024)"
python generate_data.py
```

Always `cd` to the folder first, then run Python commands.

## Error: "ModuleNotFoundError: No module named 'faker'"

**What it means:** A required library is not installed.

**Fix:**
```bash
pip install -r requirements.txt
```

Run this from inside the project folder.

## Error: "FileNotFoundError: data/raw/sales_2013.xlsx"

**What it means:** You ran `pipeline.py` before running `generate_data.py`.

**Fix:** Always run in this order:
```bash
python generate_data.py   ← First
python pipeline.py        ← Second
```

## Error: "PermissionError: [Errno 13] Permission denied"

**What it means:** The output Excel file is currently open in Excel. Python cannot
write to a file that is already open.

**Fix:** Close the Excel file first, then re-run the pipeline.

## Error: "TypeError: Invalid value for dtype 'str'"

**What it means:** You are using pandas 3.0+ which enforces strict data types.
Datetime objects cannot be stored in a string column.

**Fix:** This was already fixed in `generate_data.py` — the dates are now stored as
differently formatted strings instead of datetime objects.

## Error: MySQL connection refused

**What it means:** MySQL is not running, or the credentials in `config.py` are wrong.

**Fix:**
1. Open `config.py`
2. Update `"user"` and `"password"` to match your MySQL login
3. Make sure MySQL Server is running (check Windows Services or MySQL Workbench)
4. Re-run `python pipeline.py`

The MySQL step is optional for the rest of the project. If it fails, Steps 1–6 still
complete successfully and you still get all the Excel and chart outputs.

---

# 12. LEARNING CHECKPOINT ANSWERS

**Q: What is the difference between pandas and openpyxl?**

Pandas is for data manipulation — filtering, grouping, merging, calculating. It handles
the data itself. openpyxl is for Excel file formatting — colours, fonts, borders, column
widths. Pandas writes the numbers; openpyxl makes them look professional. In this project,
pandas uses openpyxl as its engine when saving `.xlsx` files, and we also call openpyxl
directly to apply the bold headers and coloured rows.

**Q: Why use `np.nan` instead of Python's `None`?**

`None` is Python's general-purpose "nothing" value. `np.nan` is specifically designed for
numeric contexts. When a pandas column contains `None`, pandas may convert the column to
"object" dtype which is slow and incompatible with math operations. When it contains
`np.nan`, the column stays as float dtype, which supports `.mean()`, `.median()`, `.fillna()`,
and `.isna()` correctly and efficiently.

**Q: Why use median instead of mean for missing Shipping_Cost?**

Shipping costs have outliers — most orders cost $5–$15 to ship, but occasionally a large
piece of furniture ships for $300. The mean gets pulled toward those rare high values and
overstates what "typical" shipping costs. The median — the middle value — is not affected
by outliers and gives a fair representation of the typical cost.

**Q: What does the `*` wildcard in glob do?**

It matches any sequence of characters. `sales_*.xlsx` matches any filename that starts with
"sales_", ends with ".xlsx", and has anything in between — including `sales_2013.xlsx`,
`sales_2024.xlsx`, `sales_q1.xlsx`. It would NOT match `revenue_2013.xlsx` (wrong start)
or `sales_2013.csv` (wrong ending).

**Q: What is an idempotent operation?**

An operation that produces the same result whether you run it once or ten times. In our
pipeline, `TRUNCATE TABLE` before inserting data makes Step 7 idempotent — running the
pipeline twice does not double the data in MySQL. Idempotency is essential for production
pipelines because they often need to be re-run after failures.

**Q: What does `LAG()` return for the first row?**

NULL. There is no row before the first row, so LAG has nothing to look back at.
This is why 2013 always shows NULL for YoY growth — there is no 2012 data to compare against.

**Q: What is the difference between RANK() and ROW_NUMBER()?**

If three products have the same revenue:
- `RANK()` gives them all rank 1, then jumps to rank 4 (skips 2 and 3)
- `ROW_NUMBER()` gives them 1, 2, 3 — every row gets a unique number regardless of ties

Use RANK() when you want ties to share a position (like sports rankings).
Use ROW_NUMBER() when you need a unique identifier for every row.

**Q: Why is a line chart used for time-series instead of a bar chart?**

A line chart implies continuity — the connecting line between points suggests the values
flow from one period to the next, which is true for time-series data. Bar charts imply
independence — each bar stands alone. Revenue did not suddenly appear in 2014; it grew
from 2013. The line visually communicates that progression. Bar charts are better for
comparing discrete, unrelated categories.

---

# 13. HOW TO EXTEND THIS PROJECT

## Add a New Year's File

Drop `sales_2025.xlsx` into `data/raw/` and run:
```bash
python pipeline.py
```

Everything — the CSV, the Excel reports, the charts, the MySQL table — updates automatically.
No code changes needed. This is why the glob pattern approach was used in Step 1.

## Add a New Metric to the YoY Report

In `pipeline.py`, find Step 4's `.agg()` block and add a new line:

```python
yoy = (
    master_df.groupby("Year")
    .agg(
        Total_Revenue  = ("Sales",    "sum"),
        Total_Profit   = ("Profit",   "sum"),
        Total_Orders   = ("Order_ID", "nunique"),
        Avg_Discount   = ("Discount", "mean"),    ← ADD THIS LINE
    )
    .reset_index()
)
```

Re-run the pipeline and the new column appears in `yoy_trend_analysis.xlsx` automatically.

## Add a New SQL Query

Open `sql/03_analysis_queries.sql` in any text editor and add your query at the bottom.
Try writing one that finds the top 5 regions by profit margin:

```sql
-- My custom query: Top regions by profit margin
SELECT
    Region,
    ROUND(SUM(Profit) / SUM(Sales) * 100, 2) AS Profit_Margin_Pct
FROM consolidated_sales
GROUP BY Region
ORDER BY Profit_Margin_Pct DESC;
```

## Add a New Excel Sheet to the Summary Report

In `step5_summary_report()`, add a new aggregation and sheet:

```python
# New: By Ship Mode
by_ship = master_df.groupby("Ship_Mode").agg(
    Revenue = ("Sales", "sum"),
    Orders  = ("Order_ID", "count")
).reset_index()

# Then inside the ExcelWriter block:
by_ship.to_excel(writer, index=False, sheet_name="By Ship Mode")
apply_excel_formatting(writer.sheets["By Ship Mode"])
```

---

# 14. INTERVIEW QUESTIONS AND STRONG ANSWERS

**Q: "Walk me through this project."**

> "I built an automated ETL pipeline in Python that consolidates 12 years of annual Excel
> sales files into a single master dataset. The pipeline auto-detects files, cleans each
> year's data — fixing casing inconsistencies, imputing missing values with strategies
> appropriate to each column, removing duplicates — then merges everything into one CSV.
> It then generates Year-over-Year trend analysis, a 6-sheet Excel summary, three charts,
> and loads everything into MySQL. The key achievement: what took 3 days manually now runs
> in under 10 minutes with one command."

**Q: "Why did you fill missing Discount with 0 but missing Shipping_Cost with the median?"**

> "Different columns have different business logic. A missing Discount almost certainly
> means no discount was applied — the customer paid full price. Replacing with 0 is accurate.
> A missing Shipping_Cost is genuinely unknown — we cannot assume it was zero because
> shipping always has a cost. We use the median rather than the mean because shipping costs
> have outliers — one large furniture shipment could cost hundreds of dollars and would
> pull the mean upward, making it unrepresentative of typical costs."

**Q: "What is a window function? Give me an example."**

> "A window function performs a calculation across related rows without collapsing the result
> into a group. GROUP BY aggregates 11,000 rows into 12 year-rows. A window function keeps
> all 11,000 rows and adds a new calculated column alongside each. For example, LAG() is a
> window function that looks back at the previous row. I used it to calculate YoY revenue
> growth: `(Revenue - LAG(Revenue) OVER (ORDER BY Year)) / LAG(Revenue) * 100`. For each
> year row, LAG provides the previous year's revenue, allowing a simple subtraction."

**Q: "How does your pipeline handle a new yearly file?"**

> "It uses Python's `glob` module with a wildcard pattern: `glob.glob('data/raw/sales_*.xlsx')`.
> The `*` matches any text, so when `sales_2025.xlsx` is dropped into the raw folder, it is
> automatically detected on the next run. No code changes are required — that was a deliberate
> design decision to make the pipeline reusable for years to come."

**Q: "What would you do differently if this were a production system?"**

> "Several things. First, I would replace print statements with proper logging to a log file
> using Python's logging module — so errors are recorded even if no one is watching the terminal.
> Second, I would store credentials in environment variables, not a config file — config files
> should never be committed to version control with passwords. Third, I would add unit tests
> for each cleaning function using pytest. Fourth, I would containerise with Docker so the
> pipeline runs identically on any machine. And fifth, I would schedule it with Apache Airflow
> for automatic daily runs."

**Q: "What is the difference between median and mean? Why does it matter for data cleaning?"**

> "The mean is the sum divided by count — it is sensitive to extreme values. The median is
> the middle value when sorted — it is resistant to extremes. In data cleaning, columns like
> Shipping_Cost or Salary often have outliers. Using the mean to fill missing values would
> overestimate what 'typical' looks like. The median gives a fairer, more robust imputation
> that does not get distorted by the one $500 shipment in a dataset of mostly $10 shipments."

---

# 15. HOW TO PRESENT THIS PROJECT

## On Your Resume

```
Multi-Year Sales Consolidation Pipeline  |  Python · MySQL · Excel · pandas
• Built an automated ETL pipeline consolidating 12 years of annual Excel sales data
  into a single master dataset (11,593 rows)
• Reduced manual consolidation time from 3 days to under 10 minutes
• Implemented 6-category data quality cleaning (casing, nulls, duplicates, date formats)
• Generated automated YoY trend reports, 6-sheet Excel summaries, and 3 PNG charts
• Authored 15 analytical SQL queries using window functions, CTEs, LAG(), and RANK()
```

## In an Interview — The 60-Second Version

Start with the problem, not the code:

> "Before this project, analysts spent 3 working days every quarter manually combining
> 12 Excel files. Any mistake in the copy-paste process would go undetected. I built a
> Python pipeline that does the same thing automatically in under 10 minutes — cleaning
> the data, combining it, generating trend analysis, producing charts, and loading it into
> MySQL for SQL querying. The pipeline is also reusable — adding next year's file requires
> no code changes."

Then wait. The interviewer will ask about the part they care about. Answer only what they ask.

## On GitHub

1. Make sure `README.md` is clean and complete — it is your GitHub landing page.
2. Add the 3 chart images directly to the README so visitors can see the output.
3. Do not commit your MySQL password to GitHub. Either delete `config.py` or add it to
   `.gitignore`.
4. Include `requirements.txt` — if someone cannot install and run your project, they
   will not bother reviewing it.

## What to Highlight vs. What to Downplay

**Highlight:**
- The business impact (3 days → 10 minutes)
- The cleaning decisions and the reasoning behind each (this shows analytical thinking)
- The reusable design (no code changes for new years)
- The SQL window functions (shows advanced SQL knowledge)

**Downplay:**
- `generate_data.py` — mention it exists, explain it is synthetic data, move on quickly.
  It is scaffolding, not the actual work.
- The exact number of rows — it varies each run. Say "approximately 11,000–12,000 rows."
- The chart aesthetics — interviewers care about the insight, not the colour scheme.
