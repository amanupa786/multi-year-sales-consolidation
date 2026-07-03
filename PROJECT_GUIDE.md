# PROJECT GUIDE — Multi-Year Sales Consolidation (2013–2024)

> **What this document is:** A step-by-step learning guide explaining every concept, decision,
> and tool in this project. Read it from top to bottom before running any code.

---

## Table of Contents

1. [What Problem Does This Project Solve?](#1-what-problem-does-this-project-solve)
2. [Project Architecture Overview](#2-project-architecture-overview)
3. [Python Libraries — What They Do and Why We Use Them](#3-python-libraries)
4. [Step-by-Step Pipeline Walkthrough](#4-pipeline-walkthrough)
5. [Data Cleaning — Why Each Step Matters](#5-data-cleaning)
6. [SQL Concepts Explained](#6-sql-concepts)
7. [Charts — What They Show and Their Business Insight](#7-charts)
8. [Common Errors and How to Fix Them](#8-common-errors)
9. [How to Extend This Project](#9-extending-the-project)
10. [Learning Checkpoints](#10-learning-checkpoints)

---

## 1. What Problem Does This Project Solve?

### The Business Problem

Imagine you are a data analyst at a retail company. Every year, your sales team saves its annual performance data into a separate Excel file. After 12 years, you now have 12 files — `sales_2013.xlsx` through `sales_2024.xlsx`.

Every time your manager asks "how did our revenue grow over the last decade?", you have to:

1. Manually open all 12 files
2. Copy-paste data into a master sheet
3. Fix formatting inconsistencies (TECHNOLOGY vs technology vs Technology)
4. Remove duplicates
5. Re-create charts
6. Build a summary table

This process takes **3 full working days** and is error-prone. One wrong copy-paste can invalidate your entire trend analysis.

### The Solution

This project replaces that 3-day manual process with a **reusable automated pipeline** that:

- Reads all 12 files automatically
- Cleans and standardises every inconsistency
- Produces a single master dataset
- Generates charts, trend reports, and a multi-sheet Excel report
- Loads data into MySQL for SQL-based analysis
- Can be re-run with one command when a new year's file arrives

**Key Achievement: Reduced manual consolidation time from 3 days to under 10 minutes.**

---

## 2. Project Architecture Overview

```
generate_data.py ──► data/raw/*.xlsx
                                │
                                ▼
                          pipeline.py
                         ┌──────────────────────────────────────┐
                         │  Step 1: Load all .xlsx files        │
                         │  Step 2: Clean each file             │
                         │  Step 3: Merge → consolidated_sales  │
                         │  Step 4: YoY trend analysis          │
                         │  Step 5: Multi-sheet summary report  │
                         │  Step 6: Generate 3 charts (PNG)     │
                         │  Step 7: Load into MySQL             │
                         └──────────────────────────────────────┘
                                │
                         ┌──────┼──────────────────────┐
                         ▼      ▼                      ▼
                    CSV file  Excel reports        MySQL DB
                                │                      │
                                └──────────────────────┘
                                          │
                                   Jupyter Notebook
                                   (EDA & insights)
```

The pipeline is **linear** — each step depends on the output of the previous one. This is called a **sequential ETL pipeline** (Extract, Transform, Load).

---

## 3. Python Libraries

### pandas
The backbone of the project. Pandas provides the `DataFrame` — a table structure for Python that works like Excel but is programmable.

- `pd.read_excel()` — reads .xlsx files into a DataFrame
- `df.groupby()` — aggregates rows (like a Pivot Table)
- `df.merge()` — joins two DataFrames (like VLOOKUP/SQL JOIN)
- `pd.concat()` — stacks DataFrames on top of each other (like Copy-Paste in Excel)

**Why pandas over pure Python?** Pure Python loops over millions of rows slowly. Pandas uses NumPy under the hood, performing operations on entire columns at once (vectorised operations), which is 10–100× faster.

### numpy
NumPy provides fast mathematical operations on arrays. We use it for:
- `np.nan` — Python's representation of a missing value (NaN = Not a Number)
- `np.where()` — a vectorised if-else (like Excel's IF formula)
- `np.zeros()` — create arrays of zeros for chart stacking
- Random number generation in `generate_data.py`

### openpyxl
A library for reading and writing Excel `.xlsx` files. Pandas uses it as its Excel "engine". We also use it directly to:
- Apply bold headers and coloured fills (`PatternFill`, `Font`)
- Freeze the header row (`freeze_panes`)
- Auto-fit column widths

### faker
Generates realistic fake data — names, addresses, dates. We use `fake.name()` to generate customer names that look real without using actual personal data.

### matplotlib
Python's core plotting library. Every chart ultimately comes from matplotlib, even when using seaborn. We use it to:
- Create figure and axes objects (`fig, ax = plt.subplots()`)
- Control axis labels, titles, tick formatters
- Save charts to PNG files (`plt.savefig()`)

### seaborn
A higher-level library built on top of matplotlib. It provides:
- `sns.heatmap()` — colour-coded grid charts
- `sns.boxplot()` — distribution comparison by category
- `sns.set_theme()` — one-line styling for professional charts

### mysql-connector-python
The official MySQL driver for Python. It creates a direct connection to MySQL without needing SQLAlchemy. We use:
- `mysql.connector.connect()` — open a connection
- `cursor.execute()` — run a single SQL statement
- `cursor.executemany()` — insert thousands of rows efficiently in one call

---

## 4. Pipeline Walkthrough

### Step 1 — Load All Files

**What it does:** Uses Python's `glob` module to find all files matching the pattern `data/raw/sales_*.xlsx`. The `*` is a wildcard that matches any text — so it finds `sales_2013.xlsx`, `sales_2014.xlsx`, all the way to `sales_2024.xlsx`.

For each file, it:
1. Extracts the year from the filename using string manipulation
2. Loads the file using `pd.read_excel()`
3. Adds a `Year` column so we know which year each row belongs to after merging

**Why this approach?** Hard-coding 12 `pd.read_excel()` calls would work, but the pipeline would break when you add `sales_2025.xlsx`. The glob pattern approach means adding a new yearly file automatically includes it in the next pipeline run — zero code changes needed.

### Step 2 — Data Cleaning

This is the most important step. Raw data is almost never ready for analysis. See [Section 5](#5-data-cleaning) for a detailed explanation of each cleaning operation.

### Step 3 — Consolidate

**What it does:** Uses `pd.concat()` to stack all 12 cleaned DataFrames into one master DataFrame. Then saves it to `data/processed/consolidated_sales.csv`.

**Why CSV?** CSV is the most universally compatible data format. It can be opened by Python, Excel, Power BI, Tableau, MySQL, and virtually any other tool. The `data/processed/` folder is our single source of truth — every downstream tool reads from here.

### Step 4 — YoY Analysis

**What it does:** Groups the master DataFrame by Year and calculates:
- Total Revenue (sum of Sales column)
- Total Profit (sum of Profit column)
- Profit Margin % (Profit ÷ Revenue × 100)
- Total Orders (count of unique Order IDs)
- Average Order Value (Revenue ÷ Orders)
- YoY Growth % (using pandas `.shift()` to compare each year to the previous)

**Key technique — `.shift(1)`:** When you call `.shift(1)` on a column, each value shifts down by one row. So row 2 now contains the value that was in row 1 (the previous year). This lets you calculate change between consecutive years with simple subtraction.

### Step 5 — Summary Report

**What it does:** Creates a 6-sheet Excel file covering the same data from different angles — by category, region, segment, and product. Also includes a Data Quality Log showing exactly what issues were found and fixed per file.

**Why multiple sheets?** Different stakeholders care about different dimensions. The CEO wants the YoY trend. The Category Manager wants the category breakdown. The Operations team wants the region breakdown. One Excel file, different sheets for different audiences.

### Step 6 — Charts

**What it does:** Generates 3 PNG files:
1. A line chart showing revenue growth over time
2. A bar chart showing YoY growth % (positive = blue, negative = red)
3. A stacked bar chart showing how each category contributes to yearly revenue

See [Section 7](#7-charts) for the business insight behind each chart.

### Step 7 — MySQL Load

**What it does:** Connects to MySQL, creates the `consolidated_sales` table if it doesn't exist, then inserts the entire cleaned dataset. Uses batch inserts (1,000 rows at a time) for performance.

**Why TRUNCATE before INSERT?** This makes the pipeline **idempotent** — running it twice produces the same result as running it once. Without TRUNCATE, re-running would duplicate every row in the database.

---

## 5. Data Cleaning

### 5a. Standardise Column Names
**Why:** Different Excel files may have columns named `order_id`, `Order ID`, or `OrderID`. By stripping whitespace and applying title case, all 12 files use identical column names before merging.

### 5b. Fix Text Casing
**Problem:** `generate_data.py` intentionally introduces `"TECHNOLOGY"`, `"technology"`, and `"Technology"` in the Category column.  
**Fix:** `.str.strip().str.title()` — removes leading/trailing spaces and converts to Title Case.  
**Why it matters:** If you group by Category with mixed casing, you'll get 3 separate groups instead of 1. Your revenue totals will be wrong.

### 5c. Convert Dates
**Problem:** Some Order_Date values are strings (`"2020-03-15"`), others are Python date objects. Mixing types causes errors in date arithmetic and chart plotting.  
**Fix:** `pd.to_datetime(df["Order_Date"], errors="coerce")` — converts everything to a consistent datetime type. The `errors="coerce"` argument turns any unparseable values into `NaT` (Not a Time) instead of crashing.

### 5d. Fill Missing Discount
**Business logic:** A missing discount almost certainly means no discount was applied (0%), not that the discount is unknown.  
**Fix:** `df["Discount"].fillna(0)` — safe assumption, preserves data integrity.

### 5e. Fill Missing Shipping Cost
**Business logic:** The true value is unknown, but we can't remove rows just because shipping cost is missing.  
**Fix:** Replace with the median of that year's shipping costs. We use the **median** rather than the **mean** because the mean is sensitive to outliers (one $500 shipment would skew the mean upward).

### 5f. Fill Missing Region
**Business logic:** Unknown regions should be assigned the most common region for that file.  
**Fix:** `df["Region"].mode()[0]` — the `.mode()` method returns the most frequent value.

### 5g. Remove Duplicates
**Problem:** `generate_data.py` adds 2–3 duplicate rows per file, simulating real-world data entry mistakes.  
**Fix:** `df.drop_duplicates()` — removes rows that are identical across all columns.  
**Why not remove partial duplicates?** An Order_ID might appear twice for different products — that's a legitimate multi-item order. Full-row duplicates are the problem.

### 5h. Validate Sales Formula
**Business logic:** Sales should equal `Quantity × Unit_Price × (1 – Discount)`. If the actual Sales value differs by more than 1% from this formula, it may indicate data entry errors.  
**Action:** We flag mismatches in the Data Quality Log but do NOT remove them — the Sales column is the official record; the formula check is a warning, not a filter.

### 5i. Add Profit_Margin_%
**Formula:** `(Profit / Sales) × 100`  
**Business insight:** Raw profit numbers are hard to compare across different-sized orders. A $500 profit on a $10,000 sale (5% margin) is worse than a $50 profit on a $200 sale (25% margin). The margin percentage makes comparison fair.

### 5j. Add Revenue_Band
**Formula:** Uses `pd.cut()` to bin Sales into Low (<$500), Medium ($500–$2,000), High (>$2,000).  
**Business insight:** Allows quick segmentation of orders by deal size. High-value orders may warrant priority shipping and dedicated account management.

---

## 6. SQL Concepts

### CTEs — Common Table Expressions
A CTE (written as `WITH name AS (SELECT ...)`) creates a temporary, named result set that you can reference in the same query. Think of it as giving a subquery a name so you can read it more easily.

**Without CTE:**
```sql
SELECT Year, Revenue, Revenue - (SELECT Revenue FROM ...) AS Change FROM ...
```

**With CTE:**
```sql
WITH yearly_rev AS (SELECT Year, SUM(Sales) AS Revenue FROM ... GROUP BY Year)
SELECT Year, Revenue, Revenue - LAG(Revenue) OVER (ORDER BY Year) AS Change
FROM yearly_rev;
```
CTEs make complex queries readable and maintainable.

### Window Functions — OVER()
A window function performs a calculation across a "window" of rows related to the current row, without collapsing the result into a group.

**GROUP BY** collapses many rows into one summary row.  
**Window functions** keep all rows and add a calculated column alongside them.

### LAG() — Look Backward
`LAG(column, n)` returns the value from `n` rows earlier in the window order. Used in Query C1 to get the previous year's revenue so we can calculate growth rate.

```sql
LAG(SUM(Sales)) OVER (ORDER BY Year)
```
For Year 2014, this returns the 2013 revenue. For 2015, it returns 2014. For 2013 (the first row), it returns NULL.

### RANK() — Position Within a Group
`RANK() OVER (PARTITION BY category ORDER BY revenue DESC)` assigns a rank number within each partition (group). Products ranked #1 have the highest revenue in their category.

The difference from `ROW_NUMBER()`: if two products have the same revenue, they both get rank 1, and the next rank is 3 (not 2). This is the standard competition ranking (gold-gold-bronze).

### FIRST_VALUE() and LAST_VALUE()
These return the first or last value in a window frame. Used in Query C4 to find the best and worst year. The frame `ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING` means "consider all rows in the partition".

### Running Total — SUM() OVER()
```sql
SUM(Revenue) OVER (ORDER BY Year ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
```
This adds up revenue from the first year up to the current year — a cumulative (running) total. Useful for showing how total business has grown cumulatively.

---

## 7. Charts

### Chart 1: revenue_trend.png — Line Chart
**What it shows:** Total annual revenue from 2013 to 2024 as a connected line with data point labels.

**Business insight:** The upward trajectory confirms long-term business growth. Visible dips in 2016 and 2020 mark market slowdowns. The recovery steepness after 2020 shows the business bounced back from COVID stronger than before. Stakeholders use this chart in board presentations to demonstrate year-on-year momentum.

**Chart choice rationale:** A line chart is the standard choice for time-series data because it emphasises continuity and trend direction. Bar charts are better for comparing discrete categories; line charts are better for showing change over time.

### Chart 2: yoy_growth.png — Bar Chart
**What it shows:** The percentage change in revenue compared to the previous year, with negative growth bars shown in red.

**Business insight:** This chart is more informative than the trend line for understanding *pace* of growth. A company growing at 15% YoY is doing better than one growing at 5%, even if the 5%-growth company has higher absolute revenue. Red bars immediately flag years of concern — a senior leader can identify 2020 at a glance without reading numbers.

**Chart choice rationale:** Bar charts are ideal for comparing discrete values (one per year). Color-coding positive vs. negative makes the insight instantaneous.

### Chart 3: category_breakdown.png — Stacked Bar Chart
**What it shows:** Revenue broken down by category (Technology, Furniture, Office Supplies) for each year, stacked so the total bar height represents total revenue.

**Business insight:** Shows both total growth AND category mix shift. If Technology's share grows over time while Furniture shrinks, that tells the product team where to invest. If Office Supplies remains a flat percentage, that's a mature, stable category. Portfolio composition matters as much as total revenue.

**Chart choice rationale:** Stacked bars are the standard for part-to-whole comparisons over time. They combine a trend view (bar height) with a composition view (segment proportions).

---

## 8. Common Errors

### Error: `ModuleNotFoundError: No module named 'faker'`
**Cause:** The required library is not installed.  
**Fix:** Run `pip install -r requirements.txt` from the project root.

### Error: `FileNotFoundError: data/raw/sales_2013.xlsx`
**Cause:** You ran `pipeline.py` before `generate_data.py`.  
**Fix:** Always run `python generate_data.py` first. It takes ~10 seconds.

### Error: `PermissionError: [Errno 13] Permission denied`
**Cause:** The Excel output file is open in Excel while Python is trying to write to it.  
**Fix:** Close the file in Excel, then re-run the pipeline.

### Error: `mysql.connector.errors.ProgrammingError: Table doesn't exist`
**Cause:** The database or table was not created before inserting.  
**Fix:** The pipeline handles this automatically (CREATE TABLE IF NOT EXISTS). If you see this error, check that your `config.py` credentials are correct and MySQL is running.

### Error: `ValueError: Excel file format cannot be determined`
**Cause:** A file in the `data/raw/` folder is not a valid .xlsx file (might be a CSV or corrupted file).  
**Fix:** Check the `data/raw/` folder. Only `.xlsx` files generated by `generate_data.py` should be there.

### Error: `KeyError: 'Sales'` during cleaning
**Cause:** Column names in the raw file don't match expected names after standardisation.  
**Fix:** Run Step 2a manually for that file to print `df.columns` and see what names were detected.

### Warning: `DtypeWarning: Columns have mixed types`
**Cause:** Some Order_Date values are strings and others are date objects (intentional in this project).  
**Fix:** The cleaning step handles this with `pd.to_datetime(..., errors='coerce')`. The warning is harmless here.

---

## 9. Extending This Project

### Add a New Yearly File
1. Drop `sales_2025.xlsx` into `data/raw/`
2. Run `python pipeline.py`
3. All reports, charts, and the MySQL table are automatically updated — no code changes needed.

### Add a New KPI (e.g., Return Rate)
1. Add a `Return_Rate` column to `generate_data.py`
2. In `step2_clean()`, add validation logic for the new column
3. In `step4_yoy_analysis()`, add `Return_Rate=("Return_Rate", "mean")` to the groupby aggregation
4. The summary report and YoY Excel will automatically include the new column

### Add a New Analysis Sheet
In `step5_summary_report()`, add a new DataFrame aggregation and `.to_excel(writer, sheet_name="New Sheet")` call. Then call `apply_excel_formatting()` on the new sheet.

### Connect to a Different Database (PostgreSQL, Snowflake)
Replace `import mysql.connector` with `import psycopg2` or `import snowflake.connector`. The SQL syntax in the CREATE TABLE statement may need minor dialect adjustments (e.g., `SERIAL` instead of `AUTO_INCREMENT` in PostgreSQL).

### Schedule Daily Runs
Add a cron job (Linux/Mac) or Task Scheduler entry (Windows) to run `python pipeline.py` automatically. This is how production ETL pipelines work.

---

## 10. Learning Checkpoints

After reading each section, answer these questions. Check your answers against `SOLUTION_GUIDE.md`.

### Section 3 — Libraries
1. What is the difference between pandas and openpyxl? When would you use one over the other?
2. Why do we use `np.nan` instead of Python's `None` for missing values in a numeric column?
3. What advantage does `cursor.executemany()` have over calling `cursor.execute()` in a loop?

### Section 4 — Pipeline
4. What does the `*` wildcard in `glob.glob("data/raw/sales_*.xlsx")` match?
5. In Step 3, why do we sort by `["Year", "Order_Date"]` before saving the CSV?
6. What is an idempotent operation? Why does it matter for a data pipeline?

### Section 5 — Cleaning
7. Why do we use the **median** to fill missing Shipping_Cost instead of the mean?
8. What would happen to a GROUP BY query if Category contained both "Technology" and "TECHNOLOGY"?
9. What is the difference between `df.drop_duplicates()` and removing rows where Order_ID is duplicated?

### Section 6 — SQL
10. What does `LAG(column, 1)` return for the very first row in a window? Why?
11. What is the difference between `RANK()` and `ROW_NUMBER()`? Give an example.
12. Why must you write `ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW` in a running total window?

### Section 7 — Charts
13. Why is a line chart preferred over a bar chart for a time-series like annual revenue?
14. In the stacked bar chart, what does the total height of each bar represent?
15. Why are declining years shown in red in the YoY growth chart?
