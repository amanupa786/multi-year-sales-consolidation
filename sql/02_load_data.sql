-- ============================================================
-- 02_load_data.sql
-- Multi-Year Sales Consolidation (2013–2024)
-- ============================================================
-- PURPOSE : Load the consolidated CSV into the MySQL table
--           using LOAD DATA LOCAL INFILE for fast bulk insert.
--
-- NOTE    : The pipeline.py script (Step 7) handles loading
--           automatically via Python. Use this file as an
--           alternative for manual or scheduled loads.
--
-- PRE-REQ : Run 01_create_tables.sql first.
--           Enable local_infile:
--             SET GLOBAL local_infile = 1;
--
-- COMMAND : mysql --local-infile=1 -u root -p sales_analytics
--               < sql/02_load_data.sql
-- ============================================================

USE sales_analytics;

-- Enable local file loading for this session
SET GLOBAL local_infile = 1;

-- Truncate before loading to ensure idempotent re-runs
TRUNCATE TABLE consolidated_sales;

-- Bulk load from CSV
-- Update the file path below to match your system's absolute path
LOAD DATA LOCAL INFILE 'data/processed/consolidated_sales.csv'
INTO TABLE consolidated_sales
FIELDS TERMINATED BY ','
ENCLOSED BY '"'
LINES TERMINATED BY '\n'
IGNORE 1 ROWS                -- Skip the header row
(
    Order_ID,
    Order_Date,
    Customer_Name,
    Customer_Segment,
    Region,
    Country,
    Category,
    Sub_Category,
    Product_Name,
    Quantity,
    Unit_Price,
    Discount,
    Sales,
    Profit,
    Shipping_Cost,
    Ship_Mode,
    Year,
    Profit_Margin_Pct,
    Revenue_Band
);

-- Confirm row count after load
SELECT
    COUNT(*)           AS Total_Rows_Loaded,
    MIN(Year)          AS First_Year,
    MAX(Year)          AS Last_Year,
    MIN(Order_Date)    AS Earliest_Order,
    MAX(Order_Date)    AS Latest_Order
FROM consolidated_sales;
