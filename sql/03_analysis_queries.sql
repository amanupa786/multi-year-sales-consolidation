-- ============================================================
-- 03_analysis_queries.sql
-- Multi-Year Sales Consolidation (2013–2024)
-- ============================================================
-- PURPOSE : Comprehensive SQL analysis covering basic,
--           intermediate, and advanced analytical patterns.
-- DATABASE: sales_analytics
-- ENGINE  : MySQL 8.0+  (uses window functions)
-- ============================================================

USE sales_analytics;


-- ============================================================
-- SECTION A — BASIC QUERIES
-- ============================================================

-- A1. Overall summary: total revenue, profit, and orders across all years
SELECT
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    ROUND(SUM(Profit) / SUM(Sales) * 100, 2) AS Overall_Profit_Margin_Pct,
    COUNT(DISTINCT Order_ID)        AS Total_Orders,
    COUNT(DISTINCT Customer_Name)   AS Unique_Customers
FROM consolidated_sales;


-- A2. Revenue and profit by year — fundamental trend view
SELECT
    Year,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    ROUND(SUM(Profit) / SUM(Sales) * 100, 2) AS Profit_Margin_Pct,
    COUNT(DISTINCT Order_ID)        AS Total_Orders
FROM consolidated_sales
GROUP BY Year
ORDER BY Year;


-- A3. Revenue and profit by product category
SELECT
    Category,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    ROUND(SUM(Profit) / SUM(Sales) * 100, 2) AS Profit_Margin_Pct,
    COUNT(DISTINCT Order_ID)        AS Total_Orders
FROM consolidated_sales
GROUP BY Category
ORDER BY Total_Revenue DESC;


-- A4. Revenue by geographic region
SELECT
    Region,
    ROUND(SUM(Sales),  2)                     AS Total_Revenue,
    ROUND(SUM(Profit), 2)                     AS Total_Profit,
    ROUND(SUM(Sales) / (SELECT SUM(Sales) FROM consolidated_sales) * 100, 2) AS Revenue_Share_Pct
FROM consolidated_sales
GROUP BY Region
ORDER BY Total_Revenue DESC;


-- A5. Revenue and orders by customer segment
SELECT
    Customer_Segment,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    COUNT(DISTINCT Order_ID)        AS Total_Orders,
    ROUND(AVG(Sales), 2)            AS Avg_Order_Value
FROM consolidated_sales
GROUP BY Customer_Segment
ORDER BY Total_Revenue DESC;


-- ============================================================
-- SECTION B — INTERMEDIATE QUERIES
-- ============================================================

-- B1. Top 10 customers by total revenue
SELECT
    Customer_Name,
    Customer_Segment,
    Region,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    COUNT(DISTINCT Order_ID)        AS Total_Orders
FROM consolidated_sales
GROUP BY Customer_Name, Customer_Segment, Region
ORDER BY Total_Revenue DESC
LIMIT 10;


-- B2. Monthly revenue trend for each year
--     Shows seasonality patterns across the 12-year period
SELECT
    Year,
    MONTH(Order_Date)               AS Month_Num,
    MONTHNAME(Order_Date)           AS Month_Name,
    ROUND(SUM(Sales),  2)           AS Monthly_Revenue,
    COUNT(DISTINCT Order_ID)        AS Monthly_Orders
FROM consolidated_sales
GROUP BY Year, Month_Num, Month_Name
ORDER BY Year, Month_Num;


-- B3. Products with a negative profit margin (loss-making SKUs)
SELECT
    Product_Name,
    Category,
    Sub_Category,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(SUM(Profit), 2)           AS Total_Profit,
    ROUND(SUM(Profit) / SUM(Sales) * 100, 2) AS Profit_Margin_Pct,
    COUNT(*)                        AS Times_Sold
FROM consolidated_sales
GROUP BY Product_Name, Category, Sub_Category
HAVING Profit_Margin_Pct < 0
ORDER BY Profit_Margin_Pct ASC;


-- B4. Average order value by shipping method
SELECT
    Ship_Mode,
    COUNT(DISTINCT Order_ID)        AS Total_Orders,
    ROUND(AVG(Sales),  2)           AS Avg_Order_Value,
    ROUND(AVG(Shipping_Cost), 2)    AS Avg_Shipping_Cost,
    ROUND(SUM(Sales),  2)           AS Total_Revenue
FROM consolidated_sales
GROUP BY Ship_Mode
ORDER BY Avg_Order_Value DESC;


-- B5. Order count and revenue by year and region combined
--     Useful for regional growth analysis
SELECT
    Year,
    Region,
    COUNT(DISTINCT Order_ID)        AS Total_Orders,
    ROUND(SUM(Sales),  2)           AS Total_Revenue,
    ROUND(AVG(Sales),  2)           AS Avg_Order_Value
FROM consolidated_sales
GROUP BY Year, Region
ORDER BY Year, Total_Revenue DESC;


-- ============================================================
-- SECTION C — ADVANCED QUERIES
-- ============================================================

-- C1. Year-over-Year revenue growth % using LAG()
--     LAG() looks back at the previous row's value in the ordered set.
SELECT
    Year,
    ROUND(SUM(Sales), 2)            AS Total_Revenue,
    LAG(ROUND(SUM(Sales), 2)) OVER (ORDER BY Year) AS Prev_Year_Revenue,
    ROUND(
        (SUM(Sales) - LAG(SUM(Sales)) OVER (ORDER BY Year))
        / LAG(SUM(Sales)) OVER (ORDER BY Year) * 100,
    2)                              AS YoY_Revenue_Growth_Pct
FROM consolidated_sales
GROUP BY Year
ORDER BY Year;


-- C2. Running (cumulative) total revenue using SUM() OVER()
--     Shows how total revenue has accumulated from 2013 onwards.
SELECT
    Year,
    ROUND(SUM(Sales), 2)             AS Annual_Revenue,
    ROUND(
        SUM(SUM(Sales)) OVER (ORDER BY Year ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW),
    2)                               AS Cumulative_Revenue
FROM consolidated_sales
GROUP BY Year
ORDER BY Year;


-- C3. Rank products by revenue within each category using RANK()
--     RANK() assigns the same rank to ties; numbers skip after a tie.
SELECT
    Category,
    Product_Name,
    ROUND(SUM(Sales), 2)             AS Total_Revenue,
    RANK() OVER (
        PARTITION BY Category
        ORDER BY SUM(Sales) DESC
    )                                AS Revenue_Rank_In_Category
FROM consolidated_sales
GROUP BY Category, Product_Name
ORDER BY Category, Revenue_Rank_In_Category;


-- C4. Best and worst year by total revenue using FIRST_VALUE() and LAST_VALUE()
--     Window frames over all years ordered by revenue ascending then descending.
WITH yearly_revenue AS (
    SELECT
        Year,
        ROUND(SUM(Sales), 2) AS Total_Revenue
    FROM consolidated_sales
    GROUP BY Year
)
SELECT DISTINCT
    FIRST_VALUE(Year)          OVER (ORDER BY Total_Revenue DESC
                                     ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
                               )    AS Best_Year,
    FIRST_VALUE(Total_Revenue) OVER (ORDER BY Total_Revenue DESC
                                     ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
                               )    AS Best_Revenue,
    LAST_VALUE(Year)           OVER (ORDER BY Total_Revenue DESC
                                     ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
                               )    AS Worst_Year,
    LAST_VALUE(Total_Revenue)  OVER (ORDER BY Total_Revenue DESC
                                     ROWS BETWEEN UNBOUNDED PRECEDING AND UNBOUNDED FOLLOWING
                               )    AS Worst_Revenue
FROM yearly_revenue
LIMIT 1;


-- C5. CTE-based year classification: Growth, Decline, or Stable
--     Based on YoY revenue change compared to previous year.
WITH yearly_rev AS (
    SELECT
        Year,
        SUM(Sales) AS Revenue
    FROM consolidated_sales
    GROUP BY Year
),
yoy_change AS (
    SELECT
        Year,
        Revenue,
        LAG(Revenue) OVER (ORDER BY Year)    AS Prev_Revenue,
        Revenue - LAG(Revenue) OVER (ORDER BY Year) AS Revenue_Change,
        ROUND(
            (Revenue - LAG(Revenue) OVER (ORDER BY Year))
            / LAG(Revenue) OVER (ORDER BY Year) * 100,
        2)                                   AS Growth_Pct
    FROM yearly_rev
)
SELECT
    Year,
    ROUND(Revenue,        2)                 AS Total_Revenue,
    ROUND(Revenue_Change, 2)                 AS Revenue_Change_vs_Prior,
    Growth_Pct,
    CASE
        WHEN Growth_Pct IS NULL  THEN 'Baseline'
        WHEN Growth_Pct >= 5     THEN 'Growth'
        WHEN Growth_Pct <= -5    THEN 'Decline'
        ELSE                          'Stable'
    END                                      AS Year_Classification
FROM yoy_change
ORDER BY Year;


-- C6. Cumulative profit margin trend using window functions
--     Shows how blended profit margin has evolved over the 12 years.
WITH yearly_agg AS (
    SELECT
        Year,
        SUM(Sales)  AS Revenue,
        SUM(Profit) AS Profit
    FROM consolidated_sales
    GROUP BY Year
)
SELECT
    Year,
    ROUND(Revenue, 2)                        AS Annual_Revenue,
    ROUND(Profit,  2)                        AS Annual_Profit,
    ROUND(Profit / Revenue * 100, 2)         AS Annual_Margin_Pct,
    ROUND(
        SUM(Profit)  OVER (ORDER BY Year) /
        SUM(Revenue) OVER (ORDER BY Year) * 100,
    2)                                       AS Cumulative_Margin_Pct
FROM yearly_agg
ORDER BY Year;
