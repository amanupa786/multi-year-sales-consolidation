-- ============================================================
-- 01_create_tables.sql
-- Multi-Year Sales Consolidation (2013–2024)
-- ============================================================
-- PURPOSE : Create the MySQL database and consolidated_sales
--           table with appropriate data types and indexes.
-- RUN     : mysql -u root -p < sql/01_create_tables.sql
-- ============================================================

-- Create database if it doesn't already exist
CREATE DATABASE IF NOT EXISTS sales_analytics
    CHARACTER SET utf8mb4
    COLLATE utf8mb4_unicode_ci;

USE sales_analytics;

-- Drop and recreate table for clean setup
DROP TABLE IF EXISTS consolidated_sales;

CREATE TABLE consolidated_sales (
    id                INT           AUTO_INCREMENT PRIMARY KEY COMMENT 'Surrogate key',
    Order_ID          VARCHAR(30)   NOT NULL                   COMMENT 'Unique order identifier',
    Order_Date        DATE                                     COMMENT 'Date the order was placed',
    Customer_Name     VARCHAR(100)                             COMMENT 'Full name of the customer',
    Customer_Segment  VARCHAR(50)                              COMMENT 'Consumer | Corporate | Home Office',
    Region            VARCHAR(50)                              COMMENT 'East | West | Central | South',
    Country           VARCHAR(50)                              COMMENT 'Country of the order',
    Category          VARCHAR(50)                              COMMENT 'Technology | Furniture | Office Supplies',
    Sub_Category      VARCHAR(50)                              COMMENT 'Product sub-category',
    Product_Name      VARCHAR(150)                             COMMENT 'Name of the product sold',
    Quantity          INT                                      COMMENT 'Number of units ordered',
    Unit_Price        DECIMAL(10,2)                            COMMENT 'Price per unit before discount',
    Discount          DECIMAL(5,2)  DEFAULT 0.00               COMMENT 'Discount rate applied (0–1)',
    Sales             DECIMAL(12,2)                            COMMENT 'Net revenue after discount',
    Profit            DECIMAL(12,2)                            COMMENT 'Net profit (can be negative)',
    Shipping_Cost     DECIMAL(10,2)                            COMMENT 'Cost to ship the order',
    Ship_Mode         VARCHAR(50)                              COMMENT 'Standard Class | Second Class | First Class | Same Day',
    Year              INT                                      COMMENT 'Fiscal year of the order',
    Profit_Margin_Pct DECIMAL(8,2)                             COMMENT '(Profit / Sales) * 100',
    Revenue_Band      VARCHAR(10)                              COMMENT 'Low | Medium | High based on Sales',

    -- Indexes for common query patterns
    INDEX idx_year        (Year),
    INDEX idx_category    (Category),
    INDEX idx_region      (Region),
    INDEX idx_segment     (Customer_Segment),
    INDEX idx_order_date  (Order_Date),
    INDEX idx_product     (Product_Name)
) ENGINE=InnoDB
  DEFAULT CHARSET=utf8mb4
  COMMENT='Consolidated 12-year sales data (2013–2024)';

-- Verify table was created
SHOW COLUMNS FROM consolidated_sales;
