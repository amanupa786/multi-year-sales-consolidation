# ============================================================
# config.py — MySQL Connection Settings
# ============================================================
# Edit ONLY this file to change database credentials.
# Never hardcode credentials inside pipeline.py or other scripts.

DB_CONFIG = {
    "host":     "localhost",
    "port":     3306,
    "user":     "root",           # Change to your MySQL username
    "password": "your_password",  # Change to your MySQL password
    "database": "sales_analytics"
}

# Output paths
RAW_DATA_PATH       = "data/raw"
PROCESSED_DATA_PATH = "data/processed"
OUTPUTS_PATH        = "outputs"
CHARTS_PATH         = "outputs/charts"
SQL_PATH            = "sql"
