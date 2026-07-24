# LAPTOP Daily Reporting System

PostgreSQL-based ETL and Excel reporting for the LAPTOP category.

The system reads two daily files:

- `REVENUEERA` sales file
- `GeneralInventory` stock snapshot file

It generates an Excel report with:

- `LAPTOP REPORT`: brand-level category breakdown
- `SHOP LEVEL`: store-level performance with brands as dynamic column groups

## Project Structure

```text
laptop_report_system/
├── config/
│   └── database.py
├── models/
│   └── schema.sql
├── etl/
│   ├── upload_sales.py
│   └── upload_stock.py
├── reports/
│   └── generate_laptop_report.py
├── .env.example
├── requirements.txt
└── README.md
```

## 1. PostgreSQL Setup

Create the database:

```bash
createdb laptop_report
```

Or from `psql`:

```sql
CREATE DATABASE laptop_report;
```

Install the schema:

```bash
psql -d laptop_report -f models/schema.sql
```

## 2. Python Setup

Create and activate a virtual environment:

```bash
python -m venv .venv
.venv\Scripts\activate
```

Install dependencies:

```bash
pip install -r requirements.txt
```

## 3. Environment File

Copy `.env.example` to `.env`:

```bash
copy .env.example .env
```

Edit `.env` with your PostgreSQL credentials:

```text
DB_HOST=localhost
DB_PORT=5432
DB_NAME=laptop_report
DB_USER=postgres
DB_PASSWORD=your_password_here
```

## 4. Daily Workflow

Run commands from the `laptop_report_system` directory.

## Streamlit Cloud Deployment

The app can run on Streamlit Community Cloud, but the PostgreSQL database must be hosted online. A database running on your laptop with `DB_HOST=localhost` cannot be reached by Streamlit Cloud.

Set the database credentials in Streamlit Cloud:

1. Open the deployed app settings.
2. Go to `Secrets`.
3. Add either one connection string:

```toml
DATABASE_URL = "postgresql://USER:PASSWORD@HOST:5432/DB_NAME"
```

Or add separate values:

```toml
DB_HOST = "your-cloud-db-host"
DB_PORT = "5432"
DB_NAME = "laptop_report"
DB_USER = "your-db-user"
DB_PASSWORD = "your-db-password"
```

After changing secrets, reboot the Streamlit app from the app menu.

## Railway Deployment

For faster uploads, deploy the Streamlit service and PostgreSQL database in the same Railway project and region.

1. Create a Railway project from this GitHub repository.
2. Add a PostgreSQL service to the project.
3. Add `DATABASE_URL=${{Postgres.DATABASE_URL}}` to the Streamlit service variables.
4. Deploy the Streamlit service. The included `Procfile` starts the app using Railway's `PORT` value.

The application creates the required tables and indexes automatically when it connects to an empty database.

### Web Upload

Start the web app:

```bash
streamlit run web_app.py
```

Open the local URL shown by Streamlit, usually:

```text
http://localhost:8501
```

Use the web tabs:

- `Sales`: upload `REVENUEERA` and click `Upload Sales`.
- `Stock`: upload `GeneralInventory`, choose the stock date, and click `Upload Stock`.
- `Report`: choose the report date, generate the Excel file, and download it.

### Command Line

Upload sales:

```bash
python etl/upload_sales.py --file REVENUEERA_YYYYMMDD.xlsx --uploaded_by "YourName"
```

Upload stock:

```bash
python etl/upload_stock.py --file GeneralInventory_YYYYMMDD.xlsx --uploaded_by "YourName" --date YYYY-MM-DD
```

Generate report with separate sales and stock dates:

```bash
python reports/generate_laptop_report.py --date YYYY-MM-DD --stock-date YYYY-MM-DD --output LAPTOP_REPORT_YYYYMMDD.xlsx
```

Re-uploading the same file is safe. Sales and stock tables use upserts on date, store, and brand, so rows are updated instead of duplicated.

## Input Rules

### Sales File

Required columns:

```text
DATE, IDStore, Store Name, MainCategory, SubCategory, Product code,
Product name, Brand name, InventoryStatus, TypeofOutput, Quantity_1, REVENUE
```

Processing:

- Parses `DATE` as `DD/MM/YYYY`.
- Strips whitespace from string columns.
- Keeps rows where `MainCategory` contains `Laptop`, case-insensitive.
- Aggregates by `sale_date`, `IDStore`, `Store Name`, and `Brand name`.
- Sales quantity counts transaction rows after the LAPTOP filter, matching the operational manual report.
- `REVENUE` is summed as IDR numeric value.

### Stock File

Required columns:

```text
ID Store, Store Name, Main Category, Sub Category,
ID Model, Product code, Product name, Inventory Status, Quantity_1, QUANTITYEX
```

Processing:

- Strips numeric category prefixes, for example `1584 - INDO - Laptop` becomes `INDO - Laptop`.
- Keeps rows where both cleaned `Main Category` and cleaned `Sub Category` contain `Laptop`, case-insensitive.
- Extracts brand from `Product name` pattern `Laptop [Brand] [model]`.
- Maps product names starting with `Macbook` to `Apple`.
- Uses `UNKNOWN` when the product name does not identify a brand; `COMPANYBRANDNAME=ERABLUE` is not treated as a laptop brand.
- Detects the stock date from an `YYYYMMDD` value in the filename and allows correction in the web form.
- `1-New` is counted as `new_stock`.
- `5-Error (New)` is counted separately as `error_new_units`.
- `3-Show` and `7-Show (Sample)` are counted as `demo_units`.
- `stock_volume = new_stock + error_new_units + demo_units`.

## Report Logic

`LAPTOP REPORT`:

- Day sales are taken from `fact_sales.sale_date = target date`.
- MTD sales are month-to-date through the target date.
- Stock is taken from the exact stock snapshot date selected in the web form or `--stock-date` argument.
- Stock Day is calculated as `stock_volume / (mtd_qty / target_date.day)`.
- Same-period quantity, same-period revenue, and growth rate are currently set to `0`.
- Brands are sorted by MTD revenue descending.
- Grand total is shown at the top and bottom.

`SHOP LEVEL`:

- Store rows are sorted by total MTD revenue descending.
- Brand columns are generated dynamically from uploaded sales data.
- A total revenue group is added at the far right.
- Grand total is shown at the bottom.
