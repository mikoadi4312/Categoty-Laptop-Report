import argparse
import re
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from psycopg2.extras import execute_values


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config.database import get_connection  # noqa: E402


REQUIRED_COLUMNS = [
    "ID Store",
    "Store Name",
    "Main Category",
    "Sub Category",
    "ID Model",
    "Product code",
    "Product name",
    "Inventory Status",
    "Quantity_1",
    "QUANTITYEX",
]

NEW_STATUSES = {"1-New", "5-Error (New)"}
DEMO_STATUSES = {"3-Show", "7-Show (Sample)"}

BRAND_CANONICAL = {
    "LENOVO": "Lenovo",
    "ASUS": "Asus",
    "ACER": "Acer",
    "HP": "HP",
    "DELL": "Dell",
    "MSI": "MSI",
    "APPLE": "Apple",
    "AXIOO": "Axioo",
    "ADVAN": "Advan",
}


def normalize_brand(value) -> str:
    text = " ".join(str(value or "").strip().split())
    if not text or text.lower() == "nan":
        return "UNKNOWN"
    upper = text.upper()
    return BRAND_CANONICAL.get(upper, text.title())


def strip_numeric_prefix(category: str) -> str:
    text = " ".join(str(category or "").strip().split())
    parts = text.split(" - ")
    if parts and parts[0].isdigit() and len(parts) > 1:
        return " - ".join(parts[1:])
    return text


def extract_brand(product_name) -> str:
    name = " ".join(str(product_name or "").strip().split())
    if not name or name.lower() == "nan":
        return "UNKNOWN"
    if name.lower().startswith("macbook"):
        return "Apple"

    match = re.match(r"^Laptop\s+([^\s/]+)", name, flags=re.IGNORECASE)
    if not match:
        return "UNKNOWN"
    return normalize_brand(match.group(1).strip(".,;:()[]{}"))


def stock_date_from_filename(filename: str, fallback: date) -> date:
    candidates = []
    for value in re.findall(r"20\d{6}", filename):
        try:
            candidates.append(datetime.strptime(value, "%Y%m%d").date())
        except ValueError:
            continue
    return max(candidates) if candidates else fallback


def read_input_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm", ".xls"}:
        return pd.read_excel(path)
    if suffix == ".csv":
        return pd.read_csv(path)
    raise ValueError(f"Unsupported file type: {path.suffix}. Use .xlsx, .xls, .xlsm, or .csv")


def validate_columns(df: pd.DataFrame) -> None:
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")


def prepare_stock(file_path: Path, stock_date) -> pd.DataFrame:
    df = read_input_file(file_path)
    validate_columns(df)

    string_cols = df.select_dtypes(include=["object"]).columns
    for col in string_cols:
        df[col] = df[col].astype(str).str.strip()

    df["main_category_clean"] = df["Main Category"].apply(strip_numeric_prefix)
    df["sub_category_clean"] = df["Sub Category"].apply(strip_numeric_prefix)

    laptop_mask = (
        df["main_category_clean"].astype(str).str.contains("Laptop", case=False, na=False)
        & df["sub_category_clean"].astype(str).str.contains("Laptop", case=False, na=False)
    )
    df = df[laptop_mask].copy()

    df["id_store"] = pd.to_numeric(df["ID Store"], errors="coerce")
    df["store_name"] = df["Store Name"].replace({"": "UNKNOWN STORE"}).fillna("UNKNOWN STORE")
    df["quantity"] = pd.to_numeric(df["Quantity_1"], errors="coerce").fillna(0).round().astype(int)
    df["brand"] = df["Product name"].apply(extract_brand)
    df["stock_date"] = stock_date
    df = df.dropna(subset=["id_store"])
    df["id_store"] = df["id_store"].astype(int)

    df["new_stock"] = df["quantity"].where(df["Inventory Status"].isin(NEW_STATUSES), 0).astype(int)
    df["demo_units"] = df["quantity"].where(df["Inventory Status"].isin(DEMO_STATUSES), 0).astype(int)

    grouped = (
        df.groupby(["stock_date", "id_store", "brand"], as_index=False)
        .agg(
            store_name=("store_name", "last"),
            new_stock=("new_stock", "sum"),
            demo_units=("demo_units", "sum"),
        )
    )
    grouped["stock_volume"] = grouped["new_stock"] + grouped["demo_units"]
    return grouped


def upsert_stores(cur, grouped: pd.DataFrame) -> dict[int, int]:
    stores = [
        (int(id_store), str(store_name))
        for id_store, store_name in grouped[["id_store", "store_name"]]
        .drop_duplicates(subset=["id_store"], keep="last")
        .itertuples(index=False, name=None)
    ]
    if not stores:
        return {}
    rows = execute_values(
        cur,
        """
        INSERT INTO dim_store (id_store, store_name)
        VALUES %s
        ON CONFLICT (id_store) DO UPDATE
        SET store_name = EXCLUDED.store_name
        RETURNING id_store, id
        """,
        stores,
        page_size=len(stores),
        fetch=True,
    )
    return {int(id_store): int(database_id) for id_store, database_id in rows}


def upsert_brands(cur, grouped: pd.DataFrame) -> dict[str, int]:
    brands = [(str(brand),) for brand in grouped["brand"].drop_duplicates().tolist()]
    if not brands:
        return {}
    rows = execute_values(
        cur,
        """
        INSERT INTO dim_brand (brand)
        VALUES %s
        ON CONFLICT (brand) DO UPDATE
        SET brand = EXCLUDED.brand
        RETURNING brand, id
        """,
        brands,
        page_size=len(brands),
        fetch=True,
    )
    return {str(brand): int(database_id) for brand, database_id in rows}


def upsert_fact_stock(
    cur,
    grouped: pd.DataFrame,
    store_ids: dict[int, int],
    brand_ids: dict[str, int],
    uploaded_by: str,
) -> tuple[int, int]:
    values = [
        (
            row.stock_date,
            store_ids[int(row.id_store)],
            brand_ids[row.brand],
            int(row.new_stock),
            int(row.demo_units),
            int(row.stock_volume),
            uploaded_by,
        )
        for row in grouped.itertuples(index=False)
    ]
    if not values:
        return 0, 0
    results = execute_values(
        cur,
        """
        INSERT INTO fact_stock (
            stock_date, store_id, brand_id, new_stock,
            demo_units, stock_volume, uploaded_by
        )
        VALUES %s
        ON CONFLICT (stock_date, store_id, brand_id) DO UPDATE
        SET new_stock = EXCLUDED.new_stock,
            demo_units = EXCLUDED.demo_units,
            stock_volume = EXCLUDED.stock_volume,
            uploaded_at = NOW(),
            uploaded_by = EXCLUDED.uploaded_by
        RETURNING (xmax = 0) AS inserted
        """,
        values,
        page_size=len(values),
        fetch=True,
    )
    inserted = sum(bool(row[0]) for row in results)
    return inserted, len(results) - inserted


def write_upload_log(
    conn,
    upload_type: str,
    file_name: str,
    upload_date,
    uploaded_by: str,
    rows_inserted: int,
    rows_updated: int,
    status: str,
    error_message: str | None = None,
) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO upload_log (
                upload_type, file_name, upload_date, uploaded_by,
                rows_inserted, rows_updated, status, error_message
            )
            VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """,
            (
                upload_type,
                file_name,
                upload_date,
                uploaded_by,
                rows_inserted,
                rows_updated,
                status,
                error_message,
            ),
        )
    conn.commit()


def upload_stock(file_path: Path, uploaded_by: str, stock_date) -> tuple[int, int]:
    grouped = prepare_stock(file_path, stock_date)
    conn = get_connection()
    inserted = 0
    updated = 0

    try:
        with conn.cursor() as cur:
            store_ids = upsert_stores(cur, grouped)
            brand_ids = upsert_brands(cur, grouped)
            inserted, updated = upsert_fact_stock(cur, grouped, store_ids, brand_ids, uploaded_by)

        write_upload_log(
            conn,
            "stock",
            file_path.name,
            stock_date,
            uploaded_by,
            inserted,
            updated,
            "success",
        )
        return inserted, updated
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def parse_date(value: str):
    return datetime.strptime(value, "%Y-%m-%d").date()


def main() -> None:
    parser = argparse.ArgumentParser(description="Upload LAPTOP stock snapshot into PostgreSQL.")
    parser.add_argument("--file", required=True, help="Path to GeneralInventory Excel/CSV file")
    parser.add_argument("--uploaded_by", required=True, help="Uploader name")
    parser.add_argument("--date", required=True, help="Stock snapshot date, YYYY-MM-DD")
    args = parser.parse_args()

    file_path = Path(args.file).expanduser().resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    stock_date = parse_date(args.date)

    try:
        inserted, updated = upload_stock(file_path, args.uploaded_by, stock_date)
        print(f"Stock upload completed. Inserted: {inserted}, Updated: {updated}")
    except Exception as exc:
        conn = get_connection()
        try:
            write_upload_log(
                conn,
                "stock",
                file_path.name,
                stock_date,
                args.uploaded_by,
                0,
                0,
                "failed",
                str(exc),
            )
        finally:
            conn.close()
        raise


if __name__ == "__main__":
    main()
