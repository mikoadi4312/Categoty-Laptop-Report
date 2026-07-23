import argparse
import sys
from pathlib import Path

import pandas as pd
from psycopg2.extras import execute_values


PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from config.database import get_connection  # noqa: E402


REQUIRED_COLUMNS = [
    "DATE",
    "IDStore",
    "Store Name",
    "MainCategory",
    "SubCategory",
    "Product code",
    "Product name",
    "Brand name",
    "InventoryStatus",
    "TypeofOutput",
    "Quantity_1",
    "REVENUE",
]

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


def prepare_sales(file_path: Path) -> pd.DataFrame:
    df = read_input_file(file_path)
    validate_columns(df)

    string_cols = df.select_dtypes(include=["object"]).columns
    for col in string_cols:
        df[col] = df[col].astype(str).str.strip()

    df["sale_date"] = pd.to_datetime(df["DATE"], format="%d/%m/%Y", errors="coerce").dt.date
    df["id_store"] = pd.to_numeric(df["IDStore"], errors="coerce")
    df["day_revenue"] = pd.to_numeric(df["REVENUE"], errors="coerce").fillna(0).round(2)
    df["store_name"] = df["Store Name"].replace({"": "UNKNOWN STORE"}).fillna("UNKNOWN STORE")
    df["brand"] = df["Brand name"].apply(normalize_brand)

    laptop_mask = (
        df["MainCategory"].astype(str).str.contains("Laptop", case=False, na=False)
        & df["SubCategory"].astype(str).str.contains("Laptop", case=False, na=False)
    )
    df = df[laptop_mask].copy()
    df = df.dropna(subset=["sale_date", "id_store"])
    df["id_store"] = df["id_store"].astype(int)

    grouped = (
        df.groupby(["sale_date", "id_store", "brand"], as_index=False)
        .agg(
            store_name=("store_name", "last"),
            day_qty=("Quantity_1", "size"),
            day_revenue=("day_revenue", "sum"),
        )
    )
    grouped["day_revenue"] = grouped["day_revenue"].round(2)
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


def upsert_fact_sales(
    cur,
    grouped: pd.DataFrame,
    store_ids: dict[int, int],
    brand_ids: dict[str, int],
    uploaded_by: str,
) -> tuple[int, int]:
    values = [
        (
            row.sale_date,
            store_ids[int(row.id_store)],
            brand_ids[row.brand],
            int(row.day_qty),
            float(row.day_revenue),
            uploaded_by,
        )
        for row in grouped.itertuples(index=False)
    ]
    if not values:
        return 0, 0
    results = execute_values(
        cur,
        """
        INSERT INTO fact_sales (
            sale_date, store_id, brand_id, day_qty, day_revenue, uploaded_by
        )
        VALUES %s
        ON CONFLICT (sale_date, store_id, brand_id) DO UPDATE
        SET day_qty = EXCLUDED.day_qty,
            day_revenue = EXCLUDED.day_revenue,
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


def upload_sales(file_path: Path, uploaded_by: str) -> tuple[int, int]:
    grouped = prepare_sales(file_path)
    conn = get_connection()
    inserted = 0
    updated = 0

    try:
        with conn.cursor() as cur:
            store_ids = upsert_stores(cur, grouped)
            brand_ids = upsert_brands(cur, grouped)
            inserted, updated = upsert_fact_sales(cur, grouped, store_ids, brand_ids, uploaded_by)

        upload_date = grouped["sale_date"].max() if not grouped.empty else None
        write_upload_log(
            conn,
            "sales",
            file_path.name,
            upload_date,
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


def main() -> None:
    parser = argparse.ArgumentParser(description="Upload LAPTOP daily sales into PostgreSQL.")
    parser.add_argument("--file", required=True, help="Path to REVENUEERA Excel/CSV file")
    parser.add_argument("--uploaded_by", required=True, help="Uploader name")
    args = parser.parse_args()

    file_path = Path(args.file).expanduser().resolve()
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    try:
        inserted, updated = upload_sales(file_path, args.uploaded_by)
        print(f"Sales upload completed. Inserted: {inserted}, Updated: {updated}")
    except Exception as exc:
        conn = get_connection()
        try:
            write_upload_log(
                conn,
                "sales",
                file_path.name,
                None,
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
