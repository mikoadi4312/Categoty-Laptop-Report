import argparse
import sys
from pathlib import Path

import pandas as pd


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
    df["day_qty"] = pd.to_numeric(df["Quantity_1"], errors="coerce").fillna(0).round().astype(int)
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
        df.groupby(["sale_date", "id_store", "store_name", "brand"], as_index=False)
        .agg(day_qty=("day_qty", "sum"), day_revenue=("day_revenue", "sum"))
    )
    grouped["day_revenue"] = grouped["day_revenue"].round(2)
    return grouped


def upsert_store(cur, id_store: int, store_name: str) -> int:
    cur.execute(
        """
        INSERT INTO dim_store (id_store, store_name)
        VALUES (%s, %s)
        ON CONFLICT (id_store) DO UPDATE
        SET store_name = EXCLUDED.store_name
        RETURNING id
        """,
        (id_store, store_name),
    )
    return cur.fetchone()[0]


def upsert_brand(cur, brand: str) -> int:
    cur.execute(
        """
        INSERT INTO dim_brand (brand)
        VALUES (%s)
        ON CONFLICT (brand) DO UPDATE
        SET brand = EXCLUDED.brand
        RETURNING id
        """,
        (brand,),
    )
    return cur.fetchone()[0]


def upsert_fact_sales(cur, row, uploaded_by: str) -> bool:
    cur.execute(
        """
        INSERT INTO fact_sales (
            sale_date, store_id, brand_id, day_qty, day_revenue, uploaded_by
        )
        VALUES (%s, %s, %s, %s, %s, %s)
        ON CONFLICT (sale_date, store_id, brand_id) DO UPDATE
        SET day_qty = EXCLUDED.day_qty,
            day_revenue = EXCLUDED.day_revenue,
            uploaded_at = NOW(),
            uploaded_by = EXCLUDED.uploaded_by
        RETURNING (xmax = 0) AS inserted
        """,
        (
            row["sale_date"],
            row["store_id"],
            row["brand_id"],
            int(row["day_qty"]),
            float(row["day_revenue"]),
            uploaded_by,
        ),
    )
    return bool(cur.fetchone()[0])


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
            for _, row in grouped.iterrows():
                store_id = upsert_store(cur, int(row["id_store"]), row["store_name"])
                brand_id = upsert_brand(cur, row["brand"])
                fact_row = row.to_dict()
                fact_row["store_id"] = store_id
                fact_row["brand_id"] = brand_id

                if upsert_fact_sales(cur, fact_row, uploaded_by):
                    inserted += 1
                else:
                    updated += 1

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
