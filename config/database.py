import os
from pathlib import Path
from urllib.parse import quote_plus

import psycopg2
from dotenv import load_dotenv
from sqlalchemy import create_engine


load_dotenv()


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCHEMA_PATH = PROJECT_ROOT / "models" / "schema.sql"
_schema_initialized = False


def _secret(name: str, default: str | None = None) -> str | None:
    value = os.getenv(name)
    if value:
        return value

    try:
        import streamlit as st

        if name in st.secrets:
            return str(st.secrets[name])
        if "database" in st.secrets and name in st.secrets["database"]:
            return str(st.secrets["database"][name])
    except Exception:
        pass

    return default


def _database_url() -> str | None:
    return _secret("DATABASE_URL") or _secret("DB_URL")


def _sqlalchemy_url(url: str) -> str:
    if url.startswith("postgres://"):
        return "postgresql+psycopg2://" + url.removeprefix("postgres://")
    if url.startswith("postgresql://"):
        return "postgresql+psycopg2://" + url.removeprefix("postgresql://")
    return url


def _db_config() -> dict:
    return {
        "host": _secret("DB_HOST", "localhost"),
        "port": int(_secret("DB_PORT", "5432")),
        "dbname": _secret("DB_NAME", "laptop_report"),
        "user": _secret("DB_USER", "postgres"),
        "password": _secret("DB_PASSWORD", ""),
    }


def get_connection():
    """Return a psycopg2 PostgreSQL connection."""
    url = _database_url()
    if url:
        return psycopg2.connect(url)
    return psycopg2.connect(**_db_config())


def get_engine():
    """Return a SQLAlchemy engine for pandas/report queries."""
    url = _database_url()
    if url:
        return create_engine(_sqlalchemy_url(url), future=True)

    cfg = _db_config()
    user = quote_plus(cfg["user"])
    password = quote_plus(cfg["password"])
    host = cfg["host"]
    port = cfg["port"]
    dbname = cfg["dbname"]
    url = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{dbname}"
    return create_engine(url, future=True)


def ensure_schema() -> None:
    global _schema_initialized
    if _schema_initialized:
        return

    schema_sql = SCHEMA_PATH.read_text(encoding="utf-8")
    conn = get_connection()
    try:
        with conn.cursor() as cur:
            cur.execute(schema_sql)
        conn.commit()
        _schema_initialized = True
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()
