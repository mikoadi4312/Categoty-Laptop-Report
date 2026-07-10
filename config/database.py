import os
from urllib.parse import quote_plus

import psycopg2
from dotenv import load_dotenv
from sqlalchemy import create_engine


load_dotenv()


def _db_config() -> dict:
    return {
        "host": os.getenv("DB_HOST", "localhost"),
        "port": int(os.getenv("DB_PORT", "5432")),
        "dbname": os.getenv("DB_NAME", "laptop_report"),
        "user": os.getenv("DB_USER", "postgres"),
        "password": os.getenv("DB_PASSWORD", ""),
    }


def get_connection():
    """Return a psycopg2 PostgreSQL connection."""
    return psycopg2.connect(**_db_config())


def get_engine():
    """Return a SQLAlchemy engine for pandas/report queries."""
    cfg = _db_config()
    user = quote_plus(cfg["user"])
    password = quote_plus(cfg["password"])
    host = cfg["host"]
    port = cfg["port"]
    dbname = cfg["dbname"]
    url = f"postgresql+psycopg2://{user}:{password}@{host}:{port}/{dbname}"
    return create_engine(url, future=True)
