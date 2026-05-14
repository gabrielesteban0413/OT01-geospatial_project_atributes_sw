import re
import os
import logging
from pathlib import Path
from urllib.parse import quote_plus

import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from sqlalchemy import create_engine, text
from dotenv import load_dotenv

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)

SCHEMA_NAME = "raw"   


def load_db_config(env_path: Path) -> dict:
    load_dotenv(dotenv_path=env_path)
    config = {
        "host": os.getenv("PGHOST", "localhost"),
        "port": os.getenv("PGPORT", "5432"),
        "database": os.getenv("PGDATABASE", "quality"),
        "user": os.getenv("PGUSER", "postgres"),
        "password": os.getenv("PGPASSWORD"),
    }
    if not config["password"]:
        raise ValueError("PGPASSWORD not found in environment")
    return config


def clean_column_name(col: str) -> str:
    col = str(col).strip().lower()
    col = re.sub(r"[^a-z0-9_]", "_", col)
    col = re.sub(r"_+", "_", col)
    return col


def read_excel_sheet(path: Path, sheet_name: str) -> pd.DataFrame:
    logger.info("Reading sheet '%s' from %s", sheet_name, path)
    df = pd.read_excel(path, sheet_name=sheet_name)
    logger.info("Loaded %d rows, %d columns", len(df), len(df.columns))
    df.columns = [clean_column_name(col) for col in df.columns]
    df = df.dropna(how="all")
    return df


def create_postgres_engine(db_config: dict):
    encoded_password = quote_plus(db_config["password"])
    url = (
        f"postgresql://{db_config['user']}:{encoded_password}"
        f"@{db_config['host']}:{db_config['port']}/{db_config['database']}"
    )
    engine = create_engine(url, client_encoding="utf8")
    with engine.connect() as conn:
        conn.execute(text("SELECT 1"))
    logger.info("Database connection successful")
    return engine


def create_table_from_dataframe(df: pd.DataFrame, table_name: str, engine) -> None:
    logger.info("Creating/replacing table '%s' in schema '%s'", table_name, SCHEMA_NAME)
    df.head(0).to_sql(table_name, engine, if_exists="replace", index=False, schema=SCHEMA_NAME)


def bulk_insert_dataframe(df: pd.DataFrame, table_name: str, engine, chunk_size: int = 5000) -> None:
    qualified_name = f"{SCHEMA_NAME}.{table_name}"
    logger.info("Inserting data into %s", qualified_name)

    psycopg_conn = engine.raw_connection()
    try:
        with psycopg_conn.cursor() as cursor:
            data_tuples = [tuple(row) for row in df.to_numpy()]
            columns = list(df.columns)
            insert_sql = f"INSERT INTO {qualified_name} ({', '.join(columns)}) VALUES %s"
            total = len(data_tuples)
            for i in range(0, total, chunk_size):
                chunk = data_tuples[i:i + chunk_size]
                execute_values(cursor, insert_sql, chunk, page_size=len(chunk))
                psycopg_conn.commit()
                logger.info("Inserted %d / %d rows", min(i + chunk_size, total), total)
    finally:
        psycopg_conn.close()


def main():
    base_dir = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs")
    env_path = base_dir / ".env"
    excel_path = base_dir / "database" / "output" / "03_cardidwg_excel.xlsx"

    db_config = load_db_config(env_path)
    df = read_excel_sheet(excel_path, sheet_name="Archivos_DWG")
    engine = create_postgres_engine(db_config)

    table_name = "cardiseño_origin"
    create_table_from_dataframe(df, table_name, engine)
    bulk_insert_dataframe(df, table_name, engine)

    logger.info("Upload completed successfully")


if __name__ == "__main__":
    main()