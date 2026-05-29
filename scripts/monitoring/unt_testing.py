import os
import sys
import re
import csv
import pickle
from datetime import datetime, timedelta
from typing import Set, Tuple, Optional, List

import pandas as pd
import psycopg2
from psycopg2 import sql
from psycopg2.extensions import connection as PgConnection
from rapidfuzz import process, fuzz
from dotenv import load_dotenv

load_dotenv()

DB_CONFIG = {
    "host": os.getenv("PGHOST", "localhost"),
    "port": int(os.getenv("PGPORT", 5432)),
    "dbname": os.getenv("PGDATABASE", "quality"),
    "user": os.getenv("PGUSER", "postgres"),
    "password": os.getenv("PGPASSWORD", ""),
    "connect_timeout": 10,
}

DB_TABLE = "staging.merge_asphia_portafolio"
DB_ID_COLUMN = "id_servicio"
DB_CLIENT_COLUMN = "cuenta_cliente"

GS_BASE_PATH = os.getenv("GS_BASE_PATH", ".")
SW_FILE_PATH = os.path.join(GS_BASE_PATH, "01-Bulk export of ports.xlsx")

SW_PHYSICAL_STATUS_COL = "Physical Status"
SW_SERVICE_ID_COL = "Id Servicio"
SW_CLIENT_NAME_COL = "Nombre Cliente"
SW_FILTER_VALUE = "In Service"
SW_DESCRIPTION_SHELF_COL = "Description Shelf"

# Añadimos Description Shelf para poder filtrar OT / OI
SW_USECOLS = [
    "Id",
    SW_SERVICE_ID_COL,
    "Cambio Service Manager",
    "Proyecto De Red",
    "Numero De Factibilidad-Cable",
    SW_CLIENT_NAME_COL,
    SW_PHYSICAL_STATUS_COL,
    SW_DESCRIPTION_SHELF_COL,
]

OUTPUT_DIR = GS_BASE_PATH
OUTPUT_FILENAME = "04-SW-OTHERS.csv"
OUTPUT_PATH = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
OUTPUT_SEP = ";"
OUTPUT_ENCODING = "utf-8-sig"

REPORT_COLUMNS = [
    "Id",
    "Id Servicio",
    "Cambio Service Manager",
    "Proyecto De Red",
    "Numero De Factibilidad-Cable",
    "Nombre Cliente",
    "coincidencia_nombre",
]

USE_CACHE = True
CACHE_DIR = os.path.join(OUTPUT_DIR, "cache")
CACHE_FILE = os.path.join(CACHE_DIR, "other_ids.pkl")
CACHE_EXPIRY_HOURS = 24

SIMILARITY_THRESHOLD = 80

def normalize_client_name(name: str) -> str:
    if not isinstance(name, str) or not name.strip():
        return ""
    name = name.upper().strip()
    name = re.sub(r'[^A-Z0-9\s]', ' ', name)
    name = re.sub(r'\s+', ' ', name).strip()
    suffixes = [
        r'\bS\s*A\s*S\b', r'\bS\s*A\b', r'\bLTDA\b', r'\bLTD\b',
        r'\bS\s*L\b', r'\bS\s*A\s*S\s*DE\s*C\s*V\b', r'\bS\s*C\s*A\b', r'\bE\s*U\b',
    ]
    for pattern in suffixes:
        name = re.sub(pattern, '', name).strip()
        name = re.sub(r'\s+', ' ', name).strip()
    return name

def get_db_connection() -> PgConnection:
    if not DB_CONFIG["password"]:
        sys.stderr.write("PGPASSWORD no definida\n")
    return psycopg2.connect(**DB_CONFIG)

def is_cache_valid() -> bool:
    if not os.path.exists(CACHE_FILE):
        return False
    mod_time = datetime.fromtimestamp(os.path.getmtime(CACHE_FILE))
    return datetime.now() - mod_time <= timedelta(hours=CACHE_EXPIRY_HOURS)

def load_cache() -> Optional[Tuple[Set[str], List[str]]]:
    try:
        with open(CACHE_FILE, "rb") as f:
            data = pickle.load(f)
            if isinstance(data, tuple) and len(data) == 2:
                return data
    except Exception:
        pass
    return None

def save_cache(ids_set: Set[str], names_list: List[str]) -> None:
    try:
        with open(CACHE_FILE, "wb") as f:
            pickle.dump((ids_set, names_list), f)
    except Exception:
        pass

def fetch_others_data(conn: PgConnection) -> Tuple[Set[str], List[str]]:
    if "." in DB_TABLE:
        schema, table = DB_TABLE.split(".", 1)
        query = sql.SQL("SELECT DISTINCT {}, {} FROM {}.{}").format(
            sql.Identifier(DB_ID_COLUMN),
            sql.Identifier(DB_CLIENT_COLUMN),
            sql.Identifier(schema),
            sql.Identifier(table),
        )
    else:
        query = sql.SQL("SELECT DISTINCT {}, {} FROM {}").format(
            sql.Identifier(DB_ID_COLUMN),
            sql.Identifier(DB_CLIENT_COLUMN),
            sql.Identifier(DB_TABLE),
        )
    with conn.cursor() as cur:
        cur.execute(query)
        rows = cur.fetchall()
        ids_set = set()
        names_set = set()
        for row in rows:
            if row[0] is not None:
                ids_set.add(str(row[0]))
            if row[1] is not None:
                norm = normalize_client_name(str(row[1]))
                if norm:
                    names_set.add(norm)
    return ids_set, list(names_set)

def read_sw_file(file_path: str) -> pd.DataFrame:
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"No se encuentra el archivo SW: {file_path}")
    if not file_path.endswith((".xls", ".xlsx")):
        raise ValueError("Solo se soportan archivos Excel (.xls/.xlsx)")
    return pd.read_excel(
        file_path,
        usecols=SW_USECOLS,
        dtype=str,
        keep_default_na=False,
    )

def filter_in_service_and_ot(df: pd.DataFrame) -> pd.DataFrame:
    """Filtra 'In Service' y elimina filas con 'OT' u 'OI' en Description Shelf."""
    # Excluir filas que contengan OT u OI como palabra completa
    if SW_DESCRIPTION_SHELF_COL in df.columns:
        mask_ot_oi = df[SW_DESCRIPTION_SHELF_COL].str.contains(r'\b(OT|OI)\b', case=False, na=False)
        df = df[~mask_ot_oi]
    # Filtrar por Physical Status
    mask_in_service = df[SW_PHYSICAL_STATUS_COL] == SW_FILTER_VALUE
    df = df[mask_in_service].drop(columns=[SW_PHYSICAL_STATUS_COL, SW_DESCRIPTION_SHELF_COL], errors='ignore')
    return df

def find_best_match(sw_name_norm: str, candidate_list: List[str], threshold: int) -> Optional[str]:
    if not candidate_list or not sw_name_norm:
        return None
    match = process.extractOne(
        sw_name_norm,
        candidate_list,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=threshold,
    )
    if match:
        return f"Posible coincidencia: ({match[1]:.0f}%) {match[0]} "
    return None

def check_consistency(sw_df: pd.DataFrame, existing_ids: Set[str], existing_names_list: List[str]) -> pd.DataFrame:
    mask_faltante = ~sw_df[SW_SERVICE_ID_COL].isin(existing_ids)
    inconsistencias = sw_df[mask_faltante].copy()

    def determinar_coincidencia(nombre_sw: str) -> str:
        if not isinstance(nombre_sw, str) or not nombre_sw.strip():
            return "Sin coincidencia en others"
        norm_sw = normalize_client_name(nombre_sw)
        if not norm_sw:
            return "Sin coincidencia en others"
        if norm_sw in set(existing_names_list):
            return "OK"
        best = find_best_match(norm_sw, existing_names_list, SIMILARITY_THRESHOLD)
        return best if best else "Sin coincidencia en others"

    inconsistencias["coincidencia_nombre"] = inconsistencias[SW_CLIENT_NAME_COL].apply(determinar_coincidencia)
    for col in REPORT_COLUMNS:
        if col not in inconsistencias.columns:
            inconsistencias[col] = ""
    return inconsistencias[REPORT_COLUMNS]

def _proteger_valor(valor) -> str:
    if not isinstance(valor, str):
        return str(valor)
    if re.fullmatch(r'\d{12,}', valor):
        return f'="{valor}"'
    return valor

def save_report(df: pd.DataFrame, path: str) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    df_protegido = df.map(_proteger_valor)   # Corregido FutureWarning
    df_protegido.to_csv(
        path,
        index=False,
        sep=OUTPUT_SEP,
        encoding=OUTPUT_ENCODING,
        quoting=csv.QUOTE_MINIMAL,
    )

def main():
    print("Iniciando validación SW vs Others...", flush=True)
    os.makedirs(CACHE_DIR, exist_ok=True)

    existing_ids: Set[str] = set()
    existing_names_list: List[str] = []
    conn = None
    try:
        if USE_CACHE and is_cache_valid():
            cached = load_cache()
            if cached is not None:
                existing_ids, existing_names_list = cached
                print("Usando caché de IDs de others.")
            else:
                conn = get_db_connection()
                existing_ids, existing_names_list = fetch_others_data(conn)
                save_cache(existing_ids, existing_names_list)
        else:
            conn = get_db_connection()
            existing_ids, existing_names_list = fetch_others_data(conn)
            if USE_CACHE:
                save_cache(existing_ids, existing_names_list)
    except Exception as e:
        sys.stderr.write(f"Error obteniendo datos de referencia: {e}\n")
        sys.exit(1)
    finally:
        if conn and not conn.closed:
            conn.close()

    print("Leyendo archivo SW...", flush=True)
    try:
        sw_raw = read_sw_file(SW_FILE_PATH)
    except Exception as e:
        sys.stderr.write(f"Error leyendo archivo SW: {e}\n")
        sys.exit(1)

    sw_in_service = filter_in_service_and_ot(sw_raw)

    if sw_in_service.empty:
        print("No hay registros 'In Service' sin OT/OI. Generando reporte vacío.")
        pd.DataFrame(columns=REPORT_COLUMNS).to_csv(
            OUTPUT_PATH, index=False, sep=OUTPUT_SEP, encoding=OUTPUT_ENCODING, quoting=csv.QUOTE_MINIMAL
        )
        return

    print("Analizando coincidencias...", flush=True)
    inconsistencias = check_consistency(sw_in_service, existing_ids, existing_names_list)

    save_report(inconsistencias, OUTPUT_PATH)
    print(f"Reporte guardado en: {OUTPUT_PATH}")
    print(f"Servicios 'In Service' analizados (sin OT/OI): {len(sw_in_service)}")
    print(f"Faltantes en others: {len(inconsistencias)}")
    print("Proceso completado.")

if __name__ == "__main__":
    main()