import os
from pathlib import Path
import re
import unicodedata
import hashlib
import csv
from io import StringIO
from typing import Optional

import psycopg2
from dotenv import load_dotenv

ENV_PATH = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\.env")
load_dotenv(dotenv_path=ENV_PATH, override=True)

DB_CONFIG = {
    'host': os.environ.get('PGHOST', 'localhost'),
    'port': os.environ.get('PGPORT', '5432'),
    'database': os.environ.get('PGDATABASE_CONNECTIVITY', 'connectivity'),
    'user': os.environ.get('PGUSER', 'postgres'),
    'password': os.environ.get('PGPASSWORD', '')
}

ESQUEMA = 'raw'

TABLAS = {
    'asphia':        'r_asphia',
    'baseodf':       'r_baseodf',
    'cardiseño':     'r_cardiseño',
    'cardiseno':     'r_cardiseño',
    'nce':           'r_nce',
    'portafolio':    'r_portafolio',
    'vias':          'r_vias',
    'clientes-corp': 'r_clientes_corp',
    'clientes_corp': 'r_clientes_corp',
}


REGLAS_COLUMNAS = {
    'r_nce': {
        'mapa': {
            'ne':               'equipo central',
            'port_full_name':   'port_ec',
            'port_description': 'etiqueta',
        },
        'prioridad': ['equipo central', 'port_ec'],
        'limpiar':   {'port_ec': re.compile(r'GigabitEthernet', re.IGNORECASE)},
    },
    'r_vias': {
        'mapa': {
            'equipo':              'equipo central',
            'interfaz_puerto_ge':  'port_ec',
        },
        'prioridad': ['equipo central', 'port_ec'],
    },
}


# Hoja especifica por tabla (solo aplica a Excel). Si no esta, se usa la primera hoja.
HOJAS = {
    'r_clientes_corp': 'Clientes priorizados conectivid',
}


# GS_OUTPUT es obligatorio: no hay fallback hardcodeado.
_gs_output = os.environ.get('GS_OUTPUT')
if not _gs_output:
    raise RuntimeError("Falta GS_OUTPUT en el .env")
GS_OUTPUT = Path(_gs_output)

COPY_CHUNK = 50_000


# ---------------------------------------------------------------- utilidades

def sanitizar_columna(nombre: str) -> str:
    if not nombre:
        return 'columna'
    nombre = unicodedata.normalize('NFKD', nombre).encode('ASCII', 'ignore').decode('ASCII')
    nombre = re.sub(r'[^\w]', '_', nombre)
    nombre = re.sub(r'_+', '_', nombre)
    if nombre and nombre[0].isdigit():
        nombre = '_' + nombre
    return nombre.lower().strip('_')


def normalizar_valor(v) -> str:
    """Convierte cualquier valor a texto. Elimina el '.0' de floats enteros."""
    if v is None:
        return ''
    if isinstance(v, float):
        if v.is_integer():
            return str(int(v))
        return repr(v)
    return str(v)


def normalizar_nombres(columnas: list, tabla: str):
    """
    Devuelve (nombres_finales, orden) donde orden[i] es el indice original
    de la i-esima columna final.
    """
    saneadas = [
        sanitizar_columna(str(c)) if c not in (None, '') else f'columna_{i}'
        for i, c in enumerate(columnas)
    ]

    vistos: dict = {}
    resultado: list = []
    for c in saneadas:
        if c in vistos:
            vistos[c] += 1
            c = f"{c}_{vistos[c]}"
        else:
            vistos[c] = 0
        resultado.append(c)

    orden = list(range(len(resultado)))
    reglas = REGLAS_COLUMNAS.get(tabla) or {}

    mapa = reglas.get('mapa') or {}
    if mapa:
        resultado = [mapa.get(c, c) for c in resultado]

    prioridad = reglas.get('prioridad') or []
    for col in reversed(prioridad):
        if col in resultado:
            idx = resultado.index(col)
            resultado.insert(0, resultado.pop(idx))
            orden.insert(0, orden.pop(idx))

    return resultado, orden


def hash_archivo(ruta: Path) -> str:
    h = hashlib.md5()
    with ruta.open('rb') as f:
        for chunk in iter(lambda: f.read(1 << 20), b''):
            h.update(chunk)
    return h.hexdigest()


# ---------------------------------------------------------------- control de cargas

def tabla_hash(conn, tabla: str) -> Optional[str]:
    with conn.cursor() as cur:
        cur.execute(f"CREATE SCHEMA IF NOT EXISTS {ESQUEMA}")
        cur.execute(f"""
            CREATE TABLE IF NOT EXISTS {ESQUEMA}._cargas_hash (
                tabla  TEXT PRIMARY KEY,
                hash   TEXT NOT NULL,
                filas  INTEGER,
                actualizado TIMESTAMP DEFAULT NOW()
            )
        """)
        conn.commit()
        cur.execute(f"SELECT hash FROM {ESQUEMA}._cargas_hash WHERE tabla = %s", (tabla,))
        row = cur.fetchone()
        return row[0] if row else None


def guardar_hash(conn, tabla: str, hash_: str, filas: int) -> None:
    with conn.cursor() as cur:
        cur.execute(f"""
            INSERT INTO {ESQUEMA}._cargas_hash (tabla, hash, filas, actualizado)
            VALUES (%s, %s, %s, NOW())
            ON CONFLICT (tabla) DO UPDATE
                SET hash = EXCLUDED.hash,
                    filas = EXCLUDED.filas,
                    actualizado = NOW()
        """, (tabla, hash_, filas))
        conn.commit()


# ---------------------------------------------------------------- lectura

def iterar_excel(ruta: Path, hoja_nombre: Optional[str] = None):
    from python_calamine import CalamineWorkbook
    wb = CalamineWorkbook.from_path(str(ruta))
    hoja = wb.get_sheet_by_name(hoja_nombre) if hoja_nombre else wb.get_sheet_by_index(0)
    datos = hoja.to_python(skip_empty_area=False)
    if not datos:
        return [], iter(())
    columnas = [c for c in datos[0]]
    return columnas, iter(datos[1:])


def iterar_csv(ruta: Path):
    for enc in ('utf-8-sig', 'utf-8', 'latin-1'):
        try:
            f = ruta.open('r', encoding=enc, newline='')
            sample = f.read(8192)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
            except csv.Error:
                dialect = csv.excel
            reader = csv.reader(f, dialect)
            columnas = next(reader)
            return columnas, reader
        except UnicodeDecodeError:
            f.close()
            continue
    raise ValueError(f"No se pudo leer: {ruta}")


def abrir_archivo(ruta: Path, hoja_nombre: Optional[str] = None):
    ext = ruta.suffix.lower()
    if ext in ('.xlsx', '.xlsm', '.xls'):
        return iterar_excel(ruta, hoja_nombre)
    return iterar_csv(ruta)


# ---------------------------------------------------------------- escritura

def _columnas_actuales(conn, tabla: str) -> Optional[list]:
    with conn.cursor() as cur:
        cur.execute("""
            SELECT column_name
            FROM information_schema.columns
            WHERE table_schema = %s AND table_name = %s
            ORDER BY ordinal_position
        """, (ESQUEMA, tabla))
        rows = cur.fetchall()
        return [r[0] for r in rows] if rows else None


def _crear_tabla(conn, tabla: str, columnas: list) -> None:
    with conn.cursor() as cur:
        cur.execute(f"DROP TABLE IF EXISTS {ESQUEMA}.{tabla} CASCADE")
        cols_def = ', '.join([f'"{c}" TEXT' for c in columnas])
        cur.execute(f'CREATE UNLOGGED TABLE {ESQUEMA}.{tabla} ({cols_def})')
        conn.commit()


def _copiar_stream(conn, tabla: str, filas, ncols: int) -> int:
    sql = f"COPY {ESQUEMA}.{tabla} FROM STDIN (FORMAT csv, DELIMITER '\t', NULL '\\N')"
    total = 0
    buf = StringIO()
    writer = csv.writer(buf, delimiter='\t', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')

    def flush(cur):
        nonlocal total
        if buf.tell() == 0:
            return
        buf.seek(0)
        cur.copy_expert(sql, buf)
        total += buf.getvalue().count('\n')
        buf.seek(0)
        buf.truncate(0)

    with conn.cursor() as cur:
        contador = 0
        for fila in filas:
            valores = ['' if v is None else str(v) for v in fila]
            if len(valores) < ncols:
                valores += [''] * (ncols - len(valores))
            elif len(valores) > ncols:
                valores = valores[:ncols]
            writer.writerow(valores)
            contador += 1
            if contador >= COPY_CHUNK:
                flush(cur)
                contador = 0
        flush(cur)
        conn.commit()
    return total


# ---------------------------------------------------------------- carga

def cargar(conn, ruta: Path, tabla: str):
    if not ruta.exists():
        raise FileNotFoundError(f"No existe: {ruta}")

    h = hash_archivo(ruta)
    if h == tabla_hash(conn, tabla):
        return 'sin_cambios', 0

    hoja_nombre = HOJAS.get(tabla)
    columnas_crudas, filas = abrir_archivo(ruta, hoja_nombre)
    columnas, orden = normalizar_nombres(columnas_crudas, tabla)
    ncols = len(columnas)

    reglas = REGLAS_COLUMNAS.get(tabla) or {}
    limpiar = reglas.get('limpiar') or {}
    limpieza_pos = {columnas.index(c): p for c, p in limpiar.items() if c in columnas}

    def filas_procesadas():
        n = len(orden)
        for fila in filas:
            vals = list(fila)
            if len(vals) < n:
                vals += [None] * (n - len(vals))
            fila_ord = [normalizar_valor(vals[i]) for i in orden]
            for pos, patron in limpieza_pos.items():
                v = fila_ord[pos]
                if v:
                    fila_ord[pos] = patron.sub('', v).strip()
            yield fila_ord

    actuales = _columnas_actuales(conn, tabla)
    if actuales == columnas:
        with conn.cursor() as cur:
            cur.execute(f"TRUNCATE {ESQUEMA}.{tabla}")
            conn.commit()
    else:
        _crear_tabla(conn, tabla, columnas)

    total = _copiar_stream(conn, tabla, filas_procesadas(), ncols)
    guardar_hash(conn, tabla, h, total)
    return 'cargado', total


# ---------------------------------------------------------------- descubrimiento

def descubrir_archivos() -> list:
    if not GS_OUTPUT.exists():
        raise FileNotFoundError(f"GS_OUTPUT no existe: {GS_OUTPUT}")

    encontrados: list = []
    for f in sorted(GS_OUTPUT.iterdir()):
        if not f.is_file():
            continue
        stem = f.stem.lower()
        if not stem.startswith('02-'):
            continue
        nombre = stem[3:]
        if nombre.startswith('raw_'):
            nombre = nombre[4:]
        clave = nombre.replace('_', '-')
        tabla = TABLAS.get(clave) or TABLAS.get(nombre)
        if not tabla:
            tabla = 'r_' + sanitizar_columna(nombre)
        encontrados.append((f, tabla))
    return encontrados


# ---------------------------------------------------------------- main

def main() -> None:
    archivos = descubrir_archivos()
    if not archivos:
        print(f"Sin archivos 02-* en {GS_OUTPUT}")
        return

    conn = psycopg2.connect(**DB_CONFIG)
    try:
        with conn.cursor() as cur:
            cur.execute("SET synchronous_commit = off")
            conn.commit()
    except Exception:
        conn.rollback()

    cargados = sin_cambios = errores = 0
    try:
        for ruta, tabla in archivos:
            try:
                estado, filas = cargar(conn, ruta, tabla)
            except Exception as e:
                conn.rollback()
                errores += 1
                print(f"  [ERROR] {ruta.name}: {e}")
                continue

            if estado == 'sin_cambios':
                sin_cambios += 1
                print(f"  [SKIP]  {ruta.name}")
            else:
                cargados += 1
                print(f"  [OK]    {ruta.name} -> {ESQUEMA}.{tabla} ({filas:,} filas)")
    finally:
        conn.close()

    print()
    print(f"Resultado: {cargados} cargados, {sin_cambios} sin cambios, {errores} errores")


if __name__ == "__main__":
    main()