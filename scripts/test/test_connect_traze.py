#!/usr/bin/env python3
"""
Trazador de rutas de fibra para archivos GIGANTES (billones de líneas)
sin orden garantizado.

Formato de cada línea (separado por '|' por defecto):
 0  id_cable
 1  codigo
 2  nombre
 3  num_hilo
 4  id_hilo_destino
 5  cliente
 6  estado          (LIBRE / FUNCIONANDO / ...)
 7  tipo             (Anillo / Bus)
 8  red              (METRO)
 9  uso              (Libre / Ocupado)
10  acceso
11  rango
12  nodo_a_id
13  nodo_a_nombre
14  nodo_b_id
15  nodo_b_nombre

USO:
  # Modo automático (sin argumentos): construye índice si no existe, luego traza ruta y guarda en out
  python test_connect_traze.py

  # Forzar reconstrucción del índice
  python test_connect_traze.py --build

  # Trazar ruta específica y guardar en archivo
  python test_connect_traze.py --trace --origen "BASTIDOR B08" --destino "OB" --output "salida.txt"
"""

import sqlite3
import argparse
import sys
import time
import os
from collections import deque

# ---------------------------------------------------------------
# CONFIGURACIÓN POR DEFECTO (usa tus rutas)
# ---------------------------------------------------------------
DEFAULT_ARCHIVO_ENTRADA = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_find.txt"
DEFAULT_ARCHIVO_SALIDA  = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_out.txt"
DEFAULT_DB              = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\indice.db"
DEFAULT_SEP             = "|"
DEFAULT_ENCODING        = "utf-8"
DEFAULT_ORIGEN          = "BASTIDOR"
DEFAULT_DESTINO         = "OB"

# ---------------------------------------------------------------
# 1. CONSTRUCCIÓN DEL ÍNDICE (streaming, con mensajes de diagnóstico)
# ---------------------------------------------------------------
def build_index(archivo, db_path, sep=DEFAULT_SEP, encoding=DEFAULT_ENCODING, batch_size=200_000):
    print(f"[DIAG] Iniciando construcción del índice...", file=sys.stderr)
    print(f"[DIAG] Archivo origen: {archivo}", file=sys.stderr)
    
    if not os.path.exists(archivo):
        print(f"❌ ERROR: El archivo '{archivo}' no existe.", file=sys.stderr)
        sys.exit(1)
    
    tamaño = os.path.getsize(archivo)
    print(f"[DIAG] Tamaño del archivo: {tamaño:,} bytes ({tamaño/1024/1024:.2f} MB)", file=sys.stderr)
    
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode = OFF")
    conn.execute("PRAGMA synchronous = OFF")
    conn.execute("PRAGMA temp_store = MEMORY")
    conn.execute("PRAGMA cache_size = -300000")

    conn.execute("DROP TABLE IF EXISTS hilos")
    conn.execute("""
        CREATE TABLE hilos (
            id_cable      TEXT,
            num_hilo      INTEGER,
            id_hilo_dest  TEXT,
            cliente       TEXT,
            estado        TEXT,
            tipo          TEXT,
            uso           TEXT,
            nodo_a_id     TEXT,
            nodo_a_nombre TEXT,
            nodo_b_id     TEXT,
            nodo_b_nombre TEXT
        )
    """)

    insert_sql = """
        INSERT INTO hilos
        (id_cable, num_hilo, id_hilo_dest, cliente, estado, tipo, uso,
         nodo_a_id, nodo_a_nombre, nodo_b_id, nodo_b_nombre)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
    """

    t0 = time.time()
    buffer = []
    total_validos = 0
    total_lineas = 0
    lineas_corruptas = 0

    print(f"📂 Leyendo archivo: {archivo}", file=sys.stderr)
    with open(archivo, "r", encoding=encoding, errors="ignore") as f:
        for linea in f:
            total_lineas += 1
            linea = linea.rstrip("\n")
            if not linea:
                continue
            p = linea.split(sep)
            if len(p) < 16:
                lineas_corruptas += 1
                continue

            try:
                num_hilo = int(p[3])
            except ValueError:
                num_hilo = None

            buffer.append((
                p[0], num_hilo, p[4], p[5], p[6], p[7], p[9],
                p[12], p[13], p[14], p[15]
            ))
            total_validos += 1

            if len(buffer) >= batch_size:
                conn.executemany(insert_sql, buffer)
                conn.commit()
                buffer.clear()
                print(f"  {total_validos:,} líneas válidas procesadas... ({time.time()-t0:.1f}s)",
                      file=sys.stderr)

    if buffer:
        conn.executemany(insert_sql, buffer)
        conn.commit()

    print(f"[DIAG] Total líneas leídas: {total_lineas:,}", file=sys.stderr)
    print(f"[DIAG] Líneas corruptas (ignoradas): {lineas_corruptas:,}", file=sys.stderr)
    print(f"[DIAG] Líneas válidas insertadas: {total_validos:,}", file=sys.stderr)

    if total_validos == 0:
        print("⚠️ ADVERTENCIA: No se insertó ninguna línea válida. Verifica el separador y el formato.", file=sys.stderr)
        conn.close()
        return False

    print("🔧 Creando índices...", file=sys.stderr)
    conn.execute("CREATE INDEX ix_nodo_a ON hilos(nodo_a_id)")
    conn.execute("CREATE INDEX ix_nodo_b ON hilos(nodo_b_id)")
    conn.execute("CREATE INDEX ix_cable  ON hilos(id_cable)")
    conn.execute("CREATE INDEX ix_estado ON hilos(estado)")
    conn.commit()
    conn.close()

    print(f"✅ Índice construido: {total_validos:,} filas en {time.time()-t0:.1f}s -> {db_path}",
          file=sys.stderr)
    return True

# ---------------------------------------------------------------
# 2. BFS SOBRE EL ÍNDICE (con mensajes de diagnóstico)
# ---------------------------------------------------------------
def find_path(db_path, origen_texto, destino_texto):
    print(f"[DIAG] Buscando ruta desde '{origen_texto}' hasta '{destino_texto}'", file=sys.stderr)
    
    if not os.path.exists(db_path):
        print(f"❌ ERROR: El índice '{db_path}' no existe. Ejecuta primero --build.", file=sys.stderr)
        return None

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()

    cur.execute("SELECT COUNT(*) FROM hilos")
    count = cur.fetchone()[0]
    print(f"[DIAG] El índice contiene {count:,} registros.", file=sys.stderr)
    if count == 0:
        print("⚠️ ADVERTENCIA: El índice está vacío. Reconstruye con --build.", file=sys.stderr)
        conn.close()
        return None

    origen_like = f"%{origen_texto}%"
    cur.execute("""
        SELECT DISTINCT nodo_a_id AS id, nodo_a_nombre AS nombre
        FROM hilos
        WHERE UPPER(nodo_a_nombre) LIKE UPPER(?)
        LIMIT 1
    """, (origen_like,))
    row = cur.fetchone()
    if not row:
        cur.execute("""
            SELECT DISTINCT nodo_b_id AS id, nodo_b_nombre AS nombre
            FROM hilos
            WHERE UPPER(nodo_b_nombre) LIKE UPPER(?)
            LIMIT 1
        """, (origen_like,))
        row = cur.fetchone()

    if not row:
        print(f"❌ No se encontró ningún nodo que contenga: '{origen_texto}'", file=sys.stderr)
        print(f"   Sugerencia: prueba con un texto más corto o verifica los nombres en el archivo.", file=sys.stderr)
        conn.close()
        return None

    start_id, start_name = row["id"], row["nombre"]
    print(f"[DIAG] Nodo origen encontrado: ID={start_id}, Nombre='{start_name}'", file=sys.stderr)

    visited = {start_id}
    queue = deque([(start_id, start_name, [])])

    while queue:
        nodo_id, nodo_nombre, camino = queue.popleft()

        if destino_texto.upper() in (nodo_nombre or "").upper():
            print(f"[DIAG] Destino encontrado: '{nodo_nombre}'", file=sys.stderr)
            return camino + [(None, nodo_id, nodo_nombre)]

        cur.execute("""
            SELECT DISTINCT id_cable
            FROM hilos
            WHERE nodo_a_id = ? OR nodo_b_id = ?
        """, (nodo_id, nodo_id))

        for (cable_id,) in cur.fetchall():
            cur.execute("""
                SELECT nodo_a_id, nodo_a_nombre, nodo_b_id, nodo_b_nombre
                FROM hilos
                WHERE id_cable = ?
                LIMIT 1
            """, (cable_id,))
            r = cur.fetchone()
            if not r:
                continue

            if r["nodo_a_id"] == nodo_id:
                otro_id, otro_nombre = r["nodo_b_id"], r["nodo_b_nombre"]
            else:
                otro_id, otro_nombre = r["nodo_a_id"], r["nodo_a_nombre"]

            if otro_id in visited:
                continue
            visited.add(otro_id)

            nuevo_camino = camino + [(cable_id, otro_id, otro_nombre)]
            queue.append((otro_id, otro_nombre, nuevo_camino))

    print("❌ No se encontró camino hacia el destino.", file=sys.stderr)
    print(f"   Se visitaron {len(visited)} nodos. Puede que el destino no esté conectado.", file=sys.stderr)
    conn.close()
    return None

# ---------------------------------------------------------------
# 3. GENERAR SALIDA (consola y archivo)
# ---------------------------------------------------------------
def generar_salida(db_path, camino, archivo_salida, origen, destino):
    if not camino:
        print("No hay ruta para mostrar.", file=sys.stderr)
        return

    conn = sqlite3.connect(db_path)
    cur = conn.cursor()

    # Construir el contenido de salida
    lineas = []
    lineas.append("\n=== 🗺️ RUTA (bastidores / empalmes) ===")
    ruta_str = ""
    for i, (cable_id, nodo_id, nodo_nombre) in enumerate(camino):
        if i == 0:
            ruta_str += nodo_nombre
        else:
            ruta_str += f" --[{cable_id}]--> {nodo_nombre}"
    lineas.append(ruta_str)
    lineas.append("")

    lineas.append("=== 🔍 CAMBIOS DE FIBRA EN LA RUTA (hilos con estado != LIBRE) ===")
    cables = [cable for cable, _, _ in camino if cable is not None]
    if not cables:
        lineas.append("No hay cables en la ruta (origen y destino son el mismo nodo).")
    else:
        for cable_id in cables:
            cur.execute("""
                SELECT num_hilo, cliente, estado, uso
                FROM hilos
                WHERE id_cable = ? AND UPPER(estado) != 'LIBRE'
            """, (cable_id,))
            filas = cur.fetchall()
            if filas:
                lineas.append(f"\n📡 Cable {cable_id}:")
                for num_hilo, cliente, estado, uso in filas:
                    lineas.append(f"  hilo {num_hilo:>3}  cliente={cliente:<20} "
                                  f"estado={estado:<12} uso={uso}")
            else:
                lineas.append(f"\n📡 Cable {cable_id}: (todos los hilos están LIBRES)")

    conn.close()

    # Imprimir en consola
    for linea in lineas:
        print(linea)

    # Escribir en archivo de salida
    try:
        with open(archivo_salida, "w", encoding="utf-8") as f:
            f.write(f"Origen: {origen}\nDestino: {destino}\n")
            for linea in lineas:
                f.write(linea + "\n")
        print(f"\n✅ Resultados guardados en: {archivo_salida}", file=sys.stderr)
    except Exception as e:
        print(f"❌ Error al escribir archivo de salida: {e}", file=sys.stderr)

# ---------------------------------------------------------------
# MAIN (automático, sin preguntas)
# ---------------------------------------------------------------
if __name__ == "__main__":
    ap = argparse.ArgumentParser(
        description="Trazador de fibra sobre archivos gigantes - modo automático",
        epilog="Sin argumentos: construye índice si no existe, luego traza ruta y guarda en 00_out.txt"
    )
    ap.add_argument("--build", action="store_true", help="Fuerza reconstrucción del índice")
    ap.add_argument("--trace", action="store_true", help="Traza la ruta usando el índice existente")
    ap.add_argument("--archivo", default=DEFAULT_ARCHIVO_ENTRADA, help=f"Archivo de entrada (default: {DEFAULT_ARCHIVO_ENTRADA})")
    ap.add_argument("--db", default=DEFAULT_DB, help=f"Archivo de índice (default: {DEFAULT_DB})")
    ap.add_argument("--output", default=DEFAULT_ARCHIVO_SALIDA, help=f"Archivo de salida (default: {DEFAULT_ARCHIVO_SALIDA})")
    ap.add_argument("--origen", default=DEFAULT_ORIGEN, help=f"Texto origen (default: {DEFAULT_ORIGEN})")
    ap.add_argument("--destino", default=DEFAULT_DESTINO, help=f"Texto destino (default: {DEFAULT_DESTINO})")
    ap.add_argument("--sep", default=DEFAULT_SEP, help=f"Separador (default: '{DEFAULT_SEP}')")
    ap.add_argument("--encoding", default=DEFAULT_ENCODING, help=f"Codificación (default: {DEFAULT_ENCODING})")
    args = ap.parse_args()

    # Comportamiento automático si no se dan argumentos:
    # - Si el índice no existe o se pide --build, construir.
    # - Luego, si el índice existe, trazar (ya sea por --trace o automáticamente).
    if not args.build and not args.trace:
        # Modo automático: construir si no existe el índice
        if not os.path.exists(args.db):
            print("🔧 Índice no encontrado. Construyendo automáticamente...", file=sys.stderr)
            args.build = True
        else:
            args.trace = True

    if args.build:
        ok = build_index(args.archivo, args.db, args.sep, args.encoding)
        if not ok:
            sys.exit(1)
        # Después de construir, trazar automáticamente
        args.trace = True

    if args.trace:
        camino = find_path(args.db, args.origen, args.destino)
        if camino:
            generar_salida(args.db, camino, args.output, args.origen, args.destino)
        else:
            print("❌ No se encontró ruta.", file=sys.stderr)
            sys.exit(1)