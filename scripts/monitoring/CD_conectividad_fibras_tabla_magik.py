import duckdb
import pandas as pd
import re
import os

RUTA_ENTRADA = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\smallworld\private_collections\00_out_Trace-connection.txt"
RUTA_SALIDA_EXCEL = r"C:\A_GS1_PROYECTOS\reporte_fibra.xlsx"
MEMORY_LIMIT = "12GB"
THREADS = 8

con = duckdb.connect("auditoria_fibra.duckdb")
con.execute(f"PRAGMA memory_limit='{MEMORY_LIMIT}'")
con.execute(f"PRAGMA threads={THREADS}")
con.execute("PRAGMA enable_object_cache")
con.execute("SET preserve_insertion_order=false")

create_table_sql = f"""
CREATE OR REPLACE TABLE fibra AS
SELECT
    SHEATH_ID,
    SHEATH_NAME AS FUNDA,
    SHEATH_OLDNAME AS NOMBRE_ANTIGUO,
    HILO,
    ID_HILO,
    "NOMBRE" AS NOMBRE,
    ESTADO,
    TOPOLOGIA,
    TECNOLOGIA,
    STATUS,
    USO,
    RANGO,
    SPLICE_ENTRADA_ID,
    SPLICE_ENTRADA_NOMBRE,
    SPLICE_SALIDA_ID,
    SPLICE_SALIDA_NOMBRE
FROM read_csv(
    '{RUTA_ENTRADA}',
    delim='|',
    header=True,
    ignore_errors=True,
    nullstr=['', '<none>'],
    quote='',
    escape='',
    strict_mode=False,
    columns={{
        'SHEATH_ID': 'VARCHAR',
        'SHEATH_NAME': 'VARCHAR',
        'SHEATH_OLDNAME': 'VARCHAR',
        'HILO': 'VARCHAR',
        'ID_HILO': 'VARCHAR',
        'NOMBRE': 'VARCHAR',
        'ESTADO': 'VARCHAR',
        'TOPOLOGIA': 'VARCHAR',
        'TECNOLOGIA': 'VARCHAR',
        'STATUS': 'VARCHAR',
        'USO': 'VARCHAR',
        'RANGO': 'VARCHAR',
        'SPLICE_ENTRADA_ID': 'VARCHAR',
        'SPLICE_ENTRADA_NOMBRE': 'VARCHAR',
        'SPLICE_SALIDA_ID': 'VARCHAR',
        'SPLICE_SALIDA_NOMBRE': 'VARCHAR'
    }}
);
"""
con.execute(create_table_sql)

query = """
WITH
hilos_filtrados AS (
    SELECT
        SHEATH_ID,
        FUNDA,
        NOMBRE_ANTIGUO,
        TRY_CAST(HILO AS INTEGER) AS HILO_NUM,
        NOMBRE,
        NULLIF(TRIM(REPLACE(NOMBRE, '''', '')), '') AS NOMBRE_LIMPIO,
        STATUS,
        RANGO,
        SPLICE_ENTRADA_ID,
        SPLICE_ENTRADA_NOMBRE,
        SPLICE_SALIDA_ID,
        SPLICE_SALIDA_NOMBRE,
        CASE
            WHEN STATUS = 'Ocupado' OR (STATUS = 'Libre' AND NOMBRE_LIMPIO IS NOT NULL)
            THEN 'OCUPADO'
            ELSE 'LIBRE'
        END AS estado_real,
        CASE
            WHEN STATUS = 'Libre' AND NOMBRE_LIMPIO IS NOT NULL AND UPPER(NOMBRE_LIMPIO) != 'RETIRO'
            THEN HILO || ':' || NOMBRE_LIMPIO
            ELSE NULL
        END AS error_libre_con_nombre
    FROM fibra
),

sheath_agg AS (
    SELECT
        FUNDA,
        SHEATH_ID,
        STRING_AGG(DISTINCT NOMBRE_ANTIGUO, ', ') AS NOMBRES_ANTIGUOS,
        COUNT(*) AS total_hilos,
        STRING_AGG(
            CASE WHEN estado_real = 'OCUPADO' AND HILO_NUM IS NOT NULL
                 THEN CAST(HILO_NUM AS VARCHAR)
                 ELSE NULL END,
            ', ' ORDER BY HILO_NUM
        ) AS hilos_ocupados_str,
        STRING_AGG(
            DISTINCT CASE WHEN estado_real = 'OCUPADO' AND NOMBRE_LIMPIO IS NOT NULL
                          THEN NOMBRE_LIMPIO
                          ELSE NULL END,
            ', '
        ) AS clientes_sheath,
        STRING_AGG(
            DISTINCT error_libre_con_nombre,
            ', '
        ) AS errores_libre_con_nombre,
        STRING_AGG(DISTINCT SPLICE_ENTRADA_ID, ', ') AS splice_entrada_ids,
        STRING_AGG(DISTINCT SPLICE_ENTRADA_NOMBRE, ', ') AS splice_entrada_nombres,
        STRING_AGG(DISTINCT SPLICE_SALIDA_ID, ', ') AS splice_salida_ids,
        STRING_AGG(DISTINCT SPLICE_SALIDA_NOMBRE, ', ') AS splice_salida_nombres
    FROM hilos_filtrados
    GROUP BY FUNDA, SHEATH_ID
),

sheath_valid AS (
    SELECT
        FUNDA,
        SHEATH_ID,
        NOMBRES_ANTIGUOS,
        total_hilos,
        hilos_ocupados_str,
        clientes_sheath,
        errores_libre_con_nombre,
        splice_entrada_ids,
        splice_entrada_nombres,
        splice_salida_ids,
        splice_salida_nombres,
        CASE WHEN total_hilos != 24 THEN 'ERROR_HILOS' ELSE 'OK' END AS validacion_hilos
    FROM sheath_agg
),

total_sheath_por_cable AS (
    SELECT FUNDA, COUNT(*) AS TOTAL_SHEATHS
    FROM sheath_valid
    GROUP BY FUNDA
),

clientes_expandidos AS (
    SELECT
        s.FUNDA,
        s.SHEATH_ID,
        UNNEST(STRING_SPLIT(s.clientes_sheath, ', ')) AS cliente
    FROM sheath_valid s
    WHERE s.clientes_sheath IS NOT NULL AND s.clientes_sheath != ''
),

clientes_comunes AS (
    SELECT
        c.FUNDA,
        c.cliente,
        COUNT(DISTINCT c.SHEATH_ID) AS sheath_con_cliente,
        MAX(t.TOTAL_SHEATHS) AS total_sheaths_cable
    FROM clientes_expandidos c
    JOIN total_sheath_por_cable t ON c.FUNDA = t.FUNDA
    GROUP BY c.FUNDA, c.cliente
),

cable_con_comun AS (
    SELECT
        FUNDA,
        MAX(CASE WHEN sheath_con_cliente = total_sheaths_cable THEN 1 ELSE 0 END) AS tiene_cliente_comun
    FROM clientes_comunes
    GROUP BY FUNDA
),

cable_agrupado AS (
    SELECT
        s.FUNDA,
        STRING_AGG(s.SHEATH_ID, ', ' ORDER BY s.SHEATH_ID) AS SHEATHS_IDS,
        STRING_AGG(DISTINCT s.NOMBRES_ANTIGUOS, ', ') AS NOMBRES_ANTIGUOS,
        LIST(s.total_hilos ORDER BY s.SHEATH_ID) AS lista_hilos,
        LIST(s.hilos_ocupados_str ORDER BY s.SHEATH_ID) AS lista_ocupados,
        STRING_AGG(DISTINCT s.clientes_sheath, ', ') AS CLIENTES_TODOS,
        STRING_AGG(s.validacion_hilos, ', ' ORDER BY s.SHEATH_ID) AS VALIDACIONES,
        COALESCE(c.tiene_cliente_comun, 0) AS tiene_cliente_comun,
        STRING_AGG(DISTINCT s.errores_libre_con_nombre, ', ') AS ERRORES_LIBRE_CON_NOMBRE,
        STRING_AGG(DISTINCT s.splice_entrada_ids, ', ') AS SPLICE_ENTRADA_IDS,
        STRING_AGG(DISTINCT s.splice_entrada_nombres, ', ') AS SPLICE_ENTRADA_NOMBRES,
        STRING_AGG(DISTINCT s.splice_salida_ids, ', ') AS SPLICE_SALIDA_IDS,
        STRING_AGG(DISTINCT s.splice_salida_nombres, ', ') AS SPLICE_SALIDA_NOMBRES
    FROM sheath_valid s
    LEFT JOIN cable_con_comun c ON s.FUNDA = c.FUNDA
    GROUP BY s.FUNDA, c.tiene_cliente_comun
),

resultado AS (
    SELECT
        FUNDA,
        REPLACE(REPLACE(SHEATHS_IDS, CHR(10), ''), CHR(13), '') AS SHEATHS_IDS,
        REPLACE(REPLACE(NOMBRES_ANTIGUOS, CHR(10), ''), CHR(13), '') AS NOMBRES_ANTIGUOS,
        CASE
            WHEN LEN(LIST_DISTINCT(lista_hilos)) = 1
            THEN CAST(LIST_DISTINCT(lista_hilos)[1] AS VARCHAR)
            ELSE array_to_string(lista_hilos, ', ')
        END AS HILOS_POR_SHEATH,
        CASE
            WHEN LEN(LIST_DISTINCT(lista_ocupados)) = 1
            THEN LIST_DISTINCT(lista_ocupados)[1]
            ELSE array_to_string(lista_ocupados, ', ')
        END AS HILOS_OCUPADOS_INDIVIDUALES,
        REPLACE(REPLACE(CLIENTES_TODOS, CHR(10), ''), CHR(13), '') AS CLIENTES,
        REPLACE(REPLACE(SPLICE_ENTRADA_IDS, CHR(10), ''), CHR(13), '') AS SPLICE_ENTRADA_IDS,
        REPLACE(REPLACE(SPLICE_ENTRADA_NOMBRES, CHR(10), ''), CHR(13), '') AS SPLICE_ENTRADA_NOMBRES,
        REPLACE(REPLACE(SPLICE_SALIDA_IDS, CHR(10), ''), CHR(13), '') AS SPLICE_SALIDA_IDS,
        REPLACE(REPLACE(SPLICE_SALIDA_NOMBRES, CHR(10), ''), CHR(13), '') AS SPLICE_SALIDA_NOMBRES,
        CASE
            WHEN VALIDACIONES LIKE '%ERROR_HILOS%'
            THEN 'RECHAZADO - HILOS INCONSISTENTES'
            WHEN ERRORES_LIBRE_CON_NOMBRE IS NOT NULL AND ERRORES_LIBRE_CON_NOMBRE != ''
            THEN 'RECHAZADO - LIBRE CON NOMBRE | ' || REPLACE(REPLACE(ERRORES_LIBRE_CON_NOMBRE, CHR(10), ''), CHR(13), '')
            WHEN LENGTH(CLIENTES_TODOS) > 0 AND CLIENTES_TODOS LIKE '%,%' AND CLIENTES_TODOS NOT LIKE '%SCJ%'
            THEN 'RECHAZADO - CLIENTES DIFERENTES'
            WHEN CLIENTES_TODOS IS NULL OR CLIENTES_TODOS = ''
            THEN 'SIN CLIENTES'
            ELSE 'APROBADO'
        END AS ESTADO_CALIDAD
    FROM cable_agrupado
)
SELECT * FROM resultado ORDER BY FUNDA;
"""

df_resumen = con.execute(query).fetchdf()

def clean_text(value):
    if isinstance(value, str):
        return re.sub(r'[\x00-\x1f\x7f-\x9f]', '', value)
    return value

df_resumen = df_resumen.applymap(clean_text)

with pd.ExcelWriter(RUTA_SALIDA_EXCEL, engine='openpyxl') as writer:
    df_resumen.to_excel(writer, sheet_name='Resumen', index=False)

con.close()


