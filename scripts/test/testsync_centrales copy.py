"""
SISTEMA DE SINCRONIZACION DE DATOS
Version: 20.5 - CON HOJA ANILLO
"""

import duckdb
import pandas as pd
import os
import sys
from datetime import datetime
import logging
import traceback
from openpyxl import load_workbook

logging.basicConfig(
    level=logging.INFO,
    format='%(message)s',
    handlers=[
        logging.FileHandler('sincronizacion.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

FECHA_INICIO = "01/08/2026"
FECHA_FIN = "10/08/2026"

RUTA_REPORTE_FIBRA = r"C:\A_GS1_PROYECTOS\reporte_fibra.xlsx"
RUTA_CLIENTES = r"C:\A_GS1_PROYECTOS\kineabaseetbClientes priorizados conectividad.xlsx"
RUTA_PORTS = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk export of ports.xlsx"
RUTA_PREPROCESAMIENTO = r"C:\A_GS1_PROYECTOS\preprocesamiento.xlsx"
RUTA_CENTRALES = r"C:\A_GS1_PROYECTOS\01-Bulk Merge_Centrales.xlsx"
RUTA_REPORTE_ANILLOS = r"C:\Users\gabrboa1\Documents\Reporte anillos conectados.xlsx"
RUTA_SALIDA = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\04-Bulk Merge_Clientes.xlsx"

MEMORY_LIMIT = '16GB'
THREADS = 8


class SincronizadorDatos:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = self._parse_fecha(fecha_inicio)
        self.fecha_fin = self._parse_fecha(fecha_fin)
        self.con = None
        self.df_resultado = None
        self.df_anillos = None
    
    def _parse_fecha(self, fecha_str):
        try:
            return datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            try:
                return datetime.strptime(fecha_str, "%Y-%m-%d")
            except:
                raise ValueError(f"Formato invalido: {fecha_str}")
    
    def _conectar_duckdb(self):
        self.con = duckdb.connect(":memory:")
        self.con.execute(f"SET memory_limit='{MEMORY_LIMIT}';")
        self.con.execute(f"SET threads={THREADS};")
        logger.info("Conectado a DuckDB")
    
    def _cargar_reporte_fibra(self):
        if not os.path.exists(RUTA_REPORTE_FIBRA):
            raise FileNotFoundError(f"Fibra no encontrado: {RUTA_REPORTE_FIBRA}")
        
        logger.info("Cargando fibra...")
        df_fibra = pd.read_excel(RUTA_REPORTE_FIBRA, sheet_name='Resumen', dtype=str)
        self.con.register('fibra_temp', df_fibra)
        
        self.con.execute("""
            CREATE OR REPLACE TABLE fibra_raw AS
            WITH expanded AS (
                SELECT 
                    UNNEST(STRING_SPLIT(REPLACE(SHEATHS_IDS, ';', ','), ',')) AS SHEATH_ID_UNICO,
                    UNNEST(STRING_SPLIT(REPLACE(NOMBRES_ANTIGUOS, ';', ','), ',')) AS NOMBRE_ANTIGUO,
                    FUNDA,
                    HILOS_OCUPADOS_INDIVIDUALES
                FROM fibra_temp
                WHERE SHEATHS_IDS IS NOT NULL 
                  AND SHEATHS_IDS != 'nan'
                  AND SHEATHS_IDS != 'None'
                  AND NOMBRES_ANTIGUOS IS NOT NULL
                  AND NOMBRES_ANTIGUOS != 'nan'
                  AND NOMBRES_ANTIGUOS != 'None'
            )
            SELECT DISTINCT
                TRIM(SHEATH_ID_UNICO) AS SHEATH_ID_UNICO,
                TRIM(NOMBRE_ANTIGUO) AS NOMBRE_ANTIGUO,
                FUNDA,
                HILOS_OCUPADOS_INDIVIDUALES,
                REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(NOMBRE_ANTIGUO)), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS NOMBRE_NORMALIZADO
            FROM expanded
            WHERE TRIM(SHEATH_ID_UNICO) != '' 
              AND TRIM(SHEATH_ID_UNICO) NOT IN ('nan', 'None')
              AND TRIM(NOMBRE_ANTIGUO) != ''
              AND TRIM(NOMBRE_ANTIGUO) NOT IN ('nan', 'None')
        """)
        
        count = self.con.execute("SELECT COUNT(*) FROM fibra_raw").fetchone()[0]
        logger.info(f"Fibra: {count:,} registros")
        del df_fibra
        return count
    
    def _cargar_preprocesamiento(self):
        if not os.path.exists(RUTA_PREPROCESAMIENTO):
            logger.warning("Preprocesamiento no encontrado")
            return 0
        
        try:
            logger.info("Cargando preprocesamiento...")
            df_pre = pd.read_excel(RUTA_PREPROCESAMIENTO, sheet_name='Hoja2', dtype=str)
            df_pre.columns = df_pre.columns.str.strip().str.replace(' ', '_')
            
            col_cable = None
            for col in df_pre.columns:
                if 'cable' in col.lower() or 'acceso' in col.lower():
                    col_cable = col
                    break
            
            if not col_cable:
                logger.warning("No se encontro columna de cable en preprocesamiento")
                return 0
            
            self.con.register('pre_temp', df_pre)
            
            self.con.execute(f"""
                CREATE OR REPLACE TABLE preprocesamiento_raw AS
                WITH expanded AS (
                    SELECT 
                        *,
                        UNNEST(STRING_SPLIT(REPLACE("{col_cable}", ';', ','), ',')) AS CABLE_UNICO,
                        ROW_NUMBER() OVER () AS ID_PRE
                    FROM pre_temp
                    WHERE "{col_cable}" IS NOT NULL 
                      AND "{col_cable}" != 'nan'
                      AND "{col_cable}" != 'None'
                      AND "{col_cable}" != ''
                )
                SELECT 
                    ID_PRE,
                    TRIM(CABLE_UNICO) AS Cable_de_Acceso,
                    REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CABLE_UNICO)), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO
                FROM expanded
                WHERE TRIM(CABLE_UNICO) != '' 
                  AND TRIM(CABLE_UNICO) NOT IN ('nan', 'None')
            """)
            
            count = self.con.execute("SELECT COUNT(*) FROM preprocesamiento_raw").fetchone()[0]
            logger.info(f"Preprocesamiento: {count:,} registros")
            del df_pre
            return count
            
        except Exception as e:
            logger.error(f"Error preprocesamiento: {e}")
            return 0
    
    def _cargar_centrales(self):
        if not os.path.exists(RUTA_CENTRALES):
            logger.warning("Centrales no encontrado")
            return 0
        
        try:
            logger.info("Cargando centrales...")
            
            # Cargar hoja 'bus'
            df_bus = pd.read_excel(RUTA_CENTRALES, sheet_name='bus', dtype=str)
            df_bus.columns = df_bus.columns.str.strip().str.replace(' ', '_')
            self.con.register('centrales_bus_temp', df_bus)
            
            # Cargar hoja 'anillo'
            df_anillo = pd.read_excel(RUTA_CENTRALES, sheet_name='anillo', dtype=str)
            df_anillo.columns = df_anillo.columns.str.strip().str.replace(' ', '_')
            self.con.register('centrales_anillo_temp', df_anillo)
            
            # Detectar columnas en bus
            col_anillo_bus = None
            col_cable_bus = None
            col_fibra_bus = None
            
            for col in df_bus.columns:
                col_lower = col.lower()
                if 'anillo' in col_lower or 'bocu' in col_lower:
                    col_anillo_bus = col
                if 'acceso' in col_lower or 'cable' in col_lower or 'meacbl' in col_lower:
                    if col_cable_bus is None:
                        col_cable_bus = col
                if 'fibra' in col_lower:
                    col_fibra_bus = col
            
            if not col_anillo_bus:
                for col in df_bus.columns:
                    if 'anillo' in col.lower():
                        col_anillo_bus = col
                        break
            
            if not col_cable_bus:
                for col in df_bus.columns:
                    if 'acceso' in col.lower() or 'cable' in col.lower():
                        col_cable_bus = col
                        break
            
            if not col_fibra_bus:
                for col in df_bus.columns:
                    if 'fibra' in col.lower():
                        col_fibra_bus = col
                        break
            
            logger.info(f"BUS: anillo='{col_anillo_bus}', cable='{col_cable_bus}', fibra='{col_fibra_bus}'")
            
            # Detectar columnas en anillo (misma estructura que bus)
            col_anillo_anillo = None
            col_cable_anillo = None
            col_fibra_anillo = None
            
            for col in df_anillo.columns:
                col_lower = col.lower()
                if 'anillo' in col_lower or 'bocu' in col_lower:
                    col_anillo_anillo = col
                if 'acceso' in col_lower or 'cable' in col_lower or 'meacbl' in col_lower:
                    if col_cable_anillo is None:
                        col_cable_anillo = col
                if 'fibra' in col_lower:
                    col_fibra_anillo = col
            
            if not col_anillo_anillo:
                for col in df_anillo.columns:
                    if 'anillo' in col.lower():
                        col_anillo_anillo = col
                        break
            
            if not col_cable_anillo:
                for col in df_anillo.columns:
                    if 'acceso' in col.lower() or 'cable' in col.lower():
                        col_cable_anillo = col
                        break
            
            if not col_fibra_anillo:
                for col in df_anillo.columns:
                    if 'fibra' in col.lower():
                        col_fibra_anillo = col
                        break
            
            logger.info(f"ANILLO: anillo='{col_anillo_anillo}', cable='{col_cable_anillo}', fibra='{col_fibra_anillo}'")
            
            # Crear tabla bus_raw
            if col_anillo_bus and col_cable_bus:
                campos_bus = [
                    f'TRIM(CAST("{col_anillo_bus}" AS VARCHAR)) AS ANILLO_ORIGINAL',
                    f"REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST(\"{col_anillo_bus}\" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS ANILLO_NORMALIZADO",
                    f'TRIM(CAST("{col_cable_bus}" AS VARCHAR)) AS CABLE_ORIGINAL',
                    f"REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST(\"{col_cable_bus}\" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO"
                ]
                
                if col_fibra_bus:
                    campos_bus.append(f'TRIM(CAST("{col_fibra_bus}" AS VARCHAR)) AS FIBRA')
                
                query_bus = f"""
                CREATE OR REPLACE TABLE centrales_bus_raw AS
                SELECT 
                    ROW_NUMBER() OVER () AS ID_CENTRAL,
                    {', '.join(campos_bus)}
                FROM centrales_bus_temp
                WHERE CAST("{col_anillo_bus}" AS VARCHAR) IS NOT NULL
                  AND CAST("{col_anillo_bus}" AS VARCHAR) != 'nan'
                  AND CAST("{col_anillo_bus}" AS VARCHAR) != 'None'
                  AND CAST("{col_anillo_bus}" AS VARCHAR) != ''
                  AND CAST("{col_cable_bus}" AS VARCHAR) IS NOT NULL
                  AND CAST("{col_cable_bus}" AS VARCHAR) != 'nan'
                  AND CAST("{col_cable_bus}" AS VARCHAR) != 'None'
                  AND CAST("{col_cable_bus}" AS VARCHAR) != ''
                """
                self.con.execute(query_bus)
                count_bus = self.con.execute("SELECT COUNT(*) FROM centrales_bus_raw").fetchone()[0]
                logger.info(f"Centrales BUS: {count_bus:,} registros")
            else:
                logger.warning("No se encontraron columnas en BUS")
                count_bus = 0
            
            # Crear tabla anillo_raw (misma estructura que bus)
            if col_anillo_anillo and col_cable_anillo:
                campos_anillo = [
                    f'TRIM(CAST("{col_anillo_anillo}" AS VARCHAR)) AS ANILLO_ORIGINAL',
                    f"REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST(\"{col_anillo_anillo}\" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS ANILLO_NORMALIZADO",
                    f'TRIM(CAST("{col_cable_anillo}" AS VARCHAR)) AS CABLE_ORIGINAL',
                    f"REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST(\"{col_cable_anillo}\" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO"
                ]
                
                if col_fibra_anillo:
                    campos_anillo.append(f'TRIM(CAST("{col_fibra_anillo}" AS VARCHAR)) AS FIBRA')
                
                query_anillo = f"""
                CREATE OR REPLACE TABLE centrales_anillo_raw AS
                SELECT 
                    ROW_NUMBER() OVER () AS ID_CENTRAL,
                    {', '.join(campos_anillo)}
                FROM centrales_anillo_temp
                WHERE CAST("{col_anillo_anillo}" AS VARCHAR) IS NOT NULL
                  AND CAST("{col_anillo_anillo}" AS VARCHAR) != 'nan'
                  AND CAST("{col_anillo_anillo}" AS VARCHAR) != 'None'
                  AND CAST("{col_anillo_anillo}" AS VARCHAR) != ''
                  AND CAST("{col_cable_anillo}" AS VARCHAR) IS NOT NULL
                  AND CAST("{col_cable_anillo}" AS VARCHAR) != 'nan'
                  AND CAST("{col_cable_anillo}" AS VARCHAR) != 'None'
                  AND CAST("{col_cable_anillo}" AS VARCHAR) != ''
                """
                self.con.execute(query_anillo)
                count_anillo = self.con.execute("SELECT COUNT(*) FROM centrales_anillo_raw").fetchone()[0]
                logger.info(f"Centrales ANILLO: {count_anillo:,} registros")
            else:
                logger.warning("No se encontraron columnas en ANILLO")
                count_anillo = 0
            
            del df_bus
            del df_anillo
            return count_bus + count_anillo
            
        except Exception as e:
            logger.error(f"Error centrales: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def _cargar_reporte_anillos(self):
        """Carga el reporte de anillos conectados si existe"""
        if not os.path.exists(RUTA_REPORTE_ANILLOS):
            logger.info("Reporte anillos conectados no encontrado (opcional)")
            return 0
        
        try:
            logger.info("Cargando reporte anillos conectados...")
            df_anillos = pd.read_excel(RUTA_REPORTE_ANILLOS, dtype=str)
            df_anillos.columns = df_anillos.columns.str.strip().str.replace(' ', '_')
            
            logger.info(f"  Columnas: {list(df_anillos.columns)}")
            
            col_anillo = None
            col_cable_acceso = None
            col_hilo_troncal = None
            col_odf_troncal = None
            col_id_servicio = None
            col_cliente = None
            col_troncal = None
            
            for col in df_anillos.columns:
                col_lower = col.lower()
                if 'anillo' in col_lower and not col_anillo:
                    col_anillo = col
                if 'cable_acceso' in col_lower or 'cable acceso' in col_lower:
                    col_cable_acceso = col
                if 'hilo_troncal' in col_lower or 'hilo troncal' in col_lower:
                    col_hilo_troncal = col
                if 'odf_troncal_salida' in col_lower or 'odf troncal salida' in col_lower:
                    col_odf_troncal = col
                if 'id_servicio' in col_lower or 'id servicio' in col_lower:
                    col_id_servicio = col
                if 'cliente' in col_lower:
                    col_cliente = col
                if 'troncal' in col_lower and 'hilo' not in col_lower and 'odf' not in col_lower:
                    col_troncal = col
            
            if not col_odf_troncal:
                for col in df_anillos.columns:
                    if 'odf' in col.lower() and 'troncal' in col.lower():
                        col_odf_troncal = col
                        break
            
            if not col_anillo:
                for col in df_anillos.columns:
                    if 'anillo' in col.lower():
                        col_anillo = col
                        break
            
            if not col_cable_acceso:
                for col in df_anillos.columns:
                    if 'cable' in col.lower() and 'acceso' in col.lower():
                        col_cable_acceso = col
                        break
            
            if not col_troncal:
                for col in df_anillos.columns:
                    if col.lower() == 'troncal':
                        col_troncal = col
                        break
            
            logger.info(f"  Anillo: '{col_anillo}'")
            logger.info(f"  Cable Acceso: '{col_cable_acceso}'")
            logger.info(f"  Hilo Troncal: '{col_hilo_troncal}'")
            logger.info(f"  ODF Troncal Salida: '{col_odf_troncal}'")
            logger.info(f"  TRONCAL: '{col_troncal}'")
            
            if not col_anillo or not col_cable_acceso:
                logger.warning("No se encontraron columnas necesarias en reporte anillos")
                return 0
            
            self.con.register('anillos_temp', df_anillos)
            
            campos = []
            
            # ANILLO normalizado para cruce
            campos.append(f"""
                REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST("{col_anillo}" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS ANILLO_NORMALIZADO
            """)
            
            # CABLE ACCESO normalizado para cruce
            if col_cable_acceso:
                campos.append(f"""
                    REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST("{col_cable_acceso}" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO
                """)
                campos.append(f"TRIM(CAST(\"{col_cable_acceso}\" AS VARCHAR)) AS CABLE_ORIGINAL_ANILLOS")
            else:
                campos.append("'' AS CABLE_NORMALIZADO")
                campos.append("'' AS CABLE_ORIGINAL_ANILLOS")
            
            # HILO TRONCAL
            if col_hilo_troncal:
                campos.append(f"TRIM(CAST(\"{col_hilo_troncal}\" AS VARCHAR)) AS HILO_TRONCAL")
            else:
                campos.append("'' AS HILO_TRONCAL")
            
            # ODF TRONCAL SALIDA
            if col_odf_troncal:
                campos.append(f"TRIM(CAST(\"{col_odf_troncal}\" AS VARCHAR)) AS ODF_TRONCAL")
            else:
                campos.append("'' AS ODF_TRONCAL")
            
            # TRONCAL (para cruce con fibra de centrales)
            if col_troncal:
                campos.append(f"TRIM(CAST(\"{col_troncal}\" AS VARCHAR)) AS TRONCAL")
            else:
                campos.append("'' AS TRONCAL")
            
            # ID Servicio
            if col_id_servicio:
                campos.append(f"TRIM(CAST(\"{col_id_servicio}\" AS VARCHAR)) AS ID_SERVICIO")
            else:
                campos.append("'' AS ID_SERVICIO")
            
            # Cliente
            if col_cliente:
                campos.append(f"TRIM(CAST(\"{col_cliente}\" AS VARCHAR)) AS CLIENTE")
            else:
                campos.append("'' AS CLIENTE")
            
            query = f"""
            CREATE OR REPLACE TABLE reporte_anillos_raw AS
            SELECT 
                ROW_NUMBER() OVER () AS ID_ANILLO,
                {', '.join(campos)}
            FROM anillos_temp
            WHERE CAST("{col_anillo}" AS VARCHAR) IS NOT NULL
              AND CAST("{col_anillo}" AS VARCHAR) != 'nan'
              AND CAST("{col_anillo}" AS VARCHAR) != 'None'
              AND CAST("{col_anillo}" AS VARCHAR) != ''
              AND CAST("{col_cable_acceso}" AS VARCHAR) IS NOT NULL
              AND CAST("{col_cable_acceso}" AS VARCHAR) != 'nan'
              AND CAST("{col_cable_acceso}" AS VARCHAR) != 'None'
              AND CAST("{col_cable_acceso}" AS VARCHAR) != ''
            """
            
            self.con.execute(query)
            
            count = self.con.execute("SELECT COUNT(*) FROM reporte_anillos_raw").fetchone()[0]
            logger.info(f"Reporte anillos: {count:,} registros")
            
            del df_anillos
            return count
            
        except Exception as e:
            logger.error(f"Error cargando reporte anillos: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def _cargar_datos_clientes(self):
        if not os.path.exists(RUTA_CLIENTES):
            raise FileNotFoundError(f"Clientes no encontrado: {RUTA_CLIENTES}")
        
        logger.info("Cargando clientes...")
        df_clientes = pd.read_excel(RUTA_CLIENTES, dtype=str)
        df_clientes.columns = df_clientes.columns.str.strip().str.replace(' ', '_')
        self.con.register('clientes_temp', df_clientes)
        
        col_fecha = None
        for col in df_clientes.columns:
            if col.strip() in ['Fecha', 'FECHA', 'fecha']:
                col_fecha = col
                break
        
        col_cable = None
        for col in df_clientes.columns:
            if 'Cable_de_Acceso' in col or 'CABLE_DE_ACCESO' in col:
                col_cable = col
                break
            if 'cable' in col.lower() or 'acceso' in col.lower():
                col_cable = col
        
        if not col_cable:
            for col in df_clientes.columns:
                if 'cable' in col.lower():
                    col_cable = col
                    break
        
        if not col_cable:
            raise ValueError("No se encontro columna de Cable_de_Acceso")
        
        col_anillo = None
        for col in df_clientes.columns:
            if col.strip() in ['Anillo', 'ANILLO', 'anillo']:
                col_anillo = col
                break
            if 'anillo' in col.lower():
                col_anillo = col
                break
        
        fecha_inicio_str = self.fecha_inicio.strftime('%Y-%m-%d')
        fecha_fin_str = self.fecha_fin.strftime('%Y-%m-%d')
        
        if col_anillo:
            anillo_select = f"""
                TRIM(CAST("{col_anillo}" AS VARCHAR)) AS ANILLO_ORIGINAL,
                REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CAST("{col_anillo}" AS VARCHAR))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS ANILLO_NORMALIZADO
            """
        else:
            anillo_select = "'' AS ANILLO_ORIGINAL, '' AS ANILLO_NORMALIZADO"
        
        if col_fecha:
            query = f"""
            CREATE OR REPLACE TABLE clientes_raw AS
            WITH filtered AS (
                SELECT *
                FROM clientes_temp
                WHERE TRY_CAST("{col_fecha}" AS DATE) IS NULL 
                   OR (TRY_CAST("{col_fecha}" AS DATE) BETWEEN '{fecha_inicio_str}' AND '{fecha_fin_str}')
            ),
            expanded AS (
                SELECT 
                    *,
                    UNNEST(STRING_SPLIT(REPLACE("{col_cable}", ';', ','), ',')) AS CABLE_UNICO,
                    ROW_NUMBER() OVER () AS ID_ORIGINAL
                FROM filtered
                WHERE "{col_cable}" IS NOT NULL 
                  AND "{col_cable}" != 'nan'
                  AND "{col_cable}" != 'None'
                  AND "{col_cable}" != ''
            )
            SELECT 
                ID_ORIGINAL,
                TRIM(CABLE_UNICO) AS Cable_de_Acceso,
                {anillo_select},
                Nombre AS NOMBRE_CLIENTE,
                Anillo AS ANILLO_VISIBLE,
                IDServicio,
                Responsable,
                Estatus,
                Observaciones,
                "{col_fecha}" AS FECHA,
                REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CABLE_UNICO)), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO
            FROM expanded
            WHERE TRIM(CABLE_UNICO) != '' 
              AND TRIM(CABLE_UNICO) NOT IN ('nan', 'None')
            """
        else:
            query = f"""
            CREATE OR REPLACE TABLE clientes_raw AS
            WITH expanded AS (
                SELECT 
                    *,
                    UNNEST(STRING_SPLIT(REPLACE("{col_cable}", ';', ','), ',')) AS CABLE_UNICO,
                    ROW_NUMBER() OVER () AS ID_ORIGINAL
                FROM clientes_temp
                WHERE "{col_cable}" IS NOT NULL 
                  AND "{col_cable}" != 'nan'
                  AND "{col_cable}" != 'None'
                  AND "{col_cable}" != ''
            )
            SELECT 
                ID_ORIGINAL,
                TRIM(CABLE_UNICO) AS Cable_de_Acceso,
                {anillo_select},
                Nombre AS NOMBRE_CLIENTE,
                Anillo AS ANILLO_VISIBLE,
                IDServicio,
                Responsable,
                Estatus,
                Observaciones,
                '' AS FECHA,
                REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(CABLE_UNICO)), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS CABLE_NORMALIZADO
            FROM expanded
            WHERE TRIM(CABLE_UNICO) != '' 
              AND TRIM(CABLE_UNICO) NOT IN ('nan', 'None')
            """
        
        self.con.execute(query)
        count = self.con.execute("SELECT COUNT(*) FROM clientes_raw").fetchone()[0]
        logger.info(f"Clientes: {count:,} registros")
        del df_clientes
        return count
    
    def _cargar_datos_ports(self):
        if not os.path.exists(RUTA_PORTS):
            logger.warning("Ports no encontrado")
            return 0
        
        try:
            logger.info("Cargando ports...")
            
            df_sample = pd.read_excel(RUTA_PORTS, sheet_name='Port', nrows=5, dtype=str)
            df_sample.columns = df_sample.columns.str.strip()
            
            col_id = None
            col_servicio = None
            col_description = None
            col_description_shelf = None
            col_physical_status = None
            
            for col in df_sample.columns:
                col_clean = col.strip()
                col_lower = col_clean.lower()
                
                if col_lower == 'id':
                    col_id = col
                elif col_lower == 'id servicio':
                    col_servicio = col
                elif col_clean == 'Description':
                    col_description = col
                elif col_clean == 'Description Shelf':
                    col_description_shelf = col
                elif col_clean == 'Physical Status' or col_lower == 'physical status':
                    col_physical_status = col
            
            if not col_id:
                for col in df_sample.columns:
                    if col.lower() == 'id':
                        col_id = col
                        break
            
            if not col_servicio:
                for col in df_sample.columns:
                    if 'id servicio' in col.lower():
                        col_servicio = col
                        break
            
            if not col_description:
                for col in df_sample.columns:
                    if col == 'Description':
                        col_description = col
                        break
            
            if not col_description_shelf:
                for col in df_sample.columns:
                    if col == 'Description Shelf':
                        col_description_shelf = col
                        break
                if not col_description_shelf:
                    for col in df_sample.columns:
                        if 'description shelf' in col.lower():
                            col_description_shelf = col
                            break
            
            if not col_physical_status:
                for col in df_sample.columns:
                    if 'physical status' in col.lower():
                        col_physical_status = col
                        break
            
            logger.info(f"  ID: '{col_id}'")
            logger.info(f"  Servicio: '{col_servicio}'")
            logger.info(f"  Description: '{col_description}'")
            logger.info(f"  Description Shelf: '{col_description_shelf}'")
            logger.info(f"  Physical Status: '{col_physical_status}'")
            
            if not col_id:
                logger.warning("No se encontro columna ID")
                return 0
            
            if not col_servicio:
                logger.warning("No se encontro columna de servicio")
                return 0
            
            columnas_a_leer = [col_id, col_servicio]
            if col_description:
                columnas_a_leer.append(col_description)
            if col_description_shelf:
                columnas_a_leer.append(col_description_shelf)
            if col_physical_status:
                columnas_a_leer.append(col_physical_status)
            
            df_ports = pd.read_excel(RUTA_PORTS, sheet_name='Port', dtype=str, usecols=columnas_a_leer)
            logger.info(f"Ports cargados: {len(df_ports):,} filas")
            
            self.con.register('ports_temp', df_ports)
            
            campos = [
                f'CAST("{col_id}" AS VARCHAR) AS port_id',
                f'TRIM(CAST("{col_servicio}" AS VARCHAR)) AS servicio_id'
            ]
            
            if col_description:
                campos.append(f'TRIM(CAST("{col_description}" AS VARCHAR)) AS description_raw')
            else:
                campos.append("'' AS description_raw")
            
            if col_description_shelf:
                campos.append(f'TRIM(CAST("{col_description_shelf}" AS VARCHAR)) AS description_shelf_raw')
            else:
                campos.append("'' AS description_shelf_raw")
            
            if col_physical_status:
                campos.append(f'TRIM(CAST("{col_physical_status}" AS VARCHAR)) AS physical_status_raw')
            
            where_conditions = [
                f'CAST("{col_id}" AS VARCHAR) IS NOT NULL',
                f'CAST("{col_id}" AS VARCHAR) != \'nan\'',
                f'CAST("{col_id}" AS VARCHAR) != \'None\'',
                f'CAST("{col_id}" AS VARCHAR) != \'\'',
                f'CAST("{col_servicio}" AS VARCHAR) IS NOT NULL',
                f'CAST("{col_servicio}" AS VARCHAR) != \'nan\'',
                f'CAST("{col_servicio}" AS VARCHAR) != \'None\'',
                f'CAST("{col_servicio}" AS VARCHAR) != \'\''
            ]
            
            if col_physical_status:
                where_conditions.append(f'UPPER(TRIM(CAST("{col_physical_status}" AS VARCHAR))) LIKE UPPER(\'%In Service%\')')
                logger.info("  Aplicando filtro: Physical Status contiene 'In Service'")
            
            query_extracted = f"""
                CREATE OR REPLACE TABLE ports_extracted AS
                SELECT {', '.join(campos)}
                FROM ports_temp
                WHERE {' AND '.join(where_conditions)}
            """
            
            self.con.execute(query_extracted)
            
            count_extracted = self.con.execute("SELECT COUNT(*) FROM ports_extracted").fetchone()[0]
            logger.info(f"Registros extraídos (solo In Service): {count_extracted:,}")
            
            if count_extracted == 0:
                logger.warning("  No se encontraron puertos con Physical Status = 'In Service'")
                self.con.execute("""
                    CREATE OR REPLACE TABLE ports_mapping AS
                    SELECT 
                        '' AS servicio_normalizado,
                        [] AS puertos_ids,
                        '' AS DESCRIPTION,
                        '' AS DESCRIPTION_SHELF
                    WHERE 1=0
                """)
                del df_ports
                del df_sample
                return 0
            
            self.con.execute("""
                CREATE OR REPLACE TABLE ports_mapping AS
                SELECT 
                    REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(COALESCE(servicio_id, ''))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g') AS servicio_normalizado,
                    LIST(port_id) AS puertos_ids,
                    FIRST(description_raw) AS DESCRIPTION,
                    FIRST(description_shelf_raw) AS DESCRIPTION_SHELF
                FROM ports_extracted
                WHERE servicio_id IS NOT NULL
                  AND servicio_id != ''
                  AND servicio_id != 'nan'
                  AND servicio_id != 'None'
                GROUP BY servicio_normalizado
            """)
            
            count = self.con.execute("SELECT COUNT(*) FROM ports_mapping").fetchone()[0]
            logger.info(f"Servicios mapeados (solo In Service): {count:,}")
            
            del df_ports
            del df_sample
            return count
            
        except Exception as e:
            logger.error(f"Error cargando ports: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def _realizar_cruce(self):
        try:
            logger.info("Realizando cruce...")
            
            has_anillos = False
            try:
                count_anillos = self.con.execute("SELECT COUNT(*) FROM reporte_anillos_raw").fetchone()[0]
                has_anillos = count_anillos > 0
                logger.info(f"  Reporte anillos disponible: {count_anillos} registros")
            except:
                logger.info("  Reporte anillos no disponible")
            
            # Verificar si existe centrales_bus_raw
            has_bus = False
            try:
                count_bus = self.con.execute("SELECT COUNT(*) FROM centrales_bus_raw").fetchone()[0]
                has_bus = count_bus > 0
                logger.info(f"  Centrales BUS disponible: {count_bus} registros")
            except:
                logger.info("  Centrales BUS no disponible")
            
            # Verificar si existe centrales_anillo_raw
            has_anillo = False
            try:
                count_anillo = self.con.execute("SELECT COUNT(*) FROM centrales_anillo_raw").fetchone()[0]
                has_anillo = count_anillo > 0
                logger.info(f"  Centrales ANILLO disponible: {count_anillo} registros")
            except:
                logger.info("  Centrales ANILLO no disponible")
            
            if has_anillos:
                if has_bus and has_anillo:
                    # Bus y Anillo disponibles
                    self.con.execute("""
                        CREATE OR REPLACE TABLE resultado_cruce AS
                        SELECT 
                            c.NOMBRE_CLIENTE,
                            c.ANILLO_VISIBLE,
                            c.IDServicio,
                            c.Responsable,
                            c.Estatus,
                            c.Observaciones,
                            c.FECHA,
                            c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
                            f.SHEATH_ID_UNICO AS SHEATHS_IDS,
                            f.FUNDA,
                            f.HILOS_OCUPADOS_INDIVIDUALES,
                            CASE 
                                WHEN p.ID_PRE IS NOT NULL AND (ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL) 
                                    THEN 'PREPROCESAMIENTO Y CENTRALES'
                                WHEN p.ID_PRE IS NOT NULL THEN 'PREPROCESAMIENTO'
                                WHEN ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL THEN 'CENTRALES'
                                WHEN ra.ID_ANILLO IS NOT NULL AND ce_bus.ID_CENTRAL IS NOT NULL THEN 'BUS'
                                WHEN ra.ID_ANILLO IS NOT NULL AND ce_anillo_central.ID_CENTRAL IS NOT NULL THEN 'ANILLO'
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'REPORTE_ANILLOS'
                                ELSE 'SOLO PRINCIPAL'
                            END AS ORIGEN,
                            CASE 
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'SI' 
                                ELSE 'NO' 
                            END AS EN_REPORTE_ANILLOS,
                            ra.HILO_TRONCAL AS HILO_TRONCAL_ANILLOS,
                            ra.ODF_TRONCAL AS ODF_TRONCAL_ANILLOS,
                            ra.ANILLO_NORMALIZADO AS ANILLO_NORMALIZADO,
                            ra.CABLE_ORIGINAL_ANILLOS AS CABLE_ORIGINAL_ANILLOS,
                            ra.ID_SERVICIO AS ID_SERVICIO_ANILLOS,
                            ra.CLIENTE AS CLIENTE_ANILLOS,
                            ra.TRONCAL AS TRONCAL_ANILLOS,
                            CASE WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO' ELSE 'OK' END AS ERROR
                        FROM clientes_raw c
                        LEFT JOIN fibra_raw f ON c.CABLE_NORMALIZADO = f.NOMBRE_NORMALIZADO
                        LEFT JOIN preprocesamiento_raw p ON c.CABLE_NORMALIZADO = p.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce ON c.CABLE_NORMALIZADO = ce.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce_anillo ON c.ANILLO_NORMALIZADO = ce_anillo.ANILLO_NORMALIZADO 
                            AND c.ANILLO_NORMALIZADO != '' 
                            AND ce_anillo.ANILLO_NORMALIZADO != ''
                        LEFT JOIN reporte_anillos_raw ra ON c.ANILLO_NORMALIZADO = ra.ANILLO_NORMALIZADO 
                            AND c.CABLE_NORMALIZADO = ra.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce_bus ON ra.TRONCAL = ce_bus.FIBRA
                        LEFT JOIN centrales_anillo_raw ce_anillo_central ON ra.TRONCAL = ce_anillo_central.FIBRA
                    """)
                elif has_bus:
                    # Solo BUS disponible
                    self.con.execute("""
                        CREATE OR REPLACE TABLE resultado_cruce AS
                        SELECT 
                            c.NOMBRE_CLIENTE,
                            c.ANILLO_VISIBLE,
                            c.IDServicio,
                            c.Responsable,
                            c.Estatus,
                            c.Observaciones,
                            c.FECHA,
                            c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
                            f.SHEATH_ID_UNICO AS SHEATHS_IDS,
                            f.FUNDA,
                            f.HILOS_OCUPADOS_INDIVIDUALES,
                            CASE 
                                WHEN p.ID_PRE IS NOT NULL AND (ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL) 
                                    THEN 'PREPROCESAMIENTO Y CENTRALES'
                                WHEN p.ID_PRE IS NOT NULL THEN 'PREPROCESAMIENTO'
                                WHEN ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL THEN 'CENTRALES'
                                WHEN ra.ID_ANILLO IS NOT NULL AND ce_bus.ID_CENTRAL IS NOT NULL THEN 'BUS'
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'REPORTE_ANILLOS'
                                ELSE 'SOLO PRINCIPAL'
                            END AS ORIGEN,
                            CASE 
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'SI' 
                                ELSE 'NO' 
                            END AS EN_REPORTE_ANILLOS,
                            ra.HILO_TRONCAL AS HILO_TRONCAL_ANILLOS,
                            ra.ODF_TRONCAL AS ODF_TRONCAL_ANILLOS,
                            ra.ANILLO_NORMALIZADO AS ANILLO_NORMALIZADO,
                            ra.CABLE_ORIGINAL_ANILLOS AS CABLE_ORIGINAL_ANILLOS,
                            ra.ID_SERVICIO AS ID_SERVICIO_ANILLOS,
                            ra.CLIENTE AS CLIENTE_ANILLOS,
                            ra.TRONCAL AS TRONCAL_ANILLOS,
                            CASE WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO' ELSE 'OK' END AS ERROR
                        FROM clientes_raw c
                        LEFT JOIN fibra_raw f ON c.CABLE_NORMALIZADO = f.NOMBRE_NORMALIZADO
                        LEFT JOIN preprocesamiento_raw p ON c.CABLE_NORMALIZADO = p.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce ON c.CABLE_NORMALIZADO = ce.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce_anillo ON c.ANILLO_NORMALIZADO = ce_anillo.ANILLO_NORMALIZADO 
                            AND c.ANILLO_NORMALIZADO != '' 
                            AND ce_anillo.ANILLO_NORMALIZADO != ''
                        LEFT JOIN reporte_anillos_raw ra ON c.ANILLO_NORMALIZADO = ra.ANILLO_NORMALIZADO 
                            AND c.CABLE_NORMALIZADO = ra.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce_bus ON ra.TRONCAL = ce_bus.FIBRA
                    """)
                else:
                    # Sin BUS ni ANILLO
                    self.con.execute("""
                        CREATE OR REPLACE TABLE resultado_cruce AS
                        SELECT 
                            c.NOMBRE_CLIENTE,
                            c.ANILLO_VISIBLE,
                            c.IDServicio,
                            c.Responsable,
                            c.Estatus,
                            c.Observaciones,
                            c.FECHA,
                            c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
                            f.SHEATH_ID_UNICO AS SHEATHS_IDS,
                            f.FUNDA,
                            f.HILOS_OCUPADOS_INDIVIDUALES,
                            CASE 
                                WHEN p.ID_PRE IS NOT NULL AND (ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL) 
                                    THEN 'PREPROCESAMIENTO Y CENTRALES'
                                WHEN p.ID_PRE IS NOT NULL THEN 'PREPROCESAMIENTO'
                                WHEN ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL THEN 'CENTRALES'
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'REPORTE_ANILLOS'
                                ELSE 'SOLO PRINCIPAL'
                            END AS ORIGEN,
                            CASE 
                                WHEN ra.ID_ANILLO IS NOT NULL THEN 'SI' 
                                ELSE 'NO' 
                            END AS EN_REPORTE_ANILLOS,
                            ra.HILO_TRONCAL AS HILO_TRONCAL_ANILLOS,
                            ra.ODF_TRONCAL AS ODF_TRONCAL_ANILLOS,
                            ra.ANILLO_NORMALIZADO AS ANILLO_NORMALIZADO,
                            ra.CABLE_ORIGINAL_ANILLOS AS CABLE_ORIGINAL_ANILLOS,
                            ra.ID_SERVICIO AS ID_SERVICIO_ANILLOS,
                            ra.CLIENTE AS CLIENTE_ANILLOS,
                            ra.TRONCAL AS TRONCAL_ANILLOS,
                            CASE WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO' ELSE 'OK' END AS ERROR
                        FROM clientes_raw c
                        LEFT JOIN fibra_raw f ON c.CABLE_NORMALIZADO = f.NOMBRE_NORMALIZADO
                        LEFT JOIN preprocesamiento_raw p ON c.CABLE_NORMALIZADO = p.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce ON c.CABLE_NORMALIZADO = ce.CABLE_NORMALIZADO
                        LEFT JOIN centrales_bus_raw ce_anillo ON c.ANILLO_NORMALIZADO = ce_anillo.ANILLO_NORMALIZADO 
                            AND c.ANILLO_NORMALIZADO != '' 
                            AND ce_anillo.ANILLO_NORMALIZADO != ''
                        LEFT JOIN reporte_anillos_raw ra ON c.ANILLO_NORMALIZADO = ra.ANILLO_NORMALIZADO 
                            AND c.CABLE_NORMALIZADO = ra.CABLE_NORMALIZADO
                    """)
            else:
                # Sin reporte anillos
                self.con.execute("""
                    CREATE OR REPLACE TABLE resultado_cruce AS
                    SELECT 
                        c.NOMBRE_CLIENTE,
                        c.ANILLO_VISIBLE,
                        c.IDServicio,
                        c.Responsable,
                        c.Estatus,
                        c.Observaciones,
                        c.FECHA,
                        c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
                        f.SHEATH_ID_UNICO AS SHEATHS_IDS,
                        f.FUNDA,
                        f.HILOS_OCUPADOS_INDIVIDUALES,
                        CASE 
                            WHEN p.ID_PRE IS NOT NULL AND (ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL) 
                                THEN 'PREPROCESAMIENTO Y CENTRALES'
                            WHEN p.ID_PRE IS NOT NULL THEN 'PREPROCESAMIENTO'
                            WHEN ce.ID_CENTRAL IS NOT NULL OR ce_anillo.ID_CENTRAL IS NOT NULL THEN 'CENTRALES'
                            ELSE 'SOLO PRINCIPAL'
                        END AS ORIGEN,
                        'NO' AS EN_REPORTE_ANILLOS,
                        '' AS HILO_TRONCAL_ANILLOS,
                        '' AS ODF_TRONCAL_ANILLOS,
                        '' AS ANILLO_NORMALIZADO,
                        '' AS CABLE_ORIGINAL_ANILLOS,
                        '' AS ID_SERVICIO_ANILLOS,
                        '' AS CLIENTE_ANILLOS,
                        '' AS TRONCAL_ANILLOS,
                        CASE WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO' ELSE 'OK' END AS ERROR
                    FROM clientes_raw c
                    LEFT JOIN fibra_raw f ON c.CABLE_NORMALIZADO = f.NOMBRE_NORMALIZADO
                    LEFT JOIN preprocesamiento_raw p ON c.CABLE_NORMALIZADO = p.CABLE_NORMALIZADO
                    LEFT JOIN centrales_bus_raw ce ON c.CABLE_NORMALIZADO = ce.CABLE_NORMALIZADO
                    LEFT JOIN centrales_bus_raw ce_anillo ON c.ANILLO_NORMALIZADO = ce_anillo.ANILLO_NORMALIZADO 
                        AND c.ANILLO_NORMALIZADO != '' 
                        AND ce_anillo.ANILLO_NORMALIZADO != ''
                """)
            
            count = self.con.execute("SELECT COUNT(*) FROM resultado_cruce").fetchone()[0]
            logger.info(f"Cruce: {count:,} registros")
            return count
            
        except Exception as e:
            logger.error(f"Error cruce: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _enriquecer_con_puertos(self):
        try:
            logger.info("Enriqueciendo con puertos...")
            
            try:
                count_ports = self.con.execute("SELECT COUNT(*) FROM ports_mapping").fetchone()[0]
                logger.info(f"  ports_mapping tiene {count_ports} registros")
            except Exception as e:
                logger.warning(f"  ports_mapping no existe: {e}")
                count_ports = 0
            
            if count_ports == 0:
                logger.warning("  No hay datos de ports para enriquecer")
                self.con.execute("""
                    CREATE OR REPLACE TABLE resultado_final AS
                    SELECT 
                        SHEATHS_IDS, FUNDA, CABLE_ACCESO_CLIENTE, HILOS_OCUPADOS_INDIVIDUALES,
                        NOMBRE_CLIENTE, ANILLO_VISIBLE, IDServicio, '' AS PUERTOS_IDS,
                        '' AS DESCRIPTION, '' AS DESCRIPTION_SHELF,
                        Responsable, Estatus, Observaciones, FECHA, ORIGEN,
                        EN_REPORTE_ANILLOS, HILO_TRONCAL_ANILLOS, ODF_TRONCAL_ANILLOS,
                        ANILLO_NORMALIZADO, CABLE_ORIGINAL_ANILLOS,
                        ID_SERVICIO_ANILLOS, CLIENTE_ANILLOS, TRONCAL_ANILLOS, ERROR
                    FROM resultado_cruce
                """)
                return
            
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_con_puertos AS
                SELECT 
                    r.*,
                    COALESCE(pm.puertos_ids, []) AS PUERTOS_IDS_LIST,
                    COALESCE(pm.DESCRIPTION, '') AS DESCRIPTION,
                    COALESCE(pm.DESCRIPTION_SHELF, '') AS DESCRIPTION_SHELF,
                    CASE 
                        WHEN r.ERROR != 'OK' THEN r.ERROR
                        WHEN pm.puertos_ids IS NULL THEN 'SYNC PORT | SIN PUERTO ASIGNADO'
                        ELSE 'OK'
                    END AS ERROR_ACTUALIZADO
                FROM resultado_cruce r
                LEFT JOIN ports_mapping pm
                    ON REGEXP_REPLACE(REGEXP_REPLACE(REGEXP_REPLACE(UPPER(TRIM(COALESCE(r.IDServicio, ''))), '-', '', 'g'), '_', '', 'g'), ' ', '', 'g')
                    = pm.servicio_normalizado
            """)
            
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_final AS
                SELECT 
                    SHEATHS_IDS, FUNDA, CABLE_ACCESO_CLIENTE, HILOS_OCUPADOS_INDIVIDUALES,
                    NOMBRE_CLIENTE, ANILLO_VISIBLE, IDServicio,
                    LIST_AGGREGATE(PUERTOS_IDS_LIST, 'string_agg', ', ') AS PUERTOS_IDS,
                    DESCRIPTION, DESCRIPTION_SHELF,
                    Responsable, Estatus, Observaciones, FECHA, ORIGEN,
                    EN_REPORTE_ANILLOS, HILO_TRONCAL_ANILLOS, ODF_TRONCAL_ANILLOS,
                    ANILLO_NORMALIZADO, CABLE_ORIGINAL_ANILLOS,
                    ID_SERVICIO_ANILLOS, CLIENTE_ANILLOS, TRONCAL_ANILLOS,
                    ERROR_ACTUALIZADO AS ERROR
                FROM resultado_con_puertos
            """)
            
            count = self.con.execute("SELECT COUNT(*) FROM resultado_final").fetchone()[0]
            logger.info(f"Enriquecido: {count:,} registros")
            
        except Exception as e:
            logger.error(f"Error enriquecimiento: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _agrupar_resultados(self):
        try:
            logger.info("Agrupando resultados...")
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_agrupado AS
                WITH agrupados AS (
                    SELECT 
                        NOMBRE_CLIENTE,
                        ANILLO_VISIBLE,
                        IDServicio,
                        Responsable,
                        Estatus,
                        Observaciones,
                        FECHA,
                        CABLE_ACCESO_CLIENTE,
                        ORIGEN,
                        EN_REPORTE_ANILLOS,
                        HILO_TRONCAL_ANILLOS,
                        ODF_TRONCAL_ANILLOS,
                        ANILLO_NORMALIZADO,
                        CABLE_ORIGINAL_ANILLOS,
                        ID_SERVICIO_ANILLOS,
                        CLIENTE_ANILLOS,
                        TRONCAL_ANILLOS,
                        LIST_DISTINCT(LIST_FILTER(LIST(SHEATHS_IDS), x -> x IS NOT NULL AND x != '' AND x != 'nan' AND x != 'None')) AS sheath_ids_list,
                        LIST_DISTINCT(LIST_FILTER(LIST(HILOS_OCUPADOS_INDIVIDUALES), x -> x IS NOT NULL AND x != '' AND x != 'nan' AND x != 'None')) AS hilos_list,
                        LIST_DISTINCT(FLATTEN(LIST_FILTER(LIST(CASE WHEN PUERTOS_IDS IS NOT NULL AND PUERTOS_IDS != '' AND PUERTOS_IDS != 'nan' AND PUERTOS_IDS != 'None' THEN STRING_SPLIT(PUERTOS_IDS, ', ') ELSE [] END), x -> LENGTH(x) > 0))) AS puertos_list,
                        FIRST(FUNDA) AS FUNDA,
                        FIRST(DESCRIPTION) AS DESCRIPTION,
                        FIRST(DESCRIPTION_SHELF) AS DESCRIPTION_SHELF,
                        LIST(ERROR) AS errores
                    FROM resultado_final
                    GROUP BY NOMBRE_CLIENTE, ANILLO_VISIBLE, IDServicio, Responsable, Estatus, Observaciones, FECHA, CABLE_ACCESO_CLIENTE, ORIGEN, EN_REPORTE_ANILLOS, HILO_TRONCAL_ANILLOS, ODF_TRONCAL_ANILLOS, ANILLO_NORMALIZADO, CABLE_ORIGINAL_ANILLOS, ID_SERVICIO_ANILLOS, CLIENTE_ANILLOS, TRONCAL_ANILLOS
                )
                SELECT 
                    NOMBRE_CLIENTE,
                    ANILLO_VISIBLE,
                    IDServicio,
                    Responsable,
                    Estatus,
                    Observaciones,
                    FECHA,
                    CABLE_ACCESO_CLIENTE,
                    ORIGEN,
                    EN_REPORTE_ANILLOS,
                    HILO_TRONCAL_ANILLOS,
                    ODF_TRONCAL_ANILLOS,
                    ANILLO_NORMALIZADO,
                    CABLE_ORIGINAL_ANILLOS,
                    ID_SERVICIO_ANILLOS,
                    CLIENTE_ANILLOS,
                    TRONCAL_ANILLOS,
                    LIST_AGGREGATE(sheath_ids_list, 'string_agg', ', ') AS SHEATHS_IDS,
                    FUNDA,
                    LIST_AGGREGATE(hilos_list, 'string_agg', ', ') AS HILOS_OCUPADOS_INDIVIDUALES,
                    LIST_AGGREGATE(puertos_list, 'string_agg', ', ') AS PUERTOS_IDS,
                    DESCRIPTION,
                    DESCRIPTION_SHELF,
                    CASE 
                        WHEN LIST_CONTAINS(errores, 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO') 
                            THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                        WHEN LENGTH(hilos_list) = 0 OR LIST_AGGREGATE(hilos_list, 'string_agg', ', ') = '' 
                            THEN 'SYNC FIBER | HILO LIBRE'
                        WHEN LIST_CONTAINS(errores, 'SYNC PORT | SIN PUERTO ASIGNADO') 
                            OR LENGTH(puertos_list) = 0 
                            THEN 'SYNC PORT | SIN PUERTO ASIGNADO'
                        ELSE 'OK'
                    END AS ERROR
                FROM agrupados
            """)
            
            self.df_resultado = self.con.execute("SELECT * FROM resultado_agrupado").fetchdf()
            logger.info(f"Agrupado: {len(self.df_resultado):,} registros")
            
            # Crear DataFrame para la hoja de anillos
            if 'EN_REPORTE_ANILLOS' in self.df_resultado.columns:
                df_anillos = self.df_resultado[self.df_resultado['EN_REPORTE_ANILLOS'] == 'SI'].copy()
                if len(df_anillos) > 0:
                    self.df_anillos = df_anillos
                    logger.info(f"Registros de REPORTE_ANILLOS: {len(df_anillos):,}")
                    bus_count = self.df_resultado[self.df_resultado['ORIGEN'] == 'BUS'].shape[0]
                    anillo_count = self.df_resultado[self.df_resultado['ORIGEN'] == 'ANILLO'].shape[0]
                    logger.info(f"Registros marcados como BUS: {bus_count:,}")
                    logger.info(f"Registros marcados como ANILLO: {anillo_count:,}")
            
        except Exception as e:
            logger.error(f"Error agrupacion: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _guardar_resultados(self):
        try:
            if self.df_resultado is None or len(self.df_resultado) == 0:
                logger.warning("No hay datos para guardar")
                return
            
            os.makedirs(os.path.dirname(RUTA_SALIDA), exist_ok=True)
            
            columnas = [
                'SHEATHS_IDS', 'FUNDA', 'CABLE_ACCESO_CLIENTE', 'HILOS_OCUPADOS_INDIVIDUALES',
                'NOMBRE_CLIENTE', 'ANILLO_VISIBLE', 'IDServicio', 'PUERTOS_IDS',
                'DESCRIPTION', 'DESCRIPTION_SHELF',
                'Responsable', 'Estatus', 'Observaciones', 'FECHA',
                'ORIGEN', 'EN_REPORTE_ANILLOS',
                'HILO_TRONCAL_ANILLOS', 'ODF_TRONCAL_ANILLOS',
                'ANILLO_NORMALIZADO', 'CABLE_ORIGINAL_ANILLOS',
                'ID_SERVICIO_ANILLOS', 'CLIENTE_ANILLOS', 'TRONCAL_ANILLOS',
                'ERROR'
            ]
            
            existentes = [col for col in columnas if col in self.df_resultado.columns]
            faltantes = [col for col in self.df_resultado.columns if col not in columnas]
            self.df_resultado = self.df_resultado[existentes + faltantes]
            
            # Cargar libro existente si existe
            if os.path.exists(RUTA_SALIDA):
                logger.info("Cargando libro existente para preservar hojas...")
                with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                    self.df_resultado.to_excel(writer, sheet_name='Cruce', index=False)
                    
                    # Guardar hoja de anillos si existe
                    if self.df_anillos is not None and len(self.df_anillos) > 0:
                        columnas_anillos = [
                            'ANILLO_NORMALIZADO', 'CABLE_ORIGINAL_ANILLOS',
                            'HILO_TRONCAL_ANILLOS', 'ODF_TRONCAL_ANILLOS',
                            'ID_SERVICIO_ANILLOS', 'CLIENTE_ANILLOS', 'TRONCAL_ANILLOS',
                            'NOMBRE_CLIENTE', 'ANILLO_VISIBLE', 'IDServicio', 'ORIGEN'
                        ]
                        columnas_existentes_anillos = [col for col in columnas_anillos if col in self.df_anillos.columns]
                        df_anillos_export = self.df_anillos[columnas_existentes_anillos].copy()
                        df_anillos_export.to_excel(writer, sheet_name='Reporte_Anillos', index=False)
                        logger.info(f"Hoja 'Reporte_Anillos' guardada con {len(df_anillos_export):,} registros")
            else:
                logger.info("Creando nuevo archivo...")
                with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl') as writer:
                    self.df_resultado.to_excel(writer, sheet_name='Cruce', index=False)
                    
                    if self.df_anillos is not None and len(self.df_anillos) > 0:
                        columnas_anillos = [
                            'ANILLO_NORMALIZADO', 'CABLE_ORIGINAL_ANILLOS',
                            'HILO_TRONCAL_ANILLOS', 'ODF_TRONCAL_ANILLOS',
                            'ID_SERVICIO_ANILLOS', 'CLIENTE_ANILLOS', 'TRONCAL_ANILLOS',
                            'NOMBRE_CLIENTE', 'ANILLO_VISIBLE', 'IDServicio', 'ORIGEN'
                        ]
                        columnas_existentes_anillos = [col for col in columnas_anillos if col in self.df_anillos.columns]
                        df_anillos_export = self.df_anillos[columnas_existentes_anillos].copy()
                        df_anillos_export.to_excel(writer, sheet_name='Reporte_Anillos', index=False)
                        logger.info(f"Hoja 'Reporte_Anillos' guardada con {len(df_anillos_export):,} registros")
            
            logger.info(f"Archivo guardado: {RUTA_SALIDA}")
            
        except Exception as e:
            logger.error(f"Error guardando: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _limpiar_recursos(self):
        if self.con:
            self.con.close()
    
    def ejecutar(self):
        try:
            inicio_total = datetime.now()
            logger.info("=" * 60)
            logger.info("INICIANDO SINCRONIZACION")
            logger.info("=" * 60)
            
            self._conectar_duckdb()
            
            self._cargar_reporte_fibra()
            self._cargar_datos_clientes()
            self._cargar_preprocesamiento()
            self._cargar_centrales()
            self._cargar_reporte_anillos()
            self._cargar_datos_ports()
            
            self._realizar_cruce()
            self._enriquecer_con_puertos()
            self._agrupar_resultados()
            
            if self.df_resultado is not None and len(self.df_resultado) > 0:
                total = len(self.df_resultado)
                ok = self.df_resultado[self.df_resultado['ERROR'] == 'OK'].shape[0]
                error_fiber = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC FIBER', na=False)].shape[0]
                error_port = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC PORT', na=False)].shape[0]
                
                print("\n" + "=" * 60)
                print("RESUMEN FINAL")
                print("=" * 60)
                print(f"Total registros: {total:,}")
                print(f"OK: {ok:,} ({ok/total*100:.1f}%)")
                print(f"SYNC FIBER: {error_fiber:,} ({error_fiber/total*100:.1f}%)")
                print(f"SYNC PORT: {error_port:,} ({error_port/total*100:.1f}%)")
                
                if 'EN_REPORTE_ANILLOS' in self.df_resultado.columns:
                    en_anillos = self.df_resultado[self.df_resultado['EN_REPORTE_ANILLOS'] == 'SI'].shape[0]
                    print("-" * 60)
                    print(f"EN REPORTE ANILLOS: {en_anillos:,} ({en_anillos/total*100:.1f}%)")
                
                if 'ORIGEN' in self.df_resultado.columns:
                    for origen in self.df_resultado['ORIGEN'].unique():
                        count = self.df_resultado[self.df_resultado['ORIGEN'] == origen].shape[0]
                        print(f"  {origen}: {count:,} ({count/total*100:.1f}%)")
                
                print("=" * 60)
            else:
                print("\nNo se generaron resultados")
            
            print(f"Tiempo total: {(datetime.now() - inicio_total).total_seconds():.1f}s")
            
            self._guardar_resultados()
            return True
            
        except Exception as e:
            print(f"\nERROR: {e}")
            print(traceback.format_exc())
            return False
        finally:
            self._limpiar_recursos()


if __name__ == "__main__":
    try:
        sincronizador = SincronizadorDatos(FECHA_INICIO, FECHA_FIN)
        exito = sincronizador.ejecutar()
        sys.exit(0 if exito else 1)
    except KeyboardInterrupt:
        print("\nProceso interrumpido")
        sys.exit(1)
    except Exception as e:
        print(f"Error: {e}")
        print(traceback.format_exc())
        sys.exit(1)