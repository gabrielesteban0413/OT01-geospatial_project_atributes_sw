"""
SISTEMA DE SINCRONIZACION DE DATOS
Version: 17.2 - AGREGADAS COLUMNAS DE SPLICE
"""

import duckdb
import pandas as pd
import os
import sys
from datetime import datetime
import logging
import re
import traceback

logging.basicConfig(
    level=logging.INFO,
    format='%(message)s',
    handlers=[
        logging.FileHandler('sincronizacion.log', encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)
logger = logging.getLogger(__name__)

FECHA_INICIO = "01/07/2026"
FECHA_FIN = "24/08/2026"

RUTA_REPORTE_FIBRA = r"C:\A_GS1_PROYECTOS\reporte_fibra.xlsx"
RUTA_CLIENTES = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-client_report_filtrado.xlsx"
RUTA_PORTS = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\01-Bulk export of ports.xlsx"
RUTA_SALIDA = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\04-Bulk Merge_Clientes.xlsx"

MEMORY_LIMIT = '16GB'
THREADS = 8


class SincronizadorDatos:
    def __init__(self, fecha_inicio, fecha_fin):
        self.fecha_inicio = self._parse_fecha(fecha_inicio)
        self.fecha_fin = self._parse_fecha(fecha_fin)
        self.con = None
        self.df_resultado = None
    
    def _parse_fecha(self, fecha_str):
        try:
            return datetime.strptime(fecha_str, "%d/%m/%Y")
        except ValueError:
            try:
                return datetime.strptime(fecha_str, "%Y-%m-%d")
            except:
                raise ValueError(f"Formato invalido: {fecha_str}")
    
    def _conectar_duckdb(self):
        logger.info("Conectando a DuckDB...")
        self.con = duckdb.connect(":memory:")
        self.con.execute(f"SET memory_limit='{MEMORY_LIMIT}';")
        self.con.execute(f"SET threads={THREADS};")
        logger.info("Conexion DuckDB establecida")
    
    def _cargar_reporte_fibra(self):
        if not os.path.exists(RUTA_REPORTE_FIBRA):
            raise FileNotFoundError(f"Fibra no encontrado: {RUTA_REPORTE_FIBRA}")
        
        logger.info(f"Cargando reporte de fibra...")
        df_fibra = pd.read_excel(RUTA_REPORTE_FIBRA, sheet_name='Resumen', dtype=str)
        logger.info(f"Archivo leido: {len(df_fibra):,} filas")
        
        self.con.register('fibra_temp', df_fibra)
        
        self.con.execute("""
            CREATE OR REPLACE TABLE fibra_raw AS
            WITH expanded AS (
                SELECT 
                    UNNEST(STRING_SPLIT(REPLACE(SHEATHS_IDS, ';', ','), ',')) AS SHEATH_ID_UNICO,
                    UNNEST(STRING_SPLIT(REPLACE(NOMBRES_ANTIGUOS, ';', ','), ',')) AS NOMBRE_ANTIGUO,
                    FUNDA,
                    HILOS_OCUPADOS_INDIVIDUALES,
                    SPLICE_ENTRADA_IDS,
                    SPLICE_ENTRADA_NOMBRES,
                    SPLICE_SALIDA_IDS,
                    SPLICE_SALIDA_NOMBRES
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
                SPLICE_ENTRADA_IDS,
                SPLICE_ENTRADA_NOMBRES,
                SPLICE_SALIDA_IDS,
                SPLICE_SALIDA_NOMBRES,
                UPPER(REPLACE(REPLACE(REPLACE(TRIM(NOMBRE_ANTIGUO), '-', ''), '_', ''), ' ', '')) AS NOMBRE_NORMALIZADO
            FROM expanded
            WHERE TRIM(SHEATH_ID_UNICO) != '' 
              AND TRIM(SHEATH_ID_UNICO) NOT IN ('nan', 'None')
              AND TRIM(NOMBRE_ANTIGUO) != ''
              AND TRIM(NOMBRE_ANTIGUO) NOT IN ('nan', 'None')
        """)
        
        count = self.con.execute("SELECT COUNT(*) FROM fibra_raw").fetchone()[0]
        logger.info(f"Fibra cargada: {count:,} registros")
        
        del df_fibra
        return count
    
    def _cargar_datos_clientes(self):
        if not os.path.exists(RUTA_CLIENTES):
            raise FileNotFoundError(f"Clientes no encontrado: {RUTA_CLIENTES}")
        
        logger.info(f"Cargando datos de clientes...")
        df_clientes = pd.read_excel(RUTA_CLIENTES, dtype=str)
        logger.info(f"Archivo leido: {len(df_clientes):,} filas")
        
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
        
        logger.info(f"Columna de cable: {col_cable}")
        
        fecha_inicio_str = self.fecha_inicio.strftime('%Y-%m-%d')
        fecha_fin_str = self.fecha_fin.strftime('%Y-%m-%d')
        
        if col_fecha:
            query = f"""
            CREATE OR REPLACE TABLE clientes_raw AS
            WITH filtered AS (
                SELECT *
                FROM clientes_temp
                WHERE TRY_CAST(\"{col_fecha}\" AS DATE) IS NULL 
                   OR (TRY_CAST(\"{col_fecha}\" AS DATE) BETWEEN '{fecha_inicio_str}' AND '{fecha_fin_str}')
            ),
            expanded AS (
                SELECT 
                    *,
                    UNNEST(STRING_SPLIT(REPLACE(\"{col_cable}\", ';', ','), ',')) AS CABLE_UNICO,
                    ROW_NUMBER() OVER () AS ID_ORIGINAL
                FROM filtered
                WHERE \"{col_cable}\" IS NOT NULL 
                  AND \"{col_cable}\" != 'nan'
                  AND \"{col_cable}\" != 'None'
                  AND \"{col_cable}\" != ''
            )
            SELECT 
                ID_ORIGINAL,
                TRIM(CABLE_UNICO) AS Cable_de_Acceso,
                Nombre AS NOMBRE_CLIENTE,
                Anillo,
                IDServicio,
                Responsable,
                Estatus,
                Observaciones,
                \"{col_fecha}\" AS FECHA,
                UPPER(REPLACE(REPLACE(REPLACE(TRIM(CABLE_UNICO), '-', ''), '_', ''), ' ', '')) AS CABLE_NORMALIZADO
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
                    UNNEST(STRING_SPLIT(REPLACE(\"{col_cable}\", ';', ','), ',')) AS CABLE_UNICO,
                    ROW_NUMBER() OVER () AS ID_ORIGINAL
                FROM clientes_temp
                WHERE \"{col_cable}\" IS NOT NULL 
                  AND \"{col_cable}\" != 'nan'
                  AND \"{col_cable}\" != 'None'
                  AND \"{col_cable}\" != ''
            )
            SELECT 
                ID_ORIGINAL,
                TRIM(CABLE_UNICO) AS Cable_de_Acceso,
                Nombre AS NOMBRE_CLIENTE,
                Anillo,
                IDServicio,
                Responsable,
                Estatus,
                Observaciones,
                '' AS FECHA,
                UPPER(REPLACE(REPLACE(REPLACE(TRIM(CABLE_UNICO), '-', ''), '_', ''), ' ', '')) AS CABLE_NORMALIZADO
            FROM expanded
            WHERE TRIM(CABLE_UNICO) != '' 
              AND TRIM(CABLE_UNICO) NOT IN ('nan', 'None')
            """
        
        self.con.execute(query)
        
        count = self.con.execute("SELECT COUNT(*) FROM clientes_raw").fetchone()[0]
        logger.info(f"Clientes cargados: {count:,} registros")
        
        del df_clientes
        return count
    
    def _cargar_datos_ports(self):
        if not os.path.exists(RUTA_PORTS):
            logger.warning("Archivo de ports no encontrado")
            return 0
        
        logger.info(f"Cargando archivo de ports...")
        
        try:
            df_sample = pd.read_excel(RUTA_PORTS, sheet_name='Port', nrows=5, dtype=str)
            df_sample.columns = df_sample.columns.str.strip()
            
            col_id = None
            col_servicio = None
            col_comentarios = None
            
            for col in df_sample.columns:
                col_lower = col.strip().lower()
                if col_lower == 'id':
                    col_id = col
                if 'id servicio' in col_lower or 'idservicio' in col_lower or 'service id' in col_lower:
                    col_servicio = col
                if 'comentarios' in col_lower or 'comentario' in col_lower or 'comments' in col_lower:
                    col_comentarios = col
            
            if not col_servicio:
                for col in df_sample.columns:
                    if 'servicio' in col.strip().lower() or 'service' in col.strip().lower():
                        col_servicio = col
                        break
            
            if not col_comentarios:
                for col in df_sample.columns:
                    if 'coment' in col.strip().lower() or 'obs' in col.strip().lower():
                        col_comentarios = col
                        break
            
            logger.info(f"Columnas: ID='{col_id}', Servicio='{col_servicio}', Comentarios='{col_comentarios}'")
            
            if not col_id:
                logger.warning("No se encontro columna ID")
                return 0
            
            columnas_a_leer = [col_id]
            if col_servicio:
                columnas_a_leer.append(col_servicio)
            if col_comentarios:
                columnas_a_leer.append(col_comentarios)
            
            logger.info(f"Cargando {len(columnas_a_leer)} columnas...")
            df_ports = pd.read_excel(
                RUTA_PORTS, 
                sheet_name='Port', 
                dtype=str,
                usecols=columnas_a_leer
            )
            
            logger.info(f"Ports cargados: {len(df_ports):,} filas")
            
            self.con.register('ports_temp', df_ports)
            
            if col_servicio and col_comentarios:
                query = f"""
                CREATE OR REPLACE TABLE ports_mapping AS
                SELECT 
                    UPPER(REPLACE(REPLACE(REPLACE(TRIM(COALESCE(ids.servicio_id, '')), '-', ''), '_', ''), ' ', '')) AS servicio_normalizado,
                    LIST(port_id) AS puertos_ids
                FROM (
                    SELECT 
                        CAST(\"{col_id}\" AS VARCHAR) AS port_id,
                        COALESCE(
                            NULLIF(TRIM(CAST(\"{col_servicio}\" AS VARCHAR)), ''),
                            REGEXP_EXTRACT(CAST(\"{col_comentarios}\" AS VARCHAR), '(?:ID|Id|id|Servicio|servicio|SERVICIO)[\\s:]*([A-Za-z0-9\\-]+)', 1),
                            REGEXP_EXTRACT(CAST(\"{col_comentarios}\" AS VARCHAR), '([A-Z0-9]{{5,}})', 1)
                        ) AS servicio_id
                    FROM ports_temp
                    WHERE CAST(\"{col_id}\" AS VARCHAR) IS NOT NULL
                      AND CAST(\"{col_id}\" AS VARCHAR) != 'nan'
                      AND CAST(\"{col_id}\" AS VARCHAR) != 'None'
                      AND CAST(\"{col_id}\" AS VARCHAR) != ''
                ) ids
                WHERE servicio_id IS NOT NULL
                  AND servicio_id != ''
                  AND servicio_id != 'nan'
                  AND servicio_id != 'None'
                GROUP BY servicio_normalizado
                """
            elif col_servicio:
                query = f"""
                CREATE OR REPLACE TABLE ports_mapping AS
                SELECT 
                    UPPER(REPLACE(REPLACE(REPLACE(TRIM(CAST(\"{col_servicio}\" AS VARCHAR)), '-', ''), '_', ''), ' ', '')) AS servicio_normalizado,
                    LIST(CAST(\"{col_id}\" AS VARCHAR)) AS puertos_ids
                FROM ports_temp
                WHERE CAST(\"{col_servicio}\" AS VARCHAR) IS NOT NULL
                  AND CAST(\"{col_servicio}\" AS VARCHAR) != 'nan'
                  AND CAST(\"{col_servicio}\" AS VARCHAR) != 'None'
                  AND CAST(\"{col_servicio}\" AS VARCHAR) != ''
                  AND CAST(\"{col_id}\" AS VARCHAR) IS NOT NULL
                  AND CAST(\"{col_id}\" AS VARCHAR) != 'nan'
                  AND CAST(\"{col_id}\" AS VARCHAR) != 'None'
                  AND CAST(\"{col_id}\" AS VARCHAR) != ''
                GROUP BY servicio_normalizado
                """
            else:
                logger.warning("No se encontraron columnas para extraer IDs de servicio")
                return 0
            
            self.con.execute(query)
            
            count = self.con.execute("SELECT COUNT(*) FROM ports_mapping").fetchone()[0]
            logger.info(f"Servicios con puerto: {count:,}")
            
            del df_ports
            del df_sample
            
            return count
            
        except Exception as e:
            logger.error(f"Error cargando ports: {e}")
            logger.error(traceback.format_exc())
            return 0
    
    def _realizar_cruce(self):
        logger.info("Realizando cruce de datos...")
        
        try:
            count_fibra = self.con.execute("SELECT COUNT(*) FROM fibra_raw").fetchone()[0]
            count_clientes = self.con.execute("SELECT COUNT(*) FROM clientes_raw").fetchone()[0]
            
            logger.info(f"Fibra: {count_fibra:,} registros, Clientes: {count_clientes:,} registros")
            
            if count_fibra == 0 or count_clientes == 0:
                logger.warning("No hay datos para cruzar")
                return 0
            
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_cruce AS
                SELECT 
                    c.ID_ORIGINAL,
                    c.NOMBRE_CLIENTE,
                    c.Anillo,
                    c.IDServicio,
                    c.Responsable,
                    c.Estatus,
                    c.Observaciones,
                    c.FECHA,
                    c.Cable_de_Acceso AS CABLE_ACCESO_CLIENTE,
                    f.SHEATH_ID_UNICO AS SHEATHS_IDS,
                    f.FUNDA,
                    f.HILOS_OCUPADOS_INDIVIDUALES,
                    f.SPLICE_ENTRADA_IDS,
                    f.SPLICE_ENTRADA_NOMBRES,
                    f.SPLICE_SALIDA_IDS,
                    f.SPLICE_SALIDA_NOMBRES,
                    CASE 
                        WHEN f.SHEATH_ID_UNICO IS NULL THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                        ELSE 'OK'
                    END AS ERROR
                FROM clientes_raw c
                LEFT JOIN fibra_raw f 
                    ON c.CABLE_NORMALIZADO = f.NOMBRE_NORMALIZADO
            """)
            
            count = self.con.execute("SELECT COUNT(*) FROM resultado_cruce").fetchone()[0]
            logger.info(f"Cruce completado: {count:,} registros")
            
            return count
            
        except Exception as e:
            logger.error(f"Error en cruce: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _enriquecer_con_puertos(self):
        logger.info("Enriqueciendo con puertos...")
        
        try:
            count_ports = self.con.execute("SELECT COUNT(*) FROM ports_mapping").fetchone()[0]
            if count_ports == 0:
                logger.warning("No hay datos de ports para enriquecer")
                self.con.execute("""
                    CREATE OR REPLACE TABLE resultado_final AS
                    SELECT 
                        ID_ORIGINAL,
                        SHEATHS_IDS,
                        FUNDA,
                        CABLE_ACCESO_CLIENTE,
                        HILOS_OCUPADOS_INDIVIDUALES,
                        NOMBRE_CLIENTE,
                        Anillo,
                        IDServicio,
                        '' AS PUERTOS_IDS,
                        SPLICE_ENTRADA_IDS,
                        SPLICE_ENTRADA_NOMBRES,
                        SPLICE_SALIDA_IDS,
                        SPLICE_SALIDA_NOMBRES,
                        Responsable,
                        Estatus,
                        Observaciones,
                        FECHA,
                        ERROR
                    FROM resultado_cruce
                """)
                return
            
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_con_puertos AS
                SELECT 
                    r.*,
                    COALESCE(pm.puertos_ids, []) AS PUERTOS_IDS_LIST,
                    CASE 
                        WHEN r.ERROR != 'OK' THEN r.ERROR
                        WHEN pm.puertos_ids IS NULL THEN 'SYNC PORT | SIN PUERTO ASIGNADO'
                        ELSE 'OK'
                    END AS ERROR_ACTUALIZADO
                FROM resultado_cruce r
                LEFT JOIN ports_mapping pm
                    ON UPPER(REPLACE(REPLACE(REPLACE(COALESCE(r.IDServicio, ''), '-', ''), '_', ''), ' ', '')) 
                    = pm.servicio_normalizado
            """)
            
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_final AS
                SELECT 
                    ID_ORIGINAL,
                    SHEATHS_IDS,
                    FUNDA,
                    CABLE_ACCESO_CLIENTE,
                    HILOS_OCUPADOS_INDIVIDUALES,
                    NOMBRE_CLIENTE,
                    Anillo,
                    IDServicio,
                    LIST_AGGREGATE(PUERTOS_IDS_LIST, 'string_agg', ', ') AS PUERTOS_IDS,
                    SPLICE_ENTRADA_IDS,
                    SPLICE_ENTRADA_NOMBRES,
                    SPLICE_SALIDA_IDS,
                    SPLICE_SALIDA_NOMBRES,
                    Responsable,
                    Estatus,
                    Observaciones,
                    FECHA,
                    ERROR_ACTUALIZADO AS ERROR
                FROM resultado_con_puertos
            """)
            
            count = self.con.execute("SELECT COUNT(*) FROM resultado_final").fetchone()[0]
            logger.info(f"Enriquecimiento completado: {count:,} registros")
            
        except Exception as e:
            logger.error(f"Error en enriquecimiento: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _agrupar_resultados(self):
        logger.info("Agrupando resultados...")
        
        try:
            self.con.execute("""
                CREATE OR REPLACE TABLE resultado_agrupado AS
                WITH agrupados AS (
                    SELECT 
                        ID_ORIGINAL,
                        LIST_DISTINCT(
                            LIST_FILTER(
                                LIST(SHEATHS_IDS), 
                                x -> x IS NOT NULL AND x != '' AND x != 'nan' AND x != 'None'
                            )
                        ) AS sheath_ids_list,
                        LIST_DISTINCT(
                            LIST_FILTER(
                                LIST(HILOS_OCUPADOS_INDIVIDUALES), 
                                x -> x IS NOT NULL AND x != '' AND x != 'nan' AND x != 'None'
                            )
                        ) AS hilos_list,
                        LIST_DISTINCT(
                            FLATTEN(
                                LIST_FILTER(
                                    LIST(
                                        CASE 
                                            WHEN PUERTOS_IDS IS NOT NULL AND PUERTOS_IDS != '' AND PUERTOS_IDS != 'nan' AND PUERTOS_IDS != 'None'
                                            THEN STRING_SPLIT(PUERTOS_IDS, ', ')
                                            ELSE []
                                        END
                                    ),
                                    x -> LENGTH(x) > 0
                                )
                            )
                        ) AS puertos_list,
                        FIRST(FUNDA) AS FUNDA,
                        FIRST(CABLE_ACCESO_CLIENTE) AS CABLE_ACCESO_CLIENTE,
                        FIRST(NOMBRE_CLIENTE) AS NOMBRE_CLIENTE,
                        FIRST(Anillo) AS Anillo,
                        FIRST(IDServicio) AS IDServicio,
                        FIRST(Responsable) AS Responsable,
                        FIRST(Estatus) AS Estatus,
                        FIRST(Observaciones) AS Observaciones,
                        FIRST(FECHA) AS FECHA,
                        FIRST(SPLICE_ENTRADA_IDS) AS SPLICE_ENTRADA_IDS,
                        FIRST(SPLICE_ENTRADA_NOMBRES) AS SPLICE_ENTRADA_NOMBRES,
                        FIRST(SPLICE_SALIDA_IDS) AS SPLICE_SALIDA_IDS,
                        FIRST(SPLICE_SALIDA_NOMBRES) AS SPLICE_SALIDA_NOMBRES,
                        LIST(ERROR) AS errores
                    FROM resultado_final
                    GROUP BY ID_ORIGINAL
                )
                SELECT 
                    ID_ORIGINAL,
                    LIST_AGGREGATE(sheath_ids_list, 'string_agg', ', ') AS SHEATHS_IDS,
                    FUNDA,
                    CABLE_ACCESO_CLIENTE,
                    LIST_AGGREGATE(hilos_list, 'string_agg', ', ') AS HILOS_OCUPADOS_INDIVIDUALES,
                    NOMBRE_CLIENTE,
                    Anillo,
                    IDServicio,
                    LIST_AGGREGATE(puertos_list, 'string_agg', ', ') AS PUERTOS_IDS,
                    SPLICE_ENTRADA_IDS,
                    SPLICE_ENTRADA_NOMBRES,
                    SPLICE_SALIDA_IDS,
                    SPLICE_SALIDA_NOMBRES,
                    Responsable,
                    Estatus,
                    Observaciones,
                    FECHA,
                    CASE 
                        WHEN LIST_CONTAINS(errores, 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO') 
                            THEN 'SYNC FIBER | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                        WHEN LENGTH(hilos_list) = 0 OR LIST_AGGREGATE(hilos_list, 'string_agg', ', ') = '' 
                            THEN 'SYNC FIBER | HILO LIBRE'
                        WHEN LIST_CONTAINS(errores, 'SYNC PORT | SMALLWORLD DIFERENTE A BASE CORPORATIVO') 
                            THEN 'SYNC PORT | SMALLWORLD DIFERENTE A BASE CORPORATIVO'
                        WHEN LIST_CONTAINS(errores, 'SYNC PORT | SIN PUERTO ASIGNADO') 
                            OR LENGTH(puertos_list) = 0 
                            THEN 'SYNC PORT | SIN PUERTO ASIGNADO'
                        ELSE 'OK'
                    END AS ERROR
                FROM agrupados
            """)
            
            self.df_resultado = self.con.execute("SELECT * FROM resultado_agrupado").fetchdf()
            logger.info(f"Agrupacion completada: {len(self.df_resultado):,} registros")
            
        except Exception as e:
            logger.error(f"Error en agrupacion: {e}")
            logger.error(traceback.format_exc())
            raise
    
    def _guardar_resultados(self):
        try:
            logger.info(f"Guardando resultados en {RUTA_SALIDA}...")
            
            if self.df_resultado is None or len(self.df_resultado) == 0:
                logger.warning("No hay datos para guardar")
                return
            
            os.makedirs(os.path.dirname(RUTA_SALIDA), exist_ok=True)
            
            columnas_orden = [
                'SHEATHS_IDS', 'FUNDA', 'CABLE_ACCESO_CLIENTE', 'HILOS_OCUPADOS_INDIVIDUALES',
                'NOMBRE_CLIENTE', 'Anillo', 'IDServicio', 'PUERTOS_IDS',
                'SPLICE_ENTRADA_IDS', 'SPLICE_ENTRADA_NOMBRES',
                'SPLICE_SALIDA_IDS', 'SPLICE_SALIDA_NOMBRES',
                'Responsable', 'Estatus', 'Observaciones', 'FECHA', 'ERROR'
            ]
            
            columnas_existentes = [col for col in columnas_orden if col in self.df_resultado.columns]
            columnas_faltantes = [col for col in self.df_resultado.columns if col not in columnas_orden]
            self.df_resultado = self.df_resultado[columnas_existentes + columnas_faltantes]
            
            with pd.ExcelWriter(RUTA_SALIDA, engine='openpyxl') as writer:
                self.df_resultado.to_excel(writer, sheet_name='Cruce', index=False)
            
            logger.info("Archivo guardado exitosamente")
            
        except Exception as e:
            logger.error(f"Error guardando: {e}")
            raise
    
    def _limpiar_recursos(self):
        if self.con:
            self.con.close()
            logger.info("Conexion DuckDB cerrada")
    
    def ejecutar(self):
        try:
            inicio_total = datetime.now()
            logger.info("=" * 60)
            logger.info("INICIANDO SINCRONIZACION DE DATOS")
            logger.info("=" * 60)
            
            self._conectar_duckdb()
            
            inicio = datetime.now()
            logger.info("\n[1/5] Cargando reporte de fibra")
            self._cargar_reporte_fibra()
            logger.info(f"Tiempo: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            inicio = datetime.now()
            logger.info("\n[2/5] Cargando datos de clientes")
            self._cargar_datos_clientes()
            logger.info(f"Tiempo: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            inicio = datetime.now()
            logger.info("\n[3/5] Cargando datos de ports")
            self._cargar_datos_ports()
            logger.info(f"Tiempo: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            inicio = datetime.now()
            logger.info("\n[4/5] Procesando cruce de datos")
            self._realizar_cruce()
            self._enriquecer_con_puertos()
            logger.info(f"Tiempo: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            inicio = datetime.now()
            logger.info("\n[5/5] Agrupando resultados")
            self._agrupar_resultados()
            logger.info(f"Tiempo: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            if self.df_resultado is not None and len(self.df_resultado) > 0:
                total = len(self.df_resultado)
                ok = self.df_resultado[self.df_resultado['ERROR'] == 'OK'].shape[0]
                error_fiber = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC FIBER', na=False)].shape[0]
                error_port = self.df_resultado[self.df_resultado['ERROR'].str.contains('SYNC PORT', na=False)].shape[0]
                
                print("\n" + "=" * 60)
                print("RESUMEN DE SINCRONIZACION")
                print("=" * 60)
                print(f"Total registros: {total:,}")
                print(f"OK: {ok:,} ({ok/total*100:.1f}%)")
                print(f"SYNC FIBER: {error_fiber:,} ({error_fiber/total*100:.1f}%)")
                print(f"SYNC PORT: {error_port:,} ({error_port/total*100:.1f}%)")
                print("=" * 60)
            else:
                print("\n" + "=" * 60)
                print("No se generaron resultados")
                print("=" * 60)
            
            print(f"Tiempo total: {(datetime.now() - inicio_total).total_seconds():.1f}s")
            
            inicio = datetime.now()
            self._guardar_resultados()
            logger.info(f"Tiempo guardado: {(datetime.now() - inicio).total_seconds():.1f}s")
            
            print(f"\nArchivo guardado: {RUTA_SALIDA}")
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