"""
Script para buscar clientes en el archivo Excel de ports.
Filtra clientes por ODF TRONCAL SALIDA vacio (NaN).
Formato de salida: ELEMENT | ID | VALUE | OWNER | ERROR | COMMENTS
"""

import pandas as pd
from pathlib import Path
import warnings
from datetime import datetime
from typing import Tuple, Dict
import logging

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

BASE_PATH = Path(r"C:\A_GS1_PROYECTOS")
RUTA_CLIENTES = BASE_PATH / "usuarios.xlsx"
RUTA_PORTS = BASE_PATH / r"0_Documents_gs\output\quality\01-Bulk export of ports.xlsx"
RUTA_SALIDA = BASE_PATH / "resultado_busqueda_clientes_normalizado.xlsx"

COLUMNAS_PORTS = {
    "cambio": "Cambio Service Manager",
    "id_servicio": "Id Servicio",
    "anillo": "Proyecto De Red"
}


def normalizar_texto(valor) -> str:
    if pd.isna(valor):
        return ""
    return str(valor).strip().lower()


def cargar_clientes() -> pd.DataFrame:
    logger.info(f"Cargando clientes: {RUTA_CLIENTES}")
    
    df = pd.read_excel(RUTA_CLIENTES, engine='openpyxl', dtype=str)
    logger.info(f"Total filas: {len(df):,}")
    logger.info(f"Columnas encontradas: {list(df.columns)}")
    
    col_odf = "ODF TRONCAL SALIDA(Puerto_Nemonico)"
    
    if col_odf not in df.columns:
        raise ValueError(f"No se encontro columna: {col_odf}. Columnas disponibles: {list(df.columns)}")
    
    # Clientes son los que tienen valor VACIO (NaN) en la columna ODF
    mask_vacio = df[col_odf].isna() | (df[col_odf].astype(str).str.strip() == "")
    df_clientes = df[mask_vacio].copy()
    
    logger.info(f"Clientes encontrados (ODF vacio): {len(df_clientes):,}")
    return df_clientes


def cargar_ports_diccionario() -> Tuple[Dict[tuple, dict], Dict[str, dict]]:
    """
    Carga los datos de ports y retorna:
    - datos_busqueda: Diccionario para búsqueda por coincidencia
    - ports_por_id: Diccionario que mapea ID con sus datos completos
    """
    logger.info(f"Cargando ports: {RUTA_PORTS}")
    
    if not RUTA_PORTS.exists():
        raise FileNotFoundError(f"No se encuentra: {RUTA_PORTS}")
    
    # Cargar todas las columnas para tener el ID
    df = pd.read_excel(RUTA_PORTS, engine='openpyxl', dtype=str)
    logger.info(f"Filas en ports: {len(df):,}")
    logger.info(f"Columnas en ports: {list(df.columns)}")
    
    # Identificar la columna ID (primera columna o columna llamada 'Id')
    col_id = df.columns[0]  # Primera columna
    if 'Id' in df.columns:
        col_id = 'Id'
    elif 'ID' in df.columns:
        col_id = 'ID'
    
    logger.info(f"Usando columna ID: '{col_id}'")
    
    datos_busqueda = {}
    ports_por_id = {}
    
    for _, row in df.iterrows():
        # Obtener el ID del puerto
        id_port = row.get(col_id, "")
        if pd.notna(id_port):
            id_port = str(id_port).strip()
        
        cambio = normalizar_texto(row.get(COLUMNAS_PORTS["cambio"], ""))
        id_servicio = normalizar_texto(row.get(COLUMNAS_PORTS["id_servicio"], ""))
        anillo = normalizar_texto(row.get(COLUMNAS_PORTS["anillo"], ""))
        
        if cambio or id_servicio or anillo:
            key = (cambio, id_servicio, anillo)
            if key not in datos_busqueda:
                datos_busqueda[key] = {
                    "cambio_sw": row.get(COLUMNAS_PORTS["cambio"], ""),
                    "id_servicio_sw": row.get(COLUMNAS_PORTS["id_servicio"], ""),
                    "anillo_sw": row.get(COLUMNAS_PORTS["anillo"], ""),
                    "id_port": id_port
                }
            
            # Guardar en diccionario por ID
            if id_port and id_port not in ports_por_id:
                ports_por_id[id_port] = {
                    "cambio": row.get(COLUMNAS_PORTS["cambio"], ""),
                    "id_servicio": row.get(COLUMNAS_PORTS["id_servicio"], ""),
                    "anillo": row.get(COLUMNAS_PORTS["anillo"], "")
                }
    
    logger.info(f"Registros unicos para busqueda: {len(datos_busqueda):,}")
    logger.info(f"Puertos por ID: {len(ports_por_id):,}")
    return datos_busqueda, ports_por_id


def buscar_coincidencia(
    cambio_user: str,
    id_user: str,
    anillo_user: str,
    datos_busqueda: Dict[tuple, dict]
) -> Tuple[str, str]:
    """
    Retorna: (ERROR, ID_PORT_ENCONTRADO)
    ERROR: V|Completo o F|...
    ID_PORT_ENCONTRADO: ID del puerto si hay coincidencia, sino vacío
    """
    
    c_norm = normalizar_texto(cambio_user)
    i_norm = normalizar_texto(id_user)
    a_norm = normalizar_texto(anillo_user)
    
    if not datos_busqueda:
        return "F|Sin datos en SW", ""
    
    key = (c_norm, i_norm, a_norm)
    
    if key in datos_busqueda:
        # Coincidencia exacta - retornar el ID del puerto
        return "V|Completo", datos_busqueda[key].get("id_port", "")
    
    errores = []
    
    # Verificar CAMBIO
    if c_norm:
        enc_c = any(k[0] == c_norm for k in datos_busqueda.keys())
        if not enc_c:
            errores.append(f"CAMBIO(sw:No existe, user:{cambio_user})")
    else:
        if cambio_user and cambio_user != "":
            errores.append(f"CAMBIO(sw:No existe, user:{cambio_user})")
    
    # Verificar ID SERVICIO
    if i_norm:
        enc_i = any(k[1] == i_norm for k in datos_busqueda.keys())
        if not enc_i:
            errores.append(f"ID_SERVICIO(sw:No existe, user:{id_user})")
    else:
        if id_user and id_user != "":
            errores.append(f"ID_SERVICIO(sw:No existe, user:{id_user})")
    
    # Verificar ANILLO
    if a_norm:
        enc_a = any(k[2] == a_norm for k in datos_busqueda.keys())
        if not enc_a:
            errores.append(f"ANILLO(sw:No existe, user:{anillo_user})")
    else:
        if anillo_user and anillo_user != "":
            errores.append(f"ANILLO(sw:No existe, user:{anillo_user})")
    
    # Si no hay errores de existencia, buscar coincidencias parciales
    id_encontrado = ""
    if not errores:
        for k in datos_busqueda.keys():
            if k[0] == c_norm and k[1] == i_norm and k[2] != a_norm:
                sw_valor = datos_busqueda[k]["anillo_sw"]
                errores.append(f"ANILLO(sw:{sw_valor}, user:{anillo_user})")
                id_encontrado = datos_busqueda[k].get("id_port", "")
                break
            elif k[0] == c_norm and k[2] == a_norm and k[1] != i_norm:
                sw_valor = datos_busqueda[k]["id_servicio_sw"]
                errores.append(f"ID_SERVICIO(sw:{sw_valor}, user:{id_user})")
                id_encontrado = datos_busqueda[k].get("id_port", "")
                break
            elif k[1] == i_norm and k[2] == a_norm and k[0] != c_norm:
                sw_valor = datos_busqueda[k]["cambio_sw"]
                errores.append(f"CAMBIO(sw:{sw_valor}, user:{cambio_user})")
                id_encontrado = datos_busqueda[k].get("id_port", "")
                break
    
    if errores:
        return f"F|{','.join(errores)}", id_encontrado
    
    return "F|Sin coincidencias", ""


def generar_reporte_normalizado(df_clientes: pd.DataFrame, datos_busqueda: Dict[tuple, dict]) -> pd.DataFrame:
    """
    Genera reporte en formato normalizado:
    ELEMENT | ID | VALUE | OWNER | ERROR | COMMENTS
    
    ELEMENT: Siempre será "PORT"
    ID: Si es V| trae ID del archivo port, si es F| trae TRAZA COMPLETA del usuario
    VALUE: Valores abreviados separados por | (CB:, ID:, AN:)
    OWNER: Responsable del archivo usuarios.xlsx
    ERROR: Resultado de la validación (V|Completo o F|...)
    COMMENTS: Columna "OBSERVACIONES" del archivo usuarios.xlsx
    """
    logger.info("Generando reporte normalizado...")
    
    if len(df_clientes) == 0:
        return pd.DataFrame()
    
    resultados = []
    
    for _, row in df_clientes.iterrows():
        # Obtener valores del cliente
        cambio_user = row.get("CAMBIO", "")
        id_user = row.get("ID SERVICIO", "")
        anillo_user = row.get("ANILLO", "")
        responsable = row.get("RESPONSABLE", "")
        observaciones = row.get("OBSERVACIONES", "")
        traza_completa = row.get("TRAZA COMPLETA", "")  # Columna TRAZA COMPLETA
        
        # Validar contra SW
        error, id_port_encontrado = buscar_coincidencia(cambio_user, id_user, anillo_user, datos_busqueda)
        
        # Determinar el valor para la columna ID
        if error.startswith("V|"):
            # Si es V|, usar el ID del puerto encontrado
            id_valor = id_port_encontrado
        else:
            # Si es F|, usar TRAZA COMPLETA
            id_valor = traza_completa if pd.notna(traza_completa) else ""
        
        # Construir VALUE con formato abreviado: CB:valor|ID:valor|AN:valor
        value_parts = []
        if cambio_user:
            value_parts.append(f"CB:{cambio_user}")
        if id_user:
            value_parts.append(f"ID:{id_user}")
        if anillo_user:
            value_parts.append(f"AN:{anillo_user}")
        
        value_str = "|".join(value_parts) if value_parts else "SIN_DATOS"
        
        # Crear registro único por cliente
        resultados.append({
            "ELEMENT": "PORT",
            "ID": id_valor,
            "VALUE": value_str,
            "OWNER": responsable,
            "ERROR": error,
            "COMMENTS": observaciones if pd.notna(observaciones) else ""
        })
    
    return pd.DataFrame(resultados)


def guardar_reporte_normalizado(df: pd.DataFrame, ruta: Path):
    if df.empty:
        logger.warning("No hay datos para guardar")
        return
    
    with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Reporte_Normalizado', index=False)
        
        ws = writer.sheets['Reporte_Normalizado']
        for column in ws.columns:
            max_len = 0
            col_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_len:
                        max_len = len(str(cell.value))
                except:
                    pass
            ws.column_dimensions[col_letter].width = min(max_len + 2, 50)
    
    logger.info(f"Reporte normalizado guardado: {ruta}")


def mostrar_resumen_normalizado(df: pd.DataFrame):
    if df.empty:
        print("\nNo hay datos para mostrar")
        return
    
    total = len(df)
    validos = df["ERROR"].str.startswith("V|", na=False).sum()
    fallidos = df["ERROR"].str.startswith("F|", na=False).sum()
    
    print("\n" + "=" * 60)
    print("RESUMEN DE VALIDACION - FORMATO NORMALIZADO")
    print("=" * 60)
    print(f"Total clientes evaluados: {total:,}")
    if total > 0:
        print(f"VALIDOS (V|Completo): {validos:,} ({validos/total*100:.1f}%)")
        print(f"FALLIDOS (F|): {fallidos:,} ({fallidos/total*100:.1f}%)")
    else:
        print("VALIDOS (V|Completo): 0")
        print("FALLIDOS (F|): 0")
    print("=" * 60)


def main():
    start_time = datetime.now()
    logger.info("INICIO - VALIDACION CLIENTES vs PORTS EXCEL (FORMATO NORMALIZADO)")
    
    try:
        df_clientes = cargar_clientes()
        
        if len(df_clientes) == 0:
            logger.warning("No hay clientes para procesar")
            return
        
        datos_busqueda, ports_por_id = cargar_ports_diccionario()
        
        if len(datos_busqueda) == 0:
            logger.warning("No hay datos en el archivo de ports")
            return
        
        df_resultado = generar_reporte_normalizado(df_clientes, datos_busqueda)
        
        if df_resultado.empty:
            logger.warning("No se generaron resultados")
            return
        
        guardar_reporte_normalizado(df_resultado, RUTA_SALIDA)
        
        mostrar_resumen_normalizado(df_resultado)
        
        print("\nMUESTRA DE RESULTADOS (primeros 10 registros):")
        print(df_resultado.head(10).to_string(index=False))
        
        elapsed = (datetime.now() - start_time).total_seconds()
        logger.info(f"Proceso completado en {elapsed:.2f} segundos")
        
    except Exception as e:
        logger.error(f"Error: {e}")
        raise


if __name__ == "__main__":
    main()