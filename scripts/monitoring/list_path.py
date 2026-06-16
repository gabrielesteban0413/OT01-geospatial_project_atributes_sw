import re
import csv
from pathlib import Path
from typing import List, Dict, Tuple

# ========== CONFIGURACIÓN (modifica aquí) ==========
FOLDER_PATH = r"C:\A_GS1_PROYECTOS\path"
FILE_PATTERN = "*.txt"
OUTPUT_CSV = r"C:\A_GS1_PROYECTOS\extraidos.csv"

# Patrones de búsqueda y nombres de columna (ajusta según necesites)
CAMPO_PATRONES = [
    ("IP DEMARCADOR:",   "IP_DEMARCADOR"),
    ("PTO METRO:",       "PTO_METRO"),
    ("EQUIPO:",          "NEMONICO"),
    ("ANILLO:",          "ANILLO"),
    ("ETIQUETA METRO:",  "ETIQUETA_METRO"),
    ("ID DE SERVICIO:",  "ID_SERVICIO"),
]
# ====================================================

def colapsar_espacios(texto: str) -> str:
    """Reemplaza cualquier secuencia de espacios o tabuladores por un espacio simple."""
    return re.sub(r'[ \t]+', ' ', texto).strip()

def extraer_valor(linea: str, patron: str) -> str:
    """Busca el patrón en la línea y devuelve el texto que le sigue, con espacios colapsados."""
    m = re.search(r'^.*' + re.escape(patron) + r'\s*(.*)', linea)
    if not m:
        return ""
    valor = m.group(1).strip()
    return colapsar_espacios(valor)

def leer_lineas_con_codificacion(ruta: Path) -> List[str]:
    """
    Intenta leer el archivo con varias codificaciones comunes.
    Retorna las líneas del archivo.
    """
    codificaciones = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for codif in codificaciones:
        try:
            with open(ruta, 'r', encoding=codif) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"No se pudo decodificar {ruta.name} con ninguna codificación probada.")

def procesar_archivo(ruta: Path, campos: List[Tuple[str, str]]) -> Dict[str, str]:
    """
    Lee un archivo y extrae los campos definidos.
    Retorna un diccionario con el nombre del archivo y cada columna.
    """
    try:
        lineas = leer_lineas_con_codificacion(ruta)
    except UnicodeDecodeError as e:
        print(f"Error de decodificación en {ruta.name}: {e}")
        return {}
    except OSError as e:
        print(f"Error al leer {ruta.name}: {e}")
        return {}

    registro = {"Archivo": ruta.name}

    for patron, columna in campos:
        valor = ""
        for linea in lineas:
            extraido = extraer_valor(linea, patron)
            if extraido:
                valor = extraido
                break
        registro[columna] = valor

    return registro

def main() -> None:
    carpeta = Path(FOLDER_PATH)
    if not carpeta.is_dir():
        print(f"Error: La carpeta '{FOLDER_PATH}' no existe.")
        return

    archivos = list(carpeta.glob(FILE_PATTERN))
    if not archivos:
        print(f"No se encontraron archivos con el patrón '{FILE_PATTERN}' en '{FOLDER_PATH}'.")
        return

    resultados = []
    for archivo in archivos:
        registro = procesar_archivo(archivo, CAMPO_PATRONES)
        if registro:
            resultados.append(registro)

    if not resultados:
        print("No se extrajo ningún registro.")
        return

    columnas = ["Archivo"] + [col for _, col in CAMPO_PATRONES]

    try:
        with open(OUTPUT_CSV, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=columnas)
            writer.writeheader()
            writer.writerows(resultados)
        print(f"Extracción completada. Archivo generado: {OUTPUT_CSV}")
        print(f"Registros procesados: {len(resultados)} de {len(archivos)} archivos.")
    except OSError as e:
        print(f"Error al escribir el CSV: {e}")

if __name__ == "__main__":
    main()