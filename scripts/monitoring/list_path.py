import re
import csv
from pathlib import Path
from typing import List, Dict, Tuple

# ========== CONFIGURACIÓN ==========
FOLDER_PATH = r"C:\A_GS1_PROYECTOS\path"
FILE_PATTERN = "*.txt"
OUTPUT_CSV = r"C:\A_GS1_PROYECTOS\extraidos.csv"

CAMPO_PATRONES = [
    ("IP DEMARCADOR:",   "IP_DEMARCADOR"),
    ("PTO METRO:",       "PTO_METRO"),
    ("EQUIPO:",          "NEMONICO"),
    ("ANILLO:",          "ANILLO"),
    ("ETIQUETA METRO:",  "ETIQUETA_METRO"),
    ("ID DE SERVICIO:",  "ID_SERVICIO"),
]
# ====================================

def colapsar_espacios(texto: str) -> str:
    return re.sub(r'[ \t]+', ' ', texto).strip()

def extraer_valor_general(linea: str, patron: str) -> str:
    m = re.search(r'^.*' + re.escape(patron) + r'\s*(.*)', linea)
    if not m:
        return ""
    valor = m.group(1).strip()
    return colapsar_espacios(valor)

def extraer_valor_pto_metro(linea: str) -> str:
    """
    Extrae el valor para PTO_METRO con reglas especiales:
    1. Busca "PTO METRO:" y si existe, toma lo que sigue.
    2. Si no, busca "GigabitEthernet" o "GE" y toma lo que sigue.
    3. Si encuentra "GigabitEthernet" o "GE" en el valor, elimina esa palabra y deja el resto.
    """
    # Primero intentar con el patrón estándar
    valor = extraer_valor_general(linea, "PTO METRO:")
    if valor:
        # Quitar "GigabitEthernet" o "GE" si están al inicio
        valor = re.sub(r'^\s*(?:GigabitEthernet|GE)\s*', '', valor)
        return colapsar_espacios(valor)

    # Si no se encontró "PTO METRO:", buscar "GigabitEthernet" o "GE" en la línea
    m = re.search(r'(?:GigabitEthernet|GE)\s*(.*)', linea)
    if m:
        valor = m.group(1).strip()
        return colapsar_espacios(valor)

    return ""

def leer_lineas_con_codificacion(ruta: Path) -> List[str]:
    codificaciones = ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252', 'iso-8859-1']
    for codif in codificaciones:
        try:
            with open(ruta, 'r', encoding=codif) as f:
                return f.readlines()
        except UnicodeDecodeError:
            continue
    raise UnicodeDecodeError(f"No se pudo decodificar {ruta.name}")

def procesar_archivo(ruta: Path, campos: List[Tuple[str, str]]) -> Dict[str, str]:
    try:
        lineas = leer_lineas_con_codificacion(ruta)
    except (UnicodeDecodeError, OSError) as e:
        print(f"Error en {ruta.name}: {e}")
        return {}

    registro = {"Archivo": ruta.name}

    for patron, columna in campos:
        if columna == "PTO_METRO":
            valor = ""
            for linea in lineas:
                extraido = extraer_valor_pto_metro(linea)
                if extraido:
                    valor = extraido
                    break
        else:
            valor = ""
            for linea in lineas:
                extraido = extraer_valor_general(linea, patron)
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
        print(f"No se encontraron archivos con el patrón '{FILE_PATTERN}'.")
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