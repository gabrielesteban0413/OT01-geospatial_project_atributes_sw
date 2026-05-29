import gc
import os
import re
import time
import warnings
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor, as_completed

import openpyxl
import pandas as pd
import tkinter as tk
from dotenv import load_dotenv
from openpyxl.styles import Border, Font, Side
from tkinter import simpledialog

load_dotenv()

BASE_PATH_STR = os.getenv("GS_BASE_PATH")
if not BASE_PATH_STR:
    raise ValueError("Variable GS_BASE_PATH no encontrada en .env")

BASE_PATH = Path(BASE_PATH_STR)
PROPIETARIO_VALIDO = os.getenv("PROPIETARIO_VALIDO", "DESCONOCIDO")

RUTA_MACRO          = BASE_PATH / "04-AuditoriaCalidadVisualAtributiva.xlsm"
RUTA_RESPONSABLES   = BASE_PATH / "03-linea_cables_responsables.xlsx"
RUTA_ODFS           = BASE_PATH / "04-Sync_SW_SP.xlsx"
RUTA_ASPHIA_CLIENTES= BASE_PATH / "03-asph_report_origin.xlsx"
RUTA_EDIFICIOS_BULK = BASE_PATH / "01-Bulk export of buildings.xlsx"
RUTA_FIBRAS_BULK    = BASE_PATH / "01-Bulk export of fibers.xlsx"
RUTA_EMPALMES_BULK  = BASE_PATH / "01-Bulk export of splices.xlsx"
RUTA_PORTS_BULK     = BASE_PATH / "01-Bulk export of ports.xlsx"

COLUMNAS_MASIVO = {
    "FIBRA": {
        "id": "id text",
        "nombre": "name",
        "colocacion": "colocacion type",
        "especificacion": "nombre especificación",
        "propietario": "propietario",
        "estado": "construction status",
        "nombre_ant": "nombre antiguo",
        "nodo_a": "nodo a",
        "nodo_b": "nodo b",
        "municipio_a": "municipio a id",
        "municipio_b": "municipio b id",
    },
    "EMPALME": {
        "id": "id texto",
        "nombre": "name",
        "especificacion": "nombre especificación",
        "propietario": "propietario",
        "estado": "construction status",
    },
    "EDIFICIO": {
        "id": "idtexto",
        "nombre": "name",
        "tipo": "type",
        "categoria": "categoria",
    },
    "PORT": {
        "id": "id",
        "description": "description",
        "tipo_senal": "tipo señal",
        "physical_status": "Physical Status",
        "type": "type",
        "connector_type": "connector type",
        "ancho_banda": "ancho de banda",
        "unidad_ancho_banda": "unidad ancho de banda",
        "service_level": "service level",
        "proyecto_red": "proyecto de red",
        "contratista": "contratista",
        "id_servicio": "id servicio",
        "equipo_central": "equipo central",
        "puerto_central": "puerto central",
        "tipo_servicio": "tipo servicio",
        "nombre_cliente": "nombre cliente",
        "direccion_cliente_id": "direccion cliente id",
        "comentarios": "comentarios",
        "connected": "connected?",
        "anillo": "anillo",
        "mit_card": "mit card",
        "direccion_puerto": "direccion puerto",
        "num_factibilidad": "numero de factibilidad-cable",
        "id_shelf": "id shelf",
        "cambio_service_manager": "cambio service manager",
    },
    "SHELF": {
        "id": "id shelf",
        "description": "description shelf",
        "estado": "construction status",
        "tipo": "tipo",
        "fabricante": "fabricante",
        "tecnologia": "tecnologia",
        "direccion_ip": "direccion ip",
        "nivel_bastidor": "nivel en bastidor",
        "contratista": "contratista",
        "comentarios": "comentarios",
        "nombre_anterior": "nombre anterior",
        "migracion_olt": "migracion olt",
        "id_bay": "bay.id",
    },
    "BAY": {
        "id": "bay.id",
        "number": "number",
        "description": "description bay",
        "estado": "construction status",
        "numero_salon": "numero de salon",
        "numero_fila": "numero de fila",
        "contratista": "contratista",
        "comentarios": "comentarios",
        "id_edificio": "building.id",
        "cambio_service_manager": "cambio service manager",
    },
}

COLUMNAS_ESTRUCTURA = {
    "nombre": "name",
    "floor_dimension_length": "floor dimension (length)",
    "floor_dimension_width": "floor dimension (width)",
    "idtexto": "idtexto",
    "number_of_rooms_length": "number of rooms (length)",
    "number_of_rooms_width": "number of rooms (width)",
    "number_of_floors": "number of floors",
}

JERARQUIA_FIBRA = {
    1:   {"campo": "Id",                            "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:   {"campo": "Segmento",                      "errores": []},
    3:   {"campo": "Nombre",                        "errores": ["NOMBRE (Valor Inválido)", "NOMBRE (Debe estar en MAYÚSCULAS)", "NOMBRE (Corregir a", "NOMBRE (en", "COLOCACION (Troncal)", "COLOCACION (Ultima Milla)", "DE", "CARACTER INVALIDO", "el símbolo / solo se permite", "nombre no puede ser solo", "NOMBRE"]},
    4:   {"campo": "Nombre Antiguo",                "errores": ["NOMBRE_ANT", "NOMBRE_ANT (Falta guion cuando NOMBRE tiene guion)"]},
    5:   {"campo": "Municipio A",                   "errores": ["MUNICIPIO (Vacío)", "MUNICIPIO A INCONSISTENTE", "MUNICIPIO"]},
    6:   {"campo": "Municipio B",                   "errores": ["MUNICIPIO (Vacío)", "MUNICIPIO B INCONSISTENTE", "MUNICIPIO"]},
    7:   {"campo": "Nodo A",                        "errores": ["NODO INVÁLIDO", "NODO A Y B"]},
    8:   {"campo": "Nodo B",                        "errores": ["NODO INVÁLIDO", "NODO A Y B"]},
    9:   {"campo": "Especificación",                "errores": ["ESPECI (senc)", "ESPECI (Tipo 2)", "ESPECI ('revisar'Ultima Milla cantidad de Hilos)", "ESPECI (NO acorde a lo establecido)", "ESPECI EA (revisar Tipo)", "ESPECI INCONSISTENTE EN CABLE TRONCAL", "ESPECI INCONSISTENTE"]},
    10:  {"campo": "Estado de Construcción",        "errores": ["STATUS ADMITIDOS (En servicio/Plan Abandono/Futuro)", "C.STATUS", "STATUS"]},
    11:  {"campo": "Propietario",                   "errores": [f"PROPIETARIO ({PROPIETARIO_VALIDO})"]},
    12:  {"campo": "Tipo de Colocacion",            "errores": ["TIPO DE COLOCACION", "COLOCACION"]},
    13:  {"campo": "Longitud Calculada de Cable",   "errores": []},
    14:  {"campo": "Longitud Calculada de Fibra",   "errores": []},
    15:  {"campo": "Perdida Db Medida por Inserccion", "errores": []},
    16:  {"campo": "Fecha de Instalacion",          "errores": []},
    17:  {"campo": "Abcisa Inicial",                "errores": []},
    18:  {"campo": "Abcisa Final",                  "errores": []},
    19:  {"campo": "Reservas",                      "errores": []},
    20:  {"campo": "Contratista",                   "errores": []},
    21:  {"campo": "Comentarios",                   "errores": []},
    999: {"campo": "MERGE",                         "errores": ["TIPO DE CABLE INCONSISTENTE", "MERGE", "PROBLEMA EN FIBRA", "PROBLEMA EN SUB-ANILLO", "RECOMENDACIÓN"]},
}

JERARQUIA_EMPALME = {
    1:   {"campo": "Id",                            "errores": ["CONSECUTIVO", "DUPLICADO"]},
    2:   {"campo": "Segmento",                      "errores": []},
    3:   {"campo": "Nombre",                        "errores": ["NOMBRE(ERRONEO)", "NOMBRE(CONTIENE ESPACIOS)", "NOMBRE(LETRA CONSECUTIVA)", "NOMBRE(CARACTERES ERRONEOS)", "usar:"]},
    4:   {"campo": "Propietario",                   "errores": [f"PROPIETARIO ({PROPIETARIO_VALIDO})"]},
    5:   {"campo": "Municipio",                     "errores": []},
    6:   {"campo": "Estado de Construcción",        "errores": ["STATUS (En servicio/Plan Abandono/Futuro)"]},
    7:   {"campo": "Tipo de Caja",                  "errores": []},
    8:   {"campo": "Especificación",                "errores": ["ESPECIFICACION ET ERRONEA", "ESPECIFICACION (NO acorde a lo establecido)", "TIPO DE GRUPO", "esperado:", "actual:"]},
    9:   {"campo": "E2",                            "errores": []},
    10:  {"campo": "Fecha de Instalacion",          "errores": []},
    11:  {"campo": "Distancia a Splitter 1er nivel","errores": []},
    12:  {"campo": "Distancia a Splitter 2do nivel","errores": []},
    13:  {"campo": "Disponible Para",               "errores": []},
    14:  {"campo": "Tipo Restriccion",              "errores": []},
    15:  {"campo": "Fecha Inicio Restriccion",      "errores": []},
    16:  {"campo": "Responsable Restriccion",       "errores": []},
    17:  {"campo": "Sangria?",                      "errores": []},
    18:  {"campo": "Método",                        "errores": []},
    19:  {"campo": "fx Available Mounting Positions","errores": []},
    20:  {"campo": "Tiene Riser?",                  "errores": []},
    21:  {"campo": "Red Intema",                    "errores": []},
    22:  {"campo": "Puertos Derivacion Libres",     "errores": []},
    23:  {"campo": "Contratista",                   "errores": []},
    999: {"campo": "MERGE",                         "errores": ["TIPO DE CABLE INCONSISTENTE", "INCONSISTENCIA MERGE"]},
}

JERARQUIA_EDIFICIO = {
    1:  {"campo": "Id",                     "errores": []},
    2:  {"campo": "Nombre",                 "errores": ["NOMBRE (Corregir a", "NOMBRE", "NOMBRE (Valor Inválido)", "NOMBRE (Sufijo con espacios)"]},
    3:  {"campo": "Tipo",                   "errores": ["TIPO DE GRUPO"]},
    4:  {"campo": "Estado de Construcción", "errores": ["ESTADO"]},
    5:  {"campo": "Categoria",              "errores": ["CATEGORIA (AAA)"]},
    6:  {"campo": "Direccion Edificio",     "errores": []},
    7:  {"campo": "Numero de Factibilidad", "errores": []},
    8:  {"campo": "Comentarios",            "errores": []},
    9:  {"campo": "Limite",                 "errores": []},
    10: {"campo": "Direccion Cliente ID",   "errores": []},
    11: {"campo": "Posición",               "errores": []},
    12: {"campo": "Posición",               "errores": []},
}

JERARQUIA_ESTRUCTURA = {
    1: {"campo": "ESTRUCTURA", "errores": ["ESTRUCTURA (NUMBER_OF + FLOOR + ROOMS = 0)", "ESTRUCTURA_FLOOR_DIMENSION", "ESTRUCTURA_NUMBER_OF_ROOMS", "NUM_OF:", "FLOOR:", "ROOMS:"]},
}

JERARQUIA_PORT = {
    1:  {"campo": "Id",                         "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:  {"campo": "Tipo Señal",                  "errores": ["TIPO SEÑAL INVALIDO", "PUERTO RX/TX IDENTIFICADO"]},
    3:  {"campo": "Description",                 "errores": ["DESCRIPTION VACIO"]},
    4:  {"campo": "Physical Status",             "errores": ["PHYSICAL STATUS INVALIDO", "PHYSICAL STATUS VACIO"]},
    5:  {"campo": "Type",                        "errores": ["TYPE INVALIDO", "TYPE VACIO", "TYPE (solo Equipo Cliente)"]},
    6:  {"campo": "Connector Type",              "errores": ["CONNECTOR TYPE INVALIDO", "CONNECTOR TYPE VACIO", "CONNECTOR TYPE (solo Unknown o Biconic)"]},
    7:  {"campo": "Ancho De Banda",              "errores": ["ANCHO DE BANDA INVALIDO"]},
    8:  {"campo": "Unidad Ancho De Banda",       "errores": ["UNIDAD ANCHO DE BANDA INVALIDA", "UNIDAD ANCHO DE BANDA VACIA"]},
    9:  {"campo": "Service Level",               "errores": ["SERVICE LEVEL INVALIDO", "SERVICE LEVEL VACIO", "SERVICE LEVEL (solo Fiber)"]},
    10: {"campo": "Proyecto De Red",             "errores": ["PROYECTO DE RED VACIO"]},
    11: {"campo": "Contratista",                 "errores": ["CONTRATISTA VACIO"]},
    12: {"campo": "Id Servicio",                 "errores": ["ID SERVICIO FORMATO INVÁLIDO"]},
    13: {"campo": "Equipo Central",              "errores": ["EQUIPO CENTRAL INVALIDO", "EQUIPO CENTRAL VACIO", "PREFIJO EQUIPO CENTRAL (BO,VI,PA)", "MINIMO 4 NUMEROS", "PRIMEROS 6 CARACTERES DEBEN SER LETRAS"]},
    14: {"campo": "Puerto Central",              "errores": ["PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL VACIO", "PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL INVALIDO", "PUERTO CENTRAL INVALIDO"]},
    15: {"campo": "Tipo Servicio",               "errores": ["TIPO SERVICIO VACIO"]},
    16: {"campo": "Nombre Cliente",              "errores": ["NOMBRE CLIENTE (Valor Inválido)", "NOMBRE CLIENTE (Debe estar en MAYÚSCULAS)", "NOMBRE (Corregir a", "NOMBRE CLIENTE VACIO", "NOMBRE CLIENTE", "NOMBRE CLIENTE (Sufijo)"]},
    17: {"campo": "Direccion Cliente Id",        "errores": ["DIRECCION CLIENTE"]},
    18: {"campo": "Comentarios",                 "errores": []},
    19: {"campo": "Connected?",                  "errores": ["CONNECTED INVALIDO"]},
    20: {"campo": "Anillo",                      "errores": ["ANILLO VACIO"]},
    21: {"campo": "Mit Card",                    "errores": ["MIT CARD VACIO"]},
    22: {"campo": "Direccion Puerto",            "errores": ["DIRECCION PUERTO VACIO"]},
    23: {"campo": "Numero De Factibilidad-Cable","errores": ["NUMERO FACTIBILIDAD VACIO"]},
    24: {"campo": "Id shelf",                    "errores": ["ID SHELF VACIO", "ID SHELF INVALIDO"]},
    25: {"campo": "Cambio Service Manager",      "errores": ["SERVICE MANAGER"]},
    26: {"campo": "Id Servicio",                 "errores": ["ID SERVICIO"]},
}

JERARQUIA_SHELF = {
    1:  {"campo": "Id",                 "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2:  {"campo": "Description",        "errores": ["DESCRIPTION VACIO", "DESCRIPTION INVALIDA", "DESCRIPTION CON ESPACIOS"]},
    3:  {"campo": "Construction Status","errores": ["CONSTRUCTION STATUS INVALIDO"]},
    4:  {"campo": "Tipo",               "errores": ["TIPO INVALIDO"]},
    5:  {"campo": "Fabricante",         "errores": []},
    6:  {"campo": "Tecnologia",         "errores": ["TECNOLOGIA VACIO"]},
    7:  {"campo": "Direccion Ip",       "errores": ["DIRECCION IP INVALIDA"]},
    8:  {"campo": "Nivel En Bastidor",  "errores": ["NIVEL BASTIDOR INVALIDO"]},
    9:  {"campo": "Contratista",        "errores": ["CONTRATISTA VACIO"]},
    10: {"campo": "Comentarios",        "errores": []},
    11: {"campo": "Nombre Anterior",    "errores": []},
    12: {"campo": "Migracion Olt",      "errores": ["MIGRACION OLT INVALIDA"]},
    13: {"campo": "Bay.id",             "errores": ["BAY ID VACIO", "BAY ID INVALIDO"]},
}

JERARQUIA_BAY = {
    1: {"campo": "Id",                  "errores": ["DUPLICADO", "CONSECUTIVO"]},
    2: {"campo": "Number",              "errores": ["NUMBER INVALIDO"]},
    3: {"campo": "Description",         "errores": ["DESCRIPTION VACIO", "DESCRIPTION CON ESPACIADO", "DESCRIPTION INVÁLIDA"]},
    4: {"campo": "Construction Status", "errores": ["CONSTRUCTION STATUS INVALIDO"]},
    5: {"campo": "Numero De Salon",     "errores": ["NUMERO SALON INVALIDO"]},
    6: {"campo": "Numero De Fila",      "errores": ["NUMERO FILA INVALIDO"]},
    7: {"campo": "Contratista",         "errores": ["CONTRATISTA VACIO"]},
    8: {"campo": "Comentarios",         "errores": []},
    9: {"campo": "Building.id",         "errores": ["BUILDING ID VACIO", "BUILDING ID INVALIDO"]},
}

RE_NOMBRE_FIBRA_TRONCAL        = re.compile(r"^[\dP/]+$")
RE_NOMBRE_FIBRA_TRONCAL_SIMPLE = re.compile(r"^[\d]+$")
RE_NOMBRE_FIBRA_UM_CON_P       = re.compile(r"^[\d\-P/]+$")
RE_NOMBRE_FIBRA_UM_SIN_P       = re.compile(r"^[\d\-]+$")
RE_NOMBRE_EMPALME_VALIDO       = re.compile(r"^(ET|NAP|EA|OB)[0-9_\-P]+$")
RE_NOMBRE_EMPALME_SUB          = re.compile(r"^(ET|NAP)\d+\-SUB_\d+$")
RE_NOMBRE_EMPALME_BASE         = re.compile(r"^(ET|NAP|EA|OB)[0-9\-_P]*")
RE_NOMBRE_FIBRA_BASE           = re.compile(r'^(\d+(?:-\d+)?)')
RE_NOMBRE_FIBRA_NUMERO         = re.compile(r'^(\d+)')
RE_ESPEC_HILOS                 = re.compile(r'(\d+)\s+hilos', re.IGNORECASE)
RE_NOMBRE_ANT_UM               = re.compile(r"^[\d\-P\#/]+$|^SUB$")
RE_NOMBRE_ANT_NO_UM            = re.compile(r"^[\d\-P\#]+$|^SUB$")
RE_NUMERO_SALON                = re.compile(r'^OF \d+$')
RE_DESCRIPCION_BASTIDOR        = re.compile(r'^BASTIDOR CLIENTE (\d+)$')
RE_DESCRIPCION_DEMARCADOR      = re.compile(r'^DEMARCADOR CLIENTE (\d+)$')
RE_IPV4                        = re.compile(r'^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(/\d{1,2})?$')

PREFIJOS_EXCLUIR_VALIDACION = {"C.C", "ANT", "(SED)", "(REDP)", "COLEGIO", "TM", "EDIFICIO", "IED", "TORRE","NODO"}
PREFIJOS_ID_SERVICIO        = ("CAV", "ID", "IP", "PI", "SEG", "TI")
PREFIJOS_SERVICE_MANAGER    = ("OC-", "C", "WO", "APR")
PREFIJOS_EQUIPO_CENTRAL     = ("BO", "VI", "PA","BOBSH930601","BOCHHU930601","CHCHHU930601","CUZPHU930601","SOPOHU930301")

def es_valor_cero(valor_raw, valor_str: str) -> bool:
    if valor_str == "0":
        return True
    if isinstance(valor_raw, (int, float)) and valor_raw == 0:
        return True
    try:
        return float(valor_str) == 0.0
    except (ValueError, TypeError):
        return False

def normalizar_id(valor) -> str:
    if valor is None or valor == "":
        return ""
    if isinstance(valor, (int, float)):
        try:
            return str(int(valor))
        except ValueError:
            return str(valor).strip()
    return str(valor).strip()

def traducir_estado_construccion(estado, modo: str) -> str:
    if not estado:
        return ""
    estado = str(estado).strip()
    if modo == "masivo":
        return {"is": "En servicio", "pa": "Plan Abandono", "pf": "Plan Uso Futuro"}.get(estado.lower(), estado)
    return estado

def extraer_indices_columnas(worksheet) -> dict:
    return {
        str(cell.value).strip().lower(): idx - 1
        for idx, cell in enumerate(worksheet[1], start=1)
        if cell.value
    }

def validar_sufijos_nombre_empresa(nombre):
    if not nombre or pd.isna(nombre):
        return True, None, None
    nombre_str = str(nombre).strip()
    nombre_upper = nombre_str.upper()
    patrones_sufijos = [
        (r'\s+S\s+A\s+S?$', 'SA'),
        (r'\s+L\s+T\s+D\s+A$', 'LTDA'),
        (r'\s+L\s+T\s+D$', 'LTD'),
        (r'\s+I\s+N\s+C$', 'INC'),
        (r'\s+C\s+O\s+R\s+P$', 'CORP'),
        (r'\s+C\s+I\s+E$', 'CIE'),
        (r'\s+E\s+U$', 'EU'),
        (r'\s+N\s+I\s+T\s*\d*$', 'NIT'),
    ]
    for patron, _ in patrones_sufijos:
        if re.search(patron, nombre_upper):
            nombre_sin_espacios = re.sub(r'\s+', '', nombre_upper)
            msg = f"{JERARQUIA_EDIFICIO[2]['errores'][3]}'"
            return False, msg, nombre_sin_espacios
    return True, None, None

def validar_formato_nombre_cliente(nombre_original: str):
    nombre_upper = nombre_original.upper()
    errores = []
    patron_caracteres = r"^[A-Za-z0-9 ñÑ()&.,\-]+$"
    if not re.match(patron_caracteres, nombre_original):
        caracteres_invalidos = re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nombre_original)
        unicos = "".join(sorted(set(caracteres_invalidos)))
        if unicos:
            errores.append(f"CARACTER INVALIDO: '{unicos}'")
    if nombre_original != nombre_upper:
        errores.append(f"{nombre_upper}")
    if errores:
        return False, f"NOMBRE CLIENTE: {' | '.join(errores)}", None
    return True, None, None

def validar_unidad_ancho_banda(valor_unidad):
    if pd.isna(valor_unidad) or str(valor_unidad).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_unidad).strip()
    if valor_str.lower() in ("megabyte", "kilobyte"):
        return "VÁLIDO", valor_str
    sugerencias = {
        "mb": "Megabyte", "mbps": "Megabyte", "mbyte": "Megabyte",
        "kb": "Kilobyte", "kbps": "Kilobyte", "kbyte": "Kilobyte",
    }
    for clave, sugerencia in sugerencias.items():
        if clave in valor_str.lower():
            return "UNIDAD_INCORRECTA", sugerencia
    return "INVÁLIDO", "Megabyte o Kilobyte"

def validar_physical_status(valor_status):
    if pd.isna(valor_status) or str(valor_status).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_status).strip()
    if valor_str.lower() in ("in service", "spare"):
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "In Service o Spare"

def validar_service_level_fiber(valor_service_level):
    if pd.isna(valor_service_level) or str(valor_service_level).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_service_level).strip()
    if valor_str.lower() == "fiber":
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Fiber"

def validar_type_equipo_cliente(valor_type):
    if pd.isna(valor_type) or str(valor_type).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_type).strip()
    if valor_str.lower() == "equipo cliente":
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Equipo Cliente"

def validar_connector_type(valor_connector_type):
    if pd.isna(valor_connector_type) or str(valor_connector_type).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_connector_type).strip()
    if valor_str.lower() in ("unknown", "biconic"):
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Unknown o Biconic"

def validar_ipv4(direccion_ip) -> tuple:
    if pd.isna(direccion_ip) or str(direccion_ip).strip() == "":
        return "VACÍO", True
    direccion_str = str(direccion_ip).strip()
    match = RE_IPV4.match(direccion_str)
    if not match:
        return "INVÁLIDO", False
    for i in range(1, 5):
        if not (0 <= int(match.group(i)) <= 255):
            return "INVÁLIDO", False
    if match.group(5):
        mascara = int(match.group(5)[1:])
        if not (0 <= mascara <= 32):
            return "INVÁLIDO", False
    return "VÁLIDO", True

def validar_descripcion_bastidor(descripcion) -> list:
    if pd.isna(descripcion) or str(descripcion).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(descripcion).strip().upper()
    if desc_str == "BASTIDOR CLIENTE":
        return []
    match = RE_DESCRIPCION_BASTIDOR.match(desc_str)
    if match and int(match.group(1)) >= 2:
        return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(descripcion).strip()}'"]

def validar_descripcion_demarcador(descripcion) -> list:
    if pd.isna(descripcion) or str(descripcion).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(descripcion).strip().upper()
    if desc_str == "DEMARCADOR CLIENTE":
        return []
    match = RE_DESCRIPCION_DEMARCADOR.match(desc_str)
    if match and int(match.group(1)) >= 2:
        return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(descripcion).strip()}'"]

def validar_equipo_central(valor):

    if pd.isna(valor):
        return "VÁLIDO", ""
    valor_str = str(valor).strip()
    if valor_str == "" or valor_str == "0":
        return "VÁLIDO", ""

    if not re.match(r'^[A-Z0-9]+$', valor_str):
        return "CARACTERES_INVALIDOS", "Solo letras mayúsculas y números"

    if len(valor_str) < 6:
        return "LONGITUD_INSUFICIENTE", "Debe tener al menos 6 caracteres"

    primeros_6 = valor_str[:6]
    if not primeros_6.isalpha():
        return "PRIMEROS_6_NO_LETRAS", "Los primeros 6 caracteres deben ser letras"

    prefijo = primeros_6[:2]
    if prefijo not in PREFIJOS_EQUIPO_CENTRAL:
        return "PREFIJO_INVALIDO", f"Las primeras dos letras deben ser {', '.join(PREFIJOS_EQUIPO_CENTRAL)}"

    digitos = sum(c.isdigit() for c in valor_str)
    if digitos < 4:
        return "NUMEROS_INSUFICIENTES", "Debe contener al menos 4 números"
    return "VÁLIDO", valor_str

def validar_puerto_central(valor):

    if pd.isna(valor):
        return "VÁLIDO", ""
    valor_str = str(valor).strip()
    if valor_str == "" or valor_str == "0":
        return "VÁLIDO", ""
   
    if not re.match(r'^[0-9/-]+$', valor_str):
        return "CARACTERES_INVALIDOS", "Solo números, / y -"

    if re.search(r'[^0-9/-]', valor_str):
        return "CARACTERES_INVALIDOS", "Solo números, / y -"

    if '/' not in valor_str and '-' not in valor_str:
        return "SIN_SEPARADOR", "Debe contener al menos un / o -"
  
    separadores = valor_str.count('/') + valor_str.count('-')
    if separadores > 2:
        return "DEMASIADOS_SEPARADORES", "Máximo 2 separadores (/ o -)"

    if not valor_str[0].isdigit():
        return "INICIO_INVALIDO", "Debe empezar con dígito 0-9"
  
    segmentos = re.split(r'[/-]', valor_str)
    for seg in segmentos:
        if seg and not seg.isdigit():
            return "SEGMENTO_NO_NUMERICO", "Todos los segmentos deben ser números"
    return "VÁLIDO", valor_str

def extraer_cantidad_hilos(especificacion: str) -> int:
    match = RE_ESPEC_HILOS.search(str(especificacion))
    return int(match.group(1)) if match else 0

def validar_nombre_antiguo_fibra(nombre_ant, tipo_colocacion: str) -> bool:
    if pd.isna(nombre_ant) or str(nombre_ant).strip() == "":
        return True
    nombre_ant_str = str(nombre_ant).strip().upper()
    patron = RE_NOMBRE_ANT_UM if tipo_colocacion == "Ultima Milla" else RE_NOMBRE_ANT_NO_UM
    return bool(patron.match(nombre_ant_str))

def validar_nombre_fibra(nombre: str, tipo_colocacion: str) -> list:
    errores = []
    nombre_str = str(nombre).strip().upper()
    tipo_str = str(tipo_colocacion).strip()
    if tipo_str in ("Troncal-Acceso", "Troncal-Transporte"):
        if "-" in nombre_str:
            if "SUB" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} el guion solo se permite con 'SUB')")
            elif not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str.replace("-SUB", "")):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (formato inválido con SUB en {tipo_str})")
        elif "P" in nombre_str:
            if not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} debe contener solo números, P y / cuando hay P)")
            if "/" in nombre_str and "P" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][8]} (en {tipo_str})")
            if nombre_str == "/" or nombre_str.replace("/", "") == "":
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][9]} (en {tipo_str})")
        else:
            if not RE_NOMBRE_FIBRA_TRONCAL_SIMPLE.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_str} debe contener solo números cuando no hay P)")
    else:
        if "P" in nombre_str:
            if not RE_NOMBRE_FIBRA_UM_CON_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][7]} (nombre con P)")
            if "/" in nombre_str and "P" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][8]}")
            if nombre_str == "/" or nombre_str.replace("/", "").replace("-", "") == "":
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][9]}")
        else:
            if not RE_NOMBRE_FIBRA_UM_SIN_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][7]} (nombre sin P)")
    return errores

def obtener_nombre_base_fibra(nombre_completo: str) -> str:
    if "-" not in nombre_completo:
        return nombre_completo
    partes = nombre_completo.split("-")
    if len(partes) == 2 and partes[1].isdigit():
        return nombre_completo
    if len(partes) == 3 and all(p.isdigit() for p in partes[1:]):
        return f"{partes[0]}-{partes[1]}"
    match = RE_NOMBRE_FIBRA_BASE.match(nombre_completo)
    return match.group(1) if match else nombre_completo

def validar_nombre_empalme(nombre: str) -> list:
    errores = []
    nombre_str = str(nombre).strip().upper()
    if " " in nombre_str:
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][1]}")
    if "-SUB" in nombre_str:
        if not RE_NOMBRE_EMPALME_SUB.match(nombre_str):
            errores.append(f"{JERARQUIA_EMPALME[3]['errores'][3]}: Formato inválido para empalme con SUB")
    else:
        if not RE_NOMBRE_EMPALME_VALIDO.match(nombre_str):
            caracteres_invalidos = re.sub(r"(ET|NAP|EA|OB)[0-9_\-P]", "", nombre_str)
            unicos = "".join(sorted(set(caracteres_invalidos)))
            if unicos:
                errores.append(f"{JERARQUIA_EMPALME[3]['errores'][3]}: '{unicos}'")
    if re.search(r'([A-Z])\1', nombre_str):
        letras_repetidas = {m.group(0) for m in re.finditer(r'([A-Z])\1', nombre_str)}
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][2]}: {', '.join(letras_repetidas)}")
    if nombre_str.startswith("ET") and "-" in nombre_str and "-SUB" not in nombre_str:
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][0]}")
    return errores

def validar_especificacion_empalme(nombre: str, especificacion: str) -> list:
    errores = []
    nombre_upper = nombre.upper()
    especificacion_str = str(especificacion).strip()
    ESPEC_VALIDAS = {
        "Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente",
        "Nap Sotano 12", "Nap Sotano 24", "Nap Sotano 36",
        "Nap Sotano 48", "Nap Sotano 96", "Nap Sotano 144",
    }
    if especificacion_str not in ESPEC_VALIDAS:
        if especificacion_str.startswith("Nap Sotano "):
            numero_str = especificacion_str.split("Nap Sotano ")[1]
            if not numero_str.isdigit():
                errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
        else:
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    if nombre_upper.startswith("ET"):
        if especificacion_str not in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente"} and not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][0]} (actual: {especificacion_str})")
    elif nombre_upper.startswith("NAP"):
        if not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"ESPECIFICACION NAP ERRONEA (actual: {especificacion_str}, esperado: 'Nap Sotano X')")
    elif nombre_upper.startswith("EA"):
        if especificacion_str != "Tipo 2":
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    return errores

def determinar_tipo_empalme_esperado(especificacion_fibra: str) -> str:
    esp = especificacion_fibra.lower()
    if esp.startswith("ductado 144"):
        return "Tipo 0"
    if esp.startswith("ductado 96"):
        return "Tipo 1"
    if any(esp.startswith(f"ductado {n}") for n in ("12", "24", "36", "48")):
        return "Tipo 2"
    if "empalme cliente" in esp:
        return "Empalme Cliente"
    if "nap sotano" in esp:
        return "Nap Sotano"
    return "Desconocido"

def normalizar_nombre_empresa(nombre: str) -> str:
    if not nombre:
        return ""
    nombre_upper = nombre.upper().strip()
    sufijos = [
        r'\s+S\.?A\.?S\.?', r'\s+S\.?A\.?', r'\s+L\.?T\.?D\.?A\.?',
        r'\s+E\.?U\.?', r'\s+N\.?I\.?T\.?\s*\d+', r'\s+&?\s*CIE\.?',
        r'\s+LTDA\.?', r'\s+LTD\.?', r'\s+INC\.?', r'\s+CORP\.?',
        r'\s+S\.?C\.?S\.?', r'\s+E\.?I\.?R\.?L\.?',
    ]
    for sufijo in sufijos:
        nombre_upper = re.sub(sufijo, '', nombre_upper)
    nombre_upper = re.sub(r'[.,]', '', nombre_upper)
    nombre_upper = re.sub(r'\s+', ' ', nombre_upper).strip()
    return nombre_upper

def cargar_om_desde_macro(ruta_macro: Path, hoja: str = "OM") -> tuple:
    ids_om, exc_ed = set(), set()
    try:
        wb = openpyxl.load_workbook(ruta_macro, data_only=True, read_only=True, keep_vba=True)
    except FileNotFoundError:
        return ids_om, exc_ed
    if hoja not in wb.sheetnames:
        wb.close()
        return ids_om, exc_ed
    ws = wb[hoja]
    idx_id = idx_elem = idx_tb = None
    for idx, cell in enumerate(ws[1], start=1):
        v = str(cell.value).strip().lower() if cell.value else ""
        if v == "id" and idx_id is None:
            idx_id = idx
        elif v in ("elemento", "element") and idx_elem is None:
            idx_elem = idx
        elif v == "tb" and idx_tb is None:
            idx_tb = idx
    for row in ws.iter_rows(min_row=2, values_only=True):
        if idx_id:
            val = row[idx_id - 1]
            if val:
                ids_om.add(normalizar_id(val))
        if idx_elem and idx_tb:
            elem = row[idx_elem - 1]
            tb = row[idx_tb - 1]
            if elem and tb and str(tb).strip().upper() == "ED1":
                exc_ed.add(str(elem).strip())
    wb.close()
    return ids_om, exc_ed

def cargar_responsables_por_cable(ruta: Path, hoja: str, columna_clave: str = "cable acceso") -> dict:
    try:
        df = pd.read_excel(ruta, sheet_name=hoja, engine="openpyxl", dtype=str)
    except (FileNotFoundError, ValueError):
        return {}
    df.columns = [str(c).strip().lower() for c in df.columns]
    clave_norm = columna_clave.lower()
    try:
        idx_clave = df.columns.get_loc(clave_norm)
        idx_resp = df.columns.get_loc("responsable red externa")
    except KeyError:
        return {}
    responsables = {}
    for _, row in df.iterrows():
        if clave_norm == "cable acceso":
            clave = str(row.iloc[idx_clave]).split("/")[0].strip() if pd.notna(row.iloc[idx_clave]) else ""
        elif clave_norm == "id edifico":
            clave = normalizar_id(row.iloc[idx_clave])
        else:
            clave = str(row.iloc[idx_clave]).strip() if pd.notna(row.iloc[idx_clave]) else ""
        if clave:
            responsables[clave] = str(row.iloc[idx_resp]).strip() if pd.notna(row.iloc[idx_resp]) else ""
    return responsables

def cargar_especificaciones_fibra(ruta_archivo: Path, ids_om: set, modo: str, incluir_colocacion: bool = False):
    especificaciones = {}
    colocacion_por_fibra = {} if incluir_colocacion else None
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        col = extraer_indices_columnas(ws)
        cols = COLUMNAS_MASIVO["FIBRA"]
        idx_id        = col.get(cols["id"], -1)
        idx_estado    = col.get(cols["estado"], -1)
        idx_nombre    = col.get(cols["nombre"], -1)
        idx_espec     = col.get(cols["especificacion"], -1)
        idx_colocacion= col.get(cols["colocacion"], -1) if incluir_colocacion else -1
        ids_om_set = set(ids_om)
        for fila in ws.iter_rows(min_row=2, values_only=True):
            if idx_id == -1 or idx_estado == -1:
                continue
            idv = normalizar_id(fila[idx_id])
            estado_raw = fila[idx_estado]
            if not idv or idv in ids_om_set or estado_raw in (None, ""):
                continue
            if idx_nombre == -1:
                continue
            nombre = str(fila[idx_nombre] or "").strip()
            base = nombre.split("/")[0].strip()
            esp_raw = str(fila[idx_espec] or "").strip() if idx_espec != -1 else ""
            if base:
                especificaciones[base] = esp_raw.lower()
                if incluir_colocacion and idx_colocacion != -1:
                    colocacion_por_fibra[base] = str(fila[idx_colocacion] or "").strip()
        wb.close()
    except Exception:
        pass
    return (especificaciones, colocacion_por_fibra) if incluir_colocacion else especificaciones

def cargar_validaciones_odfs(ruta_odfs: Path) -> dict:
    odfs_data = {}
    try:
        wb = openpyxl.load_workbook(ruta_odfs, data_only=True, read_only=True)
        if "ODFS" not in wb.sheetnames:
            wb.close()
            return odfs_data
        ws = wb["ODFS"]
        headers = {str(cell.value).strip().lower(): idx for idx, cell in enumerate(ws[1], start=1) if cell.value}
        columnas_requeridas = ('validacion', 'troncal', 'ids_distribucion')
        if not all(col in headers for col in columnas_requeridas):
            wb.close()
            return odfs_data
        idx_validacion     = headers['validacion']
        idx_troncal        = headers['troncal']
        idx_ids_distribucion = headers['ids_distribucion']
        for row in ws.iter_rows(min_row=2, values_only=True):
            validacion = str(row[idx_validacion]).strip().upper() if row[idx_validacion] else ""
            if validacion != "OK":
                continue
            troncal = str(row[idx_troncal]).strip() if row[idx_troncal] else ""
            ids_raw = str(row[idx_ids_distribucion]).strip() if row[idx_ids_distribucion] else ""
            if not troncal or not ids_raw:
                continue
            ids_list = [i.strip() for i in ids_raw.replace('"', '').split('|') if i.strip()]
            if ids_list:
                if troncal in odfs_data:
                    odfs_data[troncal].update(ids_list)
                else:
                    odfs_data[troncal] = set(ids_list)
        wb.close()
    except Exception:
        pass
    return odfs_data

def cargar_nombres_edificios(ruta_ed: Path, modo: str) -> tuple:
    nombres_edificios = {}
    nombres_normalizados = defaultdict(set)
    try:
        wb = openpyxl.load_workbook(ruta_ed, data_only=True, read_only=True)
        nombre_hoja = next(
            (h for h in wb.sheetnames if any(kw in h.lower() for kw in ("edif", "export", "output"))),
            wb.sheetnames[0],
        )
        ws = wb[nombre_hoja]
        col = extraer_indices_columnas(ws)
        cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
        idx_id     = col.get(cols_ed["id"], -1)
        idx_nombre = col.get(cols_ed["nombre"], -1)
        if idx_id == -1 or idx_nombre == -1:
            wb.close()
            return nombres_edificios, nombres_normalizados
        for row in ws.iter_rows(min_row=2, values_only=True):
            idv = normalizar_id(row[idx_id]) if idx_id < len(row) else ""
            if not idv:
                continue
            nombre = str(row[idx_nombre] or "").strip() if idx_nombre < len(row) else ""
            if nombre:
                nombre_upper = nombre.upper()
                nombres_edificios[idv] = nombre_upper
                nombres_normalizados[normalizar_nombre_empresa(nombre_upper)].add(idv)
        wb.close()
    except Exception:
        pass
    return nombres_edificios, nombres_normalizados

def validar_nombre_cliente_contra_edificio(
    nombre_cliente, building_id, physical_status, comentarios,
    nombres_edificios: dict, nombres_normalizados: dict
) -> tuple:
    condicion_validar = False
    ps_upper = str(physical_status).strip().upper() if physical_status is not None else ""
    if ps_upper in ("IN SERVICE", "IS"):
        condicion_validar = True
    if comentarios and pd.notna(comentarios):
        if "ADICION" in str(comentarios).strip().upper():
            condicion_validar = True
    if not condicion_validar:
        return True, None, None, None
    if pd.isna(nombre_cliente) or str(nombre_cliente).strip() == "":
        return False, "", "NOMBRE CLIENTE VACIO", None
    nombre_original = str(nombre_cliente).strip()
    valido_sufijos, msg_sufijos, nombre_corregido = validar_sufijos_nombre_empresa(nombre_original)
    if not valido_sufijos:
        return False, nombre_original, msg_sufijos, nombre_corregido
    formato_valido, msg_formato, _ = validar_formato_nombre_cliente(nombre_original)
    if not formato_valido:
        return False, nombre_original, msg_formato, None
    skip_building_match = False
    building_name = None
    if building_id and building_id in nombres_edificios:
        building_name = nombres_edificios[building_id]
        building_norm = normalizar_nombre_empresa(building_name)
        if any(building_norm.startswith(p) or building_name.upper().startswith(p) for p in PREFIJOS_EXCLUIR_VALIDACION):
            skip_building_match = True
    if not skip_building_match and (not building_id or building_id not in nombres_edificios):
        cliente_norm = normalizar_nombre_empresa(nombre_original.upper())
        if cliente_norm in nombres_normalizados:
            edificios_encontrados = list(nombres_normalizados[cliente_norm])
            if edificios_encontrados:
                building_name = nombres_edificios.get(edificios_encontrados[0], "")
                if building_name:
                    building_norm = normalizar_nombre_empresa(building_name)
                    if any(building_norm.startswith(p) or building_name.upper().startswith(p) for p in PREFIJOS_EXCLUIR_VALIDACION):
                        skip_building_match = True
    if not skip_building_match and building_name:
        cliente_norm  = normalizar_nombre_empresa(nombre_original.upper())
        edificio_norm = normalizar_nombre_empresa(building_name.upper())
        if cliente_norm in edificio_norm or edificio_norm in cliente_norm or cliente_norm == edificio_norm:
            return True, None, None, None
        error_msg = f"{JERARQUIA_PORT[16]['errores'][4]} | Posible esperado: '{building_name}'"
        return False, nombre_original, error_msg, building_name
    return True, None, None, None

class OrganizadorErrores:
    def __init__(self):
        self._errores_por_id: dict = defaultdict(list)
    def agregar_error(self, id_elemento: str, nombre_entidad: str, valores_error):
        prioridad, campo = self._resolver_prioridad_campo(nombre_entidad, valores_error)
        self._errores_por_id[id_elemento].append((prioridad, campo, nombre_entidad, valores_error))
    def _resolver_prioridad_campo(self, entidad: str, valores_error) -> tuple:
        mensaje = self._extraer_mensaje(valores_error)
        jerarquia = self._seleccionar_jerarquia(entidad, mensaje)
        if jerarquia is None:
            return (999, "Otros")
        mensaje_upper = mensaje.upper()
        for prioridad, config in jerarquia.items():
            for patron in config["errores"]:
                if patron.upper() in mensaje_upper:
                    return (prioridad, config["campo"])
        return (999, "Otros")
    @staticmethod
    def _seleccionar_jerarquia(entidad: str, mensaje: str):
        mapa = {
            "FIBRA":      JERARQUIA_FIBRA,
            "EMPALME":    JERARQUIA_EMPALME,
            "EDIFICIO":   JERARQUIA_EDIFICIO,
            "PORT":       JERARQUIA_PORT,
            "SHELF":      JERARQUIA_SHELF,
            "BAY":        JERARQUIA_BAY,
        }
        if entidad in mapa:
            return mapa[entidad]
        if entidad == "MERGE FI-EM":
            return JERARQUIA_FIBRA if "MERGE" in mensaje else JERARQUIA_EMPALME
        if "ESTRUCTURA" in entidad:
            return JERARQUIA_ESTRUCTURA
        return None
    @staticmethod
    def _extraer_mensaje(valores_error) -> str:
        if isinstance(valores_error, (list, tuple)):
            if len(valores_error) > 3:
                return str(valores_error[3])
            if len(valores_error) == 1:
                return str(valores_error[0])
            return str(valores_error)
        return str(valores_error)
    def obtener_errores_organizados(self) -> list:
        errores_finales = []
        for idv in sorted(self._errores_por_id.keys()):
            errores_id = self._errores_por_id[idv]
            errores_por_prioridad = defaultdict(list)
            for prio, campo, entidad, valores in errores_id:
                errores_por_prioridad[prio].append((campo, entidad, valores))
            mejor_prioridad = min(errores_por_prioridad.keys())
            mejores_errores = errores_por_prioridad[mejor_prioridad]
            agrupados = defaultdict(list)
            for campo, entidad, valores in mejores_errores:
                agrupados[(entidad, campo)].append(valores)
            for (entidad, _), lista_valores in agrupados.items():
                if len(lista_valores) == 1:
                    errores_finales.append((entidad, lista_valores[0]))
                else:
                    errores_finales.append((entidad, self._combinar_errores(lista_valores)))
        return errores_finales
    @staticmethod
    def _combinar_errores(lista_valores) -> tuple:
        id_val    = lista_valores[0][0]
        value_val = lista_valores[0][1]
        owner_val = lista_valores[0][2]
        mensajes = {vals[3] for vals in lista_valores if len(vals) > 3}
        return (id_val, value_val, owner_val, " | ".join(sorted(mensajes)))

# -------------------------------------------------------------------
# FUNCIONES DE PROCESAMIENTO OPTIMIZADAS CON PANDAS
# -------------------------------------------------------------------

def procesar_fibras_pandas(df_fi: pd.DataFrame, ids_om: set, responsables: dict, odfs_data: dict) -> tuple:
    """Versión optimizada con pandas para procesar fibras."""
    errores = []
    especificaciones = {}
    
    # Renombrar columnas para acceso más fácil
    cols_fi = COLUMNAS_MASIVO["FIBRA"]
    rename_map = {v: k for k, v in cols_fi.items() if v}
    df_fi.rename(columns=rename_map, inplace=True)
    
    # Columnas necesarias
    required = ['id', 'estado', 'nombre']
    for col in required:
        if col not in df_fi.columns:
            return errores, especificaciones, 0
    
    # Filtrar OM
    ids_om_set = set(ids_om)
    df_fi = df_fi[~df_fi['id'].astype(str).isin(ids_om_set)]
    df_fi = df_fi[df_fi['estado'].notna()]
    df_fi = df_fi[df_fi['estado'].astype(str).str.strip() != ""]
    
    if df_fi.empty:
        return errores, especificaciones, 0
    
    filas_procesadas = len(df_fi)
    
    # Normalizar IDs
    df_fi['id'] = df_fi['id'].apply(normalizar_id)
    df_fi['nombre'] = df_fi['nombre'].astype(str).str.strip()
    df_fi['base_raw'] = df_fi['nombre'].str.split('/').str[0].str.strip()
    df_fi['base'] = df_fi['base_raw'].apply(obtener_nombre_base_fibra)
    
    if 'colocacion' in df_fi.columns:
        df_fi['colocacion'] = df_fi['colocacion'].astype(str).str.strip()
    if 'especificacion' in df_fi.columns:
        df_fi['especificacion'] = df_fi['especificacion'].astype(str).str.strip()
    if 'propietario' in df_fi.columns:
        df_fi['propietario'] = df_fi['propietario'].astype(str).str.strip()
    if 'nombre_ant' in df_fi.columns:
        df_fi['nombre_ant'] = df_fi['nombre_ant'].astype(str).str.strip()
    
    # Mapeo de estados
    estado_map = {"is": "En servicio", "pa": "Plan Abandono", "pf": "Plan Uso Futuro"}
    df_fi['estado_traducido'] = df_fi['estado'].astype(str).str.lower().map(estado_map).fillna(df_fi['estado'])
    
    # Procesar fila por fila para errores (no hay vectorización fácil por la lógica compleja)
    for idx, row in df_fi.iterrows():
        idv = row['id']
        nombre = row['nombre']
        base = row['base']
        tipo = row.get('colocacion', '')
        esp_raw = row.get('especificacion', '')
        esp = esp_raw.lower()
        prop = row.get('propietario', '')
        est = row['estado_traducido']
        nombre_ant = row.get('nombre_ant', '')
        
        if base:
            especificaciones[base] = esp
        
        errores_fila = []
        errores_fila.extend(validar_nombre_fibra(nombre, tipo))
        
        if "-" in nombre and tipo != "Ultima Milla" and not nombre.upper().endswith("-SUB"):
            errores_fila.append(f"{JERARQUIA_FIBRA[3]['errores'][6]} (actual: {tipo})")
        
        tipo_normalizado = tipo.lower().replace(" ", "-")
        if "-" not in nombre and tipo_normalizado not in ("troncal-acceso", "troncal-transporte"):
            if tipo_normalizado != "ultima-milla":
                errores_fila.append(f"{JERARQUIA_FIBRA[3]['errores'][5]} (actual: {tipo})")
        
        if "doble" in esp:
            errores_fila.append(f"{JERARQUIA_FIBRA[9]['errores'][0]} (actual: {esp})")
        
        if prop != PROPIETARIO_VALIDO:
            errores_fila.append(f"{JERARQUIA_FIBRA[11]['errores'][0]} (actual: {prop})")
        
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            errores_fila.append(f"{JERARQUIA_FIBRA[10]['errores'][0]} (actual: {est})")
        
        if tipo == "Ultima Milla":
            hilos = extraer_cantidad_hilos(esp_raw)
            if hilos > 48:
                errores_fila.append(f"{JERARQUIA_FIBRA[9]['errores'][2]} (actual: {hilos} Hilos)")
        
        if tipo == "Ultima Milla" and nombre_ant:
            if not validar_nombre_antiguo_fibra(nombre_ant, tipo):
                errores_fila.append(f"{JERARQUIA_FIBRA[4]['errores'][0]} (actual: {nombre_ant})")
            if "-" in nombre and "-" not in str(nombre_ant) and not nombre.upper().endswith("-SUB"):
                errores_fila.append(f"{JERARQUIA_FIBRA[4]['errores'][1]} (NOMBRE: {nombre}, NOMBRE_ANT: {nombre_ant})")
        
        owner = responsables.get(base, "-")
        for msg in errores_fila:
            errores.append((idv, nombre, owner, msg))
    
    return errores, especificaciones, filas_procesadas


def procesar_empalmes_pandas(df_em: pd.DataFrame, ids_om: set, responsables: dict) -> tuple:
    """Versión optimizada con pandas para procesar empalmes."""
    errores = []
    datos_em = []
    
    cols_em = COLUMNAS_MASIVO["EMPALME"]
    rename_map = {v: k for k, v in cols_em.items() if v}
    df_em.rename(columns=rename_map, inplace=True)
    
    required = ['id', 'estado', 'nombre']
    for col in required:
        if col not in df_em.columns:
            return errores, datos_em, 0
    
    ids_om_set = set(ids_om)
    df_em = df_em[~df_em['id'].astype(str).isin(ids_om_set)]
    df_em = df_em[df_em['estado'].notna()]
    df_em = df_em[df_em['estado'].astype(str).str.strip() != ""]
    
    if df_em.empty:
        return errores, datos_em, 0
    
    filas_procesadas = len(df_em)
    
    df_em['id'] = df_em['id'].apply(normalizar_id)
    df_em['nombre'] = df_em['nombre'].astype(str).str.strip()
    df_em['nombre_upper'] = df_em['nombre'].str.upper()
    df_em['corto'] = df_em['nombre_upper'].apply(
        lambda x: x[3:] if x.startswith("NAP") else (x[2:] if len(x) > 2 else x)
    )
    if 'especificacion' in df_em.columns:
        df_em['especificacion'] = df_em['especificacion'].astype(str).str.strip()
    if 'propietario' in df_em.columns:
        df_em['propietario'] = df_em['propietario'].astype(str).str.strip()
    
    estado_map = {"is": "En servicio", "pa": "Plan Abandono", "pf": "Plan Uso Futuro"}
    df_em['estado_traducido'] = df_em['estado'].astype(str).str.lower().map(estado_map).fillna(df_em['estado'])
    
    grupos_especificacion = defaultdict(list)
    nombres_vistos = defaultdict(list)
    
    for idx, row in df_em.iterrows():
        idv = row['id']
        nombre = row['nombre']
        nombre_upper = row['nombre_upper']
        corto = row['corto']
        esp = row.get('especificacion', '')
        prop = row.get('propietario', '')
        est = row['estado_traducido']
        
        grupo = ""
        if (nombre.startswith("ET") or nombre.startswith("NAP")) and "_" in nombre:
            if "-SUB" in nombre_upper:
                match = re.match(r'^(ET|NAP)(\d+-SUB)', nombre_upper)
                grupo = f"{match.group(1)}{match.group(2)}" if match else ""
            else:
                grupo = nombre.split("_")[0]
        
        datos_em.append((idv, nombre, esp, prop, est, grupo, corto))
        nombres_vistos[nombre].append(idv)
        
        if grupo:
            agrupado = "Nap Sotano" if nombre_upper.startswith("NAP") and esp.startswith("Nap Sotano") else esp
            if agrupado in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente", "Nap Sotano"}:
                grupos_especificacion[grupo].append(agrupado)
        
        errores_fila = []
        errores_fila.extend(validar_nombre_empalme(nombre))
        errores_fila.extend(validar_especificacion_empalme(nombre, esp))
        if prop != PROPIETARIO_VALIDO:
            errores_fila.append(f"{JERARQUIA_EMPALME[4]['errores'][0]} (actual: {prop})")
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            errores_fila.append(f"{JERARQUIA_EMPALME[6]['errores'][0]} (actual: {est})")
        if "-SUB" not in nombre_upper:
            if not RE_NOMBRE_EMPALME_VALIDO.match(nombre_upper):
                base_match = RE_NOMBRE_EMPALME_BASE.match(nombre_upper)
                nombre_corregido = base_match.group(0) if base_match else ""
                if nombre_corregido and nombre_corregido != nombre_upper:
                    errores_fila.append(f"{JERARQUIA_EMPALME[3]['errores'][4]} (usar: {nombre_corregido})")
        
        owner = responsables.get(corto, "-")
        for msg in errores_fila:
            errores.append((idv, nombre, owner, msg))
    
    # Duplicados
    duplicados_reportados = set()
    for nombre, ids in nombres_vistos.items():
        if len(ids) > 1 and not ((nombre.startswith("ET") or nombre.startswith("NAP")) and "_" in nombre):
            ids_str = " , ".join(ids)
            corto_dup = nombre[3:] if nombre.upper().startswith("NAP") else (nombre[2:] if len(nombre) > 2 else nombre)
            owner = responsables.get(corto_dup, "-")
            err_tuple = (ids_str, nombre, owner, f"{JERARQUIA_EMPALME[1]['errores'][1]})")
            errores.append(err_tuple)
            duplicados_reportados.add(nombre)
    
    # Mayoría por grupo
    mayor_por_grupo = {g: Counter(v).most_common(1)[0][0] for g, v in grupos_especificacion.items() if v}
    for (idv, nombre, esp, prop, est, grupo, corto) in datos_em:
        if nombre in duplicados_reportados or not grupo or not mayor_por_grupo.get(grupo):
            continue
        esp_mayor = mayor_por_grupo[grupo]
        if grupo.startswith("NAP"):
            es_nap_actual = esp.startswith("Nap Sotano")
            es_nap_mayor = esp_mayor == "Nap Sotano"
            if es_nap_actual != es_nap_mayor:
                owner = responsables.get(corto, "-")
                msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (para {grupo}, esperado: '{esp_mayor}', actual: '{esp}')"
                errores.append((idv, nombre, owner, msg))
        elif esp in {"Tipo 0", "Tipo 1", "Tipo 2", "Empalme Cliente"} and esp != esp_mayor:
            owner = responsables.get(corto, "-")
            msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (esperado: '{esp_mayor}', actual: '{esp}')"
            errores.append((idv, nombre, owner, msg))
    
    return errores, datos_em, filas_procesadas


def procesar_edificios_estructuras_pandas(df_ed: pd.DataFrame, ids_om: set, exc_ed: set, responsables_edificio: dict) -> tuple:
    """Versión optimizada para edificios y estructuras."""
    errores_edificio = []
    errores_estructura = []
    
    cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
    rename_map = {v: k for k, v in cols_ed.items() if v}
    df_ed.rename(columns=rename_map, inplace=True)
    
    required = ['id', 'nombre']
    if 'id' not in df_ed.columns or 'nombre' not in df_ed.columns:
        return errores_edificio, errores_estructura, 0
    
    ids_om_set = set(ids_om)
    df_ed = df_ed[~df_ed['id'].astype(str).isin(ids_om_set)]
    
    if df_ed.empty:
        return errores_edificio, errores_estructura, 0
    
    filas_procesadas = len(df_ed)
    
    df_ed['id'] = df_ed['id'].apply(normalizar_id)
    df_ed['nombre'] = df_ed['nombre'].astype(str).str.strip()
    df_ed['base'] = df_ed['nombre'].str.split('/').str[0].str.strip()
    if 'tipo' in df_ed.columns:
        df_ed['tipo'] = df_ed['tipo'].astype(str).str.strip()
    if 'categoria' in df_ed.columns:
        df_ed['categoria'] = df_ed['categoria'].astype(str).str.strip().str.upper()
    
    # Agrupar tipos por base
    grupos_tipo = defaultdict(list)
    for idx, row in df_ed.iterrows():
        base = row.get('base', '')
        tipo = row.get('tipo', '')
        if base:
            grupos_tipo[base].append(tipo)
    
    tipo_mayoria = {g: Counter(t).most_common(1)[0][0] for g, t in grupos_tipo.items() if t}
    
    # Procesar cada fila
    for idx, row in df_ed.iterrows():
        idv = row['id']
        nombre = row['nombre']
        base = row.get('base', '')
        tipo = row.get('tipo', '')
        cat = row.get('categoria', '')
        
        # Validar sufijos
        valido_sufijos, msg_sufijos, _ = validar_sufijos_nombre_empresa(nombre)
        if not valido_sufijos:
            owner = responsables_edificio.get(idv, "RESPONSABLE NO ENCONTRADO")
            errores_edificio.append((idv, nombre, owner, msg_sufijos))
        
        msgs = []
        if cat != "AAA":
            msgs.append(f"{JERARQUIA_EDIFICIO[5]['errores'][0]} (actual: {cat})")
        if base and tipo != tipo_mayoria.get(base, tipo) and nombre not in exc_ed:
            msgs.append(f"{JERARQUIA_EDIFICIO[3]['errores'][0]} (esperado: '{tipo_mayoria[base]}', actual: '{tipo}')")
        nombre_limpio = nombre.strip()
        if not (nombre_limpio.upper().startswith("C.C") or nombre_limpio.upper().startswith("NODO")):
            if not re.fullmatch(r"[A-Za-z0-9 ñÑ()&.,\-]+", nombre_limpio):
                invalidos = "".join(sorted(set(re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nombre_limpio))))
                if invalidos:
                    msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][2]}: '{invalidos}'")
            if nombre_limpio != nombre_limpio.upper():
                msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][1]} (usar: {nombre_limpio.upper()})")
        if msgs:
            errores_edificio.append((idv, nombre, "...", " | ".join(msgs)))
    
    # Estructuras - se necesitan columnas específicas
    col_est = COLUMNAS_ESTRUCTURA
    # Buscar columnas por nombre
    id_est_col = None
    nombre_est_col = None
    floor_len_col = None
    floor_wid_col = None
    rooms_len_col = None
    rooms_wid_col = None
    floors_col = None
    
    for col in df_ed.columns:
        col_lower = col.lower()
        if col_lower == col_est["idtexto"].lower():
            id_est_col = col
        elif col_lower == col_est["nombre"].lower():
            nombre_est_col = col
        elif col_lower == col_est["floor_dimension_length"].lower():
            floor_len_col = col
        elif col_lower == col_est["floor_dimension_width"].lower():
            floor_wid_col = col
        elif col_lower == col_est["number_of_rooms_length"].lower():
            rooms_len_col = col
        elif col_lower == col_est["number_of_rooms_width"].lower():
            rooms_wid_col = col
        elif col_lower == col_est["number_of_floors"].lower():
            floors_col = col
    
    if id_est_col and nombre_est_col:
        df_est = df_ed[[id_est_col, nombre_est_col]].copy()
        df_est.columns = ['id', 'nombre']
        df_est['id'] = df_est['id'].apply(normalizar_id)
        df_est['nombre'] = df_est['nombre'].astype(str).str.strip()
        df_est = df_est[~df_est['id'].isin(ids_om_set)]
        df_est = df_est[~df_est['nombre'].str.upper().str.startswith("NODO")]
        
        # Añadir columnas de estructura si existen
        if floor_len_col:
            df_est['floor_len'] = df_ed[floor_len_col].apply(lambda x: normalizar_id(x) if pd.notna(x) else "")
        else:
            df_est['floor_len'] = ""
        if floor_wid_col:
            df_est['floor_wid'] = df_ed[floor_wid_col].apply(lambda x: normalizar_id(x) if pd.notna(x) else "")
        else:
            df_est['floor_wid'] = ""
        if rooms_len_col:
            df_est['rooms_len'] = df_ed[rooms_len_col].apply(lambda x: normalizar_id(x) if pd.notna(x) else "")
        else:
            df_est['rooms_len'] = ""
        if rooms_wid_col:
            df_est['rooms_wid'] = df_ed[rooms_wid_col].apply(lambda x: normalizar_id(x) if pd.notna(x) else "")
        else:
            df_est['rooms_wid'] = ""
        if floors_col:
            df_est['num_pisos'] = df_ed[floors_col].apply(lambda x: normalizar_id(x) if pd.notna(x) else "")
        else:
            df_est['num_pisos'] = ""
        
        for idx, row in df_est.iterrows():
            idv = row['id']
            nombre_est = row['nombre']
            floor_len = row['floor_len']
            floor_wid = row['floor_wid']
            rooms_len = row['rooms_len']
            rooms_wid = row['rooms_wid']
            num_pisos = row['num_pisos']
            owner = responsables_edificio.get(idv, "RESPONSABLE NO ENCONTRADO")
            
            if num_pisos == "0" and floor_len == "0" and floor_wid == "0" and rooms_len == "0" and rooms_wid == "0":
                detalles = f"NUM_OF:{num_pisos}|FLOOR:{floor_len},{floor_wid}|ROOMS:{rooms_len},{rooms_wid}"
                err_msg = f"{JERARQUIA_ESTRUCTURA[1]['errores'][0]} ({detalles})"
                errores_estructura.append((idv, nombre_est, owner, err_msg))
                continue
            
            FLOOR_VALIDOS = {"700", "40000", "10000", "20000"}
            ROOMS_VALIDOS = {"3", "0"}
            regs = []
            if floor_len and floor_len not in FLOOR_VALIDOS:
                regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_len})")
            if floor_wid and floor_wid not in FLOOR_VALIDOS:
                regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_wid})")
            if rooms_len and rooms_len not in ROOMS_VALIDOS:
                regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {rooms_len})")
            if rooms_wid and rooms_wid not in ROOMS_VALIDOS:
                regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {rooms_wid})")
            if regs:
                errores_estructura.append((idv, nombre_est, owner, "; ".join(regs)))
    
    return errores_edificio, errores_estructura, filas_procesadas


def procesar_ports_pandas(df_ports: pd.DataFrame, ids_om: set, nombres_edificios: dict, nombres_normalizados: dict) -> tuple:
    """Versión optimizada para ports, shelf y bay."""
    errores_port = []
    errores_shelf = []
    errores_bay = []
    
    ids_om_set = set(ids_om)
    
    # Renombrar columnas para facilitar acceso
    cols_port = COLUMNAS_MASIVO["PORT"]
    cols_shelf = COLUMNAS_MASIVO["SHELF"]
    cols_bay = COLUMNAS_MASIVO["BAY"]
    
    rename_port = {v: f"port_{k}" for k, v in cols_port.items() if v and v in df_ports.columns}
    rename_shelf = {v: f"shelf_{k}" for k, v in cols_shelf.items() if v and v in df_ports.columns}
    rename_bay = {v: f"bay_{k}" for k, v in cols_bay.items() if v and v in df_ports.columns}
    
    rename_all = {**rename_port, **rename_shelf, **rename_bay}
    df_ports.rename(columns=rename_all, inplace=True)
    
    if 'port_id' not in df_ports.columns:
        return errores_port, errores_shelf, errores_bay, 0
    
    df_ports['port_id'] = df_ports['port_id'].apply(normalizar_id)
    df_ports = df_ports[~df_ports['port_id'].isin(ids_om_set)]
    
    if df_ports.empty:
        return errores_port, errores_shelf, errores_bay, 0
    
    filas_procesadas = len(df_ports)
    
    # Procesar cada fila (la lógica es compleja, no vectorizable fácilmente)
    for idx, row in df_ports.iterrows():
        idv = row.get('port_id')
        if not idv:
            continue
        
        # Leer campos del shelf
        desc_shelf_raw = row.get('shelf_description', '')
        tipo_shelf_raw = row.get('shelf_tipo', '')
        type_raw = row.get('port_type', '')
        shelf_id = row.get('shelf_id', '')
        
        # Detectar central
        es_central = False
        tipo_central = None
        if desc_shelf_raw:
            desc_upper = str(desc_shelf_raw).upper()
            if "OT" in desc_upper:
                es_central = True
                tipo_central = "OT"
            elif "OI" in desc_upper:
                es_central = True
                tipo_central = "OI"
            elif "PP" in desc_upper:
                es_central = True
                tipo_central = "PP"
        
        if es_central:
            # Validar Type del PORT
            type_val = str(type_raw).strip() if type_raw is not None else ""
            if type_val.upper() != "ODF":
                msg = f"TYPE INVALIDO. Debe ser 'ODF' para central (actual: '{type_val}')"
                errores_port.append((idv, str(row.get('port_description', '')) or "Sin descripción", "", msg))
            # Validar Tipo (2) del SHELF
            esperado = None
            if tipo_central == "OT":
                esperado = "ODF Troncal"
            elif tipo_central == "OI":
                esperado = "ODF Interconexión"
            elif tipo_central == "PP":
                esperado = "ODF Pacheo"
            if esperado:
                if str(tipo_shelf_raw).strip().upper() != esperado.upper():
                    msg = f"TIPO INVALIDO. Esperado: '{esperado}', actual: '{tipo_shelf_raw}'"
                    shelf_id_val = normalizar_id(shelf_id) if shelf_id else idv
                    errores_shelf.append((shelf_id_val, f"SHELF: {desc_shelf_raw or 'Sin descripción'}", "", msg))
            continue  # Saltar el resto de validaciones
        
        # Si no es central, aplicar validaciones normales
        name_building = row.get('name_building', '')
        building_id = row.get('building_id', '')
        tipo_senal = row.get('port_tipo_senal', '')
        description = row.get('port_description', '')
        ps_raw = row.get('port_physical_status')
        if ps_raw is not None:
            ps_upper = str(ps_raw).strip().upper()
            ps = "IN SERVICE" if ps_upper == "IS" else ps_raw
        else:
            ps = None
        nombre_cliente_raw = row.get('port_nombre_cliente')
        comentarios_raw = row.get('port_comentarios')
        direccion_cliente_raw = row.get('port_direccion_cliente_id')
        proyecto_red_raw = row.get('port_proyecto_red')
        anillo_raw = row.get('port_anillo')
        cambio_sm_raw = row.get('port_cambio_service_manager')
        id_servicio_raw = row.get('port_id_servicio')
        unidad_ancho_banda_raw = row.get('port_unidad_ancho_banda')
        connector_type_raw = row.get('port_connector_type')
        service_level_raw = row.get('port_service_level')
        equipo_central_raw = row.get('port_equipo_central')
        puerto_central_raw = row.get('port_puerto_central')
        
        mensajes = []
        error_equipo_central = None
        error_puerto_central = None
        
        if pd.notna(equipo_central_raw):
            estado_ec, valor_ec = validar_equipo_central(equipo_central_raw)
            if estado_ec != "VÁLIDO":
                valor_mostrar = str(equipo_central_raw).strip() if equipo_central_raw is not None else ""
                error_equipo_central = f"EQUIPO CENTRAL ({valor_mostrar})"
        
        if pd.notna(puerto_central_raw):
            estado_pc, valor_pc = validar_puerto_central(puerto_central_raw)
            if estado_pc != "VÁLIDO":
                valor_mostrar = str(puerto_central_raw).strip() if puerto_central_raw is not None else ""
                error_puerto_central = f"PUERTO CENTRAL ({valor_mostrar})"
        
        if pd.notna(nombre_cliente_raw):
            es_valido, valor_actual, error_msg, _ = validar_nombre_cliente_contra_edificio(
                nombre_cliente_raw, building_id, ps, comentarios_raw,
                nombres_edificios, nombres_normalizados
            )
            if not es_valido and error_msg is not None:
                errores_port.append((idv, valor_actual or "", "", error_msg))
                continue
        
        if ps is not None and str(ps).strip().upper() == "IN SERVICE":
            dir_str = str(direccion_cliente_raw).strip() if direccion_cliente_raw is not None else ""
            if es_valor_cero(direccion_cliente_raw, dir_str):
                mensajes.append(f"{JERARQUIA_PORT[17]['errores'][0]} = 0")
            proy_str = str(proyecto_red_raw).strip() if proyecto_red_raw is not None else ""
            if es_valor_cero(proyecto_red_raw, proy_str):
                mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]} = 0")
            elif proyecto_red_raw is not None and proy_str:
                proy_clean = re.sub(r'[\x00-\x1f\x7f]', '', proy_str).strip()
                if not re.match(r'^[A-Z0-9]+$', proy_clean):
                    mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]}'{proy_str}'")
                elif not any(proy_clean.startswith(p) for p in ("ME", "HU", "SD")):
                    mensajes.append(f"{JERARQUIA_PORT[10]['errores'][0]}: '{proy_str}'")
            if anillo_raw is not None and str(anillo_raw).strip() == "":
                mensajes.append(f"{JERARQUIA_PORT[20]['errores'][0]}")
            if cambio_sm_raw is not None:
                csm_str = str(cambio_sm_raw).strip()
                csm_upper = csm_str.upper()
                if csm_str:
                    if len(csm_upper) < 7:
                        mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
                    elif not any(csm_upper.startswith(p) for p in PREFIJOS_SERVICE_MANAGER):
                        mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
                    elif not csm_upper.replace("-", "").isalnum():
                        mensajes.append(f"{JERARQUIA_PORT[25]['errores'][0]}: '{csm_str}'")
            if id_servicio_raw is not None:
                id_serv_str = str(id_servicio_raw).strip()
                id_serv_clean = re.sub(r'[\x00-\x1f\x7f]', '', id_serv_str).strip()
                if id_serv_str and id_serv_str != "0":
                    if re.search(r'[\s\t]+', id_serv_clean):
                        mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                    elif not re.match(r'^[A-Z0-9]+$', id_serv_clean):
                        mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                    elif len(id_serv_clean) < 8:
                        mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
                    elif not id_serv_clean.startswith(PREFIJOS_ID_SERVICIO):
                        mensajes.append(f"{JERARQUIA_PORT[26]['errores'][0]}: '{id_serv_str}'")
        
        if error_equipo_central:
            mensajes.append(error_equipo_central)
        if error_puerto_central:
            mensajes.append(error_puerto_central)
        
        if tipo_senal != "Rx - Tx":
            mensajes.append(f"{JERARQUIA_PORT[2]['errores'][0]} ({'(VACÍO)' if not tipo_senal else f'(actual: {tipo_senal!r})'})")
        
        st, _ = validar_type_equipo_cliente(type_raw)
        if st != "VÁLIDO":
            if st == 'VACÍO':
                mensajes.append(f"{JERARQUIA_PORT[5]['errores'][1]}")
            else:
                mensajes.append(f"{JERARQUIA_PORT[5]['errores'][2]} (actual: '{type_raw}')")
        
        stc, _ = validar_connector_type(connector_type_raw)
        if stc != "VÁLIDO":
            if stc == 'VACÍO':
                mensajes.append(f"{JERARQUIA_PORT[6]['errores'][1]}")
            else:
                mensajes.append(f"{JERARQUIA_PORT[6]['errores'][2]} (actual: '{connector_type_raw}')")
        
        sts, _ = validar_service_level_fiber(service_level_raw)
        if sts != "VÁLIDO":
            if sts == 'VACÍO':
                mensajes.append(f"{JERARQUIA_PORT[9]['errores'][1]}")
            else:
                mensajes.append(f"{JERARQUIA_PORT[9]['errores'][2]} (actual: '{service_level_raw}')")
        
        if ps is not None:
            stp, _ = validar_physical_status(ps)
            if stp != "VÁLIDO":
                if stp == 'VACÍO':
                    mensajes.append(f"{JERARQUIA_PORT[4]['errores'][1]}")
                else:
                    mensajes.append(f"{JERARQUIA_PORT[4]['errores'][0]} (actual: '{ps}')")
        
        stu, _ = validar_unidad_ancho_banda(unidad_ancho_banda_raw)
        if stu != "VÁLIDO":
            if stu == 'VACÍO':
                mensajes.append(f"{JERARQUIA_PORT[8]['errores'][1]}")
            else:
                mensajes.append(f"{JERARQUIA_PORT[8]['errores'][0]} (actual: '{unidad_ancho_banda_raw}')")
        
        if mensajes:
            msg_final = " | ".join(mensajes)
            errores_port.append((idv, description or "Sin descripción", "", msg_final))
        
        # Validaciones de SHELF (si no es central)
        if 'shelf_id' in df_ports.columns:
            shelf_idv = normalizar_id(row.get('shelf_id'))
            if shelf_idv and shelf_idv not in ids_om_set:
                # Verificar si es central por description shelf (ya lo tenemos)
                if not es_central:
                    description_shelf = str(row.get('shelf_description', ''))
                    estado_shelf = str(row.get('shelf_estado', ''))
                    contratista_shelf = str(row.get('shelf_contratista', ''))
                    tecnologia = str(row.get('shelf_tecnologia', ''))
                    direccion_ip = str(row.get('shelf_direccion_ip', ''))
                    nivel_bastidor = str(row.get('shelf_nivel_bastidor', ''))
                    migracion_olt = str(row.get('shelf_migracion_olt', ''))
                    id_bay = normalizar_id(row.get('shelf_id_bay')) if 'shelf_id_bay' in row else ""
                    
                    msgs_shelf = []
                    if not description_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[2]['errores'][0]}")
                    else:
                        if description_shelf != description_shelf.upper():
                            msgs_shelf.append("DESCRIPTION CON ESPACIOS")
                        if description_shelf.endswith(' '):
                            msgs_shelf.append("DESCRIPTION CON ESPACIOS")
                        msgs_shelf.extend(validar_descripcion_demarcador(description_shelf))
                    if not contratista_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[9]['errores'][0]}")
                    if not estado_shelf:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]}")
                    elif estado_shelf.lower() not in ("is", "spare", "under construction", "retired"):
                        msgs_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]} (actual: '{estado_shelf}')")
                    if not tecnologia:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[6]['errores'][0]}")
                    if direccion_ip:
                        est_ip, _ = validar_ipv4(direccion_ip)
                        if est_ip != "VÁLIDO":
                            msgs_shelf.append(f"{JERARQUIA_SHELF[7]['errores'][0]} (actual: '{direccion_ip}')")
                    if nivel_bastidor:
                        try:
                            nivel = int(nivel_bastidor)
                            if not (0 <= nivel <= 50):
                                msgs_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                        except ValueError:
                            msgs_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                    if migracion_olt and migracion_olt.lower() not in ("yes", "no", "sí", "true", "false"):
                        msgs_shelf.append(f"{JERARQUIA_SHELF[12]['errores'][0]} (actual: '{migracion_olt}')")
                    if not id_bay:
                        msgs_shelf.append(f"{JERARQUIA_SHELF[13]['errores'][0]}")
                    if msgs_shelf:
                        unicos = list(dict.fromkeys(msgs_shelf))
                        error_msg = " | ".join(unicos)
                        errores_shelf.append((shelf_idv, f"SHELF: {description_shelf or 'Sin descripción'}", "", error_msg))
        
        # Validaciones de BAY (similar)
        if 'bay_id' in df_ports.columns:
            bay_idv = normalizar_id(row.get('bay_id'))
            if bay_idv and bay_idv not in ids_om_set:
                if not es_central:
                    description_bay = str(row.get('bay_description', ''))
                    number_bay = str(row.get('bay_number', ''))
                    estado_bay = str(row.get('bay_estado', ''))
                    numero_salon = str(row.get('bay_numero_salon', ''))
                    numero_fila = str(row.get('bay_numero_fila', ''))
                    contratista_bay = str(row.get('bay_contratista', ''))
                    id_edificio = normalizar_id(row.get('bay_id_edificio')) if 'bay_id_edificio' in row else ""
                    
                    msgs_bay = []
                    if not description_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[3]['errores'][0]}")
                    msgs_bay.extend(validar_descripcion_bastidor(description_bay))
                    if not number_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[2]['errores'][0]}")
                    if not estado_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[4]['errores'][0]}")
                    if not numero_salon:
                        msgs_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]}")
                    elif not RE_NUMERO_SALON.match(numero_salon):
                        msgs_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]} (actual: '{numero_salon}')")
                    if not numero_fila:
                        msgs_bay.append(f"{JERARQUIA_BAY[6]['errores'][0]}")
                    elif numero_fila != "1":
                        msgs_bay.append(f"NUMERO FILA INVALIDO. ACTUAL: {numero_fila}")
                    if not contratista_bay:
                        msgs_bay.append(f"{JERARQUIA_BAY[7]['errores'][0]}")
                    if not id_edificio:
                        msgs_bay.append(f"{JERARQUIA_BAY[9]['errores'][0]}")
                    if msgs_bay:
                        unicos = list(dict.fromkeys(msgs_bay))
                        error_msg = " | ".join(unicos)
                        errores_bay.append((bay_idv, f"BAY: {description_bay or 'Sin descripción'}", "", error_msg))
    
    return errores_port, errores_shelf, errores_bay, filas_procesadas


def detectar_edificios_con_nombres_similares_pandas(ruta_archivo: Path) -> list:
    """Versión optimizada con pandas para nombres similares."""
    try:
        df = pd.read_excel(ruta_archivo, dtype=str)
    except Exception:
        return []
    
    # Buscar columnas de nombre e id
    nombre_col = None
    id_col = None
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ("nombre", "name"):
            nombre_col = col
        if col_lower in ("id", "idtexto", "id texto"):
            id_col = col
    
    if nombre_col is None or id_col is None:
        return []
    
    df = df[[id_col, nombre_col]].dropna(subset=[id_col, nombre_col])
    df[id_col] = df[id_col].apply(normalizar_id)
    df[nombre_col] = df[nombre_col].astype(str).str.strip()
    
    conteo = df[nombre_col].value_counts().to_dict()
    ids_por_nombre = df.groupby(nombre_col)[id_col].apply(lambda x: ", ".join(x.astype(str))).to_dict()
    nombres_vistos = df[nombre_col].unique().tolist()
    
    def preparar_comparacion(texto: str) -> str:
        t = texto.lower()
        t = re.sub(r'[.,;()]', '', t)
        t = re.sub(r'\s+', ' ', t).strip()
        palabras = t.split()
        resultado = []
        i = 0
        while i < len(palabras):
            if len(palabras[i]) == 1 and palabras[i].isalpha():
                letras = []
                while i < len(palabras) and len(palabras[i]) == 1 and palabras[i].isalpha():
                    letras.append(palabras[i])
                    i += 1
                resultado.append(''.join(letras) if len(letras) > 1 else letras[0])
            else:
                resultado.append(palabras[i])
                i += 1
        return ' '.join(resultado)
    
    grupos_normalizados = defaultdict(list)
    for nombre in nombres_vistos:
        prep = preparar_comparacion(nombre)
        if prep:
            grupos_normalizados[prep].append(nombre)
    
    def excluir(nombre: str) -> bool:
        nl = nombre.lower()
        return "subred" in nl or "sede" in nl or nl.startswith("ant cm")
    
    def puntuar(nombre: str) -> int:
        pts = 0
        if '.' not in nombre: pts += 1
        if '  ' not in nombre and nombre.count(' ') < 3: pts += 1
        if nombre.strip() == nombre: pts += 1
        if not any(nombre.upper().startswith(p) for p in ("C.C", "NODO")) and nombre == nombre.upper():
            pts += 1
        return pts
    
    salida = []
    for prep, grupo in grupos_normalizados.items():
        grupo_filtrado = [n for n in grupo if not excluir(n)]
        if len(grupo_filtrado) < 2:
            continue
        max_count = max(conteo[n] for n in grupo_filtrado)
        mejores = [n for n in grupo_filtrado if conteo[n] == max_count]
        mejor = max(mejores, key=puntuar)
        for nombre in grupo_filtrado:
            if nombre == mejor:
                continue
            ids = ids_por_nombre.get(nombre, "")
            salida.append((ids, nombre, "...", f"{JERARQUIA_EDIFICIO[2]['errores'][0]} ({mejor})"))
    
    return salida


def ejecutar_validacion_calidad_paralelo(modo: str) -> int:
    start_time = time.time()
    
    # Cargar datos comunes primero
    ids_om, exc_ed = cargar_om_desde_macro(RUTA_MACRO)
    responsables = cargar_responsables_por_cable(RUTA_RESPONSABLES, "TD ACCES0s", "cable acceso")
    responsables_edificio = cargar_responsables_por_cable(RUTA_RESPONSABLES, "TD ACCES0s", "id edifico")
    odfs_data = cargar_validaciones_odfs(RUTA_ODFS)
    nombres_edificios, nombres_normalizados = cargar_nombres_edificios(RUTA_EDIFICIOS_BULK, modo)
    
    # Leer DataFrames
    print("Cargando archivos con pandas...")
    df_fibra = pd.read_excel(RUTA_FIBRAS_BULK, dtype=str) if RUTA_FIBRAS_BULK.exists() else pd.DataFrame()
    df_empalme = pd.read_excel(RUTA_EMPALMES_BULK, dtype=str) if RUTA_EMPALMES_BULK.exists() else pd.DataFrame()
    df_edificio = pd.read_excel(RUTA_EDIFICIOS_BULK, dtype=str) if RUTA_EDIFICIOS_BULK.exists() else pd.DataFrame()
    df_ports = pd.read_excel(RUTA_PORTS_BULK, dtype=str) if RUTA_PORTS_BULK.exists() else pd.DataFrame()
    
    # Procesar en paralelo
    with ProcessPoolExecutor(max_workers=4) as executor:
        futures = {}
        
        if not df_fibra.empty:
            futures['fibra'] = executor.submit(procesar_fibras_pandas, df_fibra, ids_om, responsables, odfs_data)
        if not df_empalme.empty:
            futures['empalme'] = executor.submit(procesar_empalmes_pandas, df_empalme, ids_om, responsables)
        if not df_edificio.empty:
            futures['edificio'] = executor.submit(procesar_edificios_estructuras_pandas, df_edificio, ids_om, exc_ed, responsables_edificio)
        if not df_ports.empty:
            futures['ports'] = executor.submit(procesar_ports_pandas, df_ports, ids_om, nombres_edificios, nombres_normalizados)
        
        # Recoger resultados
        resultados = {}
        for key, future in futures.items():
            resultados[key] = future.result()
    
    # Procesar resultados
    organizador = OrganizadorErrores()
    contadores = {'fibra': 0, 'empalme': 0, 'edificio': 0, 'estructura': 0, 'ports': 0}
    errores_por_tipo = {'FIBRA': 0, 'EMPALME': 0, 'EDIFICIO': 0, 'ESTRUCTURA': 0, 'PORTS': 0, 'MERGE': 0}
    
    # Fibras
    if 'fibra' in resultados:
        err_fi, espec_fi, filas_fi = resultados['fibra']
        contadores['fibra'] = filas_fi
        errores_por_tipo['FIBRA'] = len(err_fi)
        for error in err_fi:
            organizador.agregar_error(error[0], "FIBRA", error)
    else:
        espec_fi = {}
    
    # Empalmes
    if 'empalme' in resultados:
        err_em, datos_em, filas_em = resultados['empalme']
        contadores['empalme'] = filas_em
        errores_por_tipo['EMPALME'] = len(err_em)
        for error in err_em:
            organizador.agregar_error(error[0], "EMPALME", error)
    else:
        datos_em = []
    
    # Edificios
    if 'edificio' in resultados:
        err_ed, err_estructura, filas_ed = resultados['edificio']
        contadores['edificio'] = filas_ed
        contadores['estructura'] = filas_ed
        errores_por_tipo['EDIFICIO'] = len(err_ed)
        errores_por_tipo['ESTRUCTURA'] = len(err_estructura)
        for error in err_ed:
            organizador.agregar_error(error[0], "EDIFICIO", error)
        for error in err_estructura:
            organizador.agregar_error(error[0], "ESTRUCTURA", error)
    
    # Ports
    if 'ports' in resultados:
        err_ports, err_shelf, err_bay, filas_ports = resultados['ports']
        contadores['ports'] = filas_ports
        total_ports_errors = len(err_ports) + len(err_shelf) + len(err_bay)
        errores_por_tipo['PORTS'] = total_ports_errors
        for error in err_ports:
            organizador.agregar_error(error[0], "PORT", error)
        for error in err_shelf:
            organizador.agregar_error(error[0], "SHELF", error)
        for error in err_bay:
            organizador.agregar_error(error[0], "BAY", error)
    
    # Nombres similares en edificios (no paralelizable fácilmente, rápido)
    err_similares = detectar_edificios_con_nombres_similares_pandas(RUTA_EDIFICIOS_BULK)
    errores_por_tipo['EDIFICIO'] += len(err_similares)
    for error in err_similares:
        organizador.agregar_error(error[0], "EDIFICIO", error)
    
    # Validaciones adicionales de empalmes (consecutividad y consistencia)
    if datos_em:
        err_consec = validar_consecutividad_empalmes(datos_em, ids_om, responsables)
        err_tipo_cable = validar_consistencia_tipo_cable_empalmes(datos_em, ids_om, responsables)
        errores_por_tipo['EMPALME'] += len(err_consec) + len(err_tipo_cable)
        for error in err_consec:
            organizador.agregar_error(error[0], "EMPALME", error)
        for error in err_tipo_cable:
            organizador.agregar_error(error[0], "EMPALME", error)
    
    # Merge (no paralelizable, necesita datos de empalmes y especificaciones)
    # Se ejecuta al final después de tener todo
    if datos_em and espec_fi and contadores['empalme'] > 0 and contadores['fibra'] > 0:
        err_merge = validar_merge_fibra_empalme(datos_em, espec_fi, responsables, ids_om, None, None, None, None, None)
        errores_por_tipo['MERGE'] = len(err_merge)
        for error in err_merge:
            organizador.agregar_error(error[0], "MERGE FI-EM", error)
    
    # Guardar resultados en la macro
    try:
        wb = openpyxl.load_workbook(RUTA_MACRO, keep_vba=True)
    except Exception as e:
        print(f"ERROR: No se pudo abrir el archivo macro: {e}")
        return 0
    
    if "CD" in wb.sheetnames:
        wb.remove(wb["CD"])
    ws_salida = wb.create_sheet("CD")
    ws_salida.append(["ELEMENT", "ID", "VALUE", "OWNER", "ERROR", "COMMENTS"])
    
    errores_finales = organizador.obtener_errores_organizados()
    for elemento, valores in errores_finales:
        if isinstance(valores, (list, tuple)):
            v = list(valores) + [""] * max(0, 4 - len(valores))
            ws_salida.append([elemento, v[0], v[1], v[2], v[3], "..."])
        else:
            ws_salida.append([elemento, str(valores), "", "", "Formato de error no reconocido", ""])
    
    aplicar_formato_hoja_resultados(ws_salida, len(errores_finales))
    
    total_procesados = sum(contadores.values())
    total_errores_encontrados = sum(errores_por_tipo.values())
    
    print(f"\n═" * 55)
    print(f" TOTAL PROCESADOS........: {total_procesados:,}")
    print(f" ERRORES ENCONTRADOS.....: {total_errores_encontrados:,}")
    print(f" TIEMPO DE EJECUCIÓN.....: {time.time() - start_time:.2f} s")
    
    try:
        wb.save(str(RUTA_MACRO))
        print(" GS_EXPORT guardado correctamente")
    except Exception as e:
        print(f" ERROR al guardar: {e}")
    
    wb.close()
    return total_errores_encontrados


def aplicar_formato_hoja_resultados(ws, total_filas: int):
    borde_fino = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin'),
    )
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = borde_fino
    limite_filas = min(total_filas + 1, ws.max_row)
    for row_idx in range(2, limite_filas + 1):
        for col_idx in range(1, 7):
            ws.cell(row=row_idx, column=col_idx).border = borde_fino
        if row_idx % 10000 == 0:
            gc.collect()
    for col_idx, col in enumerate(ws.columns, 1):
        max_len = max((len(str(cell.value)) for cell in list(col)[:500] if cell.value), default=0)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 50)


if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    print("═" * 55)
    print()
    print(" RUTAS / COLUMNAS")
    archivos = [
        ("FIBRA",     RUTA_FIBRAS_BULK),
        ("EMPALME",   RUTA_EMPALMES_BULK),
        ("EDIFICIO",  RUTA_EDIFICIOS_BULK),
        ("ESTRUCTURA",RUTA_EDIFICIOS_BULK),
        ("PORTS",     RUTA_PORTS_BULK),
    ]
    hay_problemas = False
    for i, (nombre, ruta) in enumerate(archivos):
        status = "(OK)" if ruta.exists() else "(NO ENCONTRADO)"
        if not ruta.exists():
            hay_problemas = True
        prefijo = "└──" if i == len(archivos) - 1 else "├──"
        print(f"{prefijo} {nombre:12}: {status}")
    if not hay_problemas:
        total_errores = ejecutar_validacion_calidad_paralelo("masivo")
        print(f"\n└──GS ERRORES......:[{total_errores}]")
    else:
        print("\nERROR: Se detectaron problemas en los archivos.")