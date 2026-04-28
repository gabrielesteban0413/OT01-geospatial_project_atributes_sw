import openpyxl
import tkinter as tk
from tkinter import simpledialog
from collections import defaultdict, Counter
import re
from pathlib import Path
import pandas as pd
import warnings
import time
from datetime import datetime
from openpyxl.styles import Border, Side, Font
import gc
import os
from dotenv import load_dotenv

load_dotenv()

BASE_PATH_STR = os.getenv("GS_BASE_PATH")
if not BASE_PATH_STR:
    raise ValueError("Variable GS_BASE_PATH no encontrada en .env")

BASE_PATH = Path(BASE_PATH_STR)

RUTA_MACRO = BASE_PATH / "04-AuditoriaCalidadVisualAtributiva.xlsm"
RUTA_RESPONS = BASE_PATH / "03-LINEA CABLES PRIORIZADOS.xlsx"
RUTA_ODFS = BASE_PATH / "04-Sync_SW_SP.xlsx"
RUTA_ASPHIA_CLIENTES = BASE_PATH / "03-ASPHIA_CLIENTES.xlsx"

RUTA_ED_MS = BASE_PATH / "01-Bulk export of buildings.xlsx"
RUTA_FI_MASIVO = BASE_PATH / "01-Bulk export of fibers.xlsx"
RUTA_EM_MASIVO = BASE_PATH / "01-Bulk export of splices.xlsx"
RUTA_PORTS_MASIVO = BASE_PATH / "01-Bulk export of ports.xlsx"

RUTA_EDIFICIO = BASE_PATH / "02-Individual export of buildings.xlsx"
RUTA_FI_DIARIO = BASE_PATH / "02-Individual export of fibers.xlsx"
RUTA_EM_DIARIO = BASE_PATH / "02-Individual export of splices.xlsx"
RUTA_PORTS_DIARIO = BASE_PATH / "02-Individual export of ports.xlsx"



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
        "municipio_b": "municipio b id"
    },
    "EMPALME": {
        "id": "id texto",
        "nombre": "name",
        "especificacion": "nombre especificación",
        "propietario": "propietario",
        "estado": "construction status"
    },
    "EDIFICIO": {
        "id": "idtexto",
        "nombre": "name",
        "tipo": "type",
        "categoria": "categoria"
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
        "id_shelf": "id shelf"
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
        "id_bay": "bay.id"
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
        "cambio_service_manager": "cambio service manager"
    }
}

COLUMNAS_DIARIO = {
    "FIBRA": {
        "id": "id",
        "nombre": "nombre",
        "colocacion": "tipo de colocacion",
        "especificacion": "especificación",
        "propietario": "propietario",
        "estado": "estado de construcción",
        "nombre_ant": "Nombre Antiguo",
        "nodo_a": "nodo a",
        "nodo_b": "nodo b",
        "municipio_a": "municipio a id",  
        "municipio_b": "municipio b id"   
    },
    "EMPALME": {
        "id": "id",
        "nombre": "name",
        "especificacion": "Nombre Especificación",
        "propietario": "propietario",
        "estado": "Construction Status"
    },
    "EDIFICIO": {
        "id": "id",
        "nombre": "nombre",
        "tipo": "tipo",
        "categoria": "categoria"
    },
    "PORT": {
        "id": "Id",
        "description": "Description",
        "tipo_senal": "Tipo Señal",
        "physical_status": "Physical Status",
        "type": "Type",
        "connector_type": "Connector Type",
        "ancho_banda": "Ancho De Banda",
        "unidad_ancho_banda": "Unidad Ancho De Banda",
        "service_level": "Service Level",
        "proyecto_red": "Proyecto De Red",
        "contratista": "Contratista",
        "id_servicio": "Id Servicio",
        "equipo_central": "Equipo Central",
        "puerto_central": "Puerto Central",
        "tipo_servicio": "Tipo Servicio",
        "nombre_cliente": "Nombre Cliente",
        "direccion_cliente_id": "Direccion Cliente Id",
        "comentarios": "Comentarios",
        "connected": "Connected?",
        "anillo": "Anillo",
        "mit_card": "Mit Card",
        "direccion_puerto": "Direccion Puerto",
        "num_factibilidad": "Numero De Factibilidad-Cable",
        "id_shelf": "Id shelf"
    },
    "SHELF": {
        "id": "Id shelf",
        "description": "Description Shelf",
        "estado": "Construction Status",
        "tipo": "Tipo",
        "fabricante": "Fabricante",
        "tecnologia": "Tecnologia",
        "direccion_ip": "Direccion Ip",
        "nivel_bastidor": "Nivel En Bastidor",
        "contratista": "Contratista",
        "comentarios": "Comentarios",
        "nombre_anterior": "Nombre Anterior",
        "migracion_olt": "Migracion Olt",
        "id_bay": "Bay.id"
    },
    "BAY": {
        "id": "Bay.id",
        "number": "Number",
        "description": "Description Bay",
        "estado": "Construction Status",
        "numero_salon": "Numero De Salon",
        "numero_fila": "Numero De Fila",
        "contratista": "Contratista",
        "comentarios": "Comentarios",
        "id_edificio": "Building.id",
        "cambio_service_manager": "Cambio Service Manager"
    }
}

COLUMNAS_ESTRUCTURA = {
    "nombre": "name",
    "floor_dimension_length": "floor dimension (length)",
    "floor_dimension_width": "floor dimension (width)",
    "idtexto": "idtexto",
    "number_of_rooms_length": "number of rooms (length)",
    "number_of_rooms_width": "number of rooms (width)",
    "number_of_floors": "number of floors"
}

JERARQUIA_FIBRA = {
    1: {"campo": "Id", "errores": ["DUPLICADO","CONSECUTIVO"]},
    2: {"campo": "Segmento", "errores": []},
    3: {"campo": "Nombre", "errores": ["NOMBRE (Valor Inválido)","NOMBRE (Debe estar en MAYÚSCULAS)","NOMBRE (Corregir a","NOMBRE (en","COLOCACION (Troncal)","COLOCACION (Ultima Milla)","DE","caracteres inválidos","el símbolo / solo se permite","nombre no puede ser solo","NOMBRE"]},
    4: {"campo": "Nombre Antiguo", "errores": ["NOMBRE_ANT","NOMBRE_ANT (Falta guion cuando NOMBRE tiene guion)"]},
    5: {"campo": "Municipio A", "errores": ["MUNICIPIO (Vacío)","MUNICIPIO A INCONSISTENTE","MUNICIPIO"]},
    6: {"campo": "Municipio B", "errores": ["MUNICIPIO (Vacío)","MUNICIPIO B INCONSISTENTE","MUNICIPIO"]},
    7: {"campo": "Nodo A", "errores": ["NODO INVÁLIDO","NODO A Y B"]},
    8: {"campo": "Nodo B", "errores": ["NODO INVÁLIDO","NODO A Y B"]},
    9: {"campo": "Especificación", "errores": ["ESPECI (senc)","ESPECI (Tipo 2)","ESPECI ('revisar'Ultima Milla cantidad de Hilos)","ESPECI (NO acorde a lo establecido)","ESPECI EA (revisar Tipo)","ESPECI INCONSISTENTE EN CABLE TRONCAL","ESPECI INCONSISTENTE"]},
    10: {"campo": "Estado de Construcción", "errores": ["STATUS ADMITIDOS (En servicio/Plan Abandono/Futuro)","C.STATUS","STATUS"]},
    11: {"campo": "Propietario", "errores": ["PROPIETARIO (ETB)"]},
    12: {"campo": "Tipo de Colocacion", "errores": ["TIPO DE COLOCACION","COLOCACION"]},
    13: {"campo": "Longitud Calculada de Cable", "errores": []},
    14: {"campo": "Longitud Calculada de Fibra", "errores": []},
    15: {"campo": "Perdida Db Medida por Inserccion", "errores": []},
    16: {"campo": "Fecha de Instalacion", "errores": []},
    17: {"campo": "Abcisa Inicial", "errores": []},
    18: {"campo": "Abcisa Final", "errores": []},
    19: {"campo": "Reservas", "errores": []},
    20: {"campo": "Contratista", "errores": []},
    21: {"campo": "Comentarios", "errores": []},
    999: {"campo": "MERGE", "errores": ["TIPO DE CABLE INCONSISTENTE","MERGE","PROBLEMA EN FIBRA","PROBLEMA EN SUB-ANILLO","RECOMENDACIÓN"]}
}

JERARQUIA_EMPALME = {
    1: {"campo": "Id", "errores": ["CONSECUTIVO","DUPLICADO"]},
    2: {"campo": "Segmento", "errores": []},
    3: {"campo": "Nombre", "errores": ["NOMBRE(ERRONEO)","NOMBRE(CONTIENE ESPACIOS)","NOMBRE(LETRA CONSECUTIVA)","NOMBRE(CARACTERES ERRONEOS)","usar:"]},
    4: {"campo": "Propietario", "errores": ["PROPIETARIO (ETB)"]},
    5: {"campo": "Municipio", "errores": []},
    6: {"campo": "Estado de Construcción", "errores": ["STATUS (En servicio/Plan Abandono/Futuro)"]},
    7: {"campo": "Tipo de Caja", "errores": []},
    8: {"campo": "Especificación", "errores": ["ESPECIFICACION ET ERRONEA","ESPECIFICACION (NO acorde a lo establecido)","TIPO DE GRUPO","esperado:","actual:"]},
    9: {"campo": "E2", "errores": []},
    10: {"campo": "Fecha de Instalacion", "errores": []},
    11: {"campo": "Distancia a Splitter 1er nivel", "errores": []},
    12: {"campo": "Distancia a Splitter 2do nivel", "errores": []},
    13: {"campo": "Disponible Para", "errores": []},
    14: {"campo": "Tipo Restriccion", "errores": []},
    15: {"campo": "Fecha Inicio Restriccion", "errores": []},
    16: {"campo": "Responsable Restriccion", "errores": []},
    17: {"campo": "Sangria?", "errores": []},
    18: {"campo": "Método", "errores": []},
    19: {"campo": "fx Available Mounting Positions", "errores": []},
    20: {"campo": "Tiene Riser?", "errores": []},
    21: {"campo": "Red Intema", "errores": []},
    22: {"campo": "Puertos Derivacion Libres", "errores": []},
    23: {"campo": "Contratista", "errores": []},
    999: {"campo": "MERGE", "errores": ["TIPO DE CABLE INCONSISTENTE","INCONSISTENCIA MERGE"]}
}

JERARQUIA_EDIFICIO = {
    1: {"campo": "Id", "errores": []},
    2: {"campo": "Nombre", "errores": ["NOMBRE (Corregir a","NOMBRE (Debe estar en MAYÚSCULAS)","NOMBRE (Valor Inválido)","NOMBRE (Sufijo con espacios)"]},
    3: {"campo": "Tipo", "errores": ["TIPO DE GRUPO"]},
    4: {"campo": "Estado de Construcción", "errores": ["ESTADO"]},
    5: {"campo": "Categoria", "errores": ["CATEGORIA (AAA)"]},
    6: {"campo": "Direccion Edificio", "errores": []},
    7: {"campo": "Numero de Factibilidad", "errores": []},
    8: {"campo": "Comentarios", "errores": []},
    9: {"campo": "Limite", "errores": []},
    10: {"campo": "Direccion Cliente ID", "errores": []},
    11: {"campo": "Posición", "errores": []},
    12: {"campo": "Posición", "errores": []}
}

JERARQUIA_ESTRUCTURA = {
    1: {"campo": "ESTRUCTURA", "errores": ["ESTRUCTURA (NUMBER_OF + FLOOR + ROOMS = 0)","ESTRUCTURA_FLOOR_DIMENSION","ESTRUCTURA_NUMBER_OF_ROOMS","NUM_OF:","FLOOR:","ROOMS:"]}
}

JERARQUIA_PORT = {
    1: {"campo": "Id", "errores": ["DUPLICADO","CONSECUTIVO"]},
    2: {"campo": "Tipo Señal", "errores": ["TIPO SEÑAL INVALIDO","PUERTO RX/TX IDENTIFICADO"]},
    3: {"campo": "Description", "errores": ["DESCRIPTION VACIO"]},
    4: {"campo": "Physical Status", "errores": ["PHYSICAL STATUS INVALIDO","PHYSICAL STATUS VACIO"]},
    5: {"campo": "Type", "errores": ["TYPE INVALIDO","TYPE VACIO","TYPE (solo Equipo Cliente)"]},
    6: {"campo": "Connector Type", "errores": ["CONNECTOR TYPE INVALIDO","CONNECTOR TYPE VACIO","CONNECTOR TYPE (solo Unknown o Biconic)"]},
    7: {"campo": "Ancho De Banda", "errores": ["ANCHO DE BANDA INVALIDO"]},
    8: {"campo": "Unidad Ancho De Banda", "errores": ["UNIDAD ANCHO DE BANDA INVALIDA","UNIDAD ANCHO DE BANDA VACIA"]},
    9: {"campo": "Service Level", "errores": ["SERVICE LEVEL INVALIDO","SERVICE LEVEL VACIO","SERVICE LEVEL (solo Fiber)"]},
    10: {"campo": "Proyecto De Red", "errores": ["PROYECTO DE RED VACIO"]},
    11: {"campo": "Contratista", "errores": ["CONTRATISTA VACIO"]},
    12: {"campo": "Id Servicio", "errores": ["ID SERVICIO VACIO"]},
    13: {"campo": "Equipo Central", "errores": ["EQUIPO CENTRAL VACIO"]},
    14: {"campo": "Puerto Central", "errores": ["PUERTO CENTRAL VACIO"]},
    15: {"campo": "Tipo Servicio", "errores": ["TIPO SERVICIO VACIO"]},
    16: {"campo": "Nombre Cliente", "errores": ["NOMBRE CLIENTE (Valor Inválido)","NOMBRE CLIENTE (Debe estar en MAYÚSCULAS)","NOMBRE (Corregir a","NOMBRE CLIENTE VACIO","NOMBRE CLIENTE INCONSISTENTE","NOMBRE CLIENTE (Sufijo con espacios)"]},
    17: {"campo": "Direccion Cliente Id", "errores": ["DIRECCION CLIENTE ID VACIO"]},
    18: {"campo": "Comentarios", "errores": []},
    19: {"campo": "Connected?", "errores": ["CONNECTED INVALIDO"]},
    20: {"campo": "Anillo", "errores": ["ANILLO VACIO"]},
    21: {"campo": "Mit Card", "errores": ["MIT CARD VACIO"]},
    22: {"campo": "Direccion Puerto", "errores": ["DIRECCION PUERTO VACIO"]},
    23: {"campo": "Numero De Factibilidad-Cable", "errores": ["NUMERO FACTIBILIDAD VACIO"]},
    24: {"campo": "Id shelf", "errores": ["ID SHELF VACIO","ID SHELF INVALIDO"]}
}

JERARQUIA_SHELF = {
    1: {"campo": "Id", "errores": ["DUPLICADO","CONSECUTIVO"]},
    2: {"campo": "Description", "errores": ["DESCRIPTION VACIO","DESCRIPTION INVALIDA","DESCRIPTION CON ESPACIOS","DESCRIPTION INVALIDA","DESCRIPTION INVALIDA"]},  
    3: {"campo": "Construction Status", "errores": ["CONSTRUCTION STATUS INVALIDO"]},
    4: {"campo": "Tipo", "errores": ["TIPO INVALIDO"]},
    5: {"campo": "Fabricante", "errores": []},
    6: {"campo": "Tecnologia", "errores": ["TECNOLOGIA VACIO"]},
    7: {"campo": "Direccion Ip", "errores": ["DIRECCION IP INVALIDA"]},
    8: {"campo": "Nivel En Bastidor", "errores": ["NIVEL BASTIDOR INVALIDO"]},
    9: {"campo": "Contratista", "errores": ["CONTRATISTA VACIO"]},
    10: {"campo": "Comentarios", "errores": []},
    11: {"campo": "Nombre Anterior", "errores": []},
    12: {"campo": "Migracion Olt", "errores": ["MIGRACION OLT INVALIDA"]},
    13: {"campo": "Bay.id", "errores": ["BAY ID VACIO","BAY ID INVALIDO"]}
}

JERARQUIA_BAY = {
    1: {"campo": "Id", "errores": ["DUPLICADO","CONSECUTIVO"]},
    2: {"campo": "Number", "errores": ["NUMBER INVALIDO"]},
    3: {"campo": "Description", "errores": ["DESCRIPTION VACIO","DESCRIPTION CON DOBLE ESPACIADO","DESCRIPTION CON CARACTERES INVÁLIDOS"]},
    4: {"campo": "Construction Status", "errores": ["CONSTRUCTION STATUS INVALIDO"]},
    5: {"campo": "Numero De Salon", "errores": ["NUMERO SALON INVALIDO"]},
    6: {"campo": "Numero De Fila", "errores": ["NUMERO FILA INVALIDO"]},
    7: {"campo": "Contratista", "errores": ["CONTRATISTA VACIO"]},
    8: {"campo": "Comentarios", "errores": []},
    9: {"campo": "Building.id", "errores": ["BUILDING ID VACIO","BUILDING ID INVALIDO"]}
}

RE_NOMBRE_FIBRA_TRONCAL = re.compile(r"^[\dP/]+$")
RE_NOMBRE_FIBRA_TRONCAL_SIMPLE = re.compile(r"^[\d]+$")
RE_NOMBRE_FIBRA_UM_CON_P = re.compile(r"^[\d\-P/]+$")
RE_NOMBRE_FIBRA_UM_SIN_P = re.compile(r"^[\d\-]+$")
RE_NOMBRE_EMPALME_VALIDO = re.compile(r"^(ET|NAP|EA|OB)[0-9_\-P]+$")
RE_NOMBRE_EMPALME_SUB = re.compile(r"^(ET|NAP)\d+\-SUB_\d+$")
RE_NOMBRE_EMPALME_BASE = re.compile(r"^(ET|NAP|EA|OB)[0-9\-_P]*")
RE_NOMBRE_FIBRA_BASE = re.compile(r'^(\d+(?:-\d+)?)')
RE_NOMBRE_FIBRA_NUMERO = re.compile(r'^(\d+)')
RE_ESPEC_HILOS = re.compile(r'(\d+)\s+hilos', re.IGNORECASE)
RE_NOMBRE_ANT_UM = re.compile(r"^[\d\-P\#/]+$|^SUB$")
RE_NOMBRE_ANT_NO_UM = re.compile(r"^[\d\-P\#]+$|^SUB$")
RE_NUMERO_SALON = re.compile(r'^OF \d+$')
RE_DESCRIPCION_BASTIDOR = re.compile(r'^BASTIDOR CLIENTE (\d+)$')
RE_DESCRIPCION_DEMARCADOR = re.compile(r'^DEMARCADOR CLIENTE (\d+)$')
RE_IPV4 = re.compile(r'^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})(/\d{1,2})?$')

def validar_sufijos_pegados(nombre):
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
        (r'\s+N\s+I\s+T\s*\d*$', 'NIT')
    ]
    for patron, sustituto in patrones_sufijos:
        if re.search(patron, nombre_upper):
            nombre_sin_espacios = re.sub(r'\s+', '', nombre_upper)
            return False, f"{JERARQUIA_EDIFICIO[2]['errores'][3]} | Debe estar pegado: '{nombre_sin_espacios}'", nombre_sin_espacios
    return True, None, None

class ErrorOrganizer:
    def __init__(self):
        self.errores_por_id = defaultdict(list)
    
    def agregar_error(self, id_elemento, elemento_nombre, valores_error):
        prioridad, campo = self._obtener_prioridad_campo(elemento_nombre, valores_error)
        self.errores_por_id[id_elemento].append((prioridad, campo, elemento_nombre, valores_error))
    
    def _obtener_prioridad_campo(self, elemento, valores_error):
        mensaje_str = self._extraer_mensaje_error(valores_error)
        
        if elemento == "FIBRA":
            jerarquia = JERARQUIA_FIBRA
        elif elemento == "EMPALME":
            jerarquia = JERARQUIA_EMPALME
        elif elemento == "EDIFICIO":
            jerarquia = JERARQUIA_EDIFICIO
        elif elemento == "MERGE FI-EM":
            jerarquia = JERARQUIA_FIBRA if "MERGE" in mensaje_str else JERARQUIA_EMPALME
        elif "ESTRUCTURA" in elemento:
            jerarquia = JERARQUIA_ESTRUCTURA
        elif elemento == "PORT":
            jerarquia = JERARQUIA_PORT
        elif elemento == "SHELF":
            jerarquia = JERARQUIA_SHELF
        elif elemento == "BAY":
            jerarquia = JERARQUIA_BAY
        else:
            return (999, "Otros")
        
        mensaje_upper = mensaje_str.upper()
        for prioridad, config in jerarquia.items():
            for error_patron in config["errores"]:
                if error_patron.upper() in mensaje_upper:
                    return (prioridad, config["campo"])
        
        return (999, "Otros")
    
    def _extraer_mensaje_error(self, valores_error):
        if isinstance(valores_error, (list, tuple)):
            if len(valores_error) > 3:
                return str(valores_error[3])
            elif len(valores_error) == 1:
                return str(valores_error[0])
            else:
                return str(valores_error)
        else:
            return str(valores_error)
    
    def obtener_errores_organizados(self):
        ids_ordenados = sorted(self.errores_por_id.keys())
        errores_finales = []
        
        for idv in ids_ordenados:
            errores_id = self.errores_por_id[idv]
            
            errores_por_prioridad = defaultdict(list)
            for prio, campo, elem, valores in errores_id:
                errores_por_prioridad[prio].append((campo, elem, valores))
            
            if errores_por_prioridad:
                mejor_prioridad = min(errores_por_prioridad.keys())
                mejores_errores = errores_por_prioridad[mejor_prioridad]
                
                errores_agrupados = defaultdict(list)
                for campo, elemento, valores in mejores_errores:
                    errores_agrupados[(elemento, campo)].append(valores)
                
                for (elemento, campo), lista_valores in errores_agrupados.items():
                    if len(lista_valores) == 1:
                        errores_finales.append((elemento, lista_valores[0]))
                    else:
                        valores_combinados = self._combinar_errores(lista_valores)
                        errores_finales.append((elemento, valores_combinados))
        
        return errores_finales
    
    def _combinar_errores(self, lista_valores):
        id_val = lista_valores[0][0]
        value_val = lista_valores[0][1]
        owner_val = lista_valores[0][2]
        
        mensajes = set()
        for vals in lista_valores:
            if len(vals) > 3:
                mensajes.add(vals[3])
        
        error_combinado = " | ".join(sorted(mensajes))
        return (id_val, value_val, owner_val, error_combinado)

def _to_str_id(val):
    if val is None or val == "":
        return ""
    if isinstance(val, (int, float)):
        try:
            return str(int(val))
        except ValueError:
            return str(val).strip()
    return str(val).strip()

def traducir_estado(estado, modo):
    if not estado:
        return ""
    estado = str(estado).strip()
    if modo == "masivo":
        return {
            "is": "En servicio",
            "pa": "Plan Abandono",
            "pf": "Plan Uso Futuro"
        }.get(estado.lower(), estado)
    return estado

def obtener_columnas_openpyxl(ws):
    return {
        str(c.value).strip().lower(): i - 1
        for i, c in enumerate(ws[1], start=1)
        if c.value
    }

def cargar_om_openpyxl(ruta_macro, hoja="OM"):
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
    for idx, c in enumerate(ws[1], start=1):
        v = str(c.value).strip().lower() if c.value else ""
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
                ids_om.add(_to_str_id(val))
        if idx_elem and idx_tb:
            elem = row[idx_elem - 1]
            tb = row[idx_tb - 1]
            if elem and tb and str(tb).strip().upper() == "ED1":
                exc_ed.add(str(elem).strip())
    wb.close()
    return ids_om, exc_ed

def cargar_responsables(ruta, hoja, columna_clave="cable acceso"):
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
    resp = {}
    for _, row in df.iterrows():
        if clave_norm == "cable acceso":
            clave = str(row.iloc[idx_clave]).split("/")[0].strip() if pd.notna(row.iloc[idx_clave]) else ""
        elif clave_norm == "id edifico":
            clave = _to_str_id(row.iloc[idx_clave])
        else:
            clave = str(row.iloc[idx_clave]).strip() if pd.notna(row.iloc[idx_clave]) else ""
        if clave:
            resp[clave] = str(row.iloc[idx_resp]).strip() if pd.notna(row.iloc[idx_resp]) else ""
    return resp

def cargar_especificaciones_fibra(ruta_archivo, ids_om, modo, incluir_colocacion=False):
    espec = {}
    colocacion_por_fibra = {} if incluir_colocacion else None
    try:
        wb_fi = openpyxl.load_workbook(ruta_archivo, data_only=True, read_only=True)
        ws_fi = wb_fi.worksheets[0]
        col_fi = obtener_columnas_openpyxl(ws_fi)
        cols = COLUMNAS_MASIVO["FIBRA"] if modo == "masivo" else COLUMNAS_DIARIO["FIBRA"]
        idx_id = col_fi.get(cols["id"], -1)
        idx_estado = col_fi.get(cols["estado"], -1)
        idx_nombre = col_fi.get(cols["nombre"], -1)
        idx_espec = col_fi.get(cols["especificacion"], -1)
        idx_colocacion = col_fi.get(cols["colocacion"], -1) if incluir_colocacion else -1
        ids_om_set = set(ids_om)
        for f in ws_fi.iter_rows(min_row=2, values_only=True):
            if idx_id == -1 or idx_estado == -1:
                continue
            idv_raw = f[idx_id]
            idv = _to_str_id(idv_raw)
            estado_raw = f[idx_estado]
            if idv == "" or idv in ids_om_set or estado_raw in (None, ""):
                continue
            if idx_nombre != -1:
                nom = str(f[idx_nombre] or "").strip()
                base = nom.split("/")[0].strip()
            else:
                continue
            if idx_espec != -1:
                esp_raw = str(f[idx_espec] or "").strip()
            else:
                esp_raw = ""
            if base:
                espec[base] = esp_raw.lower()
                if incluir_colocacion and idx_colocacion != -1:
                    tipo_colocacion = str(f[idx_colocacion] or "").strip()
                    colocacion_por_fibra[base] = tipo_colocacion
        wb_fi.close()
    except Exception:
        pass
    return (espec, colocacion_por_fibra) if incluir_colocacion else espec

def cargar_odfs_validaciones(ruta_odfs):
    odfs_data = {}
    try:
        wb = openpyxl.load_workbook(ruta_odfs, data_only=True, read_only=True)
        if "ODFS" not in wb.sheetnames:
            wb.close()
            return odfs_data
        ws = wb["ODFS"]
        headers = {str(cell.value).strip().lower(): idx for idx, cell in enumerate(ws[1], start=1) if cell.value}
        if not all(col in headers for col in ['validacion', 'troncal', 'ids_distribucion']):
            wb.close()
            return odfs_data
        idx_validacion = headers['validacion']
        idx_troncal = headers['troncal']
        idx_ids_distribucion = headers['ids_distribucion']
        for row in ws.iter_rows(min_row=2, values_only=True):
            validacion = str(row[idx_validacion]).strip().upper() if row[idx_validacion] else ""
            if validacion != "OK":
                continue
            troncal = str(row[idx_troncal]).strip() if row[idx_troncal] else ""
            if not troncal:
                continue
            ids_distribucion = str(row[idx_ids_distribucion]).strip() if row[idx_ids_distribucion] else ""
            if not ids_distribucion:
                continue
            ids_list = []
            try:
                ids_limpios = ids_distribucion.replace('"', '').split('|')
                for id_str in ids_limpios:
                    id_clean = id_str.strip()
                    if id_clean:
                        ids_list.append(id_clean)
            except Exception:
                continue
            if ids_list:
                if troncal in odfs_data:
                    odfs_data[troncal].update(ids_list)
                else:
                    odfs_data[troncal] = set(ids_list)
        wb.close()
    except Exception:
        pass
    return odfs_data

def cargar_nombres_edificios(ruta_ed, modo):
    nombres_edificios = {}
    nombres_normalizados = defaultdict(set)
    
    try:
        wb = openpyxl.load_workbook(ruta_ed, data_only=True, read_only=True)
        nombre_hoja = next((h for h in wb.sheetnames 
                           if "edif" in h.lower() or "export" in h.lower() or "output" in h.lower()), 
                           wb.sheetnames[0])
        ws = wb[nombre_hoja]
        col = obtener_columnas_openpyxl(ws)
        
        if modo == "masivo":
            cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
        else:
            cols_ed = COLUMNAS_DIARIO["EDIFICIO"]
        
        idx_id = col.get(cols_ed["id"], -1)
        idx_nombre = col.get(cols_ed["nombre"], -1)
        
        if idx_id == -1 or idx_nombre == -1:
            wb.close()
            return nombres_edificios, nombres_normalizados
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            idv = _to_str_id(row[idx_id]) if idx_id < len(row) else ""
            if not idv:
                continue
            nombre = str(row[idx_nombre] or "").strip() if idx_nombre < len(row) else ""
            if nombre:
                nombre_upper = nombre.upper()
                nombres_edificios[idv] = nombre_upper
                nombre_normalizado = normalizar_nombre_empresa(nombre_upper)
                nombres_normalizados[nombre_normalizado].add(idv)
        
        wb.close()
    except Exception:
        pass
    
    return nombres_edificios, nombres_normalizados

def normalizar_nombre_empresa(nombre):
    if not nombre:
        return ""
    nombre_upper = nombre.upper().strip()
    
    sufijos = [
        r'\s+S\.?A\.?S\.?', r'\s+S\.?A\.?', r'\s+L\.?T\.?D\.?A\.?',
        r'\s+E\.?U\.?', r'\s+N\.?I\.?T\.?\s*\d+', r'\s+&?\s*CIE\.?',
        r'\s+LTDA\.?', r'\s+LTD\.?', r'\s+INC\.?', r'\s+CORP\.?',
        r'\s+S\.?C\.?S\.?', r'\s+E\.?I\.?R\.?L\.?'
    ]
    
    for sufijo in sufijos:
        nombre_upper = re.sub(sufijo, '', nombre_upper)
    
    nombre_upper = re.sub(r'[.,]', '', nombre_upper)
    nombre_upper = re.sub(r'\s+', ' ', nombre_upper).strip()
    
    return nombre_upper

def validar_unidad_ancho_banda_megabyte(valor_unidad):
    if pd.isna(valor_unidad) or str(valor_unidad).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_unidad).strip()
    valor_lower = valor_str.lower()
    if valor_lower in ("megabyte", "kilobyte"):
        return "VÁLIDO", valor_str
    sugerencias = {
        "mb": "Megabyte", "mbps": "Megabyte", "mbyte": "Megabyte",
        "kb": "Kilobyte", "kbps": "Kilobyte", "kbyte": "Kilobyte"
    }
    for clave, sugerencia in sugerencias.items():
        if clave in valor_lower:
            return "UNIDAD_INCORRECTA", sugerencia
    return "INVÁLIDO", "Megabyte o Kilobyte"

def validar_physical_status(valor_status):
    if pd.isna(valor_status) or str(valor_status).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_status).strip()
    if valor_str.lower() in ["in service", "spare"]:
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

def validar_connector_type_unknown_biconic(valor_connector_type):
    if pd.isna(valor_connector_type) or str(valor_connector_type).strip() == "":
        return "VACÍO", ""
    valor_str = str(valor_connector_type).strip()
    if valor_str.lower() in ["unknown", "biconic"]:
        return "VÁLIDO", valor_str
    return "INVÁLIDO", "Unknown o Biconic"

def validar_direccion_ip(direccion_ip):
    if pd.isna(direccion_ip) or str(direccion_ip).strip() == "":
        return "VACÍO", True
    direccion_str = str(direccion_ip).strip()
    match = RE_IPV4.match(direccion_str)
    if not match:
        return "INVÁLIDO", False
    for i in range(1, 5):
        octeto = int(match.group(i))
        if octeto < 0 or octeto > 255:
            return "INVÁLIDO", False
    if match.group(5):
        mascara = int(match.group(5)[1:])
        if mascara < 0 or mascara > 32:
            return "INVÁLIDO", False
    return "VÁLIDO", True

def validar_bastidor_cliente(description):
    if pd.isna(description) or str(description).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(description).strip().upper()
    if desc_str == "BASTIDOR CLIENTE":
        return []
    match = RE_DESCRIPCION_BASTIDOR.match(desc_str)
    if match:
        numero = int(match.group(1))
        if numero >= 2:
            return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(description).strip()}'"]

def validar_demarcador_cliente(description):
    if pd.isna(description) or str(description).strip() == "":
        return ["DESCRIPTION VACIO"]
    desc_str = str(description).strip().upper()
    if desc_str == "DEMARCADOR CLIENTE":
        return []
    match = RE_DESCRIPCION_DEMARCADOR.match(desc_str)
    if match:
        numero = int(match.group(1))
        if numero >= 2:
            return []
    return [f"DESCRIPTION INVALIDA. ACTUAL:'{str(description).strip()}'"]

def extraer_hilos(especificacion):
    match = RE_ESPEC_HILOS.search(str(especificacion))
    return int(match.group(1)) if match else 0

def validar_nombre_ant(nombre_ant, tipo_colocacion):
    if pd.isna(nombre_ant) or str(nombre_ant).strip() == "":
        return True
    nombre_ant_str = str(nombre_ant).strip()
    if tipo_colocacion == "Ultima Milla":
        return bool(RE_NOMBRE_ANT_UM.match(nombre_ant_str.upper()))
    else:
        return bool(RE_NOMBRE_ANT_NO_UM.match(nombre_ant_str.upper()))

def validar_nombre_fibra(nombre, tipo_colocacion):
    errores = []
    nombre_str = str(nombre).strip().upper()
    tipo_colocacion_str = str(tipo_colocacion).strip()
    if tipo_colocacion_str in ("Troncal-Acceso", "Troncal-Transporte"):
        if "-" in nombre_str:
            if "SUB" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_colocacion_str} el guion solo se permite con 'SUB')")
            else:
                if not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str.replace("-SUB", "")):
                    errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (formato inválido con SUB en {tipo_colocacion_str})")
        else:
            if "P" in nombre_str:
                if not RE_NOMBRE_FIBRA_TRONCAL.match(nombre_str):
                    errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_colocacion_str} debe contener solo números, P y / cuando hay P)")
                if "/" in nombre_str and "P" not in nombre_str:
                    errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (el símbolo / solo se permite cuando hay una P en el nombre)")
                if nombre_str == "/" or nombre_str.replace("/", "") == "":
                    errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (nombre no puede ser solo /)")
            else:
                if not RE_NOMBRE_FIBRA_TRONCAL_SIMPLE.match(nombre_str):
                    errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (en {tipo_colocacion_str} debe contener solo números cuando no hay P)")
    else:
        if "P" in nombre_str:
            if not RE_NOMBRE_FIBRA_UM_CON_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (caracteres inválidos en nombre con P)")
            if "/" in nombre_str and "P" not in nombre_str:
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (el símbolo / solo se permite cuando hay una P en el nombre)")
            if nombre_str == "/" or nombre_str.replace("/", "").replace("-", "") == "":
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (nombre no puede ser solo / o -/)")
        else:
            if not RE_NOMBRE_FIBRA_UM_SIN_P.match(nombre_str):
                errores.append(f"{JERARQUIA_FIBRA[3]['errores'][0]} (caracteres inválidos en nombre sin P)")
    return errores

def obtener_nombre_base_fibra(nombre_completo):
    if "-" not in nombre_completo:
        return nombre_completo
    partes = nombre_completo.split("-")
    if len(partes) == 2 and partes[1].isdigit():
        return nombre_completo
    elif len(partes) == 3 and all(p.isdigit() for p in partes[1:]):
        return f"{partes[0]}-{partes[1]}"
    else:
        match = RE_NOMBRE_FIBRA_BASE.match(nombre_completo)
        return match.group(1) if match else nombre_completo

def validar_nombre_empalme(nombre):
    errores = []
    nombre_str = str(nombre).strip().upper()
    if ' ' in nombre_str:
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
        letras_repetidas = set(m.group(0) for m in re.finditer(r'([A-Z])\1', nombre_str))
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][2]}: {', '.join(letras_repetidas)}")
    if nombre_str.startswith("ET") and "-" in nombre_str and "-SUB" not in nombre_str:
        errores.append(f"{JERARQUIA_EMPALME[3]['errores'][0]}")
    return errores

def validar_especificacion_empalme(nombre, especificacion):
    errores = []
    nombre_upper = nombre.upper()
    especificacion_str = str(especificacion).strip()
    ESPEC_VALIDAS = {"Tipo 0","Tipo 1","Tipo 2","Empalme Cliente",
                     "Nap Sotano 12","Nap Sotano 24","Nap Sotano 36","Nap Sotano 48",
                     "Nap Sotano 96","Nap Sotano 144"}
    if especificacion_str not in ESPEC_VALIDAS:
        if especificacion_str.startswith("Nap Sotano "):
            try:
                numero_str = especificacion_str.split("Nap Sotano ")[1]
                if not numero_str.isdigit():
                    errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
            except:
                errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
        else:
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    if nombre_upper.startswith("ET"):
        if especificacion_str not in {"Tipo 0","Tipo 1","Tipo 2","Empalme Cliente"} and not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][0]} (actual: {especificacion_str})")
    elif nombre_upper.startswith("NAP"):
        if not especificacion_str.startswith("Nap Sotano"):
            errores.append(f"ESPECIFICACION NAP ERRONEA (actual: {especificacion_str}, esperado: 'Nap Sotano X')")
    elif nombre_upper.startswith("EA"):
        if especificacion_str != "Tipo 2":
            errores.append(f"{JERARQUIA_EMPALME[8]['errores'][1]} (actual: {especificacion_str})")
    return errores

def determinar_tipo_esperado_por_especificacion(especificacion_fibra):
    esp = especificacion_fibra.lower()
    if esp.startswith("ductado 144"):
        return "Tipo 0"
    elif esp.startswith("ductado 96"):
        return "Tipo 1"
    elif esp.startswith("ductado 12") or esp.startswith("ductado 24") or esp.startswith("ductado 36") or esp.startswith("ductado 48"):
        return "Tipo 2"
    elif "empalme cliente" in esp:
        return "Empalme Cliente"
    elif "nap sotano" in esp:
        return "Nap Sotano"
    else:
        return "Desconocido"

def validar_nombre_cliente_con_edificio(nombre_cliente, building_id, physical_status, comentarios, nombres_edificios, nombres_normalizados):
    CONDICION_VALIDAR = False
    
    if physical_status and str(physical_status).strip().upper() == "IN SERVICE":
        CONDICION_VALIDAR = True
    
    if comentarios and pd.notna(comentarios):
        comentarios_str = str(comentarios).strip().upper()
        if "ADICION" in comentarios_str:
            CONDICION_VALIDAR = True
    
    if not CONDICION_VALIDAR:
        return True, None, None, None
    
    if pd.isna(nombre_cliente) or str(nombre_cliente).strip() == "":
        return False, "", "NOMBRE CLIENTE VACIO", None
    
    nombre_original = str(nombre_cliente).strip()
    nombre_upper = nombre_original.upper()
    
    valido_sufijos, msg_sufijos, nombre_corregido = validar_sufijos_pegados(nombre_original)
    if not valido_sufijos:
        return False, nombre_original, msg_sufijos, nombre_corregido
    
    if building_id and building_id in nombres_edificios:
        nombre_edificio = nombres_edificios[building_id]
        nombre_cliente_upper = nombre_upper
        nombre_edificio_upper = nombre_edificio.upper()
        
        if nombre_cliente_upper in nombre_edificio_upper:
            return True, None, None, None
        
        if nombre_edificio_upper in nombre_cliente_upper:
            return True, None, None, None
        
        return False, nombre_original, f"NOMBRE CLIENTE INCONSISTENTE | Esperado (contiene o está contenido en): '{nombre_edificio}'", nombre_edificio
    
    nombre_cliente_norm = normalizar_nombre_empresa(nombre_upper)
    
    if nombre_cliente_norm in nombres_normalizados:
        edificios_encontrados = list(nombres_normalizados[nombre_cliente_norm])
        nombre_edificio = nombres_edificios.get(edificios_encontrados[0], "")
        if nombre_edificio:
            if nombre_upper in nombre_edificio.upper() or nombre_edificio.upper() in nombre_upper:
                return True, None, None, None
            else:
                return False, nombre_original, f"NOMBRE CLIENTE INCONSISTENTE | Posible edificio: '{nombre_edificio}'", nombre_edificio
    
    errores = []
    patron_caracteres = r"^[A-Za-z0-9 ñÑ()&.,\-]+$"
    if not re.match(patron_caracteres, nombre_original):
        caracteres_invalidos = re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nombre_original)
        unicos = "".join(sorted(set(caracteres_invalidos)))
        if unicos:
            errores.append(f"Caracteres inválidos: '{unicos}'")
    
    if nombre_original != nombre_upper:
        errores.append(f"Debe estar en MAYÚSCULAS: {nombre_upper}")
    
    if errores:
        return False, nombre_original, f"NOMBRE CLIENTE: {' | '.join(errores)}", None
    
    return True, None, None, None

def procesar_fibra_un_paso(ws_fi, col_fi, cols_fi, ids_om, resp, modo, odfs_data, error_organizer):
    errores = []
    espec_fi = {}
    fibras_por_cable = defaultdict(list)
    fibras_por_cable_muni = defaultdict(list)
    fibras_procesadas_nodos = set()
    
    idx_id = col_fi.get(cols_fi["id"], -1)
    idx_nombre = col_fi.get(cols_fi["nombre"], -1)
    idx_colocacion = col_fi.get(cols_fi["colocacion"], -1)
    idx_especificacion = col_fi.get(cols_fi["especificacion"], -1)
    idx_propietario = col_fi.get(cols_fi["propietario"], -1)
    idx_estado = col_fi.get(cols_fi["estado"], -1)
    idx_nombre_ant = col_fi.get(cols_fi.get("nombre_ant", -1), -1)
    idx_nodo_a = col_fi.get(cols_fi.get("nodo_a", -1), -1)
    idx_nodo_b = col_fi.get(cols_fi.get("nodo_b", -1), -1)
    idx_municipio_a = col_fi.get(cols_fi.get("municipio_a", -1), -1)
    idx_municipio_b = col_fi.get(cols_fi.get("municipio_b", -1), -1)
    
    ids_om_set = set(ids_om)
    filas_procesadas = 0
    
    for f in ws_fi.iter_rows(min_row=2, values_only=True):
        filas_procesadas += 1
        if idx_id == -1 or idx_estado == -1:
            continue
        idv_raw = f[idx_id]
        idv = _to_str_id(idv_raw)
        estado_raw = f[idx_estado]
        if idv == "" or idv in ids_om_set or estado_raw in (None, ""):
            continue
        
        if idx_nombre != -1:
            nom = str(f[idx_nombre] or "").strip()
            base_raw = nom.split("/")[0].strip()
            base = obtener_nombre_base_fibra(base_raw)
        else:
            continue
        
        if idx_colocacion != -1:
            tipo = str(f[idx_colocacion] or "").strip()
        else:
            tipo = ""
        
        if idx_especificacion != -1:
            esp_raw = str(f[idx_especificacion] or "").strip()
            esp = esp_raw.lower()
        else:
            esp_raw = ""
            esp = ""
            
        
        if idx_propietario != -1:
            prop = str(f[idx_propietario] or "").strip()
        else:
            prop = ""
        
        est = traducir_estado(estado_raw, modo)
        
        if idx_nombre_ant != -1:
            nombre_ant = f[idx_nombre_ant]
        else:
            nombre_ant = ""
        
        if base:
            espec_fi[base] = esp
        
        e = []
        errores_nombre = validar_nombre_fibra(nom, tipo)
        e.extend(errores_nombre)
        if "-" in nom and tipo != "Ultima Milla" and not nom.upper().endswith("-SUB"):
            e.append(f"{JERARQUIA_FIBRA[3]['errores'][6]} (actual: {tipo})")
        tipo_normalizado = tipo.lower().replace(" ", "-")
        if "-" not in nom and tipo_normalizado not in ("troncal-acceso", "troncal-transporte"):
            e.append(f"{JERARQUIA_FIBRA[3]['errores'][5]} (actual: {tipo})")
        if "doble" in esp:
            e.append(f"{JERARQUIA_FIBRA[9]['errores'][0]} (actual: {esp})")
        if prop != "ETB":
            e.append(f"{JERARQUIA_FIBRA[11]['errores'][0]} (actual: {prop})")
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            e.append(f"{JERARQUIA_FIBRA[10]['errores'][0]} (actual: {est})")
        if tipo == "Ultima Milla":
            hilos = extraer_hilos(esp_raw)
            if hilos > 48:
                e.append(f"{JERARQUIA_FIBRA[9]['errores'][2]} (actual: {hilos} Hilos)")
        if tipo == "Ultima Milla" and nombre_ant:
            if not validar_nombre_ant(nombre_ant, tipo):
                e.append(f"{JERARQUIA_FIBRA[4]['errores'][0]} (actual: {nombre_ant})")
            if "-" in nom and "-" not in str(nombre_ant) and not nom.upper().endswith("-SUB"):
                e.append(f"{JERARQUIA_FIBRA[4]['errores'][1]} (NOMBRE: {nom}, NOMBRE_ANT: {nombre_ant})")
        
        if e:
            owner = resp.get(base, "-")
            for err_msg in e:
                err_tuple = (idv, nom, owner, err_msg)
                errores.append(err_tuple)
                error_organizer.agregar_error(idv, "FIBRA", err_tuple)
        
        if "-" not in nom and tipo in ("Troncal-Acceso", "Troncal-Transporte") and base:
            fibras_por_cable[base].append({
                'id': idv,
                'nombre': nom,
                'especificacion': esp,
                'estado': est,
                'tipo_colocacion': tipo,
                'base_raw': base_raw
            })
        
        if idx_municipio_a != -1 and idx_municipio_b != -1:
            muni_a = _to_str_id(f[idx_municipio_a]) if idx_municipio_a != -1 else ""
            muni_b = _to_str_id(f[idx_municipio_b]) if idx_municipio_b != -1 else ""
            match_num = RE_NOMBRE_FIBRA_NUMERO.match(nom)
            cable_num = match_num.group(1) if match_num else None
            if cable_num:
                fibras_por_cable_muni[cable_num].append({
                    'id': idv,
                    'nombre': nom,
                    'municipio_a': muni_a,
                    'municipio_b': muni_b,
                    'tipo_colocacion': tipo
                })
        
        if odfs_data and idv not in fibras_procesadas_nodos:
            fibras_procesadas_nodos.add(idv)
            match_base = RE_NOMBRE_FIBRA_BASE.match(nom)
            if match_base:
                base_nombre = match_base.group(1)
                cable_odfs = None
                ids_permitidos = None
                if base_nombre in odfs_data:
                    cable_odfs = base_nombre
                    ids_permitidos = odfs_data[base_nombre]
                if not cable_odfs and "-" in base_nombre:
                    partes = base_nombre.split("-")
                    if partes[0].isdigit() and partes[0] in odfs_data:
                        cable_odfs = partes[0]
                        ids_permitidos = odfs_data[partes[0]]
                if not cable_odfs:
                    for troncal in odfs_data.keys():
                        if base_nombre.startswith(troncal) or troncal.startswith(base_nombre.split("-")[0]):
                            cable_odfs = troncal
                            ids_permitidos = odfs_data[troncal]
                            break
                if cable_odfs and ids_permitidos:
                    nodo_a = _to_str_id(f[idx_nodo_a]) if idx_nodo_a != -1 else ""
                    nodo_b = _to_str_id(f[idx_nodo_b]) if idx_nodo_b != -1 else ""
                    owner = prop if prop else "-"
                    if len(ids_permitidos) == 1:
                        id_req = next(iter(ids_permitidos))
                        if nodo_a != id_req or nodo_b != id_req:
                            mensaje = f"{JERARQUIA_FIBRA[7]['errores'][0]} | CORREGIR SEGÚN ODFS: A='{id_req}', B='{id_req}' | ACTUAL A='{nodo_a}' B='{nodo_b}'"
                            err_tuple = (idv, nom, owner, mensaje)
                            errores.append(err_tuple)
                            error_organizer.agregar_error(idv, "FIBRA", err_tuple)
                    else:
                        nodos_invalidos_info = []
                        if nodo_a and nodo_a not in ids_permitidos:
                            nodos_invalidos_info.append(f"A:'{nodo_a}'")
                        if nodo_b and nodo_b not in ids_permitidos:
                            nodos_invalidos_info.append(f"B:'{nodo_b}'")
                        if nodos_invalidos_info:
                            mensaje = f"{JERARQUIA_FIBRA[7]['errores'][1]} | Nodos no coinciden con ODFS"
                            err_tuple = (idv, nom, owner, mensaje)
                            errores.append(err_tuple)
                            error_organizer.agregar_error(idv, "FIBRA", err_tuple)
    
    for cable_base, fibras in fibras_por_cable.items():
        if len(fibras) <= 1:
            continue
        especs = [f['especificacion'] for f in fibras]
        cont_espec = Counter(especs)
        if len(cont_espec) > 1:
            esp_mayor = cont_espec.most_common(1)[0][0]
            for fib in fibras:
                if fib['especificacion'] != esp_mayor:
                    owner = resp.get(cable_base, "-")
                    err_msg = (f"{JERARQUIA_FIBRA[999]['errores'][1]} | "
                               f"Cable '{cable_base}' tiene mayoría '{esp_mayor}' | "
                               f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{fib['especificacion']}'")
                    err_tuple = (fib['id'], fib['nombre'], owner, err_msg)
                    errores.append(err_tuple)
                    error_organizer.agregar_error(fib['id'], "FIBRA", err_tuple)
        estados = [f['estado'] for f in fibras]
        cont_est = Counter(estados)
        if len(cont_est) > 1:
            est_mayor = cont_est.most_common(1)[0][0]
            for fib in fibras:
                if fib['estado'] != est_mayor:
                    owner = resp.get(cable_base, "-")
                    err_msg = (f"{JERARQUIA_FIBRA[10]['errores'][0]} | "
                               f"MAYORÍA: '{est_mayor}' , ACTUAL: '{fib['estado']}'")
                    err_tuple = (fib['id'], fib['nombre'], owner, err_msg)
                    errores.append(err_tuple)
                    error_organizer.agregar_error(fib['id'], "FIBRA", err_tuple)
    
    municipios_pred = {}
    for cable_num, fibras in fibras_por_cable_muni.items():
        es_troncal = any(f['tipo_colocacion'] in ("Troncal-Acceso", "Troncal-Transporte") for f in fibras)
        if es_troncal:
            continue
        todos_munis = []
        for f in fibras:
            if f['municipio_a']:
                todos_munis.append(f['municipio_a'])
            if f['municipio_b']:
                todos_munis.append(f['municipio_b'])
        if todos_munis:
            cnt = Counter(todos_munis)
            mayor = cnt.most_common(1)[0][0]
            porcentaje = (cnt[mayor] / len(todos_munis)) * 100
            municipios_pred[cable_num] = (mayor, porcentaje)
    
    for cable_num, fibras in fibras_por_cable_muni.items():
        if cable_num not in municipios_pred:
            continue
        pred, pct = municipios_pred[cable_num]
        for fib in fibras:
            if fib['tipo_colocacion'] in ("Troncal-Acceso", "Troncal-Transporte"):
                continue
            idv = fib['id']
            nombre_fibra = fib['nombre']
            owner = resp.get(cable_num, "-")
            muni_a = fib['municipio_a']
            muni_b = fib['municipio_b']
            errores_fila = []
            if muni_a == "":
                msg = f"{JERARQUIA_FIBRA[5]['errores'][0]} | SUGERIDO: '{pred}' ({pct:.1f}%)"
                errores_fila.append(msg)
            elif muni_a != pred:
                msg = f"{JERARQUIA_FIBRA[5]['errores'][1]} | Municipio A '{muni_a}' no coincide con predominante '{pred}' ({pct:.1f}%)"
                errores_fila.append(msg)
            if muni_b == "":
                msg = f"{JERARQUIA_FIBRA[6]['errores'][0]} | SUGERIDO: '{pred}' ({pct:.1f}%)"
                errores_fila.append(msg)
            elif muni_b != pred:
                msg = f"{JERARQUIA_FIBRA[6]['errores'][1]} | Municipio B '{muni_b}' no coincide con predominante '{pred}' ({pct:.1f}%)"
                errores_fila.append(msg)
            for err_msg in errores_fila:
                err_tuple = (idv, nombre_fibra, owner, err_msg)
                errores.append(err_tuple)
                error_organizer.agregar_error(idv, "FIBRA", err_tuple)
    
    return errores, espec_fi, filas_procesadas

def procesar_empalme_un_paso(ws_em, col_em, cols_em, ids_om, resp, modo, error_organizer):
    errores = []
    datos_em = []
    grupos = defaultdict(list)
    nombres_vistos = defaultdict(list)
    
    idx_id = col_em.get(cols_em["id"], -1)
    idx_nombre = col_em.get(cols_em["nombre"], -1)
    idx_especificacion = col_em.get(cols_em["especificacion"], -1)
    idx_propietario = col_em.get(cols_em["propietario"], -1)
    idx_estado = col_em.get(cols_em["estado"], -1)
    
    ids_om_set = set(ids_om)
    filas_procesadas = 0
    
    for f in ws_em.iter_rows(min_row=2, values_only=True):
        filas_procesadas += 1
        if idx_id == -1 or idx_estado == -1:
            continue
        idv_raw = f[idx_id]
        idv = _to_str_id(idv_raw)
        estado_raw = f[idx_estado]
        if idv == "" or idv in ids_om_set or estado_raw in (None, ""):
            continue
        
        if idx_nombre != -1:
            nom = str(f[idx_nombre] or "").strip()
            corto = nom[3:] if nom.upper().startswith("NAP") else (nom[2:] if len(nom) > 2 else nom)
        else:
            continue
        
        if idx_especificacion != -1:
            esp = str(f[idx_especificacion] or "").strip()
        else:
            esp = ""
        
        if idx_propietario != -1:
            prop = str(f[idx_propietario] or "").strip()
        else:
            prop = ""
        
        est = traducir_estado(estado_raw, modo)
        
        g = ""
        if (nom.startswith("ET") or nom.startswith("NAP")) and "_" in nom:
            if "-SUB" in nom.upper():
                match = re.match(r'^(ET|NAP)(\d+-SUB)', nom.upper())
                g = f"{match.group(1)}{match.group(2)}" if match else ""
            else:
                g = nom.split("_")[0]
        
        datos_em.append((idv, nom, esp, prop, est, g, corto))
        nombres_vistos[nom].append(idv)
        
        if g:
            if g.startswith("NAP"):
                if esp.startswith("Nap Sotano"):
                    grupos[g].append("Nap Sotano")
                else:
                    grupos[g].append(esp)
            else:
                if esp in {"Tipo 0","Tipo 1","Tipo 2","Empalme Cliente"}:
                    grupos[g].append(esp)
        
        e = []
        e.extend(validar_nombre_empalme(nom))
        e.extend(validar_especificacion_empalme(nom, esp))
        if prop != "ETB":
            e.append(f"{JERARQUIA_EMPALME[4]['errores'][0]} (actual: {prop})")
        if est not in ("En servicio", "Plan Abandono", "Plan Uso Futuro"):
            e.append(f"{JERARQUIA_EMPALME[6]['errores'][0]} (actual: {est})")
        if "-SUB" not in nom.upper():
            if not RE_NOMBRE_EMPALME_VALIDO.match(nom.upper()):
                base_match = RE_NOMBRE_EMPALME_BASE.match(nom.upper())
                nombre_corregido = base_match.group(0) if base_match else ""
                if nombre_corregido and nombre_corregido != nom.upper():
                    e.append(f"{JERARQUIA_EMPALME[3]['errores'][4]} (usar: {nombre_corregido})")
        if e:
            owner = resp.get(corto, "-")
            for err_msg in e:
                err_tuple = (idv, nom, owner, err_msg)
                errores.append(err_tuple)
                error_organizer.agregar_error(idv, "EMPALME", err_tuple)
    
    duplicados_reportados = set()
    for nom, ids in nombres_vistos.items():
        if len(ids) > 1 and not ((nom.startswith("ET") or nom.startswith("NAP")) and "_" in nom):
            ids_str = " , ".join(ids)
            corto = nom[3:] if nom.upper().startswith("NAP") else (nom[2:] if len(nom) > 2 else nom)
            owner = resp.get(corto, "-")
            err_tuple = (ids_str, nom, owner, f"{JERARQUIA_EMPALME[1]['errores'][1]})")
            errores.append(err_tuple)
            error_organizer.agregar_error(ids_str, "EMPALME", err_tuple)
            duplicados_reportados.add(nom)
    
    mayor = {}
    for g, v in grupos.items():
        if v:
            contador = Counter(v)
            mayor[g] = contador.most_common(1)[0][0]
    
    for (idv, nom, esp, prop, est, g, corto) in datos_em:
        if nom in duplicados_reportados:
            continue
        if g and mayor.get(g):
            if g.startswith("NAP"):
                es_nap_sotano_actual = esp.startswith("Nap Sotano")
                es_nap_sotano_mayoria = mayor[g] == "Nap Sotano"
                if es_nap_sotano_actual != es_nap_sotano_mayoria:
                    err_msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (para {g}, esperado: '{mayor[g]}', actual: '{esp}')"
                    owner = resp.get(corto, "-")
                    err_tuple = (idv, nom, owner, err_msg)
                    errores.append(err_tuple)
                    error_organizer.agregar_error(idv, "EMPALME", err_tuple)
            else:
                if esp in {"Tipo 0","Tipo 1","Tipo 2","Empalme Cliente"} and esp != mayor[g]:
                    err_msg = f"{JERARQUIA_EMPALME[8]['errores'][2]} (esperado: '{mayor[g]}', actual: '{esp}')"
                    owner = resp.get(corto, "-")
                    err_tuple = (idv, nom, owner, err_msg)
                    errores.append(err_tuple)
                    error_organizer.agregar_error(idv, "EMPALME", err_tuple)
    
    return errores, datos_em, filas_procesadas

def validar_consecutividad_empalmes(datos_em, ids_om, resp):
    errs = []
    grupos_consecutivos_sin_sub = defaultdict(list)
    grupos_consecutivos_con_sub = defaultdict(list)
    duplicados_ya_reportados = set()
    ids_om_set = set(ids_om)
    
    for idv, nom, esp, prop, est, g, corto in datos_em:
        if idv in ids_om_set or idv == "":
            continue
        if not (nom.startswith("ET") or nom.startswith("NAP")) or "_" not in nom:
            continue
        tiene_sub = "SUB" in nom.upper()
        match = re.match(r'^(ET|NAP)(\d+)_(\d+)$', nom.upper())
        if not match:
            continue
        prefijo_base = match.group(1)
        numero_cable = match.group(2)
        prefijo = f"{prefijo_base}{numero_cable}"
        try:
            consecutivo = int(match.group(3))
        except ValueError:
            continue
        
        if tiene_sub:
            grupos_consecutivos_con_sub[prefijo].append({
                'id': idv, 'nombre': nom, 'consecutivo': consecutivo,
                'corto': corto, 'con_sub': True, 'tipo': prefijo_base
            })
        else:
            grupos_consecutivos_sin_sub[prefijo].append({
                'id': idv, 'nombre': nom, 'consecutivo': consecutivo,
                'corto': corto, 'con_sub': False, 'tipo': prefijo_base
            })
    
    for prefijo, lista_empalmes in grupos_consecutivos_sin_sub.items():
        lista_empalmes.sort(key=lambda x: x['consecutivo'])
        consecutivos_vistos = {}
        for em in lista_empalmes:
            consec = em['consecutivo']
            nombre_completo = f"{prefijo}_{consec}"
            if consec in consecutivos_vistos:
                if nombre_completo not in duplicados_ya_reportados:
                    id_anterior = consecutivos_vistos[consec]['id']
                    nom_anterior = consecutivos_vistos[consec]['nombre']
                    owner = resp.get(em['corto'], "-")
                    ids_duplicados = f"{id_anterior} , {em['id']}"
                    errs.append((ids_duplicados, f"{nom_anterior}", owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][1]})"))
                    duplicados_ya_reportados.add(nombre_completo)
            else:
                consecutivos_vistos[consec] = em
        
        if not duplicados_ya_reportados:
            for i in range(len(lista_empalmes) - 1):
                actual = lista_empalmes[i]
                siguiente = lista_empalmes[i + 1]
                consec_actual = actual['consecutivo']
                consec_siguiente = siguiente['consecutivo']
                diferencia = consec_siguiente - consec_actual
                owner = resp.get(actual['corto'], "-")
                
                if diferencia < 0:
                    errs.append((siguiente['id'], siguiente['nombre'], owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][0]} _{consec_actual} y un duplicado {prefijo}_{consec_siguiente})"))
                elif diferencia == 2:
                    faltante = consec_actual + 1
                    errs.append((actual['id'], actual['nombre'], owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][0]} _{faltante} después del _{consec_actual})"))
                elif diferencia > 2:
                    inicio_falta = consec_actual + 1
                    fin_falta = consec_siguiente - 1
                    if fin_falta - inicio_falta >= 1:
                        lista_faltantes = list(range(inicio_falta, fin_falta + 1))
                        if len(lista_faltantes) <= 3:
                            faltantes_str = ", ".join([f"_{n}" for n in lista_faltantes])
                        else:
                            faltantes_str = f"_{inicio_falta} al _{fin_falta}"
                        errs.append((actual['id'], actual['nombre'], owner,
                                     f"{JERARQUIA_EMPALME[1]['errores'][0]} _{consec_actual} y el _{consec_siguiente}. Faltan: {faltantes_str})"))
    
    for prefijo, lista_empalmes in grupos_consecutivos_con_sub.items():
        lista_empalmes.sort(key=lambda x: x['consecutivo'])
        consecutivos_vistos = {}
        for em in lista_empalmes:
            consec = em['consecutivo']
            nombre_completo = f"{prefijo}_{consec}"
            if consec in consecutivos_vistos:
                if nombre_completo not in duplicados_ya_reportados:
                    id_anterior = consecutivos_vistos[consec]['id']
                    nom_anterior = consecutivos_vistos[consec]['nombre']
                    owner = resp.get(em['corto'], "-")
                    ids_duplicados = f"{id_anterior}, {em['id']}"
                    errs.append((ids_duplicados, f"{nom_anterior}", owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][1]} {prefijo}_{consec})"))
                    duplicados_ya_reportados.add(nombre_completo)
            else:
                consecutivos_vistos[consec] = em
        
        if not duplicados_ya_reportados:
            for i in range(len(lista_empalmes) - 1):
                actual = lista_empalmes[i]
                siguiente = lista_empalmes[i + 1]
                consec_actual = actual['consecutivo']
                consec_siguiente = siguiente['consecutivo']
                diferencia = consec_siguiente - consec_actual
                owner = resp.get(actual['corto'], "-")
                if diferencia < 0:
                    errs.append((siguiente['id'], siguiente['nombre'], owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][0]} _{consec_actual} y un duplicado {prefijo}_{consec_siguiente})"))
                elif diferencia == 2:
                    faltante = consec_actual + 1
                    errs.append((actual['id'], actual['nombre'], owner,
                                 f"{JERARQUIA_EMPALME[1]['errores'][0]} _{faltante} después del _{consec_actual})"))
                elif diferencia > 2:
                    inicio_falta = consec_actual + 1
                    fin_falta = consec_siguiente - 1
                    if fin_falta - inicio_falta >= 1:
                        lista_faltantes = list(range(inicio_falta, fin_falta + 1))
                        if len(lista_faltantes) <= 3:
                            faltantes_str = ", ".join([f"_{n}" for n in lista_faltantes])
                        else:
                            faltantes_str = f"_{inicio_falta} al _{fin_falta}"
                        errs.append((actual['id'], actual['nombre'], owner,
                                     f"{JERARQUIA_EMPALME[1]['errores'][0]} _{consec_actual} y el _{consec_siguiente}. Faltan: {faltantes_str})"))
    return errs

def validar_tipo_consistente_cable(datos_em, ids_om, resp):
    errs = []
    cables_principales = defaultdict(list)
    sub_cables = defaultdict(list)
    ESPEC_VALIDAS = {"Tipo 0","Tipo 1","Tipo 2","Empalme Cliente",
                     "Nap Sotano 12","Nap Sotano 24","Nap Sotano 36","Nap Sotano 48",
                     "Nap Sotano 96","Nap Sotano 144"}
    ids_om_set = set(ids_om)
    
    for idv, nom, esp, prop, est, g, corto in datos_em:
        if idv in ids_om_set or idv == "":
            continue
        if not (nom.startswith("ET") or nom.startswith("NAP")):
            continue
        if esp not in ESPEC_VALIDAS and not esp.startswith("Nap Sotano "):
            continue
        match = re.match(r'^(ET|NAP)(\d+)', nom.upper())
        if not match:
            continue
        prefijo_base = match.group(1)
        cable_base = f"{prefijo_base}{match.group(2)}"
        if "-SUB" in nom.upper():
            match_sub = re.match(r'^(ET|NAP)(\d+)-SUB', nom.upper())
            if match_sub:
                sub_cable = f"{match_sub.group(1)}{match_sub.group(2)}-SUB"
                sub_cables[sub_cable].append({
                    'id': idv, 'nombre': nom, 'especificacion': esp, 'corto': corto
                })
        else:
            cables_principales[cable_base].append({
                'id': idv, 'nombre': nom, 'especificacion': esp, 'corto': corto
            })
    
    for cable, empalmes in cables_principales.items():
        if len(empalmes) <= 1:
            continue
        especs = [em['especificacion'] for em in empalmes]
        if cable.startswith("NAP"):
            agrup = []
            for x in especs:
                if x.startswith("Nap Sotano"):
                    agrup.append("Nap Sotano")
                else:
                    agrup.append(x)
            cont = Counter(agrup)
        else:
            cont = Counter(especs)
        if len(cont) == 1:
            continue
        esp_mayor = cont.most_common(1)[0][0]
        for em in empalmes:
            if cable.startswith("NAP"):
                es_nap_actual = em['especificacion'].startswith("Nap Sotano")
                es_nap_mayor = esp_mayor == "Nap Sotano"
                if es_nap_actual != es_nap_mayor:
                    numero_cable = cable[3:]
                    owner = resp.get(numero_cable, "-")
                    err_msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | "
                               f"Cable {cable} tiene mayoría '{esp_mayor}' | "
                               f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                    errs.append((em['id'], em['nombre'], owner, err_msg))
            else:
                if em['especificacion'] != esp_mayor:
                    numero_cable = cable[2:]
                    owner = resp.get(numero_cable, "-")
                    err_msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | "
                               f"Cable {cable} tiene mayoría '{esp_mayor}' | "
                               f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                    errs.append((em['id'], em['nombre'], owner, err_msg))
    
    for sub_cable, empalmes in sub_cables.items():
        if len(empalmes) <= 1:
            continue
        especs = [em['especificacion'] for em in empalmes]
        if "NAP" in sub_cable:
            agrup = []
            for x in especs:
                if x.startswith("Nap Sotano"):
                    agrup.append("Nap Sotano")
                else:
                    agrup.append(x)
            cont = Counter(agrup)
        else:
            cont = Counter(especs)
        if len(cont) == 1:
            continue
        esp_mayor = cont.most_common(1)[0][0]
        for em in empalmes:
            if "NAP" in sub_cable:
                es_nap_actual = em['especificacion'].startswith("Nap Sotano")
                es_nap_mayor = esp_mayor == "Nap Sotano"
                if es_nap_actual != es_nap_mayor:
                    numero_match = re.search(r'NAP(\d+)', sub_cable)
                    numero_cable = numero_match.group(1) if numero_match else sub_cable[3:]
                    owner = resp.get(numero_cable, "-")
                    err_msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | "
                               f"Cable {sub_cable} tiene mayoría '{esp_mayor}' | "
                               f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                    errs.append((em['id'], em['nombre'], owner, err_msg))
            else:
                if em['especificacion'] != esp_mayor:
                    numero_match = re.search(r'ET(\d+)', sub_cable)
                    numero_cable = numero_match.group(1) if numero_match else sub_cable[2:]
                    owner = resp.get(numero_cable, "-")
                    err_msg = (f"{JERARQUIA_EMPALME[999]['errores'][0]} | "
                               f"Cable {sub_cable} tiene mayoría '{esp_mayor}' | "
                               f"MAYORÍA: '{esp_mayor}' , ACTUAL: '{em['especificacion']}'")
                    errs.append((em['id'], em['nombre'], owner, err_msg))
    return errs

def procesar_ports_un_paso(ruta_ports, ids_om, modo, nombres_edificios, nombres_normalizados):
    errores_port = []
    errores_shelf = []
    errores_bay = []
    
    try:
        wb = openpyxl.load_workbook(ruta_ports, data_only=True, read_only=True)
        if "Port" in wb.sheetnames:
            ws = wb["Port"]
        else:
            ws = wb.worksheets[0]
        
        col = obtener_columnas_openpyxl(ws)
        
        if modo == "masivo":
            cols_port = COLUMNAS_MASIVO["PORT"]
            cols_shelf = COLUMNAS_MASIVO["SHELF"]
            cols_bay = COLUMNAS_MASIVO["BAY"]
        else:
            cols_port = COLUMNAS_DIARIO["PORT"]
            cols_shelf = COLUMNAS_DIARIO["SHELF"]
            cols_bay = COLUMNAS_DIARIO["BAY"]
        
        idx_id = col.get(cols_port["id"].lower(), -1)
        idx_physical_status = col.get(cols_port["physical_status"].lower(), -1)
        idx_tipo_senal = col.get(cols_port["tipo_senal"].lower(), -1)
        idx_description = col.get(cols_port["description"].lower(), -1)
        idx_unidad_ancho_banda = col.get(cols_port["unidad_ancho_banda"].lower(), -1)
        idx_type = col.get(cols_port["type"].lower(), -1)
        idx_connector_type = col.get(cols_port["connector_type"].lower(), -1)
        idx_service_level = col.get(cols_port["service_level"].lower(), -1)
        idx_contratista = col.get(cols_port["contratista"].lower(), -1)
        idx_nombre_cliente = col.get(cols_port["nombre_cliente"].lower(), -1)
        idx_comentarios = col.get(cols_port["comentarios"].lower(), -1)
        idx_direccion_cliente_id = col.get(cols_port["direccion_cliente_id"].lower(), -1)
        
        idx_shelf_id = col.get(cols_shelf["id"].lower(), -1)
        idx_shelf_desc = col.get(cols_shelf["description"].lower(), -1)
        idx_shelf_estado = col.get(cols_shelf["estado"].lower(), -1)
        idx_shelf_tipo = col.get(cols_shelf["tipo"].lower(), -1)
        idx_shelf_fabricante = col.get(cols_shelf["fabricante"].lower(), -1)
        idx_shelf_tecnologia = col.get(cols_shelf["tecnologia"].lower(), -1)
        idx_shelf_direccion_ip = col.get(cols_shelf["direccion_ip"].lower(), -1)
        idx_shelf_nivel_bastidor = col.get(cols_shelf["nivel_bastidor"].lower(), -1)
        idx_shelf_contratista = col.get(cols_shelf["contratista"].lower(), -1)
        idx_shelf_migracion_olt = col.get(cols_shelf["migracion_olt"].lower(), -1)
        idx_shelf_id_bay = col.get(cols_shelf["id_bay"].lower(), -1)
        
        idx_bay_id = col.get(cols_bay["id"].lower(), -1)
        idx_bay_number = col.get(cols_bay["number"].lower(), -1)
        idx_bay_description = col.get(cols_bay["description"].lower(), -1)
        idx_bay_estado = col.get(cols_bay["estado"].lower(), -1)
        idx_bay_numero_salon = col.get(cols_bay["numero_salon"].lower(), -1)
        idx_bay_numero_fila = col.get(cols_bay["numero_fila"].lower(), -1)
        idx_bay_contratista = col.get(cols_bay["contratista"].lower(), -1)
        idx_bay_id_edificio = col.get(cols_bay["id_edificio"].lower(), -1)
        idx_bay_cambio_service_manager = col.get(cols_bay["cambio_service_manager"].lower(), -1)
        
        col_name_building = None
        for key in col.keys():
            if "name_building" in key or "name building" in key:
                col_name_building = key
                break
        idx_name_building = col.get(col_name_building, -1) if col_name_building else -1
        
        col_building_id = None
        for key in col.keys():
            if "building.id" in key.lower() or "building id" in key.lower():
                col_building_id = key
                break
        idx_building_id = col.get(col_building_id, -1) if col_building_id else -1
        
        ids_om_set = set(ids_om)
        filas_procesadas = 0
        
        for f in ws.iter_rows(min_row=2, values_only=True):
            filas_procesadas += 1
            
            if idx_id != -1:
                idv = _to_str_id(f[idx_id]) if idx_id < len(f) else ""
                if idv and idv not in ids_om_set:
                    es_nodo = False
                    name_building = ""
                    if idx_name_building != -1 and idx_name_building < len(f):
                        nb = f[idx_name_building]
                        if nb:
                            name_building = str(nb).strip()
                            if "NODO" in name_building.upper():
                                es_nodo = True
                    if not es_nodo:
                        for key, idx in col.items():
                            if idx < len(f) and f[idx] is not None:
                                valor = str(f[idx])
                                if "NODO" in valor.upper():
                                    es_nodo = True
                                    break
                    
                    building_id = ""
                    if idx_building_id != -1 and idx_building_id < len(f):
                        bid = f[idx_building_id]
                        if bid:
                            building_id = _to_str_id(bid)
                    
                    tipo_senal = str(f[idx_tipo_senal]).strip() if idx_tipo_senal != -1 and idx_tipo_senal < len(f) else ""
                    description = str(f[idx_description]).strip() if idx_description != -1 and idx_description < len(f) else ""
                    unidad_ancho_banda_raw = f[idx_unidad_ancho_banda] if idx_unidad_ancho_banda != -1 and idx_unidad_ancho_banda < len(f) else None
                    type_raw = f[idx_type] if idx_type != -1 and idx_type < len(f) else None
                    connector_type_raw = f[idx_connector_type] if idx_connector_type != -1 and idx_connector_type < len(f) else None
                    service_level_raw = f[idx_service_level] if idx_service_level != -1 and idx_service_level < len(f) else None
                    ps = f[idx_physical_status] if idx_physical_status != -1 and idx_physical_status < len(f) else None
                    nombre_cliente_raw = f[idx_nombre_cliente] if idx_nombre_cliente != -1 and idx_nombre_cliente < len(f) else None
                    comentarios_raw = f[idx_comentarios] if idx_comentarios != -1 and idx_comentarios < len(f) else None
                    direccion_cliente_id_raw = f[idx_direccion_cliente_id] if idx_direccion_cliente_id != -1 and idx_direccion_cliente_id < len(f) else None
                    
                    mensajes = []
                    
                    if ps is not None:
                        ps_upper = str(ps).strip().upper()
                        if ps_upper == "IN SERVICE" and not es_nodo:
                            if direccion_cliente_id_raw is None or str(direccion_cliente_id_raw).strip() == "":
                                mensajes.append(f"{JERARQUIA_PORT[17]['errores'][0]}")
                    
                    if not es_nodo:
                        if tipo_senal != "Rx - Tx":
                            if tipo_senal:
                                mensajes.append(f"{JERARQUIA_PORT[2]['errores'][0]} (actual: '{tipo_senal}')")
                            else:
                                mensajes.append(f"{JERARQUIA_PORT[2]['errores'][0]} (VACÍO)")
                        st, sug = validar_type_equipo_cliente(type_raw)
                        if st != "VÁLIDO":
                            if st == "VACÍO":
                                mensajes.append(f"{JERARQUIA_PORT[5]['errores'][1]}")
                            else:
                                mensajes.append(f"{JERARQUIA_PORT[5]['errores'][2]} (actual: '{type_raw}')")
                        stc, sugc = validar_connector_type_unknown_biconic(connector_type_raw)
                        if stc != "VÁLIDO":
                            if stc == "VACÍO":
                                mensajes.append(f"{JERARQUIA_PORT[6]['errores'][1]}")
                            else:
                                mensajes.append(f"{JERARQUIA_PORT[6]['errores'][2]} (actual: '{connector_type_raw}')")
                        sts, sgs = validar_service_level_fiber(service_level_raw)
                        if sts != "VÁLIDO":
                            if sts == "VACÍO":
                                mensajes.append(f"{JERARQUIA_PORT[9]['errores'][1]}")
                            else:
                                mensajes.append(f"{JERARQUIA_PORT[9]['errores'][2]} (actual: '{service_level_raw}')")
                    
                    if ps is not None:
                        stp, sgp = validar_physical_status(ps)
                        if stp != "VÁLIDO":
                            if stp == "VACÍO":
                                mensajes.append(f"{JERARQUIA_PORT[4]['errores'][1]}")
                            else:
                                mensajes.append(f"{JERARQUIA_PORT[4]['errores'][0]} (actual: '{ps}')")
                    
                    stu, sgu = validar_unidad_ancho_banda_megabyte(unidad_ancho_banda_raw)
                    if stu != "VÁLIDO":
                        if stu == "VACÍO":
                            mensajes.append(f"{JERARQUIA_PORT[8]['errores'][1]}")
                        else:
                            mensajes.append(f"{JERARQUIA_PORT[8]['errores'][0]} (actual: '{unidad_ancho_banda_raw}')")
                    
                    if idx_nombre_cliente != -1:
                        es_valido, valor_actual, error_msg, nombre_sugerido = validar_nombre_cliente_con_edificio(
                            nombre_cliente_raw, building_id, ps, comentarios_raw, nombres_edificios, nombres_normalizados
                        )
                        if not es_valido and error_msg is not None:
                            errores_port.append((idv, valor_actual or "", "", error_msg))
                            continue
                    
                    if mensajes:
                        error_msg = " | ".join(mensajes)
                        if es_nodo:
                            error_msg += f" | NOTA: Name_building='{name_building}' → Validaciones de Tipo Señal, Type, Connector Type, Service Level y Direccion Cliente Id omitidas para nodos"
                        errores_port.append((idv, f"PORT: {description or 'Sin descripción'}", "", error_msg))
            
            if idx_shelf_id != -1:
                shelf_idv = _to_str_id(f[idx_shelf_id]) if idx_shelf_id < len(f) else ""
                if shelf_idv and shelf_idv not in ids_om_set:
                    es_nodo_shelf = False
                    name_building_shelf = ""
                    if idx_name_building != -1 and idx_name_building < len(f):
                        nb = f[idx_name_building]
                        if nb:
                            name_building_shelf = str(nb).strip()
                            if "NODO" in name_building_shelf.upper():
                                es_nodo_shelf = True
                    
                    description_shelf = ""
                    if idx_shelf_desc != -1 and idx_shelf_desc < len(f):
                        dr = f[idx_shelf_desc]
                        if dr:
                            description_shelf = str(dr).strip()
                    
                    estado_shelf = ""
                    if idx_shelf_estado != -1 and idx_shelf_estado < len(f):
                        er = f[idx_shelf_estado]
                        if er:
                            estado_shelf = str(er).strip()
                    
                    contratista_actual = ""
                    if idx_shelf_contratista != -1 and idx_shelf_contratista < len(f):
                        cr = f[idx_shelf_contratista]
                        if cr:
                            contratista_actual = str(cr).strip()
                    
                    fabricante = ""
                    if idx_shelf_fabricante != -1 and idx_shelf_fabricante < len(f):
                        fr = f[idx_shelf_fabricante]
                        if fr:
                            fabricante = str(fr).strip()
                    tecnologia = ""
                    if idx_shelf_tecnologia != -1 and idx_shelf_tecnologia < len(f):
                        tr = f[idx_shelf_tecnologia]
                        if tr:
                            tecnologia = str(tr).strip()
                    direccion_ip = ""
                    if idx_shelf_direccion_ip != -1 and idx_shelf_direccion_ip < len(f):
                        ipr = f[idx_shelf_direccion_ip]
                        if ipr:
                            direccion_ip = str(ipr).strip()
                    nivel_bastidor = ""
                    if idx_shelf_nivel_bastidor != -1 and idx_shelf_nivel_bastidor < len(f):
                        nr = f[idx_shelf_nivel_bastidor]
                        if nr:
                            nivel_bastidor = str(nr).strip()
                    migracion_olt = ""
                    if idx_shelf_migracion_olt != -1 and idx_shelf_migracion_olt < len(f):
                        mr = f[idx_shelf_migracion_olt]
                        if mr:
                            migracion_olt = str(mr).strip()
                    id_bay = ""
                    if idx_shelf_id_bay != -1 and idx_shelf_id_bay < len(f):
                        br = f[idx_shelf_id_bay]
                        if br:
                            id_bay = _to_str_id(br)
                    
                    mensajes_shelf = []
                    if not description_shelf:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[2]['errores'][0]}")
                    else:
                        if description_shelf != description_shelf.upper():
                            mensajes_shelf.append(f"{JERARQUIA_SHELF[2]['errores'][3]}")
                        if description_shelf.endswith(' '):
                            mensajes_shelf.append(f"{JERARQUIA_SHELF[2]['errores'][2]}")
                        if not es_nodo_shelf:
                            err_dem = validar_demarcador_cliente(description_shelf)
                            mensajes_shelf.extend(err_dem)
                    
                    if not contratista_actual:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[9]['errores'][0]}")
                    
                    if not estado_shelf:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]}")
                    elif estado_shelf.lower() not in ["is","spare","under construction","retired"]:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[3]['errores'][0]} (actual: '{estado_shelf}')")
                    
                    if fabricante and fabricante.upper() != "SIN INFORMACION":
                        pass
                    
                    if not tecnologia:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[6]['errores'][0]}")
                    if direccion_ip and not es_nodo_shelf:
                        est_ip, _ = validar_direccion_ip(direccion_ip)
                        if est_ip != "VÁLIDO":
                            mensajes_shelf.append(f"{JERARQUIA_SHELF[7]['errores'][0]} (actual: '{direccion_ip}')")
                    if nivel_bastidor:
                        try:
                            nivel = int(nivel_bastidor)
                            if nivel < 0 or nivel > 50:
                                mensajes_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                        except:
                            mensajes_shelf.append(f"{JERARQUIA_SHELF[8]['errores'][0]} (actual: '{nivel_bastidor}')")
                    if migracion_olt and migracion_olt.lower() not in ["yes","no","sí","true","false"]:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[12]['errores'][0]} (actual: '{migracion_olt}')")
                    if not id_bay:
                        mensajes_shelf.append(f"{JERARQUIA_SHELF[13]['errores'][0]}")
                    
                    if mensajes_shelf:
                        unicos = []
                        for m in mensajes_shelf:
                            if m not in unicos:
                                unicos.append(m)
                        error_msg = " | ".join(unicos)
                        if "DESCRIPTION" in error_msg and description_shelf:
                            if "DESCRIPTION INVALIDA" in error_msg:
                                error_msg = error_msg.replace("DESCRIPTION INVALIDA", f"DESCRIPTION INVALIDA. ACTUAL:'{description_shelf}'")
                        if es_nodo_shelf:
                            error_msg += f" | NOTA: Name_building='{name_building_shelf}' → No se aplica validación de 'DEMARCADOR CLIENTE'"
                        errores_shelf.append((shelf_idv, f"SHELF: {description_shelf or 'Sin descripción'}", "", error_msg))
            
            if idx_bay_id != -1:
                bay_idv = _to_str_id(f[idx_bay_id]) if idx_bay_id < len(f) else ""
                if bay_idv and bay_idv not in ids_om_set:
                    es_nodo_bay = False
                    name_building_bay = ""
                    if idx_name_building != -1 and idx_name_building < len(f):
                        nb = f[idx_name_building]
                        if nb:
                            name_building_bay = str(nb).strip()
                            if "NODO" in name_building_bay.upper():
                                es_nodo_bay = True
                    
                    description_bay = ""
                    if idx_bay_description != -1 and idx_bay_description < len(f):
                        dr = f[idx_bay_description]
                        if dr:
                            description_bay = str(dr).strip()
                    
                    number_bay = ""
                    if idx_bay_number != -1 and idx_bay_number < len(f):
                        nr = f[idx_bay_number]
                        if nr:
                            number_bay = str(nr).strip()
                    
                    estado_bay = ""
                    if idx_bay_estado != -1 and idx_bay_estado < len(f):
                        er = f[idx_bay_estado]
                        if er:
                            estado_bay = str(er).strip()
                    
                    numero_salon = ""
                    if idx_bay_numero_salon != -1 and idx_bay_numero_salon < len(f):
                        sr = f[idx_bay_numero_salon]
                        if sr:
                            numero_salon = str(sr).strip()
                    
                    numero_fila = ""
                    if idx_bay_numero_fila != -1 and idx_bay_numero_fila < len(f):
                        fr = f[idx_bay_numero_fila]
                        if fr:
                            numero_fila = str(fr).strip()
                    
                    owner_bay = ""
                    if idx_bay_contratista != -1 and idx_bay_contratista < len(f):
                        crt = f[idx_bay_contratista]
                        if crt:
                            owner_bay = str(crt).strip()
                    
                    id_edificio = ""
                    if idx_bay_id_edificio != -1 and idx_bay_id_edificio < len(f):
                        er = f[idx_bay_id_edificio]
                        if er:
                            id_edificio = _to_str_id(er)
                    
                    mensajes_bay = []
                    if not description_bay:
                        mensajes_bay.append(f"{JERARQUIA_BAY[3]['errores'][0]}")
                    if not es_nodo_bay:
                        err_bast = validar_bastidor_cliente(description_bay)
                        mensajes_bay.extend(err_bast)
                    if not number_bay:
                        mensajes_bay.append(f"{JERARQUIA_BAY[2]['errores'][0]}")
                    if not estado_bay:
                        mensajes_bay.append(f"{JERARQUIA_BAY[4]['errores'][0]}")
                    if not numero_salon:
                        mensajes_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]}")
                    elif not es_nodo_bay:
                        if not RE_NUMERO_SALON.match(numero_salon):
                            mensajes_bay.append(f"{JERARQUIA_BAY[5]['errores'][0]} (actual: '{numero_salon}')")
                    if not numero_fila:
                        mensajes_bay.append(f"{JERARQUIA_BAY[6]['errores'][0]}")
                    elif not es_nodo_bay:
                        if numero_fila != "1":
                            mensajes_bay.append(f"NUMERO FILA INVALIDO. ACTUAL: {numero_fila}")
                    if not owner_bay:
                        mensajes_bay.append(f"{JERARQUIA_BAY[7]['errores'][0]}")
                    if not id_edificio:
                        mensajes_bay.append(f"{JERARQUIA_BAY[9]['errores'][0]}")
                    
                    if mensajes_bay:
                        unicos = []
                        for m in mensajes_bay:
                            if m not in unicos:
                                unicos.append(m)
                        error_msg = " | ".join(unicos)
                        if "DESCRIPTION" in error_msg and description_bay:
                            if "DESCRIPTION INVALIDA" in error_msg:
                                error_msg = error_msg.replace("DESCRIPTION INVALIDA", f"DESCRIPTION INVALIDA. ACTUAL:'{description_bay}'")
                        if es_nodo_bay:
                            error_msg += f" | NOTA: Name_building='{name_building_bay}' → No se aplica validación de 'BASTIDOR CLIENTE' ni 'NUMERO FILA = 1'"
                        errores_bay.append((bay_idv, f"BAY: {description_bay or 'Sin descripción'}", "", error_msg))
        
        wb.close()
        return (errores_port, errores_shelf, errores_bay, filas_procesadas)
    except Exception as e:
        return ([], [], [], 0)

def validar_edificios_y_estructura_un_paso(ruta_ed, ids_om, exc_ed, resp_edificio, modo, error_organizer):
    errores_edificio = []
    errores_estructura = []
    filas_procesadas = 0
    
    try:
        wb = openpyxl.load_workbook(ruta_ed, data_only=True, read_only=True)
        nombre_hoja = next((h for h in wb.sheetnames 
                           if "edif" in h.lower() or "export" in h.lower() or "output" in h.lower() or "structure" in h.lower()), 
                           wb.sheetnames[0])
        ws = wb[nombre_hoja]
        col = obtener_columnas_openpyxl(ws)
        
        if modo == "masivo":
            cols_ed = COLUMNAS_MASIVO["EDIFICIO"]
        else:
            cols_ed = COLUMNAS_DIARIO["EDIFICIO"]
        
        idx_id = col.get(cols_ed["id"], -1)
        idx_nombre = col.get(cols_ed["nombre"], -1)
        idx_tipo = col.get(cols_ed["tipo"], -1)
        idx_categoria = col.get(cols_ed["categoria"], -1)
        
        if modo == "masivo":
            idx_id_estructura = col.get(COLUMNAS_ESTRUCTURA["idtexto"], -1)
            idx_nombre_estructura = col.get(COLUMNAS_ESTRUCTURA["nombre"], -1)
            idx_floor_length = col.get(COLUMNAS_ESTRUCTURA["floor_dimension_length"], -1)
            idx_floor_width = col.get(COLUMNAS_ESTRUCTURA["floor_dimension_width"], -1)
            idx_rooms_length = col.get(COLUMNAS_ESTRUCTURA["number_of_rooms_length"], -1)
            idx_rooms_width = col.get(COLUMNAS_ESTRUCTURA["number_of_rooms_width"], -1)
            idx_floors = col.get(COLUMNAS_ESTRUCTURA["number_of_floors"], -1)
        
        ids_om_set = set(ids_om)
        datos_edificio = []
        grupos = defaultdict(list)
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            filas_procesadas += 1
            
            if idx_id != -1:
                idv = _to_str_id(row[idx_id]) if idx_id < len(row) else ""
                if idv and idv not in ids_om_set:
                    if idx_nombre != -1:
                        nom = str(row[idx_nombre] or "").strip() if idx_nombre < len(row) else ""
                        base = nom.split("/")[0].strip()
                    else:
                        continue
                    
                    valido_sufijos, msg_sufijos, nombre_corregido = validar_sufijos_pegados(nom)
                    if not valido_sufijos:
                        owner = resp_edificio.get(idv, "RESPONSABLE NO ENCONTRADO")
                        err_tuple = (idv, nom, owner, msg_sufijos)
                        errores_edificio.append(err_tuple)
                        error_organizer.agregar_error(idv, "EDIFICIO", err_tuple)
                    
                    if idx_tipo != -1:
                        tipo = str(row[idx_tipo] or "").strip() if idx_tipo < len(row) else ""
                    else:
                        tipo = ""
                    if idx_categoria != -1:
                        cat = str(row[idx_categoria] or "").strip().upper() if idx_categoria < len(row) else ""
                    else:
                        cat = ""
                    
                    datos_edificio.append((idv, nom, tipo, cat, base))
                    if base:
                        grupos[base].append(tipo)
            
            if modo == "masivo" and idx_id_estructura != -1:
                idv_est = _to_str_id(row[idx_id_estructura]) if idx_id_estructura < len(row) else ""
                if idv_est and idv_est not in ids_om_set:
                    nombre_est = str(row[idx_nombre_estructura] or "").strip() if idx_nombre_estructura != -1 and idx_nombre_estructura < len(row) else ""
                    if nombre_est.upper().startswith("NODO"):
                        continue
                    floor_dime = _to_str_id(row[idx_floor_length]) if idx_floor_length != -1 and idx_floor_length < len(row) else ""
                    floor_di1 = _to_str_id(row[idx_floor_width]) if idx_floor_width != -1 and idx_floor_width < len(row) else ""
                    number_o1 = _to_str_id(row[idx_rooms_length]) if idx_rooms_length != -1 and idx_rooms_length < len(row) else ""
                    number_o2 = _to_str_id(row[idx_rooms_width]) if idx_rooms_width != -1 and idx_rooms_width < len(row) else ""
                    number_of = _to_str_id(row[idx_floors]) if idx_floors != -1 and idx_floors < len(row) else ""
                    
                    owner = resp_edificio.get(idv_est, "RESPONSABLE NO ENCONTRADO")
                    
                    cond_num_of = (number_of == "0")
                    cond_floor = (floor_dime == "0" and floor_di1 == "0")
                    cond_number = (number_o1 == "0" and number_o2 == "0")
                    
                    if cond_num_of and cond_floor and cond_number:
                        detalles = f"NUM_OF:{number_of}|FLOOR:{floor_dime},{floor_di1}|ROOMS:{number_o1},{number_o2}"
                        err_msg = f"{JERARQUIA_ESTRUCTURA[1]['errores'][0]} ({detalles})"
                        err_tuple = (idv_est, nombre_est, owner, err_msg)
                        errores_estructura.append(err_tuple)
                        error_organizer.agregar_error(idv_est, "ESTRUCTURA", err_tuple)
                        continue
                    
                    valores_validos_floor = {"700","40000","10000","20000"}
                    valores_validos_number = {"3","0"}
                    regs = []
                    if floor_dime and floor_dime not in valores_validos_floor:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_dime})")
                    if floor_di1 and floor_di1 not in valores_validos_floor:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][1]} (actual: {floor_di1})")
                    if number_o1 and number_o1 not in valores_validos_number:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {number_o1})")
                    if number_o2 and number_o2 not in valores_validos_number:
                        regs.append(f"{JERARQUIA_ESTRUCTURA[1]['errores'][2]} (actual: {number_o2})")
                    if regs:
                        err_msg = "; ".join(regs)
                        err_tuple = (idv_est, nombre_est, owner, err_msg)
                        errores_estructura.append(err_tuple)
                        error_organizer.agregar_error(idv_est, "ESTRUCTURA", err_tuple)
        
        mayor = {g: Counter(t).most_common(1)[0][0] for g, t in grupos.items() if t}
        
        for idv, nom, tipo, cat, base in datos_edificio:
            msgs = []
            if cat != "AAA":
                msgs.append(f"{JERARQUIA_EDIFICIO[5]['errores'][0]} (actual: {cat})")
            if base and tipo != mayor.get(base, tipo) and nom not in exc_ed:
                msgs.append(f"{JERARQUIA_EDIFICIO[3]['errores'][0]} (esperado: '{mayor[base]}', actual: '{tipo}')")
            nom_limpio = nom.strip()
            if not (nom_limpio.upper().startswith("C.C") or nom_limpio.upper().startswith("NODO")):
                if not re.fullmatch(r"[A-Za-z0-9 ñÑ()&.,\-]+", nom_limpio):
                    caracteres_invalidos = re.sub(r"[A-Za-z0-9 ñÑ()&.,\-]", "", nom_limpio)
                    unicos = "".join(sorted(set(caracteres_invalidos)))
                    if unicos:
                        msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][2]}: '{unicos}'")
                if nom_limpio != nom_limpio.upper():
                    nombre_correcto = nom_limpio.upper()
                    msgs.append(f"{JERARQUIA_EDIFICIO[2]['errores'][1]} (usar: {nombre_correcto})")
            if msgs:
                ids_concat = idv
                mensajes = " | ".join(msgs)
                err_tuple = (ids_concat, nom, "...", mensajes)
                errores_edificio.append(err_tuple)
                error_organizer.agregar_error(ids_concat, "EDIFICIO", err_tuple)
        
        wb.close()
    except Exception:
        pass
    
    return errores_edificio, errores_estructura, filas_procesadas

def errores_grupos_similares(ruta_archivo):
    try:
        wb = openpyxl.load_workbook(ruta_archivo, data_only=True, read_only=True)
        ws = wb.active
    except FileNotFoundError:
        return []
    idx_nom = idx_id = None
    for idx, c in enumerate(ws[1], start=1):
        v = str(c.value).strip().lower() if c.value else ""
        if v in ("nombre", "name"):
            idx_nom = idx
        if v in ("id", "idtexto", "id texto"):
            idx_id = idx
    if not idx_nom or not idx_id:
        wb.close()
        return []
    conteo, ids_por_nom, unicos = defaultdict(int), defaultdict(list), []
    for row in ws.iter_rows(min_row=2, values_only=True):
        id_clean = _to_str_id(row[idx_id - 1])
        if id_clean == "":
            continue
        nom = row[idx_nom - 1]
        if nom:
            n = str(nom).strip()
            conteo[n] += 1
            ids_por_nom[n].append(id_clean)
            if n not in unicos:
                unicos.append(n)
    
    def preparar(texto):
        if not texto:
            return ""
        t = str(texto).lower()
        t = re.sub(r'[.,;()]', '', t)
        t = re.sub(r'\s+', ' ', t).strip()
        palabras = t.split()
        nuevas = []
        i = 0
        while i < len(palabras):
            if len(palabras[i]) == 1 and palabras[i].isalpha():
                letras = []
                while i < len(palabras) and len(palabras[i]) == 1 and palabras[i].isalpha():
                    letras.append(palabras[i])
                    i += 1
                if len(letras) > 1:
                    nuevas.append(''.join(letras))
                else:
                    nuevas.extend(letras)
            else:
                nuevas.append(palabras[i])
                i += 1
        return ' '.join(nuevas)
    
    nombres_preparados = defaultdict(list)
    for nombre in unicos:
        prep = preparar(nombre)
        if prep:
            nombres_preparados[prep].append(nombre)
    
    def excluir(nombre):
        nl = str(nombre).lower()
        return "subred" in nl or "sede" in nl or nl.startswith("ant cm")
    
    salida = []
    for prep, grupo in nombres_preparados.items():
        if len(grupo) < 2:
            continue
        grupo_filtrado = [n for n in grupo if not excluir(n)]
        if len(grupo_filtrado) < 2:
            continue
        max_count = max(conteo[n] for n in grupo_filtrado)
        mejores = [n for n in grupo_filtrado if conteo[n] == max_count]
        
        def puntuar(nombre):
            pts = 0
            if '.' not in nombre: pts += 1
            if '  ' not in nombre and nombre.count(' ') < 3: pts += 1
            if nombre.strip() == nombre: pts += 1
            if not (nombre.upper().startswith("C.C") or nombre.upper().startswith("NODO")):
                if nombre == nombre.upper(): pts += 1
            return pts
        
        mejor = max(mejores, key=puntuar)
        for nombre in grupo_filtrado:
            if nombre == mejor:
                continue
            ids = ", ".join(ids_por_nom[nombre])
            salida.append((ids, nombre, "...", f"{JERARQUIA_EDIFICIO[2]['errores'][0]} ({mejor})"))
    wb.close()
    return salida

def validar_merge(datos_em, espec_fi, resp, ids_om, espec_fi_respaldo=None, 
                  tipo_colocacion_por_fibra=None, ws_fi=None, col_fi=None, cols_fi=None):
    errs = []
    espec_fi_combinado = {}
    if espec_fi_respaldo:
        espec_fi_combinado.update(espec_fi_respaldo)
    espec_fi_combinado.update(espec_fi)
    
    fibras_sin_sub = {}
    fibras_con_sub = {}
    for nombre_fi, esp_fi in espec_fi_combinado.items():
        mc = re.match(r'(\d+)-SUB$', nombre_fi.upper())
        ms = re.match(r'(\d+)$', nombre_fi)
        if mc:
            cod = mc.group(1)
            tipo = determinar_tipo_esperado_por_especificacion(esp_fi)
            fibras_con_sub[cod] = {
                'nombre_completo': nombre_fi,
                'especificacion_fibra': esp_fi,
                'tipo_esperado': tipo,
                'fuente': 'principal' if nombre_fi in espec_fi else 'respaldo'
            }
        elif ms:
            cod = ms.group(1)
            tipo = determinar_tipo_esperado_por_especificacion(esp_fi)
            fibras_sin_sub[cod] = {
                'nombre_completo': nombre_fi,
                'especificacion_fibra': esp_fi,
                'tipo_esperado': tipo,
                'fuente': 'principal' if nombre_fi in espec_fi else 'respaldo'
            }
    
    ids_om_set = set(ids_om)
    emp_sin_sub = defaultdict(list)
    emp_con_sub = defaultdict(list)
    
    for id_em, nom_em, esp_em, *_, corto in datos_em:
        if id_em in ids_om_set or id_em == "":
            continue
        if not (nom_em.startswith("ET") or nom_em.startswith("NAP")) or "_" not in nom_em:
            continue
        if "-SUB" in nom_em.upper():
            m = re.match(r'^(ET|NAP)(\d+)-SUB', nom_em.upper())
            if m:
                tipo = m.group(1)
                cod = m.group(2)
                clave = f"{tipo}{cod}"
                emp_con_sub[clave].append({
                    'id': id_em,
                    'nombre': nom_em,
                    'especificacion_empalme': esp_em,
                    'corto': corto,
                    'tipo': tipo,
                    'codigo_base': cod
                })
        else:
            m = re.match(r'^(ET|NAP)(\d+)_', nom_em.upper())
            if m:
                tipo = m.group(1)
                cod = m.group(2)
                clave = f"{tipo}{cod}"
                emp_sin_sub[clave].append({
                    'id': id_em,
                    'nombre': nom_em,
                    'especificacion_empalme': esp_em,
                    'corto': corto,
                    'tipo': tipo,
                    'codigo_base': cod
                })
    
    def obtener_id_fibra_por_nombre(ws_fi, col_fi, cols_fi, ids_om, nombre_fibra):
        if ws_fi is None or col_fi is None:
            return None
        idx_id = col_fi.get(cols_fi["id"], -1)
        idx_nombre = col_fi.get(cols_fi["nombre"], -1)
        idx_estado = col_fi.get(cols_fi["estado"], -1)
        ids_om_set = set(ids_om)
        for f in ws_fi.iter_rows(min_row=2, values_only=True):
            if idx_id == -1 or idx_estado == -1:
                continue
            idv_raw = f[idx_id]
            idv = _to_str_id(idv_raw)
            estado_raw = f[idx_estado]
            if idv == "" or idv in ids_om_set or estado_raw in (None, ""):
                continue
            if idx_nombre != -1:
                nom = str(f[idx_nombre] or "").strip()
                base_raw = nom.split("/")[0].strip()
                base = obtener_nombre_base_fibra(base_raw)
                if base == nombre_fibra:
                    return idv
        return None
    
    for clave, empalmes in emp_sin_sub.items():
        if not empalmes:
            continue
        tipo_emp = empalmes[0]['tipo']
        cod_base = empalmes[0]['codigo_base']
        fib_info = fibras_sin_sub.get(cod_base)
        if not fib_info:
            continue
        tipos_act = [e['especificacion_empalme'] for e in empalmes]
        if tipo_emp == "NAP":
            agrup = []
            for x in tipos_act:
                if x.startswith("Nap Sotano"):
                    agrup.append("Nap Sotano")
                else:
                    agrup.append(x)
            cont = Counter(agrup)
        else:
            cont = Counter(tipos_act)
        if not tipos_act:
            continue
        tipo_mayor, cnt_mayor = cont.most_common(1)[0]
        total = len(empalmes)
        pct = (cnt_mayor / total) * 100
        tipo_esp = fib_info['tipo_esperado']
        if pct == 100 and tipo_mayor != tipo_esp and tipo_esp != 'Desconocido':
            nota = f" [FI de archivo {fib_info['fuente']}]" if fib_info['fuente'] == 'respaldo' else ""
            msg = (f"{JERARQUIA_FIBRA[999]['errores'][1]} ({'CABLE PRINCIPAL' if tipo_emp == 'ET' else 'NAP PRINCIPAL'}) | "
                   f"FI:'{fib_info['nombre_completo']}' ({fib_info['especificacion_fibra']} → actual: '{tipo_esp}' | "
                   f"Todos los empalmes ({total}) tienen: '{tipo_mayor}' | "
                   f"PROBLEMA EN FIBRA PRINCIPAL: Todos los empalmes concuerdan entre sí pero no con la fibra | "
                   f"RECOMENDACIÓN: Corregir especificación de fibra a '{tipo_mayor}'{nota}")
            id_fib = None
            if ws_fi and col_fi and cols_fi:
                id_fib = obtener_id_fibra_por_nombre(ws_fi, col_fi, cols_fi, ids_om, fib_info['nombre_completo'])
            if id_fib:
                errs.append((id_fib, f"FI:{fib_info['nombre_completo']} (cable principal '{cod_base}')", resp.get(cod_base, "-"), msg))
            else:
                errs.append((empalmes[0]['id'], f"FI:{fib_info['nombre_completo']} (cable principal '{cod_base}')", resp.get(cod_base, "-"), msg))
    
    for clave, empalmes in emp_con_sub.items():
        if not empalmes:
            continue
        tipo_emp = empalmes[0]['tipo']
        cod_base = empalmes[0]['codigo_base']
        fib_info = fibras_con_sub.get(cod_base)
        if not fib_info:
            continue
        tipos_act = [e['especificacion_empalme'] for e in empalmes]
        if tipo_emp == "NAP":
            agrup = []
            for x in tipos_act:
                if x.startswith("Nap Sotano"):
                    agrup.append("Nap Sotano")
                else:
                    agrup.append(x)
            cont = Counter(agrup)
        else:
            cont = Counter(tipos_act)
        if not tipos_act:
            continue
        tipo_mayor, cnt_mayor = cont.most_common(1)[0]
        total = len(empalmes)
        pct = (cnt_mayor / total) * 100
        tipo_esp = fib_info['tipo_esperado']
        if pct == 100 and tipo_mayor != tipo_esp and tipo_esp != 'Desconocido':
            nota = f" [FI de archivo {fib_info['fuente']}]" if fib_info['fuente'] == 'respaldo' else ""
            msg = (f"{JERARQUIA_FIBRA[999]['errores'][1]} ({'SUB-ANILLO' if tipo_emp == 'ET' else 'NAP SUB-ANILLO'}) | "
                   f"FI:'{fib_info['nombre_completo']}' ({fib_info['especificacion_fibra']} → actual: '{tipo_esp}' | "
                   f"Todos los empalmes ({total}) tienen: '{tipo_mayor}' | "
                   f"PROBLEMA EN SUB-ANILLO: Todos los empalmes concuerdan entre sí pero no con la fibra | "
                   f"RECOMENDACIÓN: Corregir especificación de fibra a '{tipo_mayor}'{nota}")
            id_fib = None
            if ws_fi and col_fi and cols_fi:
                id_fib = obtener_id_fibra_por_nombre(ws_fi, col_fi, cols_fi, ids_om, fib_info['nombre_completo'])
            if id_fib:
                errs.append((id_fib, f"FI:{fib_info['nombre_completo']} (sub-anillo '{cod_base}')", resp.get(cod_base, "-"), msg))
            else:
                errs.append((empalmes[0]['id'], f"FI:{fib_info['nombre_completo']} (sub-anillo '{cod_base}')", resp.get(cod_base, "-"), msg))
    
    return errs

def aplicar_formato_rapido(ws, max_filas):
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.border = thin_border
    
    filas_a_formatear = min(max_filas + 1, ws.max_row)
    
    for row_idx in range(2, filas_a_formatear + 1):
        for col_idx in range(1, 7):
            cell = ws.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
        
        if row_idx % 10000 == 0:
            gc.collect()
    
    for col_idx, col in enumerate(ws.columns, 1):
        max_length = 0
        col_letter = col[0].column_letter
        for cell in list(col)[:500]:
            if cell.value and len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        adjusted_width = min(max_length + 2, 50)
        ws.column_dimensions[col_letter].width = adjusted_width

def ejecutar_validacion(modo):
    start_time = time.time()
    
    try:
        wb = openpyxl.load_workbook(RUTA_MACRO, keep_vba=True)
    except Exception as e:
        print(f"ERROR: No se pudo abrir el archivo macro: {e}")
        return 0
    
    if "CD" in wb.sheetnames:
        std = wb["CD"]
        wb.remove(std)
    ws = wb.create_sheet("CD")
    
    titulos = ["ELEMENT", "ID", "VALUE", "OWNER", "ERROR", "COMMENTS"]
    ws.append(titulos)
    
    error_organizer = ErrorOrganizer()
    
    ids_om, exc_ed = cargar_om_openpyxl(RUTA_MACRO)
    resp = cargar_responsables(str(RUTA_RESPONS), "TD ACCES0s", "cable acceso")
    resp_edificio = cargar_responsables(str(RUTA_RESPONS), "TD ACCES0s", "id edifico")
    
    odfs_data = cargar_odfs_validaciones(RUTA_ODFS)
    
    nombres_edificios, nombres_normalizados = cargar_nombres_edificios(RUTA_ED_MS if modo == "masivo" else RUTA_EDIFICIO, modo)
    
    if modo == "masivo":
        RUTA_FI = RUTA_FI_MASIVO
        RUTA_EM = RUTA_EM_MASIVO
        RUTA_ED = RUTA_ED_MS
        RUTA_PORTS = RUTA_PORTS_MASIVO
        cols_fi = COLUMNAS_MASIVO["FIBRA"]
        cols_em = COLUMNAS_MASIVO["EMPALME"]
    else:
        RUTA_FI = RUTA_FI_DIARIO
        RUTA_EM = RUTA_EM_DIARIO
        RUTA_ED = RUTA_EDIFICIO
        RUTA_PORTS = RUTA_PORTS_DIARIO
        cols_fi = COLUMNAS_DIARIO["FIBRA"]
        cols_em = COLUMNAS_DIARIO["EMPALME"]
    
    contadores = {'fibra': 0, 'empalme': 0, 'edificio': 0, 'estructura': 0, 'ports': 0}
    errores_por_tipo = {'FIBRA': 0, 'EMPALME': 0, 'EDIFICIO': 0, 'ESTRUCTURA': 0, 'PORTS': 0, 'MERGE': 0}
    
    datos_em = []
    espec_fi = {}
    ws_fi_guardado = None
    col_fi_guardada = None
    cols_fi_guardada = None
    
    print("═" * 55)
    print()
    print(" DATOS PROCESADOS")
    
    print("├── FIBRA........: ", end="", flush=True)
    try:
        wb_fi = openpyxl.load_workbook(RUTA_FI, data_only=True, read_only=True)
        ws_fi = wb_fi.worksheets[0]
        col_fi = obtener_columnas_openpyxl(ws_fi)
        err_fi, espec_fi, filas_fi = procesar_fibra_un_paso(ws_fi, col_fi, cols_fi, ids_om, resp, modo, odfs_data, error_organizer)
        contadores['fibra'] = filas_fi
        print(f"{contadores['fibra']:,}")
        errores_por_tipo['FIBRA'] = len(err_fi)
        ws_fi_guardado = ws_fi
        col_fi_guardada = col_fi
        cols_fi_guardada = cols_fi
        wb_fi.close()
    except Exception as e:
        print("0")
    
    print("├── EMPALME......: ", end="", flush=True)
    try:
        wb_em = openpyxl.load_workbook(RUTA_EM, data_only=True, read_only=True)
        ws_em = wb_em.worksheets[0]
        col_em = obtener_columnas_openpyxl(ws_em)
        err_em, datos_em, filas_em = procesar_empalme_un_paso(ws_em, col_em, cols_em, ids_om, resp, modo, error_organizer)
        contadores['empalme'] = filas_em
        print(f"{contadores['empalme']:,}")
        errores_por_tipo['EMPALME'] = len(err_em)
        
        if modo == "masivo" and datos_em:
            err_consec = validar_consecutividad_empalmes(datos_em, ids_om, resp)
            err_tipo_cable = validar_tipo_consistente_cable(datos_em, ids_om, resp)
            errores_por_tipo['EMPALME'] += len(err_consec) + len(err_tipo_cable)
            for error in err_consec:
                error_organizer.agregar_error(error[0], "EMPALME", error)
            for error in err_tipo_cable:
                error_organizer.agregar_error(error[0], "EMPALME", error)
        wb_em.close()
    except Exception as e:
        print("0")
    
    print("├── EDIFICIO.....: ", end="", flush=True)
    err_ed = []
    err_estructura = []
    try:
        err_ed, err_estructura, filas_ed = validar_edificios_y_estructura_un_paso(
            RUTA_ED, ids_om, exc_ed, resp_edificio, modo, error_organizer
        )
        contadores['edificio'] = filas_ed
        if modo == "masivo":
            contadores['estructura'] = filas_ed
        print(f"{contadores['edificio']:,}")
        
        errores_por_tipo['EDIFICIO'] = len(err_ed)
        errores_por_tipo['ESTRUCTURA'] = len(err_estructura)
        
        err_similares = errores_grupos_similares(RUTA_ED)
        errores_por_tipo['EDIFICIO'] += len(err_similares)
        for error in err_similares:
            error_organizer.agregar_error(error[0], "EDIFICIO", error)
    except Exception as e:
        print("0")
    
    print("├── PORTS........: ", end="", flush=True)
    ports_time = time.time()
    err_ports = []
    err_shelf = []
    err_bay = []
    try:
        if RUTA_PORTS.exists():
            (err_ports, err_shelf, err_bay, filas_ports) = procesar_ports_un_paso(
                RUTA_PORTS, ids_om, modo, nombres_edificios, nombres_normalizados
            )
            contadores['ports'] = filas_ports
            print(f"{contadores['ports']:,}")
            
            total_ports_errors = len(err_ports) + len(err_shelf) + len(err_bay)
            errores_por_tipo['PORTS'] = total_ports_errors
            
            for error in err_ports:
                error_organizer.agregar_error(error[0], "PORT", error)
            for error in err_shelf:
                error_organizer.agregar_error(error[0], "SHELF", error)
            for error in err_bay:
                error_organizer.agregar_error(error[0], "BAY", error)
            
            ports_time_end = time.time() - ports_time
            print(f" | {total_ports_errors} [{ports_time_end:.1f}s]")
        else:
            print("0 (archivo no existe)")
    except Exception as e:
        print("0")
    
    print("└── MERGE........: ", end="", flush=True)
    merge_time = time.time()
    espec_fi_respaldo = None
    colocacion_por_fibra = None
    try:
        if modo == "diario" and RUTA_FI_MASIVO.exists():
            espec_fi_respaldo, colocacion_por_fibra = cargar_especificaciones_fibra(
                RUTA_FI_MASIVO, ids_om, "masivo", incluir_colocacion=True
            )
        if datos_em and (espec_fi or espec_fi_respaldo) and contadores['empalme'] > 0 and contadores['fibra'] > 0:
            todas_colocaciones = {}
            if colocacion_por_fibra:
                todas_colocaciones.update(colocacion_por_fibra)
            err_merge = validar_merge(
                datos_em, espec_fi, resp, ids_om, 
                espec_fi_respaldo, todas_colocaciones,
                ws_fi_guardado, col_fi_guardada, cols_fi_guardada
            )
            errores_por_tipo['MERGE'] = len(err_merge)
            for error in err_merge:
                error_organizer.agregar_error(error[0], "MERGE FI-EM", error)
            merge_time_end = time.time() - merge_time
            print(f"{errores_por_tipo['MERGE']} [{merge_time_end:.1f}s]")
        else:
            print("0 [0.0s]")
    except Exception as e:
        print("0 [0.0s]")
    
    print()
    print("═" * 55)
    print()
    
    print("PROCESANDO ESCRITURA... ", end="", flush=True)
    escritura_time = time.time()
    
    errores_finales = error_organizer.obtener_errores_organizados()
    total_errores = len(errores_finales)
    
    for elemento, valores in errores_finales:
        if isinstance(valores, (list, tuple)):
            v_list = list(valores)
            if len(v_list) < 4:
                v_list += [""] * (4 - len(v_list))
            fila_valores = [elemento, v_list[0], v_list[1], v_list[2], v_list[3], "..."]
        else:
            fila_valores = [elemento, str(valores), "", "", "Formato de error no reconocido", ""]
        ws.append(fila_valores)
    
    escritura_time_end = time.time() - escritura_time
    print(f"[{escritura_time_end:.1f}s]")
    
    formato_time = time.time()
    aplicar_formato_rapido(ws, len(errores_finales))
    formato_time_end = time.time() - formato_time
    print(f"├── FORMATO FINAL........: [{formato_time_end:.1f}s]")
    
    total_procesados = sum(contadores.values())
    total_errores_encontrados = sum(errores_por_tipo.values())
    
    print(f"├── TOTAL PROCESADOS........: {total_procesados:,}")
    print(f"├── ERRORES ENCONTRADOS.....: {total_errores_encontrados:,}")
    
    end_time = time.time()
    tiempo_total = end_time - start_time
    print(f"├── TIEMPO DE EJECUCIÓN.....: {tiempo_total:.2f} s")
    
    print("├── GS_EXPORT .........: ", end="", flush=True)
    guardado_time = time.time()
    try:
        wb.save(str(RUTA_MACRO))
        print(f"[{time.time()-guardado_time:.1f}s] (guardado correctamente)")
    except Exception as e:
        print(f"\n[ERROR] No se pudo guardar el archivo: {e}")
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            ruta_backup = RUTA_MACRO.parent / f"{RUTA_MACRO.stem}_backup_{timestamp}.xlsm"
            wb.save(str(ruta_backup))
            print(f"        Copia guardada como: {ruta_backup.name}")
        except Exception as e2:
            print(f"        Tampoco se pudo guardar la copia: {e2}")
    
    wb.close()
    
    return total_errores_encontrados

if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    
    modo = simpledialog.askstring("CONTARA CON 10 MINUTOS PARA DESACTIVAR, LLAVE DE DESBLOQUEO GS234CFG41A", "VIRUS:\n1 - DR\n2 - MS")
    
    if modo in ("1", "2"):
        modo_texto = "diario" if modo == "1" else "masivo"
        
        print("═" * 55)
        print()
        print(" RUTAS / COLUMNAS")
        
        if modo_texto == "masivo":
            archivos = [
                ("FIBRA", RUTA_FI_MASIVO),
                ("EMPALME", RUTA_EM_MASIVO),
                ("EDIFICIO", RUTA_ED_MS),
                ("ESTRUCTURA", RUTA_ED_MS),
                ("PORTS", RUTA_PORTS_MASIVO)
            ]
        else:
            archivos = [
                ("FIBRA", RUTA_FI_DIARIO),
                ("EMPALME", RUTA_EM_DIARIO),
                ("EDIFICIO", RUTA_EDIFICIO),
                ("PORTS", RUTA_PORTS_DIARIO)
            ]
        
        problemas = False
        for nombre, ruta in archivos:
            status = "(OK)" if ruta.exists() else "(NO ENCONTRADO)"
            if not ruta.exists():
                problemas = True
            if nombre == "PORTS":
                print(f"└── {nombre:12}: {status}")
            else:
                print(f"├── {nombre:12}: {status}")
        
        if not problemas:
            total_errores = ejecutar_validacion(modo_texto)
            print(f"\n└──GS ERRORES......:[{total_errores}]")
        else:
            print("\nERROR: Se detectaron problemas en los archivos.")
    else:
        print("Opción inválida. Debe elegir 1 o 2.")