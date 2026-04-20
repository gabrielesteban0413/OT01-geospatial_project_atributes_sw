# ─────────────────────────── IMPORTACIONES ───────────────────────────────
import os
import shutil
import win32com.client as win32
import ctypes  
from pathlib import Path
import traceback
from datetime import datetime

try:
    from procesar_dwg_arcgis import procesar_dwg
except ImportError:
    def procesar_dwg(ruta_dwg, carpeta_destino):
        print(f"[ADVERTENCIA] procesar_dwg no está implementado. Archivo: {ruta_dwg}")

# ─────────────────────────── CONFIGURACIÓN DE RUTAS ─────────────────────────
BASE_PATH = Path(r"C:\A_GS1_PROYECTOS")
RUTA_EXCEL_PRINCIPAL = BASE_PATH / "0_Documents_gs" / "docs" / "Manual_GIS.xlsm"
RUTA_CARPETA = BASE_PATH
RUTA_EXTRA_AUTOCAD = BASE_PATH / "INSUMOS" / "Autocad"

RUTAS_ORIGEN = {
    "c": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Transporte\Cartográfico\Bogotá"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Transporte\Lógico")
    ],
    "a": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Tendidos_ANILLOS"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Logicos_ANILLOS")
    ]
}

RUTA_CORPORATIVO = Path(r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Shape_ MCB")
RUTA_40_ANILLOS = RUTA_CORPORATIVO / "00_ANILLOS"
RUTA_40_TRONCALES = RUTA_CORPORATIVO / "00_TRONCALES"

MB_ICONERROR = 0x10
MSG_TIPO_BUSQUEDA_ERROR = "ERROR: Tipo de búsqueda incorrecto. Use 'A' o 'C'."

# ─────────────────────────── UTILIDADES ────────────────────────────────
def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def buscar_archivo_troncal(ruta_base, numero_cable):
    """
    Busca archivos de troncal con diferentes patrones:
    1. C.{numero}_TO-TO_T.dwg (ej: C.71248_TO-TO_T.dwg)
    2. C.{numero}_T.dwg (ej: C.71248_T.dwg)
    3. {numero}_T.dwg (ej: 71248_T.dwg)
    """
    patrones = [
        f"C.{numero_cable}_TO-TO_T.dwg",
        f"C.{numero_cable}_T.dwg",
        f"{numero_cable}_T.dwg",
        f"C.{numero_cable}_TO-TO_L.dwg",
        f"C.{numero_cable}_L.dwg",
        f"{numero_cable}_L.dwg"
    ]
    
    archivo_mas_reciente = None
    fecha_mas_reciente = 0
    
    try:
        for root, _, files in os.walk(str(ruta_base)):
            for nombre in files:
                nombre_lower = nombre.lower()
                # Verificar si coincide con alguno de los patrones
                for patron in patrones:
                    if patron.lower() in nombre_lower and nombre_lower.endswith(('.dwg', '.DWG')):
                        ruta_completa = Path(root) / nombre
                        try:
                            fecha_mod = ruta_completa.stat().st_mtime
                            if fecha_mod > fecha_mas_reciente:
                                fecha_mas_reciente = fecha_mod
                                archivo_mas_reciente = ruta_completa
                        except Exception:
                            continue
    except Exception as e:
        print(f"[{_now()}] [ERROR] Al recorrer la ruta '{ruta_base}': {e}")
        
    return archivo_mas_reciente

def buscar_carpeta_corporativo(ruta_base, numero_cable):
    """
    Busca carpetas en corporativo que contengan el número del cable
    Puede ser: C.{numero}_TO-TO, C.{numero}, o solo el número
    """
    patrones_carpeta = [
        f"C.{numero_cable}_TO-TO",
        f"C.{numero_cable}",
        numero_cable
    ]
    
    try:
        for item in ruta_base.iterdir():
            if item.is_dir():
                nombre_carpeta = item.name
                for patron in patrones_carpeta:
                    if patron in nombre_carpeta:
                        return item
    except Exception as e:
        print(f"[{_now()}] [ERROR] Buscando carpeta en corporativo: {e}")
    
    return None

def obtener_excel():
    try:
        xl = win32.GetActiveObject("Excel.Application")
    except Exception:
        try:
            xl = win32.Dispatch("Excel.Application")
            xl.Visible = True
        except Exception:
            print(f"[{_now()}] ERROR: No se pudo conectar ni crear una instancia de Excel.")
            return None, None

    try:
        wb = xl.Workbooks(RUTA_EXCEL_PRINCIPAL.name)
    except Exception:
        try:
            wb = xl.Workbooks.Open(str(RUTA_EXCEL_PRINCIPAL))
        except Exception:
            print(f"[{_now()}] ERROR: No se pudo encontrar ni abrir el archivo: {RUTA_EXCEL_PRINCIPAL}")
            return None, None
            
    return xl, wb

# ─────────────────────────── LÓGICA PRINCIPAL ────────────────────────────
def copiar_archivos_masivo():
    xl, wb = obtener_excel() 
    if not xl or not wb:
        ctypes.windll.user32.MessageBoxW(
            0, 
            f"No se pudo conectar al archivo Excel: {RUTA_EXCEL_PRINCIPAL.name}. Asegúrese de que existe y se puede acceder.", 
            "Error de Conexión", 
            MB_ICONERROR
        )
        return

    try:
        # Ahora seleccionamos explícitamente la hoja Frontend
        ws = wb.Sheets("Frontend")
        
        tipo_busqueda = str(xl.InputBox("Ingrese tipo de búsqueda (A o C):",
                                         "Tipo de Búsqueda", "A")).strip().lower()

        if tipo_busqueda not in ["a", "c"]:
            ctypes.windll.user32.MessageBoxW(0, MSG_TIPO_BUSQUEDA_ERROR, "Error", MB_ICONERROR)
            return

        RUTA_EXTRA_AUTOCAD.mkdir(parents=True, exist_ok=True)
        print(f"[{_now()}] [INFO] Carpeta temporal creada o existente: {RUTA_EXTRA_AUTOCAD}")

        # NUEVA UBICACIÓN INICIAL: Columna D (4), Fila 14
        fila = 14
        col_entrada = 4   # D
        col_resultado_t = 5  # E
        col_resultado_l = 6  # F
        col_ruta_t = 24  # Mantenemos columna X (24) para ruta T

        procesadas = 0
        sin_datos_consecutivos = 0
        max_sin_datos = 10  

        while sin_datos_consecutivos < max_sin_datos:
            valor_celda = ws.Cells(fila, col_entrada).Value
            
            if not valor_celda or str(valor_celda).strip().lower() in ("", "none"):
                sin_datos_consecutivos += 1
                fila += 1
                continue

            sin_datos_consecutivos = 0
            numero_cable = str(valor_celda).strip()
            print(f"[{_now()}] [PROCESANDO] Fila {fila}: Cable {numero_cable}")

            ruta_base_40 = RUTA_40_ANILLOS if tipo_busqueda == "a" else RUTA_40_TRONCALES
            
            # PRIMERO: Buscar carpeta en corporativo
            carpeta_corporativo = buscar_carpeta_corporativo(ruta_base_40, numero_cable)
            
            if carpeta_corporativo:
                ws.Cells(fila, col_resultado_t).Value = str(carpeta_corporativo)
                print(f"[{_now()}] [SKIP] Carpeta encontrada en Corporativo: {carpeta_corporativo}. Saltando.")
                fila += 1
                procesadas += 1
                continue

            # SEGUNDO: Buscar archivos en corporativo (para anillos)
            if tipo_busqueda == "a":
                encontrado_archivo = False
                # Buscar archivos con diferentes patrones
                for patron_t in [f"C.{numero_cable}_T.dwg", f"{numero_cable}_T.dwg"]:
                    ruta_t = ruta_base_40 / patron_t
                    if ruta_t.exists():
                        ws.Cells(fila, col_resultado_t).Value = str(ruta_t)
                        encontrado_archivo = True
                        break
                
                for patron_l in [f"C.{numero_cable}_L.dwg", f"{numero_cable}_L.dwg"]:
                    ruta_l = ruta_base_40 / patron_l
                    if ruta_l.exists():
                        ws.Cells(fila, col_resultado_l).Value = str(ruta_l) if ruta_l.exists() else "No"
                        encontrado_archivo = True
                        break
                
                if encontrado_archivo:
                    print(f"[{_now()}] [SKIP] Archivos encontrados en Corporativo 40. Saltando.")
                    fila += 1
                    procesadas += 1
                    continue

            # TERCERO: Buscar en Cardiseño
            rutas_cardiseno = RUTAS_ORIGEN[tipo_busqueda]
            
            archivo_t = buscar_archivo_troncal(rutas_cardiseno[0], numero_cable)
            archivo_l = buscar_archivo_troncal(rutas_cardiseno[1], numero_cable)
            
            # Filtrar para asegurarnos que T sea T y L sea L
            if archivo_t and "_L." in archivo_t.name:
                archivo_t = None
            if archivo_l and "_T." in archivo_l.name:
                archivo_l = None

            if not archivo_t and not archivo_l:
                print(f"[{_now()}] [ADVERTENCIA] No se encontraron archivos para cable {numero_cable} en Cardiseño.")
                ws.Cells(fila, col_resultado_t).Value = "No encontrado"
                ws.Cells(fila, col_resultado_l).Value = "No encontrado"
                fila += 1
                continue

            # Crear carpeta de destino usando el nombre del archivo encontrado
            if archivo_t:
                nombre_base = archivo_t.stem.replace("_T", "")
            elif archivo_l:
                nombre_base = archivo_l.stem.replace("_L", "")
            else:
                nombre_base = f"C.{numero_cable}"
            
            destino = RUTA_CARPETA / nombre_base
            destino.mkdir(parents=True, exist_ok=True)

            if archivo_t:
                nombre_archivo_t = archivo_t.name
                ruta_t_destino = destino / nombre_archivo_t
                
                try:
                    shutil.copy(archivo_t, ruta_t_destino)
                except Exception as e:
                    print(f"[{_now()}] [ERROR] Copiar T: {e}")
                
                try:
                    shutil.copy(archivo_t, RUTA_EXTRA_AUTOCAD)
                except Exception as e:
                    print(f"[{_now()}] [ADVERTENCIA] No se pudo copiar a carpeta extra: {e}")
                
                ws.Cells(fila, col_ruta_t).Value = str(ruta_t_destino)
                ws.Cells(fila, col_resultado_t).Value = nombre_archivo_t
                print(f"[{_now()}] [COPIADO T] {nombre_archivo_t}")

                try:
                    procesar_dwg(ruta_dwg=str(ruta_t_destino), carpeta_destino=str(destino))
                except Exception as e:
                    print(f"[{_now()}] [ERROR ArcGIS] Falló procesar_dwg para {ruta_t_destino}: {e}")
            else:
                ws.Cells(fila, col_resultado_t).Value = "No"

            if archivo_l:
                nombre_archivo_l = archivo_l.name
                ruta_l_destino = destino / nombre_archivo_l
                try:
                    shutil.copy(archivo_l, ruta_l_destino)
                    print(f"[{_now()}] [COPIADO L] {nombre_archivo_l}")
                    ws.Cells(fila, col_resultado_l).Value = nombre_archivo_l
                except Exception as e:
                    print(f"[{_now()}] [ERROR] No se pudo copiar L: {e}")
                    ws.Cells(fila, col_resultado_l).Value = "No"
            else:
                ws.Cells(fila, col_resultado_l).Value = "No"

            procesadas += 1
            fila += 1

        print(f"[{_now()}] [FINALIZADO] Se procesaron {procesadas} entradas.")
    
    except Exception:
        print(f"[{_now()}] [ERROR GENERAL] Se produjo un error inesperado:\n{traceback.format_exc()}")
    finally:
        if 'wb' in locals() and wb:
            try:
                wb.Save()
                print(f"[{_now()}] [INFO] Cambios guardados en el libro.")
            except Exception as e:
                print(f"[{_now()}] [ADVERTENCIA] No se pudo guardar el libro: {e}")

# ───────────────────── PUNTO DE ENTRADA ─────────────────────
if __name__ == "__main__":
    copiar_archivos_masivo()