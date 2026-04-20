# ─────────────────────────── IMPORTACIONES ─────────────────────────────── #pyinstaller --onefile 02_Bulk_history.py
import os
import shutil
import win32com.client as win32
import ctypes  
from pathlib import Path
import traceback

try:
    from procesar_dwg_arcgis import procesar_dwg
except ImportError:

    def procesar_dwg(ruta_dwg, carpeta_destino):
        print(f"[ADVERTENCIA] procesar_dwg no está implementado. Archivo: {ruta_dwg}")

# ─────────────────────────── CONSTANTES Y RUTAS ────────────────────────────

# RUTA DEL ARCHIVO EXCEL QUE DEBE SER ABIERTO
RUTA_EXCEL_PRINCIPAL = Path(r"C:\A_GS1_PROYECTOS\0.GS.DB.V3.xlsm")

BASE_PROYECTOS     = Path(r"C:\A_GS1_PROYECTOS")
RUTA_CARPETA       = BASE_PROYECTOS 
RUTA_EXTRA_AUTOCAD = Path(r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Autocad")


RUTAS_ORIGEN = {

    "c": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Troncal de Acceso\Cartográfico\Bogotá"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Troncal de Acceso\Lógico")
    ],
    "a": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Tendidos_ANILLOS"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Logicos_ANILLOS")
    ]
}

RUTA_CORPORATIVO = Path(r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Shape_ MCB")
RUTA_40_ANILLOS    = RUTA_CORPORATIVO / "00_ANILLOS"
RUTA_40_TRONCALES  = RUTA_CORPORATIVO / "00_TRONCALES"

# ─────────────────────────── CONSTANTES DE MENSAJES ────────────────────────


MB_ICONERROR = 0x10
MSG_TIPO_BUSQUEDA_ERROR = "ERROR: Tipo de búsqueda incorrecto. Use 'A' o 'C'."


# ─────────────────────────── UTILIDADES ────────────────────────────────

def buscar_archivo_mas_reciente(ruta_base, patron):
    """
    Busca el archivo más reciente en la carpeta y subcarpetas.
    (Se mantuvo la implementación original con os.walk para respetar la función anterior)
    """
    archivo_mas_reciente = None
    fecha_mas_reciente = 0
    
    try:
        for root, _, files in os.walk(str(ruta_base)):
            for nombre in files:
                if patron.lower() in nombre.lower():
                    ruta_completa = Path(root) / nombre
                    try:
                        # Utilizamos Path.stat().st_mtime para obtener la fecha de modificación
                        fecha_mod = ruta_completa.stat().st_mtime
                        if fecha_mod > fecha_mas_reciente:
                            fecha_mas_reciente = fecha_mod
                            archivo_mas_reciente = ruta_completa
                    except Exception:
                        continue
    except Exception as e:
        print(f"[ERROR] Al recorrer la ruta '{ruta_base}': {e}")
        
    return archivo_mas_reciente

def obtener_excel():
    """
    Intenta conectarse al libro de trabajo específico o crea una nueva instancia de Excel y lo abre.
    Devuelve la aplicación (xl) y el libro (wb).
    """
    try:
        # Intenta conectarse a una instancia activa de Excel
        xl = win32.GetActiveObject("Excel.Application")
    except Exception:
        # Si no hay una instancia activa, crea una nueva
        try:
            xl = win32.Dispatch("Excel.Application")
            xl.Visible = True # Muestra la ventana de Excel
        except Exception:
            print("ERROR: No se pudo conectar ni crear una instancia de Excel.")
            return None, None

    # Ahora que tenemos la aplicación, intentamos acceder o abrir el libro.
    try:
        # 1. Intenta acceder al libro si ya está abierto en la instancia conectada
        wb = xl.Workbooks(RUTA_EXCEL_PRINCIPAL.name)
        print(f"[INFO] Conectado al libro abierto: {RUTA_EXCEL_PRINCIPAL.name}")
    except Exception:
        # 2. Si no está abierto, ábrelo
        try:
            wb = xl.Workbooks.Open(str(RUTA_EXCEL_PRINCIPAL))
            print(f"[INFO] Libro abierto: {RUTA_EXCEL_PRINCIPAL.name}")
        except Exception:
            print(f"ERROR: No se pudo encontrar ni abrir el archivo: {RUTA_EXCEL_PRINCIPAL}")
            return None, None
            
    return xl, wb

# ─────────────────────────── FUNCIÓN PRINCIPAL ───────────────────────────

def copiar_archivos_masivo():
    """
    Procesa múltiples valores desde Excel y copia archivos DWG.
    """
    # Se modificó para recibir xl y wb (workbook)
    xl, wb = obtener_excel() 
    if not xl or not wb:
        # Mostrar error si no se pudo acceder al Excel/Libro
        ctypes.windll.user32.MessageBoxW(
            0, 
            f"No se pudo conectar al archivo Excel: {RUTA_EXCEL_PRINCIPAL.name}. Asegúrese de que existe y se puede acceder.", 
            "Error de Conexión", 
            MB_ICONERROR
        )
        return

    try:
        # El libro ya está abierto (wb), ahora obtenemos la hoja activa
        ws = wb.ActiveSheet
        
        tipo_busqueda = str(xl.InputBox("Ingrese tipo de búsqueda (A o C):",
                                         "Tipo de Búsqueda", "A")).strip().lower()

        if tipo_busqueda not in ["a", "c"]:

            ctypes.windll.user32.MessageBoxW(0, MSG_TIPO_BUSQUEDA_ERROR, "Error", MB_ICONERROR)
            return

        RUTA_EXTRA_AUTOCAD.mkdir(parents=True, exist_ok=True)
        print(f"[INFO] Carpeta temporal creada o existente: {RUTA_EXTRA_AUTOCAD}")

        fila = 31
        procesadas = 0
        sin_datos_consecutivos = 0
        max_sin_datos = 10  

        while sin_datos_consecutivos < max_sin_datos:
            valor_celda = ws.Cells(fila, 2).Value
            
            if not valor_celda or str(valor_celda).strip().lower() in ("", "none"):
                sin_datos_consecutivos += 1
                fila += 1
                continue

            sin_datos_consecutivos = 0
            nombre_carpeta = str(valor_celda).strip()
            print(f"\n[PROCESANDO] Fila {fila}: {nombre_carpeta}")

            # ------------------ Búsqueda en Corporativo (Carpeta Existente) 
            ruta_base_40 = RUTA_40_ANILLOS if tipo_busqueda == "a" else RUTA_40_TRONCALES
            carpeta_corporativo = ruta_base_40 / nombre_carpeta

            if carpeta_corporativo.exists():
                ws.Cells(fila, 3).Value = str(carpeta_corporativo)
                print(f"[SKIP] Carpeta encontrada en Corporativo: {carpeta_corporativo}. Saltando.")
                fila += 1
                procesadas += 1
                continue

            # ------------------ Búsqueda especial de Archivos en Corporativo (Solo Anillos) 
            if tipo_busqueda == "a":
                ruta_40_t = ruta_base_40 / f"{nombre_carpeta}_T.dwg"
                ruta_40_l = ruta_base_40 / f"{nombre_carpeta}_L.dwg"

                if ruta_40_t.exists() or ruta_40_l.exists():
                    ws.Cells(fila, 3).Value = str(ruta_40_t) if ruta_40_t.exists() else "No"
                    ws.Cells(fila, 4).Value = str(ruta_40_l) if ruta_40_l.exists() else "No"
                    print("[SKIP] Archivos encontrados en Corporativo 40. Saltando.")
                    fila += 1
                    procesadas += 1
                    continue

            # ------------------ Búsqueda normal en Cardiseño 
            rutas_cardiseno = RUTAS_ORIGEN[tipo_busqueda]
            
            archivo_t = buscar_archivo_mas_reciente(
                rutas_cardiseno[0], f"{nombre_carpeta}_T.dwg"
            )
            archivo_l = buscar_archivo_mas_reciente(
                rutas_cardiseno[1], f"{nombre_carpeta}_L.dwg"
            )

            if not archivo_t and not archivo_l:
                print("[ADVERTENCIA] No se encontraron archivos T ni L en Cardiseño.")
                fila += 1
                continue

            # ------------------ Copia y Procesamiento 
            # Crear carpeta de destino
            destino = RUTA_CARPETA / nombre_carpeta
            destino.mkdir(parents=True, exist_ok=True)

            # Procesar archivo T
            if archivo_t:
                nombre_archivo_t = archivo_t.name
                ruta_t_destino = destino / nombre_archivo_t
                
                shutil.copy(archivo_t, ruta_t_destino)  
                shutil.copy(archivo_t, RUTA_EXTRA_AUTOCAD) 
                
                ws.Cells(fila, 24).Value = str(ruta_t_destino)
                ws.Cells(fila, 3).Value = nombre_archivo_t
                print(f"[COPIADO T] {nombre_archivo_t}")

                try:
                    
                    procesar_dwg(ruta_dwg=str(ruta_t_destino), carpeta_destino=str(destino))
                except Exception as e:
                    print(f"[ERROR ArcGIS] Falló procesar_dwg para {ruta_t_destino}: {e}")
            else:
                ws.Cells(fila, 3).Value = "No"

            
            if archivo_l:
                shutil.copy(archivo_l, destino)
                print(f"[COPIADO L] {archivo_l.name}")
            else:
                ws.Cells(fila, 4).Value = "No"

            procesadas += 1
            fila += 1

        print(f"\n[FINALIZADO] Se procesaron {procesadas} entradas.")
    
    except Exception:
        print(f"[ERROR GENERAL] Se produjo un error inesperado:\n{traceback.format_exc()}")
    finally:
        # Asegúrate de guardar los cambios y cerrar el libro al finalizar
        if 'wb' in locals() and wb:
            try:
                # Si el archivo era .xlsm, es importante guardarlo.
                wb.Save()
                print("[INFO] Cambios guardados en el libro.")
            except Exception as e:
                print(f"[ADVERTENCIA] No se pudo guardar el libro: {e}")
        
        # IMPORTANTE: No cierres la aplicación de Excel (xl) si la inicializaste 
        # con GetActiveObject, ya que cerrarías la instancia del usuario. 
        # Si la creaste tú (Dispatch), se recomienda cerrarla.
        pass

# ───────────────────── PUNTO DE ENTRADA ─────────────────────
if __name__ == "__main__":
    copiar_archivos_masivo()

