# ─────────────────────────── IMPORTACIONES ───────────────────────────────
import os
import shutil
import win32com.client
import traceback
from pathlib import Path

# ─────────────────────────── CONSTANTES Y RUTAS ────────────────────────────

BASE_PROYECTOS = Path(r"C:\A_GS1_PROYECTOS")
RUTA_DESTINO_BASE = BASE_PROYECTOS
RUTAS_ORIGEN = {
    "c": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Troncal de Acceso\Cartográfico"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Troncal de Acceso\Lógico")
    ],
    "a": [
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Tendidos_ANILLOS"),
        Path(r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\ANILLOS_ACCESO\Logicos_ANILLOS")
    ]
}

# ─────────────────────────── UTILIDADES ────────────────────────────────

def buscar_archivo_mas_reciente(ruta_base, patron):
    def buscar(patron_busqueda):
        archivo_mas_reciente = None
        fecha_mas_reciente = 0
        carpetas_pendientes = [ruta_base]

        while carpetas_pendientes:
            carpeta_actual = carpetas_pendientes.pop()
            try:
                with os.scandir(carpeta_actual) as it:
                    for entry in it:
                        if entry.is_dir(follow_symlinks=False):
                            carpetas_pendientes.append(entry.path)
                        elif entry.is_file() and patron_busqueda.lower() in entry.name.lower():
                            fecha_mod = entry.stat().st_mtime
                            if fecha_mod > fecha_mas_reciente:
                                fecha_mas_reciente = fecha_mod
                                archivo_mas_reciente = entry.path
            except PermissionError:
                print(f"[ADVERTENCIA] Sin permiso para acceder a: {carpeta_actual}")
                continue
            except Exception as e:
                print(f"[ERROR] Al acceder a: {carpeta_actual}\n{e}")
                continue

        return archivo_mas_reciente

    print(f"[INFO] Buscando archivo que contenga: {patron}")
    archivo = buscar(patron)

    if not archivo:
        print(f"[INFO] No se encontró con el patrón '{patron}', intentando con 'C.{patron}'")
        base_patron = Path(patron).stem 
        extension = Path(patron).suffix
        archivo = buscar(f"C.{base_patron}{extension}")

    if archivo:
        print(f"[ENCONTRADO] {archivo}")
    else:
        print("[NO ENCONTRADO] No se encontró ningún archivo con ese patrón.")

    return archivo

# ─────────────────────────── FUNCIÓN PRINCIPAL ───────────────────────────

def copiar_archivos_desde_excel():
    """Copia archivos DWG desde rutas compartidas usando valores de Excel."""
    try:
        print("[INFO] Intentando conectarse a Excel...")
        xl = win32com.client.GetActiveObject("Excel.Application")
        wb = xl.ActiveWorkbook
        ws = wb.ActiveSheet
        print("[ÉXITO] Conectado a Excel correctamente.")
    except Exception:
        print("[ERROR] No se pudo conectar con Excel. Asegúrate de que el archivo está abierto.")
        input("Presiona ENTER para salir...")
        return

    try:
        nombre_carpeta = str(ws.Range("B1").Value).strip()
        tipo_busqueda = str(ws.Range("B2").Value).strip().lower()

        print(f"[INFO] Carpeta destino: {nombre_carpeta}")
        print(f"[INFO] Tipo de búsqueda: {tipo_busqueda}")

        if not nombre_carpeta:
            print("[ERROR] Celda B1 vacía. Se requiere un nombre de carpeta.")
            xl.Application.Run("MsgBox", "ERROR: Celda B1 vacía.", 48)
            return

        if tipo_busqueda not in RUTAS_ORIGEN:
            print("[ERROR] Celda B2 con valor inválido. Debe ser 'c' o 'a'.")
            xl.Application.Run("MsgBox", "ERROR: Celda B2 con valor incorrecto.", 48)
            return

        # ------------------ Búsqueda y Copia ------------------
        rutas = RUTAS_ORIGEN[tipo_busqueda]
        archivo_t = buscar_archivo_mas_reciente(rutas[0], f"{nombre_carpeta}_T.dwg")
        archivo_l = buscar_archivo_mas_reciente(rutas[1], f"{nombre_carpeta}_L.dwg")
        destino = RUTA_DESTINO_BASE / nombre_carpeta
        os.makedirs(destino, exist_ok=True)
        print(f"[INFO] Carpeta de destino creada o existente: {destino}")

        if archivo_t:
            shutil.copy(archivo_t, destino)
            nombre_t = Path(archivo_t).name
            print(f"[COPIADO] Archivo T copiado: {nombre_t}")
            ws.Range("B24").Value = str(destino / nombre_t) 
            ws.Range("B3").Value = nombre_t  
        else:
            print("[ADVERTENCIA] No se encontró el archivo T.")


        if archivo_l:
            shutil.copy(archivo_l, destino)
            print(f"[COPIADO] Archivo L copiado: {Path(archivo_l).name}")
        else:
            print("[ADVERTENCIA] No se encontró el archivo lógico.")
            xl.Application.Run("MsgBox", "No se encontró el archivo lógico.", 48)

    except Exception as e:
        print(f"[ERROR] Se produjo un error inesperado:\n{traceback.format_exc()}")
        input("Presiona ENTER para salir...")

# ───────────────────── PUNTO DE ENTRADA ─────────────────────
if __name__ == "__main__":
    copiar_archivos_desde_excel()