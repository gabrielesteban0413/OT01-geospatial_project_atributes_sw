"TRAZE CON SINCRINIZACION TABLA INGE ARTURO, SE AREALIZA POR CABLE"

import sys
import re
import time
import logging
import multiprocessing
from pathlib import Path
from typing import Dict, List, Any
from collections import Counter

import pandas as pd
import win32com.client as win32
import pythoncom

logging.basicConfig(level=logging.CRITICAL)
logger = logging.getLogger(__name__)

RUTA_ORIGEN = Path(r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos"
                   r"\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO"
                   r"\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\11_Trace_Conectividad\70926")
RUTA_MOTOR = Path(r"C:\A_GS1_PROYECTOS\00_GV_TRAZE.xlsm")
RUTA_RESPONSABLES = Path(r"C:\A_GS1_PROYECTOS\Reporte anillos-clientes.xlsx")
RUTA_SALIDA_BASE = Path(r"C:\A_GS1_PROYECTOS")



class ResponsableLoader:
    def __init__(self, ruta: Path):
        self.ruta = ruta
        self.datos: Dict[str, Dict[str, str]] = {}

    def cargar(self) -> Dict[str, Dict[str, str]]:
        if not self.ruta.exists():
            return {}

        try:
            df = pd.read_excel(self.ruta, engine='openpyxl', dtype=str)
            df.columns = [str(c).strip().lower() for c in df.columns]

            col_anillo = None
            col_responsable = None
            col_total_clientes = None

            for col in df.columns:
                if "anillo" in col:
                    col_anillo = col
                elif "responsable" in col:
                    col_responsable = col
                elif "total clientes" in col or "total_clientes" in col:
                    col_total_clientes = col

            if col_anillo is None or col_responsable is None:
                if len(df.columns) > 2:
                    col_anillo = df.columns[2]
                if len(df.columns) > 1:
                    col_responsable = df.columns[1]
                if len(df.columns) > 3:
                    col_total_clientes = df.columns[3]

            if col_anillo is None or col_responsable is None:
                return {}

            datos = {}
            for _, row in df.iterrows():
                anillo_val = row[col_anillo]
                responsable_val = row[col_responsable]
                total_clientes_val = row[col_total_clientes] if col_total_clientes else None

                anillo = str(anillo_val).strip() if pd.notna(anillo_val) else ""
                responsable = str(responsable_val).strip() if pd.notna(responsable_val) else ""
                total_clientes = str(total_clientes_val).strip() if total_clientes_val and pd.notna(total_clientes_val) else ""

                if anillo:
                    datos[anillo] = {
                        "responsable": responsable,
                        "total_clientes": total_clientes
                    }

            return datos

        except Exception:
            return {}

    @staticmethod
    def extraer_anillo(nombre_hoja: str) -> str:
        partes = nombre_hoja.split('_')
        if len(partes) >= 3:
            return partes[2]
        return nombre_hoja


class TextFileProcessor:
    @staticmethod
    def leer_lineas(archivo: Path) -> List[str]:
        try:
            with open(archivo, "r", encoding="utf-8") as f:
                lineas = f.readlines()
        except UnicodeDecodeError:
            with open(archivo, "r", encoding="latin-1") as f:
                lineas = f.readlines()
        return [l.strip() for l in lineas if l.strip()]


class ExcelHandler:
    def __init__(self, ruta_motor: Path, ruta_salida: Path):
        self.ruta_motor = ruta_motor
        self.ruta_salida = ruta_salida
        self.excel = None
        self.motor = None
        self.libro_salida = None
        self.visible_original = None
        self.display_alerts_original = None
        self.es_instancia_existente = False

    def iniciar(self) -> bool:
        try:
            self.excel = win32.DispatchEx("Excel.Application")
            self.es_instancia_existente = False
            self.visible_original = False
            self.display_alerts_original = False
        except Exception:
            return False

        self.excel.Visible = False
        self.excel.DisplayAlerts = False
        self.excel.ScreenUpdating = False

        try:
            self.motor = self.excel.Workbooks.Open(str(self.ruta_motor))
            self.libro_salida = self.excel.Workbooks.Add()
            return True
        except Exception:
            self.cerrar()
            return False

    def cerrar(self) -> None:
        if self.excel:
            self.excel.ScreenUpdating = True
            try:
                if self.libro_salida:
                    self.libro_salida.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if self.motor:
                    self.motor.Close(SaveChanges=False)
            except Exception:
                pass
            if self.es_instancia_existente:
                try:
                    self.excel.Visible = self.visible_original
                    self.excel.DisplayAlerts = self.display_alerts_original
                except Exception:
                    pass
            else:
                self.excel.Quit()

    def crear_hoja(self, nombre: str) -> Any:
        if self.libro_salida.Sheets.Count == 1 and self.libro_salida.Sheets(1).Name == "Hoja1":
            hoja = self.libro_salida.Sheets(1)
            hoja.Name = nombre
            return hoja
        else:
            hoja = self.libro_salida.Sheets.Add(After=self.libro_salida.Sheets(self.libro_salida.Sheets.Count))
            hoja.Name = nombre
            return hoja

    def ejecutar_macros(self, hoja: Any) -> None:
        hoja.Activate()
        nombre_motor = self.ruta_motor.name
        self.excel.Application.Run(f"'{nombre_motor}'!G01_01_REPORT_TABLE")
        self.excel.Application.Run(f"'{nombre_motor}'!G01_02_DIAGRAMA_COMPONENTES")

    def escribir_lineas(self, hoja: Any, lineas: List[str]) -> None:
        fila = 4
        for linea in lineas:
            hoja.Cells(fila, 2).Value = linea
            fila += 1

    def limpiar_datos(self, hoja: Any) -> None:
        ultima_fila = hoja.Cells(hoja.Rows.Count, 2).End(-4162).Row
        if ultima_fila >= 4:
            hoja.Range(f"B4:B{ultima_fila}").ClearContents()
        hoja.Columns("I:K").AutoFit()

    def leer_clientes_desde_shapes(self, hoja: Any) -> str:
        patron = re.compile(r"CLIENTES\s*:\s*(\d+)", re.IGNORECASE)
        try:
            for shape in hoja.Shapes:
                try:
                    txt = shape.TextFrame2.TextRange.Text
                    m = patron.search(txt)
                    if m:
                        return m.group(1)
                except Exception:
                    pass
        except Exception:
            pass
        return ""

    def crear_indice(self, info_hojas: List[Dict[str, Any]],
                     responsable_por_anillo: Dict[str, Dict[str, str]]) -> None:
        hoja_indice = self.libro_salida.Sheets.Add(Before=self.libro_salida.Sheets(1))
        hoja_indice.Name = "ÍNDICE"

        headers = ["ARCHIVO", "HOJA", "RESPONSABLE", "CLIENTES_TRACE", "TOTAL_CLIENTES_RESPONSABLE", "COMPARACION"]
        for col, titulo in enumerate(headers, start=1):
            hoja_indice.Cells(1, col).Value = titulo
        hoja_indice.Range("A1:F1").Font.Bold = True

        for i, info in enumerate(info_hojas, start=2):
            nombre_hoja = info['nombre_hoja']
            anillo = info['anillo']
            clientes_trace = info['clientes_trace']

            celda = hoja_indice.Cells(i, 1)
            hoja_indice.Hyperlinks.Add(
                Anchor=celda,
                Address="",
                SubAddress=f"'{nombre_hoja}'!Q4",
                TextToDisplay=nombre_hoja,
            )
            hoja_indice.Cells(i, 2).Value = nombre_hoja

            info_anillo = responsable_por_anillo.get(anillo, {"responsable": "", "total_clientes": ""})
            responsable = info_anillo["responsable"]
            total_clientes = info_anillo["total_clientes"]

            if responsable:
                hoja_indice.Cells(i, 3).Value = responsable

            if clientes_trace.isdigit():
                hoja_indice.Cells(i, 4).Value = int(clientes_trace)
            else:
                hoja_indice.Cells(i, 4).Value = clientes_trace

            if total_clientes and total_clientes.isdigit():
                hoja_indice.Cells(i, 5).Value = int(total_clientes)
            else:
                hoja_indice.Cells(i, 5).Value = total_clientes

            hoja_indice.Cells(i, 6).Formula = f"=D{i}=E{i}"

            sh = self.libro_salida.Sheets(nombre_hoja)
            sh.Activate()
            sh.Range("Q4").Select()

        hoja_indice.Columns("A:F").AutoFit()
        hoja_indice.Activate()
        hoja_indice.Range("A1").Select()

    def guardar(self) -> bool:
        try:
            self.libro_salida.SaveAs(str(self.ruta_salida), FileFormat=52)
            return True
        except Exception:
            return False


class WorkerProcess:
    def __init__(self, ruta_origen: Path, ruta_motor: Path, ruta_salida_base: Path,
                 ruta_responsables: Path, cola: multiprocessing.Queue):
        self.ruta_origen = ruta_origen
        self.ruta_motor = ruta_motor
        self.ruta_salida_base = ruta_salida_base
        self.ruta_responsables = ruta_responsables
        self.cola = cola
        self.cable_nombre = self.ruta_origen.name

    def ejecutar(self) -> None:
        try:
            pythoncom.CoInitialize()

            loader = ResponsableLoader(self.ruta_responsables)
            responsables = loader.cargar()

            archivos_txt = sorted(self.ruta_origen.glob("*.txt"))
            total = len(archivos_txt)
            if total == 0:
                self.cola.put(("error", f"No se encontraron archivos .txt en {self.ruta_origen}"))
                return

            anillos = []
            info_hojas = []

            for idx, archivo in enumerate(archivos_txt, 1):
                nombre_base = archivo.stem
                nombre_hoja = nombre_base[:31]
                anillo = ResponsableLoader.extraer_anillo(nombre_base)
                anillos.append(anillo)

                self.cola.put(("progreso", idx, total, nombre_base))

                info_hojas.append({
                    'nombre_hoja': nombre_hoja,
                    'anillo': anillo,
                    'clientes_trace': ""  
                })

            responsables_anillos = {anillo: responsables.get(anillo, {}).get("responsable", "") for anillo in anillos}
            contador = Counter(responsables_anillos.values())
            responsable_mayoritario = contador.most_common(1)[0][0] if contador else ""

            if responsable_mayoritario:
                nombre_archivo = f"{self.cable_nombre}_{responsable_mayoritario}.xlsm"
            else:
                nombre_archivo = f"{self.cable_nombre}.xlsm"

            ruta_salida = self.ruta_salida_base / nombre_archivo

            excel_handler = ExcelHandler(self.ruta_motor, ruta_salida)
            if not excel_handler.iniciar():
                self.cola.put(("error", "No se pudo iniciar Excel"))
                return

            try:
                for idx, archivo in enumerate(archivos_txt, 1):
                    nombre_base = archivo.stem
                    nombre_hoja = nombre_base[:31]
                    anillo = ResponsableLoader.extraer_anillo(nombre_base)

                    self.cola.put(("progreso", idx, total, nombre_base))

                    hoja = excel_handler.crear_hoja(nombre_hoja)

                    lineas = TextFileProcessor.leer_lineas(archivo)
                    excel_handler.escribir_lineas(hoja, lineas)

                    excel_handler.ejecutar_macros(hoja)

                    clientes = excel_handler.leer_clientes_desde_shapes(hoja)
                    info_hojas[idx-1]['clientes_trace'] = clientes

                    excel_handler.limpiar_datos(hoja)

                    time.sleep(0.3)

                excel_handler.crear_indice(info_hojas, responsables)

                if excel_handler.guardar():
                    self.cola.put(("ok",))
                else:
                    self.cola.put(("error", "Error al guardar el archivo de salida"))

            finally:
                excel_handler.cerrar()

        except Exception as exc:
            self.cola.put(("error", str(exc)))
        finally:
            pythoncom.CoUninitialize()


def main() -> None:
    cola = multiprocessing.Queue()

    worker = WorkerProcess(RUTA_ORIGEN, RUTA_MOTOR, RUTA_SALIDA_BASE, RUTA_RESPONSABLES, cola)
    proceso = multiprocessing.Process(target=worker.ejecutar, daemon=True)
    proceso.start()

    try:
        while True:
            mensaje = cola.get(timeout=300)
            tipo = mensaje[0]
            if tipo == "progreso":
                pass
            elif tipo == "ok":
                break
            elif tipo == "error":
                proceso.terminate()
                sys.exit(1)
    except Exception:
        proceso.terminate()
        sys.exit(1)

    proceso.join()


if __name__ == "__main__":
    multiprocessing.freeze_support()
    main()