import os
from openpyxl import Workbook
from openpyxl.styles import Alignment

BASE_PATH = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Shape_ MCB\00_ANILLOS"
OUTPUT_FILE = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Shape_ MCB\00_REPORTE_SHAPES.xlsx"

def generar_excel():
    if not os.path.exists(BASE_PATH):
        print(f"Ruta no encontrada: {BASE_PATH}")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "REPORT_SHAPES"
    ws.append(["NAME", "Ruta Completa", "."])

    for carpeta in os.listdir(BASE_PATH):
        ruta_completa = os.path.join(BASE_PATH, carpeta)
        if os.path.isdir(ruta_completa):
            ws.append([carpeta, ruta_completa, "."])

    ajustar_formato(ws)
    wb.save(OUTPUT_FILE)
    print(f"----REPORT SHP-------- {OUTPUT_FILE}")

def ajustar_formato(ws):
    # Ancho columnas
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 30
    ws.column_dimensions["C"].width = 5

    # Alinear a la derecha la columna B
    for fila in ws.iter_rows(min_row=2, min_col=2, max_col=2):
        for celda in fila:
            celda.alignment = Alignment(horizontal="right")

if __name__ == "__main__":
    generar_excel()
