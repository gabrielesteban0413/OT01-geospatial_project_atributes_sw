import pandas as pd

ruta = r"C:\A_GS1_PROYECTOS\00_MODULOS\Z_CALIDAD\Z_REPORTES\Z_OTHERS\EXP_MS_FI.xlsx"
print(pd.ExcelFile(ruta).sheet_names)
