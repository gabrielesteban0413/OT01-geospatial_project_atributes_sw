import os
import pandas as pd
import win32com.client as win32
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import unicodedata


archivo_nombres = "C:\\A_GS1_PROYECTOS\\0.PB_RP_V1.xlsx"
archivo_tabla = "C:\\A_GS1_PROYECTOS\\001_MARZO_ CARGAS DE SERVICIO_GIT PROGRAMAS Y DENUNCIAS.xlsx"
archivo_correos = "C:\\A_GS1_PROYECTOS\\Planta funcionarios a marzo 31 de 2025 Fiscalización Naturales.xlsx"

try:
    excel_nombres = pd.read_excel(archivo_nombres, sheet_name="DB")
    excel_tabla = pd.read_excel(
    archivo_tabla,
    sheet_name="1Consolidado Cargas de Servicio",
    header=10,
    dtype={'Consecutivo': str}
)

    excel_correos = pd.read_excel(archivo_correos, sheet_name="PlantaFiscaNaturalmarzo31-2025")
except Exception as e:
    print(f"❌ Error cargando archivos: {e}")
    exit()

try:
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True 
    libro = excel.Workbooks.Open(archivo_nombres) 
    hoja = libro.Sheets("DB") 
except Exception as e:
    print(f"❌ No se pudo abrir el archivo Excel con COM. Error: {e}")
    exit()



# Función
def normalizar_texto(texto):
    return unicodedata.normalize('NFD', texto).encode('ascii', 'ignore').decode('ascii')


def obtener_correo(nombre):
    try:
        nombre_normalizado = normalizar_texto(nombre).strip().lower()
        fila = excel_correos[excel_correos['formula'].astype(str).str.strip().str.lower().apply(normalizar_texto) == nombre_normalizado]
        if fila.empty:
            return None
        return fila['Correo'].values[0]
    except Exception as e:
        print(f"❌ Error obteniendo correo: {e}")
        return None

def obtener_saludo(nombre):
    if not isinstance(nombre, str):
        return "Compañer@", "Compañer@"
    if nombre.strip().endswith("o"):
        return f"Estimado {nombre}", "Estimado"
    else:
        return f"Estimada {nombre}", "Estimada"

def dataframe_to_html(df):
    html = df.to_html(index=False, header=True, escape=False)
    html = f"""
    <html>
    <head>
        <style>
            table {{
                width: 70%;
                border-collapse: collapse;
                margin: 10px auto;
                font-family: Arial, sans-serif;
                font-size: 8pt;
            }}
            th {{
                background-color: rgb(0, 176, 240);
                color: white;
                font-size: 10pt;
                height: 21px;
                padding: 2px;
                text-align: left;
            }}
            td {{
                border: 1px solid black;
                height: 10px;
                padding: 1px;
                text-align: left;
                font-family: Arial, sans-serif;
                font-size: 8pt;
            }}
        </style>
    </head>
    <body style="font-family: Arial, sans-serif; font-size: 10pt;">
        {html}
    </body>
    </html>
    """
    return html


outlook = win32.Dispatch("Outlook.Application")
print("📧 Outlook iniciado correctamente.")

fila_actual = 1
columnas_a_incluir = [
    'Nombre Funcionario',
    'Grupo Interno de Trabajo / Coordinación',
    'NIT',
    'Razón Social',
    'Consecutivo',
    'Estado de la Carga no Persuasiva',
    'Fecha Apertura',
    'Fecha Vencimiento'
]


for index, nombre in enumerate(excel_nombres['NOMBRE']):
    if pd.isna(nombre):
        continue

    nombre = str(nombre).strip()
    print(f"\n🔄 Procesando: {nombre}")

    saludo, genero = obtener_saludo(nombre)
    correo_destino = obtener_correo(nombre)

    if not correo_destino:
        print(f"❌ Correo no encontrado para {nombre}")
        hoja.Cells(index + 2, 2).Value = 'Correo no encontrado'
        continue

    tabla_filtrada = excel_tabla[excel_tabla['Nombre Funcionario'].astype(str).str.strip().str.lower() == nombre.lower()]
    if tabla_filtrada.empty:
        print(f"⚠️ No se encontraron datos para {nombre}")
        hoja.Cells(index + 2, 2).Value = 'Datos no encontrados'  
        continue

    tabla_filtrada = tabla_filtrada[columnas_a_incluir]




    mail = outlook.CreateItem(0)
    mail.To = correo_destino
    mail.CC = "gabriel.boadas@colvatel.com.co; gabriel.boadas@colvatel.com.co"
    mail.Subject = "Cargas próximas a vencer"

    cuerpo = f"""<p style="font-family: Arial, sans-serif;">Buenos días {saludo},</p>
    <p style="font-family: Arial, sans-serif;">Cordial saludo,</p>
    <p style="font-family: Arial, sans-serif;">Envío las cargas que están próximas a vencer.</p>
    <p style="font-family: Arial, sans-serif;">Este mensaje se envía de carácter preventivo, lo enviamos de la manera más cordial posible y cuente con todo nuestro apoyo.</p>
    <p style="font-family: Arial, sans-serif;">Equipo de Fiscalización</p>
    <p>&nbsp;</p>"""

    tabla_html = dataframe_to_html(tabla_filtrada)
    mail.HTMLBody = cuerpo + tabla_html
    mail.Display()

    print(f" Correo listo para: {nombre}")