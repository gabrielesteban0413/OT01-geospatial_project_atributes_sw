#se presenta potsgres


import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import numpy as np


archivo_excel = r"C:\A_GS1_PROYECTOS\dbasphia.csv"  

conexion = psycopg2.connect(
    host="localhost",
    port="5432",
    database="GS01_DBASPHIA",
    user="postgres",
    password="FGS_eelgs44FGRT#44"  
)



df = pd.read_excel(archivo_excel)
df.columns = df.columns.str.strip().str.lower()

print(f" Columnas: {list(df.columns)}")
print(f" Total filas: {len(df)}")
df = df.replace(['N/A', 'n/a', '', 'nan', 'NaN', None], np.nan)
df['fecha'] = pd.to_datetime(df['fecha'], errors='coerce', format='%d/%m/%Y')



columnas_numericas = ['idcliente', 'n_hilo', 'l_cable_acceso', 'bw', 'n_hilo_troncal', 'l_cable_troncal']
for col in columnas_numericas:
    if col in df.columns:
        df[col] = pd.to_numeric(df[col], errors='coerce')

columnas_texto = df.select_dtypes(include=['object']).columns
for col in columnas_texto:
    df[col] = df[col].astype(str).replace('nan', None)

print(" INSERT...")

cursor = conexion.cursor()


cursor.execute("TRUNCATE TABLE clientes_servicios")


lote = []
total = len(df)

for idx, row in df.iterrows():
    lote.append((
        row.get('idcliente'), row.get('diseno'), row.get('cambio'), row.get('version'),
        row.get('actividad'), row.get('n_acceso'), row.get('n_hilo'), row.get('anillo'),
        row.get('cliente'), row.get('direccion'), row.get('servicio'), row.get('l_cable_acceso'),
        row.get('fecha'), row.get('observacion'), row.get('consecutivo_cable_acceso'),
        row.get('disenador'), row.get('cuadrilla'), row.get('ciudad'), row.get('bw'),
        row.get('und_bw'), row.get('idservicio'), row.get('equipo'), row.get('n_hilo_troncal'),
        row.get('n_troncal'), row.get('l_cable_troncal'), row.get('odf')
    ))
    
    if len(lote) >= 1000:
        execute_values(cursor, """
            INSERT INTO clientes_servicios VALUES %s
        """, lote)
        print(f"    Subidas {idx + 1} de {total} filas")
        lote = []


if lote:
    execute_values(cursor, "INSERT INTO clientes_servicios VALUES %s", lote)

conexion.commit()
cursor.close()
conexion.close()

print(f"\n ¡COMPLETADO! {total} filas importadas correctamente")