import os
import shutil
import win32com.client

# Conectarse a Excel
excel = win32com.client.Dispatch("Excel.Application")

# Buscar el libro abierto
ruta_objetivo = r"C:\A_GS1_PROYECTOS\0.GS.DB.V3.xlsm".lower()
wb = None
for libro in excel.Workbooks:
    if libro.FullName.lower() == ruta_objetivo:
        wb = libro
        break

if wb is None:
    print("El archivo no está abierto en Excel.")
    exit()

ws = wb.Sheets("DB")

# Leer desde fila 32
datos = []
row = 32
while True:
    origen = ws.Cells(row, 2).Value  # Columna B
    destino = ws.Cells(row, 5).Value # Columna E
    if origen is None and destino is None:
        break
    datos.append((origen, destino))
    row += 1

# Procesamiento de carpetas
ruta_base = r"C:\A_GS1_PROYECTOS"
for nombre_origen, nombre_destino in datos:
    if not nombre_origen or not nombre_destino:
        print(f"Fila con datos faltantes: origen={nombre_origen}, destino={nombre_destino}")
        continue

    nombre_origen = str(nombre_origen).strip()
    nombre_destino = str(nombre_destino).strip()

    if not nombre_origen or not nombre_destino:
        print(f"Fila con texto vacío: origen='{nombre_origen}', destino='{nombre_destino}'")
        continue

    ruta_origen = os.path.join(ruta_base, nombre_origen)
    ruta_destino = os.path.join(ruta_base, nombre_destino)

    os.makedirs(ruta_destino, exist_ok=True)

    if os.path.exists(ruta_origen):
        try:
            shutil.move(ruta_origen, ruta_destino)
            print(f"Movido: {ruta_origen} -> {ruta_destino}")
        except Exception as e:
            print(f"Error al mover {ruta_origen}: {e}")
    else:
        print(f"No existe la carpeta: {ruta_origen}")
