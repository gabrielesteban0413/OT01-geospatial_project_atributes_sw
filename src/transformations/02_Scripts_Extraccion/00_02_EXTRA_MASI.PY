import os
import shutil

origen = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Shapefile"
destino = r"C:\A_GS1_PROYECTOS"
ruta_autocad = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Autocad"


archivos = os.listdir(origen)
def obtener_nombre_base(nombre_archivo):
    nombre_sin_ext = nombre_archivo

    while '.' in nombre_sin_ext:
        nombre_sin_ext = os.path.splitext(nombre_sin_ext)[0]
 
    if nombre_sin_ext.endswith('_T'):
        nombre_sin_ext = nombre_sin_ext[:-2]
    return nombre_sin_ext


shapefiles = {} # Agrupar archivos por nombre base sin sufijo _T
for archivo in archivos:
    nombre_base = obtener_nombre_base(archivo)
    if nombre_base not in shapefiles:
        shapefiles[nombre_base] = []
    shapefiles[nombre_base].append(archivo)


for nombre_base, lista_archivos in shapefiles.items(): # Mover archivos a carpetas
    nueva_ruta = os.path.join(destino, nombre_base)
    os.makedirs(nueva_ruta, exist_ok=True)  
    
    for archivo in lista_archivos:
        origen_archivo = os.path.join(origen, archivo)
        destino_archivo = os.path.join(nueva_ruta, archivo)
        shutil.move(origen_archivo, destino_archivo)


if os.path.exists(ruta_autocad): # Borrar todos los archivos de la carpeta Autocad
    for archivo in os.listdir(ruta_autocad):
        ruta_completa = os.path.join(ruta_autocad, archivo)
        if os.path.isfile(ruta_completa):
            os.remove(ruta_completa)

print("Archivos movidos correctamente y carpeta Autocad limpiada.")
