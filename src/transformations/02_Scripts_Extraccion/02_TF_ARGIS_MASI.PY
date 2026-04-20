import arcpy
import os

carpeta_cad      = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Autocad"
gdb_intermedia   = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\GDB_paso.gdb"
salida_shp       = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Shapefile"

dataset_mcb      = "Salida"
dataset_mcbog    = "Salida_Project"
ruta_dataset_mcb    = os.path.join(gdb_intermedia, dataset_mcb)
ruta_dataset_mcbog  = os.path.join(gdb_intermedia, dataset_mcbog)

sr_mcbog = arcpy.SpatialReference(3116)  # MAGNA Bogotá

# WKT de MAGNA Ciudad Bogotá 
sr_wkt_ciudad_bogota = """PROJCS["MAGNA_Ciudad_Bogota",GEOGCS["GCS_MAGNA",DATUM["D_MAGNA",SPHEROID["GRS_1980",6378137.0,298.257222101]],PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]],PROJECTION["Transverse_Mercator"],PARAMETER["False_Easting",92334.879],PARAMETER["False_Northing",109320.965],PARAMETER["Central_Meridian",-74.14659166666668],PARAMETER["Scale_Factor",1.000399803265436],PARAMETER["Latitude_Of_Origin",4.680486111111112],UNIT["Meter",1.0]]"""
sr_ciudad_bogota = arcpy.SpatialReference()
sr_ciudad_bogota.loadFromString(sr_wkt_ciudad_bogota)

archivos_fuera_colombia = []

for archivo in os.listdir(carpeta_cad):
    if not archivo.lower().endswith((".dwg", ".dxf")):
        continue

    ruta_cad     = os.path.join(carpeta_cad, archivo)
    nombre_base  = os.path.splitext(archivo)[0].replace(" ", "").replace("-", "").replace(".", "_")

    try:
        arcpy.CADToGeodatabase_conversion(
            ruta_cad,
            gdb_intermedia,
            dataset_mcb,
            1000,
            sr_ciudad_bogota
        )

        feature_polyline   = os.path.join(ruta_dataset_mcb, "Polyline")
        feature_intermedio = os.path.join(ruta_dataset_mcb, f"{nombre_base}_mcb")
        arcpy.CopyFeatures_management(feature_polyline, feature_intermedio)

        feature_final = os.path.join(ruta_dataset_mcbog, f"{nombre_base}_mcbog")
        arcpy.Project_management(feature_intermedio, feature_final, sr_mcbog)

        salida_shapefile = os.path.join(salida_shp, f"{nombre_base}.shp")
        arcpy.CopyFeatures_management(feature_final, salida_shapefile)

        # por geometrías = ubicación
        fuera_colombia = False
        with arcpy.da.SearchCursor(feature_final, ["SHAPE@XY"]) as cursor:
            for row in cursor:
                x, y = row[0]
                if not (200000 <= x <= 1200000 and 200000 <= y <= 1700000):
                    fuera_colombia = True
                    break

        if fuera_colombia:
            print(f"{archivo} ERROR\n")
            nombre_limpio = os.path.splitext(archivo)[0].replace("_T", "")
            archivos_fuera_colombia.append(nombre_limpio)

            # Eliminar shapefile si está fuera de Colombia
            for ext in [".shp", ".shx", ".dbf", ".prj", ".cpg", ".sbn", ".sbx", ".fbn", ".fbx", ".ain", ".aih", ".ixs", ".mxs", ".atx"]:
                archivo_completo = os.path.join(salida_shp, f"{nombre_base}{ext}")
                if os.path.exists(archivo_completo):
                    os.remove(archivo_completo)
        else:
            print(f"{archivo}\n")

    except Exception as e:
        print(f"ERROR ------ {archivo}: {e}\n")

print("-------------------------TF---------------------------")

if archivos_fuera_colombia:
    print("MAL GEOREFERENCIADO:\n")
    for nombre in archivos_fuera_colombia:
        print(nombre)
else:
    print("BIEN GEOREFERENCIADO.")





#ELIMINAR MANZANAS ------------


import arcpy
import os
carpeta_shapefiles = r"C:\Users\gabrboas\Desktop\PERSONAL GIS\INSUMOS\Shapefile"

for archivo in os.listdir(carpeta_shapefiles):
    if archivo.lower().endswith(".shp"):
        ruta_shp = os.path.join(carpeta_shapefiles, archivo)
        try:
            contador = 0
            with arcpy.da.UpdateCursor(ruta_shp, ["Layer"], "Layer LIKE 'MA%'") as cursor:
                for row in cursor:
                    cursor.deleteRow()
                    contador += 1
            print(f"{archivo}: Eliminadas {contador} entidades con Layer LIKE 'MA%'")
        except Exception as e:
            print(f"ERROR en {archivo}: {e}")

print("\n---------------------- MA%----------------------")
