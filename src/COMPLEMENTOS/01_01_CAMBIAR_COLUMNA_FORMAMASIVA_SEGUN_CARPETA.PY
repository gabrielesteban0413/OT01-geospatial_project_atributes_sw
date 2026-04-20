import os
import pandas as pd

# Ruta de la carpeta con los archivos Excel
ruta_carpeta = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\12_Cables_troncales_ conectados_en central"

# Función para procesar cada archivo Excel
def procesar_excel(archivo_path):
    print(f"Procesando: {archivo_path}")
    
    # Cargar el archivo Excel
    try:
        xlsx = pd.ExcelFile(archivo_path, engine='openpyxl')
    except Exception as e:
        print(f"  ERROR al abrir archivo: {e}")
        return
    
    modificado = False
    
    # Procesar cada hoja del archivo
    for hoja in xlsx.sheet_names:
        # Verificar si el nombre de la hoja contiene un guion '-'
        if '-' in hoja:
            print(f"  - Hoja '{hoja}': Contiene '-' - Se cambiará el encabezado de columna G")
            
            try:
                df = pd.read_excel(archivo_path, sheet_name=hoja, engine='openpyxl')
                
                # Verificar si la columna G existe (índice 6)
                if df.shape[1] > 6:
                    nombre_original = df.columns[6]
                    # Cambiar el encabezado de la columna G
                    df.rename(columns={df.columns[6]: "ANILLO ORIGEN"}, inplace=True)
                    
                    # Guardar el archivo modificado
                    with pd.ExcelWriter(archivo_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                        df.to_excel(writer, sheet_name=hoja, index=False)
                    
                    print(f"    ✓ Columna G ('{nombre_original}') cambiada a 'ANILLO ORIGEN'")
                    modificado = True
                else:
                    print(f"    ✗ La hoja no tiene columna G (solo {df.shape[1]} columnas)")
                    
            except Exception as e:
                print(f"    ✗ Error al procesar: {e}")
        else:
            print(f"  - Hoja '{hoja}': No contiene '-' - Se omite")
    
    if not modificado:
        print(f"  No se modificó ninguna hoja en este archivo")

# Recorrer todos los archivos Excel en la carpeta
archivos_procesados = 0
for archivo in os.listdir(ruta_carpeta):
    if archivo.endswith(('.xlsx', '.xls')):
        ruta_completa = os.path.join(ruta_carpeta, archivo)
        procesar_excel(ruta_completa)
        archivos_procesados += 1
        print()  # Línea en blanco para separar archivos

print(f"Proceso completado. {archivos_procesados} archivos procesados.")