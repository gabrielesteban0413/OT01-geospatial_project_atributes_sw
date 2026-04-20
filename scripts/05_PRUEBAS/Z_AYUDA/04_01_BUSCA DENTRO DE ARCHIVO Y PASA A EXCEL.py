import os
import pandas as pd

directory_path = r"\\10.112.71.143\Smallworld"
palabras_clave = ["ds_date","date_created", "sys!change_info","creation_time"] 
extensiones_validas = ('.magik', '.keymap', '.xml','.dat', 'txt','dmp')

if not os.path.exists(directory_path):
    print(f"La ruta {directory_path} no existe o no es accesible.")
else:
    search_results = []

    for root, dirs, files in os.walk(directory_path):
        for file in files:
            if file.endswith(extensiones_validas):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as archivo:
                        lineas = archivo.readlines()
                        for num_linea, linea in enumerate(lineas, 1):
                            linea_limpia = linea.strip().lower()
                            if any(palabra in linea_limpia for palabra in palabras_clave):
                                search_results.append({
                                    'File Path': file_path,
                                    'Line Number': num_linea,
                                    'Line Content': linea.strip()
                                })
                except Exception as e:
                    print(f"No se pudo leer el archivo {file_path}: {e}")


    if not search_results:
        print("No se encontraron ocurrencias relacionadas.")
    else:
        df = pd.DataFrame(search_results, columns=["File Path", "Line Number", "Line Content"])
        output_path = r"C:\A_GS1_PROYECTOS\reporte2.xlsx"
        df.to_excel(output_path, index=False)
        print(f"Archivo Excel guardado en: {output_path}")
