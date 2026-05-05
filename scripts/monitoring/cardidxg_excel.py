import os
import pandas as pd
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import subprocess

ruta_base = r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo"
ruta_salida = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output"
os.makedirs(ruta_salida, exist_ok=True)

def escanear_dwg_windows_api(ruta):
    archivos = []
    try:
        cmd = f'dir "{ruta}" /s /b *.dwg *.DWG 2>nul'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        archivos = [line.strip() for line in result.stdout.splitlines() if line.strip()]
    except:
        pass
    return archivos

def obtener_metadata_rapido(ruta_archivo):
    try:
        stat = os.stat(ruta_archivo)
        ruta_carpeta = os.path.dirname(ruta_archivo)
        return {
            'ARCHIVO': os.path.basename(ruta_archivo),
            'RUTA_COMPLETA': ruta_archivo,
            'RUTA_CARPETA': ruta_carpeta,
            'TAMAÑO_MB': round(stat.st_size / (1024 * 1024), 2),
            'FECHA_MODIFICACION': datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
            'FECHA_CREACION': datetime.fromtimestamp(stat.st_ctime).strftime('%Y-%m-%d %H:%M:%S')
        }
    except:
        return None

def crear_buscador_batch(df):
    buscador_path = os.path.join(ruta_salida, "BUSCADOR_DWG.bat")
    with open(buscador_path, 'w', encoding='utf-8') as f:
        f.write("@echo off\n")
        f.write("title BUSCADOR DE ARCHIVOS DWG\n")
        f.write("echo Escribe parte del nombre del archivo DWG\n")
        f.write("set /p busqueda=Buscar: \n")
        f.write("set \"tempfile=%TEMP%\\dwg_resultados.txt\"\n")
        f.write(f'> \"%%tempfile%%\" echo Resultados para: %%busqueda%%\n')
        f.write(f'>> \"%%tempfile%%\" echo.\n')
        
        contador = 1
        for idx, row in df.iterrows():
            archivo = row['ARCHIVO']
            ruta_completa = row['RUTA_COMPLETA']
            ruta_carpeta = row['RUTA_CARPETA']
            nombre_sin_extension = archivo.replace('.dwg', '').replace('.DWG', '')
            
            f.write(f'echo "%nombre_sin_extension%" | findstr /i "%%busqueda%%" > nul\n')
            f.write(f'if not errorlevel 1 (\n')
            f.write(f'  echo [%%contador%%] {archivo} >> "%%tempfile%%"\n')
            f.write(f'  echo     {ruta_carpeta} >> "%%tempfile%%"\n')
            f.write(f'  echo. >> "%%tempfile%%"\n')
            f.write(f'  set "archivo_%%contador%%={ruta_completa}"\n')
            f.write(f'  set /a contador+=1\n')
            f.write(f')\n')
        
        f.write('type "%tempfile%" && del "%tempfile%" 2>nul\n')
        f.write('echo.\n')
        f.write('if %contador%==1 (echo No se encontraron archivos & pause & exit)\n')
        f.write('set /p opcion=Selecciona el numero: \n')
        
        for idx, row in df.iterrows():
            ruta_completa = row['RUTA_COMPLETA']
            f.write(f'if "%%opcion%%"=="{idx+1}" start explorer /select, "{ruta_completa}" & exit\n')
        
        f.write('echo Opcion no valida & pause\n')
    return buscador_path

def main():
    archivos_dwg = list(set(escanear_dwg_windows_api(ruta_base)))
    if not archivos_dwg:
        return
    
    datos = []
    with ThreadPoolExecutor(max_workers=20) as executor:
        futures = {executor.submit(obtener_metadata_rapido, f): f for f in archivos_dwg}
        for future in as_completed(futures):
            resultado = future.result()
            if resultado:
                datos.append(resultado)
    
    df = pd.DataFrame(datos)
    df = df.drop_duplicates(subset=['RUTA_COMPLETA'])
    df = df.sort_values('ARCHIVO')
    
    ruta_excel = os.path.join(ruta_salida, "03_cardidwg_excel.xlsx")
    with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Archivos_DWG', index=False)
    
    buscador_path = crear_buscador_batch(df)
    escritorio = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    with open(os.path.join(escritorio, "BUSCADOR_DWG.bat"), 'w', encoding='utf-8') as f:
        f.write(f'@start "" "{buscador_path}"')
    
    print(f" Proceso completado: {len(df)} archivos DWG encontrados")

if __name__ == "__main__":
    main()