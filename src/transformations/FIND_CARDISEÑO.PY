# ============================================================================
import os
import time
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path
import warnings
warnings.filterwarnings('ignore')

# ============================================================================
RUTA_BUSQUEDA = r"\\atlas.etb.com.co\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo"
RUTA_SALIDA = Path(r"C:\A_GS1_PROYECTOS")

MAX_WORKERS = 10
PROGRESO_INTERVALO = 1000
# CAMBIO PRINCIPAL: Solo estas columnas
COLUMNAS_SALIDA = ['ARCHIVO', 'RUTA_CARPETA', 'TAMAÑO_MB', 'FECHA_MODIFICACION', 'FECHA_CREACION']

# ============================================================================
# FUNCIONES DE UTILIDAD
# ============================================================================

def verificar_ruta_accesible(ruta):
    """Verifica si una ruta existe y es accesible."""
    try:
        return os.path.exists(ruta)
    except:
        return False

def mostrar_progreso(contador, total, etapa):
    """Muestra progreso en la consola."""
    if contador % PROGRESO_INTERVALO == 0:
        porcentaje = (contador / total) * 100 if total > 0 else 0
        print(f"   {etapa}: {contador:,} / {total:,} ({porcentaje:.1f}%)")

def formatear_fecha(timestamp):
    """Formatea timestamp a string legible."""
    try:
        return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
    except:
        return "N/A"

def obtener_ruta_carpeta(ruta_completa):
    """Obtiene solo la ruta de la carpeta (sin el nombre del archivo)."""
    return os.path.dirname(ruta_completa)

# ============================================================================
# FUNCIONES DE OBTENCIÓN DE ARCHIVOS
# ============================================================================

def obtener_lista_dwg_rapido(ruta_busqueda):
    """
    Obtiene la lista completa de archivos DWG de forma optimizada.
    """
    rutas_dwg = []
    contador = 0
    
    print("Fase 1: Recorriendo estructura de carpetas...")
    
    try:
        for root, dirs, files in os.walk(ruta_busqueda):
            for file in files:
                if file.lower().endswith('.dwg'):
                    ruta_completa = os.path.join(root, file)
                    rutas_dwg.append(ruta_completa)
                    contador += 1
                    
                    if contador % PROGRESO_INTERVALO == 0:
                        print(f"   DWG encontrados: {contador:,}...")
                        
    except Exception as e:
        print(f"Advertencia: Error durante el recorrido: {e}")
    
    return rutas_dwg

# ============================================================================
# FUNCIONES DE PROCESAMIENTO
# ============================================================================

def obtener_info_carpeta_archivo(ruta_archivo):
    """
    Obtiene información del archivo CON RUTA_CARPETA en lugar de RUTA_COMPLETA.
    """
    try:
        stat = os.stat(ruta_archivo)
        return {
            'ARCHIVO': os.path.basename(ruta_archivo),
            'RUTA_CARPETA': obtener_ruta_carpeta(ruta_archivo),  # SOLO CARPETA
            'TAMAÑO_MB': round(stat.st_size / (1024 * 1024), 2),
            'FECHA_MODIFICACION': formatear_fecha(stat.st_mtime),
            'FECHA_CREACION': formatear_fecha(stat.st_ctime)
        }
    except Exception as e:
        # Información mínima si hay error
        return {
            'ARCHIVO': os.path.basename(ruta_archivo),
            'RUTA_CARPETA': "ERROR",  # SOLO CARPETA
            'TAMAÑO_MB': 0,
            'FECHA_MODIFICACION': "ERROR",
            'FECHA_CREACION': "ERROR"
        }

def procesar_archivos_paralelo_rapido(rutas_dwg):
    """
    Procesa archivos en paralelo de forma optimizada.
    """
    datos_procesados = []
    total_archivos = len(rutas_dwg)
    
    if total_archivos == 0:
        return datos_procesados
    
    print(f"Fase 2: Procesando {total_archivos:,} archivos...")
    
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
        # Enviar todas las tareas
        futures = {executor.submit(obtener_info_carpeta_archivo, ruta): ruta 
                  for ruta in rutas_dwg}
        
        # Recoger resultados
        procesados = 0
        for future in as_completed(futures):
            try:
                datos = future.result()
                datos_procesados.append(datos)
                procesados += 1
                mostrar_progreso(procesados, total_archivos, "Procesados")
            except Exception:
                procesados += 1
                mostrar_progreso(procesados, total_archivos, "Procesados")
    
    return datos_procesados

# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def ejecutar_inventario_dwg_carpetas():
    """
    Función principal que solo obtiene RUTA_CARPETA en lugar de RUTA_COMPLETA.
    """
    print("=" * 70)
    print("INVENTARIO DE ARCHIVOS DWG (SOLO CARPETAS)")
    print("=" * 70)
    print(f"Ruta de búsqueda: {RUTA_BUSQUEDA}")
    print(f"Carpeta de salida: {RUTA_SALIDA}")
    print("=" * 70)
    
    inicio_total = time.time()
    
    # Verificar ruta
    if not verificar_ruta_accesible(RUTA_BUSQUEDA):
        print(f"ERROR: La ruta no es accesible")
        print(f"Verifique: {RUTA_BUSQUEDA}")
        return
    
    # Crear carpeta de salida
    RUTA_SALIDA.mkdir(parents=True, exist_ok=True)
    
    # 1. Obtener lista de archivos
    print("\n[1/3] Buscando archivos DWG...")
    inicio_busqueda = time.time()
    rutas_dwg = obtener_lista_dwg_rapido(RUTA_BUSQUEDA)
    tiempo_busqueda = time.time() - inicio_busqueda
    
    if not rutas_dwg:
        print("No se encontraron archivos DWG.")
        return
    
    print(f"✓ Encontrados: {len(rutas_dwg):,} archivos DWG")
    print(f"⏱️  Tiempo de búsqueda: {tiempo_busqueda:.2f} segundos")
    
    # 2. Procesar archivos (solo carpeta)
    print(f"\n[2/3] Obteniendo información de carpetas...")
    inicio_proceso = time.time()
    datos_procesados = procesar_archivos_paralelo_rapido(rutas_dwg)
    tiempo_proceso = time.time() - inicio_proceso
    
    if not datos_procesados:
        print("ERROR: No se pudo procesar la información.")
        return
    
    # 3. Crear DataFrame
    df = pd.DataFrame(datos_procesados)
    
    # Ordenar por carpeta y luego por archivo
    df = df.sort_values(['RUTA_CARPETA', 'ARCHIVO'])
    
    # 4. Guardar en Excel
    print(f"\n[3/3] Guardando en Excel...")
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_archivo = f"inventario_dwg_carpetas_{timestamp}.xlsx"
    ruta_excel = RUTA_SALIDA / nombre_archivo
    
    try:
        # Crear Excel
        with pd.ExcelWriter(ruta_excel, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='TODOS_LOS_DWG', index=False)
            
            # Ajustar ancho de columnas
            worksheet = writer.sheets['TODOS_LOS_DWG']
            
            # Archivo (col A)
            worksheet.column_dimensions['A'].width = 40
            
            # Ruta Carpeta (col B) - más ancha para mostrar carpetas
            worksheet.column_dimensions['B'].width = 60
            
            # Tamaño MB (col C)
            worksheet.column_dimensions['C'].width = 15
            
            # Fechas (cols D y E)
            worksheet.column_dimensions['D'].width = 20
            worksheet.column_dimensions['E'].width = 20
        
        tiempo_total = time.time() - inicio_total
        
        # Mostrar resumen
        print("\n" + "=" * 70)
        print("✅ INVENTARIO COMPLETADO")
        print("=" * 70)
        print(f"\n📊 RESUMEN:")
        print(f"   Archivos procesados: {len(df):,}")
        print(f"   Carpetas únicas: {df['RUTA_CARPETA'].nunique():,}")
        print(f"   Tiempo total: {tiempo_total:.2f} segundos")
        
        # Top carpetas con más archivos
        print(f"\n📁 TOP 5 CARPETAS CON MÁS ARCHIVOS:")
        top_carpetas = df['RUTA_CARPETA'].value_counts().head(5)
        for carpeta, cantidad in top_carpetas.items():
            nombre_carpeta = os.path.basename(carpeta) if carpeta != "ERROR" else "ERROR"
            print(f"   {nombre_carpeta[:30]:30} : {cantidad:4} archivos")
        
        # Vista previa
        print(f"\n👁️  VISTA PREVIA (primeros 5):")
        for i, (_, fila) in enumerate(df.head(5).iterrows(), 1):
            carpeta_base = os.path.basename(fila['RUTA_CARPETA']) if fila['RUTA_CARPETA'] != "ERROR" else "ERROR"
            print(f"   {i}. {fila['ARCHIVO'][:25]:25} | Carpeta: {carpeta_base[:20]:20}")
        
        print(f"\n💾 ARCHIVO GENERADO:")
        print(f"   📄 {ruta_excel}")
        
        # Intentar abrir Excel
        try:
            os.startfile(ruta_excel)
            print(f"   📂 Archivo abierto automáticamente")
        except:
            print(f"   📂 Para abrir manualmente: {ruta_excel}")
            
    except Exception as e:
        print(f"\n❌ ERROR al guardar Excel: {e}")
        # Guardar como CSV
        try:
            csv_file = ruta_excel.with_suffix('.csv')
            df.to_csv(csv_file, index=False, sep=';', encoding='utf-8')
            print(f"✓ Datos guardados como CSV: {csv_file}")
            os.startfile(csv_file)
        except Exception as e2:
            print(f"❌ También falló CSV: {e2}")
    
    print("\n" + "=" * 70)

# ============================================================================
# VERSIÓN ULTRA SIMPLE (solo carpeta y archivo)
# ============================================================================

def generar_lista_simple_carpetas():
    """
    Versión ultra simple: solo archivo y carpeta.
    """
    print("=" * 70)
    print("LISTA SIMPLE: ARCHIVO Y CARPETA")
    print("=" * 70)
    
    inicio = time.time()
    
    if not verificar_ruta_accesible(RUTA_BUSQUEDA):
        print(f"ERROR: Ruta no accesible")
        return
    
    # Crear carpeta de salida
    RUTA_SALIDA.mkdir(parents=True, exist_ok=True)
    
    # Recopilar datos
    datos = []
    contador = 0
    
    print("Recopilando información...")
    
    try:
        for root, dirs, files in os.walk(RUTA_BUSQUEDA):
            for file in files:
                if file.lower().endswith('.dwg'):
                    datos.append({
                        'ARCHIVO': file,
                        'CARPETA': root  # Solo la carpeta, no la ruta completa
                    })
                    contador += 1
                    
                    if contador % 5000 == 0:
                        print(f"   Procesados: {contador:,}...")
    except Exception as e:
        print(f"Error: {e}")
    
    if not datos:
        print("No se encontraron archivos DWG.")
        return
    
    # Crear DataFrame
    df = pd.DataFrame(datos)
    
    # Ordenar por carpeta y archivo
    df = df.sort_values(['CARPETA', 'ARCHIVO'])
    
    # Guardar
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    nombre_csv = f"lista_simple_carpetas_{timestamp}.csv"
    ruta_csv = RUTA_SALIDA / nombre_csv
    
    df.to_csv(ruta_csv, index=False, sep=';', encoding='utf-8')
    
    tiempo = time.time() - inicio
    
    print(f"\n✅ COMPLETADO")
    print(f"   Archivos: {len(df):,}")
    print(f"   Carpetas únicas: {df['CARPETA'].nunique():,}")
    print(f"   Tiempo: {tiempo:.2f} segundos")
    print(f"   Archivo: {ruta_csv}")
    
    # Mostrar muestra
    print(f"\n📋 MUESTRA (primeros 5):")
    for i, (_, fila) in enumerate(df.head(5).iterrows(), 1):
        carpeta_base = os.path.basename(fila['CARPETA'])
        print(f"   {i}. {fila['ARCHIVO']}")
        print(f"      📁 {carpeta_base}")
    
    # Abrir carpeta
    os.startfile(RUTA_SALIDA)

# ============================================================================
# EJECUCIÓN
# ============================================================================

if __name__ == "__main__":
    # Ejecutar inventario con carpetas
    ejecutar_inventario_dwg_carpetas()
    
    # O ejecutar versión simple
    # generar_lista_simple_carpetas()
    
    # Mantener consola abierta
    input("\nPresiona Enter para salir...")