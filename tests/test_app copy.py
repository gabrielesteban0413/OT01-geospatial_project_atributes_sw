import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
from collections import defaultdict
import re

# ========== CONFIGURACIÓN ==========
# Conexión a PostgreSQL
conn = psycopg2.connect(
    host="localhost",
    port="5432",
    database="GS01_DBASPHIA",
    user="postgres",
    password="Feelgs44#4"  # <--- CAMBIA AQUÍ
)

# ========== 1. CARGAR DATOS DESDE POSTGRESQL ==========
print("📖 Cargando datos desde PostgreSQL...")

query = """
SELECT 
    idservicio, 
    cambio, 
    n_acceso as fibra, 
    anillo, 
    actividad, 
    fecha,
    idcliente,
    cliente,
    direccion,
    ciudad
FROM clientes_servicios 
WHERE idservicio IS NOT NULL AND idservicio != 'None'
ORDER BY idservicio, fecha
"""

df = pd.read_sql(query, conn)
conn.close()

print(f"✅ Cargadas {len(df)} filas")
print(f"📊 Servicios únicos: {df['idservicio'].nunique()}")

# ========== 2. FUNCIONES DE DETECCIÓN DE CASOS ATÍPICOS ==========

def detectar_casos_atipicos(grupo):
    """
    Detecta casos atípicos para un grupo de un mismo idservicio
    Retorna: (observacion, actividad_final, fecha_final, fibra_final, cambio_final)
    """
    observaciones = []
    
    # Ordenar por fecha
    grupo = grupo.sort_values('fecha')
    actividades = grupo['actividad'].tolist()
    fechas = grupo['fecha'].tolist()
    fibras = grupo['fibra'].tolist()
    cambios = grupo['cambio'].tolist()
    
    # Último registro
    ultimo = grupo.iloc[-1]
    actividad_final = ultimo['actividad']
    fecha_final = ultimo['fecha']
    fibra_final = ultimo['fibra']
    cambio_final = ultimo['cambio']
    
    # --- CASO 1: ROLL BACK ---
    if 'ROLL BACK' in actividades:
        # Verificar si el último evento es ROLL BACK
        if actividad_final == 'ROLL BACK':
            # Buscar el evento anterior al ROLL BACK
            if len(grupo) > 1:
                evento_anterior = grupo.iloc[-2]
                actividad_final = evento_anterior['actividad']
                fecha_final = evento_anterior['fecha']
                fibra_final = evento_anterior['fibra']
                cambio_final = evento_anterior['cambio']
                observaciones.append("ROLL_BACK: el estado final es el anterior al rollback")
        else:
            observaciones.append("ROLL_BACK_DETECTADO_EN_HISTORIAL")
    
    # --- CASO 2: Múltiples instalaciones con diferente fibra ---
    instalaciones = grupo[grupo['actividad'] == 'INSTALACION']
    if len(instalaciones) > 1:
        fibras_instalacion = instalaciones['fibra'].tolist()
        if len(set(fibras_instalacion)) > 1:
            observaciones.append(f"MULTIPLES_INSTALACIONES_DIFERENTE_FIBRA: {len(instalaciones)} instalaciones con fibras {fibras_instalacion}")
    
    # --- CASO 3: Instalación después de Desprogramación ---
    fechas_lista = grupo['fecha'].tolist()
    actividades_lista = grupo['actividad'].tolist()
    for i in range(1, len(grupo)):
        if actividades_lista[i] == 'INSTALACION' and actividades_lista[i-1] == 'DESPROGRAMACION':
            observaciones.append(f"INSTALACION_DESPUES_DESPROGRAMACION: reactivación el {fechas_lista[i]}")
            break
    
    # --- CASO 4: Fibra cambiada sin RETIRO ---
    fibras_unicas = grupo['fibra'].unique()
    if len(fibras_unicas) > 1:
        # Verificar si hay RETIRO entre los cambios de fibra
        tiene_retiro = 'RETIRO' in actividades
        if not tiene_retiro:
            observaciones.append(f"FIBRA_CAMBIADA_SIN_RETIRO: fibras {list(fibras_unicas)}")
    
    # --- CASO 5: Secuencia de actividad inusual ---
    # Normal: INSTALACION -> ADICION -> AMPLIACION -> RETIRO -> DESPROGRAMACION
    # Detectar si hay RETIRO sin INSTALACION previa, etc.
    if 'RETIRO' in actividades and 'INSTALACION' not in actividades:
        observaciones.append("RETIRO_SIN_INSTALACION_PREVIA")
    
    if 'ADICION' in actividades and 'INSTALACION' not in actividades:
        if 'INSTALACION' not in actividades:
            observaciones.append("ADICION_SIN_INSTALACION_PREVIA")
    
    # --- CASO 6: Misma fecha, diferentes actividades (posible error) ---
    fechas_duplicadas = grupo[grupo.duplicated(subset=['fecha'], keep=False)]
    if len(fechas_duplicadas) > 0:
        act_misma_fecha = fechas_duplicadas['actividad'].tolist()
        if len(set(act_misma_fecha)) > 1:
            observaciones.append(f"MULTIPLES_ACTIVIDADES_MISMA_FECHA: {act_misma_fecha}")
    
    # Construir observación final
    if observaciones:
        observacion_final = " | ".join(observaciones)
    else:
        observacion_final = "SIN_CASOS_ATIPICOS"
    
    # Construir historial de actividades
    historial = " → ".join(actividades)
    
    return {
        'observacion': observacion_final,
        'actividad_final': actividad_final,
        'fecha_final': fecha_final,
        'fibra_final': fibra_final,
        'cambio_final': cambio_final,
        'historial': historial,
        'total_eventos': len(grupo)
    }

# ========== 3. PROCESAR POR GRUPO DE IDSERVICIO ==========
print("🔄 Procesando servicios (puede tomar unos segundos)...")

resultados = []
servicios = df['idservicio'].unique()
total_servicios = len(servicios)

for i, servicio in enumerate(servicios):
    if (i + 1) % 1000 == 0:
        print(f"   Procesados {i + 1} de {total_servicios} servicios...")
    
    grupo = df[df['idservicio'] == servicio].copy()
    analisis = detectar_casos_atipicos(grupo)
    
    # Tomar datos del último registro (o el ajustado por rollback)
    ultimo_registro = grupo[grupo['fecha'] == analisis['fecha_final']]
    if len(ultimo_registro) == 0:
        # Si no encuentra por fecha exacta, tomar el último
        ultimo_registro = grupo.iloc[-1:]
    
    resultados.append({
        'idservicio': servicio,
        'cambio': analisis['cambio_final'],
        'fibra': analisis['fibra_final'],
        'anillo': ultimo_registro['anillo'].iloc[0] if len(ultimo_registro) > 0 else None,
        'actividad_final': analisis['actividad_final'],
        'fecha_final': analisis['fecha_final'],
        'historial_actividades': analisis['historial'],
        'total_eventos': analisis['total_eventos'],
        'observacion_caso_atipico': analisis['observacion'],
        'cliente': ultimo_registro['cliente'].iloc[0] if len(ultimo_registro) > 0 else None,
        'direccion': ultimo_registro['direccion'].iloc[0] if len(ultimo_registro) > 0 else None,
        'ciudad': ultimo_registro['ciudad'].iloc[0] if len(ultimo_registro) > 0 else None
    })

# ========== 4. CREAR DATAFRAME FINAL ==========
df_resultado = pd.DataFrame(resultados)

print(f"\n✅ Procesados {len(df_resultado)} servicios únicos")

# ========== 5. GUARDAR RESULTADOS ==========
# Guardar a CSV
df_resultado.to_csv('servicios_depurados.csv', index=False, encoding='utf-8-sig')
print("📁 Archivo guardado: servicios_depurados.csv")

# Guardar a Excel (si no es muy grande)
if len(df_resultado) < 100000:
    df_resultado.to_excel('servicios_depurados.xlsx', index=False)
    print("📁 Archivo guardado: servicios_depurados.xlsx")

# ========== 6. MOSTRAR ESTADÍSTICAS DE CASOS ATÍPICOS ==========
print("\n" + "="*60)
print("📊 ESTADÍSTICAS DE CASOS ATÍPICOS")
print("="*60)

casos_counts = df_resultado['observacion_caso_atipico'].value_counts()
print("\nTipos de casos detectados:")
for caso, count in casos_counts.head(10).items():
    print(f"   {caso}: {count} servicios")

# Mostrar servicios con casos atípicos
casos_atipicos = df_resultado[df_resultado['observacion_caso_atipico'] != 'SIN_CASOS_ATIPICOS']
print(f"\n🔍 Servicios con casos atípicos: {len(casos_atipicos)} de {len(df_resultado)}")

if len(casos_atipicos) > 0:
    print("\n📋 Ejemplos de servicios con casos atípicos:")
    print(casos_atipicos[['idservicio', 'actividad_final', 'total_eventos', 'observacion_caso_atipico']].head(10).to_string(index=False))

# ========== 7. MOSTRAR EL CASO QUE DETECTASTE ==========
print("\n" + "="*60)
print("🔍 VERIFICACIÓN DEL CASO CCCH00013942")
print("="*60)

caso_especifico = df_resultado[df_resultado['idservicio'] == 'CCCH00013942']
if len(caso_especifico) > 0:
    print(caso_especifico[['idservicio', 'fibra', 'actividad_final', 'fecha_final', 'historial_actividades', 'observacion_caso_atipico']].to_string(index=False))
else:
    print("No se encontró el servicio CCCH00013942 en los datos")

print("\n🎉 PROCESO COMPLETADO!")



