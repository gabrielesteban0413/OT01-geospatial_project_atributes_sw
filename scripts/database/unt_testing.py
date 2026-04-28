import pandas as pd
import re
from pathlib import Path

# Datos del registro específico de IS
datos_is = {
    'FIBRA': '70907',
    'EC': 'BOBAHU930602',
    'PUERTO C': '3/0/0',
    'ODF P': '5',
    'CASSETERA': 'A',
    'PATCHEO': '5-1A',
    'BANDEJA': '1',
    'ODF CALLE': '33B',
    'HILO': '64',
    'DESCRIPTION PORT OT': '16',
    'DESCRIPCION SHELF OT': 'FO49/96',
    'ID_PORT_CABECERA': '3492000048808684396',
    'ID_PORT_PACHEO': '3492000048808725495',
    'ID_PORT_ODF-TRONCAL': '',
    'SINCRONIZA': 'V',
    'ANILLO ORIGEN': 'TO_DEM_H64_70907_G1_MEADALBOBA02(10.252.91.171)',
    'ANILLO': 'MEADALBOBA02',
    'CAMBIO': 'V|MEADALBOBA02|C572197|2016-12-06'
}

# Función para limpiar puertos
def limpiar_puerto_completo(valor):
    if pd.isna(valor):
        return valor
    cadena = re.sub(r'[^0-9/-]', '', str(valor))
    if cadena.endswith('/'):
        cadena = cadena[:-1]
    partes = cadena.split('/')
    partes_limpias = [p.lstrip('0') or '0' for p in partes]
    return '/'.join(partes_limpias)

def analizar_id_odf_troncal():
    print("="*80)
    print("ANÁLISIS DE POR QUÉ NO SE GENERA ID_PORT_ODF-TRONCAL")
    print("="*80)
    
    # Paso 1: Extraer datos del registro
    fibra = datos_is['FIBRA']
    odf_calle = datos_is['ODF CALLE']
    hilo = datos_is['HILO']
    
    print(f"\n1. DATOS DEL REGISTRO:")
    print(f"   FIBRA (cable): '{fibra}'")
    print(f"   ODF CALLE: '{odf_calle}'")
    print(f"   HILO: '{hilo}'")
    
    # Paso 2: Verificar condiciones necesarias
    print(f"\n2. VERIFICANDO CONDICIONES NECESARIAS:")
    
    condiciones = {
        'ot_mapping existe?': 'Pendiente (depende de datos OT)',
        'ot_base existe?': 'Pendiente (depende de datos OT)',
        'FIBRA no es NaN?': pd.notna(fibra),
        'ODF CALLE no es NaN?': pd.notna(odf_calle),
        'HILO no es NaN?': pd.notna(hilo),
        'HILO es dígito?': str(hilo).isdigit()
    }
    
    for cond, valor in condiciones.items():
        if isinstance(valor, bool):
            print(f"   - {cond}: {'✅ SI' if valor else '❌ NO'}")
        else:
            print(f"   - {cond}: {valor}")
    
    # Paso 3: Análisis del HILO
    print(f"\n3. ANÁLISIS DEL HILO:")
    hilo_str = str(hilo).strip()
    hilo_int = int(hilo_str) if hilo_str.isdigit() else None
    
    print(f"   HILO original: '{hilo_str}'")
    print(f"   ¿Es número?: {hilo_str.isdigit()}")
    if hilo_int:
        print(f"   Como entero: {hilo_int}")
    
    # Paso 4: Simular lo que haría el código con OT
    print(f"\n4. SIMULACIÓN DE BÚSQUEDA EN OT:")
    
    # Aquí normalmente se buscaría en ot_mapping y ot_base
    # Como no tenemos el archivo real, simulamos las posibles causas
    
    print(f"\n   POSIBLES CAUSAS DEL FALLOS:")
    print(f"   {'='*60}")
    
    causas = []
    
    # Causa 1: El archivo OT no tiene datos para esta fibra y ODF
    causas.append({
        'causa': 'No existe registro en OT para la combinación (FIBRA, ODF CALLE)',
        'detalle': f"FIBRA='{fibra}', ODF CALLE='{odf_calle}'",
        'solucion': f"Verificar si en la hoja OT existe un registro con FIBRA='{fibra}' y NOMBRE ODF='{odf_calle}'"
    })
    
    # Causa 2: El HILO no es numérico
    if not hilo_str.isdigit():
        causas.append({
            'causa': 'El HILO no es un número válido',
            'detalle': f"HILO='{hilo_str}' - Debe ser un número entero",
            'solucion': 'Corregir el valor de HILO en el archivo origen'
        })
    
    # Causa 3: El cálculo del puerto está fuera de rango
    causas.append({
        'causa': 'El cálculo (HILO - BASE) + 1 podría dar un número inválido',
        'detalle': f"HILO={hilo_int}, se necesita BASE para calcular PORT = (HILO - BASE) + 1",
        'solucion': 'Verificar que la BASE en OT sea correcta y que HILO >= BASE'
    })
    
    # Causa 4: El puerto calculado no existe en ot_mapping
    causas.append({
        'causa': 'El puerto calculado no existe en el mapeo OT',
        'detalle': 'Una vez calculado el PORT, debe existir en ot_mapping[(FIBRA, ODF, PORT)]',
        'solucion': 'Verificar que en OT exista el puerto específico'
    })
    
    # Causa 5: Problemas con el formato de nombres
    causas.append({
        'causa': 'Inconsistencia en nombres (mayúsculas/minúsculas, espacios)',
        'detalle': f"ODF CALLE='{odf_calle}' - Puede tener espacios o mayúsculas diferentes",
        'solucion': 'Normalizar nombres: .strip() y .upper()'
    })
    
    for i, causa in enumerate(causas, 1):
        print(f"\n   CAUSA {i}: {causa['causa']}")
        print(f"   └─ Detalle: {causa['detalle']}")
        print(f"   └─ Solución: {causa['solucion']}")
    
    # Paso 5: Validación específica para este caso
    print(f"\n5. VALIDACIÓN ESPECÍFICA PARA ESTE CASO:")
    print(f"   {'='*60}")
    
    # Analizar el formato de ODF CALLE
    print(f"\n   a) ODF CALLE = '{odf_calle}'")
    if 'B' in odf_calle:
        print(f"      → Contiene 'B' (posible bandeja o cassetera)")
        print(f"      → En el código, ot_base usa solo el nombre ODF sin sufijos")
        nombre_odf_limpio = re.sub(r'[A-Za-z]', '', odf_calle)
        print(f"      → Nombre limpio sugerido: '{nombre_odf_limpio}'")
    
    # Analizar el HILO
    print(f"\n   b) HILO = {hilo}")
    if hilo_int and hilo_int < 10:
        print(f"      → HILO es pequeño ({hilo_int})")
        print(f"      → Posiblemente la BASE en OT sea > {hilo_int}")
    
    # Recomendación final
    print(f"\n6. RECOMENDACIONES PARA DEPURAR:")
    print(f"   {'='*60}")
    print(f"""
    Para resolver este problema, NECESITAS VERIFICAR EN EL ARCHIVO OT:

    1. Buscar en la hoja OT (del archivo Port.xlsx):
       - ¿Existe un registro con FIBRA = '{fibra}'?
       - ¿Existe un registro con NOMBRE ODF = '{odf_calle}'?
       
    2. Si existe, verificar el valor de DESCRIPCION SHELF:
       - Debe tener formato: OT{numero}_{nombre}
       - Ejemplo: OT33B_33B 1/2 (donde el número ANTES de la / es la BASE)
       
    3. Calcular: PORT = (HILO - BASE) + 1
       - Con HILO=64, si BASE=33, entonces PORT = (64-33)+1 = 32
       - Luego debe existir PORT=32 en esa misma fila o en otra
       
    4. Si todo falla, ejecuta este código de debugging en tu programa principal:
    
    ```python
    # Agregar antes de calcular id_odf_troncal
    print(f"DEBUG OT: fibra={fibra}, odf_calle={odf_calle}, hilo={hilo}")
    print(f"DEBUG ot_base: {ot_base.get((fibra, odf_calle), 'NO ENCONTRADO')}")
    print(f"DEBUG ot_mapping keys: {[k for k in ot_mapping.keys() if k[0]==fibra and k[1]==odf_calle]}")