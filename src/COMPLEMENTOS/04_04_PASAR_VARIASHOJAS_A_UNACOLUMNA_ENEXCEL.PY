import pandas as pd
import numpy as np

# Leer el archivo que generaste
ruta_archivo = r"C:\Users\gabrboa1\Downloads\VIAS_CENTRALES_UNIFICADO.xlsx"
df = pd.read_excel(ruta_archivo)

print("Columnas originales:", list(df.columns))

# 1. Eliminar columnas completamente vacías o con solo NaN
df = df.dropna(axis=1, how='all')

# 2. Eliminar columnas basura (Unnamed, @, columnas numéricas sin nombre)
columnas_a_eliminar = []
for col in df.columns:
    col_lower = str(col).lower()
    if 'unnamed' in col_lower:
        columnas_a_eliminar.append(col)
    elif col.startswith('@'):
        columnas_a_eliminar.append(col)
    elif col.startswith('Unnamed'):
        columnas_a_eliminar.append(col)
    elif col in ['HILO.1', 'INTERFAZ-PUERTO-', 'INTERFAZ/PUERTO/GE']:
        # Estas son duplicadas de otras columnas
        columnas_a_eliminar.append(col)

df = df.drop(columns=columnas_a_eliminar, errors='ignore')

print(f"\nColumnas eliminadas: {columnas_a_eliminar}")

# 3. Unificar columnas equivalentes (combinar datos)
# Mapeo de columnas destino -> columnas origen que deben combinarse
unificar_columnas = {
    'PATCHEO': ['PATCHEO', ' PACHEO', 'PACHEO', 'ODF PACHEO'],
    'INTERFAZ-PUERTO-GE': ['INTERFAZ-PUERTO-GE', 'INTERFAZ-PUERTO', 'INTERFAZ-PUERTO-', 'INTERFAZ/PUERTO/GE'],
    'HILO': ['HILO', 'HILO.1'],
    'OBSERVACION': ['OBSERVACION', 'BRAZO - CLIENTE']
}

for col_destino, col_origenes in unificar_columnas.items():
    for col_origen in col_origenes:
        if col_origen in df.columns and col_origen != col_destino:
            # Combinar datos: si la columna destino está vacía, tomar el valor de la columna origen
            mascara_vacia = df[col_destino].isna() | (df[col_destino].astype(str).str.strip() == '')
            df.loc[mascara_vacia, col_destino] = df.loc[mascara_vacia, col_origen]
            # Eliminar la columna origen
            df = df.drop(columns=[col_origen], errors='ignore')

# 4. Reordenar columnas dejando 'centrales' primero
columnas_principales = ['centrales', 'SALON', 'FILA', 'BASTIDOR', 'EQUIPO', 
                        'INTERFAZ-PUERTO-GE', 'ANILLO', 'PATCHEO', 
                        'TRONCAL-CABLE', 'ODF CALLE', 'HILO', 'count', 'OBSERVACION']

# Mantener solo las columnas que existen
columnas_finales = [col for col in columnas_principales if col in df.columns]
# Agregar cualquier otra columna que no esté en la lista principal al final
otras_columnas = [col for col in df.columns if col not in columnas_finales]
columnas_finales = columnas_finales + otras_columnas

df = df[columnas_finales]

# 5. Guardar versión limpia
ruta_limpia = r"C:\Users\gabrboa1\Downloads\VIAS_CENTRALES_LIMPIO.xlsx"
df.to_excel(ruta_limpia, index=False)

print("\n" + "="*60)
print("✅ LIMPIEZA COMPLETADA!")
print(f"✓ Columnas finales ({len(df.columns)}): {list(df.columns)}")
print(f"✓ Total filas: {len(df)}")
print(f"✓ Archivo guardado: {ruta_limpia}")

# Mostrar primeras filas
print("\nVista previa de datos limpios:")
print(df.head(5).to_string())