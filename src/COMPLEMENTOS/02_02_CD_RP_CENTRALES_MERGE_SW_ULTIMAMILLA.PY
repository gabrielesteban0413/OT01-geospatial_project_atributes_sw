import pandas as pd
from pathlib import Path
import re

# CONFIGURACION DE RUTAS
BASE_PATH = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality")

ARCHIVO_FIBRAS = BASE_PATH / "01-Bulk export of fibers.xlsx"
ARCHIVO_ODFS = BASE_PATH / "04-Sync_SW_SP.xlsx"
ARCHIVO_SALIDA = BASE_PATH / "04-Sync_SW_SP.xlsx"
HOJA_ASPHIA = "SW_CLT_FI"
HOJA_ODFS = "ODFS"

# MAPEO DE COLUMNAS ESPAÑOL A INGLES
COLUMNAS_MAPEO = {
    'Id Text': 'ID_TEXT',
    'Name': 'NAME', 
    'Nombre Especificación': 'NOMBRE_ESP',
    'Construction Status': 'CONSTRUCTI',
    'Nodo A': 'NODO A',
    'Nodo B': 'NODO B',
    'Municipio A': 'MUNICIPIO A',
    'Municipio B': 'MUNICIPIO B'
}

COLUMNAS_SALIDA = [
    'ID_TEXT', 'NAME', 'NODO A', 'NODO B', 
    'MUNICIPIO A', 'MUNICIPIO B', 'NOMBRE_ESP', 'CONSTRUCTI',
    'TRONCAL', 'C_ORIGINAL', 'Distribucion_centrales', 'Ids_distribucion'
]


class FiberDataProcessor:
    def __init__(self):
        self.odfs_map = {}
        self.fibras_data = None
    
    def cargar_datos_odfs(self):
        try:
            # Usar engine específico openpyxl
            df_odfs = pd.read_excel(ARCHIVO_ODFS, sheet_name=HOJA_ODFS, dtype=str, engine='openpyxl')
            
            for _, row in df_odfs.iterrows():
                troncal = str(row['TRONCAL']).strip() if pd.notna(row['TRONCAL']) else ''
                if troncal and troncal.lower() != 'nan' and troncal != '':
                    self.odfs_map[troncal] = {
                        'C_ORIGINAL': str(row['C_ORIGINAL']) if pd.notna(row['C_ORIGINAL']) else '',
                        'Distribucion_centrales': str(row['Distribucion_centrales']) if pd.notna(row['Distribucion_centrales']) else '',
                        'Ids_distribucion': str(row['Ids_distribucion']) if pd.notna(row['Ids_distribucion']) else ''
                    }
            
            print(f"ODFS datos cargados: {len(self.odfs_map)} troncales")
            
        except Exception as e:
            print(f"Error cargando ODFS: {e}")
            self.odfs_map = {}
    
    def normalizar_columnas(self, df):
        df_normalizado = pd.DataFrame()
        columnas_originales = list(df.columns)
        
        print("Columnas en archivo de fibras:")
        for i, col in enumerate(columnas_originales, 1):
            print(f"  [{i:2d}] {col}")
        
        for col_espanol, col_ingles in COLUMNAS_MAPEO.items():
            col_encontrada = None
            
            for col_original in columnas_originales:
                if col_espanol.lower() in str(col_original).lower():
                    col_encontrada = col_original
                    break
            
            if col_encontrada and col_encontrada in df.columns:
                df_normalizado[col_ingles] = df[col_encontrada]
                print(f"Mapeo: {col_ingles} <- {col_encontrada}")
            else:
                df_normalizado[col_ingles] = ''
                print(f"Columna no encontrada: {col_espanol} (creando {col_ingles} vacia)")
        
        return df_normalizado
    
    def cargar_datos_fibras(self):
        try:
            # Usar engine específico openpyxl
            df_fibras = pd.read_excel(ARCHIVO_FIBRAS, dtype=str, engine='openpyxl')
            self.fibras_data = self.normalizar_columnas(df_fibras)
            
            # Verificar columnas criticas
            if 'NOMBRE_ESP' not in self.fibras_data.columns or 'CONSTRUCTI' not in self.fibras_data.columns:
                print("ERROR: Columnas criticas no encontradas")
                return False
            
            print(f"Fibras datos cargados: {len(self.fibras_data):,} registros")
            
            return True
            
        except Exception as e:
            print(f"Error cargando fibras: {e}")
            return False
    
    def extraer_troncal(self, name):
        if pd.isna(name) or not isinstance(name, str):
            return ''
        
        name_str = str(name).strip()
        match = re.match(r'^(\d+)', name_str)
        
        if match:
            return match.group(1)
        
        if '-' in name_str:
            parts = name_str.split('-')
            if parts and parts[0].isdigit():
                return parts[0]
        
        return ''
    
    def procesar_fibras(self):
        if self.fibras_data is None:
            return None
        
        # Extraer troncal
        self.fibras_data['TRONCAL'] = self.fibras_data['NAME'].apply(self.extraer_troncal)
        
        # Añadir informacion ODFS
        self.fibras_data['C_ORIGINAL'] = self.fibras_data['TRONCAL'].map(
            lambda x: self.odfs_map.get(x, {}).get('C_ORIGINAL', '')
        )
        
        self.fibras_data['Distribucion_centrales'] = self.fibras_data['TRONCAL'].map(
            lambda x: self.odfs_map.get(x, {}).get('Distribucion_centrales', '')
        )
        
        self.fibras_data['Ids_distribucion'] = self.fibras_data['TRONCAL'].map(
            lambda x: self.odfs_map.get(x, {}).get('Ids_distribucion', '')
        )
        
        # Convertir todo a texto
        for col in self.fibras_data.columns:
            self.fibras_data[col] = self.fibras_data[col].astype(str).fillna('')
        
        return self.fibras_data
    
    def preparar_datos_finales(self):
        if self.fibras_data is None:
            return None
        
        for col in COLUMNAS_SALIDA:
            if col not in self.fibras_data.columns:
                self.fibras_data[col] = ''
        
        df_final = self.fibras_data[COLUMNAS_SALIDA].copy()
        return df_final
    
    def guardar_resultados(self, df_final):
        try:
            if ARCHIVO_SALIDA.exists():
                # Intentar leer con openpyxl primero
                try:
                    xls = pd.ExcelFile(ARCHIVO_SALIDA, engine='openpyxl')
                    existing_sheets = xls.sheet_names
                except:
                    # Si falla, crear nuevo archivo
                    existing_sheets = []
                
                all_sheets = {}
                
                for sheet in existing_sheets:
                    if sheet != HOJA_ASPHIA:
                        try:
                            df_sheet = pd.read_excel(ARCHIVO_SALIDA, sheet_name=sheet, dtype=str, engine='openpyxl')
                            all_sheets[sheet] = df_sheet
                        except:
                            all_sheets[sheet] = pd.DataFrame()
                
                all_sheets[HOJA_ASPHIA] = df_final
                
                # Guardar usando openpyxl
                with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
                    for sheet_name, sheet_data in all_sheets.items():
                        sheet_data.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print(f"Archivo actualizado: {ARCHIVO_SALIDA}")
                return True
                
            else:
                # Crear nuevo archivo con openpyxl
                with pd.ExcelWriter(ARCHIVO_SALIDA, engine='openpyxl') as writer:
                    df_final.to_excel(writer, sheet_name=HOJA_ASPHIA, index=False)
                
                print(f"Archivo creado: {ARCHIVO_SALIDA}")
                return True
            
        except Exception as e:
            print(f"Error guardando resultados: {e}")
            
            # Intento alternativo: usar xlsxwriter
            try:
                with pd.ExcelWriter(ARCHIVO_SALIDA, engine='xlsxwriter') as writer:
                    df_final.to_excel(writer, sheet_name=HOJA_ASPHIA, index=False)
                print(f"Archivo creado con xlsxwriter: {ARCHIVO_SALIDA}")
                return True
            except Exception as e2:
                print(f"Error alternativo: {e2}")
                return False
    
    def ejecutar(self):
        print("=" * 60)
        print("PROCESADOR DE FIBRAS - SW_CLT_FI")
        print("=" * 60)
        
        if not ARCHIVO_FIBRAS.exists():
            print(f"Error: Archivo no encontrado: {ARCHIVO_FIBRAS}")
            return
        
        self.cargar_datos_odfs()
        
        if not self.cargar_datos_fibras():
            return
        
        result_df = self.procesar_fibras()
        
        if result_df is not None:
            df_final = self.preparar_datos_finales()
            
            if df_final is not None:
                if self.guardar_resultados(df_final):
                    troncales = (df_final['TRONCAL'] != '').sum()
                    match_odfs = (df_final['C_ORIGINAL'] != '').sum()
                    
                    print(f"\nRESUMEN:")
                    print(f"Registros procesados: {len(df_final):,}")
                    print(f"Troncales extraidos: {troncales:,}")
                    print(f"Match con ODFS: {match_odfs:,}")
                    print(f"Columna NOMBRE_ESP: {df_final['NOMBRE_ESP'].str.strip().ne('').sum():,} registros con datos")
                    print(f"Columna CONSTRUCTI: {df_final['CONSTRUCTI'].str.strip().ne('').sum():,} registros con datos")
        
        print("=" * 60)
        print("PROCESO COMPLETADO")


if __name__ == "__main__":
    processor = FiberDataProcessor()
    processor.ejecutar()