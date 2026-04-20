import pandas as pd
from pathlib import Path
import time
from collections import defaultdict

BASE_PATH = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality")
ARCHIVO_PRINCIPAL = BASE_PATH / "03-2025-08-31_BASE DE ODF´S.xlsx"
ARCHIVO_CENTRALES = BASE_PATH / "centrales.xlsx"
ARCHIVO_SALIDA = BASE_PATH / "04-Sync_SW_SP.xlsx"
HOJA_SALIDA = "ODFS"

ALIAS_CENTRALES = {
    'AV CHILE': 'AVENIDA CHILE', 'AV. CHILE': 'AVENIDA CHILE',  
    'CHILE': 'AVENIDA CHILE',
    'CIUDAD UNIVERSITARIA': 'CUNI',
    'ESPERANZA': 'VILLAV_ESPERANZA',
    'FACATATIVA': 'FACATATIVÁ',
    'SB': 'SANTA BARBARA',
    'SANTA FE': 'CENTRO',
    'ZIPAQUIRA': 'ZIPAQUIRÁ',
    'PORFIA': 'VILLAV_PORFIA',
    'NORMAND': 'NORMANDIA',
    'POMPILIO': 'POMPILIO PAEZ',
    'TM': 'FONTIBON', 'FO': 'FONTIBON',
    'ZF': 'ZONA FRANCA',
}

COLUMNAS_PRINCIPALES = {
    'Central': str,
    'Abr_central': str,
    'num_troncal_actual': str
}

# Columnas en el orden solicitado, sin COMMENTS
COLUMNAS_SALIDA = [
    'VALIDACION',
    'TRONCAL', 
    'C_ORIGINAL', 
    'C_RESULTADO_ID',
    'C_RESULTADO_NOMBRE',
    'ABREVIATURA', 
    'CANTIDAD DE HILOS',
    'Distribucion_centrales', 
    'Ids_distribucion'
]


class CentralProcessor:
    def __init__(self, archivo_centrales=None):
        self.centrales_map = {}
        self.centrales_id_map = {}
        self.abreviatura_id_map = {}
        
        if archivo_centrales and Path(archivo_centrales).exists():
            self.cargar_mapeo_centrales(archivo_centrales)
        else:
            self._crear_mapeo_completo()
    
    def cargar_mapeo_centrales(self, archivo_path):
        try:
            df_centrales = pd.read_excel(
                archivo_path, 
                dtype={'Id Texto': str, 'Name': str, 'Diminutivo': str}
            )
            
            for _, row in df_centrales.iterrows():
                id_texto = str(row['Id Texto']).strip()
                nombre = str(row['Name']).strip().upper()
                diminutivo = str(row['Diminutivo']).strip()
                
                if nombre and pd.notna(nombre) and nombre != 'NAN':
                    if not diminutivo or diminutivo == 'AUN NO SE DEFINE':
                        diminutivo = nombre
                    
                    self.centrales_map[nombre] = diminutivo
                    self.centrales_id_map[nombre] = id_texto
                    self.abreviatura_id_map[diminutivo] = id_texto
            
            for alias, nombre_real in ALIAS_CENTRALES.items():
                alias_upper = alias.upper()
                nombre_real_upper = nombre_real.upper()
                if nombre_real_upper in self.centrales_map:
                    self.centrales_map[alias_upper] = self.centrales_map[nombre_real_upper]
                    self.centrales_id_map[alias_upper] = self.centrales_id_map.get(nombre_real_upper, '')
            
        except Exception:
            self._crear_mapeo_completo()
    
    def _crear_mapeo_completo(self):
        mapeo_completo = {
            'ACACIAS': ('AUN NO SE DEFINE', '3492000048794233723'),
            'AEROPUERTO': ('AE', '612802304624240477'),
            'AUTOPISTA': ('AU', '612802304624240362'),
            'AVENIDA CHILE': ('AC', '612802304624240368'),
            'BACHUE': ('BA', '612802304593062437'),
            'BOCHICA': ('BC', '612802304624240374'),
            'BOSA': ('BS', '612802304624240483'),
            'CALANDAIMA': ('AUN NO SE DEFINE', '3364551064720075568'),
            'CANDELARIA': ('AUN NO SE DEFINE', '612802304624240489'),
            'CASTILLA LA NUEVA': ('AUN NO SE DEFINE', '3364551064719059922'),
            'CENTRO': ('CE', '612802304624240386'),
            'CHAPINERO': ('CH', '612802304624240392'),
            'CHIA': ('CHIA', '612802304620015698'),
            'CHICO': ('CO', '612802304561225334'),
            'CRUCES': ('CR', '612802304624240495'),
            'CUCUTA-NODO CTI': ('AUN NO SE DEFINE', '3364551064707426861'),
            'CUNI': ('CU', '612802304584530822'),
            'FACATATIVÁ': ('FA', '3364551064701882002'),
            'FONTIBON': ('FO', '612802304593062415'),
            'G_RICAURTE': ('AUN NO SE DEFINE', '1974723527736222316'),
            'GUAYMARAL': ('GU', '612802304614727499'),
            'HOLANDA': ('HO', '612802304598409030'),
            'JUAN REY': ('JR', '612802304624240398'),
            'KENNEDY': ('KE', '612802304624240501'),
            'LA CALERA': ('LA CALERA', '3492000048816529166'),
            'LLANO GAS': ('LLANO GAS', '3492000048801818968'),
            'LUCERO': ('AUN NO SE DEFINE', '612802304614727492'),
            'MADRID': ('MA', '1974723527731707021'),
            'MARANDU': ('AUN NO SE DEFINE', '3364551064720075638'),
            'MARRUECOS': ('AUN NO SE DEFINE', '612802304625319576'),
            'MOSQUERA': ('MO', '1974723527731707015'),
            'MUZU': ('MU', '612802304614727471'),
            'NIZA': ('NI', '612802304624240410'),
            'NORMANDIA': ('NO', '612802304624240507'),
            'OLAYA': ('OL', '612802304624240441'),
            'POMPILIO PAEZ': ('POMPILIO', '3492000048814726176'),
            'PUENTE ARANDA': ('PA', '612802304624240513'),
            'PUENTE VARGAS': ('PUENTE VARGAS', '3364551064701907190'),
            'RICAURTE': ('RI', '612802304624240471'),
            'SALITRE': ('SA', '612802304624240417'),
            'SAN CARLOS': ('SC', '612802304624240465'),
            'SAN FERNANDO': ('SF', '612802304624240453'),
            'SAN JOSE': ('SJ', '612802304624240447'),
            'SANTA BARBARA': ('SB', '612802304624240423'),
            'SANTA HELENITA': ('SH', '612802304624240435'),
            'SANTA INES': ('SI', '612802304614727478'),
            'SANTA PAZ': ('AUN NO SE DEFINE', '3364551064720075599'),
            'SAUZALITO': ('SZ', '612802304624240429'),
            'SOACHA': ('SO', '612802304625319570'),
            'SOPÓ': ('SP', '3364551064701873482'),
            'SUBA': ('SU', '612802304594145113'),
            'SUBA RINCON': ('AUN NO SE DEFINE', '3364551064725655323'),
            'TEUSAQUILLO': ('TE', '612802304624240404'),
            'TIBABUYES': ('TI', '612802304624240459'),
            'TOBERIN': ('TO', '612802304564998594'),
            'TOCANCIPÁ': ('TC', '3364551064701872955'),
            'VILLAV_BUQUE': ('AUN NO SE DEFINE', '3364551064707426855'),
            'VILLAV_CATAMA': ('AUN NO SE DEFINE', '3364551064707426849'),
            'VILLAV_CENTRO': ('AUN NO SE DEFINE', '3364551064707426822'),
            'VILLAV_ESPERANZA': ('ESPERANZA', '3364551064707426837'),
            'VILLAV_PORFIA': ('PORFIA', '3364551064707426830'),
            'VILLAV_VILLA_ORTIZ': ('AUN NO SE DEFINE', '3364551064707426843'),
            'YOMASA': ('YO', '612802304614727485'),
            'ZIPAQUIRÁ': ('ZP', '3364551064701223934'),
            'ZONA FRANCA': ('ZF', '3364551064706915291'),
        }
        
        for nombre, (abr, id_texto) in mapeo_completo.items():
            nombre_upper = nombre.upper()
            
            if abr == 'AUN NO SE DEFINE':
                abr = self._generar_abreviatura_automatica(nombre)
            
            self.centrales_map[nombre_upper] = abr
            self.centrales_id_map[nombre_upper] = id_texto
            self.abreviatura_id_map[abr] = id_texto
        
        for alias, real in ALIAS_CENTRALES.items():
            real_upper = real.upper()
            if real_upper in mapeo_completo:
                abr, id_texto = mapeo_completo[real_upper]
                
                if abr == 'AUN NO SE DEFINE':
                    abr = self._generar_abreviatura_automatica(real)
                
                self.centrales_map[alias.upper()] = abr
                self.centrales_id_map[alias.upper()] = id_texto
    
    def _generar_abreviatura_automatica(self, nombre):
        if ' ' in nombre or '_' in nombre:
            palabras = nombre.replace('_', ' ').split()
            if len(palabras) >= 2:
                return ''.join(p[0].upper() for p in palabras[:min(3, len(palabras))])
        
        return nombre[:min(3, len(nombre))].upper()
    
    def normalizar_nombre_central(self, nombre):
        if pd.isna(nombre):
            return None
        
        nombre_str = str(nombre).strip().upper()
        
        if nombre_str in ['0', 'NAN', 'NONE', 'NULL', '#N/A', '#N/A (BOGOTA)', '']:
            return None
        
        nombre_str = ''.join(c for c in nombre_str if c.isalnum() or c in [' ', '_', '-', '.', 'Á', 'É', 'Í', 'Ó', 'Ú', 'Ñ'])
        
        if nombre_str in ALIAS_CENTRALES:
            nombre_str = ALIAS_CENTRALES[nombre_str].upper()
        
        return nombre_str
    
    def obtener_abreviatura_e_id(self, nombre):
        if pd.isna(nombre):
            return None, None
        
        nombre_norm = self.normalizar_nombre_central(nombre)
        
        if not nombre_norm:
            return None, None
        
        if nombre_norm in self.centrales_map:
            abr = self.centrales_map[nombre_norm]
            id_central = self.centrales_id_map.get(nombre_norm, '')
            return abr, id_central
        
        nombre_simple = self._simplificar_nombre(nombre_norm)
        
        for clave in self.centrales_map:
            clave_simple = self._simplificar_nombre(clave)
            
            if (clave_simple in nombre_simple or nombre_simple in clave_simple) and len(nombre_simple) > 2:
                abr = self.centrales_map[clave]
                id_central = self.centrales_id_map.get(clave, '')
                return abr, id_central
        
        return None, None
    
    def _simplificar_nombre(self, nombre):
        if not nombre:
            return ''
        
        nombre = str(nombre).upper().strip()
        
        reemplazos = {
            'Á': 'A', 'É': 'E', 'Í': 'I', 'Ó': 'O', 'Ú': 'U', 'Ñ': 'N',
            '.': '', ',': '', '-': ' ', '_': ' '
        }
        
        for orig, repl in reemplazos.items():
            nombre = nombre.replace(orig, repl)
        
        while '  ' in nombre:
            nombre = nombre.replace('  ', ' ')
        
        return nombre
    
    def es_central_reconocida(self, abreviatura):
        if not abreviatura:
            return False
        
        return abreviatura in self.abreviatura_id_map
    
    def obtener_id_por_abreviatura(self, abreviatura):
        """Obtiene el ID de una central por su abreviatura"""
        if not abreviatura:
            return ''
        return self.abreviatura_id_map.get(abreviatura, '')
    
    def procesar_archivo(self, archivo_path):
        start_time = time.time()
        
        try:
            df = pd.read_excel(
                archivo_path, 
                usecols=COLUMNAS_PRINCIPALES.keys(),
                dtype=COLUMNAS_PRINCIPALES
            )
            
            resultados, centrales_por_cable = self._procesar_primera_pasada(df)
            self._procesar_segunda_pasada(resultados, centrales_por_cable)
            
            # Crear DataFrame con las columnas en el orden solicitado
            df_resultado = pd.DataFrame({
                'VALIDACION': [r['VALIDACION'] for r in resultados],
                'TRONCAL': [r['TRONCAL'] for r in resultados],
                'C_ORIGINAL': [r['C_ORIGINAL'] for r in resultados],
                'C_RESULTADO_ID': [r['C_RESULTADO_ID'] for r in resultados],
                'C_RESULTADO_NOMBRE': [r['C_RESULTADO_NOMBRE'] for r in resultados],
                'ABREVIATURA': [r['ABREVIATURA'] for r in resultados],
                'CANTIDAD DE HILOS': ['' for _ in resultados],  # Columna vacía como solicitaste
                'Distribucion_centrales': [r['Distribucion_centrales'] for r in resultados],
                'Ids_distribucion': [r['Ids_distribucion'] for r in resultados]
            })
            
            return df_resultado
            
        except Exception as e:
            raise
    
    def _procesar_primera_pasada(self, df):
        resultados = []
        centrales_por_cable = defaultdict(set)
        
        for idx, row in df.iterrows():
            troncal = row['num_troncal_actual'] if pd.notna(row['num_troncal_actual']) else ''
            central_original = row['Central'] if pd.notna(row['Central']) else ''
            abr_original = row['Abr_central'] if pd.notna(row['Abr_central']) else ''
            
            abreviatura, id_central = self.obtener_abreviatura_e_id(central_original)
            
            if abreviatura:
                validacion = 'OK'
                abr_final = abreviatura
                id_final = id_central
                if self.es_central_reconocida(abr_final):
                    centrales_por_cable[troncal].add(abr_final)
            else:
                validacion = 'NO_RECONOCIDA'
                abr_final = self._obtener_abreviatura_valida(abr_original)
                id_final = ''
            
            resultados.append({
                'TRONCAL': troncal,
                'C_ORIGINAL': central_original,
                'C_RESULTADO_NOMBRE': self.normalizar_nombre_central(central_original) or '',
                'C_RESULTADO_ID': id_final,
                'ABREVIATURA': abr_final,
                'VALIDACION': validacion,
                'Distribucion_centrales': '',  # Se llenará en segunda pasada
                'Ids_distribucion': ''  # Se llenará en segunda pasada
            })
        
        return resultados, centrales_por_cable
    
    def _obtener_abreviatura_valida(self, abr_original):
        if not abr_original or pd.isna(abr_original):
            return ''
        
        abr_str = str(abr_original).strip()
        if abr_str in ['0', '#N/A', 'NAN', '#N/A (BOGOTA)', '']:
            return ''
        
        return abr_str
    
    def _procesar_segunda_pasada(self, resultados, centrales_por_cable):
        for resultado in resultados:
            troncal = resultado['TRONCAL']
            if troncal in centrales_por_cable:
                # Obtener las abreviaturas ordenadas
                centrales_ordenadas = sorted(centrales_por_cable[troncal])
                
                # Formatear distribución de centrales
                centrales_str = ', '.join(centrales_ordenadas)
                resultado['Distribucion_centrales'] = centrales_str
                
                # Formatear IDs con comillas como solicitaste
                ids_con_comillas = []
                for abr in centrales_ordenadas:
                    id_central = self.obtener_id_por_abreviatura(abr)
                    if id_central:
                        ids_con_comillas.append(f'"{id_central}"')
                
                resultado['Ids_distribucion'] = '|'.join(ids_con_comillas)
            else:
                resultado['Distribucion_centrales'] = ''
                resultado['Ids_distribucion'] = ''
    
    def guardar_resultado(self, df_resultado, archivo_salida=ARCHIVO_SALIDA, hoja=HOJA_SALIDA):
        """Guarda el resultado en un archivo Excel específico, en una hoja específica"""
        archivo_salida = Path(archivo_salida)
        
        # Verificar si el archivo ya existe
        if archivo_salida.exists():
            # Cargar el archivo existente con todas las hojas
            with pd.ExcelWriter(archivo_salida, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                # Guardar la hoja ODFS con los resultados
                df_resultado.to_excel(writer, sheet_name=hoja, index=False)
        else:
            # Si el archivo no existe, crear uno nuevo
            df_resultado.to_excel(archivo_salida, sheet_name=hoja, index=False)
        
        return str(archivo_salida)


def main():
    archivo_principal = Path(ARCHIVO_PRINCIPAL)
    if not archivo_principal.exists():
        print(f"Archivo principal no encontrado: {archivo_principal}")
        return
    
    print("Procesando archivo...")
    processor = CentralProcessor(ARCHIVO_CENTRALES)
    
    resultado = processor.procesar_archivo(archivo_principal)
    
    # Guardar en la ruta específica y hoja específica
    output_path = processor.guardar_resultado(resultado)

    print(f"MG -ODFS")


if __name__ == "__main__":
    main()