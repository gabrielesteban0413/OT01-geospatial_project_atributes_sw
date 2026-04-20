# =============================== IMPORTS ===============================
import pandas as pd
import re
from collections import defaultdict, Counter
import os
from datetime import datetime

# =============================== CONFIGURACIÓN DE RUTAS ===============================
BASE_PATH = r"C:\A_GS1_PROYECTOS"

RUTA_ORIGEN = os.path.join(BASE_PATH, "0_Documents_gs", "output", "quality", 
                          "002-----------------------------Línea Base Priorización Cables.xlsx")

RUTA_DESTINO = os.path.join(BASE_PATH, "CABLES_FINAL.xlsx")

# =============================== CONSTANTES ===============================
SEPARADOR_VARIANTES = ' | '

COLUMNAS_SALIDA = [
    'GRUPO', 
    'TOTAL FILAS', 
    'CABLE_PRINCIPAL', 
    'VARIANTES',
    'INCORPORACION_CABLE',
    'RESPONSABLE_RED_EXTERNA',
    'CENTRAL'
]

# =============================== CLASE PRINCIPAL ===============================
class ProcesadorCablesFinal:
    """Clase para procesar todos los cables con las columnas necesarias."""
    
    def __init__(self, ruta_entrada, ruta_salida):
        self.ruta_entrada = ruta_entrada
        self.ruta_salida = ruta_salida
        self.df_datos = None
        self.columna_cables = None
        self.columna_responsable = None
        self.columna_central = None
        self.columna_incorporacion = None
        self.grupos = defaultdict(list)
        self.resultados = []
        
    # =============================== MÉTODOS DE CARGA ===============================
    
    def cargar_datos(self):
        """Carga los datos desde el archivo Excel de entrada."""
        print("Cargando datos...")
        try:
            self.df_datos = pd.read_excel(self.ruta_entrada, sheet_name='TD ACCES0S')
            print(f"Datos cargados: {len(self.df_datos):,} filas")
        except Exception as e:
            print(f"Error al cargar archivo: {str(e)}")
            raise
            
        self._identificar_columnas()
        
    def _identificar_columnas(self):
        """Identifica automáticamente las columnas necesarias."""
        print("\nIdentificando columnas...")
        
        # Mostrar todas las columnas disponibles para referencia
        print(f"Columnas disponibles en el archivo:")
        for i, col in enumerate(self.df_datos.columns, 1):
            print(f"  {i:2}. {col}")
        
        # BUSCAR COLUMNA DE CABLE (Cable acceso)
        for col in self.df_datos.columns:
            col_str = str(col).strip().lower()
            if col_str == 'cable acceso':
                self.columna_cables = col
                break
        
        if not self.columna_cables:
            for col in self.df_datos.columns:
                col_str = str(col).strip().lower()
                if 'cable' in col_str and 'acceso' in col_str:
                    self.columna_cables = col
                    break
        
        print(f"\nColumna de cables identificada: '{self.columna_cables}'")
        
        # BUSCAR COLUMNA DE RESPONSABLE (Responsable red externa) - PRIMERA
        for col in self.df_datos.columns:
            col_str = str(col).strip().lower()
            if col_str == 'responsable red externa':
                self.columna_responsable = col
                break
        
        if not self.columna_responsable:
            for col in self.df_datos.columns:
                col_str = str(col).strip().lower()
                if 'responsable' in col_str and 'red' in col_str:
                    self.columna_responsable = col
                    break
        
        print(f"Columna de responsable identificada: '{self.columna_responsable}'")
        
        # BUSCAR COLUMNA DE CENTRAL
        for col in self.df_datos.columns:
            col_str = str(col).strip().lower()
            if col_str == 'central':
                self.columna_central = col
                break
        
        print(f"Columna de central identificada: '{self.columna_central}'")
        
        # BUSCAR COLUMNA DE INCORPORACIÓN CABLE - PRIMERA OCURRENCIA
        for col in self.df_datos.columns:
            col_str = str(col).strip().lower()
            if col_str == 'incorporación cable':
                self.columna_incorporacion = col
                break
        
        if not self.columna_incorporacion:
            for col in self.df_datos.columns:
                col_str = str(col).strip().lower()
                if 'incorporación' in col_str or 'incorporacion' in col_str:
                    self.columna_incorporacion = col
                    break
        
        print(f"Columna de incorporacion cable identificada: '{self.columna_incorporacion}'")
        
        # Verificar que todas las columnas existen
        columnas_necesarias = [self.columna_cables, self.columna_responsable, 
                              self.columna_central, self.columna_incorporacion]
        
        for col in columnas_necesarias:
            if col not in self.df_datos.columns:
                print(f"\nADVERTENCIA: Columna '{col}' no encontrada en el archivo")
        
        # Mostrar vista previa con valores reales
        print(f"\nVista previa de datos con columnas identificadas:")
        print("-" * 100)
        
        # Mostrar cabecera
        headers = ["Cable", "Responsable", "Central", "Incorporación"]
        print(f"{'Cable':<30} | {'Responsable':<30} | {'Central':<15} | {'Incorporación':<15}")
        print("-" * 100)
        
        # Mostrar primeras 5 filas
        for idx in range(min(5, len(self.df_datos))):
            fila = self.df_datos.iloc[idx]
            cable = str(fila[self.columna_cables])[:30] if self.columna_cables in self.df_datos.columns else "N/A"
            responsable = str(fila[self.columna_responsable])[:30] if self.columna_responsable in self.df_datos.columns else "N/A"
            central = str(fila[self.columna_central])[:15] if self.columna_central in self.df_datos.columns else "N/A"
            incorporacion = str(fila[self.columna_incorporacion])[:15] if self.columna_incorporacion in self.df_datos.columns else "N/A"
            
            print(f"{cable:<30} | {responsable:<30} | {central:<15} | {incorporacion:<15}")
    
    # =============================== MÉTODOS DE PROCESAMIENTO ===============================
    
    def procesar(self):
        """Ejecuta todo el proceso de procesamiento."""
        self._agrupar_cables()
        self._procesar_grupos()
        self._generar_resultados()
        self._guardar_resultados()
        self._mostrar_resumen()
    
    def _agrupar_cables(self):
        """Agrupa todos los cables por su clave normalizada."""
        print("\nAgrupando cables...")
        
        total_filas = len(self.df_datos)
        for indice, fila in self.df_datos.iterrows():
            # Mostrar progreso cada 10,000 filas
            if indice % 10000 == 0 and indice > 0:
                print(f"Procesadas {indice:,}/{total_filas:,} filas...")
            
            # Obtener valor del cable
            cable = ""
            if self.columna_cables in fila:
                cable = fila[self.columna_cables]
            
            if pd.isna(cable) or not str(cable).strip():
                continue
            
            clave = self._normalizar_cable(str(cable).strip())
            if not clave:
                continue
            
            # Obtener valores de las otras columnas
            responsable = ""
            if self.columna_responsable in fila:
                responsable = self._limpiar_valor(fila[self.columna_responsable])
            
            central = ""
            if self.columna_central in fila:
                central = self._limpiar_valor(fila[self.columna_central])
            
            incorporacion = ""
            if self.columna_incorporacion in fila:
                incorporacion = self._limpiar_valor(fila[self.columna_incorporacion])
            
            # Agregar al grupo
            self.grupos[clave].append({
                'cable': str(cable).strip(),
                'responsable': responsable,
                'central': central,
                'incorporacion': incorporacion
            })
        
        print(f"Total grupos identificados: {len(self.grupos):,}")
    
    def _procesar_grupos(self):
        """Procesa todos los grupos."""
        print("\nProcesando grupos...")
        
        total_grupos = len(self.grupos)
        for i, (clave_grupo, filas) in enumerate(self.grupos.items(), 1):
            # Mostrar progreso cada 1000 grupos
            if i % 1000 == 0:
                print(f"Procesados {i:,}/{total_grupos:,} grupos...")
            
            # Obtener variantes únicas de cables
            variantes_set = {f['cable'] for f in filas}
            
            # Filtrar variantes que contienen "P"
            variantes_filtradas = [v for v in variantes_set if 'P' not in v.upper()]
            
            # Si todas las variantes tienen "P", usar las originales
            if not variantes_filtradas:
                variantes_filtradas = list(variantes_set)
            
            # Ordenar variantes: principal primero, luego las demás
            variantes_ordenadas = self._ordenar_variantes(variantes_filtradas)
            cable_principal = variantes_ordenadas[0] if variantes_ordenadas else clave_grupo
            
            # Formatear variantes (columna VARIANTES)
            variantes_str = SEPARADOR_VARIANTES.join(variantes_ordenadas)
            
            # Obtener valores más comunes para cada campo
            responsable_comun = self._obtener_valor_comun([f['responsable'] for f in filas])
            central_comun = self._obtener_valor_comun([f['central'] for f in filas])
            incorporacion_comun = self._obtener_valor_comun([f['incorporacion'] for f in filas])
            
            # Si no hay valor común, usar el de la primera fila
            if not responsable_comun and filas:
                responsable_comun = filas[0]['responsable']
            
            if not central_comun and filas:
                central_comun = filas[0]['central']
            
            if not incorporacion_comun and filas:
                incorporacion_comun = filas[0]['incorporacion']
            
            # Verificar que incorporacion_comun no sea un valor de cable
            # Si parece ser un valor de cable (contiene números y guiones), buscar otro valor
            if incorporacion_comun and self._es_valor_de_cable(incorporacion_comun):
                # Buscar otro valor que no sea cable
                otros_valores = [f['incorporacion'] for f in filas 
                               if f['incorporacion'] and not self._es_valor_de_cable(f['incorporacion'])]
                if otros_valores:
                    incorporacion_comun = self._obtener_valor_comun(otros_valores)
                else:
                    # Si todos son valores de cable, usar vacío
                    incorporacion_comun = ""
            
            # Asegurar que incorporacion_comun sea un estado válido (COMPLETADO, PENDIENTE, etc.)
            if incorporacion_comun:
                incorporacion_comun = self._validar_estado_incorporacion(incorporacion_comun)
            
            self.resultados.append({
                'GRUPO': clave_grupo,
                'TOTAL FILAS': len(filas),
                'CABLE_PRINCIPAL': cable_principal,
                'VARIANTES': variantes_str,
                'INCORPORACION_CABLE': incorporacion_comun,
                'RESPONSABLE_RED_EXTERNA': responsable_comun,
                'CENTRAL': central_comun
            })
        
        print(f"Grupos procesados: {len(self.resultados):,}")
    
    # =============================== MÉTODOS AUXILIARES ===============================
    
    @staticmethod
    def _limpiar_valor(valor):
        """Limpia un valor (quita espacios, maneja NaN)."""
        if pd.isna(valor):
            return ""
        valor_str = str(valor).strip()
        # Limpiar valores como #N/A
        if valor_str.upper() in ['#N/A', 'N/A', 'NA', '#VALOR!', '#REF!', '#DIV/0!']:
            return ""
        return valor_str
    
    def _normalizar_cable(self, cable_str):
        """Normaliza un cable para agrupación."""
        if not cable_str or pd.isna(cable_str):
            return ""
        
        # Remover espacios innecesarios
        cable_str = str(cable_str).strip()
        
        # Extraer parte principal (antes del primer / si existe)
        if '/' in cable_str:
            partes = cable_str.split('/')
            principal = partes[0].strip()
        else:
            principal = cable_str
        
        # Normalizar: quitar espacios, estandarizar guiones
        principal = re.sub(r'\s*-\s*', '-', principal)
        principal = principal.upper()
        principal = re.sub(r'\s+', '', principal)  # Quitar todos los espacios
        principal = re.sub(r'-+', '-', principal)  # Normalizar múltiples guiones
        
        # Remover guiones al final
        while principal.endswith('-'):
            principal = principal[:-1]
        
        return principal
    
    def _ordenar_variantes(self, variantes):
        """Ordena variantes: principal primero, luego las demás."""
        if not variantes:
            return []
        
        # Identificar cable principal
        # Preferir: 1) sin diagonales, 2) más corto, 3) alfabético
        sin_diagonales = [v for v in variantes if '/' not in v]
        
        if sin_diagonales:
            # Ordenar por longitud y luego alfabéticamente
            sin_diagonales.sort(key=lambda x: (len(x), x))
            principal = sin_diagonales[0]
        else:
            # Si todos tienen diagonales, usar el más corto
            variantes.sort(key=lambda x: (len(x), x))
            principal = variantes[0]
        
        # Crear lista ordenada
        resultado = [principal]
        
        # Agregar las demás variantes ordenadas
        otras = [v for v in variantes if v != principal]
        otras.sort(key=lambda x: ('/' in x, len(x), x))
        resultado.extend(otras)
        
        return resultado
    
    @staticmethod
    def _obtener_valor_comun(valores):
        """Obtiene el valor más común de una lista."""
        # Filtrar valores vacíos
        valores_validos = [v for v in valores if v and str(v).strip()]
        if not valores_validos:
            return ""
        
        # Contar frecuencia
        contador = Counter(valores_validos)
        
        # Devolver el más común
        return contador.most_common(1)[0][0]
    
    @staticmethod
    def _es_valor_de_cable(valor):
        """Determina si un valor parece ser un código de cable."""
        if not valor:
            return False
        
        valor_str = str(valor).strip()
        
        # Un código de cable típicamente contiene números y guiones
        # pero no palabras como COMPLETADO, PENDIENTE, etc.
        
        # Si contiene palabras de estado, no es un cable
        palabras_estado = ['COMPLETADO', 'PENDIENTE', 'EXCLUIDO', 'EN PROGRESO', 
                          'CANCELADO', 'APROBADO', 'RECHAZADO']
        
        valor_upper = valor_str.upper()
        for palabra in palabras_estado:
            if palabra in valor_upper:
                return False
        
        # Si contiene principalmente números y guiones, probablemente es un cable
        # Buscar patrones de cable como 70953-55, 71212-1, etc.
        patron_cable = r'^\d{5,6}(-\d+)*(\/\d{5,6}(-\d+)*)*$'
        if re.match(patron_cable, valor_str.replace(' ', '')):
            return True
        
        # Si contiene números y guiones pero no palabras de estado
        if any(c.isdigit() for c in valor_str) and ('-' in valor_str or '/' in valor_str):
            return True
        
        return False
    
    @staticmethod
    def _validar_estado_incorporacion(valor):
        """Valida y normaliza el estado de incorporación."""
        if not valor:
            return ""
        
        valor_str = str(valor).strip().upper()
        
        # Lista de estados válidos
        estados_validos = {
            'COMPLETADO': 'COMPLETADO',
            'PENDIENTE': 'PENDIENTE', 
            'EXCLUIDO': 'EXCLUIDO',
            'EN PROCESO': 'EN PROCESO',
            'EN PROGRESO': 'EN PROGRESO',
            'APROBADO': 'APROBADO',
            'RECHAZADO': 'RECHAZADO',
            'CANCELADO': 'CANCELADO',
            'SUSPENDIDO': 'SUSPENDIDO'
        }
        
        # Verificar si el valor contiene algún estado válido
        for estado_key, estado_val in estados_validos.items():
            if estado_key in valor_str:
                return estado_val
        
        # Si no es un estado válido, verificar si parece ser un cable
        if any(c.isdigit() for c in valor_str) and ('-' in valor_str or '/' in valor_str):
            # Parece ser un cable, no un estado
            return ""
        
        return valor_str
    
    def _generar_resultados(self):
        """Genera el DataFrame final."""
        self.df_resultados = pd.DataFrame(self.resultados)
        
        # Ordenar por TOTAL FILAS descendente
        self.df_resultados = self.df_resultados.sort_values('TOTAL FILAS', ascending=False)
        self.df_resultados = self.df_resultados[COLUMNAS_SALIDA]
        
        # Rellenar valores NaN con cadena vacía
        self.df_resultados = self.df_resultados.fillna('')
    
    # =============================== MÉTODOS DE SALIDA ===============================
    
    def _guardar_resultados(self):
        """Guarda los resultados en Excel."""
        print(f"\nGuardando resultados en: {self.ruta_salida}")
        
        with pd.ExcelWriter(self.ruta_salida, engine='openpyxl') as writer:
            self.df_resultados.to_excel(writer, sheet_name='CABLES', index=False)
            
            # Ajustar anchos de columna
            workbook = writer.book
            worksheet = writer.sheets['CABLES']
            
            # Definir anchos personalizados para cada columna
            anchos = {
                'A': 20,  # GRUPO
                'B': 12,  # TOTAL FILAS
                'C': 25,  # CABLE_PRINCIPAL
                'D': 50,  # VARIANTES
                'E': 20,  # INCORPORACION_CABLE
                'F': 30,  # RESPONSABLE_RED_EXTERNA
                'G': 20   # CENTRAL
            }
            
            for col_letter, width in anchos.items():
                worksheet.column_dimensions[col_letter].width = width
            
            # Formatear TOTAL FILAS como número sin decimales
            for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, min_col=2, max_col=2):
                for cell in row:
                    cell.number_format = '0'
        
        print(f"Archivo guardado exitosamente")
        print(f"Total registros: {len(self.df_resultados):,}")
    
    def _mostrar_resumen(self):
        """Muestra un resumen del procesamiento."""
        print("\n" + "="*80)
        print("RESUMEN FINAL DEL PROCESAMIENTO")
        print("="*80)
        
        total_grupos = len(self.resultados)
        grupos_unicos = sum(1 for r in self.resultados if r['TOTAL FILAS'] == 1)
        grupos_variantes = sum(1 for r in self.resultados if r['TOTAL FILAS'] > 1)
        
        print(f"\nESTADISTICAS:")
        print(f"   Total filas originales: {len(self.df_datos):,}")
        print(f"   Total grupos identificados: {total_grupos:,}")
        print(f"   - Grupos unicos (1 fila): {grupos_unicos:,}")
        print(f"   - Grupos con variantes (>1 fila): {grupos_variantes:,}")
        
        if total_grupos > 0:
            porcentaje_unicos = (grupos_unicos / total_grupos) * 100
            porcentaje_variantes = (grupos_variantes / total_grupos) * 100
            
            print(f"\n   Distribucion porcentual:")
            print(f"   - Unicos: {porcentaje_unicos:.1f}%")
            print(f"   - Con variantes: {porcentaje_variantes:.1f}%")
        
        # Mostrar distribución de valores en Incorporación Cable
        if len(self.resultados) > 0:
            valores_incorporacion = [r['INCORPORACION_CABLE'] for r in self.resultados if r['INCORPORACION_CABLE']]
            if valores_incorporacion:
                print(f"\nDISTRIBUCION EN COLUMNA INCORPORACION_CABLE:")
                contador = Counter(valores_incorporacion)
                total_valores = len(valores_incorporacion)
                
                print(f"   Total valores no vacios: {total_valores:,}")
                print(f"   Valores unicos: {len(contador):,}")
                
                print(f"\n   Top 10 estados mas comunes:")
                for i, (valor, cantidad) in enumerate(contador.most_common(10), 1):
                    porcentaje = (cantidad / total_valores) * 100
                    print(f"   {i:2}. {valor:<15} : {cantidad:>6,} ({porcentaje:>5.1f}%)")
        
        # Mostrar algunos ejemplos CORRECTOS
        print(f"\nEJEMPLOS DE RESULTADOS CORRECTOS:")
        print("-" * 80)
        
        # Buscar ejemplos con diferentes estados de incorporación
        estados_ejemplo = {}
        for resultado in self.resultados:
            estado = resultado['INCORPORACION_CABLE']
            if estado and estado not in estados_ejemplo:
                estados_ejemplo[estado] = resultado
                if len(estados_ejemplo) >= 3:
                    break
        
        for i, (estado, ejemplo) in enumerate(estados_ejemplo.items(), 1):
            print(f"\n{i}. Grupo con estado '{estado}':")
            print(f"   - GRUPO: {ejemplo['GRUPO']}")
            print(f"   - TOTAL FILAS: {ejemplo['TOTAL FILAS']}")
            print(f"   - CABLE_PRINCIPAL: {ejemplo['CABLE_PRINCIPAL']}")
            print(f"   - VARIANTES: {ejemplo['VARIANTES'][:50]}..." if len(ejemplo['VARIANTES']) > 50 else f"   - VARIANTES: {ejemplo['VARIANTES']}")
            print(f"   - INCORPORACION_CABLE: {ejemplo['INCORPORACION_CABLE']}")
            print(f"   - RESPONSABLE: {ejemplo['RESPONSABLE_RED_EXTERNA']}")
            print(f"   - CENTRAL: {ejemplo['CENTRAL']}")
        
        print(f"\nCOLUMNAS GENERADAS ({len(COLUMNAS_SALIDA)}):")
        for i, col in enumerate(COLUMNAS_SALIDA, 1):
            print(f"   {i}. {col}")
        
        print("\n" + "="*80)

# =============================== EJECUCIÓN PRINCIPAL ===============================
def main():
    """Función principal de ejecución."""
    print("="*80)
    print("PROCESADOR DE CABLES - VERSION CORREGIDA")
    print("="*80)
    
    if not os.path.exists(RUTA_ORIGEN):
        print(f"ERROR: Archivo no encontrado")
        print(f"  Ruta: {RUTA_ORIGEN}")
        return
    
    try:
        inicio = datetime.now()
        print(f"Inicio: {inicio.strftime('%H:%M:%S')}")
        
        # Crear y ejecutar procesador
        procesador = ProcesadorCablesFinal(RUTA_ORIGEN, RUTA_DESTINO)
        procesador.cargar_datos()
        procesador.procesar()
        
        fin = datetime.now()
        duracion = fin - inicio
        
        print(f"\nPROCESO COMPLETADO")
        print(f"   Duracion: {duracion.total_seconds():.2f} segundos")
        print(f"   Archivo generado: {RUTA_DESTINO}")
        
    except Exception as e:
        print(f"\nERROR DURANTE EL PROCESO:")
        print(f"   {str(e)}")
        
        # Mostrar error detallado para depuración
        import traceback
        print(f"\nDetalle del error:")
        print(traceback.format_exc())

# =============================== ENTRADA DEL PROGRAMA ===============================
if __name__ == "__main__":
    main()