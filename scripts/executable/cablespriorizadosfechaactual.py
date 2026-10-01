import pandas as pd
import os
import sys
from datetime import datetime
import gc

RUTA_ORIGEN = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-client_report_origin.xlsx"
RUTA_DESTINO = r"C:\A_GS1_PROYECTOS\0_Documents_gs\database\output\03-client_report_filtrado.xlsx"
FECHA_INICIO = "24/08/2026"
FECHA_FIN = "31/08/2026"

class FiltradorMasivo:
    def __init__(self):
        self.df_resultado = None
        self.fecha_ini = datetime.strptime(FECHA_INICIO, "%d/%m/%Y")
        self.fecha_fin = datetime.strptime(FECHA_FIN, "%d/%m/%Y")
    
    def _parse_fecha(self, valor):
        if pd.isna(valor):
            return None
        if isinstance(valor, (datetime, pd.Timestamp)):
            return pd.Timestamp(valor)
        try:
            valor_str = str(valor).strip()
            for fmt in ['%Y-%m-%d %H:%M:%S', '%Y-%m-%d', '%d/%m/%Y', '%d/%m/%y', '%d-%m-%Y', '%m/%d/%Y']:
                try:
                    return pd.Timestamp(datetime.strptime(valor_str, fmt))
                except:
                    continue
            return pd.Timestamp(valor_str)
        except:
            return None
    
    def _encontrar_hoja(self):
        hojas = pd.ExcelFile(RUTA_ORIGEN).sheet_names
        for hoja in hojas:
            if 'cliente' in hoja.lower() or 'priorizado' in hoja.lower() or 'conectividad' in hoja.lower():
                return hoja
            if 'client' in hoja.lower() or 'report' in hoja.lower():
                return hoja
        return hojas[0]
    
    def _encontrar_columna_fecha(self, df):
        for col in df.columns:
            col_lower = col.lower().strip()
            if col_lower in ['fecha', 'date', 'fechacreacion', 'fecha_creacion']:
                return col
            if 'fecha' in col_lower or 'date' in col_lower:
                return col
        return df.columns[0]
    
    def ejecutar(self):
        if not os.path.exists(RUTA_ORIGEN):
            raise FileNotFoundError(f"Archivo no encontrado: {RUTA_ORIGEN}")
        
        hoja = self._encontrar_hoja()
        df_sample = pd.read_excel(RUTA_ORIGEN, nrows=5, sheet_name=hoja, dtype=str)
        col_fecha = self._encontrar_columna_fecha(df_sample)
        
        df_completo = pd.read_excel(RUTA_ORIGEN, sheet_name=hoja, dtype=str)
        df_completo['_fecha_temp'] = df_completo[col_fecha].apply(self._parse_fecha)
        
        mask = (
            (df_completo['_fecha_temp'] >= pd.Timestamp(self.fecha_ini)) & 
            (df_completo['_fecha_temp'] <= pd.Timestamp(self.fecha_fin))
        )
        
        self.df_resultado = df_completo[mask].copy()
        if '_fecha_temp' in self.df_resultado.columns:
            self.df_resultado = self.df_resultado.drop('_fecha_temp', axis=1)
        
        del df_completo
        gc.collect()
        
        if self.df_resultado is not None and len(self.df_resultado) > 0:
            os.makedirs(os.path.dirname(RUTA_DESTINO), exist_ok=True)
            with pd.ExcelWriter(RUTA_DESTINO, engine='openpyxl') as writer:
                self.df_resultado.to_excel(writer, sheet_name='Filtrado', index=False)
            return True
        else:
            return False

if __name__ == "__main__":
    try:
        filtrador = FiltradorMasivo()
        exito = filtrador.ejecutar()
        sys.exit(0 if exito else 1)
    except Exception as e:
        print(f"Error: {e}")
        sys.exit(1)