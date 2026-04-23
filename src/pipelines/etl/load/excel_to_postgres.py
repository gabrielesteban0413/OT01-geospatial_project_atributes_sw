import pandas as pd
from sqlalchemy import create_engine
import psycopg2
from psycopg2.extras import execute_values
from pathlib import Path
import glob

class ExcelToPostgresLoader:
    def __init__(self, connection_string: str):
        self.connection_string = connection_string
    
    def _sanitize_column_name(self, name: str) -> str:
        name = name.replace(' ', '_').replace('-', '_')
        name = ''.join(c for c in name if c.isalnum() or c == '_')
        return name.lower()
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.where(pd.notnull(df), None)
        df.columns = [self._sanitize_column_name(col) for col in df.columns]
        df = df.dropna(how='all')
        return df
    
    def load_excel(self, excel_path: str, table_name: str, batch_size: int = 10000, if_exists: str = 'replace'):
        df = pd.read_excel(excel_path)
        df = self._clean_dataframe(df)
        
        engine = create_engine(self.connection_string)
        df.to_sql(table_name, engine, if_exists=if_exists, index=False, chunksize=batch_size)
        return len(df)
    
    def load_excel_with_copy(self, excel_path: str, table_name: str, batch_size: int = 10000, if_exists: str = 'replace'):
        df = pd.read_excel(excel_path)
        df = self._clean_dataframe(df)
        
        conn = psycopg2.connect(self.connection_string)
        cursor = conn.cursor()
        
        if if_exists == 'replace':
            cursor.execute(f"DROP TABLE IF EXISTS {table_name} CASCADE")
            conn.commit()
        
        columns = df.columns.tolist()
        columns_str = ', '.join([f'"{col}"' for col in columns])
        
        create_sql = f"""
        CREATE TABLE IF NOT EXISTS {table_name} (
            {', '.join([f'"{col}" TEXT' for col in columns])}
        )
        """
        cursor.execute(create_sql)
        conn.commit()
        
        data_tuples = [tuple(x) for x in df.to_numpy()]
        insert_sql = f"INSERT INTO {table_name} ({columns_str}) VALUES %s"
        
        for i in range(0, len(data_tuples), batch_size):
            batch = data_tuples[i:i+batch_size]
            execute_values(cursor, insert_sql, batch)
            conn.commit()
        
        cursor.close()
        conn.close()
        return len(df)
    
    def load_multiple(self, pattern: str, table_prefix: str = '', **kwargs):
        files = glob.glob(pattern)
        results = {}
        
        for file_path in files:
            base_name = Path(file_path).stem
            table_name = f"{table_prefix}{self._sanitize_column_name(base_name)}"
            rows = self.load_excel(file_path, table_name, **kwargs)
            results[table_name] = rows
        
        return results


def fast_excel_to_postgres(excel_path: str, postgres_url: str, table_name: str = None, drop_table: bool = False):
    if table_name is None:
        table_name = Path(excel_path).stem.replace(' ', '_').lower()
    
    loader = ExcelToPostgresLoader(postgres_url)
    if_exists = 'replace' if drop_table else 'append'
    rows = loader.load_excel(excel_path, table_name, if_exists=if_exists)
    
    return {'rows_loaded': rows, 'table': table_name}