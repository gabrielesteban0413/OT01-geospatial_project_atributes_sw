import sys
from pathlib import Path
from concurrent.futures import ProcessPoolExecutor
import time
import psycopg2
import pandas as pd
import tempfile

sys.path.insert(0, str(Path(__file__).parent.parent.parent))

from database.credentials import DatabaseCredentials

def load_single_file(excel_path, table_name, connection_string):
    try:
        conn = psycopg2.connect(connection_string)
        conn.autocommit = False
        cursor = conn.cursor()
        
        df = pd.read_excel(excel_path)
        df.columns = [col.lower().replace(' ', '_').replace('-', '_') for col in df.columns]
        df = df.where(pd.notnull(df), None)
        
        with tempfile.NamedTemporaryFile(mode='w', suffix='.csv', delete=False) as f:
            df.to_csv(f, index=False, header=True)
            csv_path = f.name
        
        cursor.execute(f"DROP TABLE IF EXISTS {table_name} CASCADE")
        
        columns = df.columns.tolist()
        columns_def = ', '.join([f'"{col}" TEXT' for col in columns])
        cursor.execute(f"CREATE TABLE {table_name} ({columns_def})")
        
        with open(csv_path, 'r') as f:
            cursor.copy_expert(f"COPY {table_name} FROM STDIN WITH CSV HEADER", f)
        
        conn.commit()
        
        cursor.execute(f"SELECT COUNT(*) FROM {table_name}")
        row_count = cursor.fetchone()[0]
        
        cursor.close()
        conn.close()
        Path(csv_path).unlink()
        
        return table_name, row_count, None
        
    except Exception as e:
        return table_name, 0, str(e)

def load_file_wrapper(args):
    excel_path, table_name, connection_string = args
    return load_single_file(excel_path, table_name, connection_string)

if __name__ == '__main__':
    FILES_TO_LOAD = [
        ('database/output/01-Bulk export of buildings.xlsx', 'buildings'),
        ('database/output/01-Bulk export of fibers.xlsx', 'fibers'),
        ('database/output/01-Bulk export of ports.xlsx', 'ports'),
        ('database/output/01-Bulk export of splices.xlsx', 'splices'),
    ]
    
    creds = DatabaseCredentials()
    connection_string = creds.get_connection_string()
    
    args_list = [(path, name, connection_string) for path, name in FILES_TO_LOAD]
    
    start_time = time.time()
    
    with ProcessPoolExecutor(max_workers=4) as executor:
        results = executor.map(load_file_wrapper, args_list)
        
        for table_name, rows, error in results:
            if error:
                print(f"ERROR: {table_name} - {error}")
            else:
                print(f"{table_name}: {rows} filas")
    
    elapsed = time.time() - start_time
    print(f"Tiempo total: {elapsed:.2f} segundos")