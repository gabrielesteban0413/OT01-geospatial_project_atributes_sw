import json
from pathlib import Path
from typing import Dict

class DatabaseCredentials:
    def __init__(self):
        self.credentials_path = Path('config/security/credentials/postgres.json')
    
    def get_connection_string(self) -> str:
        with open(self.credentials_path) as f:
            creds = json.load(f)
        
        return f"postgresql://{creds['user']}:{creds['password']}@{creds['host']}:{creds['port']}/{creds['database']}"