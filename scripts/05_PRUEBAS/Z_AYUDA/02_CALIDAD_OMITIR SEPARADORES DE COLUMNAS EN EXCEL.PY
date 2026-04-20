import re
import csv
import os


archivo_entrada = r'C:\A_GS1_PROYECTOS\FI.txt'
salida_csv = r'C:\A_GS1_PROYECTOS\datos_corregidos.csv'


with open(archivo_entrada, 'r', encoding='utf-8') as f:
    texto_crudo = f.read()


todo_en_una = texto_crudo.replace('\n', ' ').strip()
registros = re.findall(r'(349\d{16}.*?)((?=349\d{16})|$)', todo_en_una)
with open(salida_csv, mode='w', newline='', encoding='utf-8') as f:
    writer = csv.writer(f, delimiter='\t')
    for registro, _ in registros:
        columnas = registro.strip().split('\t')
        writer.writerow([c.strip() for c in columnas if c.strip() != ''])

print(f"OK")
