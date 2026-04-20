from oletools.olevba import VBA_Parser

# Ruta al archivo .bin extraído de tu Excel
archivo_vba = 'ruta/a/vbaProject.bin'  # <-- ajusta la ruta

# Analiza el archivo
vba_parser = VBA_Parser(archivo_vba)

if vba_parser.detect_vba_macros():
    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
        print(f"\n--- Módulo: {vba_filename} ---")
        print(vba_code)
else:
    print("No se detectaron macros.")
