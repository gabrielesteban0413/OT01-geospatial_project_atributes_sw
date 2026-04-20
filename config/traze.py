import os
import re
import pandas as pd
from concurrent.futures import ProcessPoolExecutor, as_completed
from tqdm import tqdm
from collections import defaultdict

# ============================================
# CONFIGURACIÓN
# ============================================
CARPETA_TRAZOS = r"C:\A_GS1_PROYECTOS\TRAZE"
ARCHIVO_SALIDA = r"C:\A_GS1_PROYECTOS\traces_procesados.xlsx"

# ============================================
# PATRONES DE BÚSQUEDA
# ============================================
PATRON_CENTRAL = re.compile(r'Edificio \(ETB, \'([^\']+)\'\)', re.IGNORECASE)
PATRON_TRONCAL = re.compile(
    r'Fibre \(([^:]+):\s*(\d+)\)\s+in\s+Cable\s+de\s+FO\s*\([^,]+, \'([^\']+)\'\)',
    re.IGNORECASE
)
PATRON_ODF = re.compile(
    r'Port \((\d+)\)\s+in\s+Shelf\s+\(([^)]+)\)\s+in\s+Bay\s+\(([^)]+)\)',
    re.IGNORECASE
)
PATRON_EQUIPO = re.compile(
    r'Port \(G-(\d+)\)\s+in\s+Card\s+\(Slot(\d+)\s+[^\)]+\)\s+in\s+Shelf\s+\(([^\)]+)\)',
    re.IGNORECASE
)
PATRON_BANDEJA = re.compile(
    r'Port \((\d+)\)\s+in\s+Card\s+\(BANDEJA\s+([A-Z])\)\s+in\s+Shelf\s+\(([^)]+)\)',
    re.IGNORECASE
)
PATRON_CLIENTE = re.compile(
    r'^\s*Fibre\s*\([^:]+:\s*(\d+)\)\s+in\s+Cable\s+de\s+FO\s*\([^,]+, \'([^\']+)\'\)\s+at\s+Edificio\s*\(ETB,\s*\'([^\']+)\'\)',
    re.IGNORECASE | re.MULTILINE
)

# ============================================
# FUNCIONES AUXILIARES
# ============================================
def obtener_archivos(carpeta):
    archivos = []
    for root, _, files in os.walk(carpeta):
        for file in files:
            archivos.append(os.path.join(root, file))
    return [a for a in archivos if es_valido(a)]

def es_valido(ruta):
    try:
        with open(ruta, 'r', encoding='utf-8', errors='ignore') as f:
            return 'Signal Level Trace' in f.read(4096)
    except:
        return False

def procesar_archivo(ruta):
    try:
        with open(ruta, 'r', encoding='utf-8', errors='ignore') as f:
            contenido = f.read()
    except:
        with open(ruta, 'r', encoding='latin-1', errors='ignore') as f:
            contenido = f.read()

    resultados = []
    nombre = os.path.basename(ruta)

    for traza in re.split(r'Start Node Number \d+', contenido):
        if 'End of Trace' not in traza:
            continue

        central_match = PATRON_CENTRAL.search(traza)
        central = central_match.group(1) if central_match else ''

        troncal_match = PATRON_TRONCAL.search(traza)
        if troncal_match:
            anillo, hilo_troncal, troncal = troncal_match.groups()
        else:
            anillo = hilo_troncal = troncal = ''

        # ODFs
        odfs = list(PATRON_ODF.finditer(traza))
        odf_troncal_salida = ''
        odf_troncal_entrada = ''
        if odfs:
            primero = odfs[0]
            puerto, shelf_content, _ = primero.groups()
            shelf_id = shelf_content.split()[0] if shelf_content else ''
            odf_troncal_salida = f"{puerto}_{shelf_id}"
            ultimo = odfs[-1]
            puerto, shelf_content, _ = ultimo.groups()
            shelf_id = shelf_content.split()[0] if shelf_content else ''
            odf_troncal_entrada = f"{puerto}_{shelf_id}"

        # Equipos
        equipos = list(PATRON_EQUIPO.finditer(traza))
        equipo_salida = ''
        equipo_entrada = ''
        if equipos:
            primero = equipos[0]
            gport, slot, shelf = primero.groups()
            equipo_salida = f"G-{gport}_{slot}_{shelf}"
            ultimo = equipos[-1]
            gport, slot, shelf = ultimo.groups()
            equipo_entrada = f"G-{gport}_{slot}_{shelf}"

        # Bandejas de patcheo
        bandejas = list(PATRON_BANDEJA.finditer(traza))
        odf_patcheo_salida = ''
        odf_patcheo_entrada = ''

        if bandejas:
            # Referencia para salida: primer ODF (si existe), sino primer equipo
            if odfs:
                ref_salida = odfs[0].start()
            elif equipos:
                ref_salida = equipos[0].start()
            else:
                ref_salida = None

            # Referencia para entrada: último ODF (si existe), sino último equipo
            if odfs:
                ref_entrada = odfs[-1].end()
            elif equipos:
                ref_entrada = equipos[-1].end()
            else:
                ref_entrada = None

            def formatear_bandejas_individual(lista_bandejas):
                """Formatea cada bandeja por separado y las une con '|'."""
                if not lista_bandejas:
                    return ''
                elementos = []
                for b in lista_bandejas:
                    puerto, letra, shelf = b.groups()
                    shelf_id = shelf.split()[0] if shelf else ''
                    elementos.append(f"{puerto}_{letra}_{shelf_id}")
                # Unir con '|'
                return '|'.join(elementos)

            if ref_salida is not None:
                bandejas_antes = [b for b in bandejas if b.end() < ref_salida]
                odf_patcheo_salida = formatear_bandejas_individual(bandejas_antes)
            if ref_entrada is not None:
                bandejas_despues = [b for b in bandejas if b.start() > ref_entrada]
                odf_patcheo_entrada = formatear_bandejas_individual(bandejas_despues)

        # Clientes
        clientes = []
        for match in PATRON_CLIENTE.finditer(traza):
            hilo, cable, cliente = match.groups()
            if cliente != central:
                clientes.append((cliente, cable, int(hilo)))

        grupos = defaultdict(set)
        for cliente, cable, hilo in clientes:
            grupos[(cliente, cable)].add(hilo)

        for (cliente, cable), hilos in grupos.items():
            hilo_acceso = ' Y '.join(str(h) for h in sorted(hilos))
            fila = {
                'central': central,
                'troncal': troncal,
                'hilo_troncal': hilo_troncal,
                'odf_troncal_salida': odf_troncal_salida,
                'odf_troncal_entrada': odf_troncal_entrada,
                'odf_patcheo_salida': odf_patcheo_salida,
                'odf_patcheo_entrada': odf_patcheo_entrada,
                'equipo_salida': equipo_salida,
                'equipo_entrada': equipo_entrada,
                'cable_acceso': cable,
                'hilo_acceso': hilo_acceso,
                'anillo': anillo,
                'direccion_cliente': '',
                'id_servicio': '',
                'cliente': cliente,
                'nombre_archivo': nombre
            }
            resultados.append(fila)

    return resultados

def main():
    archivos = obtener_archivos(CARPETA_TRAZOS)
    if not archivos:
        print("No se encontraron archivos válidos.")
        return

    todos = []
    with ProcessPoolExecutor(max_workers=os.cpu_count()) as executor:
        futures = {executor.submit(procesar_archivo, a): a for a in archivos}
        for f in tqdm(as_completed(futures), total=len(archivos),
                      desc="Procesando", bar_format='{l_bar}{r_bar}'):
            todos.extend(f.result())

    if not todos:
        print("No se extrajeron datos.")
        return

    df = pd.DataFrame(todos)
    columnas = [
        'central', 'troncal', 'hilo_troncal',
        'odf_troncal_salida', 'odf_troncal_entrada',
        'odf_patcheo_salida', 'odf_patcheo_entrada',
        'equipo_salida', 'equipo_entrada',
        'cable_acceso', 'hilo_acceso', 'anillo',
        'direccion_cliente', 'id_servicio', 'cliente',
        'nombre_archivo'
    ]
    for col in columnas:
        if col not in df.columns:
            df[col] = ''
    df = df[columnas]

    df.to_excel(ARCHIVO_SALIDA, index=False, engine='openpyxl')
    print(f"Archivo guardado en: {ARCHIVO_SALIDA}")

if __name__ == "__main__":
    main()