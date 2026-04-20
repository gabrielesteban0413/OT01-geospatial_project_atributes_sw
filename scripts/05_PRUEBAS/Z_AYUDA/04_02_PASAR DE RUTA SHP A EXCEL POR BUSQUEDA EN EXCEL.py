import os
import win32com.client


BASE_PATH = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\02_ConversionCAD_Shape\Shape_ MCB"
MAX_FILAS_SIN_DATOS = 10
FILA_INICIO = 31
COLUMNA_NOMBRE = 2
COLUMNA_RESULTADO = 3
COLUMNA_RUTA = 24

#utilidades
def _add_to_map_list(mapa, clave, valor):
    clave = clave.lower()
    if clave not in mapa:
        mapa[clave] = []
    mapa[clave].append(valor)

def buscar_shp_prioritario_en_lista(paths, nombre_base, tipo_busqueda):
    """
    Dada una lista de rutas .shp dentro de la carpeta objetivo,
    devuelve la ruta priorizando:
      1) Polyline_1.shp
      2) {nombre_base}_T o {nombre_base}_P según tipo A/C
      3) El .shp más reciente
    """
    if not paths:
        return None

    
    for p in paths:
        if os.path.basename(p).lower() == "polyline_1.shp":
            return p

    nombre_base = str(nombre_base).strip()
    pref1 = f"{nombre_base}_T".lower() if tipo_busqueda == "a" else f"{nombre_base}_P".lower()
    pref2 = f"{nombre_base}_P".lower() if tipo_busqueda == "a" else f"{nombre_base}_T".lower()

    for p in paths:
        if os.path.splitext(os.path.basename(p))[0].lower() == pref1:
            return p

    for p in paths:
        if os.path.splitext(os.path.basename(p))[0].lower() == pref2:
            return p

    paths_exist = [p for p in paths if os.path.exists(p)]
    if not paths_exist:
        return None
    return max(paths_exist, key=lambda p: os.path.getmtime(p))


def indexar_archivos_shp_eficiente(base_path):

    por_nombre = {}
    por_carpeta = {}

    for root, _, files in os.walk(base_path):
        carpeta_actual = os.path.basename(root)
        carpeta_padre = os.path.basename(os.path.dirname(root))

        for nombre in files:
            if not nombre.lower().endswith(".shp"):
                continue

            ruta = os.path.join(root, nombre)
            if not os.path.exists(ruta):
                continue

            stem = os.path.splitext(nombre)[0]

            if (stem not in por_nombre) or (os.path.getmtime(ruta) > os.path.getmtime(por_nombre[stem])):
                por_nombre[stem] = ruta

            if carpeta_actual:
                _add_to_map_list(por_carpeta, carpeta_actual, ruta)
            if carpeta_padre:
                _add_to_map_list(por_carpeta, carpeta_padre, ruta)

    return por_nombre, por_carpeta


def obtener_excel_activo():
    try:
        xl = win32com.client.Dispatch("Excel.Application")
        wb = xl.ActiveWorkbook
        ws = wb.ActiveSheet
        return xl, wb, ws
    except Exception:
        return None, None, None


def pedir_tipo_busqueda(xl):
    """
    A = prioriza _T
    C = prioriza _P
    """
    try:
        tipo = str(xl.InputBox("Ingrese tipo de búsqueda (A o C):", "Tipo de Búsqueda", "A")).strip().lower()
        return tipo if tipo in ("a", "c") else None
    except Exception:
        return None


def seleccionar_mejor_ruta(nombre_base, tipo_busqueda, por_nombre, por_carpeta):
    """
    Intenta primero por nombre exacto con prioridad según A/C,
    luego por carpeta (con prioridad Polyline_1.shp), y devuelve la ruta.
    """
    nombre_base = str(nombre_base).strip()
    candidatos = (
        [f"{nombre_base}_T", f"{nombre_base}_P"]
        if tipo_busqueda == "a"
        else [f"{nombre_base}_P", f"{nombre_base}_T"]
    )

    for cand in candidatos:
        ruta = por_nombre.get(cand)
        if ruta and os.path.exists(ruta):
            return ruta

    claves_carpeta = [nombre_base.lower()]
    claves_carpeta += [nombre_base.replace(" ", "_").lower(),
                       nombre_base.replace("_", "").lower()]

    for key in claves_carpeta:
        paths = por_carpeta.get(key)
        if paths:
            ruta = buscar_shp_prioritario_en_lista(paths, nombre_base, tipo_busqueda)
            if ruta:
                return ruta

    return None



# PRINCIPAL

def buscar_y_actualizar_shp_directo():
    xl, wb, ws = obtener_excel_activo()
    if not xl:
        return

    tipo_busqueda = pedir_tipo_busqueda(xl)
    if not tipo_busqueda:
        return


    por_nombre, por_carpeta = indexar_archivos_shp_eficiente(BASE_PATH)
    rango_nombres = ws.Range(
        ws.Cells(FILA_INICIO, COLUMNA_NOMBRE),
        ws.Cells(FILA_INICIO + MAX_FILAS_SIN_DATOS + 100, COLUMNA_NOMBRE)
    ).Value

    resultados, rutas = [], []
    sin_datos_consecutivos = 0

    for fila_val in rango_nombres:
        valor_celda = fila_val[0] if isinstance(fila_val, (list, tuple)) else fila_val

        if not valor_celda or str(valor_celda).strip().lower() == "none":
            sin_datos_consecutivos += 1
            if sin_datos_consecutivos >= MAX_FILAS_SIN_DATOS:
                break
            resultados.append("")
            rutas.append("")
            continue

        sin_datos_consecutivos = 0
        nombre_base = str(valor_celda).strip()

        ruta = seleccionar_mejor_ruta(nombre_base, tipo_busqueda, por_nombre, por_carpeta)

        if ruta:
            resultados.append(ruta)  
            rutas.append(ruta)  
        else:
            resultados.append("")
            rutas.append("")

    if resultados:
        ws.Range(ws.Cells(FILA_INICIO, COLUMNA_RESULTADO),
                 ws.Cells(FILA_INICIO + len(resultados) - 1, COLUMNA_RESULTADO)).Value = [(x,) for x in resultados]

    if rutas:
        ws.Range(ws.Cells(FILA_INICIO, COLUMNA_RUTA),
                 ws.Cells(FILA_INICIO + len(rutas) - 1, COLUMNA_RUTA)).Value = [(x,) for x in rutas]



if __name__ == "__main__":
    buscar_y_actualizar_shp_directo()
