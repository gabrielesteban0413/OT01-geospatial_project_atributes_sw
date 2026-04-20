from pathlib import Path
import pandas as pd

# ───────────────────────────── CONSTANTES Y RUTAS ─────────────────────────────
RUTA_DB         = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality\03-LINEA CABLES PRIORIZADOS.xlsx")
BASE_PATH       = Path(r"C:\A_GS1_PROYECTOS\0_Documents_gs\output\quality")
RUTA_SW         = BASE_PATH / "01-Bulk export of fibers.xlsx"
RUTA_EXPORT     = BASE_PATH / "03FibrasEliminadas-2025-12-01.xlsx"


HOJA_DB         = "TD ACCES0S"
HOJA_SW         = "CableFibra"

COLUMNA_DB_ID   = "Cable acceso"
COLUMNA_DB_EST  = "Incorporación cable"
COLUMNA_SW_ID = "Name"  
COLUMNA_SW_EXTRA = "Id Text"  
ESTADO_VALIDO   = "COMPLETADO"

# ───────────────────────────── UTILIDADES ─────────────────────────────


def cargar_excel(ruta, hoja):
    df = pd.read_excel(ruta, sheet_name=hoja, dtype=str)
    df.fillna("", inplace=True)
    return df

def limpiar_codigo(codigo: str) -> str:
    return codigo.split("/")[0].strip()

def es_puenteo(codigo: str) -> bool:
    if not isinstance(codigo, str):
        return False
    codigo = codigo.strip().upper().replace(" ", "")
    return "/" in codigo or "P" in codigo

def cables_son_equivalentes(db_id: str, sw_id: str) -> bool:
    """Evalúa si un cable en DB y uno en SW pueden considerarse equivalentes bajo reglas de puenteo"""
    if not db_id or not sw_id:
        return False
    db_id, sw_id = db_id.strip(), sw_id.strip()
    if db_id == sw_id:
        return True

    if not (es_puenteo(db_id) and es_puenteo(sw_id)):
        return False

    partes_db = set(db_id.replace("P", "").split("/"))
    partes_sw = set(sw_id.replace("P", "").split("/"))

    return len(partes_db.intersection(partes_sw)) > 0


# ───────────────────────────── PROCESAMIENTO ─────────────────────────────
def procesar_cables(df_db, df_sw):
    df_db["Es_Puenteo_DB"] = df_db[COLUMNA_DB_ID].apply(es_puenteo)
    df_sw["Es_Puenteo_SW"] = df_sw[COLUMNA_SW_ID].apply(es_puenteo)
    df_db["Cable_Limpio"] = df_db[COLUMNA_DB_ID].apply(limpiar_codigo)
    df_sw["Cable_Limpio"] = df_sw[COLUMNA_SW_ID].apply(limpiar_codigo)


    df_completados = df_db[df_db[COLUMNA_DB_EST].str.upper() == ESTADO_VALIDO].copy()
    df_sw_match = df_sw[["Cable_Limpio", COLUMNA_SW_ID, COLUMNA_SW_EXTRA, "Es_Puenteo_SW"]].drop_duplicates()

    # PRIMER MATCH: EXACTO
    df_merge_exacto = df_completados.merge(
        df_sw_match,
        left_on=COLUMNA_DB_ID,
        right_on=COLUMNA_SW_ID,
        how="left",
        suffixes=("", "_sw")
    )
    df_merge_exacto["Match_Exacto"] = df_merge_exacto[COLUMNA_SW_ID].notna()

    # SEGUNDO MATCH: POR CABLE LIMPIO
    df_restantes = df_merge_exacto[df_merge_exacto[COLUMNA_SW_ID].isna()].copy()
    df_merge_limpio = df_restantes.drop(columns=[COLUMNA_SW_ID, COLUMNA_SW_EXTRA, "Es_Puenteo_SW", "Match_Exacto"])
    df_merge_limpio = df_merge_limpio.merge(
        df_sw_match,
        on="Cable_Limpio",
        how="left",
        suffixes=("", "_sw")
    )
    df_merge_limpio["Match_Limpio"] = (
        df_merge_limpio["Es_Puenteo_DB"] == df_merge_limpio["Es_Puenteo_SW"]
    )

    # TERCER MATCH: PUENTEO FLEXIBLE (OPTIMIZADO)
    df_restantes_2 = df_merge_limpio[~df_merge_limpio["Match_Limpio"] | df_merge_limpio[COLUMNA_SW_ID].isna()].copy()

    # Crear índice por tipo y cable limpio
    index_sw = {}
    for _, row in df_sw_match.iterrows():
        clave = (row["Es_Puenteo_SW"], row["Cable_Limpio"])
        index_sw.setdefault(clave, []).append(row)

    posibles_matches = []
    for _, fila in df_restantes_2.iterrows():
        clave = (fila["Es_Puenteo_DB"], fila["Cable_Limpio"])
        candidatos = index_sw.get(clave, [])

        for row_sw in candidatos:
            if cables_son_equivalentes(fila[COLUMNA_DB_ID], row_sw[COLUMNA_SW_ID]):
                nueva_fila = fila.copy()
                nueva_fila[COLUMNA_SW_ID] = row_sw[COLUMNA_SW_ID]
                nueva_fila[COLUMNA_SW_EXTRA] = row_sw[COLUMNA_SW_EXTRA]
                nueva_fila["Es_Puenteo_SW"] = row_sw["Es_Puenteo_SW"]
                nueva_fila["Match_Parcial"] = True
                posibles_matches.append(nueva_fila)
                break

    df_match_parcial = pd.DataFrame(posibles_matches)



    df_final = pd.concat([
        df_merge_exacto[df_merge_exacto["Match_Exacto"]],
        df_merge_limpio[df_merge_limpio["Match_Limpio"]],
        df_match_parcial
    ], ignore_index=True)

    # ENCONTRADOS
    df_encontrados = df_final.copy()
    df_encontrados["Nombre en SW"] = df_encontrados[COLUMNA_SW_ID]
    df_encontrados["Nombre en DB"] = df_encontrados[COLUMNA_DB_ID]
    df_encontrados = df_encontrados.drop_duplicates(subset=["Cable acceso"], keep="first")

    # FALTANTES
    encontrados_ids = set(df_encontrados["Cable acceso"])
    df_faltantes_raw = df_completados[~df_completados["Cable acceso"].isin(encontrados_ids)].copy()

    def agrupar_coincidencias(grupo):
        cable_limpio = limpiar_codigo(grupo[COLUMNA_DB_ID].iloc[0])
        posibles = df_sw[df_sw["Cable_Limpio"] == cable_limpio][COLUMNA_SW_ID].unique()
        coincidencias = ", ".join(sorted(posibles)) if len(posibles) else ""
        fila = grupo.iloc[0].copy()
        fila["Posibles coincidencias en SW"] = coincidencias
        return fila

    df_faltantes = df_faltantes_raw.groupby("Cable acceso", as_index=False).apply(agrupar_coincidencias)

    columnas_exportar = [
        COLUMNA_SW_EXTRA,
        "Nombre en SW",
        "Nombre en DB",
        "Cable acceso",
        "Responsable red externa",
        "Fecha inicio",
        "Fecha Fin",
        "Posibles coincidencias en SW"
    ]
    df_encontrados = df_encontrados[columnas_exportar[:-1]]

    return (
        df_faltantes.drop(columns=[
            "Cable_Limpio", "Es_Puenteo_DB", "Es_Puenteo_SW",
            "Match_Exacto", "Match_Limpio", "Match_Parcial"
        ], errors="ignore"),
        df_encontrados
    )

# ───────────────────────────── EXPORTADOR ─────────────────────────────
def exportar_resultados(df_faltantes, df_encontrados, ruta_export):
    with pd.ExcelWriter(ruta_export, engine="openpyxl") as writer:
        df_faltantes.to_excel(writer, sheet_name="N", index=False)
        df_encontrados.to_excel(writer, sheet_name="S", index=False) 
        

        for hoja in writer.book.sheetnames:
            ws = writer.book[hoja]
            for col in ws.iter_cols(min_row=2, max_row=ws.max_row):
                col_letter = col[0].column_letter
                col_name = ws[f"{col_letter}1"].value or ""
                if "FECHA" in col_name.upper():
                    for cell in col:
                        if isinstance(cell.value, str):
                            try:
                                cell.value = pd.to_datetime(cell.value).date()
                                cell.number_format = "dd/mm/yyyy"
                            except:
                                continue
                        elif hasattr(cell.value, "strftime"):
                            cell.number_format = "dd/mm/yyyy"

    print(f"---------MERGE_DB_SW-----------: N-{len(df_faltantes)} | S-{len(df_encontrados)}")

# ───────────────────────────── FUNCIÓN PRINCIPAL ─────────────────────────────
def main():
    df_db = cargar_excel(RUTA_DB, HOJA_DB)
    df_sw = cargar_excel(RUTA_SW, HOJA_SW)
    print("Procesando comparación")
    df_faltantes, df_encontrados = procesar_cables(df_db, df_sw)
    exportar_resultados(df_faltantes, df_encontrados, RUTA_EXPORT)

# ───────────────────────────── PUNTO DE ENTRADA ─────────────────────────────
if __name__ == "__main__":
    main()
