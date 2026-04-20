import arcpy
import xlwings as xw

excel_path = r"C:\A_GS1_PROYECTOS\0.GS.DB.V3.xlsm"
sheet_name = "DB"
cell_refs = {"cable": "E6", "empalme": "E7", "reservas": "E8", "nodo": "E9"}
wb = xw.Book(excel_path)
sheet = wb.sheets[sheet_name]
def_queries = {key: sheet.range(cell).value for key, cell in cell_refs.items()}

mxd = arcpy.mapping.MapDocument("CURRENT")
df = arcpy.mapping.ListDataFrames(mxd)[0]

layer_names = {
    "reservas": "Reservas",
    "cable": "Cable_Corporativo_Existente",
    "empalme": "Empalme_Corporativo_Existente",
    "nodo": "Nodo_Central"
}

def find_layer(layers, name):
    for lyr in layers:
        if lyr.isGroupLayer:
            found_layer = find_layer(lyr, name)
            if found_layer:
                return found_layer
        elif lyr.name == name:
            return lyr
    return None

layers_list = arcpy.mapping.ListLayers(mxd, "", df)
layers = {key: find_layer(layers_list, name) for key, name in layer_names.items()}

for key, layer in layers.items():
    if layer:
        layer.definitionQuery = def_queries[key]
        print("FT {}: {}".format(layer_names[key], def_queries[key].encode('utf-8')))
    else:
        print("No se encontró la capa '{}'.".format(layer_names[key].encode('utf-8')))

# Hacer zoom a la capa de Cable_Corporativo_Existente si existe
df.extent = layers["cable"].getExtent() if layers["cable"] else df.extent

arcpy.RefreshActiveView()
arcpy.RefreshTOC()
print("====================FT======================".encode('utf-8'))
