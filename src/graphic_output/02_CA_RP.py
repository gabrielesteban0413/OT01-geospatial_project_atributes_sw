import arcpy

mxd = arcpy.mapping.MapDocument("CURRENT")
df = arcpy.mapping.ListDataFrames(mxd)[0]  


layer_names = {
    "reservas": "Reservas",
    "cable": "Cable_Corporativo_Existente",
    "empalme": "Empalme_Corporativo_Existente"
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


sum_reservas = sum_cable = 0
count_reservas = count_cable = 0
empalme_names = []


arcpy.RefreshActiveView()
arcpy.RefreshTOC()

try:
    
    if layers["reservas"]:
        with arcpy.da.SearchCursor(layers["reservas"], ["Longitud_Medida"]) as cursor:
            for row in cursor:
                if row[0]:
                    sum_reservas += row[0]
                    count_reservas += 1

    if layers["cable"]:
        with arcpy.da.SearchCursor(layers["cable"], ["Shape_Length"]) as cursor:
            for row in cursor:
                if row[0]:
                    sum_cable += row[0]
                    count_cable += 1

    if layers["empalme"]:
        arcpy.SelectLayerByAttribute_management(layers["empalme"], "CLEAR_SELECTION") 
        with arcpy.da.SearchCursor(layers["empalme"], ["NAME"]) as cursor:
            empalme_names = sorted(row[0] for row in cursor if row[0])

    highest_empalme_name = empalme_names[-1] if empalme_names else "N/A"

    print "================= REPORT ================="

    if layers["empalme"]:
        print "{}  E.N".format(highest_empalme_name)
    else:
        print "No se encontró la capa 'Empalme_Corporativo_Existente'."
        
    if layers["reservas"]:
        print "{}   R.C".format(count_reservas)
        print "{}   R.S".format(sum_reservas)
    else:
        print "No se encontró la capa 'Reservas'."

    if layers["cable"]:
        print "{}   F.S".format(sum_cable)
    else:
        print "No se encontró la capa 'Cable_Corporativo_Existente'."

    

finally:
    arcpy.RefreshActiveView()
    arcpy.RefreshTOC()
print "==================RP========================"

