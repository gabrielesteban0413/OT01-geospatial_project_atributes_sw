import arcpy
import os

dwg = r"C:\A_GS1_PROYECTOS\MEADALBOSI34_L.dwg"
output_file = r"C:\A_GS1_PROYECTOS\dwg_textos.txt"

if not arcpy.Exists(dwg):
    print("ERROR")
else:
    with open(output_file, 'w') as f:
        f.write("TEXTOS DEL DWG\n")
        f.write("="*50 + "\n\n")
        
        annotation_path = os.path.join(dwg, "Annotation")
        
        if arcpy.Exists(annotation_path):
            cursor = arcpy.da.SearchCursor(annotation_path, ["Text"])
            
            for row in cursor:
                if row[0]:
                    texto = str(row[0])
                    texto = texto.replace('\\P', '\n')
                    texto = texto.replace('\\pxqc', '')
                    texto = texto.replace(';', '')
                    texto = texto.strip()
                    
                    if texto:
                        f.write(texto + "\n")
                        f.write("-"*40 + "\n")
            
            del cursor
    
    print("Archivo creado:", output_file)