# import os
# from oletools.olevba3 import VBA_Parser, VBA_Scanner

# ruta_archivo = r"C:\A_GS1_PROYECTOS\ARROBA DE PAN DEMO - copia.xlsm"

# # Copia del archivo original por seguridad
# ruta_backup = ruta_archivo.replace(".xlsm", "_backup.xlsm")
# os.system(f'copy "{ruta_archivo}" "{ruta_backup}"')

# # Abrimos el archivo con VBA_Parser
# vbaparser = VBA_Parser(ruta_archivo)

# if vbaparser.detect_vba_macros():
#     for (filename, stream_path, vba_filename, vba_code) in vbaparser.extract_macros():
#         if "UserForm" in vba_filename or "Login" in vba_filename or "frmLogin" in vba_filename:
#             print(f"❌ Eliminar: {vba_filename}")
#             # Aquí deberíamos eliminar, pero VBA_Parser solo permite lectura
#             # Eliminar el módulo requiere otra herramienta como vba-editor o cambiar a Open XML
#         else:
#             print(f"✅ Conservar: {vba_filename}")
# else:
#     print("No se detectaron macros.")

# vbaparser.close()















import os
import shutil
import win32com.client

# === CONFIGURACIÓN ===
ruta_archivo = r"C:\A_GS1_PROYECTOS\ARROBA DE PAN DEMO - copia.xlsm"
nombre_login = "login"  # Parte del nombre del formulario (en minúsculas)

# === 1. CREAR BACKUP ===
ruta_backup = ruta_archivo.replace(".xlsm", "_backup.xlsm")
shutil.copy(ruta_archivo, ruta_backup)
print(f"🔒 Backup creado: {ruta_backup}")

# === 2. INICIAR EXCEL Y ABRIR ARCHIVO ===
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False  # Evita ventanas emergentes

wb = excel.Workbooks.Open(ruta_archivo)
vbproject = wb.VBProject

# === 3. ELIMINAR FORMULARIO LOGIN ===
for component in list(vbproject.VBComponents):
    nombre = component.Name.lower()
    if nombre_login in nombre and component.Type == 3:  # 3 = Formulario
        vbproject.VBComponents.Remove(component)
        print(f"❌ Formulario eliminado: {component.Name}")

# === 4. ELIMINAR CÓDIGO QUE LLAMA AL FORMULARIO ===
def eliminar_lineas_referencias(module):
    total_lineas = module.CountOfLines
    lineas = module.Lines(1, total_lineas).split('\n')

    nuevas_lineas = [
        linea for linea in lineas if nombre_login not in linea.lower()
    ]

    if len(nuevas_lineas) < len(lineas):
        print(f"🧹 Limpiando referencias en: {module.Parent.Name}")
        module.DeleteLines(1, total_lineas)
        module.InsertLines(1, "\n".join(nuevas_lineas))

# Recorre todos los componentes (módulos, hojas, formularios)
for component in vbproject.VBComponents:
    try:
        module = component.CodeModule
        eliminar_lineas_referencias(module)
    except Exception as e:
        print(f"⚠️ No se pudo analizar: {component.Name} ({e})")

# === 5. GUARDAR Y CERRAR ===
wb.Save()
wb.Close(False)
excel.Quit()

print("✅ Proceso completado.")
