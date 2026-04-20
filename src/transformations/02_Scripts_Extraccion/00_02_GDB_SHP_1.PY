import pyautogui
import time
import pygetwindow as gw
import win32com.client
import os

# === Abrir Excel y leer nombre base ===
excel = win32com.client.Dispatch("Excel.Application")
wb = excel.Workbooks("0.GS.DB.V3.xlsm")
sheet = wb.ActiveSheet
base_name = str(sheet.Range("B1").Value).strip()
target_file = f"{base_name}_T.dwg"

# === Verificar si ArcCatalog está abierto ===
arc_catalog_open = False
for win in gw.getWindowsWithTitle("ArcCatalog"):
    if "ArcCatalog" in win.title:
        arc_catalog_open = True
        win.minimize()
        win.restore()
        break

# === Si no está abierto, abrirlo ===
if not arc_catalog_open:
    pyautogui.hotkey("win", "r")
    time.sleep(0.5)
    pyautogui.write("ArcCatalog")
    pyautogui.press("enter")
    time.sleep(5)

time.sleep(0.3)

# === Interacción con ArcCatalog ===
pyautogui
pyautogui.doubleClick(929, 384)  # CAD de entrada 
time.sleep(0.3)
pyautogui.press("o")  # Folder
pyautogui.press("enter")

time.sleep(0.5)
pyautogui.click(879, 164)  # Carpeta 

time.sleep(0.2)
pyautogui.write(base_name, interval=0.05)
pyautogui.press("enter")
time.sleep(1)

pyautogui.write(target_file)
time.sleep(0.2)
pyautogui.press("enter")
time.sleep(0.5)
pyautogui.press("down")
time.sleep(0.2)
pyautogui.press("enter")

print(f" archivo correcto: {target_file}")

# === Click en el botón RUN === 
time.sleep(0.5)
pyautogui.click(1311, 196) 

# === Esperar a que aparezca el botón "Close" ===
close_img_path = r"C:\A_GS1_PROYECTOS\00_GS_IMPORT\A_PROPS_VS\IMG\close_button.png"
intentos = 0
max_intentos = 30

while intentos < max_intentos:
    location = pyautogui.locateOnScreen(close_img_path, confidence=0.8)
    intentos += 1
    print(f"Intento {intentos} - {'Encontrado' if location else 'No encontrado'}")

    if location:
        pyautogui.click(pyautogui.center(location))
        time.sleep(0.8)
        pyautogui.click(1050, 88)  # "X"  
        time.sleep(0.8)
        pyautogui.click(948, 592)  # "Sí"
        print("✅ Botón 'Close' detectado y clickeado.")
        break

    time.sleep(1)

if intentos == max_intentos:
    print("❌ No se encontró el botón 'Close' tras varios intentos.")

print("✅ Proceso finalizado correctamente.")



