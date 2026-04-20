import pyautogui
import time
import openpyxl

# 📍 1. Leer el valor de la celda E11 en Excel
ruta_excel = r"C:\A_GS1_PROYECTOS\0.GS.DB.V3.xlsm"
wb = openpyxl.load_workbook(ruta_excel, data_only=True)
hoja = wb.active
valor_celda = hoja["E11"].value
wb.close()

print(f"📌 Valor obtenido de Excel: {valor_celda}")


pyautogui.hotkey('alt', 'tab')
time.sleep(0.5) 

pyautogui.click(925, 469)  

pyautogui.click(1707, 103)

pyautogui.doubleClick(1705, 254)

# 📍 6. Escribir el valor y presionar Enter
pyautogui.write(str(valor_celda), interval=0)  # 🔹 Escritura instantánea
pyautogui.press('enter')

# 📍 7. Hacer clic en el icono de los anteojos y presionar Enter
pyautogui.click(1781, 97)
pyautogui.press('enter')

# 📍 8. Confirmar búsqueda
pyautogui.click(1508, 130)
pyautogui.press('enter')

print("✅ Búsqueda realizada en Smallworld (Ultra Rápida)")






import pyautogui
import time

while True:
    print("Mueve el mouse a la posición deseada y espera 3 segundos...")
    time.sleep(3)
    x, y = pyautogui.position()
    print(f"📍 Coordenadas: x={x}, y={y}")
    continuar = input("¿Quieres capturar otra posición? (s/n): ")
    if continuar.lower() != 's':
        break

