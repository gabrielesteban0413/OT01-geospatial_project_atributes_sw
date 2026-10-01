import keyboard
import time

INTERVALO = 60 
DURACION_TOTAL = 4 * 60 * 60

inicio = time.time()
print("[INICIO] delete {} segundos.".format(INTERVALO))

while time.time() - inicio < DURACION_TOTAL:
    keyboard.send('f15')  
    time.sleep(INTERVALO)

print("[FIN] delete.")