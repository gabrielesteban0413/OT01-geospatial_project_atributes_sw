import os

ruta = r"\\atlas\VP_DESARROLLO_DE_RED_Y_SERVICIOS\G_Aprovisionamiento_Servicios_Datos_Internet\D3910_Carr_Diseno_fo\Cables_Fibra_Óptica\Troncal de Acceso\Cartográfico\Regional"

for nombre in os.listdir(ruta):
    print(nombre)
