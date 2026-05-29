import os

carpeta = r"\\atlas\VP_INFRAESTRUCTURA\G_Plan_Gestion_Proyectos\RI200_A_Admon_Capacidad_P_R_Infr\2023\REPORTES_INF_CART_GEO\DICCIONARIO_DE_DATOS\40_CABLES_CORPORATIVOS\12_Cables_troncales_ conectados_en central"

try:
    archivos = [f for f in os.listdir(carpeta) if os.path.isfile(os.path.join(carpeta, f))]
    for archivo in archivos:
        print(archivo)
except FileNotFoundError:
    print("La carpeta no existe o la ruta es incorrecta.")
except PermissionError:
    print("No tienes permisos para acceder a la carpeta.")