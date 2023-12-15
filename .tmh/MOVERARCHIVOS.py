import os
import shutil
def mover_archivos(origen, destino):
    try:
        # Verificar si la carpeta de destino existe, si no, crearla
        if not os.path.exists(destino):
            os.makedirs(destino)

        # Listar archivos en la carpeta de origen
        archivos = os.listdir(origen)

        # Mover cada archivo al destino
        for archivo in archivos:
            origen_path = os.path.join(origen, archivo)
            destino_path = os.path.join(destino, archivo)
            shutil.move(origen_path, destino_path)

        print("Archivos movidos exitosamente.")
    except Exception as e:
        print(f"Error al mover archivos: {e}")

# Definir las carpetas de origen y destino
carpeta_origen = "C:\\Users\\alan.riquelmes\\Desktop\\DESCARGAS"
carpeta_destino = "C:\\Users\\alan.riquelmes\\Desktop\\FCA"

# Llamar a la función para mover archivos
mover_archivos(carpeta_origen, carpeta_destino)
