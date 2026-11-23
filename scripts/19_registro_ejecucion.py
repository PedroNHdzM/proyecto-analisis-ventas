# datetime.now() obtiene la fecha y hora.
# strftime() permite darles formato.
# open(..., "a") conserva el historial y agrega nuevos registros.
# Un archivo log ayuda a revisar qué ocurrió durante las ejecuciones.

from datetime import datetime
from pathlib import Path

# Ubicacion principal del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"

# Archivo donde guardaremos el historial
ARCHIVO_LOG = CARPETA_REPORTES / "ejecuciones.log"

# Obtener la fecha y hora actuales
fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Registrar el resultado de una ejecucion
mensaje = f"{fecha_hora} - PROCESO COMPLETADO CORRECTAMENTE\n"

with open(ARCHIVO_LOG, "a", encoding="utf-8") as archivo:
    archivo.write(mensaje)

print("Registro guardado correctamente.")
print(f"Archivo: {ARCHIVO_LOG}")
print(f"Mensaje: {mensaje.strip()}")
