#Validar archivos: comprobar que existan los archivos de Excel necesarios.
#Path(__file__).resolve().parent.parent
# Localiza la carpeta Proyecto_Final a partir de la ubicación del propio script. 
# Así evitamos depender de la carpeta desde la que ejecutemos Python.
# archivo.exists()
# Devuelve True si el archivo existe y False si no lo encuentra.
# todos_existen
# Es una variable lógica que empieza en True. Si falta algún archivo, cambia a False.
from pathlib import Path

# Carpeta principal del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent

# Archivos que necesita nuestro proyecto
archivos_necesarios = [
    CARPETA_PROYECTO / "datos" / "ventas.xlsx",
    CARPETA_PROYECTO / "datos" / "costos_productos.xlsx",
    CARPETA_PROYECTO / "reportes" / "reporte_rentabilidad.xlsx"
]

print("=" * 45)
print("VALIDACION DE ARCHIVOS DEL PROYECTO")
print("=" * 45)

todos_existen = True


for archivo in archivos_necesarios:
    try:
        if archivo.is_file():
            print(f"[OK] Encontrado: {archivo.name}")
        else:
            print(f"[ERROR] No encontrado: {archivo.name}")
            todos_existen = False

    except OSError as error:
        print(f"[ERROR] No se pudo comprobar {archivo.name}")
        print(f"Detalle: {error}")
        todos_existen = False


print("-" * 45)

if todos_existen:
    print("Todos los archivos necesarios existen.")
else:
    print("Faltan archivos. Revisa las rutas del proyecto.")
