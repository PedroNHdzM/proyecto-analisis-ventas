# Explicación de lo aprendido
# - Path: permite trabajar con rutas de archivos y carpetas.
# - datetime: obtiene la fecha y hora actuales.
# - shutil.copy2(): copia un archivo conservando sus datos de modificación.
# - mkdir(): crea la carpeta si todavía no existe.
# - exists(): comprueba que el reporte esté disponible.
# - strftime(): convierte la fecha y hora en texto.
from pathlib import Path
from datetime import datetime
import shutil
MAX_RESPALDOS = 10

# Rutas del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"
CARPETA_RESPALDOS = CARPETA_REPORTES / "respaldos"

# Crear carpeta de respaldos si no existe
CARPETA_RESPALDOS.mkdir(exist_ok=True)

# Lista de los cinco reportes
reportes = [
    "reporte_ventas.xlsx",
    "reporte_rentabilidad.xlsx",
    "dashboard_ventas.xlsx",
    "historial_por_fecha.xlsx",
    "reporte_calidad_datos.xlsx",
    "reporte_calidad_avanzado.xlsx",
    "reporte_general.xlsx",
    "dashboard_ejecutivo.xlsx"
]


def obtener_respaldos(nombre_reporte):
    # ¿Qué hace esta función?
    #     - glob() localiza los respaldos de un reporte específico.
    #     - stat().st_mtime obtiene la fecha de modificación.
    #     - sort() ordena los archivos.
    #     - reverse=True coloca primero el más reciente.
    nombre_base = Path(nombre_reporte).stem
    extension = Path(nombre_reporte).suffix

    patron = f"{nombre_base}_*{extension}"

    archivos = list(CARPETA_RESPALDOS.glob(patron))

    # Ordenar del mas reciente al mas antiguo
    archivos.sort(key=lambda archivo: archivo.stat().st_mtime,
                  reverse=True)

    return archivos



def eliminar_respaldos_antiguos(nombre_reporte):
    archivos = obtener_respaldos(nombre_reporte)
    eliminados = 0

    if len(archivos) > MAX_RESPALDOS:
        antiguos = archivos[MAX_RESPALDOS:]

        for archivo in antiguos:
            archivo.unlink()
            eliminados += 1
            print(f"[ELIMINADO] {archivo.name}")

    # Consultar cuantos respaldos quedan
    conservados = len(obtener_respaldos(nombre_reporte))

    return eliminados, conservados


# Fecha y hora para identificar el respaldo
fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S")

copiados = 0

print("CREANDO RESPALDO DE REPORTES")
print("-" * 40)

for nombre in reportes:
    archivo_origen = CARPETA_REPORTES / nombre

    if archivo_origen.exists():
        nombre_respaldo = (
            f"{archivo_origen.stem}_{fecha_hora}{archivo_origen.suffix}"
        )

        archivo_destino = CARPETA_RESPALDOS / nombre_respaldo

        shutil.copy2(archivo_origen, archivo_destino)

        copiados += 1
        print(f"[OK] Respaldo creado: {nombre_respaldo}")
        
        eliminados, conservados = eliminar_respaldos_antiguos(nombre)

        print(f"Respaldos eliminados: {eliminados}")
        print(f"Respaldos conservados: {conservados}")
    else:
        print(f"[ERROR] No se encontro: {nombre}")

print("-" * 40)
print(f"Total de respaldos creados: {copiados} de {len(reportes)}")
print("Proceso de respaldo finalizado.")
