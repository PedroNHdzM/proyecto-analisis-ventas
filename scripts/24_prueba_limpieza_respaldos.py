# glob() para localizar archivos.
# sort() para ordenar respaldos.
# stat().st_mtime para consultar fechas.
# unlink() para eliminar archivos.
# TemporaryDirectory() para hacer pruebas sin tocar los archivos reales.

from pathlib import Path
from datetime import datetime, timedelta
import tempfile
import os

MAX_RESPALDOS = 10

# Crear una carpeta temporal para la prueba
with tempfile.TemporaryDirectory() as temporal:
    carpeta_prueba = Path(temporal)

    print("CREANDO 12 RESPALDOS DE PRUEBA")
    print("-" * 40)

    # Crear 12 archivos ficticios
    fecha_base = datetime(2026, 11, 1)

    for i in range(12):
        fecha = fecha_base + timedelta(days=i)

        nombre = f"reporte_ventas_{fecha:%Y%m%d}_120000.xlsx"
        archivo = carpeta_prueba / nombre

        archivo.write_text("Respaldo de prueba")

        # Asignar fecha de modificacion
        marca_tiempo = fecha.timestamp()
        os.utime(archivo, (marca_tiempo, marca_tiempo))

    print(f"Respaldos creados: {len(list(carpeta_prueba.glob('*.xlsx')))}")

    # Obtener y ordenar los respaldos
    archivos = list(carpeta_prueba.glob("reporte_ventas_*.xlsx"))

    archivos.sort(
        key=lambda archivo: archivo.stat().st_mtime,
        reverse=True
    )

    # Identificar y eliminar los antiguos
    antiguos = archivos[MAX_RESPALDOS:]

    for archivo in antiguos:
        archivo.unlink()
        print(f"[ELIMINADO] {archivo.name}")

    # Comprobar el resultado
    restantes = list(carpeta_prueba.glob("*.xlsx"))

    print("-" * 40)
    print(f"Respaldos conservados: {len(restantes)}")
    print("Prueba finalizada.")
