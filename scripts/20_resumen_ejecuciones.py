# ¿Qué estamos aprendiendo?
# - readlines() lee todas las líneas del archivo.
# - for linea in lineas recorre cada una.
# - if y elif identifican el tipo de ejecución.
# - += 1 incrementa el contador en una unidad.
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import PieChart, Reference
from openpyxl.chart.label import DataLabelList
from datetime import datetime  #Esta herramienta nos permitirá obtener la fecha y hora actuales.
import re

# Obtener la ruta del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent

# Ruta del archivo log
ARCHIVO_LOG = CARPETA_PROYECTO / "reportes" / "ejecuciones.log"

# Contadores
exitosas = 0
errores = 0
total = 0
# Analizar cada linea
# ¿Qué estamos aprendiendo?
# - re.search() busca un patrón dentro de un texto.
# - ([\d.]+) identifica números enteros o decimales.
# - group(1) recupera el número encontrado.
# - float() convierte ese texto en un número decimal.
# - append() agrega el resultado a nuestra lista.

duraciones = []

# Leer el archivo
with open(ARCHIVO_LOG, "r", encoding="utf-8") as archivo:
    lineas = archivo.readlines()

for linea in lineas:

    if "PROCESO COMPLETADO CORRECTAMENTE" in linea:
        exitosas += 1

    elif "ERROR al ejecutar" in linea:
        errores += 1

    # Extraer la duracion
    if "Duracion:" in linea:
        try:
            texto_duracion = linea.split("Duracion:")[1]
            segundos = float(texto_duracion.split()[0])
            duraciones.append(segundos)
        except (ValueError, IndexError):
            continue

# Calcular el total real de ejecuciones
total = exitosas + errores



if total > 0:
    porcentaje = (exitosas / total) * 100
else:
    porcentaje = 0
    
# Calcular estadisticas de duracion
if duraciones:
    tiempo_promedio = sum(duraciones) / len(duraciones)
    tiempo_minimo = min(duraciones)
    tiempo_maximo = max(duraciones)
else:
    tiempo_promedio = 0
    tiempo_minimo = 0
    tiempo_maximo = 0

# Clasificar el resultado
#Python evalúa las condiciones de arriba hacia abajo .
# y asigna la primera clasificación que se cumple.
if porcentaje >= 95:
    clasificacion = "Excelente"
elif porcentaje >= 85:
    clasificacion = "Bueno"
elif porcentaje >= 75:
    clasificacion = "Regular"
else:
    clasificacion = "Necesita mejorar"
    
# Crear libro de Excel
libro = Workbook()
hoja = libro.active
hoja.title = "Resumen"

# Encabezados
hoja["A1"] = "Indicador"
hoja["B1"] = "Resultado"

# Datos
hoja["A2"] = "Ejecuciones exitosas"
hoja["B2"] = exitosas

hoja["A3"] = "Ejecuciones con errores"
hoja["B3"] = errores

hoja["A4"] = "Total de ejecuciones"
hoja["B4"] = total

hoja["A5"] = "Porcentaje de éxito"
hoja["B5"] = porcentaje / 100
hoja["B5"].number_format = "0.00%"

hoja["A6"] = "Clasificación"
hoja["B6"] = clasificacion

hoja["A7"] = "Tiempo promedio (segundos)"
hoja["B7"] = tiempo_promedio

hoja["A8"] = "Tiempo minimo (segundos)"
hoja["B8"] = tiempo_minimo

hoja["A9"] = "Tiempo maximo (segundos)"
hoja["B9"] = tiempo_maximo
# Formato de tiempos con dos decimales
for fila in range(7, 10):
    hoja[f"B{fila}"].number_format = "0.00"

# Ajustar ancho de columnas
hoja.column_dimensions["A"].width = 30
hoja.column_dimensions["B"].width = 20


# Formato de encabezados
for celda in hoja[1]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    celda.alignment = Alignment(horizontal="center")

# Formato de indicadores
for fila in range(2, 10):
    hoja[f"A{fila}"].font = Font(bold=True)
    hoja[f"B{fila}"].alignment = Alignment(horizontal="center")

# Bordes
borde = Side(style="thin", color="BFBFBF")

for fila in hoja.iter_rows(min_row=1, max_row=9, min_col=1, max_col=2):
    for celda in fila:
        celda.border = Border(
            left=borde,
            right=borde,
            top=borde,
            bottom=borde
        )

# Crear grafico circular
grafico = PieChart()

# Seleccionar los datos
datos = Reference(
    hoja,
    min_col=2,
    min_row=1,
    max_row=3
)

# Seleccionar las categorias
categorias = Reference(
    hoja,
    min_col=1,
    min_row=2,
    max_row=3
)

# Agregar datos y categorias
grafico.add_data(datos, titles_from_data=True)
grafico.set_categories(categorias)

# Titulo del grafico
grafico.title = "Ejecuciones del proyecto"

# Mostrar porcentajes
grafico.dataLabels = DataLabelList()
grafico.dataLabels.showPercent = True
grafico.dataLabels.showVal = False
grafico.dataLabels.showCatName = False
grafico.dataLabels.showSerName = False
grafico.dataLabels.showLegendKey = False
grafico.dataLabels.dLblPos = "bestFit"

# Insertar grafico en Excel
hoja.add_chart(grafico, "D2")
        
# Guardar archivo
# ARCHIVO_EXCEL = CARPETA_PROYECTO / "reportes" / "resumen_ejecuciones.xlsx"

# libro.save(ARCHIVO_EXCEL)
# Guardar reporte actual
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"

ARCHIVO_EXCEL = CARPETA_REPORTES / "resumen_ejecuciones.xlsx"

libro.save(ARCHIVO_EXCEL)

# Crear nombre con fecha y hora
fecha_hora = datetime.now().strftime("%Y%m%d_%H%M%S") #genera una fecha y hora como esta: 20261030_134200

nombre_historico = f"resumen_ejecuciones_{fecha_hora}.xlsx"

ARCHIVO_HISTORICO = CARPETA_REPORTES / nombre_historico

# Guardar copia historica
libro.save(ARCHIVO_HISTORICO) #Se crea una copia histórica con fecha y hora.

print()
print("Reportes Excel creados correctamente.")
print(f"Reporte actual: {ARCHIVO_EXCEL}")
print(f"Copia historica: {ARCHIVO_HISTORICO}")

# Mostrar resultados
print()
print("=" * 40)
print("RESUMEN DE EJECUCIONES")
print("=" * 40)

print(f"Ejecuciones exitosas: {exitosas}")
print(f"Ejecuciones con errores: {errores}")
print(f"Total de ejecuciones: {exitosas + errores}")
print(f"Porcentaje de exitos: {porcentaje:.2f}%")
print(f"Clasificacion: {clasificacion}")
print(f"Tiempo promedio: {tiempo_promedio:.2f} segundos")
print(f"Tiempo minimo: {tiempo_minimo:.2f} segundos")
print(f"Tiempo maximo: {tiempo_maximo:.2f} segundos")
print("=" * 40)

# Limpiar copias historicas antiguas
# Qué estamos aprendiendo?
# - glob() localiza archivos que coinciden con un patrón.
# - sort(reverse=True) ordena los nombres de más reciente a más antiguo.
# - [10:] selecciona las copias que exceden las primeras diez.
# - unlink() elimina un archivo del disco.
def limpiar_historial():

    archivos = list(
        CARPETA_REPORTES.glob("resumen_ejecuciones_*.xlsx")
    )

    # Identificar solamente archivos historicos validos
    patron = r"^resumen_ejecuciones_\d{8}_\d{6}\.xlsx$"

    archivos_historicos = [
        archivo for archivo in archivos
        if re.match(patron, archivo.name)
    ]

    # Ordenar del mas reciente al mas antiguo
    archivos_historicos.sort(reverse=True)

    # Eliminar los que excedan las 10 copias
    for archivo in archivos_historicos[10:]:
        archivo.unlink()
        print(f"Archivo historico eliminado: {archivo.name}")

    print(f"Copias historicas conservadas: {min(len(archivos_historicos), 10)}")
    
limpiar_historial()
    