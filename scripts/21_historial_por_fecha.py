# Qué estamos aprendiendo?
# - defaultdict: crea automáticamente un registro para cada fecha.
# - split(): separa la fecha de la hora y el estado.
# - sorted(): ordena las fechas cronológicamente.
# - += 1: incrementa el contador de ejecuciones.
from pathlib import Path

from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
# PieChart() crea un gráfico circular.
# Reference() selecciona los datos de Excel.
# add_data() incorpora los valores.
# set_categories() asigna los nombres de cada categoría.
# showPercent muestra los porcentajes.

# Ubicacion del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"

ARCHIVO_LOG = CARPETA_REPORTES / "ejecuciones.log"


# Diccionario para agrupar las ejecuciones por fecha
historial = defaultdict(lambda: {
    "exitosas": 0,
    "errores": 0
})


# Leer el archivo de registro
with open(ARCHIVO_LOG, "r", encoding="utf-8") as archivo:

    for linea in archivo:

        linea = linea.strip()

        if not linea:
            continue

        # Ignorar lineas que no sean registros principales
        if len(linea) < 19 or linea[4] != "-" or linea[7] != "-":
            continue

        if " - " not in linea:
            continue

        fecha_hora, estado = linea.split(" - ", 1)

        fecha = fecha_hora.split(" ")[0]

        if "PROCESO COMPLETADO CORRECTAMENTE" in estado:
            historial[fecha]["exitosas"] += 1
        elif "ERROR al ejecutar" in estado:
            historial[fecha]["errores"] += 1

# Mostrar el historial
print()
print("=" * 50)
print("HISTORIAL DE EJECUCIONES POR FECHA")
print("=" * 50)

for fecha in sorted(historial):

    exitosas = historial[fecha]["exitosas"]
    errores = historial[fecha]["errores"]
    total = exitosas + errores

    print(f"Fecha: {fecha}")
    print(f"  Ejecuciones exitosas: {exitosas}")
    print(f"  Ejecuciones con errores: {errores}")
    print(f"  Total: {total}")
    print("-" * 40)

# Crear libro de Excel
# Qué aprenderemos con esta parte?
# - Crear libros y hojas de Excel con Python.
# - Insertar datos desde un diccionario.
# - Aplicar formatos profesionales.
# - Crear gráficos de barras automáticamente.
# - Guardar reportes en una ubicación específica.
libro = Workbook()
# Crear segunda hoja
hoja_resumen = libro.create_sheet("Resumen general")
hoja = libro.active
hoja.title = "Historial"

# Encabezados
encabezados = [
    "Fecha",
    "Ejecuciones exitosas",
    "Ejecuciones con errores",
    "Total de ejecuciones",
    "Porcentaje de exito",
    "Clasificacion"
]

hoja.append(encabezados)

# Agregar los datos
for fecha in sorted(historial):

    exitosas = historial[fecha]["exitosas"]
    errores = historial[fecha]["errores"]
    total = exitosas + errores

    porcentaje = (exitosas / total) * 100 if total > 0 else 0
    if porcentaje == 100:
        clasificacion = "Excelente"
    elif porcentaje >= 80:
        clasificacion = "Bueno"
    elif porcentaje >= 50:
        clasificacion = "Regular"
    else:
        clasificacion = "Necesita atencion"

    hoja.append([
        fecha,
        exitosas,
        errores,
        total,
        porcentaje / 100,
        clasificacion
    ])

# Formato de porcentaje
for fila in range(2, hoja.max_row + 1):
    hoja[f"E{fila}"].number_format = "0.00%"
    
# Formato del encabezado
for celda in hoja[1]:

    celda.font = Font(
        bold=True,
        color="FFFFFF"
    )

    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )

    celda.alignment = Alignment(
        horizontal="center"
    )

# Bordes
borde = Border(
    bottom=Side(style="thin", color="808080")
)

for fila in hoja.iter_rows():

    for celda in fila:
        celda.border = borde
        celda.alignment = Alignment(
            horizontal="center"
        )

# Ajustar ancho de columnas
anchos = {
    "A": 16,
    "B": 25,
    "C": 27,
    "D": 25,
    "E": 22,
    "F": 24
}

for columna, ancho in anchos.items():
    hoja.column_dimensions[columna].width = ancho

# Crear grafico
grafico = BarChart()

grafico.type = "col"
grafico.style = 10
grafico.title = "Ejecuciones por fecha"
grafico.y_axis.title = "Cantidad"
grafico.x_axis.title = "Fecha"

datos = Reference(
    hoja,
    min_col=2,
    max_col=3,
    min_row=1,
    max_row=hoja.max_row
)

categorias = Reference(
    hoja,
    min_col=1,
    min_row=2,
    max_row=hoja.max_row
)

grafico.add_data(datos, titles_from_data=True)
grafico.set_categories(categorias)

grafico.height = 9
grafico.width = 18


#hoja.add_chart(grafico, "F2") movimos la celda yaque la grafica invade esas columnas
hoja.add_chart(grafico, "H2")

# ========================================
# RESUMEN GENERAL
# ========================================

total_dias = len(historial)
total_exitosas = sum(
    datos["exitosas"] for datos in historial.values()
)
total_errores = sum(
    datos["errores"] for datos in historial.values()
)

total_ejecuciones = total_exitosas + total_errores

porcentaje_global = (
    total_exitosas / total_ejecuciones
) if total_ejecuciones > 0 else 0

# Clasificacion general
porcentaje = porcentaje_global * 100

if porcentaje == 100:
    clasificacion_general = "Excelente"
elif porcentaje >= 80:
    clasificacion_general = "Bueno"
elif porcentaje >= 50:
    clasificacion_general = "Regular"
else:
    clasificacion_general = "Necesita atencion"

# Titulo principal
hoja_resumen.merge_cells("A1:F1")
hoja_resumen["A1"] = "DASHBOARD DE CONTROL DE EJECUCIONES"

# Encabezados de la tabla
hoja_resumen.append([])
hoja_resumen.append(["Indicador", "Resultado"])


# Indicadores
indicadores = [
    ["Dias registrados", total_dias],
    ["Total de ejecuciones", total_ejecuciones],
    ["Ejecuciones exitosas", total_exitosas],
    ["Ejecuciones con errores", total_errores],
    ["Porcentaje global de exito", porcentaje_global],
    ["Clasificacion general", clasificacion_general]
]

for indicador in indicadores:
    hoja_resumen.append(indicador)


    
# Formato del encabezado
for celda in hoja_resumen[1]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    celda.alignment = Alignment(horizontal="center")

# Formato de los datos
for fila in hoja_resumen.iter_rows(min_row=2):
    for celda in fila:
        celda.border = borde

# Porcentaje global
hoja_resumen["B8"].number_format = "0.00%"

# Ancho de columnas
hoja_resumen.column_dimensions["A"].width = 30
hoja_resumen.column_dimensions["B"].width = 25
for fila in range(1, 9):
    hoja_resumen.row_dimensions[fila].height = 25

# Alinear resultados
for fila in range(2, hoja_resumen.max_row + 1):
    hoja_resumen[f"B{fila}"].alignment = Alignment(
        horizontal="center"
    )
    

# Formato del titulo principal
hoja_resumen["A1"].font = Font(
    bold=True,
    size=16,
    color="FFFFFF"
)

hoja_resumen["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

hoja_resumen["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

hoja_resumen.row_dimensions[1].height = 35

# ========================================
# DESTACAR INDICADORES PRINCIPALES
# ========================================

# Porcentaje global de exito
hoja_resumen["A8"].fill = PatternFill(
    fill_type="solid",
    fgColor="D9EAF7"
)

hoja_resumen["B8"].fill = PatternFill(
    fill_type="solid",
    fgColor="D9EAF7"
)

hoja_resumen["B8"].font = Font(
    bold=True,
    size=12
)

# Clasificacion general
hoja_resumen["A9"].fill = PatternFill(
    fill_type="solid",
    fgColor="E2F0D9"
)

hoja_resumen["B9"].fill = PatternFill(
    fill_type="solid",
    fgColor="E2F0D9"
)

hoja_resumen["B9"].font = Font(
    bold=True,
    size=12
)   

# Formato del titulo principal
hoja_resumen["A1"].font = Font(
    bold=True,
    size=16,
    color="FFFFFF"
)

hoja_resumen["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

hoja_resumen["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

hoja_resumen.row_dimensions[1].height = 35 

# ========================================
# GRAFICO CIRCULAR
# ========================================

grafico_circular = PieChart()

# Datos: exitosas y errores
datos = Reference(
    hoja_resumen,
    min_col=2,
    min_row=6,
    max_row=7
)

categorias = Reference(
    hoja_resumen,
    min_col=1,
    min_row=6,
    max_row=7
)

grafico_circular.add_data(datos)
grafico_circular.set_categories(categorias)

grafico_circular.title = "Ejecuciones exitosas vs errores"

# Mostrar porcentajes
grafico_circular.dataLabels = DataLabelList()

grafico_circular.dataLabels.showPercent = True
grafico_circular.dataLabels.showVal = False
grafico_circular.dataLabels.showCatName = False
grafico_circular.dataLabels.showSerName = False
grafico_circular.dataLabels.showLegendKey = False
grafico_circular.dataLabels.dLblPos = "bestFit"

grafico_circular.height = 9
grafico_circular.width = 14

# Ubicar el grafico
hoja_resumen.add_chart(grafico_circular, "D3")
    
# Guardar archivo
ARCHIVO_EXCEL = CARPETA_REPORTES / "historial_por_fecha.xlsx"

libro.save(ARCHIVO_EXCEL)

print()
print("Reporte Excel creado correctamente.")
print(f"Archivo: {ARCHIVO_EXCEL}")