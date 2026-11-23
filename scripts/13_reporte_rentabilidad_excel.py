
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

# Rutas de los archivos
ruta_ventas = "Proyecto_Final/datos/ventas.xlsx"
ruta_costos = "Proyecto_Final/datos/costos_productos.xlsx"
ruta_reporte = "Proyecto_Final/reportes/reporte_rentabilidad.xlsx"

# Leer ventas
libro_ventas = load_workbook(ruta_ventas, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

# Leer costos
libro_costos = load_workbook(ruta_costos, data_only=True)
hoja_costos = libro_costos["Costos"]

# Guardar costos en un diccionario
costos = {}

for producto, costo in hoja_costos.iter_rows(
    min_row=2, max_col=2, values_only=True
):
    if producto is not None and costo is not None:
        costos[producto] = costo

# Calcular rentabilidad por producto
rentabilidad = {}

for fecha, producto, cantidad, precio in hoja_ventas.iter_rows(
    min_row=2, max_col=4, values_only=True
):
    if producto not in costos:
        raise ValueError(
            f"No se encontro el costo del producto: {producto}"
        )

    ingreso = cantidad * precio
    costo_total = cantidad * costos[producto]
    ganancia = ingreso - costo_total

    if producto not in rentabilidad:
        rentabilidad[producto] = {
            "unidades": 0,
            "ingresos": 0,
            "costos": 0,
            "ganancia": 0
        }

    rentabilidad[producto]["unidades"] += cantidad
    rentabilidad[producto]["ingresos"] += ingreso
    rentabilidad[producto]["costos"] += costo_total
    rentabilidad[producto]["ganancia"] += ganancia

# Crear el nuevo libro de Excel
libro_reporte = Workbook()
hoja = libro_reporte.active
hoja.title = "Rentabilidad"

# Titulo
hoja.merge_cells("A1:F1")
hoja["A1"] = "REPORTE DE RENTABILIDAD POR PRODUCTO"
hoja["A1"].font = Font(bold=True, color="FFFFFF", size=14)
hoja["A1"].fill = PatternFill(
    "solid", fgColor="1F4E78"
)
hoja["A1"].alignment = Alignment(horizontal="center")

# Encabezados
encabezados = [
    "Producto",
    "Unidades",
    "Ingresos",
    "Costos",
    "Ganancia",
    "Margen (%)"
]

for columna, encabezado in enumerate(encabezados, start=1):
    celda = hoja.cell(row=3, column=columna)
    celda.value = encabezado
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="4472C4")
    celda.alignment = Alignment(horizontal="center")

# Escribir los resultados
fila = 4

for producto, datos in rentabilidad.items():
    margen = datos["ganancia"] / datos["ingresos"] * 100

    hoja.cell(fila, 1, producto)
    hoja.cell(fila, 2, datos["unidades"])
    hoja.cell(fila, 3, datos["ingresos"])
    hoja.cell(fila, 4, datos["costos"])
    hoja.cell(fila, 5, datos["ganancia"])
    hoja.cell(fila, 6, margen)

    fila += 1

# Formato numérico
for fila_datos in range(4, fila):
    for columna in range(3, 6):
        hoja.cell(fila_datos, columna).number_format = '$#,##0.00'

    hoja.cell(fila_datos, 6).number_format = '0.00"%"'

# Fila de totales
fila_total = fila
hoja.cell(fila_total, 1, "TOTAL")
hoja.cell(fila_total, 2, sum(d["unidades"] for d in rentabilidad.values()))
hoja.cell(fila_total, 3, sum(d["ingresos"] for d in rentabilidad.values()))
hoja.cell(fila_total, 4, sum(d["costos"] for d in rentabilidad.values()))
hoja.cell(fila_total, 5, sum(d["ganancia"] for d in rentabilidad.values()))

ingresos_totales = hoja.cell(fila_total, 3).value
ganancia_total = hoja.cell(fila_total, 5).value
hoja.cell(
    fila_total, 6,
    ganancia_total / ingresos_totales * 100
)

for celda in hoja[fila_total]:
    celda.font = Font(bold=True)
    celda.fill = PatternFill("solid", fgColor="D9EAF7")

for columna in range(3, 6):
    hoja.cell(fila_total, columna).number_format = '$#,##0.00'

hoja.cell(fila_total, 6).number_format = '0.00"%"'

# Filtros
hoja.auto_filter.ref = f"A3:F{fila_total}"

# Grafico de ganancias
grafico = BarChart()
grafico.type = "bar"
grafico.style = 10
grafico.title = "Ganancia por producto"
grafico.x_axis.title = "Ganancia ($)"
grafico.y_axis.title = "Producto"
grafico.height = 7
grafico.width = 13

datos_grafico = Reference(
    hoja, min_col=5, min_row=3, max_row=fila - 1
)
categorias = Reference(
    hoja, min_col=1, min_row=4, max_row=fila - 1
)

grafico.add_data(datos_grafico, titles_from_data=True)
grafico.set_categories(categorias)
hoja.add_chart(grafico, "H3")

# Ajustar anchos de columnas
for columna in range(1, hoja.max_column + 1):
    letra = get_column_letter(columna)
    ancho = 0

    for fila_celda in range(1, hoja.max_row + 1):
        valor = hoja.cell(fila_celda, columna).value
        if valor is not None:
            ancho = max(ancho, len(str(valor)))

    hoja.column_dimensions[letra].width = max(ancho + 3, 14)

# Guardar el reporte
libro_reporte.save(ruta_reporte)

libro_ventas.close()
libro_costos.close()
libro_reporte.close()

print("Reporte de rentabilidad creado correctamente.")
print(f"Archivo: {ruta_reporte}")
print(f"Productos analizados: {len(rentabilidad)}")
print(f"Ingresos totales: ${ingresos_totales:,.2f}")
print(f"Costos totales: ${hoja.cell(fila_total, 4).value:,.2f}")
print(f"Ganancia total: ${ganancia_total:,.2f}")
print(f"Margen general: {ganancia_total / ingresos_totales * 100:.2f}%")
