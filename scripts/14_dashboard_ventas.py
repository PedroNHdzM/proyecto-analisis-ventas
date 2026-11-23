
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from openpyxl.utils import get_column_letter

ruta_ventas = "Proyecto_Final/datos/ventas.xlsx"
ruta_rentabilidad = "Proyecto_Final/reportes/reporte_rentabilidad.xlsx"

# Leer las ventas
libro_ventas = load_workbook(ruta_ventas, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

total_registros = 0
total_unidades = 0
ingresos_totales = 0

for fecha, producto, cantidad, precio in hoja_ventas.iter_rows(
    min_row=2, max_col=4, values_only=True
):
    if producto is not None:
        total_registros += 1
        total_unidades += cantidad
        ingresos_totales += cantidad * precio

# Leer la rentabilidad
libro_rentabilidad = load_workbook(
    ruta_rentabilidad, data_only=True
)
hoja_rentabilidad = libro_rentabilidad["Rentabilidad"]

costos_totales = 0
ganancia_total = 0

for producto, unidades, ingresos, costos, ganancia, margen in hoja_rentabilidad.iter_rows(
    min_row=4, max_col=6, values_only=True
):
    if producto == "TOTAL":
        costos_totales = costos
        ganancia_total = ganancia
        break

margen_general = ganancia_total / ingresos_totales * 100

print("\nINDICADORES DEL DASHBOARD")
print("-" * 40)
print(f"Registros de ventas: {total_registros}")
print(f"Unidades vendidas: {total_unidades}")
print(f"Ingresos totales: ${ingresos_totales:,.2f}")
print(f"Costos totales: ${costos_totales:,.2f}")
print(f"Ganancia total: ${ganancia_total:,.2f}")
print(f"Margen general: {margen_general:.2f}%")
#------------------------------------------------------------------------------
# Crear el archivo del dashboard
libro_dashboard = Workbook()
dashboard = libro_dashboard.active
dashboard.title = "Dashboard"

# Titulo principal
dashboard.merge_cells("A1:I1")
dashboard["A1"] = "DASHBOARD DE VENTAS Y RENTABILIDAD"
dashboard["A1"].font = Font(
    bold=True, size=16, color="FFFFFF"
)
dashboard["A1"].fill = PatternFill(
    "solid", fgColor="1F4E78"
)
dashboard["A1"].alignment = Alignment(horizontal="center")

# Indicadores principales
indicadores = [
    ("Ingresos totales", ingresos_totales),
    ("Costos totales", costos_totales),
    ("Ganancia total", ganancia_total),
    ("Margen general", margen_general)
]

posiciones = ["A3", "C3", "E3", "G3"]

for (titulo, valor), posicion in zip(indicadores, posiciones):
    columna = dashboard[posicion].column
    fila_titulo = dashboard[posicion].row

    celda_titulo = dashboard.cell(fila_titulo, columna)
    celda_titulo.value = titulo
    celda_titulo.font = Font(bold=True, color="FFFFFF")
    celda_titulo.fill = PatternFill(
        "solid", fgColor="4472C4"
    )
    celda_titulo.alignment = Alignment(horizontal="center")

    celda_valor = dashboard.cell(fila_titulo + 1, columna)
    celda_valor.value = valor
    celda_valor.font = Font(bold=True, size=13)
    celda_valor.alignment = Alignment(horizontal="center")

    if "Margen" in titulo:
        celda_valor.number_format = '0.00"%"'
    else:
        celda_valor.number_format = '$#,##0.00'

# Indicador de unidades vendidas
dashboard["I3"] = "Unidades vendidas"
dashboard["I3"].font = Font(bold=True, color="FFFFFF")
dashboard["I3"].fill = PatternFill(
    "solid", fgColor="4472C4"
)
dashboard["I3"].alignment = Alignment(horizontal="center")

dashboard["I4"] = total_unidades
dashboard["I4"].font = Font(bold=True, size=13)
dashboard["I4"].alignment = Alignment(horizontal="center")
dashboard["I4"].number_format = '#,##0'

# Datos para el gráfico de ventas mensuales
hoja_ventas = libro_ventas["Ventas"]
ventas_mensuales = {}

for fecha, producto, cantidad, precio in hoja_ventas.iter_rows(
    min_row=2, max_col=4, values_only=True
):
    if producto is None:
        continue

    if hasattr(fecha, "strftime"):
        mes = fecha.strftime("%Y-%m")
    else:
        mes = str(fecha)[:7]

    ventas_mensuales[mes] = (
        ventas_mensuales.get(mes, 0) + cantidad * precio
    )

dashboard["A7"] = "Mes"
dashboard["B7"] = "Ingresos"

for celda in dashboard[7][:2]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="4472C4")

for fila, (mes, ingresos) in enumerate(
    sorted(ventas_mensuales.items()), start=8
):
    dashboard.cell(fila, 1, mes)
    dashboard.cell(fila, 2, ingresos)
    dashboard.cell(fila, 2).number_format = '$#,##0.00'

grafico_mensual = BarChart()
grafico_mensual.type = "col"
grafico_mensual.title = "Ingresos por mes"
grafico_mensual.y_axis.title = "Ingresos ($)"
grafico_mensual.x_axis.title = "Mes"

datos = Reference(
    dashboard, min_col=2, min_row=7,
    max_row=7 + len(ventas_mensuales)
)
categorias = Reference(
    dashboard, min_col=1, min_row=8,
    max_row=7 + len(ventas_mensuales)
)

grafico_mensual.add_data(datos, titles_from_data=True)
grafico_mensual.set_categories(categorias)
grafico_mensual.height = 7
grafico_mensual.width = 12
dashboard.add_chart(grafico_mensual, "A13")

# Datos y gráfico de ganancias por producto
fila_productos = 7
dashboard.cell(fila_productos, 4, "Producto")
dashboard.cell(fila_productos, 5, "Ganancia")

for columna in (4, 5):
    celda = dashboard.cell(fila_productos, columna)
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill("solid", fgColor="4472C4")

# Leer la rentabilidad por producto desde Excel
rentabilidad = {}

for producto, unidades, ingresos, costos, ganancia, margen in (
    hoja_rentabilidad.iter_rows(
        min_row=4, max_col=6, values_only=True
    )
):
    if producto is None or producto == "TOTAL":
        continue

    rentabilidad[producto] = {
        "unidades": unidades,
        "ingresos": ingresos,
        "costos": costos,
        "ganancia": ganancia,
        "margen": margen
    }


for fila, (producto, datos_producto) in enumerate(
    rentabilidad.items(), start=8
):
    dashboard.cell(fila, 4, producto)
    dashboard.cell(
        fila, 5, datos_producto["ganancia"]
    )
    dashboard.cell(fila, 5).number_format = '$#,##0.00'

grafico_productos = BarChart()
grafico_productos.type = "bar"
grafico_productos.title = "Ganancia por producto"
grafico_productos.x_axis.title = "Ganancia ($)"
grafico_productos.y_axis.title = "Producto"

datos = Reference(
    dashboard, min_col=5, min_row=7,
    max_row=7 + len(rentabilidad)
)
categorias = Reference(
    dashboard, min_col=4, min_row=8,
    max_row=7 + len(rentabilidad)
)

grafico_productos.add_data(datos, titles_from_data=True)
grafico_productos.set_categories(categorias)
grafico_productos.height = 7
grafico_productos.width = 12
dashboard.add_chart(grafico_productos, "J7")

# Ajustar anchos de columnas
for columna in range(1, 10):
    letra = get_column_letter(columna)
    dashboard.column_dimensions[letra].width = 18

# Guardar el dashboard
ruta_dashboard = "Proyecto_Final/reportes/dashboard_ventas.xlsx"
libro_dashboard.save(ruta_dashboard)
libro_dashboard.close()

print(f"\nDashboard creado: {ruta_dashboard}")
#------------------------------------------------------------------------------


libro_ventas.close()
libro_rentabilidad.close()

