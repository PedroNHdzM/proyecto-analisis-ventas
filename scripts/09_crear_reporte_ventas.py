
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference
from collections import defaultdict
from datetime import datetime
from openpyxl.utils import get_column_letter
from openpyxl.chart.layout import Layout, ManualLayout

# Rutas de los archivos
ruta_entrada = "Proyecto_Final/datos/ventas.xlsx"
ruta_salida = "Proyecto_Final/reportes/reporte_ventas.xlsx"

# Leer los datos originales
libro_entrada = load_workbook(ruta_entrada, data_only=True)
hoja_entrada = libro_entrada["Ventas"]

ventas = []

for fecha, producto, cantidad, precio in hoja_entrada.iter_rows(
    min_row=2, values_only=True
):
    if isinstance(fecha, str):
        fecha = datetime.strptime(fecha, "%Y-%m-%d")

    importe = cantidad * precio

    ventas.append({
        "fecha": fecha,
        "producto": producto,
        "cantidad": cantidad,
        "precio": precio,
        "importe": importe
    })

libro_entrada.close()

# Calcular el resumen por producto y por mes
productos = defaultdict(lambda: {"unidades": 0, "ingresos": 0})
meses = defaultdict(lambda: {"unidades": 0, "ingresos": 0})

for venta in ventas:
    producto = venta["producto"]
    mes = venta["fecha"].strftime("%Y-%m")

    productos[producto]["unidades"] += venta["cantidad"]
    productos[producto]["ingresos"] += venta["importe"]

    meses[mes]["unidades"] += venta["cantidad"]
    meses[mes]["ingresos"] += venta["importe"]

total_unidades = sum(v["cantidad"] for v in ventas)
total_ingresos = sum(v["importe"] for v in ventas)

# Crear un libro Excel nuevo
libro = Workbook()

# Hoja 1: Resumen
hoja = libro.active
hoja.title = "Resumen"

hoja.append(["REPORTE GENERAL DE VENTAS"])
hoja.append(["Indicador", "Resultado"])
hoja.append(["Número de ventas", len(ventas)])
hoja.append(["Unidades vendidas", total_unidades])
hoja.append(["Ingresos totales", total_ingresos])

# Hoja 2: Por producto
hoja_productos = libro.create_sheet("Por producto")
hoja_productos.append(["Producto", "Unidades", "Ingresos"])

for producto, datos in sorted(
    productos.items(),
    key=lambda elemento: elemento[1]["ingresos"],
    reverse=True
):
    hoja_productos.append([
        producto,
        datos["unidades"],
        datos["ingresos"]
    ])

# Hoja 3: Por mes
hoja_meses = libro.create_sheet("Por mes")
hoja_meses.append(["Mes", "Unidades", "Ingresos"])

for mes in sorted(meses):
    hoja_meses.append([
        mes,
        meses[mes]["unidades"],
        meses[mes]["ingresos"]
    ])


# Hoja 4: Analisis de productos
hoja_analisis = libro.create_sheet("Analisis de productos")

hoja_analisis.append([
    "Producto",
    "Unidades vendidas",
    "Ingresos"
])

# Ordenar productos por ingresos, de mayor a menor
for producto, datos in sorted(
    productos.items(),
    key=lambda elemento: elemento[1]["ingresos"],
    reverse=True
):
    hoja_analisis.append([
        producto,
        datos["unidades"],
        datos["ingresos"]
    ])

# Crear grafica comparativa de ingresos por producto
grafica_productos = BarChart()
grafica_productos.type = "bar"
grafica_productos.style = 10
grafica_productos.title = "Ingresos por producto"
grafica_productos.x_axis.title = "Ingresos ($)"
grafica_productos.y_axis.title = "Producto"
grafica_productos.height = 8
grafica_productos.width = 15

datos_grafica = Reference(
    hoja_analisis,
    min_col=3,
    min_row=1,
    max_row=hoja_analisis.max_row
)

categorias_grafica = Reference(
    hoja_analisis,
    min_col=1,
    min_row=2,
    max_row=hoja_analisis.max_row
)

grafica_productos.add_data(
    datos_grafica,
    titles_from_data=True
)
grafica_productos.set_categories(categorias_grafica)

hoja_analisis.add_chart(grafica_productos, "E2")


# Dar formato a las hojas
for hoja_actual in libro.worksheets:
    # Congelar encabezados
    hoja_actual.freeze_panes = "A2"

    # Configurar filtros
    if hoja_actual.title == "Resumen":
        hoja_actual.auto_filter.ref = "A2:B5"
    else:
        hoja_actual.auto_filter.ref = hoja_actual.dimensions

    # Formato de encabezados
    for celda in hoja_actual[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        celda.alignment = Alignment(horizontal="center")

    # Formato especial del título de Resumen
    if hoja_actual.title == "Resumen":
        hoja_actual.merge_cells("A1:B1")
        hoja_actual["A1"].font = Font(
            bold=True,
            color="FFFFFF",
            size=14
        )
        hoja_actual["A1"].fill = PatternFill(
            fill_type="solid",
            fgColor="1F4E78"
        )
        hoja_actual["A1"].alignment = Alignment(
            horizontal="center"
        )

        # Encabezados de la fila 2
        for celda in hoja_actual[2]:
            celda.font = Font(bold=True, color="FFFFFF")
            celda.fill = PatternFill(
                fill_type="solid",
                fgColor="1F4E78"
            )
            celda.alignment = Alignment(horizontal="center")

    
    # Ajustar ancho de columnas
    for numero_columna in range(1, hoja_actual.max_column + 1):
        letra = get_column_letter(numero_columna)
        #get_column_letter() obtiene la letra de la columna a partir de su número: 1 es A, 2 es B, 3 es C, etc. 
        # Así evitamos depender del tipo de celda que devuelve Excel, incluidas las celdas combinadas.
        ancho = max(
            len(str(hoja_actual.cell(fila, numero_columna).value or ""))
            for fila in range(1, hoja_actual.max_row + 1)
        )

        hoja_actual.column_dimensions[letra].width = max(
            ancho + 3, 14
        )


# Formato de moneda
hoja["B5"].number_format = '$#,##0.00'

for fila in range(2, hoja_productos.max_row + 1):
    hoja_productos.cell(fila, 3).number_format = '$#,##0.00'

for fila in range(2, hoja_meses.max_row + 1):
    hoja_meses.cell(fila, 3).number_format = '$#,##0.00'

# Formato de moneda en las tres hojas
hoja["B5"].number_format = '$#,##0.00'

for fila in range(2, hoja_productos.max_row + 1):
    hoja_productos.cell(fila, 3).number_format = '$#,##0.00'

for fila in range(2, hoja_meses.max_row + 1):
    hoja_meses.cell(fila, 3).number_format = '$#,##0.00'
    

    
# Gráfica de ingresos mensuales
grafica = BarChart()
grafica.type = "col"
grafica.style = 10
grafica.title = "Ingresos por mes"
grafica.y_axis.title = "Ingresos ($)"
grafica.x_axis.title = "Mes"
grafica.height = 8
grafica.width = 15
grafica.layout = Layout( manualLayout=ManualLayout(
        x=0.15,
        y=0.15,
        w=0.70,
        h=0.65
    )
)
datos = Reference(
    hoja_meses,
    min_col=3,
    min_row=1,
    max_row=hoja_meses.max_row
)
categorias = Reference(
    hoja_meses,
    min_col=1,
    min_row=2,
    max_row=hoja_meses.max_row
)

grafica.add_data(datos, titles_from_data=True)
grafica.set_categories(categorias)
hoja_meses.add_chart(grafica, "E2")

# Guardar el reporte nuevo
libro.save(ruta_salida)

print("REPORTE DE VENTAS CREADO")
print("-" * 35)
print(f"Ventas registradas: {len(ventas)}")
print(f"Unidades vendidas: {total_unidades}")
print(f"Ingresos totales: ${total_ingresos:,.2f}")
print(f"Archivo generado: {ruta_salida}")
