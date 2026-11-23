# Workbook: permitirá crear un nuevo Excel.
# BarChart: permitirá construir gráficos de barras.
# Reference: permitirá seleccionar los datos que alimentarán los gráficos.
# PatternFill, Font y Alignment: nos ayudarán con el diseño.
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import BarChart, Reference

# Identificar carpetas
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent

# Rutas de archivos
RUTA_VENTAS = CARPETA_PROYECTO / "datos" / "ventas.xlsx"
RUTA_COSTOS = CARPETA_PROYECTO / "datos" / "costos_productos.xlsx"

# Carpeta de reportes
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"

print("DASHBOARD EJECUTIVO DE VENTAS")
print("----------------------------------------")

print("Archivo de ventas:", RUTA_VENTAS.name)
print("Archivo de costos:", RUTA_COSTOS.name)
print("Carpeta de reportes:", CARPETA_REPORTES.name)

# Abrir archivos de origen
libro_ventas = load_workbook(RUTA_VENTAS, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

libro_costos = load_workbook(RUTA_COSTOS, data_only=True)
hoja_costos = libro_costos["Costos"]

print("\nARCHIVOS ABIERTOS CORRECTAMENTE")
print("----------------------------------------")
print("Filas de ventas:", hoja_ventas.max_row - 1)
print("Filas del catálogo de costos:", 4)

# Calcular indicadores de ventas
total_ventas = 0
total_unidades = 0
total_ingresos = 0
total_costos = 0

# Recorrer las ventas
for fila in range(2, hoja_ventas.max_row + 1):

    producto = hoja_ventas.cell(fila, 2).value
    cantidad = hoja_ventas.cell(fila, 3).value
    precio = hoja_ventas.cell(fila, 4).value

    # Buscar el costo del producto
    costo_unitario = 0

    for fila_costo in range(2, 6):
        nombre = hoja_costos.cell(fila_costo, 1).value

        if nombre == producto:
            costo_unitario = hoja_costos.cell(fila_costo, 2).value
            break

    # Acumular resultados
    total_ventas += 1
    total_unidades += cantidad
    total_ingresos += cantidad * precio
    total_costos += cantidad * costo_unitario

# Calcular utilidad y margen
utilidad = total_ingresos - total_costos
margen = (utilidad / total_ingresos) * 100

print("\nINDICADORES DEL DASHBOARD")
print("----------------------------------------")
print("Total de ventas:", total_ventas)
print("Unidades vendidas:", total_unidades)
print("Ingresos: $", total_ingresos)
print("Costos: $", total_costos)
print("Utilidad: $", utilidad)
print("Margen: {:.2f}%".format(margen))

# Crear libro del dashboard
libro_dashboard = Workbook()

# Seleccionar hoja principal
hoja_dashboard = libro_dashboard.active
hoja_dashboard.title = "Dashboard Ejecutivo"

# Crear título
hoja_dashboard.merge_cells("A1:G1")
hoja_dashboard["A1"] = "DASHBOARD EJECUTIVO DE VENTAS"

hoja_dashboard["A1"].font = Font(
    name="Arial",
    size=10,
    bold=True,
    color="FFFFFF"
)

hoja_dashboard["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

hoja_dashboard["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

hoja_dashboard.row_dimensions[1].height = 30

print("\nESTRUCTURA DEL DASHBOARD CREADA")

# Encabezados de indicadores
hoja_dashboard["A3"] = "INDICADOR"
hoja_dashboard["B3"] = "RESULTADO"

# Lista de indicadores
indicadores = [
    ("Total de ventas", total_ventas),
    ("Unidades vendidas", total_unidades),
    ("Ingresos", total_ingresos),
    ("Costos", total_costos),
    ("Utilidad", utilidad),
    ("Margen de utilidad", margen / 100)
]

# Escribir indicadores en Excel
for fila, (nombre, valor) in enumerate(indicadores, start=4):
    hoja_dashboard.cell(fila, 1, nombre)
    hoja_dashboard.cell(fila, 2, valor)

# Aplicar formato solo a los encabezados de A3 y B3
for columna in range(1, 3):
    celda = hoja_dashboard.cell(3, columna)

    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )

# Aplicar formato numérico
for fila in range(4, 10):
    hoja_dashboard.cell(fila, 2).number_format = '#,##0.00'

# Formato de porcentaje para el margen
hoja_dashboard["B9"].number_format = "0.00%"

print("Indicadores incorporados al dashboard")

# Ajustar ancho de columnas
hoja_dashboard.column_dimensions["A"].width = 25
hoja_dashboard.column_dimensions["B"].width = 18

# Alinear los indicadores
for fila in range(3, 10):
    hoja_dashboard.cell(fila, 1).alignment = Alignment(
        horizontal="left"
    )

    hoja_dashboard.cell(fila, 2).alignment = Alignment(
        horizontal="center"
    )

print("Formato de indicadores aplicado")

# Crear carpeta de reportes si no existe
CARPETA_REPORTES.mkdir(parents=True, exist_ok=True)

# Definir ruta de salida
RUTA_DASHBOARD = CARPETA_REPORTES / "dashboard_ejecutivo.xlsx"

# Guardar archivo
libro_dashboard.save(RUTA_DASHBOARD)

print("\nDASHBOARD GUARDADO CORRECTAMENTE")
print("Archivo:", RUTA_DASHBOARD.name)

# Calcular unidades vendidas por producto
ventas_producto = {}

for fila in range(2, hoja_ventas.max_row + 1):

    producto = hoja_ventas.cell(fila, 2).value
    cantidad = hoja_ventas.cell(fila, 3).value

    if producto not in ventas_producto:
        ventas_producto[producto] = 0

    ventas_producto[producto] += cantidad

# Mostrar resultados
print("\nUNIDADES VENDIDAS POR PRODUCTO")
print("----------------------------------------")

for producto, unidades in ventas_producto.items():
    print(producto, ":", unidades)
    
# Preparar datos para el gráfico de productos
hoja_dashboard["C3"] = "Producto"
hoja_dashboard["D3"] = "Unidades"

# Escribir datos
for fila, (producto, unidades) in enumerate(
    ventas_producto.items(), start=4
):
    hoja_dashboard.cell(fila, 3, producto)
    hoja_dashboard.cell(fila, 4, unidades)

# Formatear encabezados
for columna in range(3, 5):
    celda = hoja_dashboard.cell(3, columna)

    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )

# Ajustar columnas
hoja_dashboard.column_dimensions["C"].width = 18
hoja_dashboard.column_dimensions["D"].width = 15

print("\nDATOS PARA GRAFICO DE PRODUCTOS INCORPORADOS")

# Crear gráfico de barras
grafico_productos = BarChart()

grafico_productos.type = "bar"
grafico_productos.style = 10
grafico_productos.title = "Unidades vendidas por producto"
grafico_productos.y_axis.title = "Producto"
grafico_productos.x_axis.title = "Unidades"

# Seleccionar datos
datos = Reference(
    hoja_dashboard,
    min_col=4,
    min_row=3,
    max_row=7
)

categorias = Reference(
    hoja_dashboard,
    min_col=3,
    min_row=4,
    max_row=7
)

# Incorporar datos y categorías
grafico_productos.add_data(datos, titles_from_data=True)
grafico_productos.set_categories(categorias)

# Definir tamaño
grafico_productos.width = 12
grafico_productos.height = 7

# Ubicar gráfico
hoja_dashboard.add_chart(grafico_productos, "C9")

print("Grafico de unidades vendidas creado")

# Calcular ingresos y costos por producto
resultados_producto = {}

for fila in range(2, hoja_ventas.max_row + 1):

    producto = hoja_ventas.cell(fila, 2).value
    cantidad = hoja_ventas.cell(fila, 3).value
    precio = hoja_ventas.cell(fila, 4).value

    # Buscar costo unitario
    costo_unitario = 0

    for fila_costo in range(2, 6):
        nombre = hoja_costos.cell(fila_costo, 1).value

        if nombre == producto:
            costo_unitario = hoja_costos.cell(fila_costo, 2).value
            break

    if producto not in resultados_producto:
        resultados_producto[producto] = {
            "ingresos": 0,
            "costos": 0
        }

    resultados_producto[producto]["ingresos"] += cantidad * precio
    resultados_producto[producto]["costos"] += cantidad * costo_unitario

print("\nINGRESOS Y COSTOS POR PRODUCTO")
print("----------------------------------------")

for producto, datos in resultados_producto.items():
    print(producto)
    print("Ingresos: $", datos["ingresos"])
    print("Costos: $", datos["costos"])

# Preparar datos para el gráfico de ingresos y costos
hoja_dashboard["E3"] = "Producto"
hoja_dashboard["F3"] = "Ingresos"
hoja_dashboard["G3"] = "Costos"

# Escribir resultados por producto
for fila, (producto, datos) in enumerate(
    resultados_producto.items(), start=4
):
    hoja_dashboard.cell(fila, 5, producto)
    hoja_dashboard.cell(fila, 6, datos["ingresos"])
    hoja_dashboard.cell(fila, 7, datos["costos"])

# Formatear encabezados
for columna in range(5, 8):
    celda = hoja_dashboard.cell(3, columna)

    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )

# Ajustar columnas
hoja_dashboard.column_dimensions["E"].width = 18
hoja_dashboard.column_dimensions["F"].width = 15
hoja_dashboard.column_dimensions["G"].width = 15

# Formato monetario
for fila in range(4, 8):
    for columna in range(6, 8):
        hoja_dashboard.cell(fila, columna).number_format = '$#,##0.00'

print("Datos de ingresos y costos incorporados")

# Crear gráfico de ingresos y costos
grafico_rentabilidad = BarChart()

grafico_rentabilidad.type = "col"
grafico_rentabilidad.style = 10
grafico_rentabilidad.title = "Ingresos y costos por producto"
grafico_rentabilidad.y_axis.title = "Importe ($)"
grafico_rentabilidad.x_axis.title = "Producto"

# Seleccionar datos de ingresos y costos
datos = Reference(
    hoja_dashboard,
    min_col=6,
    max_col=7,
    min_row=3,
    max_row=7
)

categorias = Reference(
    hoja_dashboard,
    min_col=5,
    min_row=4,
    max_row=7
)

# Incorporar datos y categorías
grafico_rentabilidad.add_data(datos, titles_from_data=True)
grafico_rentabilidad.set_categories(categorias)

# Definir tamaño
grafico_rentabilidad.width = 14
grafico_rentabilidad.height = 8

# Ubicar gráfico
hoja_dashboard.add_chart(grafico_rentabilidad, "J9")

print("Grafico de ingresos y costos creado")

# Guardar dashboard actualizado
libro_dashboard.save(RUTA_DASHBOARD)

print("\nDASHBOARD ACTUALIZADO CORRECTAMENTE")
print("Archivo:", RUTA_DASHBOARD.name)
print("Graficos incorporados: 2")

# Resumen de calidad de datos
hoja_dashboard["A12"] = "CALIDAD DE DATOS"
hoja_dashboard.merge_cells("A12:B12")

# Formato del encabezado
hoja_dashboard["A12"].font = Font(
    bold=True,
    color="FFFFFF",
    size=10
)

hoja_dashboard["A12"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)

hoja_dashboard["A12"].alignment = Alignment(
    horizontal="center"
)

print("Seccion de calidad de datos creada")

# Calcular calidad de los datos
ventas_validas = 0
ventas_invalidas = 0

for fila in range(2, hoja_ventas.max_row + 1):

    producto = hoja_ventas.cell(fila, 2).value
    cantidad = hoja_ventas.cell(fila, 3).value
    precio = hoja_ventas.cell(fila, 4).value

    if (
        producto is not None
        and cantidad is not None
        and precio is not None
        and cantidad > 0
        and precio > 0
    ):
        ventas_validas += 1
    else:
        ventas_invalidas += 1

total_registros = ventas_validas + ventas_invalidas
porcentaje_calidad = ventas_validas / total_registros

# Escribir indicadores
hoja_dashboard["A13"] = "Registros validos"
hoja_dashboard["B13"] = ventas_validas

hoja_dashboard["A14"] = "Registros invalidos"
hoja_dashboard["B14"] = ventas_invalidas

hoja_dashboard["A15"] = "Calidad de datos"
hoja_dashboard["B15"] = porcentaje_calidad

# Formato de porcentaje
hoja_dashboard["B15"].number_format = "0.00%"

print("\nCALIDAD DE DATOS")
print("----------------------------------------")
print("Registros validos:", ventas_validas)
print("Registros invalidos:", ventas_invalidas)
print("Calidad:", "{:.2f}%".format(porcentaje_calidad * 100))

# Formatear indicadores de calidad
for fila in range(13, 16):

    hoja_dashboard.cell(fila, 1).font = Font(
        bold=True,
        color="1F1F1F"
    )

    hoja_dashboard.cell(fila, 2).font = Font(
        bold=True,
        color="1F1F1F"
    )

    hoja_dashboard.cell(fila, 1).alignment = Alignment(
        horizontal="left"
    )

    hoja_dashboard.cell(fila, 2).alignment = Alignment(
        horizontal="center"
    )

# Aplicar bordes
from openpyxl.styles import Border, Side

borde = Side(style="thin", color="BFBFBF")

for fila in range(12, 16):
    for columna in range(1, 3):
        hoja_dashboard.cell(fila, columna).border = Border(
            left=borde,
            right=borde,
            top=borde,
            bottom=borde
        )

print("Formato de calidad de datos aplicado")

# Guardar dashboard ejecutivo final
libro_dashboard.save(RUTA_DASHBOARD)

print("\n========================================")
print("DASHBOARD EJECUTIVO FINALIZADO")
print("========================================")
print("Archivo:", RUTA_DASHBOARD.name)
print("Indicadores: 6")
print("Graficos: 2")
print("Seccion de calidad: OK")
print("Calidad de datos: {:.2f}%".format(porcentaje_calidad * 100))

