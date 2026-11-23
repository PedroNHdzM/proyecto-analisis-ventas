from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from datetime import date, datetime
from openpyxl.chart import PieChart, Reference
#PieChart: crea un gráfico circular.
#Reference: indica qué celdas contienen los datos que utilizará el gráfico.

# Ubicacion del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent

# Archivos de entrada
ruta_ventas = CARPETA_PROYECTO / "datos" / "ventas.xlsx"
ruta_costos = CARPETA_PROYECTO / "datos" / "costos_productos.xlsx"

# Archivo de salida
ruta_reporte = CARPETA_PROYECTO / "reportes" / "reporte_calidad_datos.xlsx"

print("REPORTE DE CALIDAD DE DATOS")
print("-" * 40)
print("Archivo de ventas:", ruta_ventas.name)
print("Archivo de costos:", ruta_costos.name)
print("Reporte de salida:", ruta_reporte.name)

# Cargar archivo de ventas
libro_ventas = load_workbook(ruta_ventas, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

# Cargar catalogo de productos
libro_costos = load_workbook(ruta_costos, data_only=True)
hoja_costos = libro_costos["Costos"]

# Crear conjunto de productos registrados
productos_catalogo = set()

for producto, costo in hoja_costos.iter_rows(
    min_row=2, max_col=2, values_only=True
):
    if isinstance(producto, str) and producto.strip():
        productos_catalogo.add(producto.strip().casefold())

# Contadores de calidad
ventas_validas = 0
ventas_invalidas = 0
errores_fecha = 0
errores_producto = 0
errores_cantidad = 0
errores_precio = 0 

# Revisar cada venta
for fecha, producto, cantidad, precio in hoja_ventas.iter_rows(
    min_row=2, max_col=4, values_only=True
):
    # Ignorar filas completamente vacias
    if all(valor is None for valor in (fecha, producto, cantidad, precio)):
        continue

    errores = []

    # Validar fecha
    fecha_valida = False

    if isinstance(fecha, (datetime, date)):
        fecha_valida = True

    elif isinstance(fecha, str):
        try:
            datetime.strptime(fecha.strip(), "%Y-%m-%d")
            fecha_valida = True
        except ValueError:
            pass

    if not fecha_valida:
        errores_fecha += 1
        errores.append("Fecha invalida")

    # Validar producto
    if not isinstance(producto, str) or not producto.strip():
        errores_producto += 1
        errores.append("Producto faltante")
    elif producto.strip().casefold() not in productos_catalogo:
        errores_producto += 1
        errores.append("Producto no registrado")

    # Validar cantidad
    if (
        not isinstance(cantidad, (int, float))
        or isinstance(cantidad, bool)
        or cantidad <= 0
        or not isinstance(cantidad, int)
    ):
        errores_cantidad += 1
        errores.append("Cantidad invalida")

    # Validar precio
    if (
        not isinstance(precio, (int, float))
        or isinstance(precio, bool)
        or precio < 0
    ):
        errores_precio += 1
        errores.append("Precio invalido")

    # Clasificar venta
    if errores:
        ventas_invalidas += 1
    else:
        ventas_validas += 1
        
     # Calcular total y porcentaje de calidad
    total_ventas = ventas_validas + ventas_invalidas
    if total_ventas > 0:
        porcentaje_validas = ventas_validas / total_ventas * 100
        porcentaje_invalidas = ventas_invalidas / total_ventas * 100
    else:
        porcentaje_validas = 0
        porcentaje_invalidas = 0
      
   
    
    # Clasificar calidad de datos
    if porcentaje_validas == 100:
        clasificacion = "Excelente"
    elif porcentaje_validas >= 95:
        clasificacion = "Buena"
    elif porcentaje_validas >= 80:
        clasificacion = "Regular"
    else:
        clasificacion = "Deficiente"
    # Mostrar resultados
print()
print("RESULTADOS DE CALIDAD")
print("-" * 40)

print(f"Total de ventas: {total_ventas}")
print(f"Ventas validas: {ventas_validas}")
print(f"Ventas invalidas: {ventas_invalidas}")
print(f"Porcentaje valido: {porcentaje_validas:.2f}%")
print(f"Porcentaje invalido: {porcentaje_invalidas:.2f}%")
print(f"Clasificacion de calidad: {clasificacion}")

print()
print("ERRORES DETECTADOS")
print("-" * 40)

print(f"Errores de fecha: {errores_fecha}")
print(f"Errores de producto: {errores_producto}")
print(f"Errores de cantidad: {errores_cantidad}")
print(f"Errores de precio: {errores_precio}")
        


print()
print("ARCHIVOS LEIDOS CORRECTAMENTE")
print("-" * 40)
print(f"Registros de ventas: {hoja_ventas.max_row - 1}")
print(f"Productos en catalogo: {len(productos_catalogo)}")


# Crear libro de Excel
libro_reporte = Workbook()
hoja_reporte = libro_reporte.active
hoja_reporte.title = "Calidad de datos"

# Encabezado
hoja_reporte["A1"] = "REPORTE DE CALIDAD DE DATOS"

# Indicadores generales
hoja_reporte["A3"] = "Indicador"
hoja_reporte["B3"] = "Resultado"

hoja_reporte["A4"] = "Total de ventas"
hoja_reporte["B4"] = total_ventas

hoja_reporte["A5"] = "Ventas validas"
hoja_reporte["B5"] = ventas_validas

hoja_reporte["A6"] = "Ventas invalidas"
hoja_reporte["B6"] = ventas_invalidas

hoja_reporte["A7"] = "Porcentaje valido"
hoja_reporte["B7"] = porcentaje_validas / 100

hoja_reporte["A8"] = "Porcentaje invalido"
hoja_reporte["B8"] = porcentaje_invalidas / 100

hoja_reporte["A9"] = "Clasificacion de calidad"
hoja_reporte["B9"] = clasificacion

# Detalle de errores
hoja_reporte["A10"] = "DETALLE DE ERRORES"
hoja_reporte["A11"] = "Tipo de error"
hoja_reporte["B11"] = "Cantidad"

hoja_reporte["A12"] = "Fechas invalidas"
hoja_reporte["B12"] = errores_fecha

hoja_reporte["A13"] = "Productos incorrectos"
hoja_reporte["B13"] = errores_producto

hoja_reporte["A14"] = "Cantidades invalidas"
hoja_reporte["B14"] = errores_cantidad

hoja_reporte["A15"] = "Precios invalidos"
hoja_reporte["B15"] = errores_precio

# Formato del titulo
hoja_reporte.merge_cells("A1:B1")
hoja_reporte["A1"].font = Font(
    bold=True,
    size=14,
    color="FFFFFF"
)
hoja_reporte["A1"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)
hoja_reporte["A1"].alignment = Alignment(
    horizontal="center"
)

# Formato de encabezados
for celda in hoja_reporte[3]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )
    celda.alignment = Alignment(horizontal="center")

# Formato del titulo de detalle
hoja_reporte.merge_cells("A10:B10")

hoja_reporte["A10"].font = Font(
    bold=True,
    color="FFFFFF"
)
hoja_reporte["A10"].fill = PatternFill(
    fill_type="solid",
    fgColor="1F4E78"
)
hoja_reporte["A10"].alignment = Alignment(
    horizontal="center"
)

# Formato de encabezados del detalle
for celda in hoja_reporte[11]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="4472C4"
    )
    celda.alignment = Alignment(horizontal="center")

# Formato de porcentajes
hoja_reporte["B7"].number_format = "0.00%"
hoja_reporte["B8"].number_format = "0.00%"

# Bordes de la tabla
borde = Border(
    bottom=Side(style="thin", color="BFBFBF")
)

for fila in hoja_reporte.iter_rows(min_row=3, max_row=8, min_col=1, max_col=15):
    for celda in fila:
        celda.border = borde

# Ancho de columnas
hoja_reporte.column_dimensions["A"].width = 28
hoja_reporte.column_dimensions["B"].width = 18

# Alinear resultados
for fila in range(4, 10):
    hoja_reporte[f"B{fila}"].alignment = Alignment(
        horizontal="center"
    )

# Crear grafico circular
grafico = PieChart()

# Seleccionar los datos
datos = Reference(
    hoja_reporte,
    min_col=2,
    min_row=5,
    max_row=6
)

# Seleccionar las etiquetas
etiquetas = Reference(
    hoja_reporte,
    min_col=1,
    min_row=5,
    max_row=6
)

grafico.add_data(datos)
grafico.set_categories(etiquetas)

# Titulo y posicion
grafico.title = "Calidad de las ventas"
grafico.height = 8
grafico.width = 12

# Insertar grafico
hoja_reporte.add_chart(grafico, "D3")

# Guardar archivo
libro_reporte.save(ruta_reporte)
libro_reporte.close()

print()
print("REPORTE EXCEL CREADO CORRECTAMENTE")
print(ruta_reporte)

libro_ventas.close()
libro_costos.close()

