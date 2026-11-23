# ¿Qué estamos haciendo?
# - Path(__file__) identifica dónde está guardado nuestro programa.
# - parent.parent sube dos niveles hasta la carpeta Proyecto_Final.
# - Construimos las rutas de los archivos Excel sin escribir rutas absolutas de Windows.
# Paso 1: preparar el programa
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# Identificar las carpetas del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent

# Identificar los archivos de entrada
RUTA_VENTAS = CARPETA_PROYECTO / "datos" / "ventas.xlsx"
RUTA_COSTOS = CARPETA_PROYECTO / "datos" / "costos_productos.xlsx"

# Identificar la carpeta de reportes
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"

print("REPORTE GENERAL DEL PROYECTO")
print("----------------------------------------")

print("Archivo de ventas:", RUTA_VENTAS.name)
print("Archivo de costos:", RUTA_COSTOS.name)
print("Carpeta de reportes:", CARPETA_REPORTES.name)

#Paso 2: leer los archivos Excel
# Conceptos nuevos o que estamos reforzando:
# - load_workbook() abre un archivo Excel.
# - data_only=True permite leer los resultados guardados de las fórmulas.
# - .title obtiene el nombre de la hoja.
# - .max_row devuelve la última fila utilizada.
# Abrir archivo de ventas
libro_ventas = load_workbook(RUTA_VENTAS, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

# Abrir archivo de costos
libro_costos = load_workbook(RUTA_COSTOS, data_only=True)
hoja_costos = libro_costos["Costos"]

print("\nARCHIVOS ABIERTOS CORRECTAMENTE")
print("----------------------------------------")

print("Hoja de ventas:", hoja_ventas.title)
print("Hoja de costos:", hoja_costos.title)

print("Registros de ventas:", hoja_ventas.max_row - 1)

#Esta versión cuenta solamente las celdas que contienen
# nombres de productos en la columna A, excluyendo el encabezado.
print(
    "Productos en catalogo:",
    sum(
        1
        for fila in hoja_costos.iter_rows(
            min_row=2,
            min_col=1,
            max_col=1,
            values_only=True
        )
        if fila[0] is not None
    )
)

#Paso 3: calcular las ventas
# Inicializar contadores
total_ventas = 0
total_unidades = 0

# Recorrer las ventas
for fila in hoja_ventas.iter_rows(min_row=2, values_only=True):
    #iter_rows() recorre las filas desde la segunda, porque la primera contiene los encabezados.
    fecha, producto, cantidad, precio = fila

    total_ventas += 1
    total_unidades += cantidad

print("\nINDICADORES DE VENTAS")
print("----------------------------------------")
print("Total de ventas:", total_ventas)
print("Total de unidades:", total_unidades)

#Paso 4: calcular ingresos totales

# Inicializar ingresos
total_ingresos = 0

# Recorrer las ventas
for fila in hoja_ventas.iter_rows(min_row=2, values_only=True):

    fecha, producto, cantidad, precio = fila

    importe = cantidad * precio
    total_ingresos += importe

print("\nINGRESOS POR VENTAS")
print("----------------------------------------")
print(f"Total de ingresos: ${total_ingresos:,.2f}")



#Paso 5: crear el catálogo de costos
# Crear diccionario de costos
#Una explicación importante: min_col=1 y max_col=2 
# limitan la lectura a las columnas A y B. Así evitamos las columnas del ejercicio.
costos_productos = {}

for fila in hoja_costos.iter_rows(
    min_row=2,
    max_row=hoja_costos.max_row,
    min_col=1,
    max_col=2,
    values_only=True
):

    producto, costo = fila

    if producto is not None and costo is not None:
        costos_productos[producto] = costo

print("\nCATALOGO DE COSTOS")
print("----------------------------------------")

for producto, costo in costos_productos.items():
    print(f"{producto}: ${costo:,.2f}")

print("Total de productos:", len(costos_productos))
# # print("\nESTRUCTURA DEL CATALOGO")
# # print("----------------------------------------")

# # for fila in hoja_costos.iter_rows(values_only=True):
# #     print(fila)
    
print("\nCATALOGO DE COSTOS")
print("----------------------------------------")

for producto, costo in costos_productos.items():
    print(f"{producto}: ${costo:,.2f}")

print("Total de productos en diccionario:", len(costos_productos))

#Paso 6: calcular el costo total de las ventas
# Calcular costo total de las ventas
total_costos = 0

for fila in hoja_ventas.iter_rows(min_row=2, values_only=True):

    fecha, producto, cantidad, precio = fila

    costo_unitario = costos_productos[producto]

    costo_venta = cantidad * costo_unitario

    total_costos += costo_venta

print("\nCOSTOS DE VENTAS")
print("----------------------------------------")
print(f"Total de costos: ${total_costos:,.2f}")

#Paso 7: calcular la utilidad
# Dos conceptos importantes:
# - Utilidad = ingresos − costos.
# - Margen = (utilidad ÷ ingresos) × 100.
# Calcular utilidad
utilidad = total_ingresos - total_costos

# Calcular margen de utilidad
margen_utilidad = (utilidad / total_ingresos) * 100

print("\nRENTABILIDAD DEL PROYECTO")
print("----------------------------------------")
print(f"Utilidad total: ${utilidad:,.2f}")
print(f"Margen de utilidad: {margen_utilidad:.2f}%")

#Paso 8: calcular la calidad de los datos

# Calcular calidad de los datos
ventas_correctas = 0
ventas_con_problemas = 0

for fila in hoja_ventas.iter_rows(min_row=2, values_only=True):

    fecha, producto, cantidad, precio = fila

    if (
        fecha is not None
        and producto is not None
        and producto in costos_productos
        and cantidad is not None
        and cantidad > 0
        and precio is not None
        and precio > 0
    ):
        ventas_correctas += 1
    else:
        ventas_con_problemas += 1

porcentaje_calidad = (
    ventas_correctas / total_ventas
) * 100

print("\nCALIDAD DE LOS DATOS")
print("----------------------------------------")
print("Ventas correctas:", ventas_correctas)
print("Ventas con problemas:", ventas_con_problemas)
print(f"Calidad de los datos: {porcentaje_calidad:.2f}%")

#Paso 9 — Preparar el archivo Excel

# Crear libro de Excel
libro_reporte = Workbook()

# Seleccionar hoja activa
hoja_reporte = libro_reporte.active
hoja_reporte.title = "Resumen General"

# Ocultar lineas de cuadricula
hoja_reporte.sheet_view.showGridLines = False

# Crear titulo
hoja_reporte["A1"] = "REPORTE GENERAL DEL PROYECTO DE VENTAS"

# Crear secciones
hoja_reporte["A3"] = "INDICADORES DE VENTAS"
hoja_reporte["A4"] = "Total de ventas"
hoja_reporte["A5"] = "Unidades vendidas"
hoja_reporte["A6"] = "Ingresos"

hoja_reporte["A8"] = "INDICADORES DE RENTABILIDAD"
hoja_reporte["A9"] = "Costos"
hoja_reporte["A10"] = "Utilidad"
hoja_reporte["A11"] = "Margen de utilidad"

hoja_reporte["A13"] = "CALIDAD DE LOS DATOS"
hoja_reporte["A14"] = "Ventas correctas"
hoja_reporte["A15"] = "Ventas con problemas"
hoja_reporte["A16"] = "Porcentaje de calidad"

print("\nEstructura del reporte creada correctamente.")

#Paso 10: colocar los resultados en Excel

# Colocar resultados de ventas
hoja_reporte["B4"] = total_ventas
hoja_reporte["B5"] = total_unidades
hoja_reporte["B6"] = total_ingresos

# Colocar resultados de rentabilidad
hoja_reporte["B9"] = total_costos
hoja_reporte["B10"] = utilidad
hoja_reporte["B11"] = margen_utilidad / 100

# Colocar resultados de calidad
hoja_reporte["B14"] = ventas_correctas
hoja_reporte["B15"] = ventas_con_problemas
hoja_reporte["B16"] = porcentaje_calidad / 100

# Formatos numéricos
for celda in ["B6", "B9", "B10"]:
    hoja_reporte[celda].number_format = '$#,##0.00'

hoja_reporte["B11"].number_format = "0.00%"
hoja_reporte["B16"].number_format = "0.00%"

print("Indicadores transferidos al reporte Excel.")


#Paso 11: aplicar formato profesional
# Formato del titulo principal
hoja_reporte.merge_cells("A1:B1")

hoja_reporte["A1"].font = Font(
    name="Arial",
    size=10,
    bold=True,
    color="FFFFFF"
)

hoja_reporte["A1"].fill = PatternFill(
    "solid",
    fgColor="1F4E78"
)

hoja_reporte["A1"].alignment = Alignment(
    horizontal="center",
    vertical="center"
)

hoja_reporte.row_dimensions[1].height = 30

# Formato de las secciones
for fila in [3, 8, 13]:

    celda = hoja_reporte[f"A{fila}"]

    celda.font = Font(
        bold=True,
        color="FFFFFF"
    )

    celda.fill = PatternFill(
        "solid",
        fgColor="4472C4"
    )

    hoja_reporte.merge_cells(
        start_row=fila,
        start_column=1,
        end_row=fila,
        end_column=2
    )

# Formato de las etiquetas
for fila in [4, 5, 6, 9, 10, 11, 14, 15, 16]:

    hoja_reporte[f"A{fila}"].font = Font(
        bold=True,
        color="404040"
    )

# Ajustar anchos
hoja_reporte.column_dimensions["A"].width = 32
hoja_reporte.column_dimensions["B"].width = 20

print("Formato visual aplicado correctamente.")

#Paso 12: agregar bordes y alineación

# Crear borde delgado
borde_delgado = Side(
    style="thin",
    color="D9E1F2"
)

borde = Border(
    left=borde_delgado,
    right=borde_delgado,
    top=borde_delgado,
    bottom=borde_delgado
)

# Aplicar bordes y alineacion
for fila in hoja_reporte.iter_rows(
    min_row=1,
    max_row=16,
    min_col=1,
    max_col=2
):
    for celda in fila:
        celda.border = borde

# Centrar los resultados numericos
for fila in [4, 5, 6, 9, 10, 11, 14, 15, 16]:

    hoja_reporte[f"B{fila}"].alignment = Alignment(
        horizontal="center",
        vertical="center"
    )

# Resaltar la utilidad
hoja_reporte["B10"].font = Font(
    bold=True,
    color="008000"
)

print("Bordes y alineacion aplicados correctamente.")

#Paso 13: guardar el archivo Excel
# Definir ruta del reporte
RUTA_REPORTE = (
    CARPETA_REPORTES / "reporte_general.xlsx"
)

# Guardar archivo Excel
libro_reporte.save(RUTA_REPORTE)

# Cerrar archivos originales
libro_ventas.close()
libro_costos.close()

print("\nREPORTE GENERAL CREADO CORRECTAMENTE")
print("----------------------------------------")
print("Archivo:", RUTA_REPORTE.name)
print("Ubicacion:", RUTA_REPORTE)



