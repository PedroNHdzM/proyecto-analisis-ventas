# Clase 31. Reporte de calidad de datos
# Objetivo: crear un reporte Excel que identifique problemas en los registros de ventas.
# ¿Qué estamos haciendo?
# - Path: localiza los archivos del proyecto.
# - load_workbook(): abre el Excel.
# - total_ventas: contará los registros.
# - Los demás contadores registrarán posibles problemas.

from pathlib import Path
from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

#Definir el estilo de los bordes
borde_fino = Side(style="thin", color="BFBFBF")

borde_celda = Border(
    left=borde_fino,
    right=borde_fino,
    top=borde_fino,
    bottom=borde_fino
)

# Rutas del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
RUTA_VENTAS = CARPETA_PROYECTO / "datos" / "ventas.xlsx"
RUTA_REPORTE = CARPETA_PROYECTO / "reportes" / "reporte_calidad_avanzado.xlsx"

# Abrir archivo Excel
libro = load_workbook(RUTA_VENTAS, data_only=True)
hoja = libro["Ventas"]

# Contadores
total_ventas = 0
productos_vacios = 0
cantidades_invalidas = 0
precios_invalidos = 0
ventas_duplicadas = 0

print("INICIANDO REVISION DE CALIDAD")
print("--------------------------------")
ventas_registradas = []
for fila in range(2, hoja.max_row + 1):

    total_ventas += 1
    producto = hoja.cell(fila, 2).value

    if producto is None or producto.strip() == "":
        productos_vacios += 1
        
    cantidad = hoja.cell(fila, 3).value

    if cantidad is None or cantidad <= 0:
        cantidades_invalidas += 1
    
    precio = hoja.cell(fila, 4).value
    venta = (
        hoja.cell(fila, 1).value,
        producto,
        cantidad,
        precio
    )

    if venta in ventas_registradas:
        ventas_duplicadas += 1
    else:
        ventas_registradas.append(venta)
        
    if precio is None or precio <= 0:
        precios_invalidos += 1

ventas_con_problemas = (
    productos_vacios
    + cantidades_invalidas
    + precios_invalidos
    + ventas_duplicadas
)

ventas_correctas = total_ventas - ventas_con_problemas

if total_ventas > 0:
    porcentaje_calidad = ventas_correctas / total_ventas * 100
else:
    porcentaje_calidad = 0

if porcentaje_calidad == 100:
    clasificacion = "Excelente"
elif porcentaje_calidad >= 90:
    clasificacion = "Bueno"
elif porcentaje_calidad >= 70:
    clasificacion = "Regular"
else:
    clasificacion = "Deficiente" 
    
resultados = [
    ["Indicador", "Resultado"],
    ["Ventas revisadas", total_ventas],
    ["Productos vacios", productos_vacios],
    ["Cantidades invalidas", cantidades_invalidas],
    ["Precios invalidos", precios_invalidos],
    ["Ventas duplicadas", ventas_duplicadas],
    ["Ventas con problemas", ventas_con_problemas],
    ["Ventas correctas", ventas_correctas],
    ["Porcentaje de calidad", porcentaje_calidad],
    ["Clasificacion", clasificacion]
]

print("Total de ventas revisadas:", total_ventas)
print("Productos vacios:", productos_vacios)
print("Cantidades invalidas:", cantidades_invalidas)
print("Precios invalidos:", precios_invalidos)
print("Ventas duplicadas:", ventas_duplicadas)

print("Ventas con problemas:", ventas_con_problemas)
print("Ventas correctas:", ventas_correctas)
print("Porcentaje de calidad:", round(porcentaje_calidad, 2), "%")
print("Clasificacion de calidad:", clasificacion)

libro_reporte = Workbook()
hoja_reporte = libro_reporte.active
hoja_reporte.title = "Calidad"
hoja_reporte.sheet_view.showGridLines = False #oculta las líneas de cuadrícula de Excel.

for fila in resultados:
    # Recorre cada elemento de la lista resultados.
    # append() agrega cada fila a la hoja de Excel.
    # La primera fila contendrá los encabezados: Indicador y Resultado.
    hoja_reporte.append(fila)
    
for celda in hoja_reporte[1]:
    # bold=True: coloca el texto en negritas.
    # color="FFFFFF": utiliza letras blancas.
    # PatternFill: aplica un fondo azul.
    # Alignment: centra los encabezados.
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        start_color="1F4E78",
        end_color="1F4E78",
        fill_type="solid"
    )
    celda.alignment = Alignment(horizontal="center") 
    
for fila in hoja_reporte.iter_rows():
    for celda in fila:
        celda.border = borde_celda

for fila in range(2, hoja_reporte.max_row + 1):
    hoja_reporte.cell(fila, 2).alignment = Alignment(horizontal="center")
    
  
hoja_reporte.column_dimensions["A"].width = 30
hoja_reporte.column_dimensions["B"].width = 20 
hoja_reporte["B9"].number_format = '0.00"%"'
#Ahora haremos que el color cambie automáticamente según la clasificación.
colores_clasificacion = {
    #Excelente: verde oscuro;Bueno: verde claro.
    #Regular: amarillo, Deficiente: rojo
    "Excelente": "548235",
    "Bueno": "70AD47",
    "Regular": "FFC000",
    "Deficiente": "C00000"
}

color_clasificacion = colores_clasificacion[clasificacion]
#Vamos a darle un color de fondo a la celda donde aparece la clasificación.
hoja_reporte["B10"].font = Font(bold=True, color="FFFFFF")
hoja_reporte["B10"].fill = PatternFill(
    start_color=color_clasificacion,
    end_color=color_clasificacion,
    fill_type="solid"
)
hoja_reporte["B10"].alignment = Alignment(horizontal="center")

libro_reporte.save(RUTA_REPORTE)

print("--------------------------------")
print("REPORTE EXCEL CREADO CORRECTAMENTE")
print("Archivo:", RUTA_REPORTE)
libro.close()