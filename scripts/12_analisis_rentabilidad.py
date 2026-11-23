
from openpyxl import load_workbook

# Rutas de los archivos
ruta_ventas = "Proyecto_Final/datos/ventas.xlsx"
ruta_costos = "Proyecto_Final/datos/costos_productos.xlsx"

# Leer el archivo de ventas
libro_ventas = load_workbook(ruta_ventas, data_only=True)
hoja_ventas = libro_ventas["Ventas"]

# Leer el archivo de costos
libro_costos = load_workbook(ruta_costos, data_only=True)
hoja_costos = libro_costos["Costos"]

# Crear un diccionario con los costos unitarios
costos = {}

    
for producto, costo in hoja_costos.iter_rows(
    min_row=2,
    max_col=2,
    values_only=True
):
    if producto is not None and costo is not None:
        costos[producto] = costo    

# Preparar el diccionario para acumular resultados
rentabilidad = {}

for fecha, producto, cantidad, precio in hoja_ventas.iter_rows(
    min_row=2, values_only=True
):
    ingreso = cantidad * precio
    costo_unitario = costos[producto]
    costo_total = cantidad * costo_unitario
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

# Mostrar el reporte
print("\nREPORTE DE RENTABILIDAD")
print("-" * 75)

for producto, datos in rentabilidad.items():
    margen = datos["ganancia"] / datos["ingresos"] * 100

    print(f"\nProducto: {producto}")
    print(f"Unidades vendidas: {datos['unidades']}")
    print(f"Ingresos: ${datos['ingresos']:,.2f}")
    print(f"Costos: ${datos['costos']:,.2f}")
    print(f"Ganancia: ${datos['ganancia']:,.2f}")
    print(f"Margen de ganancia: {margen:.2f}%")

# Calcular los totales generales
total_ingresos = sum(
    datos["ingresos"] for datos in rentabilidad.values()
)
total_costos = sum(
    datos["costos"] for datos in rentabilidad.values()
)
total_ganancia = sum(
    datos["ganancia"] for datos in rentabilidad.values()
)
margen_general = total_ganancia / total_ingresos * 100

print("\n" + "=" * 75)
print("RESUMEN GENERAL")
print(f"Ingresos totales: ${total_ingresos:,.2f}")
print(f"Costos totales: ${total_costos:,.2f}")
print(f"Ganancia total: ${total_ganancia:,.2f}")
print(f"Margen general: {margen_general:.2f}%")

# Cerrar los archivos
libro_ventas.close()
libro_costos.close()
