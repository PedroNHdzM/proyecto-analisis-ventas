#Validar datos: detectar cantidades negativas, precios incorrectos o datos incompletos.
# isinstance() comprueba el tipo de dato.
# errores guarda los problemas encontrados en una fila.
# append() agrega un mensaje a esa lista.
# enumerate() nos permite identificar el número de fila de Excel.
# ventas_validas y ventas_invalidas cuentan los resultados.
from pathlib import Path
from openpyxl import load_workbook
from datetime import date, datetime

CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
ruta_ventas = CARPETA_PROYECTO / "datos" / "ventas.xlsx"

libro = load_workbook(ruta_ventas, data_only=True)
hoja = libro["Ventas"]
# print("Archivo:", ruta_ventas)
# print("Hoja:", hoja.title)
# print("Filas detectadas:", hoja.max_row)
# print("Columnas detectadas:", hoja.max_column)
# Cargar catalogo de productos
ruta_costos = CARPETA_PROYECTO / "datos" / "costos_productos.xlsx"

libro_costos = load_workbook(ruta_costos, data_only=True)
hoja_costos = libro_costos["Costos"]

productos_catalogo = set()

for producto_catalogo, costo in hoja_costos.iter_rows(
    min_row=2, max_col=2, values_only=True
):
    if isinstance(producto_catalogo, str) and producto_catalogo.strip():
        productos_catalogo.add(producto_catalogo.strip().casefold())

ventas_validas = 0
ventas_invalidas = 0

print("VALIDACION DE VENTAS")
print("-" * 40)

# print("PRIMERAS FILAS DEL EXCEL:")
# for fila in hoja.iter_rows(min_row=2, max_col=4, values_only=True):
#     print(fila)
    
for numero_fila, (fecha, producto, cantidad, precio) in enumerate(
    hoja.iter_rows(min_row=2, max_col=4, values_only=True),
    start=2
):
    # Ignorar filas completamente vacias
    if all(valor is None for valor in (fecha, producto, cantidad, precio)):
        continue

    errores = []

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
        errores.append("La fecha es invalida o esta vacia")

    if not isinstance(producto, str) or not producto.strip():
        errores.append("Falta el producto")
        
    elif producto.strip().casefold() not in productos_catalogo:
        errores.append(f"Producto no registrado en el catalogo: {producto}")

    if not isinstance(cantidad, (int, float)) or isinstance(cantidad, bool) or cantidad <= 0 or not isinstance(cantidad, int):
        errores.append("La cantidad debe ser un entero positivo")

    if not isinstance(precio, (int, float)) or isinstance(precio, bool) or precio < 0:
        errores.append("El precio debe ser un numero mayor o igual a cero")

    if errores:
        ventas_invalidas += 1
        print(f"Fila {numero_fila}: VENTA INVALIDA")
        for error in errores:
            print(f"  - {error}")
    else:
        ventas_validas += 1

print("-" * 40)
print(f"Ventas validas: {ventas_validas}")
print(f"Ventas invalidas: {ventas_invalidas}")

total_ventas = ventas_validas + ventas_invalidas

if total_ventas > 0:
    porcentaje_validas = (ventas_validas / total_ventas) * 100
else:
    porcentaje_validas = 0

print(f"Porcentaje de ventas validas: {porcentaje_validas:.2f}%")

libro.close()
libro_costos.close()

print()
# print("PRUEBA DE VALIDACION")
# cantidad_prueba = -3

# if cantidad_prueba <= 0:
#     print(f"Cantidad invalida: {cantidad_prueba}")
# else:
#     print(f"Cantidad valida: {cantidad_prueba}")