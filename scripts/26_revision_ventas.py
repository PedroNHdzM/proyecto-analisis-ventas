#revisar los productos de ventas.xlsx
# ¿Qué hace este programa?
# 1. Localiza el archivo ventas.xlsx.
# 2. Abre la hoja Ventas.
# 3. Lee los nombres de los productos, sin modificar nada.
# 4. Muestra los resultados en PowerShell.

from pathlib import Path
from openpyxl import load_workbook

CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
ruta_ventas = CARPETA_PROYECTO / "datos" / "ventas.xlsx"

libro = load_workbook(ruta_ventas, data_only=True)
hoja = libro["Ventas"]

print("REVISION DE PRODUCTOS")
print("------------------------------")
print(" ")
print("Saber si alguno de los 11 nombres necesita limpieza.")
for fila in range(2, hoja.max_row + 1):
    #¿Qué buscamos? Saber si alguno de los 11 nombres necesita limpieza,
    # sin modificar todavía el Excel original.
    producto = hoja.cell(fila, 2).value
    producto_limpio = producto.strip()

    if producto != producto_limpio:
        print(f"Espacios detectados: [{producto}]")
    else:
        print(f"Correcto: {producto}")


print(" ")
print("alguna venta tiene el nombre del producto vacío.")
for fila in range(2, hoja.max_row + 1):
    #vamos a revisar si alguna venta tiene el nombre del producto vacío.
    #Objetivo: comprobar si las 11 ventas tienen un producto registrado
    producto = hoja.cell(fila, 2).value

    if producto is None or producto.strip() == "":
        print(f"Producto vacio en fila: {fila}")
    else:
        print(f"Producto correcto en fila: {fila}")

print(" ")
print("detectar cantidades vacías.")
# Siguiente ejercicio: detectar cantidades vacías
# Ahora revisaremos otra columna importante: Cantidad.
for fila in range(2, hoja.max_row + 1):
    cantidad = hoja.cell(fila, 3).value

    if cantidad is None:
        print(f"Cantidad vacia en fila: {fila}")
    else:
        print(f"Cantidad correcta en fila: {fila}")

print(" ")
print("validar la columna Cantidad numeros positivos.")
#Que una cantidad no esté vacía no significa necesariamente que sea correcta. 
# Por ejemplo, una cantidad de -3 no tendría sentido en nuestro registro de ventas.
for fila in range(2, hoja.max_row + 1):
    #Objetivo: comprobar que todas las cantidades sean mayores que cero
    cantidad = hoja.cell(fila, 3).value

    if cantidad is None:
        print(f"Cantidad vacia en fila: {fila}")

    elif cantidad <= 0:
        print(f"Cantidad incorrecta en fila: {fila}: {cantidad}")

    else:
        print(f"Cantidad valida en fila: {fila}: {cantidad}")

#Siguiente ejercicio: revisar precios
#Ahora vamos a validar la columna Precio.
print(" ")
print("validar la columna Precio.")
for fila in range(2, hoja.max_row + 1):
    precio = hoja.cell(fila, 4).value

    if precio is None:
        print(f"Precio vacio en fila: {fila}")

    elif precio <= 0:
        print(f"Precio incorrecto en fila: {fila}: {precio}")

    else:
        print(f"Precio valido en fila: {fila}: {precio}")

libro.close()


print(" ")
print("detectar ventas duplicadas.")
# ¿Qué estamos aprendiendo?
# - Las tuplas permiten agrupar los cuatro datos de una venta.
# - in comprueba si esa venta ya apareció.
# - append() guarda las ventas que ya revisamos.
# Importante: el programa solamente detecta duplicados; no elimina ningún registro.
ventas_vistas = []

for fila in range(2, hoja.max_row + 1):

    fecha = hoja.cell(fila, 1).value
    producto = hoja.cell(fila, 2).value
    cantidad = hoja.cell(fila, 3).value
    precio = hoja.cell(fila, 4).value

    venta = (fecha, producto, cantidad, precio)

    if venta in ventas_vistas:
        print(f"Venta duplicada en fila: {fila}")
    else:
        ventas_vistas.append(venta)

print("--------------------------------")
print("Total de ventas revisadas:", hoja.max_row - 1)
print("Ventas unicas:", len(ventas_vistas))

