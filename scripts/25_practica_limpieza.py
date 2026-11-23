# Una distinción importante:
# - None: ausencia de valor.
# - 0: valor numérico cero.
# - "": texto vacío.
#Ejercicio 1: detectar datos vacíos en una cadena
print("Ejercicio 1: detectar datos vacíos string")
producto = "   Teclado   "

print("Producto original:", producto)
print("Producto limpio:", producto.strip())

#Ejercicio 2: detectar datos vacíos
print(" ")
print("Ejercicio 2: detectar datos vacíos")

cantidad = None

if cantidad is None:
    print("La cantidad esta vacia")
else:
    print("Cantidad:", cantidad)
    
#Ejercicio 3: comprobar los tres casos
print(" ")
print("Ejercicio 3: comprobar los tres casos")
cantidad = 0

if cantidad is None:
    print("No hay dato")
else:
    print("La cantidad es:", cantidad)

producto = ""

if producto == "":
    print("El nombre del producto esta vacio")

#Ejercicio 4: limpiar nombres de productos
print(" ")
print("Ejercicio 4: limpiar nombres de productos")
productos = [
    "  Teclado  ",
    "Mouse",
    "  Monitor ",
    " Audifonos  "
]

productos_limpios = []

for producto in productos:
    producto_limpio = producto.strip()
    productos_limpios.append(producto_limpio)

print("Productos limpios:")

for producto in productos_limpios:
    print(producto)
    
#Ejercicio 5: detectar productos vacíos
print(" ")
print("Ejercicio 5: detectar productos vacíos")
productos = [
    "Teclado",
    "",
    "Mouse",
    "   ",
    "Monitor",
    ""
]

productos_vacios = 0
productos_validos = 0

for producto in productos:
    producto = producto.strip()

    if producto == "":
        productos_vacios += 1
    else:
        productos_validos += 1

print("Productos validos:", productos_validos)
print("Productos vacios:", productos_vacios)
print("Total de productos:", len(productos))

#Ejercicio 6: detectar duplicados
print(" ")
print("Ejercicio 6: detectar duplicados")
productos = [
    "Teclado",
    "Mouse",
    "Monitor",
    "Teclado",
    "Audifonos",
    "Monitor",
    "Mouse",
    "Teclado"
]
productos_vistos = []
#Observa especialmente cómo funciona in: comprueba si un elemento ya existe dentro de una lista.
for producto in productos:

    if producto in productos_vistos:
        print("Producto duplicado:", producto)

    else:
        productos_vistos.append(producto)
        
#Ejercicio 7: integrar limpieza y duplicados
print(" ")
print("Ejercicio 7: integrar limpieza y duplicados")
productos = [
    " Teclado ",
    "Mouse",
    "Teclado",
    "",
    " Monitor ",
    "Mouse",
    "   ",
    "Audifonos",
    "Teclado"
]

productos_vistos = []
productos_vacios = 0
productos_duplicados = 0

for producto in productos:

    producto = producto.strip()

    if producto == "":
        productos_vacios += 1

    elif producto in productos_vistos:
        productos_duplicados += 1

    else:
        productos_vistos.append(producto)

print("RESUMEN DE LIMPIEZA")
print("------------------------------")
print("Productos unicos:", len(productos_vistos))
print("Productos vacios:", productos_vacios)
print("Registros duplicados:", productos_duplicados)
print("Total de registros:", len(productos))