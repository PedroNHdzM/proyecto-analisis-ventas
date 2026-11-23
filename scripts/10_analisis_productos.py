
from openpyxl import load_workbook
from collections import defaultdict

# Leer el archivo Excel original
ruta_excel = "Proyecto_Final/datos/ventas.xlsx"

libro = load_workbook(ruta_excel, data_only=True)
hoja = libro["Ventas"]

# Acumular unidades e ingresos por producto
productos = defaultdict(lambda: {
    "unidades": 0,
    "ingresos": 0
})

for fecha, producto, cantidad, precio in hoja.iter_rows(
    min_row=2, values_only=True
):
    importe = cantidad * precio

    productos[producto]["unidades"] += cantidad
    productos[producto]["ingresos"] += importe

libro.close()

# Ordenar productos por unidades vendidas
productos_por_unidades = sorted(
    productos.items(),
    key=lambda elemento: elemento[1]["unidades"],
    reverse=True
)

# Ordenar productos por ingresos
productos_por_ingresos = sorted(
    productos.items(),
    key=lambda elemento: elemento[1]["ingresos"],
    reverse=True
)

# Mostrar resultados
print("\nPRODUCTOS POR UNIDADES VENDIDAS")
print("-" * 45)

for producto, datos in productos_por_unidades:
    print(
        f"{producto}: {datos['unidades']} unidades"
    )

print("\nPRODUCTOS POR INGRESOS")
print("-" * 45)

for producto, datos in productos_por_ingresos:
    print(
        f"{producto}: ${datos['ingresos']:,.2f}"
    )

# Identificar los líderes
mas_vendido = productos_por_unidades[0]
mayores_ingresos = productos_por_ingresos[0]

print("\nCONCLUSIONES")
print("-" * 45)
print(
    f"Producto más vendido: {mas_vendido[0]} "
    f"({mas_vendido[1]['unidades']} unidades)"
)
print(
    f"Producto con mayores ingresos: "
    f"{mayores_ingresos[0]} "
    f"(${mayores_ingresos[1]['ingresos']:,.2f})"
)
