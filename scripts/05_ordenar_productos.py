
from openpyxl import load_workbook
from pathlib import Path

# Localizar el archivo Excel
carpeta_proyecto = Path(__file__).resolve().parent.parent
ruta_excel = carpeta_proyecto / "datos" / "ventas.xlsx"

libro = load_workbook(ruta_excel, data_only=True)
hoja = libro["Ventas"]

# Diccionario para acumular los resultados
resumen = {}

for fila in hoja.iter_rows(min_row=2, values_only=True):
    fecha, producto, cantidad, precio = fila

    importe = cantidad * precio

    # Crear el producto si todavía no existe
    if producto not in resumen:
        resumen[producto] = {
            "unidades": 0,
            "ingresos": 0
        }

    # Acumular unidades e ingresos
    resumen[producto]["unidades"] += cantidad
    resumen[producto]["ingresos"] += importe

# Ordenar por ingresos, de mayor a menor
productos_ordenados = sorted(
    resumen.items(),
    key=lambda elemento: elemento[1]["ingresos"],
    reverse=True
)

# Mostrar el resumen ordenado
print("VENTAS POR PRODUCTO - ORDEN DE INGRESOS")
print("-" * 55)

for producto, datos in productos_ordenados:
    print(
        f"{producto:12} | "
        f"Unidades: {datos['unidades']:2} | "
        f"Ingresos: ${datos['ingresos']:,.2f}"
    )
# sorted() ordena los elementos.
# key= indica qué dato se utiliza para ordenar.
# lambda define una pequeña función.
# elemento[1]["ingresos"] obtiene los ingresos de cada producto.
# reverse=True invierte el orden para mostrar primero el ingreso más alto.
libro.close()
