
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

# Mostrar el resumen
print("VENTAS POR PRODUCTO")
print("-" * 55)

for producto, datos in resumen.items():
    print(
        f"{producto:12} | "
        f"Unidades: {datos['unidades']:2} | "
        f"Ingresos: ${datos['ingresos']:,.2f}"
    )

libro.close()
