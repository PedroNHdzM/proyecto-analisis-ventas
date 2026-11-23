
from openpyxl import load_workbook
from pathlib import Path

# Localizar el archivo Excel
carpeta_proyecto = Path(__file__).resolve().parent.parent
ruta_excel = carpeta_proyecto / "datos" / "ventas.xlsx"

# Abrir el archivo
libro = load_workbook(ruta_excel, data_only=True)
hoja = libro["Ventas"]

# Variables para acumular resultados
total_ventas = 0
total_unidades = 0

print("ANALISIS DE VENTAS")
print("-" * 75)

for fila in hoja.iter_rows(min_row=2, values_only=True):
    fecha, producto, cantidad, precio = fila

    # Calcular el importe de esta venta
    importe = cantidad * precio

    # Acumular los resultados
    total_ventas += importe
    total_unidades += cantidad

    print(
        f"{fecha} | "
        f"{producto:12} | "
        f"Cantidad: {cantidad:2} | "
        f"Importe: ${importe:,.2f}"
    )

print("-" * 75)
print(f"Unidades vendidas: {total_unidades}")
print(f"Importe total de ventas: ${total_ventas:,.2f}")

libro.close()
