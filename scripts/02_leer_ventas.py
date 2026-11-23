# load_workbook(): abre un archivo Excel existente, a diferencia de Workbook(), que crea uno nuevo.
# libro["Ventas"]: selecciona la hoja llamada Ventas.
# iter_rows(min_row=2, values_only=True): recorre las filas desde la segunda, evitando los encabezados.
# fecha, producto, cantidad, precio = fila: asigna cada valor de la fila a una variable.
# libro.close(): cierra el archivo después de leerlo.

from openpyxl import load_workbook
from pathlib import Path

# Localizar el archivo Excel
carpeta_proyecto = Path(__file__).resolve().parent.parent
ruta_excel = carpeta_proyecto / "datos" / "ventas.xlsx"

# Abrir el archivo existente
libro = load_workbook(ruta_excel, data_only=True)

# Seleccionar la hoja
hoja = libro["Ventas"]

# Mostrar los encabezados
print("ENCABEZADOS")
print("-" * 50)

for celda in hoja[1]:
    print(celda.value, end=" | ")

print("\n")

# Mostrar las ventas
print("REGISTRO DE VENTAS")
print("-" * 50)

for fila in hoja.iter_rows(min_row=2, values_only=True):
    fecha, producto, cantidad, precio = fila

    print(
        f"Fecha: {fecha} | "
        f"Producto: {producto} | "
        f"Cantidad: {cantidad} | "
        f"Precio: ${precio:.2f}"
    )

# Cerrar el libro
libro.close()
