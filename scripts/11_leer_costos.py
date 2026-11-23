
from openpyxl import load_workbook

ruta_costos = "Proyecto_Final/datos/costos_productos.xlsx"

libro = load_workbook(ruta_costos, data_only=True)
hoja = libro["Costos"]

print("\nCATALOGO DE COSTOS")
print("-" * 35)

for producto, costo in hoja.iter_rows(
    min_row=2, values_only=True
):
    print(f"{producto}: ${costo:,.2f}")

libro.close()
