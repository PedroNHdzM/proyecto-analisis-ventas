
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from pathlib import Path

# Ubicacion del proyecto
carpeta_proyecto = Path(__file__).resolve().parent.parent
carpeta_datos = carpeta_proyecto / "datos"

# Nombre del archivo Excel
ruta_excel = carpeta_datos / "ventas.xlsx"

# Crear el libro de Excel
libro = Workbook()
hoja = libro.active
hoja.title = "Ventas"

# Encabezados
encabezados = ["Fecha", "Producto", "Cantidad", "Precio"]
hoja.append(encabezados)

# Datos de ventas ficticias
ventas = [
    ["2026-09-02", "Teclado", 2, 350],
    ["2026-09-03", "Mouse", 3, 180],
    ["2026-09-05", "Monitor", 1, 2500],
    ["2026-09-10", "Teclado", 1, 350],
    ["2026-10-01", "Mouse", 5, 180],
    ["2026-10-04", "Monitor", 2, 2500],
    ["2026-10-12", "Teclado", 3, 350],
    ["2026-10-15", "Audífonos", 2, 600],
    ["2026-11-03", "Mouse", 4, 180],
    ["2026-11-08", "Audífonos", 1, 600],
]

# Agregar las ventas
for venta in ventas:
    hoja.append(venta)

# Dar formato a los encabezados
for celda in hoja[1]:
    celda.font = Font(bold=True, color="FFFFFF")
    celda.fill = PatternFill(
        fill_type="solid",
        fgColor="1F4E78"
    )
    celda.alignment = Alignment(horizontal="center")

# Ajustar el ancho de las columnas
for columna in hoja.columns:
    letra = get_column_letter(columna[0].column)
    hoja.column_dimensions[letra].width = 18

# Guardar el archivo
libro.save(ruta_excel)

print("Archivo Excel creado correctamente.")
print(f"Ubicacion: {ruta_excel}")
print(f"Total de ventas registradas: {len(ventas)}")
