
from openpyxl import load_workbook
from pathlib import Path

# Localizar el archivo Excel
carpeta_proyecto = Path(__file__).resolve().parent.parent
ruta_excel = carpeta_proyecto / "datos" / "ventas.xlsx"

libro = load_workbook(ruta_excel, data_only=True)
hoja = libro["Ventas"]

# Diccionario para acumular las ventas por mes
resumen_mensual = {}

for fila in hoja.iter_rows(min_row=2, values_only=True):
    fecha, producto, cantidad, precio = fila

    # Obtener el año y el mes: por ejemplo, 2026-09
    mes = fecha[:7]

    # Calcular el importe
    importe = cantidad * precio

    # Inicializar el mes si no existe
    if mes not in resumen_mensual:
        resumen_mensual[mes] = {
            "unidades": 0,
            "ingresos": 0
        }

    # Acumular resultados
    resumen_mensual[mes]["unidades"] += cantidad
    resumen_mensual[mes]["ingresos"] += importe

# Mostrar el resumen mensual
print("VENTAS POR MES")
print("-" * 55)

for mes, datos in sorted(resumen_mensual.items()):
    print(
        f"{mes} | "
        f"Unidades: {datos['unidades']:2} | "
        f"Ingresos: ${datos['ingresos']:,.2f}"
    )

libro.close()
