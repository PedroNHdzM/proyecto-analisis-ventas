
from openpyxl import load_workbook
from collections import defaultdict
import matplotlib.pyplot as plt
from datetime import datetime, date

# Leer el archivo Excel
ruta_excel = "Proyecto_Final/datos/ventas.xlsx"
libro = load_workbook(ruta_excel, data_only=True)
hoja = libro["Ventas"]

# Diccionario para acumular ingresos por mes
ingresos_mensuales = defaultdict(float)

# Leer las ventas, comenzando en la fila 2
for fecha, producto, cantidad, precio in hoja.iter_rows(
    min_row=2, values_only=True
):
    #isinstance(fecha, str) comprueba si la fecha es texto.
    #datetime.strptime(fecha, "%Y-%m-%d") convierte textos como "2026-09-02" en fechas.
    #strftime("%Y-%m") obtiene el año y el mes, por ejemplo, 2026-09.
    
    # Convertir el texto en una fecha si es necesario
    if isinstance(fecha, str):
        fecha = datetime.strptime(fecha, "%Y-%m-%d")

    mes = fecha.strftime("%Y-%m")
    importe = cantidad * precio

    ingresos_mensuales[mes] += importe


libro.close()

# Preparar los datos para la grafica
#keys() obtiene los meses encontrados en el archivo.
#sorted() los ordena cronológicamente porque tienen el formato AAAA-MM.
#La segunda instrucción obtiene los ingresos correspondientes a cada mes.
meses = sorted(ingresos_mensuales.keys())
ingresos = [ingresos_mensuales[mes] for mes in meses]

# Crear la grafica
plt.figure(figsize=(9, 5))
barras = plt.bar(meses, ingresos)

plt.title("Ventas mensuales - 2026")
plt.xlabel("Mes")
plt.ylabel("Ingresos ($)")
plt.bar_label(barras, fmt="$%.2f")

plt.tight_layout()

# Guardar la imagen
ruta_imagen = "Proyecto_Final/reportes/ventas_mensuales_automaticas.png"
plt.savefig(ruta_imagen, dpi=300, bbox_inches="tight")

plt.show()

print(f"\nGrafica guardada en: {ruta_imagen}")
