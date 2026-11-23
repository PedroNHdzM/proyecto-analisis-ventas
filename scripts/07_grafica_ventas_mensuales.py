
import matplotlib.pyplot as plt

# Datos del analisis mensual
meses = ["Septiembre", "Octubre", "Noviembre"]
ingresos = [4090, 8150, 1320]

# Crear la grafica de barras
plt.figure(figsize=(9, 5))
barras = plt.bar(meses, ingresos)

# Titulos y etiquetas
plt.title("Ventas mensuales - 2026")
plt.xlabel("Mes")
plt.ylabel("Ingresos ($)")

# Mostrar el importe encima de cada barra
plt.bar_label(barras, fmt="$%.2f")

# Ajustar la presentacion
plt.tight_layout()
# Guardar la grafica como imagen
plt.savefig(
    "Proyecto_Final/reportes/ventas_mensuales.png",
    dpi=300,
    bbox_inches="tight"
)
# dpi=300: guarda la imagen con buena resolución.
# bbox_inches="tight": ajusta los márgenes para evitar espacios innecesarios.
# La ruta indica que la imagen se guardará en la carpeta reportes.

# Mostrar la grafica
plt.show()

# plt.figure(figsize=(9, 5)): define el tamaño de la gráfica.
# plt.bar(meses, ingresos): crea las barras con los datos mensuales.
# plt.title(): establece el título.
# plt.xlabel() y plt.ylabel(): identifican los ejes.
# plt.bar_label(): muestra el importe encima de cada barra.
# plt.show(): abre una ventana con la gráfica.