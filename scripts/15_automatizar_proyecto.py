# Qué hace el código?
# - subprocess: permite ejecutar otros programas de Python desde nuestro programa principal.
# - sys.executable: utiliza el mismo intérprete de Python con el que estamos ejecutando el programa actual, normalmente el de tu entorno .venv.
# - Path(__file__): identifica dónde está guardado el archivo.
# - programas: contiene los tres scripts que vamos a ejecutar.
# - for: ejecuta los programas uno por uno.
# - returncode: permite detectar si algún programa terminó con error. Si falla, detenemos el proceso en lugar de continuar como si todo estuviera bien.
import subprocess
import sys
from pathlib import Path
from datetime import datetime
import time

# Obtener la carpeta principal del proyecto
CARPETA_PROYECTO = Path(__file__).resolve().parent.parent
CARPETA_SCRIPTS = CARPETA_PROYECTO / "scripts"
CARPETA_REPORTES = CARPETA_PROYECTO / "reportes"
ARCHIVO_LOG = CARPETA_REPORTES / "ejecuciones.log"
#Archivos necesarios para ejecutar el proyecto
archivos_necesarios = [
    CARPETA_PROYECTO / "datos" / "ventas.xlsx",
    CARPETA_PROYECTO / "datos" / "costos_productos.xlsx"
]


# Programas que se ejecutaran en orden
programas = [
    "09_crear_reporte_ventas.py",
    "13_reporte_rentabilidad_excel.py",
    "14_dashboard_ventas.py",
    "21_historial_por_fecha.py",
    "22_reporte_calidad_datos.py",
    "27_reporte_calidad_avanzado.py",
    "28_reporte_general.py",
    "29_dashboard_ejecutivo.py",
    "23_crear_respaldo_reportes.py"
   
]
def registrar_ejecucion(estado, duracion=None):
    # duracion=None permite que la función reciba un segundo dato opcional.
    # Si recibe una duración, la agrega al registro.
    # Si no recibe duración, registra el mensaje como lo hacía antes.
    fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if duracion is not None:
        mensaje = f"{fecha_hora} - {estado} - Duracion: {duracion:.2f} segundos\n"
    else:
        mensaje = f"{fecha_hora} - {estado}\n"

    with open(ARCHIVO_LOG, "a", encoding="utf-8") as archivo:
        archivo.write(mensaje)
        
inicio = time.time()        
print("=" * 45)
print(" AUTOMATIZACION DEL PROYECTO DE VENTAS")
print("=" * 45)
print("\nValidando archivos necesarios...")

for archivo in archivos_necesarios:
    if not archivo.is_file():
        mensaje_error = f"ERROR: No se encontro {archivo.name}"
        print(f"\n{mensaje_error}")

        registrar_ejecucion(mensaje_error) #guarda el error en el archivo .log antes de detener el programa.

        sys.exit(1)

    print(f"[OK] Encontrado: {archivo.name}")



print("\nTodos los archivos de entrada estan disponibles.")

# Ejecutar los programas en orden
# CARPETA_PROYECTO representa la carpeta Proyecto_Final.
# CARPETA_PROYECTO.parent representa la carpeta superior: Python_Curso.
for programa in programas:

    ruta_programa = CARPETA_SCRIPTS / programa

    print(f"\nEjecutando: {programa}")
    #capture_output=True: captura los mensajes que produce el programa, tanto normales como de error.
    #text=True: permite recibir esos mensajes como texto, en lugar de datos binarios.
    #Esto nos permitirá consultar:
    # - resultado.stdout: salida normal del programa.
    # - resultado.stderr: mensajes de error.
    # - resultado.returncode: código de terminación. 
    # Cero significa que terminó correctamente; un valor distinto 
    # de cero indica un fallo.
    resultado = subprocess.run(
    [sys.executable, str(ruta_programa)],
    cwd=CARPETA_PROYECTO.parent,
    capture_output=True,
    text=True
)
    # ¿Qué estamos agregando?
    # - returncode == 0: significa que el programa terminó correctamente.
    # - returncode != 0: significa que ocurrió un error.
    # - sys.exit(1): detiene la automatización.
    # - registrar_ejecucion(): guarda el error en nuestro archivo de registro.
    # ¿Qué estamos aprendiendo?
    #   1. strip() elimina espacios y saltos de línea sobrantes.
    #   2. stderr contiene los mensajes de error del programa ejecutado.
    #   3. if detalle_error comprueba si recibimos algún mensaje.
    #   4. registrar_ejecucion() guarda el error en nuestro archivo de historial.
    #   5. sys.exit(1) detiene la automatización cuando encuentra un fallo.
    if resultado.returncode == 0:
        print(f"Correcto: {programa}")

    else:
        print(f"Error al ejecutar: {programa}")

        detalle_error = resultado.stderr.strip()

        if detalle_error:
            print("Detalle del error:")
            print(detalle_error)
        else:
            detalle_error = "El programa termino con codigo diferente de cero"

        fin = time.time()
        duracion = fin - inicio

        registrar_ejecucion(
            f"ERROR al ejecutar {programa} - Detalle: {detalle_error}",
            duracion
        )

        sys.exit(1)
        
# if resultado.returncode != 0:
#     print("\n" + "=" * 45)
#     print("ERROR EN LA AUTOMATIZACION")
#     print("=" * 45)
#     print(f"Programa que fallo: {programa}")
#     print(f"Codigo de salida: {resultado.returncode}")
#     print("El proceso se ha detenido.")

#     mensaje_error = f"ERROR al ejecutar {programa}"
#     registrar_ejecucion(mensaje_error)

#     sys.exit(resultado.returncode)


print("\n" + "=" * 45)
print("PROCESO COMPLETADO CORRECTAMENTE")
print("=" * 45)
fin = time.time()
duracion = fin - inicio
registrar_ejecucion("PROCESO COMPLETADO CORRECTAMENTE", duracion)

print(f"Tiempo total de ejecucion: {duracion:.2f} segundos")
print("Ejecucion registrada en el archivo log.")
# - 0: normalmente significa que el programa terminó correctamente.
# - Un valor distinto de 0: normalmente indica que ocurrió un error.
#La instrucción sys.exit(resultado.returncode) transmite el código de error y evita que se ejecuten los siguientes reportes.