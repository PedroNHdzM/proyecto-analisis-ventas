#Ahora vamos a aprender una herramienta importante de Python. 
# Hasta ahora hemos validado datos con condiciones if; pero hay 
# situaciones en las que una operación puede generar
# una excepción, 
# como intentar abrir un archivo que no existe.
# try: contiene la operación que podría fallar.
# except FileNotFoundError: captura específicamente el error de archivo inexistente.
# print("El programa continua ejecutandose."): demuestra que el programa puede continuar después de controlar el error.
from pathlib import Path

ruta_archivo = Path("archivo_inexistente.txt")

try:
    with open(ruta_archivo, "r", encoding="utf-8") as archivo:
        contenido = archivo.read()
        print(contenido)

except FileNotFoundError:
    print("ERROR: El archivo no existe.")

print("El programa continua ejecutandose.")
