import os
import pandas as pd

# Ruta de la carpeta donde están los CSV
carpeta_origen = r"C:\Users\pracmt\Downloads\archivos a actualizar\archivos a actualizar"
carpeta_salida=r"C:\Users\pracmt\Downloads\archivos a actualizar"
contador=0
# Recorrer todos los archivos dentro de la carpeta
for archivo in os.listdir(carpeta_origen):
    if archivo.lower().endswith(".csv"):
        ruta_csv = os.path.join(carpeta_origen, archivo)
        contador=contador+1
        # Nombre nuevo del archivo .xlsx
        nombre_xlsx = archivo[:-4] + ".xlsx"
        ruta_xlsx = os.path.join(carpeta_salida, nombre_xlsx)

        try:
            # Leer el CSV (maneja correctamente comas y codificación)
            df = pd.read_csv(ruta_csv, sep=",", encoding="utf-8", engine="python")

            # Guardar como Excel
            df.to_excel(ruta_xlsx, index=False)

            print(f"Convertido: {archivo} → {nombre_xlsx}")

        except Exception as e:
            contador=contador-1
            print(f"Error al convertir {archivo}: {e}")
print(f"total de archivos procesados: {contador}")