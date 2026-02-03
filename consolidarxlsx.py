import os
import pandas as pd

def consolidar_excel():
    # Ruta de la carpeta con los archivos Excel
    ruta_carpeta = "" # RUTA DE LA CARPETA DONDE SE ENCUENTRAN LOS ARCHIVOS EXCEL

    # Verificar si la carpeta existe
    if not os.path.exists(ruta_carpeta):
        print(f"Error: La carpeta no existe: {ruta_carpeta}")
        return
    
    # Listar archivos Excel en la carpeta
    archivos_excel = []
    for archivo in os.listdir(ruta_carpeta):
        if archivo.endswith('.xlsx') and not archivo.startswith('CONSOLIDADO'):
            archivos_excel.append(os.path.join(ruta_carpeta, archivo))
    
    if not archivos_excel:
        print("No se encontraron archivos Excel (.xlsx) en la carpeta.")
        return
    
    print(f"Se encontraron {len(archivos_excel)} archivos Excel:")
    
    # Lista para almacenar todos los DataFrames
    datos_consolidados = []
    
    # Leer cada archivo Excel
    for archivo in archivos_excel:
        try:
            # Leer el archivo Excel sin saltar filas
            df = pd.read_excel(archivo, header=None)
            
            # Tomar datos desde la fila 2 (índice 1) hasta el último dato disponible
            if len(df) > 1:  # Si hay al menos 2 filas
                df_filtrado = df.iloc[1:]  # Tomar desde la fila 2 en adelante
                
                # Obtener solo el nombre del archivo (sin la ruta)
                nombre_archivo = os.path.basename(archivo)
                
                # Añadir una nueva columna con el nombre del archivo
                df_filtrado['Origen'] = nombre_archivo
                
                datos_consolidados.append(df_filtrado)
                print(f"  ✓ {nombre_archivo}: {len(df_filtrado)} filas")
            else:
                print(f"  ✗ {os.path.basename(archivo)}: No hay datos después de la fila 1")
                
        except Exception as e:
            print(f"  ✗ Error al leer {os.path.basename(archivo)}: {str(e)}")
    
    if not datos_consolidados:
        print("No se pudo leer datos de ningún archivo.")
        return
    
    # Consolidar todos los DataFrames en uno solo
    consolidado = pd.concat(datos_consolidados, ignore_index=True)
    
    # Ruta para el archivo consolidado
    ruta_consolidado = os.path.join(ruta_carpeta, "CONSOLIDADO.xlsx")
    
    # Guardar el archivo consolidado
    try:
        consolidado.to_excel(ruta_consolidado, index=False, header=False)
        print(f"\n✅ Se ha creado el archivo consolidado:")
        print(f"   Ubicación: {ruta_consolidado}")
        print(f"   Total de archivos procesados: {len(datos_consolidados)}")
        print(f"   Total de filas consolidadas: {len(consolidado)}")
        print(f"   Total de columnas: {len(consolidado.columns)} (incluyendo columna 'Origen')")
        
    except Exception as e:
        print(f"✗ Error al guardar el archivo consolidado: {str(e)}")

# Ejecutar la función
if __name__ == "__main__":
    consolidar_excel()