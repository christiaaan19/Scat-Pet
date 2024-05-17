import os
import pandas as pd

# Rutas de los archivos
ruta_excel_original = r'C:\Users\cristian.montoyab\Desktop\automatismo\excel\bd-rodamiento-2024.xlsm'
ruta_excel_nuevo = r'C:\Users\cristian.montoyab\Desktop\automatismo\guardado\automatismo.xlsx'

# Verificar la ruta del archivo original y permisos
print(f"Verificando la ruta del archivo original: {ruta_excel_original}")

# Listar el contenido del directorio para verificar la presencia del archivo
directorio_input = os.path.dirname(ruta_excel_original)
print(f"Contenido del directorio {directorio_input}:")
print(os.listdir(directorio_input))

if not os.path.exists(ruta_excel_original):
    raise FileNotFoundError(f"No se encontró el archivo: {ruta_excel_original}")

if not os.access(ruta_excel_original, os.R_OK):
    raise PermissionError(f"No se tienen permisos de lectura para el archivo: {ruta_excel_original}")

# Nombre de la hoja en el archivo original
nombre_hoja_original = 'FESTIVOS - LIDER'

# Leer el archivo original
df_original = pd.read_excel(ruta_excel_original, sheet_name=nombre_hoja_original, engine='openpyxl')

# Extraer las columnas A, B y C (considerando que A=0, B=1, C=2 en índice 0-basado)
columnas_abc = df_original.iloc[:, [0, 1, 2]]

# Extraer las columnas E y F (considerando que E=4, F=5 en índice 0-basado)
columnas_ef = df_original.iloc[:, [4, 5]]

# Crear un ExcelWriter para guardar múltiples DataFrames en diferentes hojas
with pd.ExcelWriter(ruta_excel_nuevo, engine='openpyxl') as writer:
    # Guardar las columnas A, B y C en una hoja llamada "FESTIVOS"
    columnas_abc.to_excel(writer, sheet_name='FESTIVOS', index=False)
    
    # Guardar las columnas E y F en una hoja llamada "OTRA INFORMACION"
    columnas_ef.to_excel(writer, sheet_name='OTRA INFORMACION', index=False)

print(f"Las columnas especificadas han sido extraídas y guardadas en {ruta_excel_nuevo}")













# Rutas y nombres de archivos
ruta_excel_original = r'C:\Users\cristian.montoyab\Desktop\automatismo\excel\bd-rodamiento-2024.xlsm'
nombre_hoja_origen = 'INACTIVAS'
ruta_excel_nuevo = r'C:\Users\cristian.montoyab\Desktop\automatismo\guardado\automatismo.xlsx'

# Verificar si el archivo de origen existe
if not os.path.exists(ruta_excel_original):
    raise FileNotFoundError(f"No se encuentra el archivo de origen: {ruta_excel_original}")

# Leer la hoja "INACTIVAS" del archivo de origen
df = pd.read_excel(ruta_excel_original, sheet_name=nombre_hoja_origen)

# Verificar si el archivo de destino ya existe
if os.path.exists(ruta_excel_nuevo):
    # Abrir el archivo existente en modo de adición con openpyxl
    with pd.ExcelWriter(ruta_excel_nuevo, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja_origen, index=False)
else:
    # Crear un nuevo archivo de Excel y agregar la hoja
    with pd.ExcelWriter(ruta_excel_nuevo, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=nombre_hoja_origen, index=False)

print("La hoja ha sido duplicada y guardada exitosamente.")









import os
from git import Repo

# Directorio local donde se encuentra o se clonará el repositorio (nueva ruta)
directorio_local = "C:\\Users\\cristian.montoyab\\Desktop\\automatismo\\guardado"

# URL del repositorio en GitHub
url_repositorio = "https://github.com/christiaaan19/rodamiento"

# Verificar si el repositorio ya existe en el directorio local
if not os.path.exists(os.path.join(directorio_local, ".git")):
    print("El repositorio no existe en el directorio especificado. Clonando...")
    Repo.clone_from(url_repositorio, directorio_local)
else:
    print("El repositorio ya existe en el directorio especificado. Abriendo...")