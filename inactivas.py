import pandas as pd
import os

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
