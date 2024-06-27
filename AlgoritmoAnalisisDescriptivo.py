import os
import pandas as pd
import matplotlib.pyplot as plt

# Carga del archivo al dataframe de trabajo
file_path = "/content/Libro1.xlsx"

if not os.path.isfile(file_path):
  raise FileNotFoundError(f"{file_path} not found.")

# DataFrame insumo DARCA
# Elimina los espacios adicionales en los nombres de las columnas
piam20241 = pd.read_excel(file_path,sheet_name='IDARCA2106_24_1_ALL')
piam20241.columns = piam20241.columns.str.strip()
#piam20241.info()

# DataFrame insumo Credito y Cartera
# Elimina los espacios adicionales en los nombres de las columnas
facturacion20241 = pd.read_excel(file_path,sheet_name='SQ2106_24_1ALL')
facturacion20241.columns = facturacion20241.columns.str.strip()
#facturacion20241.info()

# Verificar los nombres de las columnas
#print(piam20241.columns)
#print(facturacion20241.columns)

# Cruce de los DataFrames a partir de la referencia de la factura
df_piam20241fi = pd.merge(piam20241, facturacion20241[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura']], left_on='RECIBO', right_on='Documento', how='inner')
df_piam20241fl = pd.merge(piam20241, facturacion20241[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura']], left_on='RECIBO', right_on='Documento', how='left')
df_piam20241fr = pd.merge(piam20241, facturacion20241[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura']], left_on='RECIBO', right_on='Documento', how='right')
#print(df_piam20241f.head())

df_sq20241fi = pd.merge(facturacion20241,piam20241[['RECIBO', 'ID-SNIES', 'CODIGO']], left_on='Documento', right_on='RECIBO', how='inner')
df_sq20241fl = pd.merge(facturacion20241,piam20241[['RECIBO', 'ID-SNIES', 'CODIGO']], left_on='Documento', right_on='RECIBO', how='left')
df_sq20241fr = pd.merge(facturacion20241,piam20241[['RECIBO', 'ID-SNIES', 'CODIGO']], left_on='Documento', right_on='RECIBO', how='right')
#print(df_sq20241f.head())

# Guarda los registros duplicados segun la identiifacion asociada a un programa de pregrado en un nuevo archivo Excel
output_path = "/content/PIAM_2024_1_Conciliacion.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
    df_piam20241fi.to_excel(writer, sheet_name='Piam20241_Inner', index=False)
    df_piam20241fl.to_excel(writer, sheet_name='Piam20241_Left', index=False)
    df_piam20241fr.to_excel(writer, sheet_name='Piam20241_Right', index=False)
    df_sq20241fi.to_excel(writer, sheet_name='Sq20241_Inner', index=False)
    df_sq20241fl.to_excel(writer, sheet_name='Sq20241_Left', index=False)
    df_sq20241fr.to_excel(writer, sheet_name='Sq20241_Right', index=False)
