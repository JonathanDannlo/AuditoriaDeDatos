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

# DataFrame insumo Credito y Cartera
# Elimina los espacios adicionales en los nombres de las columnas
facturacion20241 = pd.read_excel(file_path,sheet_name='SQ2106_24_1ALL')
facturacion20241.columns = facturacion20241.columns.str.strip()

# Cruce de los DataFrames a partir de la referencia de la factura
df_piam20241fi = pd.merge(piam20241, facturacion20241[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura','Valor Ajuste','Valor Pagado','Valor Anulado','Saldo','Nombre de Destino']], left_on='RECIBO', right_on='Documento', how='inner')
df_piam20241fr = pd.merge(piam20241, facturacion20241[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura','Valor Ajuste','Valor Pagado','Valor Anulado','Saldo','Nombre de Destino']], left_on='RECIBO', right_on='Documento', how='right')

# Calculos de la matricula en el DataFrame Inner Academico Financiero
matriculaBruta = ['DERECHOS_MATRICULA','BIBLIOTECA_DEPORTES','LABORATORIOS','RECURSOS_COMPUTACIONALES','SEGURO_ESTUDIANTIL','VRES_COMPLEMENTARIOS','RESIDENCIAS','REPETICIONES']
meritoAcademico = ['CONVENIO_DESCENTRALIZACION','BECA','MATRICULA_HONOR','MEDIA_MATRICULA_HONOR','TRABAJO_GRADO','DOS_PROGRAMAS','DESCUENTO_HERMANO','ESTIMULO_EMP_DTE_PLANTA','ESTIMULO_CONYUGE','EXEN_HIJOS_CONYUGE_CATEDRA','EXEN_HIJOS_CONYUGE_OCASIONAL','HIJOS_TRABAJADORES_OFICIALES','ACTIVIDAES_LUDICAS_DEPOR','DESCUENTOS','SERVICIOS_RELIQUIDACION','DESCUENTO_LEY_1171']
df_piam20241fi['BRUTA'] = df_piam20241fi[matriculaBruta].sum(axis=1)
df_piam20241fi['BRUTAORD'] =  df_piam20241fi['BRUTA'] - df_piam20241fi['SEGURO_ESTUDIANTIL']
df_piam20241fi['NETAORD'] =  df_piam20241fi['BRUTAORD'] - df_piam20241fi['VOTO'].abs()
df_piam20241fi['MERITO'] = df_piam20241fi[meritoAcademico].sum(axis=1).abs()
df_piam20241fi['NETA'] =  df_piam20241fi['BRUTA'] - df_piam20241fi['VOTO'].abs() - df_piam20241fi['MERITO']
df_piam20241fi['NETAAPL'] =  df_piam20241fi['NETA'] - df_piam20241fi['SEGURO_ESTUDIANTIL']

# Validación de los campos de matricula neta a nivel academico y financiero
df_piam20241fi['FL_NETA'] = df_piam20241fi['NETA'] == df_piam20241fi['Valor Factura']
# Filtra las filas donde el valor es diferente
df_piam20241fi_diff = df_piam20241fi[~df_piam20241fi['FL_NETA']]

# Identificación de los registros del dataframe Financiero que no estan en el DataFrame Academico
unique_in_piam_right = df_piam20241fr[~df_piam20241fr['Documento'].isin(piam20241['RECIBO'])]
# Selección de columnas específicas
unique_in_piam_right_selected = unique_in_piam_right[['Documento', 'Id  factura', 'Estado Actual', 'Valor Factura','Valor Ajuste','Valor Pagado','Valor Anulado','Saldo','Nombre de Destino']]

# Identifica los registros duplicados según el RECIDO en el dataframe resultante del INNER entre el insumo academico y el financiero.
rgdpl_piam20241_boli = df_piam20241fi[df_piam20241fi.duplicated(subset='RECIBO', keep=False)]
# Identifica los registros unicos segun el RECIBO en el dataframe resultante del INNER entre el insumo academico y el financiero.
uni_piam20241_boli = df_piam20241fi[~df_piam20241fi['RECIBO'].isin(rgdpl_piam20241_boli['RECIBO'])]


# Guarda los dataframe cruzados segun el insumo academico y financiero
output_path = "/content/PIAM_2024_1_Conciliacion.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

    df_piam20241fi.to_excel(writer, sheet_name='Piam20241_Inner', index=False)
    unique_in_piam_right_selected.to_excel(writer, sheet_name='Unique_Piam_Right', index=False)
    rgdpl_piam20241_boli.to_excel(writer, sheet_name='RDB_Piam2024_1i', index=False)
    uni_piam20241_boli.to_excel(writer, sheet_name='RUI_Piam2024_1i', index=False)

print(f"Archivo guardado en {output_path}")
