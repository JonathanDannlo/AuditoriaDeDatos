import os
import pandas as pd
import matplotlib.pyplot as plt

# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/Libro1.xlsx'

if not os.path.isfile(file_path):
  raise FileNotFoundError(f"{file_path} no encontrado.")
else:
  print(f" Archivo {file_path} encontrado.")

# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abrierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")

# Carga los DataFrame de trabajo
try:
    # DataFrame insumo DARCA
    piam20241 = pd.read_excel(file_path,sheet_name='IDARCA2106_24_1_ALL', engine='openpyxl')
    piam20241.columns = piam20241.columns.str.strip()

    # DataFrame insumo Credito y Cartera
    facturacion20241 = pd.read_excel(file_path,sheet_name='SQ2106_24_1ALL', engine='openpyxl')
    facturacion20241.columns = facturacion20241.columns.str.strip()

    # DataFrame insumo PIAM 2024_1 CI
    PIAM20241CI = pd.read_excel(file_path,sheet_name='PIAM24_1_CI', engine='openpyxl')
    PIAM20241CI.columns = PIAM20241CI.columns.str.strip()

except Exception as e:
    print(f"Error al cargar los DataFrames: {e}")


# VALIDACIONES DE DATOS
# Cruza los DataFrames Academico y Facturación a partir de la referencia de la factura
df_piam20241fi = pd.merge(
    piam20241,
    facturacion20241[['Documento',
                      'Id  factura',
                      'Estado Actual',
                      'Valor Factura',
                      'Valor Ajuste','Valor Pagado',
                      'Valor Anulado',
                      'Saldo',
                      'Nombre de Destino']],
    left_on='RECIBO',
    right_on='Documento',
    how='inner')

df_piam20241fr = pd.merge(
    piam20241,
    facturacion20241[['Documento',
                      'Id  factura',
                      'Estado Actual',
                      'Valor Factura',
                      'Valor Ajuste',
                      'Valor Pagado',
                      'Valor Anulado',
                      'Saldo',
                      'Nombre de Destino']],
    left_on='RECIBO',
    right_on='Documento',
    how='right')


# Calculos del valor de la matricula en el DataFrame Inner Academico Financiero
matriculaBruta = ['DERECHOS_MATRICULA',
                  'BIBLIOTECA_DEPORTES',
                  'LABORATORIOS',
                  'RECURSOS_COMPUTACIONALES',
                  'SEGURO_ESTUDIANTIL',
                  'VRES_COMPLEMENTARIOS',
                  'RESIDENCIAS',
                  'REPETICIONES']

meritoAcademico = ['CONVENIO_DESCENTRALIZACION',
                   'BECA',
                   'MATRICULA_HONOR',
                   'MEDIA_MATRICULA_HONOR',
                   'TRABAJO_GRADO',
                   'DOS_PROGRAMAS',
                   'DESCUENTO_HERMANO',
                   'ESTIMULO_EMP_DTE_PLANTA',
                   'ESTIMULO_CONYUGE',
                   'EXEN_HIJOS_CONYUGE_CATEDRA',
                   'EXEN_HIJOS_CONYUGE_OCASIONAL',
                   'HIJOS_TRABAJADORES_OFICIALES',
                   'ACTIVIDAES_LUDICAS_DEPOR',
                   'DESCUENTOS',
                   'SERVICIOS_RELIQUIDACION',
                   'DESCUENTO_LEY_1171']

df_piam20241fi['BRUTA'] = df_piam20241fi[matriculaBruta].sum(axis=1)
df_piam20241fi['BRUTAORD'] = df_piam20241fi['BRUTA'] - df_piam20241fi['SEGURO_ESTUDIANTIL']
df_piam20241fi['NETAORD'] =  df_piam20241fi['BRUTAORD'] - df_piam20241fi['VOTO'].abs()
df_piam20241fi['MERITO'] = df_piam20241fi[meritoAcademico].sum(axis=1).abs()
df_piam20241fi['MTRNETA'] =  df_piam20241fi['BRUTA'] - df_piam20241fi['VOTO'].abs() - df_piam20241fi['MERITO']
df_piam20241fi['NETAAPL'] =  df_piam20241fi['MTRNETA'] - df_piam20241fi['SEGURO_ESTUDIANTIL']

# Validación de los campos de matricula neta a nivel academico y financiero
df_piam20241fi['FL_NETA'] = df_piam20241fi['MTRNETA'] == df_piam20241fi['Valor Factura']
# Filtra las filas donde el valor es diferente
df_piam20241fi_dif = df_piam20241fi[~df_piam20241fi['FL_NETA']]

# Identificación de los registros del dataframe Financiero que no estan en el DataFrame Academico
# Selección de columnas específicas requeridas
unique_in_piam_right = df_piam20241fr[~df_piam20241fr['Documento'].isin(piam20241['RECIBO'])]
unique_in_piam_right_selected = unique_in_piam_right[['Documento',
                                                      'Id  factura',
                                                      'Estado Actual',
                                                      'Valor Factura',
                                                      'Valor Ajuste',
                                                      'Valor Pagado',
                                                      'Valor Anulado',
                                                      'Saldo',
                                                      'Nombre de Destino']]

# Cruce de los DataFrames CI y Facturación a partir de la referencia de la factura y actualiza los registros
df_piam20241Cii = pd.merge(
    PIAM20241CI,
    facturacion20241[[
        'Documento',
        'Id  factura',
        'Estado Actual',
        'Valor Factura',
        'Valor Ajuste',
        'Valor Pagado',
        'Valor Anulado',
        'Saldo',
        'Nombre de Destino']],
    left_on='Id  factura',
    right_on='Id  factura',
    how='left')

df_piam20241Cii['Valor Factura_x'] = df_piam20241Cii['Valor Factura_y']
df_piam20241Cii['Valor Ajuste_x'] = df_piam20241Cii['Valor Ajuste_y']
df_piam20241Cii['Valor Pagado_x'] = df_piam20241Cii['Valor Pagado_y']
df_piam20241Cii['Valor Anulado_x'] = df_piam20241Cii['Valor Anulado_y']
df_piam20241Cii['Saldo_x'] = df_piam20241Cii['Saldo_y']
df_piam20241Cii.drop(columns=['Valor Factura_y', 'Valor Ajuste_y', 'Valor Pagado_y', 'Valor Anulado_y', 'Saldo_y'], inplace=True)

df_piam20241Cii['ESTADO FINANCIERO FSE'] = df_piam20241Cii['ESTADO FINANCIERO FSE'].fillna('')
df_piam20241Cii['ESTADO FINANCIERO ICETEX']  = df_piam20241Cii['ESTADO FINANCIERO ICETEX'] .fillna('')
df_piam20241Cii['ESTADO FINANCIERO BENEFICIO'] = df_piam20241Cii.apply(
    lambda row: f"{row['ESTADO FINANCIERO FSE']} - {row['ESTADO FINANCIERO ICETEX']}" if row['ESTADO FINANCIERO FSE'] and row['ESTADO FINANCIERO ICETEX']
    else row['ESTADO FINANCIERO FSE'] or row['ESTADO FINANCIERO ICETEX'] , axis=1)

# Identifica los registros duplicados según el RECIDO en el dataframe resultante del INNER entre el insumo academico y el financiero.
# Identifica los registros unicos segun el RECIBO en el dataframe resultante del INNER entre el insumo academico y el financiero.
df_piam20241fi_SinDuplicadosBoleta = df_piam20241fi.drop_duplicates(subset='RECIBO', keep='first')
df_piam20241fi_SinDuplicadosBoletaIdSnies = df_piam20241fi_SinDuplicadosBoleta.drop_duplicates(subset='ID-SNIES', keep='first')
registrosDuplicadosIDSnPiamfi = df_piam20241fi_SinDuplicadosBoleta[df_piam20241fi_SinDuplicadosBoleta.duplicated(subset='ID-SNIES', keep=False)]
registrosDuplicadosPiamfi = df_piam20241fi[df_piam20241fi.duplicated(subset='RECIBO', keep=False)]
registrosUnicosPiamfi = df_piam20241fi[~df_piam20241fi['RECIBO'].isin(registrosDuplicadosPiamfi['RECIBO'])]
registrosUnicosDuplicadosPiamfi = registrosDuplicadosPiamfi.drop_duplicates(subset='RECIBO', keep='first')

# Cruce dataframe PIAM CI con el dataframe piam fi
df_piam2024fci = pd.merge(df_piam20241fi_SinDuplicadosBoletaIdSnies,
                          PIAM20241CI[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']],
                          left_on='RECIBO', right_on='FACTURA', how='inner')


df_piam2024fcl = pd.merge(df_piam20241fi_SinDuplicadosBoletaIdSnies,
                          PIAM20241CI[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']],
                          left_on='RECIBO', right_on='FACTURA', how='left')


df_piam2024fcr = pd.merge(df_piam20241fi_SinDuplicadosBoletaIdSnies,
                          PIAM20241CI[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']],
                          left_on='RECIBO', right_on='FACTURA', how='right')

# Cruce registros duplicados por referencia de matricula insumo DARCA con CI PIAM 2024-1
df_registrosDuplicadosPiamfi = pd.merge(registrosDuplicadosPiamfi,
                          PIAM20241CI[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']],
                          left_on='RECIBO', right_on='FACTURA', how='left')

# Cruce registros duplicados por ID-SNIES de matricula insumo DARCA con CI PIAM 2024-1
df_registrosDuplicadosPiamfi = pd.merge(registrosDuplicadosIDSnPiamfi,
                          PIAM20241CI[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']],
                          left_on='RECIBO', right_on='FACTURA', how='left')


# Identificación de los registros del dataframe Piam2024_1 final que no estan en el DataFrame de registros unicos del insumo DARCA
# Identificación de los registros del dataframe Piam2024_1 final que no estan en el DataFrame de registros de la Certificacion Inicial
registrosDarcaNoCi = df_piam2024fcl[~df_piam2024fcl['RECIBO'].isin(df_piam2024fci['RECIBO'])]
registrosDarcaNoCiSelecionado = registrosDarcaNoCi[['IDENTIFICACION','CODIGO','SNIESPROGRAMA','RECIBO','Estado Actual','Valor Factura']]
registrosCiNoDarca = df_piam2024fcr[~df_piam2024fcr['RECIBO'].isin(df_piam2024fci['RECIBO'])]
registrosCiNoDarcaSelecionado = registrosCiNoDarca[['FACTURA','ESTADO CIVF','ESTADO SQ','ESTADO FINANCIERO FSE','ESTADO FINANCIERO ICETEX','AJSUTE VAL','NETA']]

# Identificación de la cantidad de boletas segun su estado de pago insumo DARCA y SQUID
filtro_estadoBeneficioPiam20241Ci = df_piam20241Cii.groupby('ESTADO FINANCIERO BENEFICIO')['FACTURA'].size().reset_index(name='Poblacion')
filtro_estadoPiam20241Ci = df_piam20241Cii.groupby('Estado Actual')['FACTURA'].size().reset_index(name='Poblacion')

filtro_estadoDarca = df_piam20241fi.groupby('Estado Actual')['RECIBO'].size().reset_index(name='Poblacion')

filtro_estadofinanciero = df_piam20241fi.groupby('Estado Actual')['RECIBO'].size().reset_index(name='Cantidad de boletas')
filtro_estado = df_piam20241fi_SinDuplicadosBoletaIdSnies.groupby('Estado Actual')['RECIBO'].size().reset_index(name='Cantidad de boletas')
filtro_estadofinancieroSQ = unique_in_piam_right_selected.groupby('Estado Actual')['Id  factura'].size().reset_index(name='Cantidad de boletas')

# Guarda los dataframe cruzados segun el insumo academico y financiero
output_path = "/content/PIAM_2024_1_Conciliacion.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

    # Obtener el workbook y el worksheet

    filtro_estadoPiam20241Ci.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=1, index=False)
    filtro_estadoBeneficioPiam20241Ci.to_excel(writer, sheet_name='Generalidades', startrow=6, startcol=1, index=False)

    filtro_estadoDarca.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=4, index=False)

    workbook  = writer.book
    worksheet = writer.sheets['Generalidades']
    formato = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    worksheet.merge_range('B1:C1', "PIAM CI 2024-1",formato)
    worksheet.merge_range('E1:F1', "PIAM DARCA 2024-1",formato)

    df_piam20241Cii.to_excel(writer, sheet_name='PIAM20241Ci', index=False)
    df_piam20241fi.to_excel(writer, sheet_name='PIAM20241Darca', index=False)

    df_registrosDuplicadosPiamfi.to_excel(writer, sheet_name='DuplicadosXBoleta', index=False)
    df_registrosDuplicadosPiamfi.to_excel(writer, sheet_name='DuplicadosXIdSnies', index=False)

    #filtro_estadofinanciero.to_excel(writer, sheet_name='Generalidades', startrow=4, startcol=1, index=False)
    #filtro_estado.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=8, index=False)
    #filtro_estadofinancieroSQ.to_excel(writer, sheet_name='Generalidades', startrow=12, startcol=7, index=False)
    #df_piam20241fi_SinDuplicadosBoletaIdSnies.to_excel(writer, sheet_name='PIAM20241DARCAVF', index=False)
    #df_piam2024fci.to_excel(writer, sheet_name='PIAM20241DARCAVFCI', index=False)
    #registrosDarcaNoCiSelecionado.to_excel(writer, sheet_name='DarcaNoCI', index=False)
    #registrosCiNoDarcaSelecionado.to_excel(writer, sheet_name='CiNoDarca', index=False)
    #unique_in_piam_right_selected.to_excel(writer, sheet_name='RegistrosSoloSQ', index=False)
    #df_piam20241fi_SinDuplicadosBoleta.to_excel(writer, sheet_name='UnicosXBoletaSinDpl', index=False)
    #registrosUnicosPiamfi.to_excel(writer, sheet_name='UnicosXBoleta', index=False)
    #registrosUnicosDuplicadosPiamfi.to_excel(writer, sheet_name='DuplicadosUnicosXBoleta', index=False)


print(f"Archivo guardado en {output_path}")
