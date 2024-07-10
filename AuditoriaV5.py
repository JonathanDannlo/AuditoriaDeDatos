import os
import pandas as pd
import matplotlib.pyplot as plt

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

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
    # Lectura de los insumos en un diccionario de dataframes
    # DataFrame insumo DARCA = IDARCA2106_24_1_ALL
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
# Cruces de información entre los diferentes insumos
# Cruza los DataFrames Academico y Facturación a partir de la referencia de la factura
df_piam20241_dfi = pd.merge(
    piam20241,
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
    left_on='RECIBO',
    right_on='Documento',
    how='inner')

# Calcula el valor de la matricula en el DataFrame Inner Academico Financiero
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

df_piam20241_dfi['BRUTA'] = df_piam20241_dfi[matriculaBruta].sum(axis=1)
df_piam20241_dfi['BRUTAORD'] = df_piam20241_dfi['BRUTA'] - df_piam20241_dfi['SEGURO_ESTUDIANTIL']
df_piam20241_dfi['NETAORD'] =  df_piam20241_dfi['BRUTAORD'] - df_piam20241_dfi['VOTO'].abs()
df_piam20241_dfi['MERITO'] = df_piam20241_dfi[meritoAcademico].sum(axis=1).abs()
df_piam20241_dfi['MTRNETA'] =  df_piam20241_dfi['BRUTA'] - df_piam20241_dfi['VOTO'].abs() - df_piam20241_dfi['MERITO']
df_piam20241_dfi['NETAAPL'] =  df_piam20241_dfi['MTRNETA'] - df_piam20241_dfi['SEGURO_ESTUDIANTIL']

# Validación de los campos de matricula neta a nivel academico y financiero
df_piam20241_dfi['FL_NETA'] = df_piam20241_dfi['MTRNETA'] == df_piam20241_dfi['Valor Factura']
# Filtra las filas donde el valor es diferente
df_piam20241_dfi_difNeta = df_piam20241_dfi[~df_piam20241_dfi['FL_NETA']]

columnas_df_piam20241_dfi = [
  'IDENTIFICACION','CODIGO','SNIESPROGRAMA','RECIBO','DERECHOS_MATRICULA','BIBLIOTECA_DEPORTES',
  'LABORATORIOS','RECURSOS_COMPUTACIONALES','SEGURO_ESTUDIANTIL','VRES_COMPLEMENTARIOS','RESIDENCIAS',
  'REPETICIONES','VOTO','CONVENIO_DESCENTRALIZACION','BECA', 'MATRICULA_HONOR','MEDIA_MATRICULA_HONOR',
  'TRABAJO_GRADO','DOS_PROGRAMAS','DESCUENTO_HERMANO','ESTIMULO_EMP_DTE_PLANTA','ESTIMULO_CONYUGE',
  'EXEN_HIJOS_CONYUGE_CATEDRA','EXEN_HIJOS_CONYUGE_OCASIONAL','HIJOS_TRABAJADORES_OFICIALES',
  'ACTIVIDAES_LUDICAS_DEPOR','DESCUENTOS','SERVICIOS_RELIQUIDACION','DESCUENTO_LEY_1171','GRATUIDAD_MATRICULA',
  'ESTRATO','TIPOIDENTIFICACION','CREDITOSPENSUM','CREDITOSAPROBADOS','CREDITOSMATRICULADOS','SUPERACREDITOS',
  'FACULTAD','PROGRAMA','PRIMERNOMBRE','SEGUNDONOMBRE','PRIMERAPELLIDO','SEGUNDOAPELLIDO','GENERO',
  'ZONARESIDENCIA','IDMUNICIPIOPROGRAMA','NACIMIENTO','ID_PAIS_NACIMIENTO','IDMUNICIPIONACIMIENTO',
  'ESTUDIANTEREINGRESO','ANIOINGRESO','PERIODOINGRESO','TELEFONO','CELULAR','EMAILPERSONAL',
  'EMAILINSTITUCIONAL','PUEBLOINDIGENA','COMUNIDADNEGRA','GRUPOSISBEN','FONDOICETEX','RESOLUCIONICETEX',
  'VALORGIROICETEX','BRUTA','BRUTAORD', 'NETAORD', 'MERITO', 'MTRNETA', 'NETAAPL','FL_NETA','Id  factura',
  'Estado Actual','Valor Factura','Valor Ajuste','Valor Pagado','Valor Anulado','Saldo','Nombre de Destino'
]

df_piam20241_dfi = df_piam20241_dfi[columnas_df_piam20241_dfi]
df_piam20241_dfi.columns = df_piam20241_dfi.columns.str.strip()


# Cruza los DataFrames CI y Facturación a partir de la referencia de la factura y actualiza los registros
df_piam20241_Cifl = pd.merge(
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

df_piam20241_Cifl['Valor Factura_x'] = df_piam20241_Cifl['Valor Factura_y']
df_piam20241_Cifl['Valor Ajuste_x'] = df_piam20241_Cifl['Valor Ajuste_y']
df_piam20241_Cifl['Valor Pagado_x'] = df_piam20241_Cifl['Valor Pagado_y']
df_piam20241_Cifl['Valor Anulado_x'] = df_piam20241_Cifl['Valor Anulado_y']
df_piam20241_Cifl['Saldo_x'] = df_piam20241_Cifl['Saldo_y']
df_piam20241_Cifl.drop(columns=['Valor Factura_y', 'Valor Ajuste_y', 'Valor Pagado_y', 'Valor Anulado_y', 'Saldo_y'], inplace=True)

df_piam20241_Cifl['ESTADO FINANCIERO FSE'] = df_piam20241_Cifl['ESTADO FINANCIERO FSE'].fillna('')
df_piam20241_Cifl['ESTADO FINANCIERO ICETEX']  = df_piam20241_Cifl['ESTADO FINANCIERO ICETEX'] .fillna('')
df_piam20241_Cifl['ESTADO FINANCIERO BENEFICIO'] = df_piam20241_Cifl.apply(
    lambda row: f"{row['ESTADO FINANCIERO FSE']} - {row['ESTADO FINANCIERO ICETEX']}" if row['ESTADO FINANCIERO FSE'] and row['ESTADO FINANCIERO ICETEX']
    else row['ESTADO FINANCIERO FSE'] or row['ESTADO FINANCIERO ICETEX'] , axis=1)

columnas_df_piam20241_Cifl = ['IDENTIFICACION','FACTURA','CODIGO','CATEGORIA','OIDESTUDIANTE','ID-PRO','ID-PRO SNIES','CODIGO SNIES','PRO SIMCA','PRO - SNIES',
 'SEMESTRE','Id  factura','DERECHOS_MATRICULA','BIBLIOTECA_DEPORTES','LABORATORIOS','RECURSOS_COMPUTACIONALES','SEGURO_ESTUDIANTIL',
 'VRES_COMPLEMENTARIOS','RESIDENCIAS','REPETICIONES','VOTO','CONVENIO_DESCENTRALIZACION','BECA','MATRICULA_HONOR','MEDIA_MATRICULA_HONOR',
 'TRABAJO_GRADO','DOS_PROGRAMAS','DESCUENTO_HERMANO','ESTIMULO_EMP_DTE_PLANTA','ESTIMULO_CONYUGE','EXEN_HIJOS_CONYUGE_CATEDRA','EXEN_HIJOS_CONYUGE_OCASIONAL',
 'HIJOS_TRABAJADORES_OFICIALES','ACTIVIDAES_LUDICAS_DEPOR','DESCUENTOS','SERVICIOS_RELIQUIDACION','DESCUENTO DE LEY  1171','GRATUIDAD_MATRICULA',
 'FACULTAD','PROGRAMA','PRIMERNOMBRE','SEGUNDONOMBRE','PRIMERAPELLIDO','SEGUNDOAPELLIDO','TIPO_IDENTIFICACION','GENERO','TELEFONO','CELULAR','TELFONULAR','EMAIL_INSTITUCIONAL',
 'EMAIL_PERSONAL','SEDE', 'SEDE ID','PAIS_PROCEDENCIA','CODIGO_IDENTIFICACION_PAIS','DEPARTAMENTO_PROCEDENCIA','CODIGO_DEPARTAMENTO','MUNICIPIO_PROCEDENCIA','CODIGO_MUNICIPIO',
 'ZONARESIDENCIA','ESTRATO','SISBEN','GRUPOSISBEN','FONDOICETEX','COMUNIDADES_NEGRAS','INDIGENA','DISCAPACIDAD','DIRECCION','NOMBRE_COLEGIO','EDAD','NACIMIENTO',
 'CREDITOS_DEL_PROGRAMA','DURACION_DEL_PROGRAMA','CREDITOS_MATRICULADOS','TOTAL_CREDITOS_APROBADOS_SIN_REQUISITO_GRADO','A','B_AJUSTADO','C_AJUSTADO','#PA',
 'BRUTA','BRUTA ORD','MERITO','SUFRAGANTE','NETA ORD','NETA','NETA AP','ESTADO VALIDADO 130634','VALOR A CUBRIR','TOTAL_PERIODOS_APROBADOS','PERIODOS_FINANCIADOS',
 'PERIODOS_A_FINANCIAR','ESTADO CIVF','RESUL VAL2','ICETEX FONDO','ICETEX RESOLUCION','ICETEX GIRO','Estado Actual','Valor Factura_x','Valor Ajuste_x',
 'Valor Pagado_x', 'Valor Anulado_x','Saldo_x','FSE APLICAR','FSE REINTEGRAR','ESTADO FINANCIERO FSE','ICETEX REINTEGRO','ICETEX APLICAR','ESTADO FINANCIERO ICETEX',
 'Nombre de Destino','GRADO PREVIO','AJSUTE VAL','ESTADO FINANCIERO BENEFICIO']

df_piam20241_Cifl = df_piam20241_Cifl[columnas_df_piam20241_Cifl]
df_piam20241_Cifl.rename(
    columns={
        'RESUL VAL2':'RESULTADO DE VALIDACION',
        'Valor Factura_x':'Valor Factura',
        'Valor Ajuste_x':'Valor Ajuste',
        'Valor Pagado_x':'Valor Pagado',
        'Valor Anulado_x':'Valor Anulado',
        'Saldo_x':'Saldo'
        },
    inplace=True)
df_piam20241_Cifl.columns = df_piam20241_Cifl.columns.str.strip()


# Cruce de los DataFrames Piam Darca y PIAM Ci actualizados financieramente a partir de la referencia de la factura
df_piam20241_Cidl = pd.merge(
    df_piam20241_dfi,
    df_piam20241_Cifl[[
        'FACTURA',
        'ESTADO CIVF',
        'ESTADO FINANCIERO BENEFICIO',
        'FSE APLICAR',
        'FSE REINTEGRAR',
        'ESTADO FINANCIERO FSE',
        'ICETEX RESOLUCION',
        'ICETEX GIRO',
        'ICETEX REINTEGRO',
        'ICETEX APLICAR',
        'ESTADO FINANCIERO ICETEX',
        'RESULTADO DE VALIDACION',
        'GRADO PREVIO']],
    left_on='RECIBO',
    right_on='FACTURA',
    how='left')

columnas_df_piam20241_Cidl = [
    'IDENTIFICACION', 'CODIGO', 'SNIESPROGRAMA', 'RECIBO','Id  factura','DERECHOS_MATRICULA', 'BIBLIOTECA_DEPORTES', 'LABORATORIOS',
    'RECURSOS_COMPUTACIONALES', 'SEGURO_ESTUDIANTIL','VRES_COMPLEMENTARIOS', 'RESIDENCIAS', 'REPETICIONES', 'VOTO','CONVENIO_DESCENTRALIZACION',
    'BECA', 'MATRICULA_HONOR','MEDIA_MATRICULA_HONOR', 'TRABAJO_GRADO', 'DOS_PROGRAMAS','DESCUENTO_HERMANO', 'ESTIMULO_EMP_DTE_PLANTA', 'ESTIMULO_CONYUGE',
    'EXEN_HIJOS_CONYUGE_CATEDRA', 'EXEN_HIJOS_CONYUGE_OCASIONAL','HIJOS_TRABAJADORES_OFICIALES', 'ACTIVIDAES_LUDICAS_DEPOR','DESCUENTOS',
    'SERVICIOS_RELIQUIDACION', 'DESCUENTO_LEY_1171','GRATUIDAD_MATRICULA', 'ESTRATO', 'TIPOIDENTIFICACION','CREDITOSPENSUM', 'CREDITOSAPROBADOS',
    'CREDITOSMATRICULADOS','SUPERACREDITOS', 'FACULTAD', 'PROGRAMA', 'PRIMERNOMBRE','SEGUNDONOMBRE', 'PRIMERAPELLIDO', 'SEGUNDOAPELLIDO', 'GENERO',
    'ZONARESIDENCIA', 'IDMUNICIPIOPROGRAMA', 'NACIMIENTO','ID_PAIS_NACIMIENTO', 'IDMUNICIPIONACIMIENTO', 'ESTUDIANTEREINGRESO','ANIOINGRESO',
    'PERIODOINGRESO', 'TELEFONO', 'CELULAR', 'EMAILPERSONAL','EMAILINSTITUCIONAL', 'PUEBLOINDIGENA', 'COMUNIDADNEGRA', 'GRUPOSISBEN',
    'FONDOICETEX', 'RESOLUCIONICETEX', 'VALORGIROICETEX','BRUTA', 'BRUTAORD', 'NETAORD', 'MERITO', 'MTRNETA', 'NETAAPL','FL_NETA',
    'Estado Actual', 'Valor Factura','Valor Ajuste', 'Valor Pagado', 'Valor Anulado', 'Saldo', 'ESTADO CIVF','ESTADO FINANCIERO BENEFICIO',
    'ESTADO FINANCIERO FSE','FSE APLICAR', 'FSE REINTEGRAR','ESTADO FINANCIERO ICETEX','ICETEX RESOLUCION', 'ICETEX GIRO','ICETEX REINTEGRO',
    'ICETEX APLICAR','RESULTADO DE VALIDACION', 'GRADO PREVIO', 'Nombre de Destino']

df_piam20241_Cidl = df_piam20241_Cidl[columnas_df_piam20241_Cidl]
df_piam20241_Cidl.columns = df_piam20241_Cidl.columns.str.strip()

# Identifica los registros que estan ene l insumo de facturacion pero no en el de darca
# Cruza los DataFrames Piam 2024-1 Darca actualizado y Facturación a partir de la referencia de la factura y actualiza los registros
df_piam20241_dfr = pd.merge(
    df_piam20241_dfi,
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
    how='right')

df_piam20241_dfr['Estado Actual_x'] = df_piam20241_dfr['Estado Actual_y']
df_piam20241_dfr['Nombre de Destino_x'] = df_piam20241_dfr['Nombre de Destino_y']
df_piam20241_dfr['Valor Factura_x'] = df_piam20241_dfr['Valor Factura_y']
df_piam20241_dfr['Valor Ajuste_x'] = df_piam20241_dfr['Valor Ajuste_y']
df_piam20241_dfr['Valor Pagado_x'] = df_piam20241_dfr['Valor Pagado_y']
df_piam20241_dfr['Valor Anulado_x'] = df_piam20241_dfr['Valor Anulado_y']
df_piam20241_dfr['Saldo_x'] = df_piam20241_dfr['Saldo_y']
df_piam20241_dfr.drop(columns=['Valor Factura_y', 'Valor Ajuste_y', 'Valor Pagado_y', 'Valor Anulado_y', 'Saldo_y','Nombre de Destino_y','Estado Actual_y'], inplace=True)

# Identificación de los registros del dataframe Financiero que no estan en el DataFrame Academico
# Selección de columnas específicas requeridas
registrosUnicosSq = df_piam20241_dfr[~df_piam20241_dfr['Documento'].isin(df_piam20241_dfi['RECIBO'])]
columnas_registrosunicosSq = [
    'Documento',
    'Id  factura',
    'Estado Actual_x',
    'Valor Factura_x',
    'Valor Ajuste_x',
    'Valor Pagado_x',
    'Valor Anulado_x',
    'Saldo_x',
    'Nombre de Destino_x'
]

registrosUnicosSq = registrosUnicosSq[columnas_registrosunicosSq]
registrosUnicosSq.rename(
    columns={
        'Nombre de Destino_x':'Nombre de Destino',
        'Estado Actual_x':'Estado Actual',
        'Valor Factura_x':'Valor Factura',
        'Valor Ajuste_x':'Valor Ajuste',
        'Valor Pagado_x':'Valor Pagado',
        'Valor Anulado_x':'Valor Anulado',
        'Saldo_x':'Saldo'
        },
    inplace=True)
registrosUnicosSq.columns = registrosUnicosSq.columns.str.strip()


# Identifica los registros duplicados según el id del RECIDO en el dataframe Darca validado
# Identifica los registros duplicados según el Tercero y el Id del programa  en el dataframe Darca validado
df_piam20241_Cidl_SinDuplicadosXBoleta = df_piam20241_Cidl.drop_duplicates(subset='RECIBO', keep='first')
duplicados = df_piam20241_Cidl.duplicated(subset='RECIBO', keep=False)
registrosDuplicadosXbOLETA = df_piam20241_Cidl[duplicados]

# Guarda los dataframe cruzados segun el insumo academico y financiero
output_path = "/content/PIAM_2024_1_Conciliacion.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

  df_piam20241_Cifl.to_excel(writer, sheet_name='PIAM20241Ci', index=False)
  df_piam20241_Cidl.to_excel(writer, sheet_name='PIAM20241Darca', index=False)
  registrosUnicosSq.to_excel(writer, sheet_name='RegistrosUnicosSq', index=False)
  registrosDuplicadosXbOLETA.to_excel(writer, sheet_name='PIAM20241DuplicadosBol', index=False)

  if not df_piam20241_dfi_difNeta.empty:
    df_piam20241_dfi_difNeta.to_excel(writer, sheet_name='DifNetaconSq', index=False)
    print('Registros con diferentes valores de matricula Neta (Simca - Squid) agregados satisfactoriamente')
  else:
    print('No hay registros con diferentes valores de matricula Neta (Simca - Squid)')

print(f"Archivo guardado en {output_path}")
