import os
import pandas as pd

pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)

valoresValidosTipoId = ['CC', 'DE', 'CE', 'TI', 'PS', 'CA', 'PT']
estado_financiero = {
    'ac': 'Activa',
    'an': 'Anulada',
    'ca': 'Cancelada'
}
estado_beneficio = {
    'B': 'Beneficiario',
    'E': 'Excluido'
}

columnasValidacionObligatoriedad = [
    'TIPOIDENTIFICACION','IDENTIFICACION','CODIGO','SNIESPROGRAMA','IDMUNICIPIOPROGRAMA',
    'NACIMIENTO','ID_PAIS_NACIMIENTO','IDMUNICIPIONACIMIENTO','ZONARESIDENCIA','ESTRATO',
    'ESTUDIANTEREINGRESO','ANIOINGRESO','PERIODOINGRESO','EMAILINSTITUCIONAL','SEGURO_ESTUDIANTIL',
    'CREDITOSPENSUM','CREDITOSAPROBADOS']



# Verifica la existencia del archivo en la ruta especifica
file_path = '/content/Libro1.xlsx'

if not os.path.isfile(file_path):
    raise FileNotFoundError(f"{file_path} no encontrado.")
else:
    print(f"Archivo {file_path} encontrado.")

# Abre el archivo en modo binario para verificar problemas de acceso
try:
    with open(file_path, 'rb') as f:
        print(f"Archivo {file_path} abierto satisfactoriamente en modo binario.")
except OSError as e:
    print(f"Error al abrir el archivo {file_path}: {e}")



# Carga los DataFrames de trabajo
try:
    # Lectura de los insumos en un diccionario de dataframes
    dic_insumos = pd.read_excel(file_path, sheet_name=['IDARCA2106_24_1_ALL', 'SQ2106_24_1ALL', 'PIAM24_1_CI'], engine='openpyxl')

    # Limpia los nombres de columnas
    for df in dic_insumos.values():
        df.columns = df.columns.str.strip()

    piam20241, facturacion20241, PIAM20241CI = dic_insumos['IDARCA2106_24_1_ALL'], dic_insumos['SQ2106_24_1ALL'], dic_insumos['PIAM24_1_CI']
except Exception as e:
    print(f"Error al cargar los DataFrames: {e}")



# Analisis estructural
def validar_tipo_documento(df):
    if 'TIPOIDENTIFICACION' in df.columns:
        df_validacion_tipoId = df[~df['TIPOIDENTIFICACION'].isin(valoresValidosTipoId)]
        return df_validacion_tipoId
    else:
        return pd.DataFrame()


def obtenerRegistrosVacios(df,columnas):
  registrosVaciosTotal = pd.DataFrame()
  resumenVacios = {}

  for columna in columnas:
    if columna not in df.columns:
      print(f"La columna '{columna}' no existe en el DataFrame.")
      continue

    registrosVacios = df[df[columna].isnull()]
    resumenVacios[columna] = len(registrosVacios)

    if registrosVacios.empty:
      print(f"No hay registros vacíos en la columna '{columna}'.")
    else:
      print(f"Hay {len(registrosVacios)} registros vacíos en la columna '{columna}'.")
      registrosVacios['BanderaRegistrosVacios'] = columna
      registrosVaciosTotal = pd.concat([registrosVaciosTotal, registrosVacios])

  return registrosVaciosTotal, resumenVacios



def verificarInconsistenciasCreditos(df, columna):

  if columna not in df.columns:
    print(f"La columna '{columna}' no existe en el DataFrame.")
    return pd.DataFrame()

  df['BanderaCreditosRC'] = df[columna] < 15

  df_inconsistenciasRC = df[df['BanderaCreditosRC']]
  
  print(f"Se encontraron {len(df_inconsistenciasRC)} programas con inconsistencia en los Creditos exigidos por el Registro Calificado")
  
  return df_inconsistenciasRC



def verificarInconsistenciasCreditosCantidad(df, creditosRC, creditosAprobados):

  if creditosRC not in df.columns:
    print(f"La columna '{creditosRC}' no existe en el DataFrame.")
    return pd.DataFrame()

  if creditosAprobados not in df.columns:
    print(f"La columna '{creditosAprobados}' no existe en el DataFrame.")
    return pd.DataFrame()

  def evaluarInconsistenciaCreditos(row):
    if row[creditosRC] < row[creditosAprobados]:
      return 'Creditos RC menor a los creditos aprobados'
    elif row[creditosRC] == row[creditosAprobados]:
      return 'Creditos RC igual a los creditos aprobados'
    else:
      return None

  df['FlCreditosRCAprobados'] = df.apply(evaluarInconsistenciaCreditos, axis=1)
  df_inconsistenciasRCAprobados = df[df['FlCreditosRCAprobados'].notnull()]
  
  print(f"Se encontraron {len(df_inconsistenciasRCAprobados)} inconsistencias en los Creditos del RC y lo aprobados ")
  
  return df_inconsistenciasRCAprobados



# CONFRONTACION DE INSUMOS
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


# Cruza los DataFrames PIAM 2024-1 CI y Facturación a partir de la referencia de la factura
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

df_piam20241_Cifl.update(df_piam20241_Cifl.filter(regex='_y$').rename(columns=lambda x: x.replace('_y', '_x')))
df_piam20241_Cifl.drop(columns=df_piam20241_Cifl.filter(regex='_y$').columns, inplace=True)

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
df_piam20241_Cifl.columns = df_piam20241_Cifl.columns.str.strip()

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

# Elimina los registros duplicados a partir del Recibo
df_piam20241_Cidl_SinDuplicadosXBoleta = df_piam20241_Cidl.drop_duplicates(subset='RECIBO', keep='first')

# Validación de registros duplicados y únicos
duplicados_final = df_piam20241_Cidl[df_piam20241_Cidl.duplicated(subset='RECIBO', keep=False)]
unicos_final = duplicados_final.drop_duplicates(subset='RECIBO', keep='first')

# Elimina los registros duplicados a partir del ID - SNIES
df_piam20241_Cidl_SinDuplicadosXBoletaXIdSnies = df_piam20241_Cidl_SinDuplicadosXBoleta.sort_values(
    by=['IDENTIFICACION', 'SNIESPROGRAMA', 'ESTADO CIVF'],
    ascending=[True, True, False])

duplicadosIdSnies = df_piam20241_Cidl_SinDuplicadosXBoletaXIdSnies.duplicated(
    subset=['IDENTIFICACION', 'SNIESPROGRAMA'],
    keep=False)

df_duplicadosIdSnies = df_piam20241_Cidl_SinDuplicadosXBoletaXIdSnies[duplicadosIdSnies]

df_piam20241_Cidl_SinDuplicados = df_piam20241_Cidl_SinDuplicadosXBoletaXIdSnies.drop_duplicates(
    subset=['IDENTIFICACION', 'SNIESPROGRAMA'],
    keep='first')

eliminados = df_duplicadosIdSnies[~df_duplicadosIdSnies.index.isin(df_piam20241_Cidl_SinDuplicados.index)]

# Validación de los registros únicos en el DataFrame de Facturación que no estan en el DataFrame de Darca
registrosUnicosSq = facturacion20241[~facturacion20241['Documento'].isin(df_piam20241_Cidl_SinDuplicados['RECIBO'])]
registrosUnicosSq.rename(
    columns={
        'Nombre de Destino_x': 'Nombre de Destino',
        'Estado Actual_x': 'Estado Actual',
        'Valor Factura_x': 'Valor Factura',
        'Valor Ajuste_x': 'Valor Ajuste',
        'Valor Pagado_x': 'Valor Pagado',
        'Valor Anulado_x': 'Valor Anulado',
        'Saldo_x': 'Saldo'
    },
    inplace=True
)

registrosUnicosSq.columns = registrosUnicosSq.columns.str.strip()


# Identificacion de registros que esan el PIAM validado y no estan en el insumo DARCA auditado
df_piam20241_Cifl_SnAnuladas = df_piam20241_Cifl[df_piam20241_Cifl['Estado Actual']!='an']

df_piam20241_CiDarcaSnDuplicados_inner = pd.merge(
    df_piam20241_Cifl_SnAnuladas,
    df_piam20241_Cidl_SinDuplicados,
    right_on='RECIBO',
    left_on='FACTURA',
    how='inner')

# Identificar registros únicos del lado izquierdo 'Solo CI No Darca'
registros_unicos_df_izquierdo = df_piam20241_Cifl_SnAnuladas[~df_piam20241_Cifl_SnAnuladas['FACTURA'].isin(df_piam20241_Cidl_SinDuplicados['RECIBO'])]

# Identificar registros únicos del lado derecho 'No CI Solo Darca'
registros_unicos_df_derecho = df_piam20241_Cidl_SinDuplicados[~df_piam20241_Cidl_SinDuplicados['RECIBO'].isin(df_piam20241_Cifl_SnAnuladas['FACTURA'])]

# Actualiza las columnas con sufijos '_x' y elimina las columnas con sufijos '_y'
for col in df_piam20241_CiDarcaSnDuplicados_inner.columns:
    if col.endswith('_x'):
        df_piam20241_CiDarcaSnDuplicados_inner[col.replace('_x', '')] = df_piam20241_CiDarcaSnDuplicados_inner[col]
    elif col.endswith('_y'):
        df_piam20241_CiDarcaSnDuplicados_inner.drop(columns=[col], inplace=True)

# Elimina las columnas con sufijo '_x'
df_piam20241_CiDarcaSnDuplicados_inner.drop(columns=df_piam20241_CiDarcaSnDuplicados_inner.filter(regex='_x$').columns, inplace=True)

# Valida las columnas obligatorias
registrosVaciosTotal, resumenVacios = obtenerRegistrosVacios(df_piam20241_Cidl_SinDuplicados, columnasValidacionObligatoriedad)

# Valida la cantidad de creditos del RC por cada programa
registrosConRCErrados = verificarInconsistenciasCreditos(df_piam20241_Cidl_SinDuplicados, 'CREDITOSPENSUM')

# Valida la cantidad de creditos del RC de cada programa con respecto a los aporbados por cada estudainte
registrosInconsistenciasRCAprobados = verificarInconsistenciasCreditosCantidad(df_piam20241_Cidl_SinDuplicados, 'CREDITOSPENSUM', 'CREDITOSAPROBADOS')

# Filtros de generalidades
filtro_piam20241Ci = df_piam20241_Cifl.groupby('Estado Actual')['FACTURA'].size().reset_index(name='Poblacion')
filtro_Facturacion = facturacion20241.groupby('Estado Actual')['Documento'].size().reset_index(name='Poblacion')
filtro_DarcaFinal = df_piam20241_Cidl_SinDuplicados.groupby('Estado Actual')['RECIBO'].size().reset_index(name='Poblacion')
filtro_relacionDarcaAuditadoCISinAn = df_piam20241_CiDarcaSnDuplicados_inner.groupby('ESTADO CIVF')['FACTURA'].size().reset_index(name='Poblacion')
filtro_relacionDarcaAuditadoCISinAn = filtro_relacionDarcaAuditadoCISinAn.rename(columns={
    'ESTADO CIVF': 'Estado de beneficio validado'
})
filtro_UnicosSq = registrosUnicosSq.groupby('Estado Actual')['Documento'].size().reset_index(name='Poblacion')
filtro_registros_unicos_df_izquierdo = registros_unicos_df_izquierdo.groupby(['ESTADO CIVF','ESTADO FINANCIERO BENEFICIO'])['FACTURA'].size().reset_index(name='Poblacion')
filtro_registros_unicos_df_izquierdo = filtro_registros_unicos_df_izquierdo.rename(columns={
    'ESTADO CIVF': 'Estado de beneficio validado',
    'ESTADO FINANCIERO BENEFICIO': 'Estado de ejecución del beneficio'
})
filtro_registros_unicos_df_derecho = registros_unicos_df_derecho.groupby('Estado Actual')['RECIBO'].size().reset_index(name='Poblacion')
filtro_ColumnasConRegistrosVacios = registrosVaciosTotal.groupby('BanderaRegistrosVacios')['RECIBO'].size().reset_index(name='Poblacion')
filtro_ColumnasConRegistrosVacios = filtro_ColumnasConRegistrosVacios.rename(columns={
    'BanderaRegistrosVacios': 'Columnas de información inconsistentes'
})
filtro_RCErrado = registrosConRCErrados.groupby(['PROGRAMA','CREDITOSPENSUM'])['RECIBO'].size().reset_index(name='Poblacion')
filtro_RCErrado = filtro_RCErrado.rename(columns={
    'CREDITOSPENSUM': 'Creditos Registro Calificado'
})

filtro_RcAprobadosGeneral = registrosInconsistenciasRCAprobados.groupby('FlCreditosRCAprobados')['RECIBO'].size().reset_index(name='Poblacion')
filtro_RcAprobadosGeneral = filtro_RcAprobadosGeneral.rename(columns={
    'FlCreditosRCAprobados': 'Observacion Creditos RC - Aprobados'
})


filtro_RcAprobados = registrosInconsistenciasRCAprobados.groupby(['FlCreditosRCAprobados','PROGRAMA'])['RECIBO'].size().reset_index(name='Poblacion')
filtro_RcAprobados = filtro_RcAprobados.rename(columns={
    'FlCreditosRCAprobados': 'Observacion Creditos RC - Aprobados'
})

# Aplica las descripciones a la columna 'Estado Actual'
filtro_UnicosSq['Estado Actual'] = filtro_UnicosSq['Estado Actual'].replace(estado_financiero)
filtro_piam20241Ci['Estado Actual'] = filtro_piam20241Ci['Estado Actual'].replace(estado_financiero)
filtro_Facturacion['Estado Actual'] = filtro_Facturacion['Estado Actual'].replace(estado_financiero)
filtro_DarcaFinal['Estado Actual'] = filtro_DarcaFinal['Estado Actual'].replace(estado_financiero)
filtro_relacionDarcaAuditadoCISinAn['Estado de beneficio validado'] = filtro_relacionDarcaAuditadoCISinAn['Estado de beneficio validado'].replace(estado_beneficio)
filtro_registros_unicos_df_izquierdo['Estado de beneficio validado'] = filtro_registros_unicos_df_izquierdo['Estado de beneficio validado'].replace(estado_beneficio)
filtro_registros_unicos_df_derecho['Estado Actual'] = filtro_registros_unicos_df_derecho['Estado Actual'].replace(estado_beneficio)


# Guarda los dataframe cruzados segun el insumo academico y financiero
output_path = "/content/AUDITORIA_PIAM_2024_1_Conciliacion.xlsx"
with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:

# Obtiene el workbook y el worksheet

  filtro_Facturacion.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=1, index=False)
  filtro_DarcaFinal.to_excel(writer, sheet_name='Generalidades', startrow=1, startcol=4, index=False)
  filtro_piam20241Ci.to_excel(writer, sheet_name='Generalidades',  startrow=1, startcol=7, index=False)
  filtro_UnicosSq.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=10, index=False)
  filtro_relacionDarcaAuditadoCISinAn.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=13, index=False)
  filtro_registros_unicos_df_izquierdo.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=16, index=False)
  filtro_registros_unicos_df_derecho.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=20, index=False)
  filtro_ColumnasConRegistrosVacios.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=23, index=False)
  filtro_RCErrado.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=26, index=False)
  filtro_RcAprobadosGeneral.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=30, index=False)
  filtro_RcAprobados.to_excel(writer, sheet_name='Generalidades',startrow=1, startcol=33, index=False)


  workbook  = writer.book
  worksheet = writer.sheets['Generalidades']
  formato = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
  worksheet.merge_range('B1:C1', "INSUMO FACTURACION 2024-1",formato)
  worksheet.merge_range('E1:F1', "INSUMO DARCA 2024-1 AUDITADO",formato)
  worksheet.merge_range('H1:I1', "INSUMO PIAM 2024-1 VALIDADO MEN",formato)
  worksheet.merge_range('K1:L1', "REGISTROS SOLO SQUID 2024-1",formato)
  worksheet.merge_range('N1:O1', "REGISTROS DARCA CON ESTADO VALIDADO",formato)
  worksheet.merge_range('Q1:S1', "REGISTROS APROBADOS MEN NO DARCA",formato)
  worksheet.merge_range('U1:V1', "REGISTROS DARCA SIN APROBACION MEN",formato)
  worksheet.merge_range('X1:Y1', "INFORMACION OBLIGATORIA CON REGISTROS VACIOS",formato)
  worksheet.merge_range('AA1:AC1', "POBLACIÓN CON INCONSISTENCIA EN LOS CREDITOS DEL REGISTRO CALIFICADO",formato)
  worksheet.merge_range('AE1:AJ1', "POBLACIÓN CON INCONSISTENCIA EN LA RELACION DE CREDITOS DEL REGISTRO CALIFICADO Y LOS CREDITOS APROBADOS",formato)



  facturacion20241.to_excel(writer, sheet_name='Facturacion20241', index=False)
  df_piam20241_Cifl.to_excel(writer, sheet_name='PIAM20241CiMEN', index=False)
  df_piam20241_Cidl.to_excel(writer, sheet_name='PIAM20241Darca', index=False)
  df_piam20241_Cidl_SinDuplicados.to_excel(writer, sheet_name='PIAM20241DarcaSinDuplicados', index=False)
  registrosUnicosSq.to_excel(writer, sheet_name='RegistrosSoloSquid', index=False)
  df_piam20241_CiDarcaSnDuplicados_inner.to_excel(writer, sheet_name='PIAM20241DarcaMen', index=False)
  registros_unicos_df_izquierdo.to_excel(writer, sheet_name='PIAM20241SoloCI', index=False)
  registros_unicos_df_derecho.to_excel(writer, sheet_name='PIAM20241SoloDarca', index=False)
  


  if not df_piam20241_dfi_difNeta.empty:
    df_piam20241_dfi_difNeta.to_excel(writer, sheet_name='DifNetaconSq', index=False)
    print('Registros con diferentes valores de matricula Neta (Simca - Squid) agregados satisfactoriamente')
  else:
    print('No hay registros con diferentes valores de matricula Neta (Simca - Squid)')


  if not df_piam20241_Cidl_SinDuplicados.empty:
    df_no_validos = validar_tipo_documento(df_piam20241_Cidl_SinDuplicados)
    if not df_no_validos.empty:
        df_no_validos.to_excel(writer, sheet_name='RegistrosNoValidos', index=False)
        print('Se han guardado los registros no válidos')
    else:
        print('No se encontraron registros por tipo Id no válidos por tipo de identificacion')
  else:
    print('El DataFrame está vacío')


  if not df_piam20241_Cidl_SinDuplicados.empty:
    
    if not registrosVaciosTotal.empty:
      registrosVaciosTotal.to_excel(writer, sheet_name='RegistrosVacios', index=False)
    
    if not registrosConRCErrados.empty:
      registrosConRCErrados.to_excel(writer, sheet_name='RegistrosConRCErrados', index=False)
    
    if not registrosInconsistenciasRCAprobados.empty:
      registrosInconsistenciasRCAprobados.to_excel(writer, sheet_name='RegistrosRCAprobados', index=False)

    

print(f"Archivo guardado en {output_path}")
