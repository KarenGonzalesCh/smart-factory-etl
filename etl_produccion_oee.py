import pandas as pd
import numpy as np
from datetime import datetime
import gspread
import os
import json
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe

# 1. Configurar la API de Google Drive y autenticación

scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

if "GOOGLE_CREDENTIALS_JSON" in os.environ:
  # Carga las credenciales directamente desde el texto secreto en GitHub
  creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
  creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
else:
  # Carga local desde tu archivo de credenciales de siempre
  creds = Credentials.from_service_account_file("credentials.json", scopes=scope)

client = gspread.authorize(creds)
gc = client


# 2. Enlaces oficiales de las Hojas de Google Sheets
url_app_iny = 'https://docs.google.com/spreadsheets/d/1k0KyYQzkPanwmkI7Sx4GUHwdzR9IeXcduUudGCqv8Ks/edit?gid=80104312#gid=80104312'
url_ren_disp = 'https://docs.google.com/spreadsheets/d/1MJ28yOstYa2XMzEuuI96F7ezi5myjTzxrb1fYAGPiZg/edit?gid=0#gid=0'
url_historia_registros = 'https://docs.google.com/spreadsheets/d/1SC9XlZJJbiNoOPcsbtq2NKCsdm29AyHpZOb-lOJGEMg/edit?gid=133778390#gid=133778390'
info_pieza = 'https://docs.google.com/spreadsheets/d/169u1bkHJ2sm26haqikejClyTEl0DoXQD7tVkkXkxMWw/edit?gid=1087534904#gid=1087534904'
url_ord_historicos = 'https://docs.google.com/spreadsheets/d/1wMpCOVaCp-6YL52nTCB6C-_yw8xkqNLOuYg58rKCBb8/edit?gid=1728967839#gid=1728967839'

# Enlaces específicos para sensores (LOT / FAL) y estándares
url_sensores = 'https://docs.google.com/spreadsheets/d/1729eD8BMbEWreyR0j-d0X7vepAByTvzE_yurYnbLyLA/edit?usp=sharing'
url_sensores_FAL = 'https://docs.google.com/spreadsheets/d/1HywiwBs3LifJb5k2n9V1Uy-btFQ4mvB-cLKPmVormpI/edit?usp=sharing'
url_estandares = 'https://docs.google.com/spreadsheets/d/1YPnahpGLbi7ETGoR8hRwE2SZsd_NR8zSeE4wti3_P18/edit?gid=1601508775#gid=1601508775'
url_estandares_nuevo = 'https://docs.google.com/spreadsheets/d/1YPnahpGLbi7ETGoR8hRwE2SZsd_NR8zSeE4wti3_P18/edit?gid=1347610677#gid=1347610677'

# 3. Abrir las hojas de trabajo (Worksheets) principales
wrks_app_iny = gc.open_by_url(url_app_iny).worksheet('Base_inyeccion')
wrks_ren_disp = gc.open_by_url(url_ren_disp).worksheet('tabla')
wrks_info_pieza = gc.open_by_url(info_pieza).worksheet('Info Pieza')
wrks_historia_registros = gc.open_by_url(url_historia_registros).worksheet('FOR45')
wrks_ordenes_historicos = gc.open_by_url(url_ord_historicos).worksheet('Sheet1')

# Abrir hojas de sensores y estándares (PMP / LOT / FAL)
wrks_sensores = gc.open_by_url(url_sensores).worksheet('LOT_historico')
wrks_sensores_FAL = gc.open_by_url(url_sensores_FAL).worksheet('FAL_prueba')
wrk_estandares = gc.open_by_url(url_estandares).worksheet('Estandar')
wrk_estandares_nuevo = gc.open_by_url(url_estandares_nuevo).worksheet('EstandarNuevo')

# ==========================================
# 4. CARGA INICIAL DE DATOS DESDE GOOGLE SHEETS
# ==========================================
datos_app_iny = wrks_app_iny.get_all_values()
ws_sensores = wrks_sensores.get_all_values()
ws_sensores_FAL = wrks_sensores_FAL.get_all_values()
ws_estandares_nuevo = wrk_estandares_nuevo.get_all_values()

# ==========================================
# 5. DEFINICIÓN DE TODOS LOS DATAFRAMES CRUDOS (RAW)
# ==========================================
df_app_iny = pd.DataFrame(datos_app_iny[1:], columns=datos_app_iny[0]) if len(datos_app_iny) > 1 else pd.DataFrame()
df_sensores = pd.DataFrame(ws_sensores[1:], columns=ws_sensores[0])
df_sensores_FAL = pd.DataFrame(ws_sensores_FAL[1:], columns=ws_sensores_FAL[0])
df_estandares_nuevo = pd.DataFrame(ws_estandares_nuevo[1:], columns=ws_estandares_nuevo[0])



# Constantes de tiempo
TOPE_9_HORAS = 9 * 3600  # 32400 segundos
MIN_2_HORAS = 0
FILTRO_7_HORAS = 6.5 * 3600 # 23400 segundos

# ==========================================
# 1. FUNCIONES DE PROCESAMIENTO (LOT y FAL)
# ==========================================
def procesar_lot_con_id(df_lot):
    df = df_lot.copy()

    # Limpieza de espacios en blanco en strings
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()

    # Conversión a numérico
    df['Cantidad Real'] = pd.to_numeric(df['Cantidad Real'], errors='coerce').fillna(0)
    df['Tiempo Neto'] = pd.to_numeric(df['Tiempo Neto'], errors='coerce').fillna(0)

    # --- FILTROS PREVIOS ---
    df_filtrado = df[
        (df['Cantidad Real'] > 0) &
        (df['Tiempo Neto'] > MIN_2_HORAS)
    ].copy()

    # Normalizar Fecha (formato YYYY-MM-DD para compatibilidad)
    df_filtrado['Fecha_dt'] = pd.to_datetime(
        df_filtrado['Fecha Inicio'],
        format='%d/%m/%Y',
        errors='coerce'
    ).dt.date

    # --- CREAR ID ANTES DE AGRUPAR ---
    df_filtrado['ID'] = (
        df_filtrado['Fecha_dt'].astype(str).str.strip() + '_' +
        df_filtrado['Maquina'].astype(str).str.strip() + '_' +
        df_filtrado['Pieza'].astype(str).str.strip()
    )

    # Asegurar columnas numéricas para sumar
    columnas_sumar = [
        'Cantidad Real', 'Tiempo Neto', 'Tiempo de Proceso',
        'Paradas Planificadas', 'Paradas No Planificadas',
        'Paradas Cortas', 'Setup Real'
    ]
    cols_existentes = [c for c in columnas_sumar if c in df_filtrado.columns]
    for c in cols_existentes:
        df_filtrado[c] = pd.to_numeric(df_filtrado[c], errors='coerce').fillna(0)

    # --- AGRUPACIÓN USANDO EL ID Y LA FECHA ---
    df_lot_agrupado = df_filtrado.groupby(['ID', 'Fecha_dt', 'Maquina', 'Pieza'], as_index=False)[cols_existentes].sum()
    df_lot_agrupado.rename(columns={'Fecha_dt': 'Fecha'}, inplace=True)

    # --- CONTROL DE TOPE MÁXIMO DE 9 HORAS (32400 s) PARA TIEMPO DE PROCESO EN LOT ---
    if 'Tiempo de Proceso' in df_lot_agrupado.columns:
        df_lot_agrupado['Tiempo de Proceso'] = np.minimum(
            df_lot_agrupado['Tiempo de Proceso'],
            TOPE_9_HORAS
        )

    return df_lot_agrupado



def procesar_fal_con_id(df_fal):
    df = df_fal.copy()
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()

    df['Fecha_dt'] = pd.to_datetime(
        df['Inicio'],
        format='mixed',
        dayfirst=True,
        errors='coerce'
    ).dt.date

    df['ID'] = (
        df['Fecha_dt'].astype(str).str.strip() + '_' +
        df['Nombre Maquina'].astype(str).str.upper().str.strip().str.replace(' ', '', regex=False) + '_' +
        df['Receta'].astype(str).str.upper().str.strip().str.replace(' ', '', regex=False)
    )

    col_segundos = next((c for c in df.columns if 'segundo' in c.lower() or 'cantidad' in c.lower()), 'Cantidad de Segundos')
    df[col_segundos] = pd.to_numeric(df[col_segundos], errors='coerce').fillna(0)

    # =========================================================================
    # 🛑 FILTRO: Ignorar registros donde la cantidad de segundos sea >= FILTRO_7_HORAS
    # =========================================================================
    df = df[df[col_segundos] < FILTRO_7_HORAS].copy()
    if 'Descripcion Mensaje' in df.columns:
        df_dummies = pd.get_dummies(df['Descripcion Mensaje'], prefix='parada')
        df = pd.concat([df, df_dummies], axis=1)
        for col_d in df_dummies.columns:
            df[col_d] = df[col_d].astype(int) * df[col_segundos]

    cols_agg = {col_segundos: 'sum'}
    for col_d in [c for c in df.columns if c.startswith('parada_')]:
        cols_agg[col_d] = 'sum'

    df_fal_agrupado = df.groupby(['ID'], as_index=False).agg(cols_agg)
    return df_fal_agrupado


# ==========================================
# 2. EJECUCIÓN Y MERGE INICIAL (CON PARADAS / DUMMIES)
# ==========================================
ws_sensores = wrks_sensores.get_all_values()
ws_sensores_FAL = wrks_sensores_FAL.get_all_values()

df_sensores = pd.DataFrame(ws_sensores[1:], columns=ws_sensores[0])
df_sensores_FAL = pd.DataFrame(ws_sensores_FAL[1:], columns=ws_sensores_FAL[0])

# Procesar ambos DataFrames
df_lot_agrupado = procesar_lot_con_id(df_sensores)
df_fal_agrupado = procesar_fal_con_id(df_sensores_FAL)

# Detectar dinámicamente el nombre exacto de la columna de segundos en df_fal_agrupado
col_segundos_fal = next((c for c in df_fal_agrupado.columns if 'segundo' in c.lower() or 'cantidad' in c.lower()), None)

# Detectar automáticamente todas las columnas de paradas (dummies) generadas en FAL
cols_paradas_fal = [c for c in df_fal_agrupado.columns if c.startswith('parada_')]

if col_segundos_fal:
    columnas_a_fusionar = ['ID', col_segundos_fal] + cols_paradas_fal

    # Merge utilizando df_lot_agrupado como principal (left join) -> AQUÍ NACE df_lot_fal_agrupado
    df_lot_fal_agrupado = df_lot_agrupado.merge(
        df_fal_agrupado[columnas_a_fusionar],
        on='ID',
        how='left'
    )

    # Renombrar la columna principal de segundos a 'Tiempo Parada'
    df_lot_fal_agrupado.rename(columns={col_segundos_fal: 'Tiempo Parada'}, inplace=True)

    # Rellenar con 0 tanto el Tiempo Parada como todas las columnas de paradas (dummies)
    cols_a_rellenar_cero = ['Tiempo Parada'] + cols_paradas_fal
    for col in cols_a_rellenar_cero:
        if col in df_lot_fal_agrupado.columns:
            df_lot_fal_agrupado[col] = df_lot_fal_agrupado[col].fillna(0)
else:
    print("⚠️ Advertencia: No se encontró la columna de segundos en FAL. Verificá los nombres de las columnas en el DataFrame.")
    df_lot_fal_agrupado = df_lot_agrupado.copy()
    df_lot_fal_agrupado['Tiempo Parada'] = 0
    for col in cols_paradas_fal:
        df_lot_fal_agrupado[col] = 0


# ==========================================
# 3. RENOMBRAR Y LIMPIAR ESPACIOS DE PIEZA
# ==========================================
if 'Pieza' in df_lot_fal_agrupado.columns:
    df_lot_fal_agrupado.rename(columns={'Pieza': 'Codigo_Pieza'}, inplace=True)

if 'Codigo_Pieza' in df_lot_fal_agrupado.columns:
    df_lot_fal_agrupado['Codigo_Pieza'] = (
        df_lot_fal_agrupado['Codigo_Pieza']
        .astype(str)
        .str.strip()
        .str.replace(' ', '', regex=False)
    )


# ==========================================
# 4. PREPARACIÓN CORRECTA DE DF_ESTANDARES_UNICO (CON TODAS LAS FILAS)
# ==========================================
ws_estandares_nuevo = wrk_estandares_nuevo.get_all_values()
df_estandares_nuevo = pd.DataFrame(ws_estandares_nuevo[1:], columns=ws_estandares_nuevo[0])

col_maq_std = 'Máquina' if 'Máquina' in df_estandares_nuevo.columns else 'MÁQUINA'

# Crear CONC en el estándar
df_estandares_nuevo['CONC'] = (
    df_estandares_nuevo[col_maq_std].astype(str).str.strip().str.replace(' ', '', regex=False) +
    df_estandares_nuevo['CodigodePieza'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

# Eliminar duplicados basándose estrictamente en la clave única 'CONC'
df_estandares_unico = df_estandares_nuevo.drop_duplicates(subset=['CONC'], keep='first').copy()

# Limpieza profunda de CONC en el estándar
df_estandares_unico['CONC'] = (
    df_estandares_unico['CONC']
    .astype(str)
    .str.strip()
    .str.upper()
    .str.replace(r'\s+', '', regex=True)
    .str.replace(r'[\s\xa0\u200b\u3000]', '', regex=True)
)

print(f"¡Estándar preparado con éxito! Total de filas únicas: {len(df_estandares_unico)}")


# ==========================================
# 5. ASEGURAR CONC EN DF_LOT_FAL_AGRUPADO Y LIMPIEZA PROFUNDA
# ==========================================
df_lot_fal_agrupado['CONC'] = (
    df_lot_fal_agrupado['Maquina'].astype(str).str.strip().str.replace(' ', '', regex=False) +
    df_lot_fal_agrupado['Codigo_Pieza'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

df_lot_fal_agrupado['CONC'] = (
    df_lot_fal_agrupado['CONC']
    .astype(str)
    .str.strip()
    .str.upper()
    .str.replace(r'\s+', '', regex=True)
    .str.replace(r'[\s\xa0\u200b\u3000]', '', regex=True)
)


# ==========================================
# 6. EJECUTAR EL MERGE PARA TRAER GPH, MUL, DIV Y LA PIEZA DESDE EL ESTÁNDAR
# ==========================================
cols_a_remover = [c for c in ['GPH', 'MUL', 'DIV', 'Pieza'] if c in df_lot_fal_agrupado.columns]
if cols_a_remover:
    df_lot_fal_agrupado = df_lot_fal_agrupado.drop(columns=cols_a_remover)

cols_estandar_a_traer = ['CONC', 'GPH', 'MUL', 'DIV', 'Pieza']
df_estandar_para_merge = df_estandares_unico[cols_estandar_a_traer].drop_duplicates(subset=['CONC'], keep='first')

df_lot_fal_agrupado = df_lot_fal_agrupado.merge(
    df_estandar_para_merge,
    on='CONC',
    how='left'
)

# Validación de cruce exitoso
coincidencias = df_lot_fal_agrupado['GPH'].notna().sum()
print(f"¡Merge de estándares completado con éxito! Registros totales: {len(df_lot_fal_agrupado)} | Cruces exitosos con GPH: {coincidencias}")


# ==========================================
# 7. CREACIÓN DE ID_Inyeccion EN df_lot_fal_agrupado
# ==========================================
fecha_limpia_lot = (
    df_lot_fal_agrupado['Fecha']
    .astype(str)
    .str.split(' ')
    .str[0]
    .str.strip()
    .str.replace('/', '-', regex=False)
)

# Usamos la columna 'Pieza' traída desde el estándar nuevo para armar el ID_Inyeccion
df_lot_fal_agrupado['ID_Inyeccion'] = (
    fecha_limpia_lot + '_' +
    df_lot_fal_agrupado['Maquina'].astype(str).str.strip().str.replace(' ', '', regex=False) + '_' +
    df_lot_fal_agrupado['Pieza'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

print(f"¡Columna 'ID_Inyeccion' creada con éxito para {len(df_lot_fal_agrupado)} registros en df_lot_fal_agrupado!")


# ==========================================
# 8. CONVERSIÓN FINAL A HORAS (0.00) DE TIEMPO PARADA Y COLUMNAS PARADA_
# ==========================================
cols_paradas_a_convertir = [c for c in df_lot_fal_agrupado.columns if c.startswith('parada_')]
if 'Tiempo Parada' in df_lot_fal_agrupado.columns:
    cols_paradas_a_convertir.append('Tiempo Parada')

for col in cols_paradas_a_convertir:
    df_lot_fal_agrupado[col] = pd.to_numeric(df_lot_fal_agrupado[col], errors='coerce').fillna(0)
    df_lot_fal_agrupado[col] = np.where(
        df_lot_fal_agrupado[col] > 0,
        (df_lot_fal_agrupado[col] / 3600.0).round(2),
        0
    )

print(f"¡Conversión a horas (0.00) aplicada correctamente a {len(cols_paradas_a_convertir)} columnas de parada!")




# ==========================================
# . EXPORTAR df_lot_fal_agrupado A LA HOJA 'LOT_FAL'
# ==========================================
url_destino = "https://docs.google.com/spreadsheets/d/1FeyGKPpSkrv_jC3uE09JqDF9hyLOqtcC0aPqz3ba_uI/edit?gid=0#gid=0"

# Abrir el documento de destino
spreadsheet_destino = gc.open_by_url(url_destino)

# Seleccionar la pestaña específica llamada 'LOT_FAL' (si no existe, podés crearla o cambiar el nombre)
try:
    worksheet_destino = spreadsheet_destino.worksheet('LOT_FAL')
except gspread.exceptions.WorksheetNotFound:
    # Si por algún motivo no existe, la crea automáticamente
    worksheet_destino = spreadsheet_destino.add_worksheet(title='LOT_FAL', rows="1000", cols="30")

# Limpiar el contenido previo de la hoja
worksheet_destino.clear()

# Escribir el DataFrame con sus encabezados en la pestaña 'LOT_FAL'
set_with_dataframe(worksheet_destino, df_lot_fal_agrupado)

print(f"¡Exportación exitosa! Se subieron {len(df_lot_fal_agrupado)} filas a la pestaña 'LOT_FAL'.")



#####################################################################

# ==========================================
# 0. OBTENER DATAFRAME DE WRKS_APP_INY (Base_inyeccion)
# ==========================================
datos_app_iny = wrks_app_iny.get_all_values()

if len(datos_app_iny) > 1:
    df_app_iny = pd.DataFrame(datos_app_iny[1:], columns=datos_app_iny[0])
else:
    df_app_iny = pd.DataFrame()

print(f"¡DataFrame 'df_app_iny' cargado con éxito! Contiene {len(df_app_iny)} filas.")



# ==========================================
# ESTANDARIZAR NOMBRES DE MÁQUINAS EN DF_APP_INY
# ==========================================

mapeo_maquinas = {
    'INYECTORA 1': 'INY 1',
    'INYECTORA 2': 'INY 2',
    'INYECTORA 3': 'INY 3',
    'INYECTORA 4': 'INY 4',
    'INYECTORA 5': 'INY 5'
}

# Detectar automáticamente si la columna se llama 'MÁQUINA' o 'Maquina'
col_maquina_app = next((c for c in df_app_iny.columns if c.strip().upper() in ['MÁQUINA', 'MAQUINA']), None)

if col_maquina_app:
    df_app_iny['Maquina'] = (
        df_app_iny[col_maquina_app]
        .astype(str)
        .str.strip()
        .replace(mapeo_maquinas)
    )
    print(f"¡Estandarización de máquinas completada usando la columna '{col_maquina_app}'!")
else:
    print("⚠️ No se encontró la columna de máquina en df_app_iny. Verificá los nombres de las columnas.")

# ==========================================
# CREACIÓN DE ID_Inyeccion CON FECHA EN FORMATO AAAA-MM-DD
# ==========================================

# 1. Convertir la fecha a formato datetime y formatearla como AAAA-MM-DD (YYYY-MM-DD)
fecha_formateada = (
    pd.to_datetime(df_app_iny['FECHA'], errors='coerce')
    .dt.strftime('%Y-%m-%d')
)

# 2. Construir el ID_Inyeccion concatenando Fecha (en formato YYYY-MM-DD), Maquina y Pieza sin espacios
df_app_iny['ID_Inyeccion'] = (
    fecha_formateada + '_' +
    df_app_iny['Maquina'].astype(str).str.strip().str.replace(' ', '', regex=False) + '_' +
    df_app_iny['PIEZA'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

print(f"¡Columna 'ID_Inyeccion' formateada con éxito en formato YYYY-MM-DD para {len(df_app_iny)} registros!")

# ==========================================
# MERGE DE df_app_iny Y df_lot_fal_agrupado USANDO ID_Inyeccion
# ==========================================

# 1. Definir las columnas a traer de df_lot_fal_agrupado, excluyendo 'Fecha' para evitar duplicados
cols_a_traer = [c for c in df_lot_fal_agrupado.columns if c != 'Fecha']

# Si 'ID_Inyeccion' no está en cols_a_traer por seguridad la agregamos
if 'ID_Inyeccion' not in cols_a_traer:
    cols_a_traer.append('ID_Inyeccion')

# 2. Realizar el merge (left join usando df_app_iny como base principal)
df_app_iny_unificado = df_app_iny.merge(
    df_lot_fal_agrupado[cols_a_traer],
    on='ID_Inyeccion',
    how='left',
    suffixes=('', '_lot')  # Evita colisiones de nombres si hay columnas repetidas
)

print(f"¡Merge completado con éxito! El DataFrame resultante 'df_app_iny_unificado' tiene {len(df_app_iny_unificado)} filas y {len(df_app_iny_unificado.columns)} columnas.")


################################################################################


from datetime import datetime
import pandas as pd
import numpy as np
from zoneinfo import ZoneInfo
import gspread
import time


# 📌 1. CALCULAR DINÁMICAMENTE LOS ÚLTIMOS 2 DÍAS HÁBILES SEGÚN LA HORA DE ARGENTINA
hora_argentina = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
hora_actual = hora_argentina.hour
fecha_base = hora_argentina.date()

if hora_actual < 17:
    fecha_tope = pd.bdate_range(end=fecha_base - pd.Timedelta(days=1), periods=1)[-1].date()
else:
    fecha_tope = fecha_base

dias_habiles_previos = pd.bdate_range(end=fecha_tope, periods=2)
fechas_esperadas = dias_habiles_previos.date

print(f"📅 Rango dinámico seleccionado (Hora ARG: {hora_argentina.strftime('%H:%M')} hs): {[str(d) for d in fechas_esperadas]}\n")

tablas_ac_b_h_ext = {}
# Orden estricto que se respetará sí o sí
maqs_ac_b_h_ext = ['B 1', 'B 2', 'B 3', 'B 4', 'B 5', 'AC 1', 'AC 2', 'AC 3', 'AC 4', 'AC 7', 'AC 10', 'AC 11', 'H1', 'Extrusora']

print("🚀 Procesando máquinas AC y B desde df_lot_fal_agrupado en orden estricto...\n")

if 'df_lot_fal_agrupado' in globals() and df_lot_fal_agrupado is not None and len(df_lot_fal_agrupado) > 0:
    df_trabajo = df_lot_fal_agrupado.copy().reset_index(drop=True)
    df_trabajo = df_trabajo.loc[:, ~df_trabajo.columns.duplicated()]

    # Asegurar formato datetime en la fecha
    if 'Fecha' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['Fecha'], format='mixed', dayfirst=True, errors='coerce')
    elif 'Fecha Inicio' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['Fecha Inicio'], format='mixed', dayfirst=True, errors='coerce')
    elif 'FECHA' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['FECHA'], format='mixed', dayfirst=True, errors='coerce')

    # Columnas numéricas
    cols_numericas = ['Cantidad Real', 'Tiempo Neto', 'Tiempo de Proceso', 'Tiempo Parada', 'GPH', 'MUL', 'DIV']
    for col in cols_numericas:
        if col in df_trabajo.columns:
            df_trabajo[col] = pd.to_numeric(df_trabajo[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    idx_fijo_date = [pd.to_datetime(d).date() for d in fechas_esperadas]
    dias_semana = {0: 'L', 1: 'M', 2: 'X', 3: 'J', 4: 'V'}
    nombres_columnas = [f"{pd.to_datetime(d).day}/{pd.to_datetime(d).month:02d} ({dias_semana.get(pd.to_datetime(d).dayofweek, '')})" for d in idx_fijo_date]

    # Recorrer estrictamente la lista en el orden indicado
    for maquina in maqs_ac_b_h_ext:
        df_maq = df_trabajo[df_trabajo['Maquina'].astype(str).str.strip() == maquina].copy()

        print(f"==================================================")
        print(f"--- MÁQUINA: {maquina} ---")
        print(f"==================================================")

        if df_maq.empty or 'FECHA' not in df_maq.columns:
            # Si no hay datos, crear estructura vacía con guiones
            df_final = pd.DataFrame(index=[
                'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
            ], columns=nombres_columnas)
            df_final[:] = "-"
        else:
            df_maq['FECHA'] = pd.to_datetime(df_maq['FECHA'], errors='coerce')
            df_maq = df_maq.dropna(subset=['FECHA'])

            if df_maq.empty:
                df_final = pd.DataFrame(index=[
                    'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                    'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
                ], columns=nombres_columnas)
                df_final[:] = "-"
            else:
                agg_dict = {
                    'Cantidad Real': 'sum',
                    'Tiempo Neto': 'sum',
                    'Tiempo de Proceso': 'sum',
                    'Tiempo Parada': 'sum',
                    'GPH': 'mean',
                    'MUL': 'mean'
                }

                df_diario = df_maq.groupby(df_maq['FECHA'].dt.date).agg(agg_dict)
                df_diario.index = pd.to_datetime(df_diario.index).date
                df_habil = df_diario[pd.to_datetime(df_diario.index).dayofweek < 5].copy()

                df_ultimos_2 = df_habil.reindex(idx_fijo_date).copy()

                df_ultimos_2['Tiempo Neto [h]'] = df_ultimos_2['Tiempo Neto'] / 3600
                df_ultimos_2['Tiempo de parada [h]'] = np.where(df_ultimos_2['Tiempo Parada'] > 24,
                                                               df_ultimos_2['Tiempo Parada'] / 3600,
                                                               df_ultimos_2['Tiempo Parada'])
                df_ultimos_2['Tiempo restante [h]'] = (9 - df_ultimos_2['Tiempo Neto [h]'] - df_ultimos_2['Tiempo de parada [h]'])

                df_ultimos_2['CANT. REAL'] = df_ultimos_2['Cantidad Real'] if 'Cantidad Real' in df_ultimos_2.columns else 0
                df_ultimos_2['CANT. NO CONFORMES'] = 0

                df_ultimos_2['Rendimiento'] = np.where(
                    (df_ultimos_2['GPH'].notna()) & (df_ultimos_2['GPH'] > 0) &
                    (df_ultimos_2['MUL'].notna()) & (df_ultimos_2['MUL'] > 0) &
                    (df_ultimos_2['Tiempo Neto'] > 0),
                    ((df_ultimos_2['CANT. REAL'] / df_ultimos_2['MUL']) / (df_ultimos_2['Tiempo Neto'] / 3600)) / df_ultimos_2['GPH'],
                    0
                )

                df_ultimos_2['Disponibilidad'] = np.where(
                    df_ultimos_2['Tiempo de Proceso'] > 0,
                    df_ultimos_2['Tiempo Neto'] / df_ultimos_2['Tiempo de Proceso'],
                    0
                )

                df_ultimos_2['Calidad'] = 1.0
                df_ultimos_2['OEE'] = (df_ultimos_2['Disponibilidad'] * df_ultimos_2['Rendimiento'] * df_ultimos_2['Calidad'])

                cols_finales = [
                    'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                    'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
                ]

                cols_existentes = [c for c in cols_finales if c in df_ultimos_2.columns]
                df_tabla = df_ultimos_2[cols_existentes].transpose()
                df_tabla.columns = nombres_columnas

                df_ultimos_2['Disponibilidad'] *= 100
                df_ultimos_2['Calidad'] *= 100
                df_ultimos_2['Rendimiento'] *= 100
                df_ultimos_2['OEE'] *= 100

                df_final = df_tabla.astype(object)

                for col_m in ['OEE', 'Disponibilidad', 'Calidad', 'Rendimiento']:
                    if col_m in df_ultimos_2.columns:
                        vals = df_ultimos_2[col_m]
                        df_final.loc[col_m] = [f"{v:.2f}%" if pd.notna(v) else "-" for v in vals]

                for col_m in ['Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]']:
                    if col_m in df_final.index:
                        df_final.loc[col_m] = df_final.loc[col_m].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "-")

                for col_m in ['CANT. NO CONFORMES', 'CANT. REAL']:
                    if col_m in df_final.index:
                        df_final.loc[col_m] = df_final.loc[col_m].apply(
                            lambda x: f"{int(round(pd.to_numeric(x, errors='coerce')))}" if pd.notna(x) and pd.notna(pd.to_numeric(x, errors='coerce')) else "-"
                        )

        tablas_ac_b_h_ext[maquina] = df_final

        try:
            from IPython.display import display
            display(df_final)
        except:
            print(df_final)
        print("\n" + "-"*50 + "\n")
else:
    print("⚠️ El DataFrame 'df_lot_fal_agrupado' no está disponible o está vacío.")



#######################################################################################




# 📌 1. CALCULAR DINÁMICAMENTE LOS ÚLTIMOS 2 DÍAS HÁBILES SEGÚN LA HORA DE ARGENTINA
hora_argentina = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
hora_actual = hora_argentina.hour
fecha_base = hora_argentina.date()

if hora_actual < 17:
    fecha_tope = pd.bdate_range(end=fecha_base - pd.Timedelta(days=1), periods=1)[-1].date()
else:
    fecha_tope = fecha_base

dias_habiles_previos = pd.bdate_range(end=fecha_tope, periods=2)
fechas_esperadas = dias_habiles_previos.date

tablas_inyectoras = {}
# Orden estricto para las inyectoras
maqs_inyectoras = ['INY 2', 'INY 3', 'INY 4', 'INY 5']

print("🚀 Procesando Inyectoras desde df_app_iny_unificado en orden estricto...\n")

if 'df_app_iny_unificado' in globals() and df_app_iny_unificado is not None and len(df_app_iny_unificado) > 0:
    df_trabajo_iny = df_app_iny_unificado.copy().reset_index(drop=True)
    df_trabajo_iny = df_trabajo_iny.loc[:, ~df_trabajo_iny.columns.duplicated()]

    # Asegurar formato datetime en la fecha
    if 'Fecha' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['Fecha'], format='mixed', dayfirst=True, errors='coerce')
    elif 'Fecha Inicio' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['Fecha Inicio'], format='mixed', dayfirst=True, errors='coerce')
    elif 'FECHA' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['FECHA'], format='mixed', dayfirst=True, errors='coerce')

    # Columnas numéricas (incluyendo CANT. SUB DIARIA)
    cols_numericas_iny = ['Cantidad Real', 'CANT. SUB DIARIA', 'Tiempo Neto', 'Tiempo de Proceso', 'Tiempo Parada', 'GPH', 'MUL', 'DIV']
    for col in cols_numericas_iny:
        if col in df_trabajo_iny.columns:
            df_trabajo_iny[col] = pd.to_numeric(df_trabajo_iny[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    idx_fijo_date = [pd.to_datetime(d).date() for d in fechas_esperadas]
    dias_semana = {0: 'L', 1: 'M', 2: 'X', 3: 'J', 4: 'V'}
    nombres_columnas = [f"{pd.to_datetime(d).day}/{pd.to_datetime(d).month:02d} ({dias_semana.get(pd.to_datetime(d).dayofweek, '')})" for d in idx_fijo_date]

    # Recorrer estrictamente la lista en el orden indicado
    for maquina in maqs_inyectoras:
        df_maq = df_trabajo_iny[df_trabajo_iny['Maquina'].astype(str).str.strip() == maquina].copy()

        print(f"==================================================")
        print(f"--- MÁQUINA: {maquina} ---")
        print(f"==================================================")

        if df_maq.empty or 'FECHA' not in df_maq.columns:
            df_final = pd.DataFrame(index=[
                'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
            ], columns=nombres_columnas)
            df_final[:] = "-"
        else:
            df_maq['FECHA'] = pd.to_datetime(df_maq['FECHA'], errors='coerce')
            df_maq = df_maq.dropna(subset=['FECHA'])

            if df_maq.empty:
                df_final = pd.DataFrame(index=[
                    'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                    'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
                ], columns=nombres_columnas)
                df_final[:] = "-"
            else:
                agg_dict_iny = {
                    'Cantidad Real': 'sum',
                    'Tiempo Neto': 'sum',
                    'Tiempo de Proceso': 'sum',
                    'Tiempo Parada': 'sum',
                    'GPH': 'mean',
                    'MUL': 'mean'
                }
                if 'CANT. SUB DIARIA' in df_maq.columns:
                    agg_dict_iny['CANT. SUB DIARIA'] = 'sum'

                df_diario = df_maq.groupby(df_maq['FECHA'].dt.date).agg(agg_dict_iny)
                df_diario.index = pd.to_datetime(df_diario.index).date
                df_habil = df_diario[pd.to_datetime(df_diario.index).dayofweek < 5].copy()

                df_ultimos_2 = df_habil.reindex(idx_fijo_date).copy()

                df_ultimos_2['Tiempo Neto [h]'] = df_ultimos_2['Tiempo Neto'] / 3600
                df_ultimos_2['Tiempo de parada [h]'] = np.where(df_ultimos_2['Tiempo Parada'] > 24,
                                                               df_ultimos_2['Tiempo Parada'] / 3600,
                                                               df_ultimos_2['Tiempo Parada'])
                df_ultimos_2['Tiempo restante [h]'] = (9 - df_ultimos_2['Tiempo Neto [h]'] - df_ultimos_2['Tiempo de parada [h]'])

                # Asignación específica de CANT. SUB DIARIA para inyectoras
                df_ultimos_2['CANT. REAL'] = df_ultimos_2['CANT. SUB DIARIA'] if 'CANT. SUB DIARIA' in df_ultimos_2.columns else 0
                df_ultimos_2['CANT. NO CONFORMES'] = 0

                df_ultimos_2['Rendimiento'] = np.where(
                    (df_ultimos_2['GPH'].notna()) & (df_ultimos_2['GPH'] > 0) &
                    (df_ultimos_2['MUL'].notna()) & (df_ultimos_2['MUL'] > 0) &
                    (df_ultimos_2['Tiempo Neto'] > 0),
                    ((df_ultimos_2['CANT. REAL'] / df_ultimos_2['MUL']) / (df_ultimos_2['Tiempo Neto'] / 3600)) / df_ultimos_2['GPH'],
                    0
                )

                df_ultimos_2['Disponibilidad'] = np.where(
                    df_ultimos_2['Tiempo de Proceso'] > 0,
                    df_ultimos_2['Tiempo Neto'] / df_ultimos_2['Tiempo de Proceso'],
                    0
                )

                df_ultimos_2['Calidad'] = 1.0
                df_ultimos_2['OEE'] = (df_ultimos_2['Disponibilidad'] * df_ultimos_2['Rendimiento'] * df_ultimos_2['Calidad'])

                cols_finales = [
                    'OEE', 'Disponibilidad', 'Calidad', 'Rendimiento',
                    'CANT. NO CONFORMES', 'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'
                ]

                cols_existentes = [c for c in cols_finales if c in df_ultimos_2.columns]
                df_tabla = df_ultimos_2[cols_existentes].transpose()
                df_tabla.columns = nombres_columnas

                df_ultimos_2['Disponibilidad'] *= 100
                df_ultimos_2['Calidad'] *= 100
                df_ultimos_2['Rendimiento'] *= 100
                df_ultimos_2['OEE'] *= 100

                df_final = df_tabla.astype(object)

                for col_m in ['OEE', 'Disponibilidad', 'Calidad', 'Rendimiento']:
                    if col_m in df_ultimos_2.columns:
                        vals = df_ultimos_2[col_m]
                        df_final.loc[col_m] = [f"{v:.2f}%" if pd.notna(v) else "-" for v in vals]

                for col_m in ['Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]']:
                    if col_m in df_final.index:
                        df_final.loc[col_m] = df_final.loc[col_m].apply(lambda x: f"{x:.2f}" if pd.notna(x) else "-")

                for col_m in ['CANT. NO CONFORMES', 'CANT. REAL']:
                    if col_m in df_final.index:
                        df_final.loc[col_m] = df_final.loc[col_m].apply(
                            lambda x: f"{int(round(pd.to_numeric(x, errors='coerce')))}" if pd.notna(x) and pd.notna(pd.to_numeric(x, errors='coerce')) else "-"
                        )

        tablas_inyectoras[maquina] = df_final

        try:
            from IPython.display import display
            display(df_final)
        except:
            print(df_final)
        print("\n" + "-"*50 + "\n")
else:
    print("⚠️ El DataFrame 'df_app_iny_unificado' no está disponible o está vacío.")




#####################################################################################################################




spreadsheet_id = "1qMbb87guSbqPx3fd0tojIfTYuPnuN-r4I2mq8crqL58"
sheet_su = gc.open_by_key(spreadsheet_id)

# Mapeo de diccionarios con sus respectivas hojas de destino
diccionarios_hojas = {
    "INYECCION": tablas_inyectoras if 'tablas_inyectoras' in globals() else {},
    "B": {k: v for k, v in tablas_ac_b_h_ext.items() if k.startswith('B')} if 'tablas_ac_b_h_ext' in globals() else {},
    "AC": {k: v for k, v in tablas_ac_b_h_ext.items() if k.startswith('AC')} if 'tablas_ac_b_h_ext' in globals() else {},
    "H1": {k: v for k, v in tablas_ac_b_h_ext.items() if k == 'H1'} if 'tablas_ac_b_h_ext' in globals() else {},
    "Extrusora": {k: v for k, v in tablas_ac_b_h_ext.items() if k == 'Extrusora'} if 'tablas_ac_b_h_ext' in globals() else {}
}

# Procesar cada hoja de manera independiente
for nombre_hoja_unica, tablas_dict in diccionarios_hojas.items():
    if not tablas_dict:
        print(f"⚠️ No hay datos para la hoja '{nombre_hoja_unica}'.")
        continue

    try:
        worksheet = sheet_su.worksheet(nombre_hoja_unica)
        worksheet.clear()
        time.sleep(1.5) # Pausa tras limpiar la hoja
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sheet_su.add_worksheet(title=nombre_hoja_unica, rows=100, cols=15)
        time.sleep(1.5)

    fila_actual = 1

    for maquina, df_final in tablas_dict.items():
        try:
            # A. Título de la máquina
            titulo_maquina = [[f"--- {maquina} ---"]]
            worksheet.update(values=titulo_maquina, range_name=f"A{fila_actual}")
            fila_actual += 1

            # B. Preparar DataFrame
            df_para_subir = df_final.reset_index()
            df_para_subir.columns.values[0] = "Métrica"
            datos_a_escribir = [df_para_subir.columns.values.tolist()] + df_para_subir.values.tolist()

            # C. Escribir tabla principal
            worksheet.update(values=datos_a_escribir, range_name=f"A{fila_actual}")

            # D. Espaciado vertical
            fila_actual += len(datos_a_escribir) + 2
            print(f"✅ [{nombre_hoja_unica}] Tabla de '{maquina}' subida correctamente.")

            # 🛑 Pausa de 1.5 segundos para respetar los límites de la API de Google
            time.sleep(1.5)

        except Exception as e:
            print(f"❌ Error al procesar '{maquina}' en la hoja '{nombre_hoja_unica}': {e}")
            time.sleep(3) # Pausa más larga si llega a haber un tropiezo para permitir que se reinicie el contador

