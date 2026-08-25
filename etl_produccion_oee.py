import pandas as pd
import numpy as np
from datetime import datetime
import gspread
import os
import json
from google.oauth2.service_account import Credentials
from gspread_dataframe import set_with_dataframe
from zoneinfo import ZoneInfo
import time

# 1. Configurar la API y autenticación para GitHub Actions
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

if "GOOGLE_CREDENTIALS_JSON" in os.environ:
    creds_dict = json.loads(os.environ["GOOGLE_CREDENTIALS_JSON"])
    creds = Credentials.from_service_account_info(creds_dict, scopes=scope)
else:
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
TOPE_9_HORAS = 9 * 3600   # 32400 segundos
MIN_2_HORAS = 0
FILTRO_7_HORAS = 6.5 * 3600 # 23400 segundos

# ==========================================
# 1. FUNCIONES DE PROCESAMIENTO (LOT y FAL)
# ==========================================
def procesar_lot_con_id(df_lot):
    df = df_lot.copy()
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()

    df['Cantidad Real'] = pd.to_numeric(df['Cantidad Real'], errors='coerce').fillna(0)
    df['Tiempo Neto'] = pd.to_numeric(df['Tiempo Neto'], errors='coerce').fillna(0)

    df_filtrado = df[
        (df['Cantidad Real'] > 0) &
        (df['Tiempo Neto'] > MIN_2_HORAS)
    ].copy()

    df_filtrado['Fecha_dt'] = pd.to_datetime(
        df_filtrado['Fecha Inicio'],
        format='%d/%m/%Y',
        errors='coerce'
    ).dt.date

    df_filtrado['ID'] = (
        df_filtrado['Fecha_dt'].astype(str).str.strip() + '_' +
        df_filtrado['Maquina'].astype(str).str.upper().str.strip().str.replace(' ', '', regex=False) + '_' +
        df_filtrado['Pieza'].astype(str).str.upper().str.strip().str.replace(' ', '', regex=False)
    )

    columnas_sumar = [
        'Cantidad Real', 'Tiempo Neto', 'Tiempo de Proceso',
        'Paradas Planificadas', 'Paradas No Planificadas',
        'Paradas Cortas', 'Setup Real'
    ]
    cols_existentes = [c for c in columnas_sumar if c in df_filtrado.columns]
    for c in cols_existentes:
        df_filtrado[c] = pd.to_numeric(df_filtrado[c], errors='coerce').fillna(0)

    df_lot_agrupado = df_filtrado.groupby(['ID', 'Fecha_dt', 'Maquina', 'Pieza'], as_index=False)[cols_existentes].sum()
    df_lot_agrupado.rename(columns={'Fecha_dt': 'Fecha'}, inplace=True)

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

    df['Inicio_dt'] = pd.to_datetime(
        df['Inicio'],
        format='mixed',
        dayfirst=True,
        errors='coerce'
    )

    df = df[df['Inicio_dt'].dt.hour < 17].copy()

    df['Fecha_dt'] = df['Inicio_dt'].dt.date
    
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
# 2. EJECUCIÓN Y MERGE INICIAL
# ==========================================
ws_sensores = wrks_sensores.get_all_values()
ws_sensores_FAL = wrks_sensores_FAL.get_all_values()

df_sensores = pd.DataFrame(ws_sensores[1:], columns=ws_sensores[0])
df_sensores_FAL = pd.DataFrame(ws_sensores_FAL[1:], columns=ws_sensores_FAL[0])

df_lot_agrupado = procesar_lot_con_id(df_sensores)
df_fal_agrupado = procesar_fal_con_id(df_sensores_FAL)

col_segundos_fal = next((c for c in df_fal_agrupado.columns if 'segundo' in c.lower() or 'cantidad' in c.lower()), None)
cols_paradas_fal = [c for c in df_fal_agrupado.columns if c.startswith('parada_')]

if col_segundos_fal:
    columnas_a_fusionar = ['ID', col_segundos_fal] + cols_paradas_fal
    df_lot_fal_agrupado = df_lot_agrupado.merge(
        df_fal_agrupado[columnas_a_fusionar],
        on='ID',
        how='left'
    )
    df_lot_fal_agrupado.rename(columns={col_segundos_fal: 'Tiempo Parada'}, inplace=True)
    cols_a_rellenar_cero = ['Tiempo Parada'] + cols_paradas_fal
    for col in cols_a_rellenar_cero:
        if col in df_lot_fal_agrupado.columns:
            df_lot_fal_agrupado[col] = df_lot_fal_agrupado[col].fillna(0)
else:
    df_lot_fal_agrupado = df_lot_fal_agrupado.copy()
    df_lot_fal_agrupado['Tiempo Parada'] = 0
    for col in cols_paradas_fal:
        df_lot_fal_agrupado[col] = 0

if 'Pieza' in df_lot_fal_agrupado.columns:
    df_lot_fal_agrupado.rename(columns={'Pieza': 'Codigo_Pieza'}, inplace=True)

if 'Codigo_Pieza' in df_lot_fal_agrupado.columns:
    df_lot_fal_agrupado['Codigo_Pieza'] = (
        df_lot_fal_agrupado['Codigo_Pieza']
        .astype(str)
        .str.strip()
        .str.replace(' ', '', regex=False)
    )

ws_estandares_nuevo = wrk_estandares_nuevo.get_all_values()
df_estandares_nuevo = pd.DataFrame(ws_estandares_nuevo[1:], columns=ws_estandares_nuevo[0])
col_maq_std = 'Máquina' if 'Máquina' in df_estandares_nuevo.columns else 'MÁQUINA'

df_estandares_nuevo['CONC'] = (
    df_estandares_nuevo[col_maq_std].astype(str).str.strip().str.replace(' ', '', regex=False) +
    df_estandares_nuevo['CodigodePieza'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

df_estandares_unico = df_estandares_nuevo.drop_duplicates(subset=['CONC'], keep='first').copy()
df_estandares_unico['CONC'] = (
    df_estandares_unico['CONC']
    .astype(str)
    .str.strip()
    .str.upper()
    .str.replace(r'\s+', '', regex=True)
    .str.replace(r'[\s\xa0\u200b\u3000]', '', regex=True)
)

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

fecha_limpia_lot = (
    df_lot_fal_agrupado['Fecha']
    .astype(str)
    .str.split(' ')
    .str[0]
    .str.strip()
    .str.replace('/', '-', regex=False)
)

df_lot_fal_agrupado['ID_Inyeccion'] = (
    fecha_limpia_lot + '_' +
    df_lot_fal_agrupado['Maquina'].astype(str).str.strip().str.replace(' ', '', regex=False) + '_' +
    df_lot_fal_agrupado['Pieza'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

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

url_destino = "https://docs.google.com/spreadsheets/d/1FeyGKPpSkrv_jC3uE09JqDF9hyLOqtcC0aPqz3ba_uI/edit?gid=0#gid=0"
spreadsheet_destino = gc.open_by_url(url_destino)

try:
    worksheet_destino = spreadsheet_destino.worksheet('LOT_FAL')
except gspread.exceptions.WorksheetNotFound:
    worksheet_destino = spreadsheet_destino.add_worksheet(title='LOT_FAL', rows="1000", cols="30")

worksheet_destino.clear()
set_with_dataframe(worksheet_destino, df_lot_fal_agrupado)

# ==========================================
# OBTENER DATAFRAME DE WRKS_APP_INY
# ==========================================
datos_app_iny = wrks_app_iny.get_all_values()
if len(datos_app_iny) > 1:
    df_app_iny = pd.DataFrame(datos_app_iny[1:], columns=datos_app_iny[0])
else:
    df_app_iny = pd.DataFrame()

mapeo_maquinas = {
    'INYECTORA 1': 'INY 1',
    'INYECTORA 2': 'INY 2',
    'INYECTORA 3': 'INY 3',
    'INYECTORA 4': 'INY 4',
    'INYECTORA 5': 'INY 5'
}

col_maquina_app = next((c for c in df_app_iny.columns if c.strip().upper() in ['MÁQUINA', 'MAQUINA']), None)
if col_maquina_app:
    df_app_iny['Maquina'] = df_app_iny[col_maquina_app].astype(str).str.strip().replace(mapeo_maquinas)

fecha_formateada = pd.to_datetime(df_app_iny['FECHA'], errors='coerce').dt.strftime('%Y-%m-%d')
df_app_iny['ID_Inyeccion'] = (
    fecha_formateada + '_' +
    df_app_iny['Maquina'].astype(str).str.strip().str.replace(' ', '', regex=False) + '_' +
    df_app_iny['PIEZA'].astype(str).str.strip().str.replace(' ', '', regex=False)
)

# ==========================================
# 1. LIMPIEZA PREVIA Y CONVERSIÓN NUMÉRICA
# ==========================================
df_app_iny['FECHA'] = pd.to_datetime(df_app_iny['FECHA'], errors='coerce', dayfirst=True).dt.strftime('%Y-%m-%d')

cols_a_sumar = [
    'CANT. NO CONFORMES', 
    'CANT. DE GOLPES INYECCIÓN', 
    'CANT. BOCAS', 
    'CANT. SUB DIARIA', 
    'CANT SUB DIARIA REAL'
]

for col in cols_a_sumar:
    if col in df_app_iny.columns:
        df_app_iny[col] = pd.to_numeric(df_app_iny[col], errors='coerce').fillna(0)

# ==========================================
# 2. AGRUPAMIENTO SEGURO SIN DICCIONARIO
# ==========================================
columnas_suma_existentes = [c for c in cols_a_sumar if c in df_app_iny.columns]
columnas_first_existentes = [c for c in df_app_iny.columns if c != 'ID_Inyeccion' and c not in columnas_suma_existentes and c != 'ID OPERARIO']

df_parte_first = df_app_iny.groupby('ID_Inyeccion')[columnas_first_existentes].first()
df_parte_suma = df_app_iny.groupby('ID_Inyeccion')[columnas_suma_existentes].sum()

df_agrupado_base = pd.concat([df_parte_first, df_parte_suma], axis=1)

if 'ID OPERARIO' in df_app_iny.columns:
    df_operarios = df_app_iny.groupby('ID_Inyeccion')['ID OPERARIO'].apply(lambda x: list(x.unique()))
    df_agrupado_base['ID OPERARIO'] = df_operarios

df_app_iny = df_agrupado_base.reset_index()

print(f"¡Agrupación exitosa y blindada! Ahora df_app_iny tiene {len(df_app_iny)} filas únicas.")

# ==========================================
# 3. MERGE DE df_app_iny Y df_lot_fal_agrupado USANDO ID_Inyeccion
# ==========================================
cols_a_traer = [c for c in df_lot_fal_agrupado.columns if c != 'Fecha']
if 'ID_Inyeccion' not in cols_a_traer:
    cols_a_traer.append('ID_Inyeccion')

df_app_iny_unificado = df_app_iny.merge(
    df_lot_fal_agrupado[cols_a_traer],
    on='ID_Inyeccion',
    how='left',
    suffixes=('', '_lot')
)

print(f"¡Merge completado con éxito! El DataFrame resultante 'df_app_iny_unificado' tiene {len(df_app_iny_unificado)} filas y {len(df_app_iny_unificado.columns)} columnas.")


# ==========================================
# PROCESAMIENTO MÁQUINAS AC, B, H1, Extrusora
# ==========================================
hora_argentina = datetime.now(ZoneInfo("America/Argentina/Buenos_Aires"))
hora_actual = hora_argentina.hour
fecha_base = hora_argentina.date()

if hora_actual < 17:
    fecha_tope = pd.bdate_range(end=fecha_base - pd.Timedelta(days=1), periods=1)[-1].date()
else:
    fecha_tope = fecha_base

dias_habiles_previos = pd.bdate_range(end=fecha_tope, periods=2)
fechas_esperadas = dias_habiles_previos.date

tablas_ac_b_h_ext = {}
maqs_ac_b_h_ext = ['B 1', 'B 2', 'B 3', 'B 4', 'B 5', 'AC 1', 'AC 2', 'AC 3', 'AC 4', 'AC 7', 'AC 10', 'AC 11', 'H1', 'Extrusora']

if 'df_lot_fal_agrupado' in globals() and df_lot_fal_agrupado is not None and len(df_lot_fal_agrupado) > 0:
    df_trabajo = df_lot_fal_agrupado.copy().reset_index(drop=True)
    df_trabajo = df_trabajo.loc[:, ~df_trabajo.columns.duplicated()]

    if 'Fecha' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['Fecha'], format='mixed', dayfirst=True, errors='coerce')
    elif 'Fecha Inicio' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['Fecha Inicio'], format='mixed', dayfirst=True, errors='coerce')
    elif 'FECHA' in df_trabajo.columns:
        df_trabajo['FECHA'] = pd.to_datetime(df_trabajo['FECHA'], format='mixed', dayfirst=True, errors='coerce')

    cols_numericas = ['Cantidad Real', 'Tiempo Neto', 'Tiempo de Proceso', 'Tiempo Parada', 'GPH', 'MUL', 'DIV']
    for col in cols_numericas:
        if col in df_trabajo.columns:
            df_trabajo[col] = pd.to_numeric(df_trabajo[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    idx_fijo_date = [pd.to_datetime(d).date() for d in fechas_esperadas]
    dias_semana = {0: 'L', 1: 'M', 2: 'X', 3: 'J', 4: 'V'}
    nombres_columnas = [f"{pd.to_datetime(d).day}/{pd.to_datetime(d).month:02d} ({dias_semana.get(pd.to_datetime(d).dayofweek, '')})" for d in idx_fijo_date]

    for maquina in maqs_ac_b_h_ext:
        df_maq = df_trabajo[df_trabajo['Maquina'].astype(str).str.strip() == maquina].copy()
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

# ==========================================
# PROCESAMIENTO INYECTORAS
# ==========================================
tablas_inyectoras = {}
maqs_inyectoras = ['INY 2', 'INY 3', 'INY 4', 'INY 5']

if 'df_app_iny_unificado' in globals() and df_app_iny_unificado is not None and len(df_app_iny_unificado) > 0:
    df_trabajo_iny = df_app_iny_unificado.copy().reset_index(drop=True)
    df_trabajo_iny = df_trabajo_iny.loc[:, ~df_trabajo_iny.columns.duplicated()]

    if 'Fecha' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['Fecha'], format='mixed', dayfirst=True, errors='coerce')
    elif 'Fecha Inicio' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['Fecha Inicio'], format='mixed', dayfirst=True, errors='coerce')
    elif 'FECHA' in df_trabajo_iny.columns:
        df_trabajo_iny['FECHA'] = pd.to_datetime(df_trabajo_iny['FECHA'], format='mixed', dayfirst=True, errors='coerce')

    cols_numericas_iny = ['Cantidad Real', 'CANT. SUB DIARIA','CANT. NO CONFORMES', 'Tiempo Neto', 'Tiempo de Proceso', 'Tiempo Parada', 'GPH', 'MUL', 'DIV']
    for col in cols_numericas_iny:
        if col in df_trabajo_iny.columns:
            df_trabajo_iny[col] = pd.to_numeric(df_trabajo_iny[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    for maquina in maqs_inyectoras:
        df_maq = df_trabajo_iny[df_trabajo_iny['Maquina'].astype(str).str.strip() == maquina].copy()
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

                if 'CANT. NO CONFORMES' in df_maq.columns:
                    agg_dict_iny['CANT. NO CONFORMES'] = 'sum'
                
                df_diario = df_maq.groupby(df_maq['FECHA'].dt.date).agg(agg_dict_iny)
                df_diario.index = pd.to_datetime(df_diario.index).date
                df_habil = df_diario[pd.to_datetime(df_diario.index).dayofweek < 5].copy()
                df_ultimos_2 = df_habil.reindex(idx_fijo_date).copy()

                df_ultimos_2['Tiempo Neto [h]'] = df_ultimos_2['Tiempo Neto'] / 3600
                df_ultimos_2['Tiempo de parada [h]'] = np.where(df_ultimos_2['Tiempo Parada'] > 24,
                                                           df_ultimos_2['Tiempo Parada'] / 3600,
                                                           df_ultimos_2['Tiempo Parada'])
                df_ultimos_2['Tiempo restante [h]'] = (9 - df_ultimos_2['Tiempo Neto [h]'] - df_ultimos_2['Tiempo de parada [h]'])
                df_ultimos_2['CANT. NO CONFORMES'] = (
                    df_ultimos_2['CANT. NO CONFORMES'].fillna(0)
                    if 'CANT. NO CONFORMES' in df_ultimos_2.columns
                    else 0
                )

                df_ultimos_2['CANT. REAL'] = (
                    df_ultimos_2['CANT. SUB DIARIA']
                    if 'CANT. SUB DIARIA' in df_ultimos_2.columns
                    else 0
                )


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
                df_ultimos_2['Calidad'] =np.where(
                    df_ultimos_2['CANT. REAL'] > 0,
                    (
                        df_ultimos_2['CANT. REAL'] - df_ultimos_2['CANT. NO CONFORMES']
                    )  / df_ultimos_2['CANT. REAL'],
                    0
                )
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

# ==========================================
# EXPORTACIÓN A GOOGLE SHEETS (LIBRO DE 2 DÍAS)
# ==========================================
spreadsheet_id = "1qMbb87guSbqPx3fd0tojIfTYuPnuN-r4I2mq8crqL58"
sheet_su = gc.open_by_key(spreadsheet_id)

diccionarios_hojas = {
    "INYECCION": tablas_inyectoras if 'tablas_inyectoras' in globals() else {},
    "B": {k: v for k, v in tablas_ac_b_h_ext.items() if k.startswith('B')} if 'tablas_ac_b_h_ext' in globals() else {},
    "AC": {k: v for k, v in tablas_ac_b_h_ext.items() if k.startswith('AC')} if 'tablas_ac_b_h_ext' in globals() else {},
    "H1": {k: v for k, v in tablas_ac_b_h_ext.items() if k == 'H1'} if 'tablas_ac_b_h_ext' in globals() else {},
    "Extrusora": {k: v for k, v in tablas_ac_b_h_ext.items() if k == 'Extrusora'} if 'tablas_ac_b_h_ext' in globals() else {}
}

for nombre_hoja_unica, tablas_dict in diccionarios_hojas.items():
    if not tablas_dict:
        continue

    try:
        worksheet = sheet_su.worksheet(nombre_hoja_unica)
        worksheet.clear()
        time.sleep(1.5)
    except gspread.exceptions.WorksheetNotFound:
        worksheet = sheet_su.add_worksheet(title=nombre_hoja_unica, rows=100, cols=15)
        time.sleep(1.5)

    fila_actual = 1
    for maquina, df_final in tablas_dict.items():
        try:
            titulo_maquina = [[f"--- {maquina} ---"]]
            worksheet.update(values=titulo_maquina, range_name=f"A{fila_actual}")
            fila_actual += 1

            df_para_subir = df_final.reset_index()
            df_para_subir.columns.values[0] = "Métrica"
            datos_a_escribir = [df_para_subir.columns.values.tolist()] + df_para_subir.values.tolist()

            worksheet.update(values=datos_a_escribir, range_name=f"A{fila_actual}")
            fila_actual += len(datos_a_escribir) + 2
            time.sleep(1.5)
        except Exception as e:
            print(f"❌ Error al procesar '{maquina}' en la hoja '{nombre_hoja_unica}': {e}")
            time.sleep(3)


# ==========================================
# EXPORTACIÓN A RESUMEN SEMANAL (LIBRO DE 5 DÍAS)
# ==========================================
dias_habiles_semana = pd.bdate_range(end=fecha_tope, periods=5)
idx_semana_date = [d.date() for d in dias_habiles_semana]
dias_labels = {0: 'L', 1: 'M', 2: 'X', 3: 'J', 4: 'V'}
nombres_cols_semana = [f"{d.day}/{d.month:02d} ({dias_labels.get(d.weekday(), '')})" for d in idx_semana_date]

maquinas_inyectoras = ['INY 2', 'INY 3', 'INY 4', 'INY 5']
maquinas_otras = ['B 1', 'B 2', 'B 3', 'B 4', 'B 5', 'AC 1', 'AC 2', 'AC 3', 'AC 4', 'AC 7', 'AC 10', 'AC 11', 'H1', 'Extrusora']
maquinas_orden_completo = maquinas_inyectoras + maquinas_otras

def calcular_resumen_maquina(df_maq, maquina, fechas_idx):
    if df_maq.empty:
        return pd.DataFrame()

    cols_a_convertir = ['Cantidad Real', 'CANT. SUB DIARIA', 'CANT. NO CONFORMES','Tiempo Neto', 'Tiempo de Proceso', 'Tiempo Parada', 'GPH', 'MUL']
    for col in cols_a_convertir:
        if col in df_maq.columns:
            df_maq[col] = pd.to_numeric(df_maq[col].astype(str).str.replace(',', '.'), errors='coerce').fillna(0)

    agg_dict = {'Cantidad Real': 'sum', 'CANT. SUB DIARIA': 'sum', 'CANT. NO CONFORMES': 'sum','Tiempo Neto': 'sum',
                'Tiempo de Proceso': 'sum', 'Tiempo Parada': 'sum', 'GPH': 'mean', 'MUL': 'mean'}
    agg_dict_clean = {k: v for k, v in agg_dict.items() if k in df_maq.columns}

    df_diario = df_maq.groupby(df_maq['FECHA'].dt.date).agg(agg_dict_clean)
    df_diario.index = pd.to_datetime(df_diario.index).date

    df_res = df_diario.reindex(fechas_idx).copy()

    df_res['Tiempo Neto [h]'] = df_res['Tiempo Neto'] / 3600
    df_res['Tiempo de parada [h]'] = np.where(df_res['Tiempo Parada'] > 24, df_res['Tiempo Parada'] / 3600, df_res['Tiempo Parada'])
    df_res['Tiempo restante [h]'] = (9 - df_res['Tiempo Neto [h]'] - df_res['Tiempo de parada [h]'])

    df_res['CANT. REAL'] = df_res['CANT. SUB DIARIA'] if maquina in maquinas_inyectoras else df_res['Cantidad Real']

    if maquina in maquinas_inyectoras:
        df_res['CANT. NO CONFORMES'] = (
            df_res['CANT. NO CONFORMES'].fillna(0)
            if 'CANT. NO CONFORMES' in df_res.columns
            else 0
        )
    else:
         df_res['CANT. NO CONFORMES'] = 0


    
    df_res['Disponibilidad'] = np.where(df_res['Tiempo de Proceso'] > 0, df_res['Tiempo Neto'] / df_res['Tiempo de Proceso'], 0) * 100
    df_res['Rendimiento'] = np.where((df_res['GPH'] > 0) & (df_res['MUL'] > 0) & (df_res['Tiempo Neto'] > 0),
                                     ((df_res['CANT. REAL'] / df_res['MUL']) / (df_res['Tiempo Neto'] / 3600)) / df_res['GPH'] * 100, 0)
    df_res['Calidad'] = (
         df_res['CANT. REAL'] > 0,
        (
            (df_res['CANT. REAL'] - df_res['CANT. NO CONFORMES'])
            / df_res['CANT. REAL']
        ) * 100,
        0
    )

    df_res['OEE'] = (
        (df_res['Disponibilidad'] / 100)
        * (df_res['Rendimiento'] / 100) 
        * (df_res['Calidad'] / 100)
    )* 100


    return df_res

todas_las_tablas = {}

def normalizar_df(df):
    if df.empty: return df
    df = df.copy()
    if 'Fecha' in df.columns:
        df = df.rename(columns={'Fecha': 'FECHA'})
    elif 'Fecha Inicio' in df.columns:
        df = df.rename(columns={'Fecha Inicio': 'FECHA'})
    return df

df_iny_clean = normalizar_df(df_app_iny_unificado.loc[:, ~df_app_iny_unificado.columns.duplicated()]) if 'df_app_iny_unificado' in globals() else pd.DataFrame()
df_lot_clean = normalizar_df(df_lot_fal_agrupado.loc[:, ~df_lot_fal_agrupado.columns.duplicated()]) if 'df_lot_fal_agrupado' in globals() else pd.DataFrame()

if not df_iny_clean.empty: df_iny_clean = pd.DataFrame(df_iny_clean.to_dict('list'))
if not df_lot_clean.empty: df_lot_clean = pd.DataFrame(df_lot_clean.to_dict('list'))

for maquina in maquinas_orden_completo:
    if maquina in maquinas_inyectoras:
        df_maq = df_iny_clean[df_iny_clean['Maquina'].astype(str).str.strip() == maquina].copy()
    else:
        df_maq = df_lot_clean[df_lot_clean['Maquina'].astype(str).str.strip() == maquina].copy()

    if not df_maq.empty and 'FECHA' in df_maq.columns:
        df_maq['FECHA'] = pd.to_datetime(df_maq['FECHA'], errors='coerce', format='mixed')

    df_calc = calcular_resumen_maquina(df_maq, maquina, idx_semana_date)

    df_final = pd.DataFrame('-', index=['OEE', 'Disponibilidad', 'Calidad', 'Rendimiento', 'CANT. NO CONFORMES',
                                       'CANT. REAL', 'Tiempo Neto [h]', 'Tiempo de parada [h]', 'Tiempo restante [h]'],
                            columns=nombres_cols_semana)

    if not df_calc.empty:
        for col_date in nombres_cols_semana:
            fecha_col = pd.to_datetime(col_date.split(' (')[0] + f"/{datetime.now().year}", dayfirst=True).date()
            if fecha_col in df_calc.index:
                row = df_calc.loc[fecha_col]
                df_final.loc['OEE', col_date] = f"{row['OEE']:.2f}%" if pd.notna(row['OEE']) else "-"
                df_final.loc['Disponibilidad', col_date] = f"{row['Disponibilidad']:.2f}%" if pd.notna(row['Disponibilidad']) else "-"
                df_final.loc['Calidad', col_date] = f"{row['Calidad']:.2f}%"
                df_final.loc['Rendimiento', col_date] = f"{row['Rendimiento']:.2f}%" if pd.notna(row['Rendimiento']) else "-"
                df_final.loc['CANT. NO CONFORMES', col_date] = (
                    f"{int(round(row['CANT. NO CONFORMES']))}"
                    if pd.notna(row['CANT. NO CONFORMES'])
                    else "-"
                )
                

                df_final.loc['CANT. REAL', col_date] = f"{int(round(row['CANT. REAL']))}" if pd.notna(row['CANT. REAL']) else "-"
                df_final.loc['Tiempo Neto [h]', col_date] = f"{row['Tiempo Neto [h]']:.2f}" if pd.notna(row['Tiempo Neto [h]']) else "-"
                df_final.loc['Tiempo de parada [h]', col_date] = f"{row['Tiempo de parada [h]']:.2f}" if pd.notna(row['Tiempo de parada [h]']) else "-"
                df_final.loc['Tiempo restante [h]', col_date] = f"{row['Tiempo restante [h]']:.2f}" if pd.notna(row['Tiempo restante [h]']) else "-"

    todas_las_tablas[maquina] = df_final

spreadsheet_id_5d = "1s2Kd8Sr0p_D76LPh1uJgIa9UAGro-UFZijK-_AOABTc"
sheet_5d = gc.open_by_key(spreadsheet_id_5d)

try:
    worksheet_resumen = sheet_5d.worksheet('Resumen Semanal')
except gspread.exceptions.WorksheetNotFound:
    worksheet_resumen = sheet_5d.add_worksheet(title='Resumen Semanal', rows=300, cols=10)

worksheet_resumen.clear()
time.sleep(1)

fila_actual = 1
for maquina, df_final in todas_las_tablas.items():
    try:
        worksheet_resumen.update(range_name=f"A{fila_actual}", values=[[f"--- {maquina} ---"]])
        fila_actual += 1
        
        df_para_subir = df_final.reset_index()
        df_para_subir.columns.values[0] = "Métrica"
        
        datos_a_escribir = [df_para_subir.columns.values.tolist()] + df_para_subir.values.tolist()
        worksheet_resumen.update(range_name=f"A{fila_actual}", values=datos_a_escribir)
        
        fila_actual += len(datos_a_escribir) + 2
        time.sleep(1.5)
    except Exception as e:
        print(f"❌ Error al exportar la máquina {maquina} al Resumen Semanal (5 días): {e}")
        time.sleep(3)

print("✅ ¡Proceso completo! Ambos libros de Google Sheets se han actualizado correctamente.")