"""
================================================================================
 DASHBOARD ESTRATÉGICO FIXY LOGÍSTICA — Backend
 -------------------------------------------------------------------------------
 Procesa automáticamente:
   - Tablero_Fixy_AAAA.xlsx       (consolidado histórico por año)
   - FACTURACION_AAAA_DD_MM_AA.xlsx (envíos + facturación detalle)
   - Fulfillment_AAAA.xlsx        (operación FixyFull)

 Modo de uso:
   1) Coloca los archivos en la carpeta /data (o subílos por la UI)
   2) Ejecutá:  python app.py
   3) Abrí:     http://localhost:5000

 Diseñado para C-Level: KPIs ejecutivos, proyecciones, alertas, drill-down.
================================================================================
"""
from flask import Flask, render_template, jsonify, request, send_file
import pandas as pd
import numpy as np
import openpyxl
import os
import json
from datetime import datetime
from pathlib import Path
import glob
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'data'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB

DATA_DIR = Path(__file__).parent / 'data'
DATA_DIR.mkdir(exist_ok=True)

# ============================================================
# CONFIGURACIÓN DE NEGOCIO
# ============================================================
OBJETIVO_GROWTH = 0.70          # +70% YoY
IPC_2025_ACUM = 0.279           # 27.9% acumulado 2025 (INDEC/Infobae)
IPC_2026_ESTIMADO_ANUAL = 0.25  # Estimación, ajustable desde UI

# Mapeo servicios → categorías estratégicas
def categorize_service(s):
    if not s:
        return 'Sin categoría'
    s_up = str(s).upper()
    if 'FIXY INTERIOR' in s_up or 'INTERIOR' in s_up:
        return 'Interior Andreani'
    if 'FLEX' in s_up:
        return 'Flex'
    if 'ENTREGA 24' in s_up or 'NEXT' in s_up:
        return 'Next Day'
    if 'STANDARD' in s_up:
        return 'Standard'
    if 'SAME DAY' in s_up:
        return 'Same Day'
    if 'PICK UP' in s_up:
        return 'Pick Up'
    if 'CAMBIO' in s_up:
        return 'Cambios'
    if 'DEVOLUCION' in s_up or 'DEVOL' in s_up:
        return 'Devoluciones'
    if 'RETIRO' in s_up or 'RETIRA' in s_up:
        return 'Retiros'
    if 'IMPRESION' in s_up:
        return 'Con Impresión'
    if 'PUERTA' in s_up:
        return 'Puerta a Puerta'
    return 'Otros'

MES_ORDER = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
             'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
MES_CORTO = {m: m[:3] for m in MES_ORDER}


# ============================================================
# CACHE EN MEMORIA — se invalida al subir archivos
# ============================================================
CACHE = {
    'envios': None,
    'facturacion': None,
    'fulfillment': None,
    'tablero': None,
    'last_update': None,
    'files_loaded': []
}


# ============================================================
# LOADERS — Auto-detectan los archivos por nombre/sheet
# ============================================================

def find_file(pattern):
    """Encuentra el archivo más reciente que coincida con un patrón."""
    files = list(DATA_DIR.glob(pattern))
    if not files:
        return None
    return max(files, key=lambda f: f.stat().st_mtime)


def load_tablero():
    """Carga consolidado mensual histórico (2024/2025/2026)."""
    f = find_file('Tablero*.xlsx') or find_file('tablero*.xlsx')
    if not f:
        return None

    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    result = {}

    for año_sheet in ['2024', '2025', '2026']:
        if año_sheet not in wb.sheetnames:
            continue
        ws = wb[año_sheet]
        envios_mes = {}
        fact_mes = {}

        # El primer renglón con fechas tiene los meses
        # Recorremos buscando filas que comienzan con "Q de transacciones" o "Facturación"
        for row in ws.iter_rows(values_only=True):
            if row[0] is None:
                continue
            label = str(row[0]).strip().lower()
            if label in ('q de transacciones', 'q de envios', 'q de envíos'):
                for i, val in enumerate(row[1:13], start=1):
                    if val is not None and isinstance(val, (int, float)):
                        envios_mes[i] = int(val)
            elif label in ('facturación', 'facturacion'):
                for i, val in enumerate(row[1:13], start=1):
                    if val is not None and isinstance(val, (int, float)):
                        fact_mes[i] = float(val)

        result[año_sheet] = {'envios': envios_mes, 'facturacion': fact_mes}

    wb.close()
    return result


def load_envios():
    """Carga el detalle de envíos (sheet ENVIOS del archivo FACTURACION)."""
    f = find_file('FACTURACION*.xlsx') or find_file('facturacion*.xlsx')
    if not f:
        return None

    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    if 'ENVIOS' not in wb.sheetnames:
        wb.close()
        return None

    ws = wb['ENVIOS']
    data = []
    headers = None
    idx = {}

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0:
            headers = row
            # Mapeo robusto de columnas
            for col_name in ['MES', 'Cliente', 'Servicio', 'Importe',
                             'Contrareembolso', 'Fecha', 'Estado', 'Provincia',
                             'Localidad', 'Peso']:
                try:
                    idx[col_name] = headers.index(col_name)
                except ValueError:
                    idx[col_name] = None
            continue

        if row[idx['MES']] is None:
            continue

        data.append({
            'mes': row[idx['MES']],
            'cliente': row[idx['Cliente']],
            'servicio': row[idx['Servicio']],
            'importe': float(row[idx['Importe']] or 0),
            'contrareembolso': float(row[idx['Contrareembolso']] or 0),
            'fecha': row[idx['Fecha']],
            'estado': row[idx['Estado']],
            'provincia': row[idx['Provincia']] if idx['Provincia'] is not None else None,
            'peso': float(row[idx['Peso']] or 0) if idx['Peso'] is not None else 0,
        })

    wb.close()

    df = pd.DataFrame(data)
    df['cat_servicio'] = df['servicio'].apply(categorize_service)
    # Ingreso real Fixy = Importe (flete) + 1% del contrareembolso
    df['ingreso_fixy'] = df['importe'] + 0.01 * df['contrareembolso']
    df['tipo_pago'] = df['contrareembolso'].apply(
        lambda x: 'Contraentrega' if x > 0 else 'Pago Anticipado'
    )
    return df


def load_fulfillment():
    """Carga la matriz de Fulfillment + facturación FixyFull."""
    f = find_file('Fulfillment*.xlsx') or find_file('fulfillment*.xlsx')
    if not f:
        return None

    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)

    # MATRIZ — operación
    matriz = []
    if 'MATRIZ' in wb.sheetnames:
        ws = wb['MATRIZ']
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0 or row[0] is None:
                continue
            matriz.append({
                'año': row[0],
                'periodo': row[1],
                'semana': row[2],
                'cliente': row[4],
                'estado': row[10] if len(row) > 10 else None,
                'cantidad': row[16] if len(row) > 16 else 0,
                'nota_pedido': row[9] if len(row) > 9 else None,
            })

    # FACTURACION fulfillment
    fact_fulf = {}
    if 'FACTURACION' in wb.sheetnames:
        ws = wb['FACTURACION']
        rows = list(ws.iter_rows(values_only=True))
        mes_actual = None
        for row in rows[3:30]:
            if row[0] is not None and isinstance(row[0], str) and row[0] in MES_ORDER:
                mes_actual = row[0]
                fact_fulf.setdefault(mes_actual, 0)
            if mes_actual and len(row) > 19 and row[19] is not None:
                try:
                    fact_fulf[mes_actual] = fact_fulf.get(mes_actual, 0) + float(row[19])
                except (TypeError, ValueError):
                    pass

    wb.close()

    return {
        'matriz': pd.DataFrame(matriz),
        'facturacion': fact_fulf,
    }


def load_facturacion_resumen():
    """Carga la pestaña FACTURACION (liquidaciones)."""
    f = find_file('FACTURACION*.xlsx') or find_file('facturacion*.xlsx')
    if not f:
        return None
    wb = openpyxl.load_workbook(f, read_only=True, data_only=True)
    if 'FACTURACION' not in wb.sheetnames:
        wb.close()
        return None
    ws = wb['FACTURACION']
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 or row[0] is None:
            continue
        rows.append({
            'mes': row[0],
            'cliente': row[3],
            'fecha': row[4],
            'parcial': float(row[5] or 0),
            'seguro': float(row[6] or 0),
            'adicionales': float(row[7] or 0),
            'costo_cr': float(row[8] or 0),
            'importe': float(row[9] or 0),
        })
    wb.close()
    return pd.DataFrame(rows)


def reconciliar_tablero_con_envios(tablero, df_envios):
    """
    El Tablero del cliente puede estar desactualizado.
    Completamos los meses faltantes de 2026 con el detalle real de envíos.
    """
    if df_envios is None or not tablero:
        return tablero

    if '2026' not in tablero:
        tablero['2026'] = {'envios': {}, 'facturacion': {}}

    # Mapear mes string → número
    mes_a_num = {m: i + 1 for i, m in enumerate(MES_ORDER)}

    # Q de envíos y facturación por mes desde el detalle
    grouped = df_envios.groupby('mes').agg(
        q=('cliente', 'count'),
        importe=('importe', 'sum')
    )

    for mes_str, row in grouped.iterrows():
        m_num = mes_a_num.get(mes_str)
        if m_num is None:
            continue
        # Si el tablero NO tiene ese mes, lo agregamos
        if m_num not in tablero['2026']['envios']:
            tablero['2026']['envios'][m_num] = int(row['q'])
        if m_num not in tablero['2026']['facturacion']:
            tablero['2026']['facturacion'][m_num] = float(row['importe'])

    return tablero


def refresh_cache():
    """Recarga todo desde los archivos en /data."""
    CACHE['envios'] = load_envios()
    CACHE['facturacion'] = load_facturacion_resumen()
    CACHE['fulfillment'] = load_fulfillment()
    CACHE['tablero'] = load_tablero()
    # Reconciliar tablero con envíos para que los meses parciales aparezcan
    CACHE['tablero'] = reconciliar_tablero_con_envios(CACHE['tablero'], CACHE['envios'])
    CACHE['last_update'] = datetime.now().isoformat()
    CACHE['files_loaded'] = [f.name for f in DATA_DIR.glob('*.xlsx')]


# ============================================================
# ENDPOINTS — JSON para el frontend
# ============================================================

@app.route('/')
def index():
    return render_template('dashboard.html')


@app.route('/api/status')
def status():
    return jsonify({
        'last_update': CACHE['last_update'],
        'files_loaded': CACHE['files_loaded'],
        'has_envios': CACHE['envios'] is not None,
        'has_fulfillment': CACHE['fulfillment'] is not None,
        'has_tablero': CACHE['tablero'] is not None,
    })


@app.route('/api/upload', methods=['POST'])
def upload():
    """Recibe múltiples archivos Excel y los guarda en /data."""
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No se enviaron archivos'}), 400

    saved = []
    for f in files:
        if not f.filename.endswith(('.xlsx', '.xlsm')):
            continue
        # Renombrar conservando el "tipo" del archivo para que find_file funcione
        path = DATA_DIR / f.filename
        f.save(path)
        saved.append(f.filename)

    refresh_cache()
    return jsonify({'ok': True, 'saved': saved, 'cache': {
        'files_loaded': CACHE['files_loaded'],
        'last_update': CACHE['last_update'],
    }})


@app.route('/api/refresh', methods=['POST'])
def refresh():
    refresh_cache()
    return jsonify({'ok': True, 'last_update': CACHE['last_update']})


# ---------- KPIs PRINCIPALES ----------

@app.route('/api/kpis')
def kpis():
    """KPIs ejecutivos — siempre comparados YoY y vs objetivo."""
    tab = CACHE['tablero']
    df = CACHE['envios']
    if not tab or df is None:
        return jsonify({'error': 'Sin datos cargados'}), 400

    año_actual = '2026'
    año_anterior = '2025'

    env_actual = tab.get(año_actual, {}).get('envios', {})
    env_ant = tab.get(año_anterior, {}).get('envios', {})
    fact_actual = tab.get(año_actual, {}).get('facturacion', {})
    fact_ant = tab.get(año_anterior, {}).get('facturacion', {})

    meses_cerrados = sorted(env_actual.keys())
    if not meses_cerrados:
        return jsonify({'error': 'Sin meses cargados en tablero'}), 400

    # YTD (Year to Date)
    ytd_env_actual = sum(env_actual.values())
    ytd_env_ant = sum(env_ant.get(m, 0) for m in meses_cerrados)
    ytd_fact_actual = sum(fact_actual.values())
    ytd_fact_ant = sum(fact_ant.get(m, 0) for m in meses_cerrados)

    # Objetivo YTD
    ytd_obj = sum(env_ant.get(m, 0) for m in meses_cerrados) * (1 + OBJETIVO_GROWTH)

    # YoY
    yoy_env = (ytd_env_actual / ytd_env_ant - 1) if ytd_env_ant else 0
    yoy_fact_nom = (ytd_fact_actual / ytd_fact_ant - 1) if ytd_fact_ant else 0
    # IPC proporcional (estimación lineal hasta el mes corriente)
    ipc_proporcional = IPC_2026_ESTIMADO_ANUAL * (len(meses_cerrados) / 12)
    yoy_fact_real = ((ytd_fact_actual / (1 + ipc_proporcional)) / ytd_fact_ant - 1) if ytd_fact_ant else 0

    # Cumplimiento objetivo
    cumplimiento = (ytd_env_actual / ytd_obj) if ytd_obj else 0

    # Precio promedio
    pp_actual = ytd_fact_actual / ytd_env_actual if ytd_env_actual else 0
    pp_ant = ytd_fact_ant / ytd_env_ant if ytd_env_ant else 0

    # Clientes activos (último mes con data)
    if df is not None:
        ult_mes_nombre = MES_ORDER[max(meses_cerrados) - 1]
        clientes_activos_ult = df[df['mes'] == ult_mes_nombre]['cliente'].nunique()
        clientes_totales_ytd = df['cliente'].nunique()
    else:
        clientes_activos_ult = 0
        clientes_totales_ytd = 0

    return jsonify({
        'ytd_envios': ytd_env_actual,
        'ytd_envios_ant': ytd_env_ant,
        'ytd_facturacion': ytd_fact_actual,
        'ytd_facturacion_ant': ytd_fact_ant,
        'ytd_objetivo_envios': int(ytd_obj),
        'yoy_envios': yoy_env,
        'yoy_facturacion_nominal': yoy_fact_nom,
        'yoy_facturacion_real': yoy_fact_real,
        'cumplimiento_objetivo': cumplimiento,
        'precio_promedio_actual': pp_actual,
        'precio_promedio_ant': pp_ant,
        'clientes_activos_ultimo_mes': clientes_activos_ult,
        'clientes_totales_ytd': clientes_totales_ytd,
        'meses_cerrados': meses_cerrados,
        'ipc_aplicado': ipc_proporcional,
    })


# ---------- EVOLUTIVO 2024/2025/2026 + OBJETIVO + PROYECCIÓN ----------

@app.route('/api/evolutivo')
def evolutivo():
    tab = CACHE['tablero']
    if not tab:
        return jsonify({'error': 'Sin tablero'}), 400

    env24 = tab.get('2024', {}).get('envios', {})
    env25 = tab.get('2025', {}).get('envios', {})
    env26 = tab.get('2026', {}).get('envios', {})
    fact25 = tab.get('2025', {}).get('facturacion', {})
    fact26 = tab.get('2026', {}).get('facturacion', {})

    # Objetivo 2026 = 2025 * 1.70
    obj26_env = {m: int(env25.get(m, 0) * (1 + OBJETIVO_GROWTH)) for m in range(1, 13)}

    # Proyección basada en la tasa de crecimiento del Q1
    meses_cerrados = sorted(env26.keys())
    if len(meses_cerrados) >= 3:
        # Tomamos los primeros 3 meses completos para la tasa
        q_real = sum(env26[m] for m in meses_cerrados[:3])
        q_ant = sum(env25.get(m, 0) for m in meses_cerrados[:3])
        tasa = (q_real / q_ant) if q_ant else 1
    else:
        tasa = 1 + OBJETIVO_GROWTH

    proy26 = dict(env26)
    for m in range(1, 13):
        if m not in proy26:
            proy26[m] = int(env25.get(m, 0) * tasa)

    return jsonify({
        '2024': [env24.get(m, None) for m in range(1, 13)],
        '2025': [env25.get(m, None) for m in range(1, 13)],
        '2026_real': [env26.get(m, None) for m in range(1, 13)],
        '2026_objetivo': [obj26_env[m] for m in range(1, 13)],
        '2026_proyeccion': [proy26[m] for m in range(1, 13)],
        'facturacion_2025': [fact25.get(m, None) for m in range(1, 13)],
        'facturacion_2026': [fact26.get(m, None) for m in range(1, 13)],
        'tasa_crecimiento_aplicada': tasa - 1,
        'total_obj_anual': sum(obj26_env.values()),
        'total_proy_anual': sum(proy26.values()),
        'total_2025': sum(env25.values()),
    })


# ---------- SERVICIOS ----------

@app.route('/api/servicios')
def servicios():
    df = CACHE['envios']
    if df is None:
        return jsonify({'error': 'Sin envíos'}), 400

    mes_filter = request.args.get('mes')
    if mes_filter and mes_filter != 'ALL':
        df = df[df['mes'] == mes_filter]

    # Por categoría
    cat = df.groupby('cat_servicio').agg(
        q_envios=('cliente', 'count'),
        importe=('importe', 'sum'),
        cr=('contrareembolso', 'sum'),
        ingreso_fixy=('ingreso_fixy', 'sum')
    ).round(2).sort_values('ingreso_fixy', ascending=False)

    cat_dict = cat.reset_index().to_dict('records')

    # Evolución por mes y categoría
    meses_data = sorted(df['mes'].unique().tolist(), key=lambda x: MES_ORDER.index(x))
    pivot = df.groupby(['mes', 'cat_servicio'])['ingreso_fixy'].sum().unstack(fill_value=0)
    pivot = pivot.reindex(meses_data)

    evolutivo = {
        'meses': [MES_CORTO[m] for m in meses_data],
        'series': {col: pivot[col].round(0).tolist() for col in pivot.columns}
    }

    return jsonify({
        'categorias': cat_dict,
        'evolutivo': evolutivo,
        'total_envios': len(df),
        'total_ingreso_fixy': float(df['ingreso_fixy'].sum()),
    })


# ---------- TOP 10 CLIENTES + COMPORTAMIENTO ----------

@app.route('/api/top-clientes')
def top_clientes():
    df = CACHE['envios']
    if df is None:
        return jsonify({'error': 'Sin envíos'}), 400

    n = int(request.args.get('n', 10))
    metric = request.args.get('metric', 'ingreso_fixy')  # ingreso_fixy | q_envios | importe

    agg = df.groupby('cliente').agg(
        q_envios=('cliente', 'count'),
        importe=('importe', 'sum'),
        cr=('contrareembolso', 'sum'),
        ingreso_fixy=('ingreso_fixy', 'sum')
    ).round(2)

    top = agg.nlargest(n, metric).reset_index()

    # Para cada cliente, evolución por mes
    top_list = top['cliente'].tolist()
    df_top = df[df['cliente'].isin(top_list)]
    evol = df_top.groupby(['cliente', 'mes']).size().unstack(fill_value=0)
    meses_disponibles = [m for m in MES_ORDER if m in evol.columns]
    evol = evol.reindex(columns=meses_disponibles, fill_value=0)
    evol = evol.reindex(top_list)

    # Detección de churn / crecimiento
    analisis = []
    for cli in top_list:
        serie = evol.loc[cli].tolist()
        primer = serie[0] if serie else 0
        ultimo = serie[-1] if serie else 0
        ene = serie[0] if len(serie) > 0 else 0
        # Comparar último vs penúltimo mes cerrado (no el parcial)
        if len(serie) >= 2:
            antepenultimo = serie[-2]
            tendencia_pct = ((serie[-1] - serie[-2]) / serie[-2] * 100) if serie[-2] else 0
        else:
            tendencia_pct = 0

        # Comparar primer mes vs último mes cerrado completo (excluyendo el actual si está parcial)
        if len(serie) >= 2:
            ref_final = serie[-2]
        else:
            ref_final = serie[-1]
        crecimiento_pct = ((ref_final - primer) / primer * 100) if primer else 0

        # Estado
        if ultimo == 0 and primer > 0:
            estado = '🔴 Inactivo / Churn'
        elif crecimiento_pct > 50:
            estado = '🟢 Crecimiento fuerte'
        elif crecimiento_pct > 0:
            estado = '🟢 Crece'
        elif crecimiento_pct > -25:
            estado = '🟡 Estable / Leve baja'
        else:
            estado = '🔴 En descenso'

        analisis.append({
            'cliente': cli,
            'evolucion': serie,
            'crecimiento_pct': round(crecimiento_pct, 1),
            'tendencia_ultimo_mes_pct': round(tendencia_pct, 1),
            'estado': estado,
        })

    return jsonify({
        'meses': [MES_CORTO[m] for m in meses_disponibles],
        'top': top.to_dict('records'),
        'analisis': analisis,
    })


# ---------- PAGO ANTICIPADO vs CONTRAENTREGA ----------

@app.route('/api/tipo-pago')
def tipo_pago():
    df = CACHE['envios']
    if df is None:
        return jsonify({'error': 'Sin envíos'}), 400

    # Totales
    total = df.groupby('tipo_pago').agg(
        q=('cliente', 'count'),
        importe=('importe', 'sum'),
        cr=('contrareembolso', 'sum'),
        ingreso=('ingreso_fixy', 'sum')
    ).round(2)

    # Por mes
    meses = sorted(df['mes'].unique().tolist(), key=lambda x: MES_ORDER.index(x))
    pivot = df.groupby(['mes', 'tipo_pago']).size().unstack(fill_value=0)
    pivot = pivot.reindex(meses)
    pivot_pct = pivot.div(pivot.sum(axis=1), axis=0) * 100

    return jsonify({
        'totales': total.reset_index().to_dict('records'),
        'meses': [MES_CORTO[m] for m in meses],
        'q_contraentrega': pivot.get('Contraentrega', pd.Series([0]*len(meses))).tolist(),
        'q_anticipado': pivot.get('Pago Anticipado', pd.Series([0]*len(meses))).tolist(),
        'pct_contraentrega': pivot_pct.get('Contraentrega', pd.Series([0]*len(meses))).round(1).tolist(),
        'pct_anticipado': pivot_pct.get('Pago Anticipado', pd.Series([0]*len(meses))).round(1).tolist(),
    })


# ---------- FULFILLMENT (FixyFull) ----------

@app.route('/api/fulfillment')
def fulfillment():
    f = CACHE['fulfillment']
    if not f or f['matriz'].empty:
        return jsonify({'error': 'Sin fulfillment'}), 400

    df = f['matriz']
    fact = f['facturacion']

    meses_disp = [m for m in MES_ORDER if m in df['periodo'].unique()]

    pedidos = df.groupby('periodo')['nota_pedido'].nunique().reindex(meses_disp).fillna(0).astype(int)
    unidades = df.groupby('periodo')['cantidad'].sum().reindex(meses_disp).fillna(0).astype(int)
    clientes = df.groupby('periodo')['cliente'].nunique().reindex(meses_disp).fillna(0).astype(int)

    # Top clientes
    top_cli = df.groupby('cliente')['nota_pedido'].nunique().nlargest(10).to_dict()

    fact_lista = [round(fact.get(m, 0), 2) for m in meses_disp]

    return jsonify({
        'meses': [MES_CORTO[m] for m in meses_disp],
        'pedidos': pedidos.tolist(),
        'unidades': unidades.tolist(),
        'clientes': clientes.tolist(),
        'facturacion': fact_lista,
        'top_clientes': top_cli,
        'totales': {
            'pedidos_ytd': int(pedidos.sum()),
            'unidades_ytd': int(unidades.sum()),
            'facturacion_ytd': sum(fact_lista),
            'clientes_unicos': int(df['cliente'].nunique()),
        }
    })


# ---------- SIMULADOR DE ESCENARIOS ----------

@app.route('/api/simulador', methods=['POST'])
def simulador():
    """Permite simular escenarios variando IPC y growth target."""
    data = request.get_json() or {}
    growth = float(data.get('growth_target', OBJETIVO_GROWTH))
    ipc = float(data.get('ipc_anual', IPC_2026_ESTIMADO_ANUAL))
    tasa_aplicada = float(data.get('tasa_real', 0.55))  # tasa de crecimiento esperada

    tab = CACHE['tablero']
    if not tab:
        return jsonify({'error': 'Sin tablero'}), 400

    env25 = tab.get('2025', {}).get('envios', {})
    env26 = tab.get('2026', {}).get('envios', {})
    fact25 = tab.get('2025', {}).get('facturacion', {})

    meses_cerrados = sorted(env26.keys())

    # Objetivo
    obj_anual = sum(env25.values()) * (1 + growth)
    # Proyección
    proy_anual_envios = sum(env26.values()) + sum(env25.get(m, 0) * (1 + tasa_aplicada) for m in range(1, 13) if m not in meses_cerrados)
    gap = obj_anual - proy_anual_envios

    # Facturación proyectada nominal
    fact_proy_anual = sum([fact25.get(m, 0) for m in range(1, 13)]) * (1 + tasa_aplicada) * (1 + ipc)

    return jsonify({
        'growth_target_pct': growth * 100,
        'ipc_anual_pct': ipc * 100,
        'tasa_aplicada_pct': tasa_aplicada * 100,
        'objetivo_envios_anual': int(obj_anual),
        'proyeccion_envios_anual': int(proy_anual_envios),
        'gap_envios': int(gap),
        'cumplimiento_pct': proy_anual_envios / obj_anual * 100 if obj_anual else 0,
        'facturacion_proyectada': fact_proy_anual,
    })


# ---------- ALERTAS ESTRATÉGICAS ----------

@app.route('/api/alertas')
def alertas():
    df = CACHE['envios']
    tab = CACHE['tablero']
    if df is None or not tab:
        return jsonify({'alertas': []})

    alertas_list = []

    # 1) Cumplimiento de objetivo bajo
    env25 = tab.get('2025', {}).get('envios', {})
    env26 = tab.get('2026', {}).get('envios', {})
    meses_cerrados = sorted(env26.keys())
    if meses_cerrados:
        ytd_real = sum(env26.values())
        ytd_obj = sum(env25.get(m, 0) for m in meses_cerrados) * (1 + OBJETIVO_GROWTH)
        cumpl = ytd_real / ytd_obj if ytd_obj else 0
        if cumpl < 0.85:
            alertas_list.append({
                'severidad': 'alta',
                'titulo': 'Cumplimiento de objetivo bajo',
                'mensaje': f'YTD al {cumpl*100:.1f}% del objetivo. Necesario recuperar {int(ytd_obj - ytd_real):,} envíos.',
            })

    # Detectar si el último mes es parcial: comparamos vs el mes anterior promedio
    # Si el último mes tiene <50% de la media de los otros, lo marcamos como parcial
    meses_validos = [m for m in MES_ORDER if m in df['mes'].unique()]
    if len(meses_validos) >= 3:
        envios_por_mes = df.groupby('mes').size()
        ultimo = meses_validos[-1]
        otros_promedio = envios_por_mes.drop(ultimo).mean()
        ultimo_q = envios_por_mes.get(ultimo, 0)
        mes_parcial = bool(ultimo_q < otros_promedio * 0.65)
        mes_referencia_idx = -2 if mes_parcial else -1
    else:
        mes_parcial = False
        mes_referencia_idx = -1

    # 2) Análisis de top clientes — usando mes de referencia (no parcial)
    top10 = df.groupby('cliente')['ingreso_fixy'].sum().nlargest(10).index.tolist()
    for cli in top10:
        serie = df[df['cliente'] == cli].groupby('mes').size().reindex(MES_ORDER, fill_value=0)
        if len(meses_validos) >= 3:
            # Mes referencia (último mes COMPLETO) y mes previo
            mes_ref = meses_validos[mes_referencia_idx]
            mes_prev_idx = mes_referencia_idx - 1
            if abs(mes_prev_idx) <= len(meses_validos):
                mes_prev = meses_validos[mes_prev_idx]
                if serie[mes_ref] == 0 and serie[mes_prev] > 50:
                    nota_extra = ' (excluyendo el mes en curso parcial)' if mes_parcial else ''
                    alertas_list.append({
                        'severidad': 'alta',
                        'titulo': f'Cliente top inactivo: {cli}',
                        'mensaje': f'{cli} dejó de operar en {mes_ref.lower()}{nota_extra}. Riesgo de churn.',
                    })
                elif serie[mes_prev] > 0 and serie[mes_ref] > 0:
                    var = (serie[mes_ref] - serie[mes_prev]) / serie[mes_prev]
                    if var < -0.4:
                        alertas_list.append({
                            'severidad': 'media',
                            'titulo': f'Caída fuerte: {cli}',
                            'mensaje': f'{cli} cayó {abs(var)*100:.0f}% en {mes_ref.lower()} vs {mes_prev.lower()}.',
                        })

    # 3) Concentración de ingresos en pocos clientes
    total = df['ingreso_fixy'].sum()
    top3 = df.groupby('cliente')['ingreso_fixy'].sum().nlargest(3).sum()
    concentracion = top3 / total if total else 0
    if concentracion > 0.65:
        alertas_list.append({
            'severidad': 'media',
            'titulo': 'Alta concentración de ingresos',
            'mensaje': f'Top 3 clientes concentran {concentracion*100:.1f}% del ingreso Fixy. Diversificar cartera es prioridad estratégica.',
        })

    # 4) Aviso de mes parcial
    if mes_parcial:
        alertas_list.append({
            'severidad': 'baja',
            'titulo': f'Mes en curso parcial: {meses_validos[-1]}',
            'mensaje': f'Los datos de {meses_validos[-1].lower()} están parciales. Las comparaciones de tendencia usan el último mes completo como referencia.',
        })

    return jsonify({'alertas': alertas_list, 'mes_parcial': mes_parcial, 'meses_validos': meses_validos})


# ============================================================
# INICIALIZACIÓN
# ============================================================
if __name__ == '__main__':
    # Carga inicial si hay archivos en /data
    refresh_cache()
    print('=' * 60)
    print(' DASHBOARD FIXY LOGÍSTICA')
    print(' http://localhost:5000')
    print(f' Archivos cargados: {CACHE["files_loaded"]}')
    print('=' * 60)
    app.run(host='0.0.0.0', port=5000, debug=False)
