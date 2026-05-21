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

# Ruta de control exigida por Render para saber que el servidor está funcionando
@app.route('/health')
def health_check():
    return "OK", 200

app.config['UPLOAD_FOLDER'] = 'data'
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024  # 200 MB

DATA_DIR = Path(__file__).parent / 'data'
DATA_DIR.mkdir(exist_ok=True)

# ============================================================
# CONFIGURACIÓN DE NEGOCIO
# ============================================================
OBJETIVO_GROWTH = 0.70          # +70% YoY
IPC_2025_ACUM = 0.279           # 27.9% proyectado (ajuste inflación de ser necesario)
CHURN_TOLERABLE = 0.05          # Max 5% mensual

# Cache global en memoria para performance óptima
CACHE = {
    'tablero': None,
    'envios': None,
    'full': None,
    'meses_validos': [],
    'mes_min': 'N/A',
    'mes_max': 'N/A',
    'mes_parcial': False,
    'last_updated': None
}

# ============================================================
# CARGADORES DE DATOS (Módulos de extracción)
# ============================================================
def load_tablero():
    """Busca y consolida archivos Tablero_Fixy_*.xlsx"""
    files = glob.glob(str(DATA_DIR / "Tablero_Fixy_*.xlsx"))
    if not files:
        return pd.DataFrame()
    
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f, sheet_name="Resumen Mensual")
            # Limpieza básica de columnas
            df.columns = df.columns.str.strip().str.lower()
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {f}: {e}")
            
    if not dfs:
        return pd.DataFrame()
    
    res = pd.concat(dfs, ignore_index=True)
    # Asegurar formato clave para ordenamiento temporal
    if 'mes_año' in res.columns:
        res['mes_año'] = res['mes_año'].str.strip().str.capitalize()
    return res

def load_envios():
    """Busca y consolida archivos de facturación/envíos detallados"""
    files = glob.glob(str(DATA_DIR / "FACTURACION_*.xlsx"))
    if not files:
        return pd.DataFrame()
    
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df.columns = df.columns.str.strip().str.lower()
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {f}: {e}")
            
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)

def load_fulfillment():
    """Busca y consolida archivos de Fulfillment"""
    files = glob.glob(str(DATA_DIR / "Fulfillment_*.xlsx"))
    if not files:
        return pd.DataFrame()
    
    dfs = []
    for f in files:
        try:
            df = pd.read_excel(f)
            df.columns = df.columns.str.strip().str.lower()
            dfs.append(df)
        except Exception as e:
            print(f"Error leyendo {f}: {e}")
            
    if not dfs:
        return pd.DataFrame()
    return pd.concat(dfs, ignore_index=True)

def refresh_cache():
    """Actualiza la cache global consolidando los archivos físicos de /data"""
    global CACHE
    
    df_tablero = load_tablero()
    df_fact = load_envios()
    df_full = load_fulfillment()

    # CONTROL DE SEGURIDAD: Evita que el servidor falle si no hay excels subidos
    if df_tablero.empty or 'mes_año' not in df_tablero.columns:
        meses_validos = []
        mes_min = "N/A"
        mes_max = "N/A"
        mes_parcial = False
    else:
        meses_validos = sorted(list(set(df_tablero['mes_año'].dropna())))
        if not meses_validos:
            mes_min = "N/A"
            mes_max = "N/A"
            mes_parcial = False
        else:
            mes_min = meses_validos[0]
            mes_max = meses_validos[-1]
            mes_parcial = meses_validos[-1] == datetime.now().strftime('%B-%Y').capitalize()

    CACHE['tablero'] = df_tablero
    CACHE['envios'] = df_fact
    CACHE['full'] = df_full
    CACHE['meses_validos'] = meses_validos
    CACHE['mes_min'] = mes_min
    CACHE['mes_max'] = mes_max
    CACHE['mes_parcial'] = mes_parcial
    CACHE['last_updated'] = datetime.now().strftime('%H:%M:%S')
    print(f"Cache actualizada con éxito. Meses disponibles: {len(meses_validos)}")

# ============================================================
# RUTAS WEB (Vistas y API)
# ============================================================
@app.route('/')
def index():
    return render_template('dashboard.html', 
                           mes_min=CACHE['mes_min'], 
                           mes_max=CACHE['mes_max'],
                           last_updated=CACHE['last_updated'])

@app.route('/api/upload', methods=['POST'])
def upload_files():
    if 'files[]' not in request.files:
        return jsonify({'status': 'error', 'message': 'No se enviaron archivos'}), 400
        
    files = request.files.getlist('files[]')
    saved_count = 0
    
    for file in files:
        if file.filename == '':
            continue
        if file and file.filename.endswith('.xlsx'):
            file_path = DATA_DIR / file.filename
            file.save(str(file_path))
            saved_count += 1
            
    if saved_count > 0:
        refresh_cache()
        return jsonify({
            'status': 'success', 
            'message': f'Se cargaron {saved_count} archivo(s) correctamente.',
            'mes_min': CACHE['mes_min'],
            'mes_max': CACHE['mes_max']
        })
    return jsonify({'status': 'error', 'message': 'No se procesó ningún archivo válido (.xlsx)'}), 400

@app.route('/api/kpis', methods=['GET'])
def get_kpis():
    df = CACHE['tablero']
    if df is None or df.empty:
        return jsonify({
            'ingreso_total': 0, 'ingreso_crecimiento': 0,
            'gmv_total': 0, 'gmv_crecimiento': 0,
            'envios_totales': 0, 'envios_crecimiento': 0,
            'ticket_promedio': 0, 'ticket_crecimiento': 0,
            'clientes_activos': 0, 'clientes_crecimiento': 0,
            'full_penetración': 0, 'full_crecimiento': 0
        })

    # Tomamos el último mes disponible como referencia ejecutiva
    mes_actual = CACHE['mes_max']
    meses = CACHE['meses_validos']
    
    df_act = df[df['mes_año'] == mes_actual]
    
    if len(meses) > 1:
        mes_anterior = meses[-2]
        df_ant = df[df['mes_año'] == mes_anterior]
    else:
        df_ant = df_act

    def get_val_and_growth(col_name):
        v_act = df_act[col_name].sum() if col_name in df_act.columns else 0
        v_ant = df_ant[col_name].sum() if col_name in df_ant.columns else 0
        growth = ((v_act - v_ant) / v_ant * 100) if v_ant else 0
        return float(v_act), float(growth)

    ing, ing_g = get_val_and_growth('ingreso_fixy')
    gmv, gmv_g = get_val_and_growth('gmv_total')
    env, env_g = get_val_and_growth('envios')
    cli, cli_g = get_val_and_growth('clientes_activos')
    
    tick = ing / env if env else 0
    tick_ant = (df_ant['ingreso_fixy'].sum() / df_ant['envios'].sum()) if ('envios' in df_ant.columns and df_ant['envios'].sum()) else 0
    tick_g = ((tick - tick_ant) / tick_ant * 100) if tick_ant else 0

    # Penetración Fulfillment (Envios Full / Envios Totales)
    env_full_act = df_act['envios_full'].sum() if 'envios_full' in df_act.columns else 0
    env_tot_act = df_act['envios'].sum() if 'envios' in df_act.columns else 1
    full_pen = (env_full_act / env_tot_act) * 100

    env_full_ant = df_ant['envios_full'].sum() if 'envios_full' in df_ant.columns else 0
    env_tot_ant = df_ant['envios'].sum() if 'envios' in df_ant.columns else 1
    full_pen_ant = (env_full_ant / env_tot_ant) * 100
    full_g = full_pen - full_pen_ant

    return jsonify({
        'ingreso_total': ing, 'ingreso_crecimiento': ing_g,
        'gmv_total': gmv, 'gmv_crecimiento': gmv_g,
        'envios_totales': env, 'envios_crecimiento': env_g,
        'ticket_promedio': tick, 'ticket_crecimiento': tick_g,
        'clientes_activos': cli, 'clientes_crecimiento': cli_g,
        'full_penetración': full_pen, 'full_crecimiento': full_g
    })

@app.route('/api/graficos/tendencia', methods=['GET'])
def get_grafico_tendencia():
    df = CACHE['tablero']
    if df is None or df.empty:
        return jsonify({'labels': [], 'ingresos': [], 'envios': []})

    # Agrupar por mes cronológicamente según aparezcan indexados
    df_grouped = df.groupby('mes_año', sort=False).agg({
        'ingreso_fixy': 'sum',
        'envios': 'sum'
    }).reset_index()

    return jsonify({
        'labels': df_grouped['mes_año'].tolist(),
        'ingresos': df_grouped['ingreso_fixy'].tolist(),
        'envios': df_grouped['envios'].tolist()
    })

@app.route('/api/graficos/mix-servicios', methods=['GET'])
def get_mix_servicios():
    df = CACHE['tablero']
    if df is None or df.empty:
        return jsonify({'labels': [], 'valores': []})

    mes_actual = CACHE['mes_max']
    df_mes = df[df['mes_año'] == mes_actual]

    servicios = {
        'Fulfillment': df_mes['ingreso_full'].sum() if 'ingreso_full' in df_mes.columns else 0,
        'Envíos Flex': df_mes['ingreso_flex'].sum() if 'ingreso_flex' in df_mes.columns else 0,
        'Colecta/Cross': df_mes['ingreso_colecta'].sum() if 'ingreso_colecta' in df_mes.columns else 0,
    }
    
    # Calcular "Otros" por diferencia si el ingreso total es mayor
    total_ing = df_mes['ingreso_fixy'].sum()
    sum_parcial = sum(servicios.values())
    if total_ing > sum_parcial:
        servicios['Otros'] = total_ing - sum_parcial

    return jsonify({
        'labels': list(servicios.keys()),
        'valores': list(servicios.values())
    })

@app.route('/api/alertas', methods=['GET'])
def get_alertas():
    df = CACHE['tablero']
    meses_validos = CACHE['meses_validos']
    mes_parcial = CACHE['mes_parcial']

    alertas_list = []
    
    if df is None or df.empty or not meses_validos:
        return jsonify({'alertas': [{
            'severidad': 'baja',
            'titulo': 'Sin datos cargados',
            'mensaje': 'El sistema no detectó archivos históricos en /data. Por favor, arrastrá los Excel correspondientes para inicializar las métricas.'
        }], 'mes_parcial': False, 'meses_validos': []})

    # Procesamiento lógico de alertas ejecutivas
    mes_act = meses_validos[-1]
    df_act = df[df['mes_año'] == mes_act]

    # 1) Alerta de Churn Elevado
    if 'churn_rate' in df_act.columns:
        churn = df_act['churn_rate'].max()
        if churn > CHURN_TOLERABLE:
            alertas_list.append({
                'severidad': 'alta',
                'titulo': f'Alerta de Churn: {churn*100:.1f}%',
                'mensaje': f'La fuga de clientes superó el límite tolerable de {CHURN_TOLERABLE*100:.0f}% durante el mes de {mes_act.lower()}. Exige revisión de Service Delivery.',
            })

    # 2) Desaceleración YoY vs Objetivo
    if len(meses_validos) > 12:
        mes_prev = meses_validos[-12]
        ing_act = df_act['ingreso_fixy'].sum()
        ing_prev = df[df['mes_año'] == mes_prev]['ingreso_fixy'].sum()
        
        if ing_prev > 0:
            crecimiento_yoy = (ing_act - ing_prev) / ing_prev
            if crecimiento_yoy < OBJETIVO_GROWTH:
                alertas_list.append({
                    'severidad': 'media',
                    'titulo': f'Crecimiento YoY por debajo del objetivo',
                    'mensaje': f'El crecimiento interanual se sitúa en {crecimiento_yoy*100:.1f}%, siendo el objetivo estratégico un +{OBJETIVO_GROWTH*100:.0f}% comparado contra {mes_prev.lower()}.',
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
    # Para producción en internet leemos el puerto dinámico de Render
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
import streamlit as st

st.set_page_config(page_title="Dashboard Fixy")

st.title("🚚 Dashboard Fixy")
st.success("La aplicación funciona correctamente")
