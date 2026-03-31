import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import datetime
import joblib
import numpy as np
import nltk
import dossier_utils as utils

# --- Configuración de la página ---
st.set_page_config(
    page_title="Dossier Intelligence · Transmilenio",
    page_icon="🚌",
    layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap');

    html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; }

    .stApp { background-color: #F7F8FA; }

    #MainMenu, footer, header { visibility: hidden; }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 1100px;
    }

    /* ── HEADER ── */
    .app-header {
        background: linear-gradient(135deg, #B71C1C 0%, #D32F2F 60%, #F44336 100%);
        border-radius: 12px;
        padding: 1.1rem 1.8rem;
        margin-bottom: 1.2rem;
        position: relative;
        overflow: hidden;
        display: flex;
        align-items: center;
        gap: 1.2rem;
    }
    .app-header::before {
        content: '';
        position: absolute;
        top: -30px; right: -30px;
        width: 120px; height: 120px;
        background: rgba(255,255,255,0.06);
        border-radius: 50%;
    }
    .app-header h1 { display: none; }
    .app-header p {
        color: rgba(255,255,255,0.9);
        font-size: 0.92rem;
        font-weight: 500;
        margin: 0;
    }
    .app-header .badge {
        display: inline-block;
        background: rgba(255,255,255,0.15);
        color: #fff;
        font-size: 0.7rem;
        font-weight: 600;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        padding: 0.2rem 0.6rem;
        border-radius: 20px;
        border: 1px solid rgba(255,255,255,0.25);
        white-space: nowrap;
        flex-shrink: 0;
    }

    /* ── TARJETAS ── */
    .card {
        background: #FFFFFF;
        border: 1px solid #E8ECEF;
        border-radius: 12px;
        padding: 1.6rem 1.8rem;
        margin-bottom: 1.2rem;
        box-shadow: 0 1px 4px rgba(183,28,28,0.06);
    }
    .card-title {
        font-size: 0.8rem;
        font-weight: 600;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        color: #D32F2F;
        margin-bottom: 1rem;
    }

    /* ── MÉTRICAS ── */
    .metrics-row {
        display: flex;
        gap: 1rem;
        margin: 1.5rem 0;
    }
    .metric-card {
        flex: 1;
        background: #FFFFFF;
        border: 1px solid #E8ECEF;
        border-radius: 12px;
        padding: 1.4rem 1.6rem;
        text-align: center;
        box-shadow: 0 1px 4px rgba(183,28,28,0.06);
    }
    .metric-card .metric-value {
        font-size: 2.2rem;
        font-weight: 700;
        color: #B71C1C;
        line-height: 1;
        margin-bottom: 0.4rem;
        font-family: 'DM Mono', monospace;
    }
    .metric-card .metric-label {
        font-size: 0.78rem;
        color: #8F6B6B;
        font-weight: 500;
        letter-spacing: 0.03em;
    }
    .metric-card.accent .metric-value { color: #D32F2F; }
    .metric-card.muted  .metric-value { color: #AC9B9B; }

    /* ── BANNER DE ÉXITO ── */
    .success-banner {
        background: linear-gradient(135deg, #B71C1C, #D32F2F);
        border-radius: 12px;
        padding: 1.2rem 1.8rem;
        display: flex;
        align-items: center;
        gap: 1rem;
        margin: 1.5rem 0;
        box-shadow: 0 2px 12px rgba(183,28,28,0.2);
    }
    .success-banner .icon { font-size: 1.6rem; line-height: 1; }
    .success-banner .text strong {
        display: block;
        color: #FFFFFF;
        font-size: 1rem;
        font-weight: 700;
        margin-bottom: 0.15rem;
    }
    .success-banner .text span {
        color: rgba(255,255,255,0.72);
        font-size: 0.85rem;
    }

    /* ── STEPS ── */
    .step {
        display: flex;
        align-items: flex-start;
        gap: 1rem;
        margin-bottom: 0.8rem;
    }
    .step-num {
        min-width: 28px; height: 28px;
        background: #D32F2F;
        color: white;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.78rem;
        font-weight: 700;
        margin-top: 1px;
    }
    .step-text { color: #5A2D2D; font-size: 0.9rem; line-height: 1.6; }

    /* ── FILE UPLOADER ── */
    [data-testid="stFileUploader"] {
        background: #FFFFFF;
        border: 2px dashed #D4B2B2;
        border-radius: 12px;
        padding: 0.5rem;
        transition: border-color 0.2s;
    }
    [data-testid="stFileUploader"]:hover { border-color: #D32F2F; }

    /* ── BOTÓN PRIMARIO ── */
    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #B71C1C, #D32F2F);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.7rem 2rem;
        font-family: 'DM Sans', sans-serif;
        font-size: 0.95rem;
        font-weight: 600;
        letter-spacing: 0.01em;
        transition: all 0.2s;
        box-shadow: 0 2px 8px rgba(183,28,28,0.25);
        width: 100%;
        height: 48px;
    }
    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #8E1515, #B71C1C);
        box-shadow: 0 4px 14px rgba(183,28,28,0.35);
        transform: translateY(-1px);
    }
    .stButton > button[kind="primary"]:disabled {
        background: #D4C5C5;
        box-shadow: none;
        transform: none;
    }

    /* ── BOTÓN DESCARGA ── */
    .stDownloadButton > button {
        background: #FFFFFF;
        color: #B71C1C;
        border: 2px solid #D32F2F;
        border-radius: 10px;
        padding: 0.65rem 1.8rem;
        font-family: 'DM Sans', sans-serif;
        font-size: 0.9rem;
        font-weight: 600;
        transition: all 0.2s;
        width: 100%;
        height: 48px;
    }
    .stDownloadButton > button:hover {
        background: #FEF2F2;
        box-shadow: 0 2px 10px rgba(183,28,28,0.15);
        transform: translateY(-1px);
    }

    /* ── ALERTAS ── */
    .stWarning { background: #FFF8ED; border: 1px solid #F5C97A; border-radius: 10px; }
    .stError   { background: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 10px; }

    /* ── PROGRESS BAR ── */
    .stProgress > div > div > div {
        background: linear-gradient(90deg, #D32F2F, #F44336);
        border-radius: 4px;
    }

    /* ── EXPANDER ── */
    .streamlit-expanderHeader {
        background: #F7F8FA;
        border-radius: 8px;
        font-weight: 500;
        color: #B71C1C;
    }

    /* ── DATAFRAME ── */
    [data-testid="stDataFrame"] {
        border-radius: 10px;
        overflow: hidden;
        border: 1px solid #E8ECEF;
    }

    hr { border: none; border-top: 1px solid #E8ECEF; margin: 1.5rem 0; }
    .stSpinner > div { border-top-color: #D32F2F !important; }

    /* ── FILE STATUS ── */
    .file-status {
        display: flex;
        align-items: center;
        gap: 0.6rem;
        padding: 0.7rem 1rem;
        border-radius: 8px;
        font-size: 0.87rem;
        font-weight: 500;
        margin-top: 0.5rem;
    }
    .file-status.ok      { background: #FEF2F2; color: #B71C1C; border: 1px solid #D4B2B2; }
    .file-status.missing { background: #FFF8ED; color: #92400E; border: 1px solid #F5C97A; }

    .results-header {
        font-size: 1.1rem;
        font-weight: 700;
        color: #B71C1C;
        margin: 1.8rem 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 2px solid #E8ECEF;
    }
</style>
""", unsafe_allow_html=True)

# --- Descarga NLTK stopwords ---
try:
    nltk.data.find('corpora/stopwords')
except LookupError:
    with st.spinner("Descargando recursos de lenguaje por primera vez..."):
        nltk.download('stopwords')


# ==============================================================================
# FUNCIONES DE CARGA Y PROCESO PRINCIPAL
# ==============================================================================

@st.cache_resource
def load_ml_models():
    try:
        sentiment_pipeline = joblib.load('pipeline_sentimiento_tm.pkl')
        topic_pipeline = joblib.load('pipeline_tema_tm.pkl')
        return sentiment_pipeline, topic_pipeline
    except FileNotFoundError as e:
        st.error(
            f"**Error Crítico:** No se encontró `{e.filename}`. "
            "Asegúrate de que `pipeline_sentimiento_tm.pkl` y `pipeline_tema_tm.pkl` "
            "estén en la misma carpeta que esta app."
        )
        st.stop()


def read_dossier(dossier_file):
    """
    Lee el dossier SIN expandir menciones.
    Para Transmilenio la mención suele ser siempre la misma empresa,
    así que no se hace split por punto y coma.
    """
    wb = load_workbook(dossier_file)
    sheet = wb.active
    original_headers = [cell.value for cell in sheet[1] if cell.value]

    rows_data = []
    for row in sheet.iter_rows(min_row=2, values_only=False):
        if all(c.value is None for c in row):
            continue

        row_values = {}
        for i, header in enumerate(original_headers):
            if header in ['Link Nota', 'Link (Streaming - Imagen)']:
                row_values[header] = utils.extract_link_from_cell(row[i])
            else:
                row_values[header] = row[i].value

        rows_data.append(row_values)

    return pd.DataFrame(rows_data)


def run_full_process(dossier_file, config_file, download_placeholder):
    """
    Ejecuta el proceso completo para Transmilenio.
    """
    st.markdown("<hr>", unsafe_allow_html=True)
    progress_bar = st.progress(0, text="Iniciando proceso...")

    # --- 1. Modelos y configuración ---
    progress_bar.progress(5, text="Paso 1 / 8 — Cargando modelos y configuración...")
    sentiment_pipeline, topic_pipeline = load_ml_models()

    try:
        config_sheets = pd.read_excel(config_file, sheet_name=None)
        region_map = pd.Series(
            config_sheets['Regiones'].iloc[:, 1].values,
            index=config_sheets['Regiones'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
        internet_map = pd.Series(
            config_sheets['Internet'].iloc[:, 1].values,
            index=config_sheets['Internet'].iloc[:, 0].astype(str).str.lower().str.strip()
        ).to_dict()
        mention_map = pd.Series(
            config_sheets['Menciones'].iloc[:, 1].values,
            index=config_sheets['Menciones'].iloc[:, 0].astype(str).str.strip()
        ).to_dict()
        final_topic_map = pd.Series(
            config_sheets['Mapa_Temas'].iloc[:, 1].values,
            index=config_sheets['Mapa_Temas'].iloc[:, 0].astype(str).str.strip()
        ).to_dict()
    except Exception as e:
        st.error(f"**Error al cargar `Configuracion.xlsx`:** {e}")
        st.stop()

    # --- 2. Lectura (sin expansión de menciones) ---
    progress_bar.progress(15, text="Paso 2 / 8 — Leyendo Dossier...")
    df = read_dossier(dossier_file)

    # --- 3. Limpieza y normalización ---
    progress_bar.progress(25, text="Paso 3 / 8 — Aplicando mapeos y normalizaciones...")
    if 'Fecha' in df.columns:
        df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce', dayfirst=True)
        if df['Fecha'].isna().any():
            st.warning("⚠️ Algunas fechas no se pudieron convertir. Revisa el archivo original.")

    if 'Título' in df.columns:
        df['Título'] = df['Título'].apply(utils.clean_title)

    if 'Resumen - Aclaracion' in df.columns:
        df['Resumen - Aclaracion'] = df['Resumen - Aclaracion'].apply(utils.corregir_resumen)

    tipo_medio_map = {
        'online': 'Internet', 'diario': 'Prensa',
        'am': 'Radio', 'fm': 'Radio',
        'aire': 'Televisión', 'cable': 'Televisión',
        'revista': 'Revistas'
    }
    df['Tipo de Medio'] = (
        df['Tipo de Medio'].str.lower().str.strip()
        .map(tipo_medio_map)
        .fillna(df['Tipo de Medio'])
    )

    if 'Medio' in df.columns:
        df['Región'] = df['Medio'].astype(str).str.lower().str.strip().map(region_map)

    if 'Menciones - Empresa' in df.columns:
        df['Menciones - Empresa'] = (
            df['Menciones - Empresa'].astype(str).str.strip()
            .map(mention_map)
            .fillna(df['Menciones - Empresa'])
        )

    is_internet = df['Tipo de Medio'] == 'Internet'
    if is_internet.any():
        df.loc[is_internet, 'Medio'] = (
            df.loc[is_internet, 'Medio'].astype(str).str.lower().str.strip()
            .map(internet_map)
            .fillna(df.loc[is_internet, 'Medio'])
        )

    # --- 4. Reorganización de columnas ---
    progress_bar.progress(40, text="Paso 4 / 8 — Reorganizando columnas de links y dimensiones...")
    is_print     = df['Tipo de Medio'].isin(['Prensa', 'Revistas'])
    is_broadcast = df['Tipo de Medio'].isin(['Radio', 'Televisión'])

    if 'Link Nota' in df.columns and 'Link (Streaming - Imagen)' in df.columns:
        df.loc[is_internet, ['Link Nota', 'Link (Streaming - Imagen)']] = (
            df.loc[is_internet, ['Link (Streaming - Imagen)', 'Link Nota']].values
        )
        cond_copy = is_print & df['Link Nota'].isnull() & df['Link (Streaming - Imagen)'].notnull()
        df.loc[cond_copy, 'Link Nota'] = df.loc[cond_copy, 'Link (Streaming - Imagen)']
        df.loc[is_print | is_broadcast, 'Link (Streaming - Imagen)'] = None

    if 'Duración - Nro. Caracteres' in df.columns and 'Dimensión' in df.columns:
        df.loc[is_broadcast, 'Dimensión'] = df.loc[is_broadcast, 'Duración - Nro. Caracteres']
        df.loc[is_broadcast, 'Duración - Nro. Caracteres'] = np.nan

    # --- 5. Detección de duplicados ---
    progress_bar.progress(50, text="Paso 5 / 8 — Detectando duplicados...")
    df = utils.detect_duplicates_optimized(df)

    # --- 6. Modelos de IA ---
    progress_bar.progress(70, text="Paso 6 / 8 — Aplicando modelos de IA a noticias únicas...")
    df_valid = df[~df['is_duplicate']].copy()
    if not df_valid.empty:
        df_valid['texto_para_ia'] = (
            df_valid['Título'].fillna('') + ' ' + df_valid['Resumen - Aclaracion'].fillna('')
        )
        preds_sent = sentiment_pipeline.predict(df_valid['texto_para_ia'])
        label_map_inv = {1: 'Positivo', 0: 'Neutro', -1: 'Negativo'}
        df_valid['Tono'] = [label_map_inv.get(p, 'Indefinido') for p in preds_sent]

        df_valid['resumen_procesado'] = df_valid['texto_para_ia'].apply(
            utils.preprocess_text_for_topic
        )
        df_valid['Temas Generales - Tema'] = topic_pipeline.predict(df_valid['resumen_procesado'])
        df.update(df_valid[['Tono', 'Temas Generales - Tema']])

    # --- 7. Homogeneización de temas ---
    progress_bar.progress(85, text="Paso 7 / 8 — Homogeneizando y mapeando temas...")
    df_valid_homog = df[~df['is_duplicate']].copy()
    if not df_valid_homog.empty and 'Temas Generales - Tema' in df_valid_homog.columns:
        df_valid_homog['titulo_norm_homog'] = df_valid_homog['Título'].apply(
            utils.normalize_title_for_comparison
        )
        homogenized_temas = df_valid_homog.groupby(
            'titulo_norm_homog'
        )['Temas Generales - Tema'].transform(
            lambda x: x.mode()[0] if not x.mode().empty else x
        )
        df_valid_homog['Temas Generales - Tema'] = homogenized_temas
        df.update(df_valid_homog[['Temas Generales - Tema']])

    if 'Temas Generales - Tema' in df.columns:
        df['Tema'] = (
            df['Temas Generales - Tema'].astype(str).str.strip()
            .map(final_topic_map)
            .fillna('Indefinido')
        )

    # --- Marcado final de duplicadas ---
    mask_dup = df['is_duplicate']
    if mask_dup.any():
        if 'Temas Generales - Tema' in df.columns:
            df.loc[mask_dup, 'Temas Generales - Tema'] = '-'
        if 'Tema' in df.columns:
            df.loc[mask_dup, 'Tema'] = '-'
        df.loc[mask_dup, 'Tono'] = 'Duplicada'

    # --- 8. Resultados ---
    progress_bar.progress(100, text="✓ Proceso completado")

    final_order = [
        "ID Noticia", "Fecha", "Hora", "Medio", "Tipo de Medio", "Sección - Programa",
        "Región", "Título", "Autor - Conductor", "Nro. Pagina", "Dimensión",
        "Duración - Nro. Caracteres", "CPE", "Tier", "Audiencia", "Tono", "Tema",
        "Temas Generales - Tema", "Resumen - Aclaracion", "Link Nota",
        "Link (Streaming - Imagen)", "Menciones - Empresa"
    ]

    total       = len(df)
    dups_count  = int(mask_dup.sum())
    unique_count = total - dups_count
    filename    = f"Dossier_Transmilenio_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    excel_data  = utils.to_excel_from_df(df, final_order)

    # ── Botón de descarga ──
    with download_placeholder:
        st.download_button(
            label="⬇ Descargar archivo procesado (.xlsx)",
            data=excel_data,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Banner de éxito
    st.markdown(f"""
    <div class="success-banner">
        <div class="icon">🚌</div>
        <div class="text">
            <strong>Proceso finalizado correctamente</strong>
            <span>{total:,} filas procesadas · {unique_count:,} únicas · {dups_count:,} duplicadas</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Métricas
    st.markdown('<p class="results-header">Resumen del proceso</p>', unsafe_allow_html=True)
    st.markdown(f"""
    <div class="metrics-row">
        <div class="metric-card">
            <div class="metric-value">{total:,}</div>
            <div class="metric-label">Filas totales procesadas</div>
        </div>
        <div class="metric-card accent">
            <div class="metric-value">{unique_count:,}</div>
            <div class="metric-label">Noticias únicas analizadas</div>
        </div>
        <div class="metric-card muted">
            <div class="metric-value">{dups_count:,}</div>
            <div class="metric-label">Filas marcadas como duplicadas</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Previsualización
    st.markdown('<p class="results-header">Previsualización de resultados</p>', unsafe_allow_html=True)
    final_cols_in_df = [col for col in final_order if col in df.columns]
    df_display = df[final_cols_in_df].copy()

    if 'Fecha' in df_display.columns:
        df_display['Fecha'] = (
            pd.to_datetime(df_display['Fecha'])
            .dt.strftime('%d/%m/%Y')
            .replace('NaT', 'FECHA INVÁLIDA')
        )

    st.dataframe(df_display, use_container_width=True, hide_index=True)


# ==============================================================================
# INTERFAZ PRINCIPAL
# ==============================================================================

st.markdown("""
<div class="app-header">
    <div class="badge">Transmilenio · Media Intelligence</div>
    <p>Limpieza, enriquecimiento y análisis automático de dossiers · v1.0</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="card">
    <div class="card-title">Cómo usar esta herramienta</div>
    <div class="step">
        <div class="step-num">1</div>
        <div class="step-text">Prepara tu archivo <strong>Dossier</strong> (.xlsx) y el archivo <strong>Configuracion.xlsx</strong>.</div>
    </div>
    <div class="step">
        <div class="step-num">2</div>
        <div class="step-text">Sube ambos archivos en el área de carga de abajo. El sistema los detecta automáticamente.</div>
    </div>
    <div class="step">
        <div class="step-num">3</div>
        <div class="step-text">Haz clic en <strong>Iniciar proceso</strong>. Al finalizar, el botón de descarga aparecerá justo a su lado.</div>
    </div>
</div>
""", unsafe_allow_html=True)

with st.expander("📋  Ver estructura requerida para Configuracion.xlsx"):
    st.markdown("""
    | Hoja | Columna A | Columna B |
    |------|-----------|-----------|
    | `Regiones` | Medio | Región |
    | `Internet` | Medio Original | Medio Mapeado |
    | `Menciones` | Mención Original | Mención Mapeada |
    | `Mapa_Temas` | Temas Generales - Tema | Tema |
    """)

st.markdown("<br>", unsafe_allow_html=True)

# Carga de archivos
st.markdown('<div class="card-title">Carga de archivos</div>', unsafe_allow_html=True)
uploaded_files = st.file_uploader(
    "Arrastra los archivos aquí o haz clic para seleccionarlos",
    type=["xlsx"],
    accept_multiple_files=True,
    label_visibility="collapsed"
)

dossier_file, config_file = None, None

if uploaded_files:
    for file in uploaded_files:
        if 'config' in file.name.lower():
            config_file = file
        else:
            dossier_file = file

    col_a, col_b = st.columns(2)
    with col_a:
        if dossier_file:
            st.markdown(
                f'<div class="file-status ok">✓ Dossier cargado — <strong>{dossier_file.name}</strong></div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="file-status missing">⚠ No se detectó el archivo Dossier</div>',
                unsafe_allow_html=True
            )
    with col_b:
        if config_file:
            st.markdown(
                f'<div class="file-status ok">✓ Configuración cargada — <strong>{config_file.name}</strong></div>',
                unsafe_allow_html=True
            )
        else:
            st.markdown(
                '<div class="file-status missing">⚠ No se detectó Configuracion.xlsx</div>',
                unsafe_allow_html=True
            )

st.markdown("<br>", unsafe_allow_html=True)

# ── Fila de botones ──
col_start, col_download = st.columns(2)

with col_start:
    start_clicked = st.button(
        "▶  Iniciar proceso completo",
        disabled=not (dossier_file and config_file),
        type="primary"
    )

with col_download:
    download_placeholder = st.empty()
    if not start_clicked:
        with download_placeholder:
            st.button(
                "⬇ Descargar archivo procesado (.xlsx)",
                disabled=True,
                type="primary"
            )

if start_clicked:
    run_full_process(dossier_file, config_file, download_placeholder)
