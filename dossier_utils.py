import re
import html
import io
import pandas as pd
from difflib import SequenceMatcher
from unidecode import unidecode

# ==============================================================================
# CONSTANTES DE PREPROCESAMIENTO IA (Idénticas al modelo de Colab)
# ==============================================================================

STOPWORDS_ES = set("""
a ante bajo cabe con contra de desde durante en entre hacia hasta mediante
para por segun sin so sobre tras y o u e la el los las un una unos unas
lo al del se su sus le les mi mis tu tus nuestro nuestros vuestra vuestras
este esta estos estas ese esa esos esas aquel aquella aquellos aquellas
que cual cuales quien quienes cuyo cuya cuyos cuyas como cuando donde es
son fue fueron era eran sera seran seria serian he ha han habia habrian
hay hubo habra habria estoy esta estan estaba estaban estamos estar estare
estaria estuvieron estarian estuvo asi ya mas menos tan tanto cada
""".split())

SINONIMOS_TM = {
    'tm': 'transmilenio',
    'transmi': 'transmilenio',
    'alimentador': 'sitp',
    'zonal': 'sitp',
    'troncal': 'transmilenio'
}

# PALABRAS POSITIVAS (Multiplicador x4)
PALABRAS_POSITIVAS_CLAVE = [
    # mejoras generales
    'mejora', 'mejoras', 'mejor_servicio', 'mejor_experiencia',
    'optimiza', 'optimizacion', 'avance', 'avances',
    'modernizacion', 'renovacion', 'expansion', 'ampliacion',

    # operación
    'puntualidad', 'puntual', 'frecuencia', 'frecuencia_alta',
    'menor_espera', 'fluido', 'fluidez', 'continuidad',
    'menos_congestion', 'flujo_constante',
    'reduce_tiempos', 'ahorro_tiempo', 'viajes_mas_rapidos',

    # cobertura y rutas
    'nuevas_rutas', 'rutas_extensas', 'cobertura_amplia',
    'mejor_conectividad', 'integracion', 'sistema_integrado',

    # infraestructura
    'nuevas_estaciones', 'estaciones_renovadas',
    'infraestructura_nueva', 'carriles_exclusivos',
    'mantenimiento', 'rehabilitacion', 'troncales_ampliadas',

    # tecnología
    'pago_digital', 'recarga_online', 'app_oficial',
    'informacion_tiempo_real', 'paneles_informativos',
    'wifi', 'sistema_inteligente', 'digitalizacion',

    # seguridad
    'seguridad_reforzada', 'vigilancia', 'camaras',
    'control_policial', 'reduccion_delito', 'convivencia',

    # ambiente
    'bajas_emisiones', 'cero_emisiones', 'aire_limpio',
    'flota_electrica', 'energia_limpia', 'reduccion_co2',

    # experiencia usuario
    'comodidad_mejorada', 'menos_hacinamiento',
    'acceso_facil', 'mejor_senalizacion',
    'orden_estaciones', 'atencion_usuario',

    # gestión
    'gestion_eficiente', 'transparencia', 'cumplimiento',
    'avance_proyecto', 'financiacion', 'resultados',

    # impacto social
    'empleo', 'desarrollo', 'movilidad_mejorada',
    'calidad_de_vida', 'integracion_ciudad',

    # frases comunes en noticias
    'entra_en_operacion',
    'beneficia_a_miles',
    'mejora_la_movilidad',
    'fortalece_el_sistema',
    'reduce_tiempos_de_viaje',
    'mejora_el_servicio',
    'aumenta_la_cobertura',
    'optimiza_los_recorridos',
    'facilita_el_transporte',
    'mejora_la_experiencia_del_usuario',
    'incrementa_la_seguridad',
    'moderniza_el_sistema',
    'amplia_la_capacidad',
    'reduce_la_congestion',
    'mejora_la_conectividad',
    'implementa_tecnologia',
    'renueva_la_flota',
    'incorpora_buses_electricos'
]

# NUEVO: PALABRAS NEUTRAS / INFORMATIVAS / LOGÍSTICAS (Multiplicador x3)
# Ahoga el sesgo negativo cuando se reportan trámites u operativos.
PALABRAS_NEUTRAS_CLAVE = [
    'tramite', 'tramites', 'inscripcion', 'inscripciones', 'habilitado','votacion','portales',
    'habilitada', 'habilitados', 'habilitadas', 'habilitaron', 'puntos', 'punto','plazo',
    'jornada', 'jornadas', 'proceso', 'procesos', 'registro', 'registros','votación','portal',
    'informacion', 'servicio', 'servicios', 'operacion', 'horario', 'horarios','Plazo para inscripción de cédulas',
    'ruta', 'rutas', 'movilidad', 'censo', 'electoral', 'ciudadanos', 'espacios',
    'comerciales', 'portales', 'instalaron', 'moviles', 'campana', 'campanas',
    'actividad', 'actividades', 'logistica'
]

# --- Funciones de limpieza de texto ---

def limpiar_texto(texto: str) -> str:
    """Limpia, normaliza el texto y aplica reglas de negocio para Transporte Público.
       ESTA FUNCIÓN DEBE SER IDÉNTICA A LA USADA EN EL ENTRENAMIENTO."""
    if pd.isna(texto) or texto == "":
        return ""

    # Convertir a string y lowercase
    texto = str(texto).lower()

    # Normalizar caracteres especiales (acentos)
    texto = unidecode(texto)

    # Remover URLs, menciones y hashtags
    texto = re.sub(r'http\S+|www\S+', '', texto)
    texto = re.sub(r'@\w+|#\w+', '', texto)

    # Remover caracteres especiales y números
    texto = re.sub(r'[^a-z\s]', ' ', texto)

    # Normalizar espacios
    texto = re.sub(r'\s+', ' ', texto).strip()

    # Procesar palabras
    palabras = []
    for p in texto.split():
        # Unificar sinónimos (ej: tm -> transmilenio)
        if p in SINONIMOS_TM:
            p = SINONIMOS_TM[p]
            
        if p not in STOPWORDS_ES and len(p) > 2:
            palabras.append(p)
            
            # REFUERZO AGRESIVO:
            # Multiplicamos x4 (añadimos 3 veces más) para vencer el logaritmo del TF-IDF
            if p in PALABRAS_POSITIVAS_CLAVE:
                palabras.extend([p] * 3) 
            
            # REFUERZO NEUTRO:
            # Multiplicamos x3 (añadimos 2 veces más) para contexto logístico
            elif p in PALABRAS_NEUTRAS_CLAVE:
                palabras.extend([p] * 2)

    return ' '.join(palabras)


def convert_html_entities(text: str) -> str:
    if not isinstance(text, str):
        return text
    text = html.unescape(text)
    custom_replacements = {
        '\u201c': '"', '\u201d': '"', '\u2018': "'", '\u2019': "'",
        'Â': '', 'â': '', '€': '', '™': '', '\x9d': '',
        '\xa0': ' '
    }
    for entity, char in custom_replacements.items():
        text = text.replace(entity, char)
    return text


def clean_title(title: str) -> str:
    if not isinstance(title, str):
        return ""
    return convert_html_entities(title)


def clean_title_for_output(title: str) -> str:
    if not isinstance(title, str):
        return ""
    title = convert_html_entities(title)
    title = title.replace('\n', ' ').replace('\r', ' ')
    title = re.sub(r'\s+\|.*$', '', title)
    title = re.sub(r'\|\s+.*$', '', title)
    title = re.sub(r'\s+-\s+.*$', '', title)
    return title.strip()


def normalize_title_for_comparison(title: str) -> str:
    if not isinstance(title, str):
        return ""
    cleaned_title = clean_title_for_output(title)
    abbreviations = {
        'tm': 'transmilenio',
        'tmsa': 'transmilenio',
        'sitp': 'sistema integrado de transporte publico',
    }
    for abbr, full_text in abbreviations.items():
        cleaned_title = re.sub(
            fr'\b{abbr}\b', full_text, cleaned_title, flags=re.IGNORECASE
        )
    normalized_title = re.sub(r'\W+', ' ', cleaned_title).lower().strip()
    return normalized_title


def normalize_url(url) -> str:
    if not isinstance(url, str):
        return ""
    url = url.strip().lower().rstrip('/')
    if not url.startswith('http'):
        return ""
    url = re.sub(r'^(https?://)www\.', r'\1', url)
    url = re.split(r'[?#]', url)[0].rstrip('/')
    return url


def corregir_resumen(text: str) -> str:
    if not isinstance(text, str):
        return text
    text = convert_html_entities(text)
    text = re.sub(r'(<br\s*/?>|\[\.\.\.\]|\s+)', ' ', text).strip()
    match = re.search(r'[A-Z]', text)
    if match:
        text = text[match.start():]
    if text and not text.endswith('...'):
        text = text.rstrip('.') + '...'
    return text


# --- Funciones de Excel y DataFrame ---

def extract_link_from_cell(cell):
    if cell.hyperlink and cell.hyperlink.target:
        return cell.hyperlink.target
    return cell.value


def to_excel_from_df(df: pd.DataFrame, final_order: list) -> bytes:
    output = io.BytesIO()
    final_columns_in_df = [col for col in final_order if col in df.columns]
    df_to_excel = df[final_columns_in_df]

    with pd.ExcelWriter(
        output,
        engine='xlsxwriter',
        datetime_format='dd/mm/yyyy',
        date_format='dd/mm/yyyy'
    ) as writer:
        df_to_excel.to_excel(writer, index=False, sheet_name='Resultado')
        workbook = writer.book
        worksheet = writer.sheets['Resultado']
        link_format = workbook.add_format({'color': 'black'})

        for col_name in ['Link Nota', 'Link (Streaming - Imagen)']:
            if col_name in df_to_excel.columns:
                col_idx = df_to_excel.columns.get_loc(col_name)
                for row_idx, url in enumerate(df_to_excel[col_name]):
                    if pd.notna(url) and isinstance(url, str) and url.startswith('http'):
                        worksheet.write_url(row_idx + 1, col_idx, url, link_format, 'Link')

    return output.getvalue()


# --- Funciones de Lógica de Negocio (Duplicados) ---

def calculate_title_quality_score(title: str) -> int:
    if not isinstance(title, str):
        return -999
    score = 100
    score -= len(re.findall(r'&[#\w]+;', title)) * 10
    score -= title.count('??') * 5
    score -= title.count('\x9d') * 5
    if len(title) > 250:
        score -= 5
    if len(title) < 15:
        score -= 20
    if '\n' in title:
        score -= 15
    if '|' in title:
        score -= 5
    return int(score)


def are_duplicates(
    row1: pd.Series,
    row2: pd.Series,
    title_similarity_threshold=0.85,
    date_proximity_days=1
) -> bool:
    titulo1 = normalize_title_for_comparison(row1['Título'])
    titulo2 = normalize_title_for_comparison(row2['Título'])
    if not titulo1 or not titulo2:
        return False

    fecha1 = row1['Fecha']
    fecha2 = row2['Fecha']
    if pd.isna(fecha1) or pd.isna(fecha2):
        return False

    tipo_medio = row1['Tipo de Medio']

    if tipo_medio == 'Internet':
        if abs((fecha1 - fecha2).days) > date_proximity_days:
            return False
    elif tipo_medio in ['Radio', 'Televisión']:
        if fecha1.date() != fecha2.date():
            return False
        hora1 = row1.get('Hora')
        hora2 = row2.get('Hora')
        if pd.notna(hora1) and pd.notna(hora2):
            if str(hora1).strip() != str(hora2).strip():
                return False
    else:
        if fecha1.date() != fecha2.date():
            return False

    if tipo_medio == 'Internet':
        url_col = 'Link (Streaming - Imagen)'
        url1 = normalize_url(row1.get(url_col, ''))
        url2 = normalize_url(row2.get(url_col, ''))
        if url1 and url2 and url1 == url2:
            return True

    titles_match = False
    if titulo1 == titulo2:
        titles_match = True
    if not titles_match and len(titulo1) > 15 and len(titulo2) > 15:
        if titulo1 in titulo2 or titulo2 in titulo1:
            titles_match = True
    if not titles_match:
        similarity = SequenceMatcher(None, titulo1, titulo2).ratio()
        if similarity >= title_similarity_threshold:
            titles_match = True

    return titles_match


def detect_duplicates_optimized(df: pd.DataFrame) -> pd.DataFrame:
    df = df.reset_index(drop=True).reset_index().rename(columns={'index': 'original_index'})
    df['title_quality'] = df['Título'].apply(calculate_title_quality_score)

    df.sort_values(
        by=['title_quality', 'Fecha', 'original_index'],
        ascending=[False, True, True],
        inplace=True,
        na_position='last'
    )

    grouping_keys = ['Medio', 'Menciones - Empresa']
    duplicate_indices = set()

    for _, group in df.groupby(grouping_keys, dropna=False):
        if len(group) < 2:
            continue
        group_rows = group.to_dict('records')
        for i in range(len(group_rows)):
            current = group_rows[i]
            if current['original_index'] in duplicate_indices:
                continue
            for j in range(i + 1, len(group_rows)):
                compare = group_rows[j]
                if compare['original_index'] in duplicate_indices:
                    continue
                if are_duplicates(pd.Series(current), pd.Series(compare)):
                    duplicate_indices.add(compare['original_index'])

    df['is_duplicate'] = df['original_index'].isin(duplicate_indices)
    return df.sort_values('original_index').set_index('original_index').drop(columns=['title_quality'])


# --- Función de Expansión por Menciones ---

def expand_by_mentions(df: pd.DataFrame, mention_col: str = 'Menciones - Empresa') -> pd.DataFrame:
    if mention_col not in df.columns:
        return df.copy()

    df = df.copy()
    df[mention_col] = df[mention_col].astype(str).fillna('')

    expanded_rows = []
    for _, row in df.iterrows():
        mentions_raw = row[mention_col]
        mentions = [m.strip() for m in mentions_raw.split(';') if m.strip()]

        if not mentions:
            expanded_rows.append(row)
        elif len(mentions) == 1:
            row_copy = row.copy()
            row_copy[mention_col] = mentions[0]
            expanded_rows.append(row_copy)
        else:
            for mention in mentions:
                row_copy = row.copy()
                row_copy[mention_col] = mention
                expanded_rows.append(row_copy)

    result = pd.DataFrame(expanded_rows).reset_index(drop=True)
    return result
