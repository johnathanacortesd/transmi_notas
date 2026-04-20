<div align="center">

<br/>

```
████████╗██████╗  █████╗ ███╗   ██╗███████╗███╗   ███╗██╗ █████╗ ██████╗ ██████╗ 
╚══██╔══╝██╔══██╗██╔══██╗████╗  ██║██╔════╝████╗ ████║██║██╔══██╗██╔══██╗██╔══██╗
   ██║   ██████╔╝███████║██╔██╗ ██║███████╗██╔████╔██║██║███████║██████╔╝██████╔╝
   ██║   ██╔══██╗██╔══██║██║╚██╗██║╚════██║██║╚██╔╝██║██║██╔══██║██╔═══╝ ██╔═══╝ 
   ██║   ██║  ██║██║  ██║██║ ╚████║███████║██║ ╚═╝ ██║██║██║  ██║██║     ██║     
   ╚═╝   ╚═╝  ╚═╝╚═╝  ╚═╝╚═╝  ╚═══╝╚══════╝╚═╝     ╚═╝╚═╝╚═╝  ╚═╝╚═╝     ╚═╝     
```

### 🚇 Limpieza y análisis automático de dossiers de prensa · Transmilenio

<br/>

![Python](https://img.shields.io/badge/Python-3.10+-B71C1C?style=for-the-badge&logo=python&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.x-D32F2F?style=for-the-badge&logo=streamlit&logoColor=white)
![scikit-learn](https://img.shields.io/badge/scikit--learn-ML-FF7043?style=for-the-badge&logo=scikit-learn&logoColor=white)
![XlsxWriter](https://img.shields.io/badge/XlsxWriter-Excel-C62828?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![Estado](https://img.shields.io/badge/Estado-Activo-4CAF50?style=for-the-badge)

[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://transmi-notas.streamlit.app/)

<br/>

> **TransmiApp** automatiza el procesamiento de dossiers de prensa del sistema de transporte masivo de Bogotá. Detecta duplicados, predice tono y tema con modelos de IA entrenados, normaliza medios y expande menciones de entidades — todo desde una interfaz web sin fricciones.

<br/>

</div>

---

## 🗺️ Tabla de contenidos

- [¿Qué hace?](#-qué-hace)
- [Demo en vivo](#-demo-en-vivo)
- [Arquitectura](#-arquitectura)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Estructura de archivos](#-estructura-de-archivos)
- [Configuración](#-configuración)
- [Modelos de IA](#-modelos-de-ia)
- [Lógica de duplicados](#-lógica-de-duplicados)
- [Stack técnico](#-stack-técnico)
- [Autor](#-autor)

---

## ✨ ¿Qué hace?

TransmiApp tiene **dos flujos de trabajo** principales, accesibles desde pestañas separadas:

### 🔄 Pestaña 1 — Procesar Dossier (flujo completo con IA)

| Paso | Descripción |
|------|-------------|
| **1. Lectura** | Parsea el archivo `.xlsx` preservando hipervínculos embebidos en celdas |
| **2. Normalización** | Unifica nombres de medios, regiones, tipos de medio y formatos de fecha |
| **3. Detección de duplicados** | Algoritmo optimizado con ventanas por grupo (medio + mención), similitud de títulos ≥ 85% y proximidad de fechas |
| **4. Predicción de Tono** | Modelo de ML (pipeline sklearn) entrena sobre texto limpio → `Positivo / Neutro / Negativo` |
| **5. Predicción de Tema** | Segundo pipeline que clasifica en categorías temáticas predefinidas |
| **6. Homogeneización** | Agrupa noticias similares y aplica la moda de tono/tema al grupo |
| **7. Reglas de negocio** | Ej: noticias con "titulares" → Tono `Neutro`, Tema `Entorno e información general` |
| **8. Exportación** | Genera `.xlsx` con hipervínculos activos, formato de fecha `dd/mm/yyyy` y orden de columnas definido |

### 📋 Pestaña 2 — Expandir por Entidades (sin IA)

Cuando una noticia menciona múltiples empresas separadas por `;`, **crea una fila por cada mención** — ideal para reportes por entidad. Incluye mapeo de alias antes de expandir.

---

## 🎬 Demo en vivo

> _Agrega aquí el link a tu Streamlit Cloud deployment o un GIF demostrativo_

```
https://transmiapp.streamlit.app  ← reemplaza con tu URL
```

---

## 🏗️ Arquitectura

```
┌─────────────────────────────────────────────────┐
│                  app.py (UI)                    │
│           Streamlit · 2 pestañas                │
└────────────┬──────────────────┬─────────────────┘
             │                  │
    ┌────────▼────────┐ ┌───────▼──────────┐
    │  run_full_      │ │ run_expand_      │
    │  process()      │ │ process()        │
    │  (con IA)       │ │ (sin IA)         │
    └────────┬────────┘ └───────┬──────────┘
             │                  │
    ┌────────▼──────────────────▼──────────┐
    │           dossier_utils.py            │
    │  · read_dossier()                    │
    │  · apply_common_transformations()    │
    │  · detect_duplicates_optimized()     │
    │  · expand_by_mentions()              │
    │  · limpiar_texto()  (NLP)            │
    │  · to_excel_from_df()               │
    └──────────────────────────────────────┘
             │
    ┌────────▼──────────────────────────┐
    │  Modelos ML (.pkl)                │
    │  pipeline_sentimiento_tm.pkl      │
    │  pipeline_tema_tm.pkl             │
    └───────────────────────────────────┘
```

---

## 🚀 Instalación

### Prerrequisitos

- Python **3.10+**
- `pip`

### Pasos

```bash
# 1. Clonar el repositorio
git clone https://github.com/tu-usuario/transmiapp.git
cd transmiapp

# 2. Crear entorno virtual (recomendado)
python -m venv .venv
source .venv/bin/activate        # Linux / macOS
.venv\Scripts\activate           # Windows

# 3. Instalar dependencias
pip install -r requirements.txt

# 4. Verificar que los modelos .pkl estén en la raíz
ls *.pkl
# pipeline_sentimiento_tm.pkl
# pipeline_tema_tm.pkl

# 5. Lanzar la app
streamlit run app.py
```

La app abre en `http://localhost:8501` 🎉

---

## 📖 Uso

### Archivos requeridos

Necesitas dos archivos `.xlsx` en cada proceso:

#### `Dossier.xlsx` (cualquier nombre sin "config")
El archivo de noticias exportado del sistema de monitoreo. Columnas clave:

```
ID Noticia | Fecha | Hora | Tipo de Medio | Medio | Título | Resumen - Aclaracion
Menciones - Empresa | Link Nota | Link (Streaming - Imagen) | Dimensión | ...
```

#### `Configuracion.xlsx`
Tablas de mapeo. Estructura por hoja:

| Hoja | Col A | Col B | Requerida en |
|------|-------|-------|--------------|
| `Regiones` | Nombre del medio | Región | Ambas pestañas |
| `Internet` | Medio original | Medio normalizado | Ambas pestañas |
| `Menciones` | Mención original | Mención normalizada | Solo Pestaña 2 |

### Flujo de uso

```
1. Arrastra los dos archivos al uploader
         ↓
2. La app detecta cuál es cuál por el nombre
   (el que tiene "config" → Configuracion.xlsx)
         ↓
3. Haz clic en ▶ Iniciar proceso
         ↓
4. Progreso visible paso a paso
         ↓
5. ⬇ Botón de descarga aparece al finalizar
```

---

## 📁 Estructura de archivos

```
transmiapp/
│
├── app.py                          # Interfaz Streamlit · lógica de UI
├── dossier_utils.py                # Toda la lógica de negocio y NLP
│
├── pipeline_sentimiento_tm.pkl     # Modelo de análisis de tono (joblib)
├── pipeline_tema_tm.pkl            # Modelo de clasificación temática (joblib)
│
├── requirements.txt                # Dependencias Python
└── README.md                       # Este archivo
```

---

## ⚙️ Configuración

### Mapeo de tipos de medio

Definido en `app.py` → `TIPO_MEDIO_MAP`:

```python
TIPO_MEDIO_MAP = {
    'online':  'Internet',
    'diario':  'Prensa',
    'am':      'Radio',
    'fm':      'Radio',
    'aire':    'Televisión',
    'cable':   'Televisión',
    'revista': 'Revistas'
}
```

### Orden de columnas en el output

**Pestaña 1 (Dossier completo):** incluye `Mes`, `Tono`, `Temas Generales - Tema`, `Resumen - Aclaracion`, etc.

**Pestaña 2 (Expandido):** columnas reducidas al mínimo necesario para análisis por entidad.

Edita `FINAL_ORDER_TAB1` / `FINAL_ORDER_TAB2` en `app.py` para personalizar.

---

## 🧠 Modelos de IA

Los modelos son **pipelines de scikit-learn** serializados con `joblib`.

### Preprocesamiento de texto (`limpiar_texto`)

```
Texto crudo (Título + Resumen)
        ↓
Lowercase + unidecode (normalizar acentos)
        ↓
Eliminar URLs, @menciones, #hashtags
        ↓
Eliminar caracteres no alfabéticos
        ↓
Eliminar stopwords en español
        ↓
Reemplazar sinónimos (tm → transmilenio, sitp, etc.)
        ↓
Refuerzo de señal:
  · Palabras positivas clave → ×4 (3 repeticiones extra)
  · Palabras neutras/logísticas → ×3 (2 repeticiones extra)
        ↓
Texto limpio para TF-IDF
```

> **¿Por qué el refuerzo de señal?** El TF-IDF aplica un logaritmo que amortigua términos frecuentes. Repetir palabras clave compensate ese efecto y mejora la precisión en textos cortos de noticias.

### Homogeneización por similitud

Tras la predicción, las noticias con títulos similares (mismo grupo) reciben el **tono y tema mayoritario** del grupo (moda), lo que reduce inconsistencias entre notas casi idénticas de distintos medios.

### Reglas de negocio post-predicción

- Título contiene `"titulares"` → Tono `Neutro` · Tema `Entorno e información general`
- Filas marcadas como duplicadas → Tono `Duplicada` · Tema `-`

---

## 🔍 Lógica de duplicados

El algoritmo es **O(n log n)** gracias al agrupamiento previo:

```
1. Calcular quality score del título
   (penaliza HTML entities, títulos cortos, saltos de línea, etc.)

2. Ordenar por (quality DESC, fecha ASC, índice ASC)
   → El "mejor" título del grupo sobrevive

3. Agrupar por [Medio, Menciones - Empresa]
   → Solo compara pares dentro del mismo medio/mención

4. Para cada par dentro del grupo:
   a. ¿Las fechas están cerca? (criterio por tipo de medio)
      · Internet: ±1 día
      · Radio/TV: mismo día Y misma hora
      · Prensa/Revistas: mismo día
   b. ¿Las URLs son iguales? (solo Internet) → DUPLICADO
   c. ¿Similitud de título ≥ 85%? → DUPLICADO
      (también detecta si un título contiene al otro)
```

---

## 🛠️ Stack técnico

| Componente | Librería |
|-----------|----------|
| UI / Web app | `streamlit` |
| Procesamiento Excel | `openpyxl`, `xlsxwriter` |
| DataFrames | `pandas`, `numpy` |
| Modelos ML | `scikit-learn`, `joblib` |
| NLP | `unidecode`, `re`, `difflib` |
| Fechas | `datetime` |

### `requirements.txt` sugerido

```txt
streamlit>=1.32.0
pandas>=2.0.0
numpy>=1.26.0
openpyxl>=3.1.0
xlsxwriter>=3.1.0
scikit-learn>=1.4.0
joblib>=1.3.0
unidecode>=1.3.6
```

---

## 👤 Autor

<div align="center">

**Johnathan Cortés** 🕵️😼

_Analista de datos · Bogotá, Colombia_

[![GitHub](https://img.shields.io/badge/GitHub-johnathanacortesd-B71C1C?style=flat-square&logo=github)](https://github.com/johnathanacortesd)

<br/>

> _"Menos Excel manual, más datos que hablan."_

<br/>

---

<sub>Construido con ❤️ y mucho café para el sistema de transporte de Bogotá · 2026</sub>

</div>
