# ============================================================
# LA TRIBU · ML Toolkit — Estilos y branding
# ------------------------------------------------------------
# Uso:
#   from estilos import aplicar_tema, encabezado
#
#   st.set_page_config(page_title="ML Toolkit", page_icon="📦", layout="wide")
#   aplicar_tema()                       # <- justo después de set_page_config
#   ...
#   encabezado("Actualizar Integraly", "Actualiza precio, stock y estado por SKU.")
# ============================================================

import streamlit as st


# ---- Paleta (usar en otros lados si hace falta) ------------
COLORES = {
    "amarillo": "#FFD400",
    "fondo": "#ff5a4d",
    "panel": "#14171c",
    "borde": "#23272e",
    "sidebar": "#0b0d10",
    "texto": "#ff5a4d",
    "apagado": "#8b95a3",
    "verde": "#3fd07f",
    "rojo": "#ff5a4d",
    "ambar": "#f4b63e",
}


def aplicar_tema():
    """Inyecta las fuentes y todo el CSS de marca. Llamar una vez por sesión."""
    st.markdown(_CSS, unsafe_allow_html=True)


def encabezado(titulo, subtitulo=None):
    """Título de sección con la barra amarilla de LA TRIBU."""
    sub = f'<p class="lt-sub">{subtitulo}</p>' if subtitulo else ""
    st.markdown(
        f"""
        <div class="lt-header">
          <div class="lt-bar"></div>
          <div>
            <h1 class="lt-title">{titulo}</h1>
            {sub}
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def logo_sidebar():
    """Lockup de marca LA TRIBU / ML TOOLKIT para el tope del sidebar."""
    st.sidebar.markdown(
        """
        <div class="lt-brand">
          <div class="lt-brand-mark">LT</div>
          <div>
            <div class="lt-brand-name">LA TRIBU</div>
            <div class="lt-brand-sub">ML TOOLKIT</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# CSS
# ============================================================
_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;600;700;800;900&family=Manrope:wght@400;500;600;700;800&family=IBM+Plex+Mono:wght@400;500;600&display=swap');


/* ---------- Base ---------- */
html, body, [class*="css"], .stApp, [data-testid="stAppViewContainer"] {
    font-family: 'Manrope', sans-serif;
}
.stApp { background: #0f1216; }
::selection { background: #FFD400; color: #0f1216; }

/* Tipografía de títulos */
h1, h2, h3, h4 { font-family: 'Archivo', sans-serif !important; letter-spacing: -0.01em; color: #ffffff; }

/* ---------- Contenedor principal ---------- */
.block-container { padding-top: 2.4rem; padding-bottom: 4rem; max-width: 1180px; }

/* ---------- Sidebar ---------- */
[data-testid="stSidebar"] {
    background: #0b0d10;
    border-right: 1px solid #1c2026;
}
[data-testid="stSidebar"] .block-container { padding-top: 1.6rem; }
[data-testid="stSidebar"] label, [data-testid="stSidebar"] p { color: #b9c1cc; }

/* Marca en sidebar */
.lt-brand { display:flex; align-items:center; gap:11px; padding:2px 2px 18px; }
.lt-brand-mark {
    width:38px; height:38px; border-radius:10px; background:#FFD400;
    display:flex; align-items:center; justify-content:center;
    font-family:'Archivo'; font-weight:900; color:#0f1216; font-size:15px; letter-spacing:-.03em;
}
.lt-brand-name { font-family:'Archivo'; font-weight:800; color:#fff; font-size:14px; line-height:1; }
.lt-brand-sub { font-family:'IBM Plex Mono'; font-size:9.5px; color:#FFD400; letter-spacing:.16em; margin-top:5px; }

/* ---------- Encabezado de sección ---------- */
.lt-header { display:flex; gap:14px; align-items:flex-start; margin-bottom:6px; }
.lt-bar { width:4px; align-self:stretch; min-height:38px; border-radius:4px; background:#FFD400; }
.lt-title { font-size:28px !important; font-weight:800 !important; margin:0 !important; }
.lt-sub { color:#8b95a3; font-size:14px; margin:6px 0 0; }

/* ---------- Botones ---------- */
.stButton > button, .stDownloadButton > button {
    font-family:'Archivo', sans-serif; font-weight:800; font-size:13.5px;
    border-radius:10px; padding:0.55rem 1.3rem; border:1px solid #2a3038;
    background:#1c2026; color:#eceef1; transition:filter .15s, border-color .15s;
}
.stButton > button:hover, .stDownloadButton > button:hover { filter:brightness(1.1); border-color:#3a4048; }

/* Botón primario (type="primary") */
.stButton > button[kind="primary"],
button[data-testid="baseButton-primary"],
button[data-testid="stBaseButton-primary"] {
    background:#FFD400 !important; color:#14100a !important; border:none !important;
}
.stButton > button[kind="primary"]:hover { filter:brightness(1.06); }

/* Botón de descarga: acento amarillo sobre oscuro */
.stDownloadButton > button { color:#FFD400; }

/* ---------- File uploader ---------- */
[data-testid="stFileUploaderDropzone"] {
    background:#14171c; border:1.5px dashed #2a3038; border-radius:14px;
}
[data-testid="stFileUploaderDropzone"]:hover { border-color:#3a4048; background:#141922; }
[data-testid="stFileUploaderDropzone"] button {
    background:#FFD400 !important; color:#14100a !important; border:none !important; font-weight:700;
}

/* ---------- Métricas (st.metric) como cards ---------- */
[data-testid="stMetric"] {
    background:#14171c; border:1px solid #23272e; border-radius:13px;
    padding:16px 18px;
}
[data-testid="stMetricLabel"] p {
    font-family:'IBM Plex Mono', monospace; font-size:10px !important;
    letter-spacing:.08em; text-transform:uppercase; color:#7f8896 !important;
}
[data-testid="stMetricValue"] {
    font-family:'Archivo', sans-serif; font-weight:800; color:#ffffff; font-size:26px;
}

/* ---------- Inputs de texto / búsqueda ---------- */
[data-testid="stTextInput"] input, [data-baseweb="input"] input, .stTextInput input {
    background:#14171c !important; border:1px solid #23272e !important; border-radius:11px !important;
    color:#ffffff !important;
}
[data-testid="stTextInput"] input:focus { border-color:#FFD400 !important; box-shadow:none !important; }

/* ---------- Selectbox ---------- */
[data-baseweb="select"] > div {
    background:#14171c !important; border:1px solid #23272e !important; border-radius:10px !important;
}

/* ---------- Radio horizontal como segmentado ---------- */
[data-testid="stRadio"] [role="radiogroup"] { gap:6px; }
[data-testid="stRadio"] label {
    background:#14171c; border:1px solid #23272e; border-radius:9px;
    padding:8px 15px; margin:0; transition:border-color .15s;
}
[data-testid="stRadio"] label:hover { border-color:#3a4048; }

/* ---------- Callouts (warning / error / info) ---------- */
[data-testid="stAlert"] { border-radius:12px; border:1px solid #23272e; }
div[data-baseweb="notification"] { border-radius:12px; }
/* warning (amarillo/ámbar) */
[data-testid="stAlert"][kind="warning"], .stAlert [data-testid="stMarkdownContainer"] {}

/* ---------- Expander ---------- */
[data-testid="stExpander"] {
    background:#14171c; border:1px solid #23272e; border-radius:12px; overflow:hidden;
}
[data-testid="stExpander"] summary { font-weight:600; color:#eceef1; }

/* ---------- Dataframe / tabla ---------- */
[data-testid="stDataFrame"], [data-testid="stTable"] {
    border:1px solid #23272e; border-radius:14px; overflow:hidden;
}
[data-testid="stDataFrame"] { background:#14171c; }

/* ---------- Divisores ---------- */
hr { border-color:#1c2026 !important; }

/* ---------- Caption ---------- */
[data-testid="stCaptionContainer"], .stCaption { color:#8b95a3 !important; }

/* Oculta el "Made with Streamlit" para look más limpio (opcional) */
footer { visibility:hidden; }
</style>
"""
