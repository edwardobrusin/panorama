import streamlit as st
from datetime import datetime
from pathlib import Path
import pandas as pd
import base64
import os

# ==========================================================================
# FUNCIONES AUXILIARES INTEGRADAS (ANTES EN ENGINE/)
# ==========================================================================

# Ajuste de rutas: al estar en la raíz (app.py), usamos .parent en lugar de .parent.parent
DATA_DIR = Path(__file__).parent / "data" / "intermediate"
MANUAL_DIR = Path(__file__).parent / "data" / "manual"

# --- De engine/formatting.py ---
MESES_ABR = {
    1: "Ene", 2: "Feb", 3: "Mar", 4: "Abr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Ago", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dic",
}

def fmt_periodo_trimestral(periodo: str) -> str:
    """'2026/02' -> '2T 2026'"""
    try:
        y, t = periodo.split("/")
        return f"{int(t)}T {y}"
    except Exception:
        return periodo

def fmt_periodo_mensual(periodo: str) -> str:
    """'2026/06' -> 'Jun 2026'"""
    try:
        y, m = periodo.split("/")
        return f"{MESES_ABR.get(int(m), m)} {y}"
    except Exception:
        return periodo

def fmt_fecha_corta(fecha) -> str:
    """Timestamp/str fecha -> 'DD-Mon-AAAA'"""
    try:
        return f"{fecha.day:02d}-{MESES_ABR.get(fecha.month, fecha.month)}-{fecha.year}"
    except Exception:
        return str(fecha)

def fmt_pct(valor, decimales=1, con_signo=True) -> str:
    """1.234 -> '+1.2%' | -1.234 -> '-1.2%'"""
    if valor is None:
        return "N/D"
    try:
        signo = "+" if (con_signo and valor > 0) else ""
        return f"{signo}{valor:.{decimales}f}%"
    except Exception:
        return "N/D"

def fmt_numero(valor, decimales=1) -> str:
    if valor is None:
        return "N/D"
    try:
        return f"{valor:,.{decimales}f}"
    except Exception:
        return "N/D"

def delta_css_class(valor) -> str:
    """Regresa la clase CSS (positivo/negativo) según el signo del valor."""
    if valor is None:
        return "metric-sub"
    try:
        return "metric-delta-pos" if valor >= 0 else "metric-delta-neg"
    except Exception:
        return "metric-sub"


# --- De engine/loaders.py ---
def _safe_read_excel(path: Path, **kwargs):
    if not path.exists():
        return None
    try:
        return pd.read_excel(path, **kwargs)
    except Exception:
        return None

def load_generico(filename: str) -> pd.DataFrame | None:
    """Lee un xlsx con estructura Indicador/Clave_Indicador/Periodo/Valor."""
    df = _safe_read_excel(DATA_DIR / filename)
    if df is None:
        return None
    df["Periodo"] = df["Periodo"].astype(str)
    return df

def ultimo_valor_generico(df: pd.DataFrame | None, indicador: str):
    """
    Regresa (periodo, valor) del último periodo disponible para un `Indicador`
    específico dentro de un DataFrame genérico. None, None si no hay datos.
    """
    if df is None:
        return None, None
    sub = df[df["Indicador"] == indicador].copy()
    if sub.empty:
        return None, None
    sub = sub.sort_values("Periodo")
    ultimo = sub.iloc[-1]
    return ultimo["Periodo"], ultimo["Valor"]

def penultimo_valor_generico(df: pd.DataFrame | None, indicador: str):
    """Regresa (periodo, valor) del penúltimo dato disponible (para variaciones vs. mes/trimestre previo)."""
    if df is None:
        return None, None
    sub = df[df["Indicador"] == indicador].copy()
    if len(sub) < 2:
        return None, None
    sub = sub.sort_values("Periodo")
    penultimo = sub.iloc[-2]
    return penultimo["Periodo"], penultimo["Valor"]

MESES_ES = {
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6,
    "julio": 7, "agosto": 8, "septiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12,
}

def load_puestos_imss(filename: str = "puestos_imss.xlsx") -> pd.DataFrame | None:
    """
    Lee puestos_imss.xlsx (Año, Mes texto, Total = nivel) y calcula
    Var. Anual y Var. Mensual a partir del nivel (no vienen precalculadas).
    """
    df = _safe_read_excel(DATA_DIR / filename)
    if df is None:
        return None
    df = df.copy()
    df["MesNum"] = df["Mes"].astype(str).str.strip().str.lower().map(MESES_ES)
    df = df.dropna(subset=["MesNum"])
    df["Periodo"] = df["Año"].astype(int).astype(str) + "/" + df["MesNum"].astype(int).map(lambda m: f"{m:02d}")
    df = df.sort_values("Periodo").reset_index(drop=True)
    df["Var_Anual"] = df["Total"].pct_change(periods=12) * 100
    df["Var_Mensual"] = df["Total"].pct_change(periods=1) * 100
    return df[["Periodo", "Total", "Var_Anual", "Var_Mensual"]]

def load_ied(filename: str = "ied.xlsx") -> pd.DataFrame | None:
    """
    Lee ied.xlsx (Año, Trimestre, Total = flujo acumulado del año a ese
    trimestre). La Var. Anual compara el mismo trimestre (mismo acumulado)
    del año previo -> pct_change(periods=4) una vez ordenado cronológicamente.
    """
    df = _safe_read_excel(DATA_DIR / filename)
    if df is None:
        return None
    df = df.copy()
    df["Trimestre"] = df["Trimestre"].astype(int)
    df["Periodo"] = df["Año"].astype(int).astype(str) + "/" + df["Trimestre"].map(lambda t: f"{t:02d}")
    df = df.sort_values("Periodo").reset_index(drop=True)
    df["Var_Anual"] = df["Total"].pct_change(periods=4) * 100
    return df[["Periodo", "Total", "Var_Anual"]]

def ultimo_valor_calculado(df: pd.DataFrame | None, col: str):
    """Regresa (periodo, valor_de_col) del último renglón de un df ya calculado (puestos_imss / ied)."""
    if df is None or df.empty:
        return None, None
    ultimo = df.iloc[-1]
    valor = ultimo[col]
    if pd.isna(valor):
        return ultimo["Periodo"], None
    return ultimo["Periodo"], valor

def load_fecha_valor(filename: str, value_col: str) -> pd.DataFrame | None:
    df = _safe_read_excel(DATA_DIR / filename)
    if df is None:
        return None
    df["fecha"] = pd.to_datetime(df["fecha"])
    df = df.sort_values("fecha")
    return df[["fecha", value_col]].rename(columns={value_col: "Valor"})

def ultimo_valor_fecha(df: pd.DataFrame | None):
    """Regresa (fecha, valor) del último dato disponible en una serie fecha/valor."""
    if df is None or df.empty:
        return None, None
    ultimo = df.iloc[-1]
    return ultimo["fecha"], ultimo["Valor"]

def cargar_todas_las_fuentes() -> dict:
    """
    Carga todas las fuentes de la Tabla 1 en un solo diccionario.
    Las claves ausentes de archivo (p. ej. 'ied', 'empleo') quedan como None
    y se muestran como N/D en el dashboard.
    """
    fuentes = {
        "pib": load_generico("pib.xlsx"),
        "inpc_mensual": load_generico("inpc_mensual.xlsx"),
        "inpc_quincenal": load_generico("inpc_quincenal.xlsx"),
        "importaciones": load_generico("importaciones.xlsx"),
        "igae": load_generico("igae.xlsx"),
        "fbkf": load_generico("fbkf.xlsx"),
        "exportaciones": load_generico("exportaciones.xlsx"),
        "desocupacion": load_generico("desocupacion.xlsx"),
        "consumo_privado": load_generico("consumo_privado.xlsx"),
        "actividad_industrial": load_generico("actividad_industrial.xlsx"),
        "remesas": load_fecha_valor("remesas.xlsx", "Ingresos"),
        "fix": load_fecha_valor("fix.xlsx", "FIX"),
        "tasa_objetivo": load_fecha_valor("tasa_objetivo.xlsx", "Tasa de Referencia"),
        "ied": load_ied("ied.xlsx"),
        "empleo": load_puestos_imss("puestos_imss.xlsx"),
    }
    return fuentes

def cargar_cuadros_manuales() -> dict:
    path = MANUAL_DIR / "cuadros.xlsx"
    if not path.exists():
        return {"expectativas": None, "finanzas": None, "pemex": None}
    
    # 1. Expectativas: omitimos título (fila 0). Tomamos 4 columnas y las renombramos exactamente como espera el script.
    try:
        df_exp = pd.read_excel(path, sheet_name="expectativas", skiprows=1).iloc[:, :4]
        df_exp.columns = ["Institucion", "Detalle", "2026", "2027"]
    except Exception:
        df_exp = None
    
    # 2. Finanzas: omitimos título, años y subencabezados (3 filas). Dejamos header=None para tomar solo los datos (los headers se arman manual más abajo).
    try:
        df_fin = pd.read_excel(path, sheet_name="finanzas", skiprows=3, header=None)
    except Exception:
        df_fin = None
        
    # 3. Pemex: omitimos el título (fila 0) y tomamos automáticamente los nombres de columna de la fila 1.
    try:
        df_pemex = pd.read_excel(path, sheet_name="pemex", skiprows=1)
    except Exception:
        df_pemex = None

    return {
        "expectativas": df_exp,
        "finanzas": df_fin,
        "pemex": df_pemex,
    }


# --- De engine/mensajes.py ---
VARIABLES_VALIDAS = {
    "pib", "igae", "act_ind", "cons_priv", "fbkf", "ied", "exportaciones",
    "importaciones", "remesas", "inpc", "tdc", "tasa_obj", "empleo", "desocupacion",
}
PERIODICIDADES_VALIDAS = {"anual", "trimestral", "mensual", "diario", "no fijo", "nivel"}

def cargar_mensajes() -> pd.DataFrame | None:
    path = MANUAL_DIR / "mensajes.xlsx"
    if not path.exists():
        return None
    df = pd.read_excel(path)
    df["variable"] = df["variable"].astype(str).str.strip().str.lower()
    df["periodicidad"] = df["periodicidad"].astype(str).str.strip().str.lower()
    return df

def obtener_bullets(df_mensajes: pd.DataFrame | None, variable: str, periodicidad: str) -> list[str]:
    """Regresa la lista de bullets (en orden) para una variable+periodicidad dadas."""
    if df_mensajes is None:
        return []
    sub = df_mensajes[
        (df_mensajes["variable"] == variable) & (df_mensajes["periodicidad"] == periodicidad)
    ]
    return sub["mensaje"].dropna().tolist()


# --- De engine/cards.py ---
def _bloque_html(bloque: dict) -> str:
    """
    bloque = {
        "label": "Anual",              # etiqueta de periodicidad
        "valor": 0.78,                 # valor numérico o None
        "tipo": "pct" | "nivel" | "indice" | "texto",
        "periodo": "2T 2026",          # periodo formateado (opcional)
        "bullets": ["...", "..."],     # mensajes relevantes (opcional)
    }
    """
    label = bloque.get("label", "")
    valor = bloque.get("valor")
    tipo = bloque.get("tipo", "pct")
    periodo = bloque.get("periodo")
    bullets = bloque.get("bullets") or []

    if valor is None:
        valor_str = "N/D"
        css_class = "metric-sub"
        icono = ""
    elif tipo == "pct":
        valor_str = fmt_pct(valor)
        css_class = delta_css_class(valor)
        icono = "▲" if valor >= 0 else "▼"
    elif tipo == "nivel":
        valor_str = fmt_numero(valor, 2)
        css_class = "metric-value-inline"
        icono = ""
    else:
        valor_str = str(valor)
        css_class = "metric-value-inline"
        icono = ""

    periodo_html = f'<span class="metric-sub" style="margin-left:6px; color:#94A3B8;">({periodo})</span>' if periodo else ""

    bullets_html = ""
    if bullets:
        items = "".join(f"<li>{b}</li>" for b in bullets)
        bullets_html = f'<ul class="bullet-list">{items}</ul>'

    return f"""<div class="bloque-periodicidad">
<div style="display:flex; justify-content:space-between; align-items:baseline;">
<span class="bloque-label">{label}</span>
<span class="{css_class}" style="font-size:1.15rem; font-weight:800;">{icono} {valor_str}</span>
</div>
{periodo_html}
{bullets_html}
</div>"""

def card_variable_html(titulo: str, bloques: list, subtitulo: str = "") -> str:
    """Tarjeta completa de una variable de la Tabla 1 (puede tener 1+ bloques de periodicidad)."""
    bloques_html = "".join(_bloque_html(b) for b in bloques)
    subtitulo_html = f'<div class="metric-sub" style="color:#94A3B8; margin-top:-4px; margin-bottom:8px;">{subtitulo}</div>' if subtitulo else ""

    return f"""<div class="metric-container card-hover">
<div class="metric-title">{titulo}</div>
{subtitulo_html}
{bloques_html}
</div>"""

def card_dato_simple_html(titulo: str, valor: str, fecha: str = "", mensaje: str = "") -> str:
    """Tarjeta simple para variables diarias (FIX, Tasa de Referencia): un valor + mensaje, sin bloques."""
    fecha_html = f'<div class="metric-sub" style="color:#94A3B8;">Dato al {fecha}</div>' if fecha else ""
    mensaje_html = f'<div class="metric-sub" style="margin-top:10px; padding:8px; background-color:#F8FAFC; border-radius:6px;">{mensaje}</div>' if mensaje else ""
    
    return f"""<div class="metric-container card-hover">
<div class="metric-title">{titulo}</div>
<div class="metric-value">{valor}</div>
{fecha_html}
{mensaje_html}
</div>"""

def tabla_html(headers: list, rows: list, highlight_cols: list | None = None, title: str = "") -> str:
    """
    Construye una tabla HTML estilizada genérica.
    highlight_cols: índices de columnas (0-based sobre `headers`) que se resaltan en negritas/color.
    """
    highlight_cols = highlight_cols or []
    title_html = f'<div class="tabla-titulo">{title}</div>' if title else ""

    thead = "".join(f"<th>{h}</th>" for h in headers)
    tbody_rows = []
    for row in rows:
        cells = []
        for i, val in enumerate(row):
            val_disp = "—" if (val is None or (isinstance(val, float) and val != val)) else val
            cls = "cell-highlight" if i in highlight_cols else ""
            cells.append(f'<td class="{cls}">{val_disp}</td>')
        tbody_rows.append(f"<tr>{''.join(cells)}</tr>")

    return f"""<div class="metric-container card-hover" style="overflow-x:auto;">
{title_html}
<table class="tabla-manual">
<thead><tr>{thead}</tr></thead>
<tbody>{''.join(tbody_rows)}</tbody>
</table>
</div>"""

# ==========================================================================
# 1. CONFIGURACIÓN DE LA PÁGINA (BRANDING NAFIN/BANCOMEXT)
# ==========================================================================
st.set_page_config(
    page_title="Panorama Económico | NAFIN - BANCOMEXT",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Noto+Color+Emoji&display=swap');

    .stApp { background-color: #F8FAFC; }

    /* Streamlit reserva por defecto un padding-top considerable en el contenedor
       principal (espacio para la toolbar superior). En la APP EN VIVO sí hace
       falta este aire para que el banner nativo de Streamlit no tape el contenido
       (ver generar_pdf.py para la variante agresiva usada solo en la exportación). */
    [data-testid="stAppViewBlockContainer"], .block-container {
        padding-top: 3rem !important;
    }

    .metric-container {
        background-color: #ffffff;
        border: 1px solid #E2E8F0;
        padding: 22px 24px;
        border-radius: 12px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.03);
        margin-bottom: 18px;
        height: 100%;
        transition: transform 0.2s ease;
    }
    .card-hover:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 25px rgba(0,0,0,0.06);
    }

    .metric-title {
        color: #64748B;
        font-size: 0.85rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        margin-bottom: 8px;
    }
    .metric-value {
        color: #0F172A;
        font-size: 2rem;
        font-weight: 800;
        margin: 5px 0;
        letter-spacing: -0.5px;
    }
    .metric-sub { color: #64748B; font-size: 0.85rem; }
    .metric-delta-pos { color: #059669; font-weight: 800; }
    .metric-delta-neg { color: #DC2626; font-weight: 800; }
    .metric-value-inline { color: #0F172A; font-weight: 800; }

    hr { margin: 15px 0; border-top: 1px solid #E2E8F0; }
    h1, h2, h3 { color: #0F172A !important; font-weight: 800 !important; letter-spacing: -0.5px; }

    /* --- Bloques de periodicidad dentro de una tarjeta de variable --- */
    .bloque-periodicidad {
        border-top: 1px dashed #E2E8F0;
        padding-top: 10px;
        margin-top: 10px;
    }
    .bloque-periodicidad:first-of-type { border-top: none; margin-top: 0; padding-top: 0; }
    .bloque-label {
        color: #334155;
        font-size: 0.8rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 0.4px;
    }
    .bullet-list {
        margin: 6px 0 0 0;
        padding-left: 18px;
        color: #475569;
        font-size: 0.82rem;
        line-height: 1.45;
    }
    .bullet-list li { margin-bottom: 3px; }

    /* --- Encabezados de sección --- */
    .section-header {
        display: flex;
        align-items: center;
        gap: 10px;
        margin-top: 8px;
        margin-bottom: 14px;
        padding-bottom: 8px;
        border-bottom: 3px solid #2596be;
    }
    .section-header h2 { margin: 0 !important; font-size: 1.4rem !important; }
    .section-icon { font-size: 1.6rem; }

    /* --- Header principal --- */
    .main-header {
        background: linear-gradient(135deg, #0F172A 0%, #1E3A5F 60%, #2596be 100%);
        border-radius: 16px;
        padding: 7px 32px;       /* banner comprimido a la mitad con márgenes simétricos */
        margin-bottom: 14px;     /* antes 28px: menos aire debajo */
        color: white;
    }
    .main-header h1 { color: white !important; margin-bottom: 4px !important; }
    .main-header .subtitle { color: #CBD5E1; font-size: 1rem; }
    .main-header .fecha-chip {
        display: inline-block;
        background-color: rgba(255,255,255,0.15);
        padding: 6px 14px;
        border-radius: 20px;
        font-size: 0.85rem;
        font-weight: 600;
        margin-top: 10px;
    }

    /* --- Tablas manuales estilizadas --- */
    .tabla-titulo {
        font-size: 1rem;
        font-weight: 800;
        color: #0F172A;
        margin-bottom: 14px;
    }
    table.tabla-manual {
        width: 100%;
        border-collapse: collapse;
        font-size: 0.85rem;
        table-layout: fixed;
    }
    table.tabla-manual th {
        background-color: #0F172A;
        color: white;
        text-align: center;
        padding: 8px 10px;
        font-weight: 700;
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.3px;
    }
    table.tabla-manual td {
        padding: 7px 10px;
        text-align: center;
        border-bottom: 1px solid #E2E8F0;
        color: #334155;
    }
    table.tabla-manual tr:nth-child(even) { background-color: #F8FAFC; }
    table.tabla-manual td:first-child {
        text-align: left;
        font-weight: 600;
        color: #0F172A;
    }
    table.tabla-manual th:first-child {
        text-align: left;
    }
    table.tabla-manual .cell-highlight {
        font-weight: 800;
        color: #0F172A;
    }

    /* --- Bloques indivisibles para paginación en PDF (motor de impresión de Chromium) --- */
    .cat-bloque {
        break-inside: avoid;
        page-break-inside: avoid; /* fallback por compatibilidad */
    }
    /* Evita un doble borde/gap visual entre tablas de categoría consecutivas */
    .cat-bloque table.tabla-manual {
        border-top: none;
    }

    /* --- Reglas generales de paginación para el resto del dashboard --- */

    /* Tarjetas/contenedores de tablas manuales (expectativas, finanzas, pemex):
       cada una se mantiene entera si cabe en el espacio restante de la hoja.
       Si el bloque es más alto que una hoja completa, esta regla no puede aplicar
       (es una limitación física, no de la implementación) y el navegador cae de
       forma segura a partir la tabla entre filas — para eso dejamos el <thead>
       repetible más abajo, así nunca se pierde el contexto de columnas. */
    .metric-container {
        break-inside: avoid;
        page-break-inside: avoid;
    }

    /* Encabezados de sección: nunca deben quedar solos al final de una hoja
       sin su contenido debajo (huérfanos) */
    .section-header {
        break-after: avoid;
        page-break-after: avoid;
        break-inside: avoid;
    }

    /* El encabezado principal del reporte tampoco debe partirse */
    .main-header {
        break-inside: avoid;
        page-break-inside: avoid;
    }

    /* Si una tabla SÍ termina partiéndose entre hojas (tablas largas que no caben
       enteras), el encabezado de columnas se repite en cada hoja nueva */
    table.tabla-manual thead {
        display: table-header-group;
    }

    /* Evita romper una fila de tabla a la mitad entre dos hojas */
    table.tabla-manual tr {
        break-inside: avoid;
        page-break-inside: avoid;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================================================
# 2. CARGA DE DATOS
# ==========================================================================
@st.cache_data
def cargar_datos():
    fuentes = cargar_todas_las_fuentes()
    mensajes = cargar_mensajes()
    cuadros = cargar_cuadros_manuales()
    return fuentes, mensajes, cuadros

FUENTES, MENSAJES, CUADROS = cargar_datos()


def bloque(label, valor, tipo, periodo_fmt, variable, periodicidad):
    """Arma un bloque estándar {label, valor, tipo, periodo, bullets} para una tarjeta."""
    return {
        "label": label,
        "valor": valor,
        "tipo": tipo,
        "periodo": periodo_fmt,
        "bullets": obtener_bullets(MENSAJES, variable, periodicidad),
    }


def section_header(icon, title):
    st.markdown(f'<div class="section-header"><span class="section-icon">{icon}</span><h2>{title}</h2></div>', unsafe_allow_html=True)


# ==========================================================================
# BOTÓN FLOTANTE DE DESCARGA
# ==========================================================================
ruta_pdf = "panorama_economico.pdf"

if os.path.exists(ruta_pdf):
    with open(ruta_pdf, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    
    svg_icon = '''
    <svg xmlns="http://www.w3.org/2000/svg" fill="none" viewBox="0 0 24 24" stroke-width="2" stroke="currentColor" style="width: 24px; height: 24px;">
      <path stroke-linecap="round" stroke-linejoin="round" d="M3 16.5v2.25A2.25 2.25 0 0 0 5.25 21h13.5A2.25 2.25 0 0 0 21 18.75V16.5M16.5 12 12 16.5m0 0L7.5 12m4.5 4.5V3" />
    </svg>
    '''
    html_boton = f"""
    <style>
        .floating-download-btn {{
            position: fixed; bottom: 50px; right: 20px; background-color: #2596be; color: white !important;
            border-radius: 50%; width: 56px; height: 56px; display: flex; justify-content: center; align-items: center;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15); z-index: 9999; transition: transform 0.2s ease, box-shadow 0.2s ease, background-color 0.2s ease;
        }}
        .floating-download-btn:hover {{ background-color: #1e7a9b; transform: translateY(-3px); box-shadow: 0 6px 16px rgba(0,0,0,0.2); }}
    </style>
    <a href="data:application/pdf;base64,{base64_pdf}" download="Panorama_Economico.pdf" class="floating-download-btn" title="Descargar PDF">
        {svg_icon}
    </a>
    """
    st.markdown(html_boton, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        .floating-warning { position: fixed; bottom: 50px; right: 20px; background-color: #64748B; color: white; padding: 10px 20px; border-radius: 20px; font-size: 0.8rem; box-shadow: 0 4px 12px rgba(0,0,0,0.1); z-index: 9999; }
    </style>
    <div class="floating-warning">Actualizando PDF...</div>
    """, unsafe_allow_html=True)


# ==========================================================================
# 3. HEADER PRINCIPAL
# ==========================================================================
fecha_revision = "25 de Agosto de 2026"
st.markdown(f"""
<div class="main-header" style="display: flex; justify-content: space-between; align-items: center;">
    <div>
        <h1 style="margin: 0; font-size: 1.6rem;">Panorama Económico</h1>
        <div class="subtitle" style="font-size: 1rem; margin-top: 2px;">Principales Variables</div>
    </div>
    <div style="text-align: right;">
        <div class="fecha-chip" style="margin: 0; font-weight: 500;">Fecha de revisión: {fecha_revision}</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ==========================================================================
# 4. TABLA PRINCIPAL DE VARIABLES MACROECONÓMICAS
# ==========================================================================

# 1. Extracción de todos los datos necesarios
p_a_pib, v_a_pib = ultimo_valor_generico(FUENTES["pib"], "Variación Anual")
p_t_pib, v_t_pib = ultimo_valor_generico(FUENTES["pib"], "Variación Trimestral")
per_pib = fmt_periodo_trimestral(p_a_pib) if p_a_pib else "N/D"

p_a_igae, v_a_igae = ultimo_valor_generico(FUENTES["igae"], "Var. Anual")
p_m_igae, v_m_igae = ultimo_valor_generico(FUENTES["igae"], "Var. Mensual")
per_igae = fmt_periodo_mensual(p_a_igae) if p_a_igae else "N/D"

p_a_act, v_a_act = ultimo_valor_generico(FUENTES["actividad_industrial"], "Var. Anual")
p_m_act, v_m_act = ultimo_valor_generico(FUENTES["actividad_industrial"], "Var. Mensual")
per_act = fmt_periodo_mensual(p_a_act) if p_a_act else "N/D"

p_a_cp, v_a_cp = ultimo_valor_generico(FUENTES["consumo_privado"], "Var. Anual")
p_m_cp, v_m_cp = ultimo_valor_generico(FUENTES["consumo_privado"], "Var. Mensual")
per_cp = fmt_periodo_mensual(p_a_cp) if p_a_cp else "N/D"

p_a_fbkf, v_a_fbkf = ultimo_valor_generico(FUENTES["fbkf"], "Var. Anual")
p_m_fbkf, v_m_fbkf = ultimo_valor_generico(FUENTES["fbkf"], "Var. Mensual")
per_fbkf = fmt_periodo_mensual(p_a_fbkf) if p_a_fbkf else "N/D"

p_a_ied, v_a_ied = ultimo_valor_calculado(FUENTES["ied"], "Var_Anual")
p_niv_ied, v_niv_ied = ultimo_valor_calculado(FUENTES["ied"], "Total")
per_ied = fmt_periodo_trimestral(p_niv_ied) if p_niv_ied else "N/D"

p_a_exp, v_a_exp = ultimo_valor_generico(FUENTES["exportaciones"], "Var. Anual")
p_m_exp, v_m_exp = ultimo_valor_generico(FUENTES["exportaciones"], "Var. Mensual")
per_exp = fmt_periodo_mensual(p_a_exp) if p_a_exp else "N/D"

p_a_imp, v_a_imp = ultimo_valor_generico(FUENTES["importaciones"], "Var. Anual")
p_m_imp, v_m_imp = ultimo_valor_generico(FUENTES["importaciones"], "Var. Mensual")
per_imp = fmt_periodo_mensual(p_a_imp) if p_a_imp else "N/D"

fecha_rem, val_rem = ultimo_valor_fecha(FUENTES["remesas"])
per_rem = fmt_periodo_mensual(fecha_rem.strftime("%Y/%m")) if fecha_rem is not None else "N/D"

p_a_inpc_m, v_a_inpc_m = ultimo_valor_generico(FUENTES["inpc_mensual"], "Var. Anual")
p_m_inpc_m, v_m_inpc_m = ultimo_valor_generico(FUENTES["inpc_mensual"], "Var. Mensual")
per_inpc_m = fmt_periodo_mensual(p_a_inpc_m) if p_a_inpc_m else "N/D"

p_a_inpc_q, v_a_inpc_q = ultimo_valor_generico(FUENTES["inpc_quincenal"], "Var. Anual")
p_q_inpc_q, v_q_inpc_q = ultimo_valor_generico(FUENTES["inpc_quincenal"], "Var. Quincenal")

mostrar_quincenal = False
per_inpc_q_str = "N/D"

if p_a_inpc_q:
    try:
        partes_q = p_a_inpc_q.split("/")
        y_q, m_q = partes_q[0], partes_q[1]
        
        if len(partes_q) >= 3:
            d_q = partes_q[2]
            meses_nombres = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
                             7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"}
            q = "1Q" if int(d_q) <= 15 else "2Q"
            per_inpc_q_str = f"{q} {meses_nombres.get(int(m_q), m_q)}"
        else:
            per_inpc_q_str = p_a_inpc_q
        
        # Evaluar si la inflación quincenal está adelantada (es de un mes posterior a la mensual)
        if p_a_inpc_m:
            y_m, m_m = p_a_inpc_m.split("/")[:2]
            if (int(y_q), int(m_q)) > (int(y_m), int(m_m)):
                mostrar_quincenal = True
        else:
            mostrar_quincenal = True
    except Exception:
        per_inpc_q_str = p_a_inpc_q

precios_lista = []
if mostrar_quincenal:
    # Por si la base de mensajes sigue diciendo "mensual", usamos fallback
    bullets_quincenal = obtener_bullets(MENSAJES, "inpc", "quincenal") or obtener_bullets(MENSAJES, "inpc", "mensual")
    precios_lista.append(
        ("INPC Quincenal", per_inpc_q_str, [
            ("Anual:", v_a_inpc_q, "pct", obtener_bullets(MENSAJES, "inpc", "anual")),
            ("Quincenal:", v_q_inpc_q, "pct", bullets_quincenal)
        ])
    )
else:
    precios_lista.append(
        ("INPC Mensual", per_inpc_m, [
            ("Anual:", v_a_inpc_m, "pct", obtener_bullets(MENSAJES, "inpc", "anual")),
            ("Mensual:", v_m_inpc_m, "pct", obtener_bullets(MENSAJES, "inpc", "mensual"))
        ])
    )

fecha_fix, val_fix = ultimo_valor_fecha(FUENTES["fix"])
per_fix = fmt_fecha_corta(fecha_fix) if fecha_fix is not None else "N/D"

fecha_tasa, val_tasa = ultimo_valor_fecha(FUENTES["tasa_objetivo"])
per_tasa = fmt_fecha_corta(fecha_tasa) if fecha_tasa is not None else "N/D"

p_a_emp, v_a_emp = ultimo_valor_calculado(FUENTES["empleo"], "Var_Anual")
p_m_emp, v_m_emp = ultimo_valor_calculado(FUENTES["empleo"], "Var_Mensual")
per_emp = fmt_periodo_mensual(p_a_emp) if p_a_emp else "N/D"

p_n_des, v_n_des = ultimo_valor_generico(FUENTES["desocupacion"], "Tasa de Desocupación")
per_des = fmt_periodo_mensual(p_n_des) if p_n_des else "N/D"

# 2. Organización jerárquica para la tabla (Categoría, Indicador, Variaciones)
secciones_datos = [
    ("Crecimiento", [
        ("PIB", per_pib, [
            ("Anual:", v_a_pib, "pct", obtener_bullets(MENSAJES, "pib", "anual")),
            ("Trimestral:", v_t_pib, "pct", obtener_bullets(MENSAJES, "pib", "trimestral"))
        ]),
        ("IGAE", per_igae, [
            ("Anual:", v_a_igae, "pct", obtener_bullets(MENSAJES, "igae", "anual")),
            ("Mensual:", v_m_igae, "pct", obtener_bullets(MENSAJES, "igae", "mensual"))
        ])
    ]),
    ("Producción", [
        ("Actividad Industrial", per_act, [
            ("Anual:", v_a_act, "pct", obtener_bullets(MENSAJES, "act_ind", "anual")),
            ("Mensual:", v_m_act, "pct", obtener_bullets(MENSAJES, "act_ind", "mensual"))
        ])
    ]),
    ("Consumo", [
        ("Consumo Privado", per_cp, [
            ("Anual:", v_a_cp, "pct", obtener_bullets(MENSAJES, "cons_priv", "anual")),
            ("Mensual:", v_m_cp, "pct", obtener_bullets(MENSAJES, "cons_priv", "mensual"))
        ])
    ]),
    ("Inversión", [
        ("Inversión Fija Bruta", per_fbkf, [
            ("Anual:", v_a_fbkf, "pct", obtener_bullets(MENSAJES, "fbkf", "anual")),
            ("Mensual:", v_m_fbkf, "pct", obtener_bullets(MENSAJES, "fbkf", "mensual"))
        ]),
        ("IED (Cifras Originales)", per_ied, [
            ("Anual:", v_a_ied, "pct", obtener_bullets(MENSAJES, "ied", "anual"))
        ])
    ]),
    ("Sector Externo", [
        ("Exportaciones", per_exp, [
            ("Anual:", v_a_exp, "pct", obtener_bullets(MENSAJES, "exportaciones", "anual")),
            ("Mensual:", v_m_exp, "pct", obtener_bullets(MENSAJES, "exportaciones", "mensual"))
        ]),
        ("Importaciones", per_imp, [
            ("Anual:", v_a_imp, "pct", obtener_bullets(MENSAJES, "importaciones", "anual")),
            ("Mensual:", v_m_imp, "pct", obtener_bullets(MENSAJES, "importaciones", "mensual"))
        ]),
        ("Remesas", per_rem, [
            ("Ingresos (mdd):", val_rem, "nivel", obtener_bullets(MENSAJES, "remesas", "mensual"))
        ])
    ]),
    ("Precios", precios_lista),
    ("Mercados", [
        ("Tipo de Cambio FIX", per_fix, [
            ("Nivel (MXN/USD):", val_fix, "nivel_4", obtener_bullets(MENSAJES, "tdc", "diario"))
        ]),
        ("Tasa de Referencia", per_tasa, [
            ("Nivel:", val_tasa, "tasa", obtener_bullets(MENSAJES, "tasa_obj", "no fijo"))
        ])
    ]),
    ("Empleo", [
        ("Empleo IMSS", per_emp, [
            ("Anual:", v_a_emp, "pct", obtener_bullets(MENSAJES, "empleo", "anual")),
            ("Mensual:", v_m_emp, "pct", obtener_bullets(MENSAJES, "empleo", "mensual"))
        ]),
        ("Desocupación (ENOE)", per_des, [
            ("Nivel:", v_n_des, "tasa", obtener_bullets(MENSAJES, "desocupacion", "nivel"))
        ])
    ])
]

# 3. Generación de bloques HTML por categoría (una <table> independiente por categoría)
#    CLAVE PARA PAGINACIÓN: cada categoría es su propia <table>, envuelta en un
#    <div class="cat-bloque"> con break-inside:avoid. Así, al exportar a PDF, Chromium
#    jamás puede partir un rowspan de categoría/indicador a la mitad entre dos hojas:
#    si la categoría completa no cabe en el espacio restante de la hoja actual,
#    se manda entera a la siguiente.
COLGROUP_TABLA_VARS = """<colgroup>
<col style="width:13%;"><col style="width:15%;"><col style="width:12%;"><col style="width:60%;">
</colgroup>"""

THEAD_TABLA_VARS = """<thead><tr>
    <th colspan="2" style="text-align:left; padding-left:14px;">Variable Económica</th>
    <th style="text-align:left;">Variación<br><span style="font-size:0.75rem; font-weight:normal; text-transform:none; color:#CBD5E1;">(cifras a.e.)</span></th>
    <th style="text-align:left;">Mensaje relevante</th>
</tr></thead>"""

bloques_categoria_html = []
for idx_cat, (cat_name, indicadores) in enumerate(secciones_datos):
    cat_rowspan = sum(len(ind[2]) for ind in indicadores)
    first_cat = True
    filas_cat = []

    for ind_name, ind_per, variaciones in indicadores:
        ind_rowspan = len(variaciones)
        first_ind = True
        for var_label, var_val, var_tipo, var_bullets in variaciones:
            # Formateo de los valores (idéntico al original)
            if var_val is None:
                val_str = "N/D"
                val_css = "color:#94A3B8;"
            else:
                if var_tipo == "pct":
                    val_str = fmt_pct(var_val)
                    val_css = "color:#059669; font-weight:800;" if var_val >= 0 else "color:#DC2626; font-weight:800;"
                elif var_tipo == "nivel":
                    val_str = fmt_numero(var_val, 2)
                    val_css = "color:#0F172A; font-weight:800;"
                elif var_tipo == "nivel_4":
                    val_str = fmt_numero(var_val, 4)
                    val_css = "color:#0F172A; font-weight:800;"
                elif var_tipo == "tasa":
                    val_str = f"{fmt_numero(var_val, 2)}%"
                    val_css = "color:#0F172A; font-weight:800;"

            # Formateo de los bullets (Mensajes)
            if var_bullets:
                bull_li = "".join(f"<li style='margin-bottom:4px;'>{b}</li>" for b in var_bullets)
                bull_html = f"<ul style='margin:0; padding-left:1.2em; text-align:left; color:#334155; font-size:0.85rem;'>{bull_li}</ul>"
            else:
                bull_html = "—"

            # Inserción de filas y mezcla de celdas (rowspan) — ahora acotado a esta categoría
            row_html = "<tr>"
            if first_cat:
                row_html += f"<td rowspan='{cat_rowspan}' style='vertical-align:middle; text-align:left; font-weight:700; background-color:#F8FAFC; color:#0F172A; border-right:1px solid #E2E8F0;'>{cat_name}</td>"
                first_cat = False

            if first_ind:
                row_html += f"<td rowspan='{ind_rowspan}' style='vertical-align:middle; text-align:left; border-right:1px solid #E2E8F0;'><span style='font-weight:700; color:#1E293B;'>{ind_name}</span><br><span style='font-size:0.85rem; color:#64748B;'>{ind_per}</span></td>"
                first_ind = False

            row_html += f"<td style='vertical-align:middle; text-align:left;'><span style='color:#64748B; font-size:0.85rem; display:block; margin-bottom:2px;'>{var_label}</span><span style='{val_css} font-size:1rem;'>{val_str}</span></td>"
            row_html += f"<td style='vertical-align:middle; text-align:left;'>{bull_html}</td>"
            row_html += "</tr>"

            filas_cat.append(row_html)

    # Solo la primera categoría lleva <thead> visible; las demás comparten el mismo
    # colgroup para que las columnas queden perfectamente alineadas entre tablas.
    thead_html = THEAD_TABLA_VARS if idx_cat == 0 else ""
    tabla_cat = f"""<table class="tabla-manual" style="margin:0; width:100%;">
{COLGROUP_TABLA_VARS}
{thead_html}
<tbody>{''.join(filas_cat)}</tbody>
</table>"""

    bloques_categoria_html.append(f'<div class="cat-bloque">{tabla_cat}</div>')

# 4. Renderizado final: un contenedor visual único que por dentro son N tablas
#    independientes (una por categoría), cada una indivisible para la paginación.
html_gran_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto; padding: 0; break-inside: auto; page-break-inside: auto;">
{''.join(bloques_categoria_html)}
</div>"""

st.markdown(html_gran_tabla, unsafe_allow_html=True)

# ==========================================================================
# 11. EXPECTATIVAS DE CRECIMIENTO ECONÓMICO (tabla manual)
# ==========================================================================
section_header("", "Expectativas de Crecimiento Económico")

if CUADROS["expectativas"] is not None:
    import re
    df_exp = CUADROS["expectativas"]
    
    # Extraemos únicamente los valores de 2026 y 2027 (columnas índice 2 y 3)
    rows = df_exp.iloc[:, 2:4].values.tolist()
    
    # Función para limpiar valores y aplicar formato condicional "De X a Y"
    def format_val(v):
        if pd.isna(v) or str(v).strip() == "":
            return "—"
        s = str(v).strip()
        
        # Buscamos el patrón ignorando mayúsculas/minúsculas. Admite % opcionales.
        match = re.match(r"^De\s+([0-9\.]+)(%?)\s+a\s+([0-9\.]+)(%?)$", s, re.IGNORECASE)
        if match:
            n1_str, p1, n2_str, p2 = match.groups()
            try:
                n1, n2 = float(n1_str), float(n2_str)
                if n1 < n2:
                    # Mejoró (Verde)
                    n2_color = f"<span style='color:#059669; font-weight:800;'>{n2_str}{p2}</span>"
                    return f"De {n1_str}{p1} a {n2_color}"
                elif n1 > n2:
                    # Empeoró (Rojo)
                    n2_color = f"<span style='color:#DC2626; font-weight:800;'>{n2_str}{p2}</span>"
                    return f"De {n1_str}{p1} a {n2_color}"
            except ValueError:
                pass
        return s
    
    # Estilo reutilizable para la columna "Detalle"
    td_style = "style='font-weight:400; text-align:center; color:#334155;'"
    
    # Armamos todo el HTML hardcodeado, inyectando format_val(rows[fila][columna])
    # rows[i][0] corresponde a 2026 y rows[i][1] a 2027
    tbody_html = f"""<tr><td rowspan='2' style='vertical-align:middle;'>SHCP<br><span style='font-weight:400; font-size:0.9em; color:#64748B;'>(rango)</span></td><td {td_style}>CGPE 2026 (sep-25)</td><td class='cell-highlight'>{format_val(rows[0][0])}</td><td class='cell-highlight'>{format_val(rows[0][1])}</td></tr>
<tr><td {td_style}>PCGPE 2027 (abr-26)</td><td class='cell-highlight'>{format_val(rows[1][0])}</td><td class='cell-highlight'>{format_val(rows[1][1])}</td></tr>
<tr><td colspan='2'>Banxico rango (1T26, informe trimestral)</td><td class='cell-highlight'>{format_val(rows[2][0])}</td><td class='cell-highlight'>{format_val(rows[2][1])}</td></tr>
<tr><td rowspan='4' style='vertical-align:middle;'>FMI<br><span style='font-weight:400; font-size:0.9em; color:#64748B;'>(ene-26 vs abr-26)</span></td><td {td_style}>México</td><td class='cell-highlight'>{format_val(rows[3][0])}</td><td class='cell-highlight'>{format_val(rows[3][1])}</td></tr>
<tr><td {td_style}>LATAM</td><td class='cell-highlight'>{format_val(rows[4][0])}</td><td class='cell-highlight'>{format_val(rows[4][1])}</td></tr>
<tr><td {td_style}>EE.UU.</td><td class='cell-highlight'>{format_val(rows[5][0])}</td><td class='cell-highlight'>{format_val(rows[5][1])}</td></tr>
<tr><td {td_style}>Mundial</td><td class='cell-highlight'>{format_val(rows[6][0])}</td><td class='cell-highlight'>{format_val(rows[6][1])}</td></tr>
<tr><td colspan='2'>OCDE (Economic Outlook mar-26 vs jun-26)</td><td class='cell-highlight'>{format_val(rows[7][0])}</td><td class='cell-highlight'>{format_val(rows[7][1])}</td></tr>
<tr><td colspan='2'>Encuesta Banxico (01-jun vs 01-jul, 2026)</td><td class='cell-highlight'>{format_val(rows[8][0])}</td><td class='cell-highlight'>{format_val(rows[8][1])}</td></tr>
<tr><td colspan='2'>Encuesta Citi (22-jun vs 07-jul, 2026)</td><td class='cell-highlight'>{format_val(rows[9][0])}</td><td class='cell-highlight'>{format_val(rows[9][1])}</td></tr>"""

    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto; padding: 0;">
<table class="tabla-manual" style="margin: 0; width: 100%;">
<thead><tr><th colspan="2">Institución</th><th>2026</th><th>2027</th></tr></thead>
<tbody>{tbody_html}</tbody>
</table>
</div>"""

    st.markdown(html_tabla, unsafe_allow_html=True)
else:
    st.warning("No se encontró `data/manual/cuadros.xlsx` (hoja `expectativas`).")

# ==========================================================================
# 12. FINANZAS PÚBLICAS 2019–2026 (tabla manual)
# ==========================================================================
section_header("", "Finanzas Públicas 2019 – 2026 (% del PIB)")

if CUADROS["finanzas"] is not None:
    df_fin = CUADROS["finanzas"]
    rows = df_fin.values.tolist()
    
    # Función rápida para limpiar vacíos (pone un guion si no hay datos)
    c = lambda v: "—" if (pd.isna(v) or str(v).strip() == "") else str(v)
    
    tbody_rows = []
    for row in rows:
        # row[0] es el nombre del indicador
        tr = f"<tr><td>{c(row[0])}</td>"
        
        # Años 2019 a 2024 (columnas índice 1 al 6, sin formato especial)
        for i in range(1, 7):
            tr += f"<td>{c(row[i])}</td>"
            
        # Años 2025 y 2026 (columnas índice 7 al 10, con clase highlight)
        for i in range(7, 11):
            tr += f"<td class='cell-highlight'>{c(row[i])}</td>"
            
        tr += "</tr>"
        tbody_rows.append(tr)
        
    tbody_html = "".join(tbody_rows)
    
    # Estilo en línea para los sub-encabezados (un tono azul/grisáceo elegante)
    sub_th = "background-color:#1E293B; font-weight:600; color:#CBD5E1; font-size:0.75rem; border-top:1px solid #334155; text-align:center;"
    
    # Construcción del thead con rowspan y colspan
    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto; padding: 0;">
<table class="tabla-manual" style="margin: 0; width: 100%;">
<thead>
    <tr>
        <th rowspan="2" style="vertical-align:middle;">% del PIB</th>
        <th>2019</th><th>2020</th><th>2021</th><th>2022</th><th>2023</th><th>2024</th>
        <th colspan="2">2025</th>
        <th colspan="2">2026</th>
    </tr>
    <tr>
        <th colspan="6" style="{sub_th}">Cierre</th>
        <th style="{sub_th}">Aprob.</th><th style="{sub_th}">Observado</th>
        <th style="{sub_th}">Aprob.</th><th style="{sub_th}">Estimado</th>
    </tr>
</thead>
<tbody>{tbody_html}</tbody>
</table>
</div>"""

    st.markdown(html_tabla, unsafe_allow_html=True)
    st.caption("Fuente: Informe sobre la situación económica, las finanzas públicas y la deuda "
               "pública, SHCP (4T25 y 1er trimestre 2026).")
else:
    st.warning("No se encontró `data/manual/cuadros.xlsx` (hoja `finanzas`).")

# ==========================================================================
# 13. PEMEX (tabla manual)
# ==========================================================================
section_header("", "Pemex")

if CUADROS["pemex"] is not None:
    df_pemex = CUADROS["pemex"]
    rows = df_pemex.values.tolist()
    
    # Asegurarnos de que hay al menos 4 filas en el df
    while len(rows) < 4:
        rows.append([""] * 11)
        
    # Tomamos exclusivamente los últimos 9 elementos (años 2018 a 2026) ignorando "Unnamed: 1"
    r0, r1, r2, r3 = rows[0][-9:], rows[1][-9:], rows[2][-9:], rows[3][-9:]
    
    # Limpiador y formateador de números con separador de miles
    def c(v):
        if pd.isna(v) or str(v).strip() in ["", "--", "---", "----", "NaN"]:
            return "—"
        if isinstance(v, (int, float)):
            if v == int(v):
                return f"{int(v):,}"
            else:
                return f"{v:,.2f}"
        return str(v).strip()

    td_ind = "padding-left: 20px; color: #475569;"

    tbody_html = f"""<tr><td>Producción de hidrocarburos (mbd)¹</td><td>{c(r0[0])}</td><td>{c(r0[1])}</td><td>{c(r0[2])}</td><td>{c(r0[3])}</td><td>{c(r0[4])}</td><td>{c(r0[5])}</td><td>{c(r0[6])}</td><td>{c(r0[7])}</td><td>{c(r0[8])}</td></tr>
<tr><td>Saldo de la deuda financiera²</td><td>{c(r1[0])}</td><td>{c(r1[1])}</td><td>{c(r1[2])}</td><td>{c(r1[3])}</td><td>{c(r1[4])}</td><td>{c(r1[5])}</td><td>{c(r1[6])}</td><td>{c(r1[7])}</td><td>{c(r1[8])}</td></tr>
<tr><td style="{td_ind}">Deuda financiera</td><td>{c(r2[0])}</td><td>{c(r2[1])}</td><td>{c(r2[2])}</td><td>{c(r2[3])}</td><td>{c(r2[4])}</td><td>{c(r2[5])}</td><td>{c(r2[6])}</td><td>{c(r2[7])}</td><td>{c(r2[8])}</td></tr>
<tr><td style="{td_ind}">Monetización pagarés del Gob. Federal</td><td>{c(r3[0])}</td><td>{c(r3[1])}</td><td>{c(r3[2])}</td><td>{c(r3[3])}</td><td>{c(r3[4])}</td><td>{c(r3[5])}</td><td>{c(r3[6])}</td><td>{c(r3[7])}</td><td>{c(r3[8])}</td></tr>"""

    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto; padding: 0;">
<table class="tabla-manual" style="margin: 0; width: 100%;">
<thead>
    <tr>
        <th>Indicador</th>
        <th>2018</th><th>2019</th><th>2020</th><th>2021</th><th>2022</th><th>2023</th><th>2024</th><th>2025</th><th>2026</th>
    </tr>
</thead>
<tbody>{tbody_html}</tbody>
</table>
</div>"""

    st.markdown(html_tabla, unsafe_allow_html=True)
    
    # Se utilizan dos espacios al final de cada línea en el string para forzar el salto de línea en Markdown
    st.caption("Fuente: Presentación a Inversionistas, Pemex (abril, 2026). ¹ Incluye producción de socios. ² Para los años 2018 a 2025 son cifras de estados financieros auditados. Para 2026 son cifras preliminares al 31 de marzo de 2026.")
else:
    st.warning("No se encontró `data/manual/cuadros.xlsx` (hoja `pemex`).")

# ==========================================================================
# FOOTER
# ==========================================================================
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("Dirección de Estudios Económicos · NAFIN / BANCOMEXT")