import streamlit as st
from datetime import datetime
from pathlib import Path
import pandas as pd

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
        padding: 32px 36px;
        margin-bottom: 28px;
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
# 3. HEADER PRINCIPAL
# ==========================================================================
fecha_revision = datetime.now().strftime("%d de %B de %Y")
st.markdown(f"""
<div class="main-header">
    <h1>📊 Panorama Económico</h1>
    <div class="subtitle">Principales variables macroeconómicas de México — Dirección de Estudios Económicos</div>
    <div class="fecha-chip">🗓️ Fecha de revisión: {fecha_revision}</div>
</div>
""", unsafe_allow_html=True)

# ==========================================================================
# 4. CRECIMIENTO — PIB, IGAE, Actividad Industrial
# ==========================================================================
section_header("📈", "Crecimiento")
col1, col2, col3 = st.columns(3)

with col1:
    p_a, v_a = ultimo_valor_generico(FUENTES["pib"], "Variación Anual")
    p_t, v_t = ultimo_valor_generico(FUENTES["pib"], "Variación Trimestral")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_trimestral(p_a) if p_a else None, "pib", "anual"),
        bloque("Variación Trimestral", v_t, "pct", fmt_periodo_trimestral(p_t) if p_t else None, "pib", "trimestral"),
    ]
    st.markdown(card_variable_html("PIB", bloques), unsafe_allow_html=True)

with col2:
    p_a, v_a = ultimo_valor_generico(FUENTES["igae"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["igae"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "igae", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "igae", "mensual"),
    ]
    st.markdown(card_variable_html("IGAE", bloques, subtitulo="Cifras desestacionalizadas"), unsafe_allow_html=True)

with col3:
    p_a, v_a = ultimo_valor_generico(FUENTES["actividad_industrial"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["actividad_industrial"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "act_ind", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "act_ind", "mensual"),
    ]
    st.markdown(card_variable_html("Actividad Industrial", bloques, subtitulo="Cifras desestacionalizadas"), unsafe_allow_html=True)

# ==========================================================================
# 5. CONSUMO — Consumo Privado
# ==========================================================================
section_header("🛒", "Consumo")
col1, col2, col3 = st.columns(3)

with col1:
    p_a, v_a = ultimo_valor_generico(FUENTES["consumo_privado"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["consumo_privado"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "cons_priv", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "cons_priv", "mensual"),
    ]
    st.markdown(card_variable_html("Consumo Privado", bloques), unsafe_allow_html=True)

# ==========================================================================
# 6. INVERSIÓN — FBKF, IED
# ==========================================================================
section_header("🏗️", "Inversión")
col1, col2, col3 = st.columns(3)

with col1:
    p_a, v_a = ultimo_valor_generico(FUENTES["fbkf"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["fbkf"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "fbkf", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "fbkf", "mensual"),
    ]
    st.markdown(card_variable_html("Inversión Fija Bruta", bloques), unsafe_allow_html=True)

with col2:
    p_a, v_a = ultimo_valor_calculado(FUENTES["ied"], "Var_Anual")
    p_niv, v_niv = ultimo_valor_calculado(FUENTES["ied"], "Total")
    bloques = [
        {"label": "Acumulado en el año (mdd)", "valor": v_niv, "tipo": "nivel",
         "periodo": fmt_periodo_trimestral(p_niv) if p_niv else None, "bullets": []},
        bloque("Variación Anual", v_a, "pct", fmt_periodo_trimestral(p_a) if p_a else None, "ied", "anual"),
    ]
    st.markdown(card_variable_html("IED (Cifras Originales)", bloques, subtitulo="Flujo acumulado al periodo"), unsafe_allow_html=True)

# ==========================================================================
# 7. SECTOR EXTERNO — Exportaciones, Importaciones, Remesas
# ==========================================================================
section_header("🌎", "Sector Externo")
col1, col2, col3 = st.columns(3)

with col1:
    p_a, v_a = ultimo_valor_generico(FUENTES["exportaciones"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["exportaciones"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "exportaciones", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "exportaciones", "mensual"),
    ]
    st.markdown(card_variable_html("Exportaciones", bloques), unsafe_allow_html=True)

with col2:
    p_a, v_a = ultimo_valor_generico(FUENTES["importaciones"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["importaciones"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "importaciones", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "importaciones", "mensual"),
    ]
    st.markdown(card_variable_html("Importaciones", bloques), unsafe_allow_html=True)

with col3:
    fecha_rem, val_rem = ultimo_valor_fecha(FUENTES["remesas"])
    periodo_rem = fmt_periodo_mensual(fecha_rem.strftime("%Y/%m")) if fecha_rem is not None else None
    bullets_rem = obtener_bullets(MENSAJES, "remesas", "mensual")
    bloques = [
        {"label": "Ingresos (mdd)", "valor": val_rem, "tipo": "nivel", "periodo": periodo_rem, "bullets": bullets_rem},
    ]
    st.markdown(card_variable_html("Remesas", bloques), unsafe_allow_html=True)

# ==========================================================================
# 8. PRECIOS — INPC
# ==========================================================================
section_header("💲", "Precios")
col1, col2 = st.columns(2)

with col1:
    p_a, v_a = ultimo_valor_generico(FUENTES["inpc_mensual"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["inpc_mensual"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "inpc", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "inpc", "mensual"),
    ]
    st.markdown(card_variable_html("INPC Mensual", bloques), unsafe_allow_html=True)

with col2:
    p_a, v_a = ultimo_valor_generico(FUENTES["inpc_quincenal"], "Var. Anual")
    p_m, v_m = ultimo_valor_generico(FUENTES["inpc_quincenal"], "Var. Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "inpc", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "inpc", "mensual"),
    ]
    st.markdown(card_variable_html("INPC Quincenal", bloques), unsafe_allow_html=True)

# ==========================================================================
# 9. MERCADOS — FIX, Tasa de Referencia
# ==========================================================================
section_header("💹", "Mercados")
col1, col2 = st.columns(2)

with col1:
    fecha_fix, val_fix = ultimo_valor_fecha(FUENTES["fix"])
    mensaje_fix = " · ".join(obtener_bullets(MENSAJES, "tdc", "diario"))
    st.markdown(card_dato_simple_html(
        "Tipo de Cambio FIX",
        f"{fmt_numero(val_fix, 4)} MXN/USD" if val_fix is not None else "N/D",
        fmt_fecha_corta(fecha_fix) if fecha_fix is not None else "",
        mensaje_fix,
    ), unsafe_allow_html=True)

with col2:
    fecha_tasa, val_tasa = ultimo_valor_fecha(FUENTES["tasa_objetivo"])
    mensaje_tasa = " · ".join(obtener_bullets(MENSAJES, "tasa_obj", "no fijo"))
    st.markdown(card_dato_simple_html(
        "Tasa de Referencia (Banxico)",
        f"{fmt_numero(val_tasa, 2)}%" if val_tasa is not None else "N/D",
        fmt_fecha_corta(fecha_tasa) if fecha_tasa is not None else "",
        mensaje_tasa,
    ), unsafe_allow_html=True)

# ==========================================================================
# 10. EMPLEO — Empleo IMSS, Desocupación (ENOE)
# ==========================================================================
section_header("👷", "Empleo")
col1, col2 = st.columns(2)

with col1:
    p_a, v_a = ultimo_valor_calculado(FUENTES["empleo"], "Var_Anual")
    p_m, v_m = ultimo_valor_calculado(FUENTES["empleo"], "Var_Mensual")
    bloques = [
        bloque("Variación Anual", v_a, "pct", fmt_periodo_mensual(p_a) if p_a else None, "empleo", "anual"),
        bloque("Variación Mensual", v_m, "pct", fmt_periodo_mensual(p_m) if p_m else None, "empleo", "mensual"),
    ]
    st.markdown(card_variable_html("Empleo IMSS (Puestos de Trabajo)", bloques, subtitulo="Trabajadores asegurados"), unsafe_allow_html=True)

with col2:
    p_n, v_n = ultimo_valor_generico(FUENTES["desocupacion"], "Tasa de Desocupación")
    bloques = [
        bloque("Nivel", v_n, "pct", fmt_periodo_mensual(p_n) if p_n else None, "desocupacion", "nivel"),
    ]
    st.markdown(card_variable_html("Desocupación (ENOE)", bloques), unsafe_allow_html=True)

# ==========================================================================
# 11. EXPECTATIVAS DE CRECIMIENTO ECONÓMICO (tabla manual)
# ==========================================================================
section_header("🔮", "Expectativas de Crecimiento Económico")

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

    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto;">
<table class="tabla-manual">
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
section_header("🏛️", "Finanzas Públicas 2019 – 2026 (% del PIB)")

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
    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto;">
<table class="tabla-manual">
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
section_header("⛽", "Pemex")

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

    html_tabla = f"""<div class="metric-container card-hover" style="overflow-x:auto;">
<table class="tabla-manual">
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
    st.caption("¹ Incluye producción de socios.  \n"
               "² Para los años 2018 a 2025 son cifras de estados financieros auditados. Para 2026 son cifras preliminares al 31 de marzo de 2026.  \n"
               "Fuente: Presentación a Inversionistas, Pemex (abril, 2026).")
else:
    st.warning("No se encontró `data/manual/cuadros.xlsx` (hoja `pemex`).")

# ==========================================================================
# FOOTER
# ==========================================================================
st.markdown("<hr>", unsafe_allow_html=True)
st.caption("Dirección de Estudios Económicos · NAFIN / BANCOMEXT — Dashboard generado automáticamente "
           "a partir de fuentes internas. Datos ficticios mientras se conectan las fuentes reales.")
