import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import unicodedata
import re
import sqlite3
import os
import io
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ─── PAGE CONFIG ─────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Productivity Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "agents.db")

# ─── CSS ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html,body,[data-testid="stAppViewContainer"]{background:#0d1117!important;font-family:'IBM Plex Sans',sans-serif;}
[data-testid="stSidebar"]{display:none;}
[data-testid="stHeader"]{background:transparent;}
[data-testid="block-container"]{padding:2rem 2.5rem 3rem!important;max-width:1700px;margin:0 auto;}

.dash-title{font-family:'IBM Plex Mono',monospace;font-size:1.85rem;font-weight:600;color:#e6edf3;letter-spacing:-.02em;}
.dash-sub{font-size:.82rem;color:#7d8590;font-family:'IBM Plex Mono',monospace;margin-bottom:2rem;}
.accent{color:#58a6ff;}

/* filter info pill */
.filter-pill{display:inline-block;background:#0f2236;border:1px solid #1f4068;border-radius:20px;padding:.25rem .85rem;font-family:'IBM Plex Mono',monospace;font-size:.72rem;color:#58a6ff;margin:.2rem .2rem .6rem 0;}
.filter-pill.green{background:#0f3320;border-color:#1f5c38;color:#3fb950;}

/* Upload cards */
.upload-card{background:#161b22;border:1px solid #21262d;border-radius:10px;padding:1.5rem;margin-bottom:1rem;position:relative;overflow:hidden;}
.upload-card::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;}
.upload-card.yeastar::before{background:linear-gradient(90deg,#58a6ff,#388bfd);}
.upload-card.glpi::before{background:linear-gradient(90deg,#3fb950,#2ea043);}
.card-label{font-family:'IBM Plex Mono',monospace;font-size:.7rem;font-weight:600;letter-spacing:.12em;text-transform:uppercase;margin-bottom:.5rem;}
.card-label.yeastar{color:#58a6ff;}
.card-label.glpi{color:#3fb950;}
.card-title{font-size:1rem;font-weight:600;color:#e6edf3;margin-bottom:.25rem;}
.card-desc{font-size:.8rem;color:#7d8590;margin-bottom:1rem;}

/* KPI */
.kpi-card{background:#161b22;border:1px solid #21262d;border-radius:10px;padding:1.1rem 1rem;text-align:center;}
.kpi-icon{font-size:1.4rem;margin-bottom:.3rem;display:block;}
.kpi-value{font-family:'IBM Plex Mono',monospace;font-size:1.6rem;font-weight:700;color:#e6edf3;line-height:1;margin-bottom:.3rem;}
.kpi-value.blue{color:#58a6ff;}.kpi-value.green{color:#3fb950;}.kpi-value.yellow{color:#d29922;}.kpi-value.red{color:#f85149;}.kpi-value.purple{color:#bc8cff;}
.kpi-label{font-size:.72rem;color:#7d8590;font-family:'IBM Plex Mono',monospace;text-transform:uppercase;letter-spacing:.06em;}

.section-header{font-family:'IBM Plex Mono',monospace;font-size:.7rem;font-weight:600;letter-spacing:.14em;text-transform:uppercase;color:#7d8590;border-bottom:1px solid #21262d;padding-bottom:.5rem;margin-bottom:1.25rem;margin-top:1.75rem;}
.divider{border:none;border-top:1px solid #21262d;margin:1.5rem 0;}

/* Agent badges */
.badge-active{background:#0f3320;color:#3fb950;border-radius:4px;padding:2px 8px;font-size:.7rem;font-family:'IBM Plex Mono',monospace;}
.badge-inactive{background:#2d1b1b;color:#f85149;border-radius:4px;padding:2px 8px;font-size:.7rem;font-family:'IBM Plex Mono',monospace;}

::-webkit-scrollbar{width:6px;height:6px;}
::-webkit-scrollbar-track{background:#0d1117;}
::-webkit-scrollbar-thumb{background:#30363d;border-radius:3px;}
.stButton>button{background:#21262d!important;color:#e6edf3!important;border:1px solid #30363d!important;border-radius:6px!important;font-family:'IBM Plex Mono',monospace!important;font-size:.8rem!important;}
.stButton>button:hover{background:#388bfd!important;border-color:#388bfd!important;}
</style>
""", unsafe_allow_html=True)


# ─── DATABASE ────────────────────────────────────────────────────────────────
SEED_AGENTS = [
    ("Alexander Cano Gutierrez",           "Alexander Cano Gutierrez",           "1021-Alexander Cano",                  True),
    ("Algiro Blandon Cuesta",              "Algiro Blandon Cuesta",              "1005-Algiro Blandon Cuesta",           True),
    ("Brian Nicolas Cardozo Cabrera",      "Brian Nicolas Cardozo Cabrera",      "1015-Brian Cardozo",                   True),
    ("Carlos David Casas Martinez",        "Carlos David Casas Martinez",        "1017-Carlos Casas",                    True),
    ("Cristian Camilo Estrella Gil",       "Cristian Camilo Estrella Gil",       "1013-Cristian Estrella",               True),
    ("Cristian Alexander Nuñez Mogollon", "Cristian Alexander Nuñez Mogollon", "1004-Cristian Alexander Nunez",        True),
    ("David Eliecer Aguirre Martinez",     "David Eliecer Aguirre Martinez",     "1024-David Eliecer Aguirre Martinez",  True),
    ("Diego Martinez",                     "Diego Alejandro Martinez Garcia",    "1008-Diego Martinez",                  True),
    ("Duvan Camilo Amaya Urrego",          "Duvan Camilo Amaya Urrego",          "1006-Camilo Amaya",                    True),
    ("Edwin Dario Rincon Sarrazola",       "Edwin Dario Rincon Sarrazola",       "1002-Edwin Rincon",                    True),
    ("Javier Fernando Novoa paez",         "Javier Fernando Novoa paez",         "1001-Javier Novoa",                    True),
    ("Johan David Velandia Ochoa",         "Johan David Velandia Ochoa",         "1012-David Velandia",                  True),
    ("Jorge Enrique Vargas Guevara",       "Jorge Enrique Vargas Guevara",       "1011-Jorge Vargas",                    True),
    ("Juan Pablo Alvarez Monsalve",        "Juan Pablo Alvarez Monsalve",        "1018-Juan Pablo Alvarez",              True),
    ("Kevin Josue Loaiza Balbin",          "Kevin Josue Loaiza Balbin",          "1022-Kevin Josue Loaiza Balvin",       True),
    ("Kevin Santiago Sanchez Castano",     "Kevin Santiago Sanchez Castaño",     "1014-Kevin Sanchez",                   True),
    ("Nicolas Yesid Hernandez Roncancio",  "Nicolas Yesid Hernandez Roncancio",  "1016-Nicolas Hernandez",               True),
    ("Santiago Andres Urrego",             "Santiago Andres Urrego",             "1010-Santiago Urrego",                 True),
    ("Sayra Jelanny Segura Contrera",      "Sayra Jelanny Segura Contrera",      "1003-Jelanny Segura",                  True),
]


def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)


NEW_YEASTAR_NAMES = {
    "Alexander Cano Gutierrez":          "1021-Alexander Cano",
    "Algiro Blandon Cuesta":             "1005-Algiro Blandon Cuesta",
    "Brian Nicolas Cardozo Cabrera":     "1015-Brian Cardozo",
    "Carlos David Casas Martinez":       "1017-Carlos Casas",
    "Cristian Camilo Estrella Gil":      "1013-Cristian Estrella",
    "Cristian Alexander Nuñez Mogollon":"1004-Cristian Alexander Nunez",
    "David Eliecer Aguirre Martinez":    "1024-David Eliecer Aguirre Martinez",
    "Diego Martinez":                    "1008-Diego Martinez",
    "Duvan Camilo Amaya Urrego":         "1006-Camilo Amaya",
    "Edwin Dario Rincon Sarrazola":      "1002-Edwin Rincon",
    "Javier Fernando Novoa paez":        "1001-Javier Novoa",
    "Johan David Velandia Ochoa":        "1012-David Velandia",
    "Jorge Enrique Vargas Guevara":      "1011-Jorge Vargas",
    "Juan Pablo Alvarez Monsalve":       "1018-Juan Pablo Alvarez",
    "Kevin Josue Loaiza Balbin":         "1022-Kevin Josue Loaiza Balvin",
    "Kevin Santiago Sanchez Castano":    "1014-Kevin Sanchez",
    "Nicolas Yesid Hernandez Roncancio": "1016-Nicolas Hernandez",
    "Santiago Andres Urrego":            "1010-Santiago Urrego",
    "Sayra Jelanny Segura Contrera":     "1003-Jelanny Segura",
}


def migrate_yeastar_names():
    """Update nombre_yeastar from old <ext> format to new ext- format in existing DB."""
    with get_conn() as con:
        rows = con.execute("SELECT id, agente_mesa, nombre_yeastar FROM agents").fetchall()
        updated = 0
        for row_id, agente_mesa, nombre_yeastar in rows:
            # Detect old format (contains '<') OR agente_mesa key exists in migration map
            new_name = NEW_YEASTAR_NAMES.get(agente_mesa)
            if new_name and new_name != nombre_yeastar:
                con.execute(
                    "UPDATE agents SET nombre_yeastar=? WHERE id=?",
                    (new_name, row_id)
                )
                updated += 1
        if updated:
            con.commit()
        return updated


def init_db():
    with get_conn() as con:
        con.execute("""
            CREATE TABLE IF NOT EXISTS agents (
                id             INTEGER PRIMARY KEY AUTOINCREMENT,
                agente_mesa    TEXT NOT NULL,
                nombre_glpi    TEXT NOT NULL,
                nombre_yeastar TEXT NOT NULL,
                activo         INTEGER NOT NULL DEFAULT 1
            )
        """)
        count = con.execute("SELECT COUNT(*) FROM agents").fetchone()[0]
        if count == 0:
            con.executemany(
                "INSERT INTO agents (agente_mesa,nombre_glpi,nombre_yeastar,activo) VALUES (?,?,?,?)",
                [(a, g, y, int(ac)) for a, g, y, ac in SEED_AGENTS]
            )
            con.commit()


def load_agents(only_active=False) -> pd.DataFrame:
    q = "SELECT id, agente_mesa, nombre_glpi, nombre_yeastar, activo FROM agents"
    if only_active:
        q += " WHERE activo=1"
    q += " ORDER BY agente_mesa"
    with get_conn() as con:
        return pd.read_sql_query(q, con)


def upsert_agent(agente_mesa, nombre_glpi, nombre_yeastar, activo, agent_id=None):
    with get_conn() as con:
        if agent_id:
            con.execute(
                "UPDATE agents SET agente_mesa=?,nombre_glpi=?,nombre_yeastar=?,activo=? WHERE id=?",
                (agente_mesa, nombre_glpi, nombre_yeastar, int(activo), agent_id)
            )
        else:
            con.execute(
                "INSERT INTO agents (agente_mesa,nombre_glpi,nombre_yeastar,activo) VALUES (?,?,?,?)",
                (agente_mesa, nombre_glpi, nombre_yeastar, int(activo))
            )
        con.commit()


def toggle_agent(agent_id, activo: bool):
    with get_conn() as con:
        con.execute("UPDATE agents SET activo=? WHERE id=?", (int(activo), agent_id))
        con.commit()


init_db()
migrate_yeastar_names()


# ─── UTILITIES ───────────────────────────────────────────────────────────────
def normalize(text: str) -> str:
    """Lowercase, strip accents, collapse spaces.
    Handles both old format  Alexander Cano<1021>
    and new Yeastar format   1021-Alexander Cano
    """
    if not isinstance(text, str):
        return ""
    text = text.strip().replace("\t", "").strip()
    text = text.lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    text = re.sub(r"<\d+>", "", text)        # old: Name<1021>
    text = re.sub(r"^\d{3,4}-\s*", "", text) # new: 1021-Name
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def build_lookup(agents_df: pd.DataFrame):
    """Returns y_lookup, g_lookup: norm_name → agente_mesa (canonical)."""
    y_lookup, g_lookup = {}, {}
    for _, row in agents_df.iterrows():
        canon = row["agente_mesa"]
        y_lookup[normalize(row["nombre_yeastar"])] = canon
        g_lookup[normalize(row["nombre_glpi"])]    = canon
        g_lookup[normalize(canon)]                  = canon
    return y_lookup, g_lookup


def clean_str_cols(df: pd.DataFrame) -> pd.DataFrame:
    """Strip whitespace and embedded tabs from all string columns."""
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].apply(
                lambda x: x.strip().replace("\t", "").strip() if isinstance(x, str) else x
            )
    return df


def load_csv(file) -> pd.DataFrame | None:
    encodings = ["utf-8", "latin-1", "cp1252"]
    separators = [",", ";", "\t"]
    for enc in encodings:
        for sep in separators:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc, sep=sep, engine="python")
                if len(df.columns) > 1:
                    return df
            except Exception:
                pass
    try:
        file.seek(0)
        return pd.read_csv(file)
    except Exception as e:
        st.error(f"Error leyendo CSV: {e}")
        return None


def load_csv_yeastar(file) -> pd.DataFrame | None:
    """Load Yeastar new-format CSV: data starts at row 6 (header = row 6)."""
    encodings = ["utf-8", "latin-1", "cp1252"]
    separators = [",", ";", "	"]
    for enc in encodings:
        for sep in separators:
            try:
                file.seek(0)
                df = pd.read_csv(file, encoding=enc, sep=sep, engine="python",
                                 skiprows=5, header=0)
                if len(df.columns) > 1:
                    return df
            except Exception:
                pass
    try:
        file.seek(0)
        return pd.read_csv(file, skiprows=5, header=0)
    except Exception as e:
        st.error(f"Error leyendo CSV Yeastar: {e}")
        return None


def find_col(df: pd.DataFrame, candidates: list) -> str | None:
    norm_map = {normalize(c): c for c in df.columns}
    for c in candidates:
        if normalize(c) in norm_map:
            return norm_map[normalize(c)]
    return None


# ─── SESSION STATE ────────────────────────────────────────────────────────────
for key in ["show_modal", "show_add_form", "show_edit_form"]:
    if key not in st.session_state:
        st.session_state[key] = False
if "edit_id" not in st.session_state:
    st.session_state.edit_id = None


# ─── HEADER ──────────────────────────────────────────────────────────────────
h1, h2 = st.columns([6, 1])
with h1:
    st.markdown("""
    <div class="dash-title">📡 <span class="accent">Productivity</span> Dashboard</div>
    <div class="dash-sub">// Yeastar Calls &amp; GLPI Tickets · Agent Performance Intelligence</div>
    """, unsafe_allow_html=True)
with h2:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("⚙️ Administrar agentes", use_container_width=True):
        st.session_state.show_modal     = not st.session_state.show_modal
        st.session_state.show_add_form  = False
        st.session_state.show_edit_form = False
        st.session_state.edit_id        = None

# ─── FILTER INFO PILLS ───────────────────────────────────────────────────────
st.markdown("""
<span class="filter-pill">📞 Yeastar → <b>lectura desde fila 6</b> · columna agente: <b>Agente</b> · conteo: <b>Contestada</b></span>
<span class="filter-pill green">🎫 GLPI → <b>Fuente de solicitud = Telefónico + Vacío</b> · columna: <b>Solicitante - Escritor</b></span>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════════════════
# AGENT MANAGEMENT PANEL
# ═══════════════════════════════════════════════════════════════════════════════
if st.session_state.show_modal:
    with st.container(border=True):
        st.markdown("### ⚙️ Administración de Agentes de Mesa de Servicio")
        st.caption("Base de datos persistente (SQLite · agents.db) · Cambios aplican de inmediato")

        agents_df = load_agents()

        add_c, _ = st.columns([1, 4])
        with add_c:
            if st.button("➕ Agregar nuevo agente", use_container_width=True):
                st.session_state.show_add_form  = not st.session_state.show_add_form
                st.session_state.show_edit_form = False
                st.session_state.edit_id        = None

        if st.session_state.show_add_form:
            with st.form("form_add", border=True):
                st.markdown("**Nuevo agente**")
                c1, c2, c3, c4 = st.columns([3, 3, 3, 1])
                new_mesa    = c1.text_input("Agente Mesa Servicio *", placeholder="Nombre canónico oficial")
                new_glpi    = c2.text_input("Nombre en GLPI *",       placeholder="Tal como aparece en CSV")
                new_yeastar = c3.text_input("Nombre en Yeastar *",    placeholder="ej: Javier Novoa<1001>")
                new_activo  = c4.checkbox("Activo", value=True)
                s, cancel   = st.columns(2)
                if s.form_submit_button("💾 Guardar", use_container_width=True):
                    if new_mesa and new_glpi and new_yeastar:
                        upsert_agent(new_mesa.strip(), new_glpi.strip(), new_yeastar.strip(), new_activo)
                        st.session_state.show_add_form = False
                        st.success(f"✅ '{new_mesa}' agregado.")
                        st.rerun()
                    else:
                        st.error("Completa los campos obligatorios *")
                if cancel.form_submit_button("Cancelar", use_container_width=True):
                    st.session_state.show_add_form = False
                    st.rerun()

        if st.session_state.show_edit_form and st.session_state.edit_id:
            row = agents_df[agents_df["id"] == st.session_state.edit_id].iloc[0]
            with st.form("form_edit", border=True):
                st.markdown(f"**Editar agente — ID {row['id']}**")
                c1, c2, c3, c4 = st.columns([3, 3, 3, 1])
                e_mesa    = c1.text_input("Agente Mesa Servicio", value=row["agente_mesa"])
                e_glpi    = c2.text_input("Nombre en GLPI",       value=row["nombre_glpi"])
                e_yeastar = c3.text_input("Nombre en Yeastar",    value=row["nombre_yeastar"])
                e_activo  = c4.checkbox("Activo",                 value=bool(row["activo"]))
                s2, c2b   = st.columns(2)
                if s2.form_submit_button("💾 Actualizar", use_container_width=True):
                    upsert_agent(e_mesa.strip(), e_glpi.strip(), e_yeastar.strip(), e_activo, agent_id=row["id"])
                    st.session_state.show_edit_form = False
                    st.session_state.edit_id        = None
                    st.success("✅ Agente actualizado.")
                    st.rerun()
                if c2b.form_submit_button("Cancelar", use_container_width=True):
                    st.session_state.show_edit_form = False
                    st.session_state.edit_id        = None
                    st.rerun()

        st.markdown('<div class="section-header">// Agentes registrados</div>', unsafe_allow_html=True)
        fc1, fc2 = st.columns([3, 1])
        search   = fc1.text_input("🔍 Buscar", placeholder="Filtrar por nombre…", label_visibility="collapsed")
        show_all = fc2.checkbox("Mostrar inactivos", value=True)

        disp = agents_df.copy()
        if not show_all:
            disp = disp[disp["activo"] == 1]
        if search:
            disp = disp[disp["agente_mesa"].str.lower().str.contains(search.lower())]

        hdr = st.columns([3, 3, 3, 1, 1, 1])
        for col, lbl in zip(hdr, ["Agente Mesa Servicio", "Nombre GLPI", "Nombre Yeastar", "Estado", "Editar", "Act/Des"]):
            col.markdown(f"<small style='color:#7d8590;font-family:monospace;text-transform:uppercase'>{lbl}</small>",
                         unsafe_allow_html=True)
        st.divider()

        for _, row in disp.iterrows():
            c = st.columns([3, 3, 3, 1, 1, 1])
            c[0].markdown(f"<span style='color:#e6edf3;font-size:.85rem'>{row['agente_mesa']}</span>", unsafe_allow_html=True)
            c[1].markdown(f"<span style='color:#8b949e;font-size:.82rem'>{row['nombre_glpi']}</span>", unsafe_allow_html=True)
            c[2].markdown(f"<code style='background:#161b22;color:#58a6ff;font-size:.78rem;padding:2px 6px;border-radius:4px'>{row['nombre_yeastar']}</code>", unsafe_allow_html=True)
            activo = bool(row["activo"])
            c[3].markdown(f"<span class='{'badge-active' if activo else 'badge-inactive'}'>{'Activo' if activo else 'Inactivo'}</span>", unsafe_allow_html=True)
            if c[4].button("✏️", key=f"e_{row['id']}"):
                st.session_state.edit_id        = row["id"]
                st.session_state.show_edit_form = True
                st.session_state.show_add_form  = False
                st.rerun()
            if c[5].button("🔴" if activo else "🟢", key=f"t_{row['id']}", help="Desactivar" if activo else "Activar"):
                toggle_agent(row["id"], not activo)
                st.rerun()

        st.markdown("---")
        st.download_button(
            "⬇️ Exportar base de agentes CSV",
            data=load_agents().to_csv(index=False).encode("utf-8"),
            file_name="agentes_mesa_servicio.csv",
            mime="text/csv",
        )

    st.markdown('<hr class="divider"/>', unsafe_allow_html=True)


# ─── UPLOAD SECTION ──────────────────────────────────────────────────────────
col_l, col_r = st.columns(2, gap="large")

with col_l:
    st.markdown("""
    <div class="upload-card yeastar">
      <div class="card-label yeastar">📞 Fuente 01 — Yeastar PBX</div>
      <div class="card-title">Reporte de Llamadas</div>
      <div class="card-desc">
        Lectura desde: <code style="color:#58a6ff;background:#0d1117;padding:1px 5px;border-radius:3px">Fila 6 (header automático)</code><br>
        Agente: <code style="color:#58a6ff;background:#0d1117;padding:1px 5px;border-radius:3px">Agente</code> &nbsp;·&nbsp;
        Conteo: <code style="color:#58a6ff;background:#0d1117;padding:1px 5px;border-radius:3px">Contestada</code>
      </div>
    </div>
    """, unsafe_allow_html=True)
    yeastar_file = st.file_uploader("Subir CSV Yeastar", type=["csv"], key="yeastar")

with col_r:
    st.markdown("""
    <div class="upload-card glpi">
      <div class="card-label glpi">🎫 Fuente 02 — GLPI</div>
      <div class="card-title">Casos / Tickets Creados</div>
      <div class="card-desc">
        Filtro activo: <code style="color:#3fb950;background:#0d1117;padding:1px 5px;border-radius:3px">Fuente de solicitud = Telefónico ó Vacío</code><br>
        Columna de agente: <code style="color:#3fb950;background:#0d1117;padding:1px 5px;border-radius:3px">Solicitante - Escritor</code>
      </div>
    </div>
    """, unsafe_allow_html=True)
    glpi_file = st.file_uploader("Subir CSV GLPI", type=["csv"], key="glpi")


# ─── LOAD & CLEAN CSVs ───────────────────────────────────────────────────────
df_yeastar = clean_str_cols(load_csv_yeastar(yeastar_file)) if yeastar_file else None
df_glpi    = clean_str_cols(load_csv(glpi_file))            if glpi_file    else None


# ─── APPLY FILTERS ───────────────────────────────────────────────────────────
df_y_filtered = None
df_g_filtered = None
y_filter_info = ""
g_filter_info = ""

if df_yeastar is not None:
    col_y_agente     = find_col(df_yeastar, ["Agente", "agente", "agent"])
    col_y_contestada = find_col(df_yeastar, ["Contestada", "contestada", "answered", "contestadas"])

    if col_y_agente and col_y_contestada:
        # Drop rows where Agente is empty / NaN — these are subtotal/total rows
        df_y_filtered = df_yeastar[
            df_yeastar[col_y_agente].notna() &
            (df_yeastar[col_y_agente].astype(str).str.strip() != "") &
            (df_yeastar[col_y_agente].astype(str).str.strip().str.lower() != "nan")
        ].copy()
        n_agents = df_y_filtered[col_y_agente].nunique()
        y_filter_info = (
            f"✅ Lectura desde fila 6 · **{col_y_agente}** (agente) + "
            f"**{col_y_contestada}** (conteo) → {n_agents} agentes / {len(df_y_filtered):,} filas"
        )
    else:
        df_y_filtered = df_yeastar.copy()
        missing = []
        if not col_y_agente:     missing.append("'Agente'")
        if not col_y_contestada: missing.append("'Contestada'")
        y_filter_info = f"⚠️ Columnas no encontradas: {', '.join(missing)} · Columnas disponibles: {list(df_yeastar.columns)}"

if df_glpi is not None:
    col_g_fuente  = find_col(df_glpi, ["Fuente de solicitud", "fuente de solicitud", "fuente solicitud"])
    col_g_escrito = find_col(df_glpi, [
        "Solicitante - Escritor", "solicitante - escritor",
        "Escritor - Solicitante", "escritor - solicitante",
        "Solicitante", "solicitante"
    ])

    if col_g_fuente and col_g_escrito:
        fuente_col = df_glpi[col_g_fuente]
        mask_telefónico = fuente_col.astype(str).str.strip() == "Telefónico"
        mask_vacío      = fuente_col.isna() | (fuente_col.astype(str).str.strip() == "")
        df_g_filtered   = df_glpi[mask_telefónico | mask_vacío].copy()
        n_tel  = int(mask_telefónico.sum())
        n_vac  = int(mask_vacío.sum())
        g_filter_info = (
            f"✅ Filtro aplicado · **{col_g_fuente} = Telefónico ({n_tel}) + Vacío ({n_vac})** "
            f"→ {len(df_g_filtered):,} / {len(df_glpi):,} registros"
        )
    else:
        col_g_escrito = find_col(df_glpi, ["Solicitante - Escritor","Escritor - Solicitante","Solicitante"])
        df_g_filtered = df_glpi.copy()
        g_filter_info = "⚠️ Columna 'Fuente de solicitud' no encontrada — usando todos los registros"


# ─── PREVIEW ─────────────────────────────────────────────────────────────────
if df_yeastar is not None or df_glpi is not None:
    st.markdown('<div class="section-header">// Vista previa · registros filtrados</div>', unsafe_allow_html=True)
    p1, p2 = st.columns(2, gap="large")

    if df_yeastar is not None and df_y_filtered is not None:
        with p1:
            st.info(y_filter_info)
            _col_ag = find_col(df_y_filtered, ["Agente","agente","agent"])
            _col_co = find_col(df_y_filtered, ["Contestada","contestada","answered"])
            if _col_ag and _col_co:
                st.success(f"Agente: **{_col_ag}** · Conteo: **{_col_co}**")
            # Show only Agente + Contestada preview
            preview_cols = [c for c in [_col_ag, _col_co] if c]
            st.dataframe(
                df_y_filtered[preview_cols].head(10) if preview_cols else df_y_filtered.head(6),
                use_container_width=True, hide_index=True
            )

    if df_glpi is not None and df_g_filtered is not None:
        with p2:
            st.info(g_filter_info)
            col_g_escrito = find_col(df_g_filtered, [
                "Solicitante - Escritor","solicitante - escritor",
                "Escritor - Solicitante","Solicitante"
            ])
            if col_g_escrito:
                st.success(f"Columna de agente: **{col_g_escrito}**")
            st.dataframe(df_g_filtered.head(6), use_container_width=True, hide_index=True)


# ─── DASHBOARD ───────────────────────────────────────────────────────────────
if df_y_filtered is not None and df_g_filtered is not None:

    col_y_agente     = find_col(df_y_filtered, ["Agente","agente","agent"])
    col_y_contestada = find_col(df_y_filtered, ["Contestada","contestada","answered","contestadas"])
    col_g_escrito    = find_col(df_g_filtered, [
        "Solicitante - Escritor","solicitante - escritor",
        "Escritor - Solicitante","Solicitante","solicitante"
    ])

    if not col_y_agente or not col_y_contestada:
        st.error(f"❌ Columnas Yeastar no encontradas. Disponibles: {list(df_y_filtered.columns)}")
        st.stop()
    if not col_g_escrito:
        st.error("❌ Columna 'Solicitante - Escritor' no encontrada en GLPI.")
        st.stop()

    # ── Lookups ───────────────────────────────────────────────────────────────
    agents   = load_agents(only_active=True)
    y_lookup, g_lookup = build_lookup(agents)

    # ── Yeastar counts: pre-aggregated, read Agente + Contestada directly ─────
    df_y_work = df_y_filtered[[col_y_agente, col_y_contestada]].copy()
    df_y_work[col_y_contestada] = pd.to_numeric(df_y_work[col_y_contestada], errors="coerce").fillna(0).astype(int)
    df_y_work["agente_mesa"] = df_y_work[col_y_agente].astype(str).apply(
        lambda x: y_lookup.get(normalize(x))
    )
    unmapped_y = df_y_work[df_y_work["agente_mesa"].isna()][col_y_agente].value_counts().head(20)
    y_counts   = (
        df_y_work.dropna(subset=["agente_mesa"])
        .groupby("agente_mesa")[col_y_contestada]
        .sum()
        .reset_index()
    )
    y_counts.columns = ["agente_mesa", "llamadas"]

    # ── GLPI counts ───────────────────────────────────────────────────────────
    g_raw    = df_g_filtered[col_g_escrito].dropna().astype(str)
    g_mapped = g_raw.apply(lambda x: g_lookup.get(normalize(x)))
    g_counts = g_mapped.dropna().value_counts().reset_index()
    g_counts.columns = ["agente_mesa", "registros"]
    unmapped_g = g_raw[g_mapped.isna()].value_counts().head(20)

    # ── Merge: all active agents as base ──────────────────────────────────────
    base    = agents[["agente_mesa"]].copy()
    merged  = base.merge(y_counts, on="agente_mesa", how="left") \
                  .merge(g_counts, on="agente_mesa", how="left")
    merged["llamadas"]  = merged["llamadas"].fillna(0).astype(int)
    merged["registros"] = merged["registros"].fillna(0).astype(int)

    # Only agents with at least one call or ticket
    merged = merged[(merged["llamadas"] > 0) | (merged["registros"] > 0)].copy()

    # ── Metrics ───────────────────────────────────────────────────────────────
    merged["Diferencia"] = merged["registros"] - merged["llamadas"]
    merged["% Registro"] = merged.apply(
        lambda r: round((r["registros"] / r["llamadas"]) * 100, 1) if r["llamadas"] > 0 else None, axis=1
    )
    merged["Faltan"]     = (merged["llamadas"] - merged["registros"]).clip(lower=0).astype(int)
    merged["Excedentes"] = (merged["registros"] - merged["llamadas"]).clip(lower=0).astype(int)
    merged["Estado"]     = merged["% Registro"].apply(
        lambda p: "✅ Cumple" if p is not None and p >= 100
        else ("❌ Le hace falta" if p is not None else "⚪ Sin llamadas")
    )
    merged = merged.sort_values("llamadas", ascending=False).reset_index(drop=True)

    # ── KPIs ──────────────────────────────────────────────────────────────────
    total_llamadas  = merged["llamadas"].sum()
    total_registros = merged["registros"].sum()
    pct_vals        = merged["% Registro"].dropna()
    prom_pct        = pct_vals.mean() if len(pct_vals) else 0
    cumplen         = int((pct_vals >= 100).sum())
    total_agentes   = len(merged)
    total_faltantes = int(merged["Faltan"].sum())

    st.markdown('<div class="section-header">// KPIs · Resumen Ejecutivo</div>', unsafe_allow_html=True)
    kpi_cols = st.columns(6, gap="small")
    kpis = [
        ("📞", f"{total_llamadas:,}",       "Total Llamadas Contestadas", "blue"),
        ("🎫", f"{total_registros:,}",      "Total Registros GLPI",    "green"),
        ("📈", f"{prom_pct:.1f}%",          "Cumplimiento Promedio",   "yellow"),
        ("✅", f"{cumplen}/{total_agentes}", "Agentes que Cumplen",     "green"),
        ("⚠️", f"{total_faltantes:,}",      "Registros Faltantes",     "red"),
        ("👥", str(total_agentes),           "Usuarios Evaluados",      "purple"),
    ]
    for col, (icon, val, lbl, color) in zip(kpi_cols, kpis):
        with col:
            st.markdown(f"""
            <div class="kpi-card">
              <span class="kpi-icon">{icon}</span>
              <div class="kpi-value {color}">{val}</div>
              <div class="kpi-label">{lbl}</div>
            </div>""", unsafe_allow_html=True)

    # ── Unmapped warnings ─────────────────────────────────────────────────────
    if len(unmapped_y) > 0 or len(unmapped_g) > 0:
        with st.expander(f"⚠️ Nombres sin homologar: {len(unmapped_y)} en Yeastar · {len(unmapped_g)} en GLPI"):
            wu1, wu2 = st.columns(2)
            if len(unmapped_y) > 0:
                wu1.markdown("**📞 Yeastar — no encontrados en base de agentes:**")
                wu1.dataframe(
                    unmapped_y.reset_index().rename(columns={"index": "Agente (Yeastar)", "count": "Ocurrencias", col_y_agente: "Agente (Yeastar)"}),
                    use_container_width=True, hide_index=True
                )
            if len(unmapped_g) > 0:
                wu2.markdown("**🎫 GLPI — no encontrados en base de agentes:**")
                wu2.dataframe(
                    unmapped_g.reset_index().rename(columns={"index": "Solicitante - Escritor", "count": "Ocurrencias"}),
                    use_container_width=True, hide_index=True
                )
            st.info("💡 Usa **⚙️ Administrar agentes** para agregar o corregir los nombres no reconocidos.")

    # ── Main Table ────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">// Dashboard Central · Productividad por Agente</div>', unsafe_allow_html=True)

    display = merged.rename(columns={
        "agente_mesa": "Agente Mesa Servicio",
        "registros":   "GLPI Registros",
        "llamadas":    "Llamadas",
    })[["Agente Mesa Servicio","GLPI Registros","Llamadas","Diferencia","% Registro","Faltan","Excedentes","Estado"]].copy()

    display["% Registro"] = display["% Registro"].apply(lambda x: f"{x:.1f}%" if x is not None else "N/A")

    # ── Row coloring: red text for "Le hace falta" ────────────────────────────
    def style_rows(row):
        if "Le hace falta" in str(row["Estado"]):
            return ["color: #f85149; font-weight: 600"] * len(row)
        elif "Cumple" in str(row["Estado"]):
            return ["color: #3fb950"] * len(row)
        else:
            return ["color: #7d8590; font-style: italic"] * len(row)

    styled = display.style.apply(style_rows, axis=1)

    st.dataframe(
        styled,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Agente Mesa Servicio": st.column_config.TextColumn("👤 Agente Mesa Servicio", width="large"),
            "GLPI Registros":       st.column_config.NumberColumn("🎫 GLPI",       format="%d"),
            "Llamadas":             st.column_config.NumberColumn("📞 Llamadas",   format="%d"),
            "Diferencia":           st.column_config.NumberColumn("∆ Diferencia",  format="%d"),
            "% Registro":           st.column_config.TextColumn("% Registro"),
            "Faltan":               st.column_config.NumberColumn("⚠️ Faltan",    format="%d"),
            "Excedentes":           st.column_config.NumberColumn("➕ Excedentes", format="%d"),
            "Estado":               st.column_config.TextColumn("Estado"),
        },
        height=min(650, 56 + 36 * len(display)),
    )

    # ── Export buttons ────────────────────────────────────────────────────────
    def build_excel(df: pd.DataFrame) -> bytes:
        """Build a formatted .xlsx: green rows = cumple, red rows = le hace falta."""

        wb  = Workbook()
        ws  = wb.active
        ws.title = "Productividad"

        # ── Colour palette ────────────────────────────────────────────────────
        CLR_HEADER_BG  = "0D1117"
        CLR_HEADER_FG  = "E6EDF3"
        CLR_CUMPLE_BG  = "0F3320"
        CLR_CUMPLE_FG  = "3FB950"
        CLR_FALTA_BG   = "2D1B1B"
        CLR_FALTA_FG   = "F85149"
        CLR_SIN_BG     = "1C2128"
        CLR_SIN_FG     = "7D8590"
        CLR_ROW_ALT    = "161B22"
        CLR_ROW_NORM   = "0D1117"
        CLR_BORDER     = "21262D"

        thin = Side(style="thin", color=CLR_BORDER)
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        # ── Header row ────────────────────────────────────────────────────────
        col_widths = [32, 14, 12, 13, 12, 10, 12, 18]
        headers    = list(df.columns)

        for ci, (hdr, width) in enumerate(zip(headers, col_widths), start=1):
            cell = ws.cell(row=1, column=ci, value=hdr)
            cell.font      = Font(bold=True, color=CLR_HEADER_FG, name="Calibri", size=10)
            cell.fill      = PatternFill("solid", fgColor=CLR_HEADER_BG)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border    = border
            ws.column_dimensions[get_column_letter(ci)].width = width

        ws.row_dimensions[1].height = 28

        # ── Data rows ─────────────────────────────────────────────────────────
        for ri, (_, row) in enumerate(df.iterrows(), start=2):
            estado = str(row.get("Estado", ""))

            if "Le hace falta" in estado:
                bg, fg = CLR_FALTA_BG, CLR_FALTA_FG
            elif "Cumple" in estado:
                bg, fg = CLR_CUMPLE_BG, CLR_CUMPLE_FG
            else:
                bg, fg = CLR_SIN_BG, CLR_SIN_FG

            row_fill = PatternFill("solid", fgColor=bg)
            row_font_base = Font(color=fg, name="Calibri", size=10)
            row_font_bold = Font(color=fg, name="Calibri", size=10, bold=True)

            for ci, val in enumerate(row, start=1):
                cell = ws.cell(row=ri, column=ci, value=val)
                cell.fill      = row_fill
                cell.border    = border
                cell.alignment = Alignment(horizontal="center" if ci > 1 else "left",
                                           vertical="center")
                # Bold the agente name and estado columns
                cell.font = row_font_bold if ci in (1, len(headers)) else row_font_base

            ws.row_dimensions[ri].height = 20

        # ── Freeze header + auto-filter ───────────────────────────────────────
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

        # ── Legend sheet ──────────────────────────────────────────────────────
        ws_leg = wb.create_sheet("Leyenda")
        legend = [
            ("Color",        "Significado",            "Criterio"),
            ("🟢 Verde",     "Cumple",                 "% Registro ≥ 100%"),
            ("🔴 Rojo",      "Le hace falta",          "% Registro < 100%"),
            ("⚪ Gris",      "Sin llamadas Inbound",   "0 llamadas registradas"),
        ]
        for ri, (color, sig, crit) in enumerate(legend, start=1):
            ws_leg.cell(ri, 1, color)
            ws_leg.cell(ri, 2, sig)
            ws_leg.cell(ri, 3, crit)
            for ci in range(1, 4):
                ws_leg.cell(ri, ci).font      = Font(bold=(ri==1), name="Calibri", size=10, color="E6EDF3")
                ws_leg.cell(ri, ci).fill      = PatternFill("solid", fgColor="161B22")
                ws_leg.cell(ri, ci).alignment = Alignment(horizontal="left", vertical="center")
                ws_leg.cell(ri, ci).border    = border
        for ci, w in enumerate([14, 24, 30], start=1):
            ws_leg.column_dimensions[get_column_letter(ci)].width = w

        buf = io.BytesIO()
        wb.save(buf)
        return buf.getvalue()

    dl1, dl2, _ = st.columns([1, 1, 4])
    dl1.download_button(
        "⬇️ Exportar CSV",
        data=display.to_csv(index=False).encode("utf-8"),
        file_name="dashboard_productividad.csv",
        mime="text/csv",
    )
    dl2.download_button(
        "📊 Exportar Excel",
        data=build_excel(display),
        file_name="dashboard_productividad.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── Charts ────────────────────────────────────────────────────────────────
    st.markdown('<div class="section-header">// Visualizaciones</div>', unsafe_allow_html=True)
    ch1, ch2 = st.columns([2, 1], gap="large")

    chart_data = merged.sort_values("llamadas", ascending=True).tail(20)

    with ch1:
        fig_bar = go.Figure()
        fig_bar.add_trace(go.Bar(
            y=chart_data["agente_mesa"], x=chart_data["llamadas"],
            name="📞 Llamadas Contestadas", orientation="h",
            marker=dict(color="#388bfd", opacity=.85),
            hovertemplate="<b>%{y}</b><br>Llamadas: %{x}<extra></extra>",
        ))
        fig_bar.add_trace(go.Bar(
            y=chart_data["agente_mesa"], x=chart_data["registros"],
            name="🎫 GLPI Telefónico + Vacío", orientation="h",
            marker=dict(color="#3fb950", opacity=.85),
            hovertemplate="<b>%{y}</b><br>Registros: %{x}<extra></extra>",
        ))
        fig_bar.update_layout(
            title=dict(text="Llamadas Contestadas vs Registros GLPI (Telefónico + Vacío) por Agente",
                       font=dict(family="IBM Plex Mono", size=13, color="#e6edf3")),
            barmode="group", paper_bgcolor="#161b22", plot_bgcolor="#161b22",
            font=dict(family="IBM Plex Sans", color="#7d8590", size=11),
            legend=dict(bgcolor="#0d1117", bordercolor="#21262d", borderwidth=1, font=dict(color="#e6edf3")),
            xaxis=dict(gridcolor="#21262d", zerolinecolor="#21262d"),
            yaxis=dict(gridcolor="#21262d"),
            margin=dict(l=10, r=10, t=50, b=10),
            height=max(350, 34 * len(chart_data)),
        )
        st.plotly_chart(fig_bar, use_container_width=True)

    with ch2:
        nc = int((pct_vals < 100).sum())
        sd = int(merged["% Registro"].isna().sum())
        labels, values, colors = [], [], []
        if cumplen > 0: labels.append("✅ Cumple");       values.append(cumplen); colors.append("#3fb950")
        if nc > 0:      labels.append("❌ Le hace falta"); values.append(nc);      colors.append("#f85149")
        if sd > 0:      labels.append("⚪ Sin llamadas");  values.append(sd);      colors.append("#484f58")

        fig_pie = go.Figure(go.Pie(
            labels=labels, values=values, hole=.55,
            marker=dict(colors=colors, line=dict(color="#0d1117", width=2)),
            hovertemplate="<b>%{label}</b><br>%{value} agentes<br>%{percent}<extra></extra>",
        ))
        fig_pie.update_layout(
            title=dict(text="Cumplimiento de Agentes", font=dict(family="IBM Plex Mono", size=13, color="#e6edf3")),
            paper_bgcolor="#161b22", plot_bgcolor="#161b22",
            font=dict(family="IBM Plex Sans", color="#7d8590"),
            legend=dict(bgcolor="#0d1117", bordercolor="#21262d", borderwidth=1, font=dict(color="#e6edf3"), orientation="h", y=-.1),
            margin=dict(l=10, r=10, t=50, b=40), height=360,
            annotations=[dict(text=f"<b>{total_agentes}</b><br>agentes", x=.5, y=.5,
                              font=dict(family="IBM Plex Mono", size=14, color="#e6edf3"), showarrow=False)],
        )
        st.plotly_chart(fig_pie, use_container_width=True)

    # % Registro bar
    pct_c = merged[merged["% Registro"].notna()].sort_values("% Registro", ascending=True)
    if len(pct_c) > 0:
        fig_pct = go.Figure(go.Bar(
            y=pct_c["agente_mesa"], x=pct_c["% Registro"], orientation="h",
            marker=dict(color=["#f85149" if p < 100 else "#3fb950" for p in pct_c["% Registro"]], opacity=.9),
            hovertemplate="<b>%{y}</b><br>Cumplimiento: %{x:.1f}%<extra></extra>",
        ))
        fig_pct.add_vline(x=100, line=dict(color="#d29922", width=2, dash="dash"),
                          annotation=dict(text="meta 100%", font=dict(color="#d29922", family="IBM Plex Mono", size=11), xanchor="left"))
        fig_pct.update_layout(
            title=dict(text="% Registro por Agente", font=dict(family="IBM Plex Mono", size=13, color="#e6edf3")),
            paper_bgcolor="#161b22", plot_bgcolor="#161b22",
            font=dict(family="IBM Plex Sans", color="#7d8590", size=11),
            xaxis=dict(gridcolor="#21262d", zerolinecolor="#21262d", ticksuffix="%"),
            yaxis=dict(gridcolor="#21262d"),
            margin=dict(l=10, r=10, t=50, b=10),
            height=max(300, 30 * len(pct_c)), showlegend=False,
        )
        st.plotly_chart(fig_pct, use_container_width=True)

# ─── EMPTY STATE ─────────────────────────────────────────────────────────────
elif df_yeastar is None and df_glpi is None:
    st.markdown("""
    <div style="text-align:center;padding:4rem 2rem;color:#7d8590;">
      <div style="font-size:3rem;margin-bottom:1rem">📂</div>
      <div style="font-family:'IBM Plex Mono',monospace;font-size:1rem;color:#e6edf3;margin-bottom:.5rem">
        Sube los dos archivos CSV para comenzar
      </div>
      <div style="font-size:.85rem">Columna izquierda → Yeastar &nbsp;·&nbsp; Columna derecha → GLPI</div>
    </div>
    """, unsafe_allow_html=True)
elif df_yeastar is not None:
    st.info("📂 Yeastar cargado. Sube el CSV de GLPI para ver el dashboard completo.")
elif df_glpi is not None:
    st.info("📂 GLPI cargado. Sube el CSV de Yeastar para ver el dashboard completo.")

# ─── FOOTER ──────────────────────────────────────────────────────────────────
st.markdown("""
<hr class="divider"/>
<div style="text-align:center;font-family:'IBM Plex Mono',monospace;font-size:.7rem;color:#484f58;padding-bottom:1rem">
  Productivity Dashboard · Yeastar (Agente + Contestada) ↔ GLPI (Telefónico + Vacío) · SQLite persistente
</div>
""", unsafe_allow_html=True)
