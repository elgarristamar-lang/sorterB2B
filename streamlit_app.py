# Version: 0.07 — MANGO corporate style
import streamlit as st
import subprocess, sys, tempfile, io, re
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

BASE_DIR = Path(__file__).parent

# ── DXC CSV conversion ────────────────────────────────────────────────────────
def gd_to_dxc_csv(xlsx_bytes):
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))[1:]
    postex_lines, sorexp_lines = [], []
    for r in rows:
        if len(r) < 7: continue
        desc     = str(r[2] or "").strip()
        tipo     = str(r[3] or "").strip().upper()
        dest_raw = str(r[4] or "").strip()
        elem     = str(r[6] or "").strip()
        if not desc or tipo not in ("POSTEX", "SOREXP") or not dest_raw or not elem: continue
        if desc.startswith("="): continue
        d = "".join(c for c in dest_raw if c.isdigit())
        dest8 = d.zfill(10)[-8:] if d else dest_raw[:8]
        line = f"{desc};{dest8};00;{elem}"
        if tipo == "POSTEX": postex_lines.append(line + ";20")
        else:                sorexp_lines.append(line + ";10")
    bom = "\ufeff"
    def _enc(lines): return (bom + "\r\n".join(lines)).encode("utf-8")
    return _enc(postex_lines), _enc(sorexp_lines)


# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")

MANGO_CSS = """
<style>
  /* ── Fonts & background ── */
  html, body, .stApp, [class*="css"] {
    font-family: "Mango New", "Aptos Display", "Trebuchet MS", Arial, sans-serif !important;
    background: #f0f0f0 !important;
  }
  /* ── White card: all known Streamlit container selectors ── */
  .block-container,
  .stMainBlockContainer,
  div[data-testid="stAppViewBlockContainer"],
  div[data-testid="stMainBlockContainer"] {
    max-width: 820px !important;
    width: 820px !important;
    padding: 0 !important;
    margin: 0 auto !important;
    background: #ffffff !important;
    box-shadow: none !important;
  }
  /* ── Kill ALL vertical gaps & margins between widgets ── */
  .stVerticalBlock,
  .stVerticalBlockBorderWrapper,
  div[data-testid="stVerticalBlock"] {
    gap: 0 !important;
    row-gap: 0 !important;
  }
  .element-container,
  div[data-testid="element-container"] {
    margin: 0 !important;
    padding: 0 !important;
  }
  /* ── Remove top padding Streamlit adds to the page ── */
  div[data-testid="stAppViewContainer"] > section > div:first-child,
  .stApp > div > div > div { padding-top: 0 !important; }
  /* ── Column gaps ── */
  div[data-testid="stHorizontalBlock"] { gap: 16px !important; }
  .mng-header {
    background: #000;
    color: #fff;
    padding: 32px 40px 28px 40px;
  }
  .mng-header .eyebrow {
    font-size: 10px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #888;
    margin-bottom: 10px;
  }
  .mng-header h1 {
    font-size: 28px;
    font-weight: 400;
    color: #fff;
    margin: 0 0 6px 0;
    line-height: 1.2;
  }
  .mng-header .sub {
    font-size: 13px;
    color: #999;
    font-weight: 400;
  }
  .mng-section {
    font-size: 10px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #999;
    padding: 28px 40px 10px 40px;
    border-top: 1px solid #ebebeb;
  }
  .mng-content { padding: 0 40px 28px 40px; }
  .mng-action-label {
    font-size: 12px;
    font-weight: 400;
    color: #1a1a1a;
    letter-spacing: 0.02em;
    margin-bottom: 2px;
  }
  .mng-action-desc {
    font-size: 11px;
    color: #999;
    margin-bottom: 10px;
  }
  .mng-action-note {
    font-size: 10px;
    color: #bbb;
    letter-spacing: 0.04em;
    text-transform: uppercase;
    margin-top: 6px;
  }
  .mng-result-band {
    background: #000;
    color: #fff;
    padding: 14px 40px;
    font-size: 11px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
  }
  .mng-result-band .semana { font-size: 13px; color: #fff; }
  .mng-dl-label {
    font-size: 10px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #999;
    margin: 16px 0 8px 0;
    padding-bottom: 6px;
    border-bottom: 1px solid #ebebeb;
  }
  .mng-status-row {
    display: flex;
    gap: 0;
    border-top: 1px solid #ebebeb;
    border-bottom: 1px solid #ebebeb;
    margin: 20px 0 0 0;
  }
  .mng-status-item {
    flex: 1;
    padding: 12px 16px;
    border-right: 1px solid #ebebeb;
    font-size: 11px;
    color: #ccc;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }
  .mng-status-item:last-child { border-right: none; }
  .mng-status-item.ready { color: #1a1a1a; }
  .mng-status-item.ready::before { content: "— "; }
  .mng-status-item.missing::before { content: "· "; color: #ddd; }
  .mng-footer {
    background: #000;
    color: #555;
    font-size: 9px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    text-align: right;
    padding: 16px 40px;
    margin-top: 40px;
  }
  div[data-testid="stFileUploader"] label p {
    font-size: 11px !important;
    font-weight: 400 !important;
    color: #888 !important;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }
  div[data-testid="stFileUploaderDropzone"] {
    border: 1px solid #d0d0d0 !important;
    border-radius: 0 !important;
    background: #fff !important;
  }
  div[data-testid="stFileUploaderDropzone"]:hover {
    border-color: #1a1a1a !important;
  }
  div[data-testid="stFileUploaderDropzone"] button {
    border-radius: 0 !important;
    border: 1px solid #1a1a1a !important;
    background: #fff !important;
    color: #1a1a1a !important;
    font-size: 11px !important;
    letter-spacing: 0.06em;
    text-transform: uppercase;
  }
  div[data-testid="stFileUploaderDropzone"] button:hover {
    background: #000 !important;
    color: #fff !important;
  }
  div[data-testid="stButton"] > button[kind="primary"] {
    background: #000 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 0 !important;
    font-size: 11px !important;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    padding: 12px 20px !important;
    font-family: "Mango New", "Aptos Display", "Trebuchet MS", Arial, sans-serif !important;
    width: 100%;
  }
  div[data-testid="stButton"] > button[kind="primary"]:hover { background: #333 !important; }
  div[data-testid="stButton"] > button[kind="primary"]:disabled {
    background: #e0e0e0 !important;
    color: #aaa !important;
  }
  div[data-testid="stDownloadButton"] > button {
    background: #fff !important;
    color: #1a1a1a !important;
    border: 1px solid #1a1a1a !important;
    border-radius: 0 !important;
    font-size: 11px !important;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    padding: 10px 16px !important;
    font-family: "Mango New", "Aptos Display", "Trebuchet MS", Arial, sans-serif !important;
    width: 100%;
  }
  div[data-testid="stDownloadButton"] > button:hover {
    background: #000 !important;
    color: #fff !important;
    border-color: #000 !important;
  }
  div[data-testid="stButton"] > button:not([kind="primary"]) {
    background: #fff !important;
    color: #1a1a1a !important;
    border: 1px solid #1a1a1a !important;
    border-radius: 0 !important;
    font-size: 11px !important;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    padding: 10px 16px !important;
    font-family: "Mango New", "Aptos Display", "Trebuchet MS", Arial, sans-serif !important;
    width: 100%;
  }
  div[data-testid="stButton"] > button:not([kind="primary"]):hover {
    background: #000 !important;
    color: #fff !important;
  }
  div[data-testid="stSelectbox"] > div > div {
    border-radius: 0 !important;
    border: 1px solid #1a1a1a !important;
  }
  div[data-testid="stMultiSelect"] > div > div {
    border-radius: 0 !important;
    border: 1px solid #1a1a1a !important;
  }
  div[data-testid="stMultiSelect"] span[data-baseweb="tag"] {
    border-radius: 0 !important;
    background: #000 !important;
    color: #fff !important;
    font-size: 11px !important;
  }
  div[data-testid="stExpander"] {
    border: 1px solid #ebebeb !important;
    border-radius: 0 !important;
    background: #fff !important;
  }
  div[data-testid="stExpander"] summary {
    font-size: 11px !important;
    letter-spacing: 0.08em;
    text-transform: uppercase;
    color: #1a1a1a !important;
    font-weight: 400 !important;
  }
  div[data-testid="stAlert"] { border-radius: 0 !important; }
  div[data-testid="stCaptionContainer"] p {
    font-size: 11px !important;
    color: #999 !important;
  }
  label[data-testid="stWidgetLabel"] p {
    font-size: 11px !important;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: #888 !important;
  }
  hr { border: none; border-top: 1px solid #ebebeb !important; margin: 0 !important; }
  /* Columns inner padding — mng-content handles outer */
  div[data-testid="column"] { padding: 0 8px !important; }
  div[data-testid="column"]:first-child { padding-left: 0 !important; }
  div[data-testid="column"]:last-child  { padding-right: 0 !important; }
  /* Expander: no extra margin */
  div[data-testid="stExpander"] { margin: 0 !important; }
  /* Streamlit injects a wrapping div with padding-top — nuke it */
  div[data-testid="stVerticalBlock"] > div { padding-top: 0 !important; }
  /* Hide Streamlit chrome */
  #MainMenu, footer, header { visibility: hidden !important; }
  .stDeployButton { display: none !important; }
  /* Tight spacing between uploaders */
  div[data-testid="stFileUploader"] { margin-bottom: 16px !important; }
</style>
"""
st.markdown(MANGO_CSS, unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="mng-header">
  <div class="eyebrow">VDL B2B · Logística</div>
  <h1>Sorter VDL B2B</h1>
  <div class="sub">Configurador de semanas especiales</div>
</div>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
_STATE_KEYS = ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map",
               "r1_day_filter","r1_postex_csv","r1_sorexp_csv",
               "r1_esp_postex_csv","r1_esp_sorexp_csv"]
for k in _STATE_KEYS:
    if k not in st.session_state:
        st.session_state[k] = None

# ── Usage guide ───────────────────────────────────────────────────────────────
st.markdown('<div style="height:20px;"></div>', unsafe_allow_html=True)
with st.expander("Cómo usar esta herramienta", expanded=False):
    st.markdown("""
Esta herramienta genera la configuración del sorter VDL B2B para semanas con salidas canceladas
o que cambian de día (festivos, Semana Santa, etc.).

---

**1 · Sube los ficheros de entrada**

| Fichero | Obligatorio | Descripción |
|---|:---:|---|
| Parrilla de salidas (.xlsx) | ✅ | Fichero `parrilla_de_salidas.xlsx`. Columnas: `PLAYA`, `DIA_SALIDA`, `DIA_SALIDA_NEW`, `TIPO_SALIDA`, `ID_CLUSTER`. Debe incluir la hoja `Resumen Bloques`. |
| GRUPO_DESTINOS (.xlsx) | ✅ | Consulta personalizada DXC 9066 — Consulta destinos por zona. |
| Capacidad de rampas (.csv) | ✅ | CSV separado por `;` con columnas `RAMP` y `PALLETS`. |
| Bloques horarios (.xlsx) | ⚠️ | Solo para Gantt 1H y Sorter Map. Columnas: `NUEVO BLOQUE`, `Día/Hora LIBERACIÓN`, `Día/Hora DESACTIVACIÓN`. |
| Superplayas (.xlsx) | ➖ | Opcional. Columnas: `AGRUPACION_PLAYA`, `SUPERPLAYA`. |

> Al subir la parrilla, la herramienta detecta automáticamente las hojas disponibles y deduce el código de semana (p.ej. `S14`).

---

**2 · Elige qué generar**

| Botón | Qué hace | Ficheros necesarios |
|---|---|---|
| Configuración DXC | Genera el `GRUPO_DESTINOS` listo para DXC/MAR + resumen HTML interactivo | Parrilla + GD + Capacidad |
| Gantt 1H | Excel visual rampas × tiempo × bloque. Hojas: `LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO` | + Bloques horarios |
| Sorter Map | Excel con slots físicos, una pestaña por día, posiciones POSTEX coloreadas por bloque | + Bloques horarios |

---

**3 · Descarga los resultados**

Tras generar la Configuración DXC encontrarás el GD completo para subir a DXC/MAR, el fichero solo-especiales con las filas nuevas, y en "Más descargas" los CSV para importación directa (POSTEX/SOREXP), la lista de canceladas y el informe HTML.

---

**Tipos de salida en la parrilla**

| Tipo | Acción |
|---|---|
| `HABITUAL` | Sin cambios — se mantiene en el GD |
| `CANCELADA` | Se elimina del GD esta semana |
| `ESPECIAL DIA CAMBIO` | Se reasigna a rampas libres en el nuevo día |
| `IRREGULAR` | Ignorada — no pasa por el GD semanal |
    """)
st.markdown('<div style="height:4px;"></div>', unsafe_allow_html=True)

# ── File upload section ───────────────────────────────────────────────────────
st.markdown('<div class="mng-section">01 — Ficheros de entrada</div>', unsafe_allow_html=True)
st.markdown('<div class="mng-content">', unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader(
        "Parrilla de salidas ✱",
        type=["xlsx"],
        help="Fichero: parrilla_de_salidas.xlsx — debe incluir la hoja 'Resumen Bloques'",
    )
with col2:
    f_gd = st.file_uploader(
        "GRUPO_DESTINOS ✱",
        type=["xlsx"],
        help="Consulta personalizada DXC 9066 — Consulta destinos por zona",
    )

col3, col4 = st.columns(2)
with col3:
    f_cap = st.file_uploader(
        "Capacidad de rampas ✱",
        type=["csv"],
        help="CSV separado por ';' con columnas RAMP y PALLETS",
    )
with col4:
    f_bloques = st.file_uploader(
        "Bloques horarios",
        type=["xlsx"],
        help="Necesario para Gantt 1H y Sorter Map. Columnas: NUEVO BLOQUE · Día/Hora LIBERACIÓN · Día/Hora DESACTIVACIÓN",
    )

f_superplaya = st.file_uploader(
    "Superplayas (opcional)",
    type=["xlsx"],
    help="Columnas: AGRUPACION_PLAYA · SUPERPLAYA — define destinos que deben ir en rampas contiguas",
)

# ── Sheet autodetect ──────────────────────────────────────────────────────────
sheet  = None
semana = None

if f_parrilla:
    try:
        wb_tmp = load_workbook(io.BytesIO(f_parrilla.read()), read_only=True, data_only=True)
        f_parrilla.seek(0)
        available_sheets = [s for s in wb_tmp.sheetnames if s != "Resumen Bloques"]
        if available_sheets:
            sheet = st.selectbox(
                "Hoja de la parrilla",
                options=available_sheets,
                help="Selecciona la hoja con la parrilla semanal (se excluye 'Resumen Bloques')",
            )
            m = re.search(r'(s\d+)', sheet, re.IGNORECASE)
            semana = m.group(1).upper() if m else sheet.upper()
        else:
            st.warning("No se encontraron hojas de parrilla en el fichero.")
    except Exception as e:
        st.error(f"Error leyendo la parrilla: {e}")
        f_parrilla.seek(0)

# ── Status row ────────────────────────────────────────────────────────────────
if any([f_parrilla, f_gd, f_cap, f_bloques]):
    items = [
        ("Parrilla",       f_parrilla, True),
        ("Grupo destinos", f_gd,       True),
        ("Capacidad",      f_cap,      True),
        ("Bloques",        f_bloques,  False),
    ]
    s = '<div class="mng-status-row">'
    for label, f, _ in items:
        cls = "ready" if f else "missing"
        s += f'<div class="mng-status-item {cls}">{label}</div>'
    s += '</div>'
    st.markdown(s, unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────
ALL_DAYS  = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
base_ok   = bool(f_parrilla and f_gd and f_cap and sheet)
vis_ok    = bool(base_ok and f_bloques)

def save_uploads(tmp: Path):
    p = {}
    for key, f, name in [
        ("parrilla",   f_parrilla,   "parrilla.xlsx"),
        ("gd",         f_gd,         "gd.xlsx"),
        ("cap",        f_cap,         "cap.csv"),
        ("bloques",    f_bloques,     "bloques.xlsx"),
        ("superplaya", f_superplaya,  "superplaya.xlsx"),
    ]:
        if f:
            path = tmp / name
            path.write_bytes(f.read())
            f.seek(0)
            p[key] = path
    return p

def run_gd(p, tmp, sc, days_arg=""):
    gd   = tmp / f"GRUPO_DESTINOS_{sc}.xlsx"
    html = tmp / f"resumen_{sc}.html"
    cmd  = [sys.executable, str(BASE_DIR / "process_parrilla.py"),
            str(p["parrilla"]), str(p["gd"]), str(p["cap"]),
            sheet.strip(), sc, str(gd), str(html), days_arg]
    if "superplaya" in p:
        cmd.append(str(p["superplaya"]))
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    return gd, html, r

def show_log(r, expanded=False):
    lines = [l for l in (r.stdout + r.stderr).splitlines() if l.strip()]
    with st.expander("Ver log de proceso", expanded=expanded or r.returncode != 0):
        for line in lines:
            if "✓" in line:                                            st.success(line)
            elif "❌" in line:                                          st.error(line)
            elif "⚠" in line or "E2" in line or "Sin config" in line:  st.warning(line)
            else:                                                        st.text(line)
    return lines


# ── Action section ────────────────────────────────────────────────────────────
st.markdown('<div class="mng-section">02 — Generar outputs</div>', unsafe_allow_html=True)
st.markdown('<div class="mng-content">', unsafe_allow_html=True)

b1, b2, b3 = st.columns(3)

with b1:
    st.markdown('<div class="mng-action-label">Configuración DXC</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-action-desc">GRUPO_DESTINOS + resumen HTML</div>', unsafe_allow_html=True)
    btn1 = st.button("Generar", key="go1", type="primary",
                     disabled=not base_ok, use_container_width=True)
    if not base_ok:
        st.markdown('<div class="mng-action-note">Requiere parrilla, GD y capacidad</div>', unsafe_allow_html=True)

with b2:
    st.markdown('<div class="mng-action-label">Gantt 1H</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-action-desc">Bloques × rampas × hora</div>', unsafe_allow_html=True)
    btn2 = st.button("Generar", key="go2", type="primary",
                     disabled=not vis_ok, use_container_width=True)
    if not vis_ok:
        st.markdown('<div class="mng-action-note">Requiere también bloques horarios</div>', unsafe_allow_html=True)

with b3:
    st.markdown('<div class="mng-action-label">Sorter Map</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-action-desc">1 pestaña por día, slots físicos</div>', unsafe_allow_html=True)
    btn3 = st.button("Generar", key="go3", type="primary",
                     disabled=not vis_ok, use_container_width=True)
    if not vis_ok:
        st.markdown('<div class="mng-action-note">Requiere también bloques horarios</div>', unsafe_allow_html=True)

st.markdown('</div>', unsafe_allow_html=True)

if btn1:
    for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter",
              "r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
        st.session_state[k] = None
    st.session_state["_run1"] = True

if btn2:
    st.session_state["r2_gantt"] = None
    st.session_state["_run2"] = True

if btn3:
    st.session_state["r3_map"] = None
    st.session_state["_run3"] = True


# ── Execute action 1 ──────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner(f"Procesando parrilla {sc} y asignando rampas…"):
            gd, html, r = run_gd(p, tmp, sc)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
        else:
            esp_path = Path(str(gd).replace('.xlsx', '_SOLO_ESPECIALES.xlsx'))
            can_path = Path(str(gd).replace('.xlsx', '_CANCELADAS.txt'))
            _gd_bytes = gd.read_bytes()
            _px, _sx  = gd_to_dxc_csv(_gd_bytes)
            st.session_state["r1_gd"]          = (gd.name, _gd_bytes)
            st.session_state["r1_esp"]         = (esp_path.name, esp_path.read_bytes()) if esp_path.exists() else None
            st.session_state["r1_postex_csv"]  = (gd.stem + "_POSTEX.csv", _px)
            st.session_state["r1_sorexp_csv"]  = (gd.stem + "_SOREXP.csv", _sx)
            if esp_path.exists():
                _epx, _esx = gd_to_dxc_csv(esp_path.read_bytes())
                st.session_state["r1_esp_postex_csv"] = (esp_path.stem + "_POSTEX.csv", _epx)
                st.session_state["r1_esp_sorexp_csv"] = (esp_path.stem + "_SOREXP.csv", _esx)
            st.session_state["r1_can"]         = (can_path.name, can_path.read_text(encoding='utf-8')) if can_path.exists() else None
            st.session_state["r1_html"]        = (html.name, html.read_bytes()) if html.exists() else None
            st.session_state["r1_day_filter"]  = None

# ── Execute action 2 ──────────────────────────────────────────────────────────
if st.session_state.get("_run2"):
    st.session_state["_run2"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, _, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando GD base.")
            show_log(r0, expanded=True)
        else:
            out = tmp / f"gantt_1h_{sc}.xlsx"
            with st.spinner("Generando Gantt 1H…"):
                r = subprocess.run(
                    [sys.executable, str(BASE_DIR / "gantt_1h.py"),
                     str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"],
                    capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.session_state["r2_gantt"] = (out.name, out.read_bytes())
            else:
                st.error("El Gantt terminó con error.")

# ── Execute action 3 ──────────────────────────────────────────────────────────
if st.session_state.get("_run3"):
    st.session_state["_run3"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, _, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando GD base.")
            show_log(r0, expanded=True)
        else:
            out = tmp / f"sorter_map_{sc}.xlsx"
            with st.spinner("Generando Sorter Map…"):
                cmd = [sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                       str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"]
                if "parrilla" in p and "gd" in p:
                    cmd += [str(p["parrilla"]), sheet.strip(), str(p["gd"])]
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.session_state["r3_map"] = (out.name, out.read_bytes())
            else:
                st.error("El Sorter Map terminó con error.")


# ── Results: action 1 ─────────────────────────────────────────────────────────
if st.session_state["r1_gd"] is not None:
    sc = semana or "SEMANA"
    st.markdown(f'<div class="mng-result-band">Configuración <span class="semana">{sc}</span> — lista para descargar</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-content">', unsafe_allow_html=True)

    if st.session_state["r1_day_filter"]:
        st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")

    st.markdown('<div class="mng-dl-label">Ficheros principales</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        name, data = st.session_state["r1_gd"]
        st.download_button("GD completo (.xlsx)", data=data, file_name=name,
                           mime=XLSX_MIME, use_container_width=True)
        st.caption("Subir a DXC / MAR")
    with c2:
        if st.session_state["r1_esp"]:
            name, data = st.session_state["r1_esp"]
            st.download_button("Solo especiales (.xlsx)", data=data, file_name=name,
                               mime=XLSX_MIME, use_container_width=True)
            st.caption("Solo las filas nuevas a añadir en DXC")

    with st.expander("Más descargas — CSV · Canceladas · Resumen HTML"):
        st.markdown('<div class="mng-dl-label">Importar en DXC (formato CSV)</div>', unsafe_allow_html=True)
        cc1, cc2, cc3, cc4 = st.columns(4)
        for col, key, label in [
            (cc1, "r1_postex_csv",     "POSTEX completo"),
            (cc2, "r1_sorexp_csv",     "SOREXP completo"),
            (cc3, "r1_esp_postex_csv", "POSTEX especiales"),
            (cc4, "r1_esp_sorexp_csv", "SOREXP especiales"),
        ]:
            with col:
                if st.session_state[key]:
                    name, data = st.session_state[key]
                    st.download_button(label, data=data, file_name=name,
                                       mime="text/csv", use_container_width=True)

        st.markdown('<div class="mng-dl-label" style="margin-top:16px;">Otros ficheros</div>', unsafe_allow_html=True)
        d1, d2 = st.columns(2)
        with d1:
            if st.session_state["r1_can"]:
                name, txt = st.session_state["r1_can"]
                st.download_button("Canceladas (.txt)", data=txt, file_name=name,
                                   mime="text/plain", use_container_width=True)
                with st.expander("Ver canceladas"):
                    st.text(txt)
        with d2:
            if st.session_state["r1_html"]:
                name, data = st.session_state["r1_html"]
                st.download_button("Resumen HTML", data=data, file_name=name,
                                   mime="text/html", use_container_width=True)
                st.caption("Informe interactivo con gráfico de capacidad")

    with st.expander("Filtrar por día y regenerar"):
        selected_days = st.multiselect(
            "Días a incluir en el GD filtrado",
            options=ALL_DAYS,
            default=[],
            key="day_filter_sel",
            placeholder="Selecciona uno o más días…",
        )
        if st.button("Regenerar con filtro", key="regen_filter",
                     disabled=not (selected_days and base_ok)):
            days_arg = ",".join(selected_days)
            with tempfile.TemporaryDirectory() as _tmp:
                tmp = Path(_tmp)
                p   = save_uploads(tmp)
                with st.spinner(f"Regenerando para {', '.join(selected_days)}…"):
                    gd, html, r = run_gd(p, tmp, sc, days_arg)
                if r.returncode == 0 and gd.exists():
                    esp_path = Path(str(gd).replace('.xlsx', '_SOLO_ESPECIALES.xlsx'))
                    can_path = Path(str(gd).replace('.xlsx', '_CANCELADAS.txt'))
                    st.session_state["r1_gd"]   = (gd.name,   gd.read_bytes())
                    st.session_state["r1_esp"]  = (esp_path.name, esp_path.read_bytes()) if esp_path.exists() else None
                    st.session_state["r1_can"]  = (can_path.name, can_path.read_text(encoding='utf-8')) if can_path.exists() else None
                    st.session_state["r1_html"] = (html.name, html.read_bytes()) if html.exists() else None
                    st.session_state["r1_day_filter"] = selected_days
                    st.rerun()
                else:
                    st.error("Error en regeneración.")
                    show_log(r, expanded=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ── Results: action 2 ─────────────────────────────────────────────────────────
if st.session_state["r2_gantt"] is not None:
    st.markdown('<div class="mng-result-band">Gantt 1H — listo para descargar</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-content">', unsafe_allow_html=True)
    name, data = st.session_state["r2_gantt"]
    st.download_button("Gantt 1H (.xlsx)", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")
    st.markdown('</div>', unsafe_allow_html=True)

# ── Results: action 3 ─────────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    st.markdown('<div class="mng-result-band">Sorter Map — listo para descargar</div>', unsafe_allow_html=True)
    st.markdown('<div class="mng-content">', unsafe_allow_html=True)
    name, data = st.session_state["r3_map"]
    st.download_button("Sorter Map (.xlsx)", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")
    st.markdown('</div>', unsafe_allow_html=True)

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="mng-footer">
  v0.07 · VDL B2B · Estrictamente confidencial · Solo uso interno
</div>
""", unsafe_allow_html=True)
