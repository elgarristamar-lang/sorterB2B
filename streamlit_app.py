# Version: 0.08 — MANGO corporate style (layout fixed)
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
        if not desc or tipo not in ("POSTEX","SOREXP") or not dest_raw or not elem: continue
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

# Strategy: let Streamlit keep its normal layout (centered, default padding),
# then use CSS to theme every native widget. Custom HTML blocks (header, section
# labels, result bands, footer) are full-bleed via negative margins.

MANGO_CSS = """
<style>
  /* ── Fonts ── */
  html, body, .stApp, * {
    font-family: "Mango New", "Aptos Display", "Trebuchet MS", Arial, sans-serif !important;
  }
  .stApp { background: #f0f0f0 !important; }

  /* ── Main card: white, constrained ── */
  .block-container,
  div[data-testid="stMainBlockContainer"],
  .stMainBlockContainer {
    background: #ffffff !important;
    max-width: 760px !important;
    padding-left: 0 !important;
    padding-right: 0 !important;
    padding-top: 0 !important;
    padding-bottom: 60px !important;
  }

  /* ── Full-bleed custom HTML blocks ── */
  .mng-header {
    background: #000;
    padding: 32px 36px 26px 36px;
    margin-bottom: 0;
  }
  .mng-header .eyebrow {
    font-size: 10px;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: #666;
    margin-bottom: 12px;
  }
  .mng-header h1 {
    font-size: 26px;
    font-weight: 400;
    color: #ffffff;
    margin: 0 0 8px 0;
    line-height: 1.15;
  }
  .mng-header .sub {
    font-size: 12px;
    color: #888;
  }

  .mng-section-label {
    font-size: 10px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: #aaa;
    padding: 24px 36px 8px 36px;
    border-top: 1px solid #ebebeb;
  }

  .mng-result-band {
    background: #000;
    color: #fff;
    padding: 13px 36px;
    font-size: 10px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin-top: 8px;
  }
  .mng-result-band strong {
    font-size: 13px;
    font-weight: 400;
    letter-spacing: 0.04em;
    text-transform: none;
  }

  .mng-dl-label {
    font-size: 10px;
    letter-spacing: 0.1em;
    text-transform: uppercase;
    color: #aaa;
    margin: 0 0 10px 0;
    padding-bottom: 8px;
    border-bottom: 1px solid #ebebeb;
  }

  .mng-status-row {
    display: flex;
    border-top: 1px solid #ebebeb;
    border-bottom: 1px solid #ebebeb;
    margin: 4px 0 0 0;
  }
  .mng-status-item {
    flex: 1;
    padding: 10px 14px;
    border-right: 1px solid #ebebeb;
    font-size: 10px;
    color: #ccc;
    letter-spacing: 0.08em;
    text-transform: uppercase;
  }
  .mng-status-item:last-child { border-right: none; }
  .mng-status-item.ready { color: #1a1a1a; }
  .mng-status-item.ready::before { content: "— "; }
  .mng-status-item.missing::before { content: "· "; color: #e0e0e0; }

  .mng-action-block {
    border: 1px solid #ebebeb;
    padding: 16px 16px 12px 16px;
    height: 100%;
  }
  .mng-action-title {
    font-size: 13px;
    color: #1a1a1a;
    margin-bottom: 3px;
  }
  .mng-action-desc {
    font-size: 11px;
    color: #aaa;
    margin-bottom: 14px;
    line-height: 1.4;
  }
  .mng-action-note {
    font-size: 10px;
    color: #ccc;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    margin-top: 8px;
  }

  .mng-footer {
    background: #000;
    color: #555;
    font-size: 9px;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    text-align: right;
    padding: 14px 36px;
    margin-top: 32px;
  }

  /* ── Streamlit widget overrides ── */

  /* Section padding via stVerticalBlock children */
  div[data-testid="stMainBlockContainer"] > div > div[data-testid="stVerticalBlock"] {
    gap: 0 !important;
  }

  /* File uploader label */
  div[data-testid="stFileUploader"] label p {
    font-size: 10px !important;
    font-weight: 400 !important;
    color: #888 !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
  }
  /* File uploader dropzone */
  div[data-testid="stFileUploaderDropzone"] {
    border: 1px solid #e0e0e0 !important;
    border-radius: 0 !important;
    background: #fafafa !important;
    padding: 10px 14px !important;
  }
  div[data-testid="stFileUploaderDropzone"] p,
  div[data-testid="stFileUploaderDropzone"] small,
  div[data-testid="stFileUploaderDropzone"] span {
    font-size: 12px !important;
    color: #bbb !important;
  }
  div[data-testid="stFileUploaderDropzone"] button {
    border-radius: 0 !important;
    border: 1px solid #1a1a1a !important;
    background: #fff !important;
    color: #1a1a1a !important;
    font-size: 10px !important;
    letter-spacing: 0.08em !important;
    text-transform: uppercase !important;
    padding: 6px 12px !important;
  }
  div[data-testid="stFileUploaderDropzone"] button:hover {
    background: #000 !important;
    color: #fff !important;
  }

  /* Primary buttons */
  div[data-testid="stButton"] > button[kind="primary"],
  div[data-testid="stButton"] > button[data-testid="baseButton-primary"] {
    background: #000 !important;
    color: #fff !important;
    border: none !important;
    border-radius: 0 !important;
    font-size: 10px !important;
    letter-spacing: 0.12em !important;
    text-transform: uppercase !important;
    padding: 11px 16px !important;
    width: 100% !important;
  }
  div[data-testid="stButton"] > button[kind="primary"]:hover { background: #222 !important; }
  div[data-testid="stButton"] > button[kind="primary"]:disabled {
    background: #e8e8e8 !important;
    color: #bbb !important;
  }

  /* Secondary / non-primary buttons */
  div[data-testid="stButton"] > button:not([kind="primary"]) {
    background: #fff !important;
    color: #1a1a1a !important;
    border: 1px solid #1a1a1a !important;
    border-radius: 0 !important;
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    padding: 10px 16px !important;
    width: 100% !important;
  }
  div[data-testid="stButton"] > button:not([kind="primary"]):hover {
    background: #000 !important; color: #fff !important;
  }

  /* Download buttons */
  div[data-testid="stDownloadButton"] > button {
    background: #fff !important;
    color: #1a1a1a !important;
    border: 1px solid #1a1a1a !important;
    border-radius: 0 !important;
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    padding: 10px 16px !important;
    width: 100% !important;
  }
  div[data-testid="stDownloadButton"] > button:hover {
    background: #000 !important; color: #fff !important; border-color: #000 !important;
  }

  /* Selectbox */
  div[data-testid="stSelectbox"] > div > div {
    border-radius: 0 !important;
    border: 1px solid #d0d0d0 !important;
    font-size: 13px !important;
  }
  div[data-testid="stSelectbox"] label p {
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    color: #888 !important;
  }

  /* Multiselect */
  div[data-testid="stMultiSelect"] > div > div {
    border-radius: 0 !important;
    border: 1px solid #d0d0d0 !important;
  }
  div[data-testid="stMultiSelect"] label p {
    font-size: 10px !important;
    text-transform: uppercase !important;
    letter-spacing: 0.1em !important;
    color: #888 !important;
  }
  span[data-baseweb="tag"] {
    border-radius: 0 !important;
    background: #000 !important;
    color: #fff !important;
    font-size: 10px !important;
  }

  /* Expander */
  div[data-testid="stExpander"] {
    border: 1px solid #ebebeb !important;
    border-radius: 0 !important;
  }
  div[data-testid="stExpander"] summary {
    font-size: 10px !important;
    letter-spacing: 0.1em !important;
    text-transform: uppercase !important;
    color: #888 !important;
    font-weight: 400 !important;
    padding: 14px 16px !important;
  }
  div[data-testid="stExpander"] summary:hover { background: #f8f8f8 !important; }

  /* Alerts */
  div[data-testid="stAlert"] { border-radius: 0 !important; border-left-width: 3px !important; }
  div[data-testid="stAlert"][data-type="success"] {
    background: #f8f8f8 !important; border-color: #1a1a1a !important;
  }
  div[data-testid="stAlert"][data-type="info"] {
    background: #fafafa !important; border-color: #aaa !important;
  }

  /* Caption */
  div[data-testid="stCaptionContainer"] p, small, .stCaption {
    font-size: 10px !important; color: #aaa !important;
  }

  /* Divider */
  hr { border: none !important; border-top: 1px solid #ebebeb !important; margin: 0 !important; }

  /* Hide Streamlit chrome */
  #MainMenu, footer, header, .stDeployButton { visibility: hidden !important; display: none !important; }
</style>
"""
st.markdown(MANGO_CSS, unsafe_allow_html=True)


# ── Helper: padded section wrapper using st.container ────────────────────────
def section_pad():
    """Returns a container — content inside gets padding via a wrapping div trick."""
    return st.container()

def html(s):
    st.markdown(s, unsafe_allow_html=True)

PAD = "padding: 0 36px;"


# ── Header (full-bleed black) ─────────────────────────────────────────────────
html("""
<div class="mng-header">
  <div class="eyebrow">VDL B2B · Logística</div>
  <h1>Sorter VDL B2B</h1>
  <div class="sub">Configurador de semanas especiales</div>
</div>
""")

# ── Session state ─────────────────────────────────────────────────────────────
_STATE_KEYS = ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map",
               "r1_day_filter","r1_postex_csv","r1_sorexp_csv",
               "r1_esp_postex_csv","r1_esp_sorexp_csv"]
for k in _STATE_KEYS:
    if k not in st.session_state:
        st.session_state[k] = None


# ── Usage guide ───────────────────────────────────────────────────────────────
with st.container():
    html(f'<div style="{PAD} padding-top:20px; padding-bottom:4px;">')
    with st.expander("Cómo usar esta herramienta"):
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

---

**2 · Elige qué generar**

| Botón | Qué hace | Ficheros necesarios |
|---|---|---|
| Configuración DXC | `GRUPO_DESTINOS` listo para DXC/MAR + resumen HTML | Parrilla + GD + Capacidad |
| Gantt 1H | Excel visual rampas × tiempo × bloque | + Bloques horarios |
| Sorter Map | Excel slots físicos, una pestaña por día | + Bloques horarios |

---

**Tipos de salida en la parrilla**

| Tipo | Acción |
|---|---|
| `HABITUAL` | Sin cambios — se mantiene en el GD |
| `CANCELADA` | Se elimina del GD esta semana |
| `ESPECIAL DIA CAMBIO` | Se reasigna a rampas libres en el nuevo día |
| `IRREGULAR` | Ignorada — no pasa por el GD semanal |
        """)
    html('</div>')


# ── 01 · Ficheros de entrada ──────────────────────────────────────────────────
html('<div class="mng-section-label">01 — Ficheros de entrada</div>')

with st.container():
    html(f'<div style="{PAD}">')

    col1, col2 = st.columns(2)
    with col1:
        f_parrilla = st.file_uploader(
            "Parrilla de salidas ✱", type=["xlsx"],
            help="Fichero: parrilla_de_salidas.xlsx — debe incluir la hoja 'Resumen Bloques'")
    with col2:
        f_gd = st.file_uploader(
            "GRUPO_DESTINOS ✱", type=["xlsx"],
            help="Consulta personalizada DXC 9066 — Consulta destinos por zona")

    col3, col4 = st.columns(2)
    with col3:
        f_cap = st.file_uploader(
            "Capacidad de rampas ✱", type=["csv"],
            help="CSV separado por ';' con columnas RAMP y PALLETS")
    with col4:
        f_bloques = st.file_uploader(
            "Bloques horarios", type=["xlsx"],
            help="Para Gantt 1H y Sorter Map. Columnas: NUEVO BLOQUE · Día/Hora LIBERACIÓN · Día/Hora DESACTIVACIÓN")

    f_superplaya = st.file_uploader(
        "Superplayas (opcional)", type=["xlsx"],
        help="Columnas: AGRUPACION_PLAYA · SUPERPLAYA")

    html('</div>')

# ── Sheet autodetect ──────────────────────────────────────────────────────────
sheet = None
semana = None

if f_parrilla:
    try:
        wb_tmp = load_workbook(io.BytesIO(f_parrilla.read()), read_only=True, data_only=True)
        f_parrilla.seek(0)
        available_sheets = [s for s in wb_tmp.sheetnames if s != "Resumen Bloques"]
        if available_sheets:
            with st.container():
                html(f'<div style="{PAD}">')
                sheet = st.selectbox(
                    "Hoja de la parrilla",
                    options=available_sheets,
                    help="Selecciona la hoja con la parrilla semanal (se excluye 'Resumen Bloques')")
                html('</div>')
            m = re.search(r'(s\d+)', sheet, re.IGNORECASE)
            semana = m.group(1).upper() if m else sheet.upper()
        else:
            st.warning("No se encontraron hojas de parrilla en el fichero.")
    except Exception as e:
        st.error(f"Error leyendo la parrilla: {e}")
        f_parrilla.seek(0)

# ── Status row ────────────────────────────────────────────────────────────────
if any([f_parrilla, f_gd, f_cap, f_bloques]):
    items = [("Parrilla", f_parrilla), ("Grupo destinos", f_gd),
             ("Capacidad", f_cap), ("Bloques", f_bloques)]
    s = '<div class="mng-status-row" style="margin: 0 36px;">'
    for label, f in items:
        cls = "ready" if f else "missing"
        s += f'<div class="mng-status-item {cls}">{label}</div>'
    s += '</div>'
    html(s)


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
            path = tmp / name; path.write_bytes(f.read()); f.seek(0); p[key] = path
    return p

def run_gd(p, tmp, sc, days_arg=""):
    gd   = tmp / f"GRUPO_DESTINOS_{sc}.xlsx"
    ht   = tmp / f"resumen_{sc}.html"
    cmd  = [sys.executable, str(BASE_DIR/"process_parrilla.py"),
            str(p["parrilla"]), str(p["gd"]), str(p["cap"]),
            sheet.strip(), sc, str(gd), str(ht), days_arg]
    if "superplaya" in p: cmd.append(str(p["superplaya"]))
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    return gd, ht, r

def show_log(r, expanded=False):
    lines = [l for l in (r.stdout + r.stderr).splitlines() if l.strip()]
    with st.expander("Ver log de proceso", expanded=expanded or r.returncode != 0):
        for line in lines:
            if   "✓" in line:                                         st.success(line)
            elif "❌" in line:                                         st.error(line)
            elif "⚠" in line or "E2" in line or "Sin config" in line: st.warning(line)
            else:                                                       st.text(line)


# ── 02 · Generar outputs ──────────────────────────────────────────────────────
html('<div class="mng-section-label">02 — Generar outputs</div>')

with st.container():
    html(f'<div style="{PAD}">')
    b1, b2, b3 = st.columns(3)

    with b1:
        html('<div class="mng-action-block">')
        html('<div class="mng-action-title">Configuración DXC</div>')
        html('<div class="mng-action-desc">GRUPO_DESTINOS + resumen HTML</div>')
        btn1 = st.button("Generar", key="go1", type="primary", disabled=not base_ok, use_container_width=True)
        if not base_ok: html('<div class="mng-action-note">Requiere parrilla, GD y capacidad</div>')
        html('</div>')

    with b2:
        html('<div class="mng-action-block">')
        html('<div class="mng-action-title">Gantt 1H</div>')
        html('<div class="mng-action-desc">Bloques × rampas × hora</div>')
        btn2 = st.button("Generar", key="go2", type="primary", disabled=not vis_ok, use_container_width=True)
        if not vis_ok: html('<div class="mng-action-note">Requiere bloques horarios</div>')
        html('</div>')

    with b3:
        html('<div class="mng-action-block">')
        html('<div class="mng-action-title">Sorter Map</div>')
        html('<div class="mng-action-desc">1 pestaña por día, slots físicos</div>')
        btn3 = st.button("Generar", key="go3", type="primary", disabled=not vis_ok, use_container_width=True)
        if not vis_ok: html('<div class="mng-action-note">Requiere bloques horarios</div>')
        html('</div>')

    html('</div>')

if btn1:
    for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter",
              "r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
        st.session_state[k] = None
    st.session_state["_run1"] = True
if btn2:
    st.session_state["r2_gantt"] = None; st.session_state["_run2"] = True
if btn3:
    st.session_state["r3_map"]   = None; st.session_state["_run3"] = True


# ── Execute 1 ─────────────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp); p = save_uploads(tmp)
        with st.spinner(f"Procesando parrilla {sc}…"): gd, ht, r = run_gd(p, tmp, sc)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
        else:
            esp  = Path(str(gd).replace('.xlsx','_SOLO_ESPECIALES.xlsx'))
            can  = Path(str(gd).replace('.xlsx','_CANCELADAS.txt'))
            gdb  = gd.read_bytes()
            px, sx = gd_to_dxc_csv(gdb)
            st.session_state["r1_gd"]         = (gd.name, gdb)
            st.session_state["r1_esp"]        = (esp.name, esp.read_bytes()) if esp.exists() else None
            st.session_state["r1_postex_csv"] = (gd.stem+"_POSTEX.csv", px)
            st.session_state["r1_sorexp_csv"] = (gd.stem+"_SOREXP.csv", sx)
            if esp.exists():
                epx, esx = gd_to_dxc_csv(esp.read_bytes())
                st.session_state["r1_esp_postex_csv"] = (esp.stem+"_POSTEX.csv", epx)
                st.session_state["r1_esp_sorexp_csv"] = (esp.stem+"_SOREXP.csv", esx)
            st.session_state["r1_can"]  = (can.name, can.read_text(encoding='utf-8')) if can.exists() else None
            st.session_state["r1_html"] = (ht.name, ht.read_bytes()) if ht.exists() else None
            st.session_state["r1_day_filter"] = None

# ── Execute 2 ─────────────────────────────────────────────────────────────────
if st.session_state.get("_run2"):
    st.session_state["_run2"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp); p = save_uploads(tmp)
        with st.spinner("Generando GD base…"): gd, _, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando GD base."); show_log(r0, expanded=True)
        else:
            out = tmp / f"gantt_1h_{sc}.xlsx"
            with st.spinner("Generando Gantt 1H…"):
                r = subprocess.run([sys.executable, str(BASE_DIR/"gantt_1h.py"),
                    str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"],
                    capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists(): st.session_state["r2_gantt"] = (out.name, out.read_bytes())
            else: st.error("El Gantt terminó con error.")

# ── Execute 3 ─────────────────────────────────────────────────────────────────
if st.session_state.get("_run3"):
    st.session_state["_run3"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp); p = save_uploads(tmp)
        with st.spinner("Generando GD base…"): gd, _, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando GD base."); show_log(r0, expanded=True)
        else:
            out = tmp / f"sorter_map_{sc}.xlsx"
            with st.spinner("Generando Sorter Map…"):
                cmd = [sys.executable, str(BASE_DIR/"sorter_map_por_dia.py"),
                       str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"]
                if "parrilla" in p and "gd" in p:
                    cmd += [str(p["parrilla"]), sheet.strip(), str(p["gd"])]
                r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists(): st.session_state["r3_map"] = (out.name, out.read_bytes())
            else: st.error("El Sorter Map terminó con error.")


# ── Results 1 ─────────────────────────────────────────────────────────────────
if st.session_state["r1_gd"] is not None:
    sc = semana or "SEMANA"
    html(f'<div class="mng-result-band">Configuración <strong>{sc}</strong> — lista para descargar</div>')

    with st.container():
        html(f'<div style="{PAD} padding-top:20px;">')

        if st.session_state["r1_day_filter"]:
            st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")

        html('<div class="mng-dl-label">Ficheros principales</div>')
        c1, c2 = st.columns(2)
        with c1:
            n, d = st.session_state["r1_gd"]
            st.download_button("GD completo (.xlsx)", data=d, file_name=n, mime=XLSX_MIME, use_container_width=True)
            st.caption("Subir a DXC / MAR")
        with c2:
            if st.session_state["r1_esp"]:
                n, d = st.session_state["r1_esp"]
                st.download_button("Solo especiales (.xlsx)", data=d, file_name=n, mime=XLSX_MIME, use_container_width=True)
                st.caption("Solo las filas nuevas a añadir en DXC")

        with st.expander("Más descargas — CSV · Canceladas · Resumen HTML"):
            html('<div class="mng-dl-label">Importar en DXC (formato CSV)</div>')
            cc1,cc2,cc3,cc4 = st.columns(4)
            for col, key, lbl in [(cc1,"r1_postex_csv","POSTEX completo"),
                                   (cc2,"r1_sorexp_csv","SOREXP completo"),
                                   (cc3,"r1_esp_postex_csv","POSTEX esp."),
                                   (cc4,"r1_esp_sorexp_csv","SOREXP esp.")]:
                with col:
                    if st.session_state[key]:
                        n,d = st.session_state[key]
                        st.download_button(lbl, data=d, file_name=n, mime="text/csv", use_container_width=True)

            html('<div class="mng-dl-label" style="margin-top:16px;">Otros ficheros</div>')
            d1, d2 = st.columns(2)
            with d1:
                if st.session_state["r1_can"]:
                    n, t = st.session_state["r1_can"]
                    st.download_button("Canceladas (.txt)", data=t, file_name=n, mime="text/plain", use_container_width=True)
                    with st.expander("Ver canceladas"): st.text(t)
            with d2:
                if st.session_state["r1_html"]:
                    n, d = st.session_state["r1_html"]
                    st.download_button("Resumen HTML", data=d, file_name=n, mime="text/html", use_container_width=True)
                    st.caption("Informe interactivo con gráfico de capacidad")

        with st.expander("Filtrar por día y regenerar"):
            sel = st.multiselect("Días a incluir", options=ALL_DAYS, default=[], key="day_filter_sel",
                                 placeholder="Selecciona uno o más días…")
            if st.button("Regenerar con filtro", key="regen_filter", disabled=not (sel and base_ok)):
                with tempfile.TemporaryDirectory() as _tmp:
                    tmp = Path(_tmp); p = save_uploads(tmp)
                    with st.spinner(f"Regenerando para {', '.join(sel)}…"):
                        gd, ht, r = run_gd(p, tmp, sc, ",".join(sel))
                    if r.returncode == 0 and gd.exists():
                        esp = Path(str(gd).replace('.xlsx','_SOLO_ESPECIALES.xlsx'))
                        can = Path(str(gd).replace('.xlsx','_CANCELADAS.txt'))
                        st.session_state["r1_gd"]   = (gd.name, gd.read_bytes())
                        st.session_state["r1_esp"]  = (esp.name, esp.read_bytes()) if esp.exists() else None
                        st.session_state["r1_can"]  = (can.name, can.read_text(encoding='utf-8')) if can.exists() else None
                        st.session_state["r1_html"] = (ht.name, ht.read_bytes()) if ht.exists() else None
                        st.session_state["r1_day_filter"] = sel
                        st.rerun()
                    else:
                        st.error("Error en regeneración."); show_log(r, expanded=True)

        html('</div>')

# ── Results 2 ─────────────────────────────────────────────────────────────────
if st.session_state["r2_gantt"] is not None:
    html('<div class="mng-result-band">Gantt 1H — listo para descargar</div>')
    with st.container():
        html(f'<div style="{PAD} padding-top:16px;">')
        n, d = st.session_state["r2_gantt"]
        st.download_button("Gantt 1H (.xlsx)", data=d, file_name=n, mime=XLSX_MIME, use_container_width=True)
        st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")
        html('</div>')

# ── Results 3 ─────────────────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    html('<div class="mng-result-band">Sorter Map — listo para descargar</div>')
    with st.container():
        html(f'<div style="{PAD} padding-top:16px;">')
        n, d = st.session_state["r3_map"]
        st.download_button("Sorter Map (.xlsx)", data=d, file_name=n, mime=XLSX_MIME, use_container_width=True)
        st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")
        html('</div>')

# ── Footer ────────────────────────────────────────────────────────────────────
html('<div class="mng-footer">v0.08 · VDL B2B · Estrictamente confidencial · Solo uso interno</div>')
