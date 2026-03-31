# Version: 0.09 — MANGO style, pure CSS approach (no HTML div wrappers)
import streamlit as st
import subprocess, sys, tempfile, io, re
from pathlib import Path

try:
    from openpyxl import load_workbook
except ImportError:
    load_workbook = None

BASE_DIR = Path(__file__).parent

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


st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")

st.markdown("""
<style>
/* === FONTS === */
* { font-family: "Mango New","Aptos Display","Trebuchet MS",Arial,sans-serif !important; }

/* === PAGE BACKGROUND === */
.stApp { background: #f0f0f0 !important; }

/* === WHITE CARD === */
[data-testid="stAppViewBlockContainer"],
.stMainBlockContainer,
.block-container {
    background: #fff !important;
    max-width: 800px !important;
    padding: 0 !important;
}

/* === KILL ALL DEFAULT GAPS === */
[data-testid="stVerticalBlock"] { gap: 0 !important; }
[data-testid="stVerticalBlockBorderWrapper"] { padding: 0 !important; }
.element-container { margin: 0 !important; }

/* === SECTION PADDING: inject via the stMarkdown that holds our banners === */
/* All widgets get left/right padding through a universal rule */
[data-testid="stVerticalBlock"] > [data-testid="stVerticalBlockBorderWrapper"] > [data-testid="stVerticalBlock"],
[data-testid="stVerticalBlock"] > div > [data-testid="stHorizontalBlock"],
[data-testid="stVerticalBlock"] > div > [data-testid="stFileUploader"],
[data-testid="stVerticalBlock"] > div > [data-testid="stSelectbox"],
[data-testid="stVerticalBlock"] > div > [data-testid="stExpander"],
[data-testid="stVerticalBlock"] > div > [data-testid="stButton"],
[data-testid="stVerticalBlock"] > div > [data-testid="stDownloadButton"],
[data-testid="stVerticalBlock"] > div > [data-testid="stMultiSelect"],
[data-testid="stVerticalBlock"] > div > [data-testid="stAlert"],
[data-testid="stVerticalBlock"] > div > [data-testid="stCaptionContainer"],
[data-testid="stVerticalBlock"] > div > [data-testid="stMarkdownContainer"] {
    padding-left: 36px !important;
    padding-right: 36px !important;
}

/* === HEADER BANNER (full bleed) === */
.mng-hdr {
    background: #000; color: #fff;
    padding: 28px 36px 24px;
    margin-bottom: 0;
}
.mng-hdr .ey { font-size:10px; letter-spacing:.16em; text-transform:uppercase; color:#666; margin-bottom:10px; }
.mng-hdr h1  { font-size:26px; font-weight:400; color:#fff; margin:0 0 6px; line-height:1.15; }
.mng-hdr .sb { font-size:12px; color:#888; }

/* === SECTION LABEL === */
.mng-sec {
    font-size:10px; letter-spacing:.14em; text-transform:uppercase; color:#aaa;
    padding: 24px 36px 10px;
    border-top: 1px solid #ebebeb;
    margin: 0;
}

/* === RESULT BAND (full bleed black) === */
.mng-res {
    background:#000; color:#fff;
    padding: 13px 36px;
    font-size:10px; letter-spacing:.12em; text-transform:uppercase;
}
.mng-res b { font-size:13px; font-weight:400; letter-spacing:.03em; text-transform:none; }

/* === STATUS ROW === */
.mng-sr { display:flex; border-top:1px solid #ebebeb; border-bottom:1px solid #ebebeb; }
.mng-si {
    flex:1; padding:10px 14px; border-right:1px solid #ebebeb;
    font-size:10px; color:#ccc; letter-spacing:.08em; text-transform:uppercase;
}
.mng-si:last-child { border-right:none; }
.mng-si.on  { color:#1a1a1a; }
.mng-si.on::before  { content:"— "; }
.mng-si.off::before { content:"· "; color:#e0e0e0; }

/* === ACTION CARDS === */
.mng-card { border:1px solid #ebebeb; padding:16px 16px 12px; margin-bottom:4px; }
.mng-ct   { font-size:13px; color:#1a1a1a; margin-bottom:4px; }
.mng-cd   { font-size:11px; color:#aaa; margin-bottom:12px; }
.mng-cn   { font-size:10px; color:#ccc; letter-spacing:.06em; text-transform:uppercase; margin-top:8px; }

/* === DL LABEL === */
.mng-dl {
    font-size:10px; letter-spacing:.1em; text-transform:uppercase; color:#aaa;
    border-bottom:1px solid #ebebeb; padding-bottom:8px; margin:16px 0 10px;
}

/* === FOOTER === */
.mng-ft {
    background:#000; color:#555;
    font-size:9px; letter-spacing:.14em; text-transform:uppercase;
    text-align:right; padding:14px 36px; margin-top:32px;
}

/* === WIDGET STYLING === */

/* File uploader label */
[data-testid="stFileUploader"] label p {
    font-size:10px !important; font-weight:400 !important;
    color:#888 !important; letter-spacing:.1em !important; text-transform:uppercase !important;
}
/* Dropzone */
[data-testid="stFileUploaderDropzone"] {
    border:1px solid #e0e0e0 !important; border-radius:0 !important; background:#fafafa !important;
}
[data-testid="stFileUploaderDropzone"] p,
[data-testid="stFileUploaderDropzone"] small { font-size:12px !important; color:#bbb !important; }
[data-testid="stFileUploaderDropzone"] button {
    border-radius:0 !important; border:1px solid #1a1a1a !important;
    background:#fff !important; color:#1a1a1a !important;
    font-size:10px !important; letter-spacing:.08em !important; text-transform:uppercase !important;
}
[data-testid="stFileUploaderDropzone"] button:hover { background:#000 !important; color:#fff !important; }

/* Primary buttons */
[data-testid="stButton"] > button[kind="primary"] {
    background:#000 !important; color:#fff !important;
    border:none !important; border-radius:0 !important;
    font-size:10px !important; letter-spacing:.12em !important; text-transform:uppercase !important;
    padding:11px 16px !important; width:100% !important;
}
[data-testid="stButton"] > button[kind="primary"]:hover    { background:#222 !important; }
[data-testid="stButton"] > button[kind="primary"]:disabled { background:#e8e8e8 !important; color:#bbb !important; }

/* Secondary buttons */
[data-testid="stButton"] > button:not([kind="primary"]) {
    background:#fff !important; color:#1a1a1a !important;
    border:1px solid #1a1a1a !important; border-radius:0 !important;
    font-size:10px !important; letter-spacing:.1em !important; text-transform:uppercase !important;
    padding:10px 16px !important; width:100% !important;
}
[data-testid="stButton"] > button:not([kind="primary"]):hover { background:#000 !important; color:#fff !important; }

/* Download buttons */
[data-testid="stDownloadButton"] > button {
    background:#fff !important; color:#1a1a1a !important;
    border:1px solid #1a1a1a !important; border-radius:0 !important;
    font-size:10px !important; letter-spacing:.1em !important; text-transform:uppercase !important;
    padding:10px 16px !important; width:100% !important;
}
[data-testid="stDownloadButton"] > button:hover { background:#000 !important; color:#fff !important; border-color:#000 !important; }

/* Selectbox */
[data-testid="stSelectbox"] > div > div { border-radius:0 !important; border:1px solid #d0d0d0 !important; }
[data-testid="stSelectbox"] label p {
    font-size:10px !important; text-transform:uppercase !important;
    letter-spacing:.1em !important; color:#888 !important;
}

/* Multiselect */
[data-testid="stMultiSelect"] > div > div { border-radius:0 !important; border:1px solid #d0d0d0 !important; }
[data-testid="stMultiSelect"] label p {
    font-size:10px !important; text-transform:uppercase !important;
    letter-spacing:.1em !important; color:#888 !important;
}
span[data-baseweb="tag"] { border-radius:0 !important; background:#000 !important; color:#fff !important; font-size:10px !important; }

/* Expander */
[data-testid="stExpander"] { border:1px solid #ebebeb !important; border-radius:0 !important; }
[data-testid="stExpander"] summary {
    font-size:10px !important; letter-spacing:.1em !important;
    text-transform:uppercase !important; color:#888 !important;
    font-weight:400 !important; padding:14px 16px !important;
}
[data-testid="stExpander"] summary:hover { background:#f8f8f8 !important; }

/* Alerts */
[data-testid="stAlert"] { border-radius:0 !important; border-left-width:3px !important; }
[data-testid="stAlert"][data-baseweb="notification"] { background:#f8f8f8 !important; }

/* Caption */
[data-testid="stCaptionContainer"] p { font-size:10px !important; color:#aaa !important; }

/* Widget labels */
label[data-testid="stWidgetLabel"] p {
    font-size:10px !important; text-transform:uppercase !important;
    letter-spacing:.08em !important; color:#888 !important;
}

/* Hide Streamlit chrome */
#MainMenu, footer, header, .stDeployButton { display:none !important; }
</style>
""", unsafe_allow_html=True)

def H(s): st.markdown(s, unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
H("""<div class="mng-hdr">
  <div class="ey">VDL B2B · Logística</div>
  <h1>Sorter VDL B2B</h1>
  <div class="sb">Configurador de semanas especiales</div>
</div>""")

# ── Session state ─────────────────────────────────────────────────────────────
for k in ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map",
          "r1_day_filter","r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
    if k not in st.session_state: st.session_state[k] = None

# ── Guide ─────────────────────────────────────────────────────────────────────
with st.expander("Cómo usar esta herramienta"):
    st.markdown("""
Esta herramienta genera la configuración del sorter VDL B2B para semanas con salidas canceladas o que cambian de día.

**1 · Ficheros de entrada**

| Fichero | Oblig. | Descripción |
|---|:---:|---|
| `parrilla_de_salidas.xlsx` | ✅ | Columnas: `PLAYA`, `DIA_SALIDA`, `DIA_SALIDA_NEW`, `TIPO_SALIDA`, `ID_CLUSTER`. Incluir hoja `Resumen Bloques`. |
| GRUPO_DESTINOS `.xlsx` | ✅ | Consulta DXC 9066 — Consulta destinos por zona |
| `ramp_capacity.csv` | ✅ | CSV `;` con columnas `RAMP` y `PALLETS` |
| Bloques horarios `.xlsx` | ⚠️ | Solo para Gantt y Sorter Map |
| Superplayas `.xlsx` | ➖ | Opcional. `AGRUPACION_PLAYA` · `SUPERPLAYA` |

**2 · Outputs**

| | Qué genera | Necesita |
|---|---|---|
| ⚙️ Config DXC | `GRUPO_DESTINOS` + resumen HTML | Parrilla + GD + Capacidad |
| 📊 Gantt 1H | Excel rampas × tiempo × bloque | + Bloques |
| 🗺 Sorter Map | Excel slots físicos por día | + Bloques |

**Tipos de salida:** `HABITUAL` → sin cambios · `CANCELADA` → eliminar · `ESPECIAL DIA CAMBIO` → reasignar · `IRREGULAR` → ignorar
    """)

# ── 01 Ficheros ───────────────────────────────────────────────────────────────
H('<div class="mng-sec">01 — Ficheros de entrada</div>')

col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader("Parrilla de salidas ✱", type=["xlsx"],
        help="parrilla_de_salidas.xlsx — debe incluir la hoja 'Resumen Bloques'")
with col2:
    f_gd = st.file_uploader("GRUPO_DESTINOS ✱", type=["xlsx"],
        help="Consulta personalizada DXC 9066 — Consulta destinos por zona")

col3, col4 = st.columns(2)
with col3:
    f_cap = st.file_uploader("Capacidad de rampas ✱", type=["csv"],
        help="CSV separado por ';' · columnas RAMP y PALLETS")
with col4:
    f_bloques = st.file_uploader("Bloques horarios", type=["xlsx"],
        help="Para Gantt 1H y Sorter Map · NUEVO BLOQUE · Día/Hora LIBERACIÓN · Día/Hora DESACTIVACIÓN")

f_superplaya = st.file_uploader("Superplayas (opcional)", type=["xlsx"],
    help="Columnas: AGRUPACION_PLAYA · SUPERPLAYA")

# ── Sheet autodetect ──────────────────────────────────────────────────────────
sheet = None; semana = None
if f_parrilla:
    try:
        wb_tmp = load_workbook(io.BytesIO(f_parrilla.read()), read_only=True, data_only=True)
        f_parrilla.seek(0)
        sheets = [s for s in wb_tmp.sheetnames if s != "Resumen Bloques"]
        if sheets:
            sheet = st.selectbox("Hoja de la parrilla", options=sheets,
                help="Hoja con la parrilla semanal (excluye 'Resumen Bloques')")
            m = re.search(r'(s\d+)', sheet, re.IGNORECASE)
            semana = m.group(1).upper() if m else sheet.upper()
        else:
            st.warning("No se encontraron hojas de parrilla.")
    except Exception as e:
        st.error(f"Error leyendo la parrilla: {e}"); f_parrilla.seek(0)

# Status row
if any([f_parrilla, f_gd, f_cap, f_bloques]):
    items = [("Parrilla", f_parrilla), ("Grupo Destinos", f_gd), ("Capacidad", f_cap), ("Bloques", f_bloques)]
    H('<div class="mng-sr">' +
      "".join(f'<div class="mng-si {"on" if f else "off"}">{l}</div>' for l, f in items) +
      '</div>')

# ── Helpers ───────────────────────────────────────────────────────────────────
ALL_DAYS  = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
base_ok   = bool(f_parrilla and f_gd and f_cap and sheet)
vis_ok    = bool(base_ok and f_bloques)

def save_uploads(tmp):
    p = {}
    for key, f, name in [("parrilla",f_parrilla,"parrilla.xlsx"),("gd",f_gd,"gd.xlsx"),
                          ("cap",f_cap,"cap.csv"),("bloques",f_bloques,"bloques.xlsx"),
                          ("superplaya",f_superplaya,"superplaya.xlsx")]:
        if f:
            path = tmp/name; path.write_bytes(f.read()); f.seek(0); p[key]=path
    return p

def run_gd(p, tmp, sc, days_arg=""):
    gd = tmp/f"GRUPO_DESTINOS_{sc}.xlsx"; ht = tmp/f"resumen_{sc}.html"
    cmd = [sys.executable, str(BASE_DIR/"process_parrilla.py"),
           str(p["parrilla"]),str(p["gd"]),str(p["cap"]),
           sheet.strip(),sc,str(gd),str(ht),days_arg]
    if "superplaya" in p: cmd.append(str(p["superplaya"]))
    return gd, ht, subprocess.run(cmd, capture_output=True, text=True, timeout=180)

def show_log(r, expanded=False):
    with st.expander("Ver log", expanded=expanded or r.returncode!=0):
        for line in (r.stdout+r.stderr).splitlines():
            if not line.strip(): continue
            if "✓" in line: st.success(line)
            elif "❌" in line: st.error(line)
            elif "⚠" in line or "E2" in line: st.warning(line)
            else: st.text(line)

# ── 02 Outputs ────────────────────────────────────────────────────────────────
H('<div class="mng-sec">02 — Generar outputs</div>')

c1, c2, c3 = st.columns(3)
with c1:
    H('<div class="mng-card"><div class="mng-ct">Configuración DXC</div><div class="mng-cd">GRUPO_DESTINOS + resumen HTML</div></div>')
    btn1 = st.button("Generar", key="go1", type="primary", disabled=not base_ok, use_container_width=True)
    if not base_ok: H('<div class="mng-cn">Requiere parrilla, GD y capacidad</div>')
with c2:
    H('<div class="mng-card"><div class="mng-ct">Gantt 1H</div><div class="mng-cd">Bloques × rampas × hora</div></div>')
    btn2 = st.button("Generar", key="go2", type="primary", disabled=not vis_ok, use_container_width=True)
    if not vis_ok: H('<div class="mng-cn">Requiere bloques horarios</div>')
with c3:
    H('<div class="mng-card"><div class="mng-ct">Sorter Map</div><div class="mng-cd">1 pestaña por día, slots físicos</div></div>')
    btn3 = st.button("Generar", key="go3", type="primary", disabled=not vis_ok, use_container_width=True)
    if not vis_ok: H('<div class="mng-cn">Requiere bloques horarios</div>')

if btn1:
    for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter",
              "r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
        st.session_state[k] = None
    st.session_state["_run1"] = True
if btn2: st.session_state["r2_gantt"]=None; st.session_state["_run2"]=True
if btn3: st.session_state["r3_map"]=None;   st.session_state["_run3"]=True

# ── Run 1 ─────────────────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _t:
        tmp=Path(_t); p=save_uploads(tmp)
        with st.spinner(f"Procesando {sc}…"): gd,ht,r=run_gd(p,tmp,sc)
        show_log(r)
        if r.returncode!=0 or not gd.exists(): st.error("Error en el proceso.")
        else:
            esp=Path(str(gd).replace('.xlsx','_SOLO_ESPECIALES.xlsx'))
            can=Path(str(gd).replace('.xlsx','_CANCELADAS.txt'))
            gdb=gd.read_bytes(); px,sx=gd_to_dxc_csv(gdb)
            st.session_state["r1_gd"]=(gd.name,gdb)
            st.session_state["r1_esp"]=(esp.name,esp.read_bytes()) if esp.exists() else None
            st.session_state["r1_postex_csv"]=(gd.stem+"_POSTEX.csv",px)
            st.session_state["r1_sorexp_csv"]=(gd.stem+"_SOREXP.csv",sx)
            if esp.exists():
                epx,esx=gd_to_dxc_csv(esp.read_bytes())
                st.session_state["r1_esp_postex_csv"]=(esp.stem+"_POSTEX.csv",epx)
                st.session_state["r1_esp_sorexp_csv"]=(esp.stem+"_SOREXP.csv",esx)
            st.session_state["r1_can"]=(can.name,can.read_text(encoding='utf-8')) if can.exists() else None
            st.session_state["r1_html"]=(ht.name,ht.read_bytes()) if ht.exists() else None
            st.session_state["r1_day_filter"]=None

# ── Run 2 ─────────────────────────────────────────────────────────────────────
if st.session_state.get("_run2"):
    st.session_state["_run2"]=False; sc=semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _t:
        tmp=Path(_t); p=save_uploads(tmp)
        with st.spinner("Generando GD base…"): gd,_,r0=run_gd(p,tmp,sc)
        if r0.returncode!=0 or not gd.exists(): st.error("Error GD base."); show_log(r0,True)
        else:
            out=tmp/f"gantt_1h_{sc}.xlsx"
            with st.spinner("Generando Gantt 1H…"):
                r=subprocess.run([sys.executable,str(BASE_DIR/"gantt_1h.py"),
                    str(p["cap"]),str(gd),str(p["bloques"]),str(out),"Hoja1"],
                    capture_output=True,text=True,timeout=180)
            show_log(r)
            if r.returncode==0 and out.exists(): st.session_state["r2_gantt"]=(out.name,out.read_bytes())
            else: st.error("Error en Gantt.")

# ── Run 3 ─────────────────────────────────────────────────────────────────────
if st.session_state.get("_run3"):
    st.session_state["_run3"]=False; sc=semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _t:
        tmp=Path(_t); p=save_uploads(tmp)
        with st.spinner("Generando GD base…"): gd,_,r0=run_gd(p,tmp,sc)
        if r0.returncode!=0 or not gd.exists(): st.error("Error GD base."); show_log(r0,True)
        else:
            out=tmp/f"sorter_map_{sc}.xlsx"
            with st.spinner("Generando Sorter Map…"):
                cmd=[sys.executable,str(BASE_DIR/"sorter_map_por_dia.py"),
                     str(p["cap"]),str(gd),str(p["bloques"]),str(out),"Hoja1"]
                if "parrilla" in p and "gd" in p: cmd+=[str(p["parrilla"]),sheet.strip(),str(p["gd"])]
                r=subprocess.run(cmd,capture_output=True,text=True,timeout=180)
            show_log(r)
            if r.returncode==0 and out.exists(): st.session_state["r3_map"]=(out.name,out.read_bytes())
            else: st.error("Error en Sorter Map.")

# ── Results 1 ─────────────────────────────────────────────────────────────────
if st.session_state["r1_gd"] is not None:
    sc=semana or "SEMANA"
    H(f'<div class="mng-res">Configuración <b>{sc}</b> — lista para descargar</div>')
    if st.session_state["r1_day_filter"]:
        st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")
    H('<div class="mng-dl">Ficheros principales</div>')
    c1,c2=st.columns(2)
    with c1:
        n,d=st.session_state["r1_gd"]
        st.download_button("GD completo (.xlsx)",data=d,file_name=n,mime=XLSX_MIME,use_container_width=True)
        st.caption("Subir a DXC / MAR")
    with c2:
        if st.session_state["r1_esp"]:
            n,d=st.session_state["r1_esp"]
            st.download_button("Solo especiales (.xlsx)",data=d,file_name=n,mime=XLSX_MIME,use_container_width=True)
            st.caption("Filas nuevas a añadir en DXC")
    with st.expander("Más descargas — CSV · Canceladas · HTML"):
        H('<div class="mng-dl">CSV para importar en DXC</div>')
        cc1,cc2,cc3,cc4=st.columns(4)
        for col,key,lbl in [(cc1,"r1_postex_csv","POSTEX completo"),(cc2,"r1_sorexp_csv","SOREXP completo"),
                             (cc3,"r1_esp_postex_csv","POSTEX esp."),(cc4,"r1_esp_sorexp_csv","SOREXP esp.")]:
            with col:
                if st.session_state[key]:
                    n,d=st.session_state[key]
                    st.download_button(lbl,data=d,file_name=n,mime="text/csv",use_container_width=True)
        H('<div class="mng-dl">Otros</div>')
        d1,d2=st.columns(2)
        with d1:
            if st.session_state["r1_can"]:
                n,t=st.session_state["r1_can"]
                st.download_button("Canceladas (.txt)",data=t,file_name=n,mime="text/plain",use_container_width=True)
                with st.expander("Ver canceladas"): st.text(t)
        with d2:
            if st.session_state["r1_html"]:
                n,d=st.session_state["r1_html"]
                st.download_button("Resumen HTML",data=d,file_name=n,mime="text/html",use_container_width=True)
    with st.expander("Filtrar por día y regenerar"):
        sel=st.multiselect("Días a incluir",options=ALL_DAYS,default=[],
                           key="day_filter_sel",placeholder="Selecciona días…")
        if st.button("Regenerar con filtro",key="regen_filter",disabled=not(sel and base_ok)):
            with tempfile.TemporaryDirectory() as _t:
                tmp=Path(_t); p=save_uploads(tmp)
                with st.spinner(f"Regenerando para {', '.join(sel)}…"):
                    gd,ht,r=run_gd(p,tmp,sc,",".join(sel))
                if r.returncode==0 and gd.exists():
                    esp=Path(str(gd).replace('.xlsx','_SOLO_ESPECIALES.xlsx'))
                    can=Path(str(gd).replace('.xlsx','_CANCELADAS.txt'))
                    st.session_state["r1_gd"]=(gd.name,gd.read_bytes())
                    st.session_state["r1_esp"]=(esp.name,esp.read_bytes()) if esp.exists() else None
                    st.session_state["r1_can"]=(can.name,can.read_text(encoding='utf-8')) if can.exists() else None
                    st.session_state["r1_html"]=(ht.name,ht.read_bytes()) if ht.exists() else None
                    st.session_state["r1_day_filter"]=sel; st.rerun()
                else: st.error("Error en regeneración."); show_log(r,True)

# ── Results 2 ─────────────────────────────────────────────────────────────────
if st.session_state["r2_gantt"] is not None:
    H('<div class="mng-res">Gantt 1H — listo para descargar</div>')
    n,d=st.session_state["r2_gantt"]
    st.download_button("Gantt 1H (.xlsx)",data=d,file_name=n,mime=XLSX_MIME,use_container_width=True)
    st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")

# ── Results 3 ─────────────────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    H('<div class="mng-res">Sorter Map — listo para descargar</div>')
    n,d=st.session_state["r3_map"]
    st.download_button("Sorter Map (.xlsx)",data=d,file_name=n,mime=XLSX_MIME,use_container_width=True)
    st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")

# ── Footer ────────────────────────────────────────────────────────────────────
H('<div class="mng-ft">v0.09 · VDL B2B · Estrictamente confidencial · Solo uso interno</div>')
