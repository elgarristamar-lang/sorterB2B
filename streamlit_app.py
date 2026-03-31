# Version: 0.10 — clean minimal style
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

# Minimal CSS — only tweak fonts and hide chrome, let Streamlit handle layout
st.markdown("""
<style>
* { font-family: "Aptos Display", "Trebuchet MS", Arial, sans-serif !important; }
#MainMenu, footer, header { visibility: hidden; }
.stDeployButton { display: none; }
</style>
""", unsafe_allow_html=True)

# ── Session state ─────────────────────────────────────────────────────────────
for k in ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map",
          "r1_day_filter","r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
    if k not in st.session_state: st.session_state[k] = None

# ── Title ─────────────────────────────────────────────────────────────────────
st.title("🏭 Sorter VDL B2B")
st.caption("Configurador de semanas especiales — VDL B2B")
st.divider()

# ── Guide ─────────────────────────────────────────────────────────────────────
with st.expander("📖 Cómo usar esta herramienta"):
    st.markdown("""
**Ficheros de entrada**

| Fichero | Oblig. | Descripción |
|---|:---:|---|
| `parrilla_de_salidas.xlsx` | ✅ | Columnas: `PLAYA`, `DIA_SALIDA`, `DIA_SALIDA_NEW`, `TIPO_SALIDA`, `ID_CLUSTER`. Incluir hoja `Resumen Bloques`. |
| GRUPO_DESTINOS `.xlsx` | ✅ | Consulta DXC 9066 — Consulta destinos por zona |
| `ramp_capacity.csv` | ✅ | CSV `;` con columnas `RAMP` y `PALLETS` |
| Bloques horarios `.xlsx` | ⚠️ | Solo para Gantt y Sorter Map |
| Superplayas `.xlsx` | ➖ | Opcional · `AGRUPACION_PLAYA` · `SUPERPLAYA` |

**Tipos de salida:** `HABITUAL` → sin cambios · `CANCELADA` → eliminar · `ESPECIAL DIA CAMBIO` → reasignar · `IRREGULAR` → ignorar
    """)

st.divider()

# ── 01 · Ficheros de entrada ──────────────────────────────────────────────────
st.subheader("📂 Ficheros de entrada")

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

# Status chips
if any([f_parrilla, f_gd, f_cap, f_bloques]):
    items = [("Parrilla", f_parrilla, True), ("Grupo Destinos", f_gd, True),
             ("Capacidad", f_cap, True), ("Bloques", f_bloques, False)]
    cols = st.columns(4)
    for col, (label, f, required) in zip(cols, items):
        with col:
            if f:
                st.success(f"✓ {label}")
            elif required:
                st.warning(f"· {label}")
            else:
                st.info(f"· {label} (opc.)")

st.divider()

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

# ── 02 · Generar ─────────────────────────────────────────────────────────────
st.subheader("⚡ Generar outputs")

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("**1 · Configuración DXC**")
    st.caption("GRUPO_DESTINOS + resumen HTML")
    btn1 = st.button("⚙️ Generar", key="go1", type="primary",
                     disabled=not base_ok, use_container_width=True)
    if not base_ok: st.caption("_Requiere parrilla, GD y capacidad_")
with c2:
    st.markdown("**2 · Gantt 1H**")
    st.caption("Bloques × rampas × hora")
    btn2 = st.button("📊 Generar", key="go2", type="primary",
                     disabled=not vis_ok, use_container_width=True)
    if not vis_ok: st.caption("_Requiere bloques horarios_")
with c3:
    st.markdown("**3 · Sorter Map**")
    st.caption("1 pestaña por día, slots físicos")
    btn3 = st.button("🗺️ Generar", key="go3", type="primary",
                     disabled=not vis_ok, use_container_width=True)
    if not vis_ok: st.caption("_Requiere bloques horarios_")

if btn1:
    for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter",
              "r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
        st.session_state[k] = None
    st.session_state["_run1"] = True
if btn2: st.session_state["r2_gantt"]=None; st.session_state["_run2"]=True
if btn3: st.session_state["r3_map"]=None;   st.session_state["_run3"]=True

st.divider()

# ── Run 1 ─────────────────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana or "SEMANA"
    with tempfile.TemporaryDirectory() as _t:
        tmp=Path(_t); p=save_uploads(tmp)
        with st.spinner(f"Procesando parrilla {sc}…"): gd,ht,r=run_gd(p,tmp,sc)
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
    sc = semana or "SEMANA"
    st.success(f"✅ Configuración **{sc}** generada")
    if st.session_state["r1_day_filter"]:
        st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")

    st.markdown("**Ficheros principales**")
    c1, c2 = st.columns(2)
    with c1:
        n,d = st.session_state["r1_gd"]
        st.download_button("⬇️ GD completo (.xlsx)", data=d, file_name=n,
                           mime=XLSX_MIME, use_container_width=True)
        st.caption("Subir a DXC / MAR")
    with c2:
        if st.session_state["r1_esp"]:
            n,d = st.session_state["r1_esp"]
            st.download_button("⬇️ Solo especiales (.xlsx)", data=d, file_name=n,
                               mime=XLSX_MIME, use_container_width=True)
            st.caption("Filas nuevas a añadir en DXC")

    with st.expander("📥 Más descargas — CSV · Canceladas · HTML"):
        st.markdown("**CSV para importar en DXC**")
        cc1,cc2,cc3,cc4 = st.columns(4)
        for col,key,lbl in [(cc1,"r1_postex_csv","POSTEX completo"),
                             (cc2,"r1_sorexp_csv","SOREXP completo"),
                             (cc3,"r1_esp_postex_csv","POSTEX esp."),
                             (cc4,"r1_esp_sorexp_csv","SOREXP esp.")]:
            with col:
                if st.session_state[key]:
                    n,d = st.session_state[key]
                    st.download_button(lbl, data=d, file_name=n,
                                       mime="text/csv", use_container_width=True)
        st.divider()
        d1,d2 = st.columns(2)
        with d1:
            if st.session_state["r1_can"]:
                n,t = st.session_state["r1_can"]
                st.download_button("⬇️ Canceladas (.txt)", data=t, file_name=n,
                                   mime="text/plain", use_container_width=True)
                with st.expander("Ver canceladas"): st.text(t)
        with d2:
            if st.session_state["r1_html"]:
                n,d = st.session_state["r1_html"]
                st.download_button("⬇️ Resumen HTML", data=d, file_name=n,
                                   mime="text/html", use_container_width=True)
                st.caption("Informe interactivo con gráfico de capacidad")

    with st.expander("🔍 Filtrar por día y regenerar"):
        sel = st.multiselect("Días a incluir", options=ALL_DAYS, default=[],
                             key="day_filter_sel", placeholder="Selecciona días…")
        if st.button("⚙️ Regenerar con filtro", key="regen_filter",
                     disabled=not (sel and base_ok)):
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
    st.success("✅ Gantt 1H generado")
    n,d = st.session_state["r2_gantt"]
    st.download_button("⬇️ Gantt 1H (.xlsx)", data=d, file_name=n,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")

# ── Results 3 ─────────────────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    st.success("✅ Sorter Map generado")
    n,d = st.session_state["r3_map"]
    st.download_button("⬇️ Sorter Map (.xlsx)", data=d, file_name=n,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")

st.divider()
st.caption("v0.10 · VDL B2B · Estrictamente confidencial")
