# Version: 0.05
import streamlit as st
import subprocess, sys, tempfile, datetime as dt
from pathlib import Path

BASE_DIR = Path(__file__).parent

st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")
st.markdown("<style>.block-container{max-width:780px}</style>", unsafe_allow_html=True)

st.markdown("## 🏭 Sorter VDL B2B")
st.markdown("Configurador de semanas especiales — VDL B2B")
st.divider()

# ── Session state ─────────────────────────────────────────────────────────────
for key in ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map","r1_day_filter"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Inputs ────────────────────────────────────────────────────────────────────
st.markdown("### Ficheros de entrada")
col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader("Parrilla de salidas", type=["xlsx"],
                                   help="Debe incluir la hoja Resumen Bloques")
    sheet  = st.text_input("Nombre de hoja", value="parrilla_test_s14")
    semana = st.text_input("Semana", value="S14")
with col2:
    f_gd  = st.file_uploader("GRUPO_DESTINOS", type=["xlsx"],
                               help="Export DXC o fichero clásico")
    f_cap = st.file_uploader("Capacidad de rampas", type=["csv"],
                               help="CSV con columnas RAMP;PALLETS")

f_bloques = st.file_uploader(
    "Bloques horarios *(necesario para Gantt y Sorter Map)*", type=["xlsx"],
    help="Columnas: NUEVO BLOQUE · Día LIBERACIÓN · Hora LIBERACIÓN · Día DESACTIVACIÓN · Hora DESACTIVACIÓN")

f_superplaya = st.file_uploader(
    "Superplayas *(opcional — mejora la agrupación de rampas)*", type=["xlsx"],
    help="Columnas: AGRUPACION_PLAYA · SUPERPLAYA — define qué destinos deben ir juntos en rampas contiguas")

st.divider()

# ── Helpers ───────────────────────────────────────────────────────────────────
ALL_DAYS = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]

def save_uploads(tmp: Path):
    p = {}
    for key, f, name in [
        ("parrilla",    f_parrilla,    "parrilla.xlsx"),
        ("gd",          f_gd,          "gd.xlsx"),
        ("cap",         f_cap,          "cap.csv"),
        ("bloques",     f_bloques,      "bloques.xlsx"),
        ("superplaya",  f_superplaya,   "superplaya.xlsx"),
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
            sheet.strip(), sc, str(gd), str(html)]
    cmd.append(days_arg)  # argv[8] — always pass (empty string = no filter)
    if "superplaya" in p:
        cmd.append(str(p["superplaya"]))  # argv[9]
    r = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    return gd, html, r

def show_log(r, expanded=False):
    lines = [l for l in (r.stdout + r.stderr).splitlines() if l.strip()]
    with st.expander("Ver log", expanded=expanded or r.returncode != 0):
        for line in lines:
            if "✓" in line:                                           st.success(line)
            elif "❌" in line:                                         st.error(line)
            elif "⚠" in line or "E2" in line or "Sin config" in line: st.warning(line)
            else:                                                       st.text(line)
    return lines

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ── Three action buttons ──────────────────────────────────────────────────────
st.markdown("### Acciones")

base_ok = bool(f_parrilla and f_gd and f_cap)
vis_ok  = bool(base_ok and f_bloques)

b1, b2, b3 = st.columns(3)
with b1:
    st.markdown("**1 · Configuración DXC**")
    st.caption("GRUPO_DESTINOS + resumen HTML")
    if st.button("⚙️ Generar", key="go1", type="primary",
                 disabled=not base_ok, use_container_width=True):
        # Clear previous results so we re-run
        for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter"]:
            st.session_state[k] = None
        st.session_state["_run1"] = True
    if not base_ok:
        st.caption("_Sube parrilla, GD y capacidad_")

with b2:
    st.markdown("**2 · Gantt 1H**")
    st.caption("Visual bloques × rampas × hora")
    if st.button("📊 Generar", key="go2", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r2_gantt"] = None
        st.session_state["_run2"] = True
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

with b3:
    st.markdown("**3 · Sorter Map por día**")
    st.caption("1 pestaña por día, slots físicos")
    if st.button("🗺 Generar", key="go3", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r3_map"] = None
        st.session_state["_run3"] = True
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

st.divider()

# ── Execute action 1 ──────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Procesando parrilla y asignando rampas…"):
            gd, html, r = run_gd(p, tmp, sc)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
        else:
            esp_path = Path(str(gd).replace('.xlsx', '_SOLO_ESPECIALES.xlsx'))
            can_path = Path(str(gd).replace('.xlsx', '_CANCELADAS.txt'))
            st.session_state["r1_gd"]    = (gd.name,   gd.read_bytes())
            st.session_state["r1_esp"]   = (esp_path.name, esp_path.read_bytes()) if esp_path.exists() else None
            st.session_state["r1_can"]   = (can_path.name, can_path.read_text(encoding='utf-8')) if can_path.exists() else None
            st.session_state["r1_html"]  = (html.name,  html.read_bytes()) if html.exists() else None
            st.session_state["r1_day_filter"] = None  # full run, no day filter

# ── Execute action 2 ──────────────────────────────────────────────────────────
if st.session_state.get("_run2"):
    st.session_state["_run2"] = False
    sc = semana.strip() or sheet.strip().upper()
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
    sc = semana.strip() or sheet.strip().upper()
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
                r = subprocess.run(
                    [sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                     str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"],
                    capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.session_state["r3_map"] = (out.name, out.read_bytes())
            else:
                st.error("El Sorter Map terminó con error.")

# ── Show results action 1 ─────────────────────────────────────────────────────
if st.session_state["r1_gd"] is not None:
    gd_name, gd_bytes = st.session_state["r1_gd"]
    sc = semana.strip() or sheet.strip().upper()
    st.success(f"✓ Configuración {sc} generada")

    # ── Day filter for filtered re-extraction ──
    with st.expander("🔍 Filtrar por día y regenerar especiales"):
        selected_days = st.multiselect(
            "Días a incluir en el GD filtrado",
            options=ALL_DAYS,
            default=[],
            key="day_filter_sel",
            placeholder="Selecciona días…",
        )
        if st.button("⚙️ Regenerar con filtro", key="regen_filter",
                     disabled=not (selected_days and f_parrilla and f_gd and f_cap)):
            days_arg = ",".join(selected_days)
            with tempfile.TemporaryDirectory() as _tmp:
                tmp = Path(_tmp)
                p   = save_uploads(tmp)
                with st.spinner(f"Regenerando para {', '.join(selected_days)}…"):
                    gd, html, r = run_gd(p, tmp, sc, days_arg)
                if r.returncode == 0 and gd.exists():
                    esp_path = Path(str(gd).replace('.xlsx','_SOLO_ESPECIALES.xlsx'))
                    can_path = Path(str(gd).replace('.xlsx','_CANCELADAS.txt'))
                    st.session_state["r1_gd"]   = (gd.name,   gd.read_bytes())
                    st.session_state["r1_esp"]  = (esp_path.name, esp_path.read_bytes()) if esp_path.exists() else None
                    st.session_state["r1_can"]  = (can_path.name, can_path.read_text(encoding='utf-8')) if can_path.exists() else None
                    st.session_state["r1_html"] = (html.name,  html.read_bytes()) if html.exists() else None
                    st.session_state["r1_day_filter"] = selected_days
                    st.rerun()
                else:
                    st.error("Error en regeneración.")
                    show_log(r, expanded=True)

    # Day filter badge
    if st.session_state["r1_day_filter"]:
        st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")

    # Downloads row 1: GD completo + solo especiales
    c1, c2 = st.columns(2)
    with c1:
        name, data = st.session_state["r1_gd"]
        st.download_button("⬇️ GD completo", data=data, file_name=name,
                           mime=XLSX_MIME, use_container_width=True)
        st.caption("GD completo — subir a DXC / MAR")
    with c2:
        if st.session_state["r1_esp"]:
            name, data = st.session_state["r1_esp"]
            st.download_button("⬇️ Solo especiales", data=data, file_name=name,
                               mime=XLSX_MIME, use_container_width=True)
            st.caption("Solo filas nuevas a añadir en DXC")

    # Downloads row 2: canceladas + resumen HTML
    c3, c4 = st.columns(2)
    with c3:
        if st.session_state["r1_can"]:
            name, txt = st.session_state["r1_can"]
            st.download_button("⬇️ Canceladas.txt", data=txt, file_name=name,
                               mime="text/plain", use_container_width=True)
            with st.expander("Ver canceladas"):
                st.text(txt)
            st.caption("Salidas a eliminar del sorter")
    with c4:
        if st.session_state["r1_html"]:
            name, data = st.session_state["r1_html"]
            st.download_button("⬇️ Resumen HTML", data=data, file_name=name,
                               mime="text/html", use_container_width=True)
            st.caption("Informe con gráfico interactivo")

# ── Show results action 2 ─────────────────────────────────────────────────────
if st.session_state["r2_gantt"] is not None:
    name, data = st.session_state["r2_gantt"]
    st.success("✓ Gantt 1H generado")
    st.download_button("⬇️ Gantt 1H.xlsx", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")

# ── Show results action 3 ─────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    name, data = st.session_state["r3_map"]
    st.success("✓ Sorter Map generado")
    st.download_button("⬇️ Sorter Map.xlsx", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")

st.divider()
st.caption("v0.05 · VDL B2B · Estrictamente confidencial")
