# Version: 0.04
import streamlit as st
import subprocess, sys, tempfile, datetime as dt
from pathlib import Path

BASE_DIR = Path(__file__).parent

st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")
st.markdown("<style>.block-container{max-width:780px}</style>", unsafe_allow_html=True)

st.markdown("## 🏭 Sorter VDL B2B")
st.markdown("Configurador de semanas especiales — VDL B2B")
st.divider()

# ── Session state init ────────────────────────────────────────────────────────
for key in ["result1", "result2", "result3"]:
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

# ── Day filter (for action 1 only) ───────────────────────────────────────────
ALL_DAYS = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]
with st.expander("Filtro de días (opcional — para el GRUPO_DESTINOS)"):
    st.caption("Selecciona solo los días que quieres incluir en el GD. Sin selección = todos los días.")
    selected_days = st.multiselect(
        "Días a incluir",
        options=ALL_DAYS,
        default=[],
        placeholder="Todos los días (sin filtro)",
    )

st.divider()

# ── Helpers ───────────────────────────────────────────────────────────────────
def save_uploads(tmp: Path):
    p = {}
    for key, f, name in [
        ("parrilla", f_parrilla, "parrilla.xlsx"),
        ("gd",       f_gd,       "gd.xlsx"),
        ("cap",      f_cap,       "cap.csv"),
        ("bloques",  f_bloques,   "bloques.xlsx"),
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
    cmd = [sys.executable, str(BASE_DIR / "process_parrilla.py"),
           str(p["parrilla"]), str(p["gd"]), str(p["cap"]),
           sheet.strip(), sc, str(gd), str(html)]
    if days_arg:
        cmd.append(days_arg)
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
        st.session_state.result1 = "running"
        st.session_state.result2 = None
        st.session_state.result3 = None
    if not base_ok:
        st.caption("_Sube parrilla, GD y capacidad_")

with b2:
    st.markdown("**2 · Gantt 1H**")
    st.caption("Visual bloques × rampas × hora")
    if st.button("📊 Generar", key="go2", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state.result1 = None
        st.session_state.result2 = "running"
        st.session_state.result3 = None
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

with b3:
    st.markdown("**3 · Sorter Map por día**")
    st.caption("1 pestaña por día, slots físicos")
    if st.button("🗺 Generar", key="go3", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state.result1 = None
        st.session_state.result2 = None
        st.session_state.result3 = "running"
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

st.divider()

# ── Action 1 ──────────────────────────────────────────────────────────────────
if st.session_state.result1 == "running":
    sc = semana.strip() or sheet.strip().upper()
    days_arg = ",".join(selected_days) if selected_days else ""
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Procesando parrilla y asignando rampas…"):
            gd, html, r = run_gd(p, tmp, sc, days_arg)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
            st.session_state.result1 = "error"
        else:
            day_label = f" · días: {', '.join(selected_days)}" if selected_days else ""
            st.success(f"✓ Configuración {sc} generada{day_label}")

            # Locate side-output files (written alongside gd_out by process_parrilla)
            esp_path = Path(str(gd).replace('.xlsx', '_SOLO_ESPECIALES.xlsx'))
            can_path = Path(str(gd).replace('.xlsx', '_CANCELADAS.txt'))

            # Row 1: GD completo + solo especiales
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇️ GD completo",
                    data=gd.read_bytes(), file_name=gd.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("GD completo — subir a DXC / MAR")
            with c2:
                if esp_path.exists():
                    st.download_button("⬇️ Solo especiales",
                        data=esp_path.read_bytes(), file_name=esp_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True)
                    st.caption("Solo filas nuevas a añadir en DXC")

            # Row 2: Canceladas + Resumen HTML
            c3, c4 = st.columns(2)
            with c3:
                if can_path.exists():
                    txt = can_path.read_text(encoding='utf-8')
                    st.download_button("⬇️ Canceladas.txt",
                        data=txt, file_name=can_path.name,
                        mime="text/plain", use_container_width=True)
                    # Show inline preview
                    with st.expander("Ver canceladas"):
                        st.text(txt)
                    st.caption("Salidas a eliminar del sorter")
            with c4:
                if html.exists():
                    st.download_button("⬇️ Resumen HTML",
                        data=html.read_bytes(), file_name=html.name,
                        mime="text/html", use_container_width=True)
                    st.caption("Informe con gráfico interactivo")
            st.session_state.result1 = "done"

# ── Action 2 ──────────────────────────────────────────────────────────────────
if st.session_state.result2 == "running":
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, html, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando el GD base.")
            show_log(r0, expanded=True)
            st.session_state.result2 = "error"
        else:
            out = tmp / f"gantt_1h_{sc}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            with st.spinner("Generando Gantt 1H…"):
                r = subprocess.run(
                    [sys.executable, str(BASE_DIR / "gantt_1h.py"),
                     str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"],
                    capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.success("✓ Gantt 1H generado")
                st.download_button("⬇️ Gantt 1H.xlsx",
                    data=out.read_bytes(), file_name=out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")
                st.session_state.result2 = "done"
            else:
                st.error("El Gantt terminó con error.")
                st.session_state.result2 = "error"

# ── Action 3 ──────────────────────────────────────────────────────────────────
if st.session_state.result3 == "running":
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, html, r0 = run_gd(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando el GD base.")
            show_log(r0, expanded=True)
            st.session_state.result3 = "error"
        else:
            out = tmp / f"sorter_map_{sc}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            with st.spinner("Generando Sorter Map…"):
                r = subprocess.run(
                    [sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                     str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1"],
                    capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.success("✓ Sorter Map generado")
                st.download_button("⬇️ Sorter Map.xlsx",
                    data=out.read_bytes(), file_name=out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")
                st.session_state.result3 = "done"
            else:
                st.error("El Sorter Map terminó con error.")
                st.session_state.result3 = "error"

st.divider()
st.caption("v0.04 · VDL B2B · Estrictamente confidencial")
