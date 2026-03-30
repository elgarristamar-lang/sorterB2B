# Version: 0.03
import streamlit as st
import subprocess, sys, tempfile, datetime as dt
from pathlib import Path

BASE_DIR = Path(__file__).parent

st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")
st.markdown("<style>.block-container{max-width:780px}</style>", unsafe_allow_html=True)

st.markdown("## 🏭 Sorter VDL B2B")
st.markdown("Configurador de semanas especiales — VDL B2B Mango")
st.divider()

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

def run_gd_base(p, tmp, semana_clean):
    """Run process_parrilla and return (gd_path, html_path, ok)."""
    gd   = tmp / f"GRUPO_DESTINOS_{semana_clean}.xlsx"
    html = tmp / f"resumen_{semana_clean}.html"
    r = subprocess.run([
        sys.executable, str(BASE_DIR / "process_parrilla.py"),
        str(p["parrilla"]), str(p["gd"]), str(p["cap"]),
        sheet.strip(), semana_clean,
        str(gd), str(html),
    ], capture_output=True, text=True, timeout=180)
    return gd, html, r

def show_log(r, expanded=False):
    lines = [l for l in (r.stdout + r.stderr).splitlines() if l.strip()]
    with st.expander("Ver log", expanded=expanded or r.returncode != 0):
        for line in lines:
            if "✓" in line:                                      st.success(line)
            elif "❌" in line:                                    st.error(line)
            elif "⚠" in line or "E2" in line or "Sin config" in line: st.warning(line)
            else:                                                  st.text(line)
    return lines

# ── Three action buttons ──────────────────────────────────────────────────────
st.markdown("### Acciones")

base_ok = bool(f_parrilla and f_gd and f_cap)
vis_ok  = bool(base_ok and f_bloques)

b1, b2, b3 = st.columns(3)

with b1:
    st.markdown("**1 · Configuración DXC**")
    st.caption("GRUPO_DESTINOS + resumen HTML")
    go1 = st.button("⚙️ Generar", key="go1", type="primary",
                    disabled=not base_ok, use_container_width=True)
    if not base_ok:
        st.caption("_Sube parrilla, GD y capacidad_")

with b2:
    st.markdown("**2 · Gantt 1H**")
    st.caption("Visual bloques × rampas × hora")
    go2 = st.button("📊 Generar", key="go2", type="primary",
                    disabled=not vis_ok, use_container_width=True)
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

with b3:
    st.markdown("**3 · Sorter Map por día**")
    st.caption("1 pestaña por día, slots físicos")
    go3 = st.button("🗺 Generar", key="go3", type="primary",
                    disabled=not vis_ok, use_container_width=True)
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

st.divider()

# ── Action 1 ──────────────────────────────────────────────────────────────────
if go1:
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        gd, html, r = run_gd_base(p, tmp, sc)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
        else:
            st.success(f"✓ Configuración S{sc} generada — 0 colisiones de rampa")
            c1, c2 = st.columns(2)
            with c1:
                st.download_button("⬇️ GRUPO_DESTINOS.xlsx",
                    data=gd.read_bytes(), file_name=gd.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("Subir a DXC / MAR")
            with c2:
                if html.exists():
                    st.download_button("⬇️ Resumen HTML",
                        data=html.read_bytes(), file_name=html.name,
                        mime="text/html", use_container_width=True)
                    st.caption("Informe con gráfico interactivo")

# ── Action 2 ──────────────────────────────────────────────────────────────────
if go2:
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, html, r0 = run_gd_base(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando el GD base.")
            show_log(r0, expanded=True)
        else:
            out = tmp / f"gantt_1h_{sc}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            r = subprocess.run([
                sys.executable, str(BASE_DIR / "gantt_1h.py"),
                str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1",
            ], capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.success("✓ Gantt 1H generado")
                st.download_button("⬇️ Gantt 1H.xlsx",
                    data=out.read_bytes(), file_name=out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")
            else:
                st.error("El Gantt terminó con error.")

# ── Action 3 ──────────────────────────────────────────────────────────────────
if go3:
    sc = semana.strip() or sheet.strip().upper()
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        with st.spinner("Generando GD base…"):
            gd, html, r0 = run_gd_base(p, tmp, sc)
        if r0.returncode != 0 or not gd.exists():
            st.error("Error generando el GD base.")
            show_log(r0, expanded=True)
        else:
            out = tmp / f"sorter_map_{sc}_{dt.datetime.now():%Y%m%d_%H%M%S}.xlsx"
            r = subprocess.run([
                sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1",
            ], capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.success("✓ Sorter Map generado")
                st.download_button("⬇️ Sorter Map.xlsx",
                    data=out.read_bytes(), file_name=out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
                st.caption("Hojas: DOM · LUN · MAR · MIÉ · JUE · VIE · SÁB · LEYENDA")
            else:
                st.error("El Sorter Map terminó con error.")

st.divider()
st.caption("v0.03 · Mango Logística VDL B2B · Estrictamente confidencial")
