# Version: 0.01
import streamlit as st
import subprocess, sys, os, glob, shutil, tempfile
from datetime import datetime
from pathlib import Path

BASE_DIR = Path(__file__).parent

st.set_page_config(
    page_title="Sorter VDL B2B",
    page_icon="🏭",
    layout="centered",
)

# ── Estilos ───────────────────────────────────────────────────────────────────
st.markdown("""
<style>
  .block-container { max-width: 780px; }
  .stAlert { border-radius: 8px; }
  div[data-testid="stFileUploader"] { margin-bottom: 0; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("## 🏭 Sorter VDL B2B")
st.markdown("Configurador de semanas especiales — genera el GRUPO_DESTINOS con asignación automática de rampas.")
st.divider()

# ── Inputs ────────────────────────────────────────────────────────────────────
st.markdown("### 01 — Ficheros de entrada")

col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader("Parrilla de salidas", type=["xlsx"], key="parrilla",
                                   help="Debe incluir la hoja Resumen Bloques")
    sheet = st.text_input("Nombre de hoja", value="parrilla_test_s14",
                           help="Hoja de la parrilla con los datos de la semana")
    semana = st.text_input("Semana", value="S14",
                            help="Identificador para los ficheros de salida")

with col2:
    f_gd  = st.file_uploader("GRUPO_DESTINOS", type=["xlsx"], key="gd",
                               help="Export DXC o fichero clásico")
    f_cap = st.file_uploader("Capacidad de rampas", type=["csv"], key="cap",
                               help="CSV con columnas RAMP;PALLETS")

st.markdown("**Visualizaciones operativas** *(opcional — activa Gantt y Sorter Map)*")
f_bloques = st.file_uploader("Bloques horarios", type=["xlsx"], key="bloques",
                               help="Columnas: NUEVO BLOQUE · Día LIBERACIÓN · Hora LIBERACIÓN · Día DESACTIVACIÓN · Hora DESACTIVACIÓN")

st.divider()

# ── Run ───────────────────────────────────────────────────────────────────────
run = st.button("⚙️ Generar configuración", type="primary",
                 disabled=not (f_parrilla and f_gd and f_cap),
                 use_container_width=True)

if not (f_parrilla and f_gd and f_cap):
    st.caption("Sube los tres ficheros obligatorios para continuar.")

if run:
    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)

        # Save uploads
        def save(f, name): p = tmp / name; p.write_bytes(f.read()); return p

        p_parrilla = save(f_parrilla, "parrilla.xlsx")
        p_gd       = save(f_gd,       "gd.xlsx")
        p_cap      = save(f_cap,       "cap.csv")
        p_bloques  = save(f_bloques, "bloques.xlsx") if f_bloques else None

        semana_clean = semana.strip() or sheet.strip().upper()

        # ── Step 1: process_parrilla ──────────────────────────────────────────
        st.markdown("### 02 — Proceso")
        log_box = st.empty()

        gd_out   = tmp / f"GRUPO_DESTINOS_{semana_clean}.xlsx"
        html_out = tmp / f"resumen_{semana_clean}.html"

        with st.spinner("Procesando parrilla y asignando rampas…"):
            r1 = subprocess.run(
                [sys.executable, str(BASE_DIR / "process_parrilla.py"),
                 str(p_parrilla), str(p_gd), str(p_cap),
                 sheet.strip(), semana_clean],
                capture_output=True, text=True, timeout=120,
                env={**os.environ, "HOME": str(tmp)},
            )

        # Move outputs from tmp HOME to our tmp dir
        for pattern in [f"GRUPO_DESTINOS_{semana_clean}*.xlsx",
                        f"resumen_sorter_{semana_clean}*.html"]:
            found = sorted((tmp).glob(pattern))
            if not found:  # also check ~
                found = sorted(Path.home().glob(pattern), key=os.path.getmtime)
            if found:
                newest = max(found, key=lambda p: p.stat().st_mtime)
                if "GRUPO_DESTINOS" in pattern:
                    shutil.copy(newest, gd_out)
                else:
                    shutil.copy(newest, html_out)

        log_lines = [l for l in (r1.stdout + r1.stderr).splitlines() if l.strip()]

        # Show log
        ok1 = r1.returncode == 0 and gd_out.exists()
        with log_box.container():
            for line in log_lines:
                if line.startswith("✓") or "Asignados OK" in line:
                    st.success(line)
                elif line.startswith("❌") or "error" in line.lower():
                    st.error(line)
                elif line.startswith("⚠") or "E2" in line or "Sin config" in line:
                    st.warning(line)
                else:
                    st.text(line)

        if not ok1:
            st.error("El proceso terminó con error. Revisa el log.")
            st.stop()

        # ── Step 2 & 3: visual scripts ────────────────────────────────────────
        gantt_out = map_out = None

        if p_bloques and gd_out.exists():
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            gantt_out = tmp / f"gantt_1h_{semana_clean}_{ts}.xlsx"
            map_out   = tmp / f"sorter_map_{semana_clean}_{ts}.xlsx"

            col_g, col_m = st.columns(2)
            with col_g:
                with st.spinner("Generando Gantt 1H…"):
                    r2 = subprocess.run(
                        [sys.executable, str(BASE_DIR / "gantt_1h.py"),
                         str(p_cap), str(gd_out), str(p_bloques),
                         str(gantt_out), "Hoja1"],
                        capture_output=True, text=True, timeout=120,
                    )
                if r2.returncode == 0:
                    st.success("Gantt 1H generado")
                else:
                    st.warning("Gantt terminó con advertencias")
                    gantt_out = None

            with col_m:
                with st.spinner("Generando Sorter Map…"):
                    r3 = subprocess.run(
                        [sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                         str(p_cap), str(gd_out), str(p_bloques),
                         str(map_out), "Hoja1"],
                        capture_output=True, text=True, timeout=120,
                    )
                if r3.returncode == 0:
                    st.success("Sorter Map generado")
                else:
                    st.warning("Sorter Map terminó con advertencias")
                    map_out = None

        # ── Downloads ─────────────────────────────────────────────────────────
        st.divider()
        st.markdown("### 03 — Descargar resultados")

        dcols = st.columns(2)
        col_idx = 0

        if gd_out.exists():
            with dcols[col_idx % 2]:
                st.download_button(
                    label="⬇️ GRUPO_DESTINOS.xlsx",
                    data=gd_out.read_bytes(),
                    file_name=gd_out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption("Subir a DXC/MAR")
            col_idx += 1

        if html_out.exists():
            with dcols[col_idx % 2]:
                st.download_button(
                    label="⬇️ Resumen HTML",
                    data=html_out.read_bytes(),
                    file_name=html_out.name,
                    mime="text/html",
                    use_container_width=True,
                )
                st.caption("Informe con gráfico interactivo")
            col_idx += 1

        if gantt_out and gantt_out.exists():
            with dcols[col_idx % 2]:
                st.download_button(
                    label="⬇️ Gantt 1H.xlsx",
                    data=gantt_out.read_bytes(),
                    file_name=gantt_out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption("Visual por bloque horario")
            col_idx += 1

        if map_out and map_out.exists():
            with dcols[col_idx % 2]:
                st.download_button(
                    label="⬇️ Sorter Map.xlsx",
                    data=map_out.read_bytes(),
                    file_name=map_out.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.caption("1 pestaña por día")
            col_idx += 1

        # Warnings summary
        warns = [l for l in log_lines if "Sin config" in l or "❌" in l]
        if warns:
            with st.expander("⚠️ Casos para revisión manual"):
                for w in warns:
                    st.warning(w)

st.divider()
st.caption("v0.01 · Mango Logística VDL B2B · Estrictamente confidencial")
