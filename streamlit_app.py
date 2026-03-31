# Version: 0.06 — rollback to v0.05 + instrucciones
import streamlit as st
import subprocess, sys, tempfile, datetime as dt
from pathlib import Path

BASE_DIR = Path(__file__).parent

def gd_to_dxc_csv(xlsx_bytes):
    import io as _io
    from openpyxl import load_workbook as _lwb
    wb = _lwb(_io.BytesIO(xlsx_bytes), read_only=True)
    rows = list(wb.active.iter_rows(values_only=True))[1:]
    postex_lines, sorexp_lines = [], []
    for r in rows:
        if len(r) < 7: continue
        desc = str(r[2] or "").strip()
        tipo = str(r[3] or "").strip().upper()
        dest_raw = str(r[4] or "").strip()
        elem = str(r[6] or "").strip()
        if not desc or tipo not in ("POSTEX","SOREXP") or not dest_raw or not elem: continue
        if desc.startswith("="): continue
        d = "".join(c for c in dest_raw if c.isdigit())
        dest8 = d.zfill(10)[-8:] if d else dest_raw[:8]
        line = f"{desc};{dest8};00;{elem}"
        if tipo == "POSTEX": postex_lines.append(line + ";20")
        else: sorexp_lines.append(line + ";10")
    bom = "\ufeff"
    def _enc(lines): return (bom + "\r\n".join(lines)).encode("utf-8")
    return _enc(postex_lines), _enc(sorexp_lines)


st.set_page_config(page_title="Sorter VDL B2B", page_icon="🏭", layout="centered")
st.markdown("<style>.block-container{max-width:780px}</style>", unsafe_allow_html=True)

st.markdown("## 🏭 Sorter VDL B2B")
st.markdown("Configurador de semanas especiales — VDL B2B")
st.divider()

# ── Instrucciones ─────────────────────────────────────────────────────────────
st.markdown("### Que hace esta herramienta")
st.markdown("""
Genera la configuración del sorter VDL B2B para semanas con salidas canceladas o que cambian de día
(festivos, Semana Santa, etc.). A partir de la parrilla semanal y el GRUPO_DESTINOS actual, calcula
qué destinos reasignar, encuentra rampas libres respetando los bloques horarios, y produce los ficheros
listos para subir a DXC/MAR.
""")

col_a, col_b, col_c = st.columns(3)
with col_a:
    st.markdown("**1 · Sube los ficheros**")
    st.caption("Parrilla semanal, GRUPO_DESTINOS actual y capacidad de rampas. Los bloques horarios son necesarios solo para Gantt y Sorter Map.")
with col_b:
    st.markdown("**2 · Genera la configuracion**")
    st.caption("Pulsa *Configuracion DXC* y la herramienta procesa las salidas, reasigna rampas y genera el fichero.")
with col_c:
    st.markdown("**3 · Descarga y sube a DXC**")
    st.caption("Descarga el GRUPO_DESTINOS generado y el resumen HTML con el analisis de ocupacion por bloque.")

st.divider()
st.markdown("**Ficheros de entrada**")

col_fi1, col_fi2 = st.columns(2)
with col_fi1:
    st.markdown("""
| Fichero | Oblig. |
|---|:---:|
| `parrilla_de_salidas.xlsx` | ✅ |
| `GRUPO_DESTINOS.xlsx` (consulta DXC 9066) | ✅ |
| `ramp_capacity.csv` | ✅ |
| `bloques_horarios.xlsx` | ⚠️ Gantt/Map |
| `superplayas.xlsx` | ➖ opcional |
""")
with col_fi2:
    st.markdown("""
**Tipos de salida en la parrilla**

| Tipo | Accion |
|---|---|
| `HABITUAL` | Sin cambios |
| `CANCELADA` | Se elimina del GD |
| `ESPECIAL DIA CAMBIO` | Se reasigna a rampas libres |
| `IRREGULAR` | Ignorada |
""")

st.divider()

# ── Session state ─────────────────────────────────────────────────────────────
for key in ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map","r1_day_filter","r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Inputs ────────────────────────────────────────────────────────────────────
st.markdown("### Ficheros de entrada")
col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader("Parrilla de salidas", type=["xlsx"],
                                   help="parrilla_de_salidas.xlsx — debe incluir la hoja Resumen Bloques")
    sheet  = st.text_input("Nombre de hoja", value="parrilla_test_s14")
    semana = st.text_input("Semana", value="S14")
with col2:
    f_gd  = st.file_uploader("GRUPO_DESTINOS", type=["xlsx"],
                               help="Consulta personalizada DXC 9066 — Consulta destinos por zona")
    f_cap = st.file_uploader("Capacidad de rampas", type=["csv"],
                               help="CSV con columnas RAMP;PALLETS")

f_bloques = st.file_uploader(
    "Bloques horarios *(necesario para Gantt y Sorter Map)*", type=["xlsx"],
    help="Columnas: NUEVO BLOQUE · Dia LIBERACION · Hora LIBERACION · Dia DESACTIVACION · Hora DESACTIVACION")

f_superplaya = st.file_uploader(
    "Superplayas *(opcional)*", type=["xlsx"],
    help="Columnas: AGRUPACION_PLAYA · SUPERPLAYA — define que destinos deben ir juntos en rampas contiguas")

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
    cmd.append(days_arg)
    if "superplaya" in p:
        cmd.append(str(p["superplaya"]))
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
    st.markdown("**1 · Configuracion DXC**")
    st.caption("GRUPO_DESTINOS + resumen HTML")
    if st.button("Generar", key="go1", type="primary",
                 disabled=not base_ok, use_container_width=True):
        for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter"]:
            st.session_state[k] = None
        st.session_state["_run1"] = True
    if not base_ok:
        st.caption("_Sube parrilla, GD y capacidad_")

with b2:
    st.markdown("**2 · Gantt 1H**")
    st.caption("Visual bloques x rampas x hora")
    if st.button("Generar", key="go2", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r2_gantt"] = None
        st.session_state["_run2"] = True
    if not vis_ok:
        st.caption("_Requiere tambien bloques horarios_")

with b3:
    st.markdown("**3 · Sorter Map por dia**")
    st.caption("1 pestana por dia, slots fisicos")
    if st.button("Generar", key="go3", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r3_map"] = None
        st.session_state["_run3"] = True
    if not vis_ok:
        st.caption("_Requiere tambien bloques horarios_")

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
            st.error("El proceso termino con error.")
        else:
            esp_path = Path(str(gd).replace('.xlsx', '_SOLO_ESPECIALES.xlsx'))
            can_path = Path(str(gd).replace('.xlsx', '_CANCELADAS.txt'))
            _gd_bytes = gd.read_bytes()
            st.session_state["r1_gd"]    = (gd.name, _gd_bytes)
            st.session_state["r1_esp"]   = (esp_path.name, esp_path.read_bytes()) if esp_path.exists() else None
            _px, _sx = gd_to_dxc_csv(_gd_bytes)
            st.session_state["r1_postex_csv"] = (gd.stem + "_POSTEX.csv", _px)
            st.session_state["r1_sorexp_csv"] = (gd.stem + "_SOREXP.csv", _sx)
            if esp_path.exists():
                _epx, _esx = gd_to_dxc_csv(esp_path.read_bytes())
                st.session_state["r1_esp_postex_csv"] = (esp_path.stem + "_POSTEX.csv", _epx)
                st.session_state["r1_esp_sorexp_csv"] = (esp_path.stem + "_SOREXP.csv", _esx)
            st.session_state["r1_can"]   = (can_path.name, can_path.read_text(encoding='utf-8')) if can_path.exists() else None
            st.session_state["r1_html"]  = (html.name,  html.read_bytes()) if html.exists() else None
            st.session_state["r1_day_filter"] = None

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
                st.error("El Gantt termino con error.")

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
                _sorter_cmd = [
                    sys.executable, str(BASE_DIR / "sorter_map_por_dia.py"),
                    str(p["cap"]), str(gd), str(p["bloques"]), str(out), "Hoja1",
                ]
                if "parrilla" in p and "gd" in p:
                    _sorter_cmd += [str(p["parrilla"]), sheet.strip(), str(p["gd"])]
                r = subprocess.run(_sorter_cmd, capture_output=True, text=True, timeout=180)
            show_log(r)
            if r.returncode == 0 and out.exists():
                st.session_state["r3_map"] = (out.name, out.read_bytes())
            else:
                st.error("El Sorter Map termino con error.")

# ── Show results action 1 ─────────────────────────────────────────────────────
if st.session_state["r1_gd"] is not None:
    gd_name, gd_bytes = st.session_state["r1_gd"]
    sc = semana.strip() or sheet.strip().upper()
    st.success(f"Configuracion {sc} generada")

    with st.expander("Filtrar por dia y regenerar especiales"):
        selected_days = st.multiselect(
            "Dias a incluir en el GD filtrado",
            options=ALL_DAYS,
            default=[],
            key="day_filter_sel",
            placeholder="Selecciona dias…",
        )
        if st.button("Regenerar con filtro", key="regen_filter",
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
                    st.error("Error en regeneracion.")
                    show_log(r, expanded=True)

    if st.session_state["r1_day_filter"]:
        st.info(f"Filtrado a: {', '.join(st.session_state['r1_day_filter'])}")

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
            st.caption("Solo filas nuevas a añadir en DXC")

    st.markdown("**Formato CSV para importar en DXC:**")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        if st.session_state["r1_postex_csv"]:
            name, data = st.session_state["r1_postex_csv"]
            st.download_button("POSTEX completo", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c6:
        if st.session_state["r1_sorexp_csv"]:
            name, data = st.session_state["r1_sorexp_csv"]
            st.download_button("SOREXP completo", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c7:
        if st.session_state["r1_esp_postex_csv"]:
            name, data = st.session_state["r1_esp_postex_csv"]
            st.download_button("POSTEX especiales", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c8:
        if st.session_state["r1_esp_sorexp_csv"]:
            name, data = st.session_state["r1_esp_sorexp_csv"]
            st.download_button("SOREXP especiales", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)

    c3, c4 = st.columns(2)
    with c3:
        if st.session_state["r1_can"]:
            name, txt = st.session_state["r1_can"]
            st.download_button("Canceladas.txt", data=txt, file_name=name,
                               mime="text/plain", use_container_width=True)
            with st.expander("Ver canceladas"):
                st.text(txt)
            st.caption("Salidas a eliminar del sorter")
    with c4:
        if st.session_state["r1_html"]:
            name, data = st.session_state["r1_html"]
            st.download_button("Resumen HTML", data=data, file_name=name,
                               mime="text/html", use_container_width=True)
            st.caption("Informe con grafico interactivo")

# ── Show results action 2 ─────────────────────────────────────────────────────
if st.session_state["r2_gantt"] is not None:
    name, data = st.session_state["r2_gantt"]
    st.success("Gantt 1H generado")
    st.download_button("Gantt 1H.xlsx", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: LEYENDA · BLOQUES_DESTINOS · GANTT_VISUAL · GANTT_OPERATIVO")

# ── Show results action 3 ─────────────────────────────────────────────────────
if st.session_state["r3_map"] is not None:
    name, data = st.session_state["r3_map"]
    st.success("Sorter Map generado")
    st.download_button("Sorter Map.xlsx", data=data, file_name=name,
                       mime=XLSX_MIME, use_container_width=True)
    st.caption("Hojas: DOM · LUN · MAR · MIE · JUE · VIE · SAB · LEYENDA")

st.divider()
st.caption("v0.06 · VDL B2B · Estrictamente confidencial")
