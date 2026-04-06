# Version: 0.07
import streamlit as st
import subprocess, sys, tempfile, datetime as dt
from pathlib import Path

BASE_DIR = Path(__file__).parent

# ── Validation panel ──────────────────────────────────────────────────────────
def _run_validation(parrilla_bytes, gd_bytes=None):
    """Run validate_parrilla.validate() safely; return issues list."""
    try:
        sys.path.insert(0, str(BASE_DIR))
        from validate_parrilla import validate
        return validate(parrilla_bytes, gd_bytes)
    except Exception as e:
        return [{"severity": "warning", "category": "estructura",
                 "title": "No se pudo ejecutar la validación previa",
                 "detail": str(e), "items": [], "autocorrected": False}]

def render_validation(parrilla_file, gd_file=None):
    """
    Run validation and render the results panel in Streamlit.
    Always returns True (never blocks generation).
    """
    par_bytes = parrilla_file.read(); parrilla_file.seek(0)
    gd_bytes  = None
    if gd_file:
        gd_bytes = gd_file.read(); gd_file.seek(0)

    issues = _run_validation(par_bytes, gd_bytes)

    # Count severities
    counts = {"error": 0, "warning": 0, "info": 0, "ok": 0}
    for iss in issues:
        counts[iss["severity"]] = counts.get(iss["severity"], 0) + 1

    # Header badge
    if counts["error"] > 0:
        badge = f"🔴 {counts['error']} error(s)"
        if counts["warning"]: badge += f"  ·  ⚠ {counts['warning']} aviso(s)"
    elif counts["warning"] > 0:
        badge = f"⚠ {counts['warning']} aviso(s)"
    else:
        badge = "✅ Todo OK"

    with st.expander(f"🔍 Validación previa — {badge}", expanded=(counts["error"] > 0 or counts["warning"] > 0)):
        # Group by category
        for cat_key, cat_label in [("estructura", "Estructura del fichero"),
                                    ("contenido",  "Contenido leído"),
                                    ("cobertura",  "Cobertura especiales en GD")]:
            cat_issues = [i for i in issues if i["category"] == cat_key]
            if not cat_issues:
                continue
            st.markdown(f"**{cat_label}**")
            for iss in cat_issues:
                sev = iss["severity"]
                icon = {"error": "🔴", "warning": "⚠️", "info": "ℹ️", "ok": "✅"}.get(sev, "·")
                title = iss["title"]
                if iss.get("autocorrected"):
                    title += "  *(autocorrección aplicada)*"

                if sev == "error":
                    st.error(f"{icon} **{title}**\n\n{iss['detail']}")
                elif sev == "warning":
                    st.warning(f"{icon} **{title}**\n\n{iss['detail']}")
                elif sev == "info":
                    st.info(f"{icon} **{title}**\n\n{iss['detail']}")
                else:
                    st.success(f"{icon} {title}")

                if iss.get("items"):
                    items_md = "  \n".join(f"- `{it}`" for it in iss["items"])
                    st.markdown(items_md)

            st.markdown("")  # spacing between categories

    return True  # never blocks

def _render_output_validation(parrilla_file, gd_output_bytes: bytes):
    """
    Post-generation validation: cross-check the produced GD against the parrilla.
    Shows results in an expander. Never blocks.
    """
    try:
        sys.path.insert(0, str(BASE_DIR))
        from validate_parrilla import validate_output, summary
        par_bytes = parrilla_file.read(); parrilla_file.seek(0)
        issues = validate_output(par_bytes, gd_output_bytes)
    except Exception as e:
        st.warning(f"⚠️ No se pudo ejecutar la validación del resultado: {e}")
        return

    counts = {"error": 0, "warning": 0, "info": 0, "ok": 0}
    for iss in issues:
        counts[iss["severity"]] = counts.get(iss["severity"], 0) + 1

    if counts["error"] > 0:
        badge = f"🔴 {counts['error']} error(s)"
        if counts["warning"]: badge += f"  ·  ⚠ {counts['warning']} aviso(s)"
        expanded = True
    elif counts["warning"] > 0:
        badge = f"⚠ {counts['warning']} aviso(s)"
        expanded = True
    else:
        badge = "✅ Sort map OK"
        expanded = False

    with st.expander(f"✅ Validación del resultado — {badge}", expanded=expanded):
        for iss in issues:
            sev  = iss["severity"]
            icon = {"error": "🔴", "warning": "⚠️", "info": "ℹ️", "ok": "✅"}.get(sev, "·")
            _is_tabla = "Resumen por dia" in iss["title"]
            if sev == "error":
                st.error(f"{icon} **{iss['title']}**\n\n{iss['detail']}")
            elif sev == "warning":
                if _is_tabla:
                    st.warning(f"{icon} **{iss['title']}**")
                    st.code(iss['detail'], language=None)
                else:
                    st.warning(f"{icon} **{iss['title']}**\n\n{iss['detail']}")
            elif sev == "info":
                st.info(f"{icon} **{iss['title']}**\n\n{iss['detail']}")
            else:
                if _is_tabla:
                    st.success(f"{icon} **{iss['title']}**")
                    st.code(iss['detail'], language=None)
                else:
                    st.success(f"{icon} {iss['title']}")
            if iss.get("items"):
                st.markdown("\n".join(f"- `{it}`" for it in iss["items"]))

def gd_to_dxc_csv(xlsx_bytes):
    # Convert GD xlsx to DXC upload CSV format (POSTEX + SOREXP separately)
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

# ── Session state ─────────────────────────────────────────────────────────────
for key in ["r1_gd","r1_esp","r1_can","r1_html","r2_gantt","r3_map","r1_day_filter","r1_postex_csv","r1_sorexp_csv","r1_esp_postex_csv","r1_esp_sorexp_csv"]:
    if key not in st.session_state:
        st.session_state[key] = None

# ── Inputs ────────────────────────────────────────────────────────────────────
st.markdown("### Ficheros de entrada")
col1, col2 = st.columns(2)
with col1:
    f_parrilla = st.file_uploader("Parrilla de salidas", type=["xlsx"],
                                   help="Debe incluir la hoja Resumen Bloques")

    # Dynamic sheet selector: read sheets from uploaded file
    if f_parrilla:
        import io as _io2
        from openpyxl import load_workbook as _lwb2
        _wb_tmp = _lwb2(_io2.BytesIO(f_parrilla.read()), read_only=True)
        f_parrilla.seek(0)
        # Filter to relevant sheets: those with TIPO_SALIDA column (parrilla sheets)
        _all_sheets = _wb_tmp.sheetnames
        _valid = []
        for _sh in _all_sheets:
            _ws_tmp = _wb_tmp[_sh]
            _first = next(_ws_tmp.iter_rows(values_only=True, max_row=1), None)
            if _first and any(str(h or "").strip().upper() in ("PLAYA","TIPO_SALIDA","DIA_PLAYA_NEW")
                              for h in _first):
                _valid.append(_sh)
        _options = _valid if _valid else _all_sheets
        # Pick best default: prefer parrilla_test_* sheets
        _default_idx = next(
            (i for i, s in enumerate(_options) if s.lower().startswith("parrilla_test")), 0
        )
        sheet = st.selectbox("Hoja de parrilla", options=_options, index=_default_idx,
                             help="Selecciona la pestaña con los datos de la semana")
    else:
        sheet = st.text_input("Nombre de hoja", value="parrilla_test_s14",
                              help="Sube la parrilla para ver las hojas disponibles")

    # Auto-detect semana from sheet name (parrilla_test_s18 → S18)
    import re as _re_sh
    _m_sem = _re_sh.search(r's([0-9]+)', sheet.lower())
    _sem_default = f"S{_m_sem.group(1).upper()}" if _m_sem else "S14"
    semana = st.text_input("Semana", value=_sem_default,
                           help="Se autodetecta del nombre de hoja")
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

# ── Action buttons — new flow: Sort Map first, then GD ───────────────────────
st.markdown("### Acciones")

base_ok   = bool(f_parrilla and f_gd and f_cap)
vis_ok    = bool(base_ok and f_bloques)
sortmap_done = bool(st.session_state.get("r3_map"))  # GD only available after sort map
gd_ok     = bool(sortmap_done and base_ok)             # GD requires sort map first

b1, b2, b3 = st.columns(3)
with b1:
    st.markdown("**1 · Sorter Map por día**")
    st.caption("Asigna especiales · valida visualmente")
    if st.button("🗺 Generar", key="go3", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r3_map"] = None
        st.session_state["r3_gd_bytes"] = None  # clear cached GD for sort map
        st.session_state["_run3"] = True
    if not vis_ok:
        st.caption("_Sube parrilla, GD, capacidad y bloques_")

with b2:
    st.markdown("**2 · Configuración DXC**")
    st.caption("Descarga el nuevo GD con los cambios")
    if st.button("⚙️ Generar", key="go1", type="primary",
                 disabled=not gd_ok, use_container_width=True):
        for k in ["r1_gd","r1_esp","r1_can","r1_html","r1_day_filter","r1_gd_filtered_bytes","r3_gd_bytes"]:
            st.session_state[k] = None
        st.session_state["_run1"] = True
    if not gd_ok:
        if not base_ok:
            st.caption("_Sube parrilla, GD y capacidad_")
        elif not sortmap_done:
            st.caption("_Genera el Sort Map primero_")

with b3:
    st.markdown("**3 · Gantt 1H**")
    st.caption("Visual bloques × rampas × hora")
    if st.button("📊 Generar", key="go2", type="primary",
                 disabled=not vis_ok, use_container_width=True):
        st.session_state["r2_gantt"] = None
        st.session_state["_run2"] = True
    if not vis_ok:
        st.caption("_Requiere también bloques horarios_")

st.divider()

# ── Execute action 1 ──────────────────────────────────────────────────────────
if st.session_state.get("_run1"):
    st.session_state["_run1"] = False
    sc = semana.strip() or sheet.strip().upper()
    # Validation (always runs, never blocks)
    render_validation(f_parrilla, f_gd)
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        # Reuse GD from sort map step if available — avoids regenerating
        _r3_gd = st.session_state.get("r3_gd_bytes")
        if _r3_gd:
            gd = tmp / f"GRUPO_DESTINOS_{sc}.xlsx"
            gd.write_bytes(_r3_gd)
            html = tmp / f"resumen_sorter_{sc}.html"
            with st.spinner("Generando resumen HTML…"):
                r = subprocess.run(
                    [sys.executable, str(BASE_DIR / "process_parrilla.py"),
                     str(p["parrilla"]), str(p["gd"]), str(p["cap"]),
                     sheet.strip(), sc, str(gd), str(html), ""],
                    capture_output=True, text=True, timeout=180)
        else:
            with st.spinner("Procesando parrilla y asignando rampas…"):
                gd, html, r = run_gd(p, tmp, sc)
        show_log(r)
        if r.returncode != 0 or not gd.exists():
            st.error("El proceso terminó con error.")
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
            st.session_state["r1_day_filter"] = None  # full run, no day filter
            # ── Post-generation sort map validation ──
            _render_output_validation(f_parrilla, _gd_bytes)

# ── Execute action 2 ──────────────────────────────────────────────────────────
if st.session_state.get("_run2"):
    st.session_state["_run2"] = False
    sc = semana.strip() or sheet.strip().upper()
    render_validation(f_parrilla, f_gd)
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
    render_validation(f_parrilla, f_gd)
    with tempfile.TemporaryDirectory() as _tmp:
        tmp = Path(_tmp)
        p   = save_uploads(tmp)
        # Use filtered GD if available, else generate full GD
        _filtered_bytes = st.session_state.get("r1_gd_filtered_bytes")
        _filter_days    = st.session_state.get("r1_day_filter")
        if _filtered_bytes:
            gd = tmp / f"GD_{sc}_filtered.xlsx"
            gd.write_bytes(_filtered_bytes)
            r0_ok = True
            st.info(f"Usando GD filtrado: {', '.join(_filter_days or [])}")
        else:
            with st.spinner("Generando GD base…"):
                gd, _, r0 = run_gd(p, tmp, sc)
            r0_ok = r0.returncode == 0 and gd.exists()
            if not r0_ok:
                st.error("Error generando GD base.")
                show_log(r0, expanded=True)
        if r0_ok:
            # Save GD bytes so action 1 can reuse without regenerating
            st.session_state["r3_gd_bytes"] = gd.read_bytes()
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
                    _gd_bytes_f = gd.read_bytes()
                    # Add day suffix to filenames so downloads are distinct from full run
                    _day_suffix = "_" + "+".join(selected_days)
                    _gd_name_f   = gd.stem + _day_suffix + ".xlsx"
                    _esp_name_f  = esp_path.stem + _day_suffix + ".xlsx" if esp_path.exists() else None
                    st.session_state["r1_gd"]   = (_gd_name_f, _gd_bytes_f)
                    st.session_state["r1_esp"]  = (_esp_name_f, esp_path.read_bytes()) if esp_path.exists() else None
                    # Regenerate DXC CSVs from filtered GD
                    _px_f, _sx_f = gd_to_dxc_csv(_gd_bytes_f)
                    st.session_state["r1_postex_csv"] = (gd.stem + _day_suffix + "_POSTEX.csv", _px_f)
                    st.session_state["r1_sorexp_csv"] = (gd.stem + _day_suffix + "_SOREXP.csv", _sx_f)
                    if esp_path.exists():
                        _epx_f, _esx_f = gd_to_dxc_csv(esp_path.read_bytes())
                        st.session_state["r1_esp_postex_csv"] = (esp_path.stem + _day_suffix + "_POSTEX.csv", _epx_f)
                        st.session_state["r1_esp_sorexp_csv"] = (esp_path.stem + _day_suffix + "_SOREXP.csv", _esx_f)
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

    # CSV DXC format downloads
    st.markdown("**Formato CSV para importar en DXC:**")
    c5, c6, c7, c8 = st.columns(4)
    with c5:
        if st.session_state["r1_postex_csv"]:
            name, data = st.session_state["r1_postex_csv"]
            st.download_button("⬇️ POSTEX completo", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c6:
        if st.session_state["r1_sorexp_csv"]:
            name, data = st.session_state["r1_sorexp_csv"]
            st.download_button("⬇️ SOREXP completo", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c7:
        if st.session_state["r1_esp_postex_csv"]:
            name, data = st.session_state["r1_esp_postex_csv"]
            st.download_button("⬇️ POSTEX especiales", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)
    with c8:
        if st.session_state["r1_esp_sorexp_csv"]:
            name, data = st.session_state["r1_esp_sorexp_csv"]
            st.download_button("⬇️ SOREXP especiales", data=data, file_name=name,
                               mime="text/csv", use_container_width=True)

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
st.caption("v0.07 · VDL B2B · Estrictamente confidencial")
