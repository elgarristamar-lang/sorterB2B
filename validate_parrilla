# validate_parrilla.py — Version 0.01
# Validación de parrilla + GD antes de generar.
# Devuelve una lista de issues estructurados; no bloquea la generación.
#
# Uso:
#   from validate_parrilla import validate
#   issues = validate(parrilla_bytes, gd_bytes)

from __future__ import annotations
import io
import re
from collections import defaultdict
from typing import List, Dict, Tuple, Optional

# ── Tipos de issue ────────────────────────────────────────────────────────────
# severity: "error" | "warning" | "info" | "ok"
# category: "estructura" | "cobertura" | "contenido"
# autocorrected: bool  — el código lo maneja igualmente, pero avisa
# items: list[str]     — detalles concretos (nombres de playas, columnas, etc.)

def _issue(severity, category, title, detail, items=None, autocorrected=False):
    return {
        "severity": severity,
        "category": category,
        "title": title,
        "detail": detail,
        "items": items or [],
        "autocorrected": autocorrected,
    }


# ── Helpers ───────────────────────────────────────────────────────────────────

_DAYS = {"DOMINGO", "LUNES", "MARTES", "MIERCOLES", "MIÉRCOLES",
         "JUEVES", "VIERNES", "SABADO"}
_DAY_PFX = re.compile(
    r"^(DOMINGO|LUNES|MARTES|MIERCOLES|MIÉRCOLES|JUEVES|VIERNES|SABADO)_(.+)$",
    re.IGNORECASE,
)
_BLOQUE_RE = re.compile(r"^(\d+BLO[A-Z]\d+)_(.+)$")
_SORTER_RE = re.compile(r"^R\d+")
_E2_RE     = re.compile(r"^(MAN|EXDOCK|DOCK)", re.IGNORECASE)


def _extract_playa(dpn_val: str) -> str:
    """
    Extract playa name from DIA_PLAYA_NEW value.
    Handles:
      - LUNES_ESPANA_GUARROMAN          → ESPANA_GUARROMAN
      - CANCELADA_ESPANA_LAS_PALMAS     → ESPANA_LAS_PALMAS
      - CANCELADO_AUSTRIA_...           → AUSTRIA_...
      - ESPANA_GUARROMAN (no prefix)    → ESPANA_GUARROMAN
    """
    s = (dpn_val or "").strip()
    su = s.upper()
    # CANCELADA_ / CANCELADO_ prefix
    if su.startswith("CANCELADA_") or su.startswith("CANCELADO_"):
        return s[s.index("_") + 1:].strip()
    # DAY_PLAYA prefix
    m = _DAY_PFX.match(s)
    if m:
        return m.group(2).strip()
    return s


def _parse_gd_desc(desc: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Returns (bloque, dia, playa) or (None, None, None)."""
    if not desc:
        return None, None, None
    core = re.sub(r"^\[B2B\]\s*", "", str(desc)).strip()
    core = re.sub(r"\s+PARA BAJAR POR.*$", "", core)
    core = re.sub(r"_CANCELADA.*$", "", core)
    m = _BLOQUE_RE.match(core)
    if not m:
        return None, None, None
    bloque, rest = m.group(1), m.group(2)
    for d in _DAYS:
        if rest.upper().startswith(d + "_"):
            return bloque, d.replace("MIÉRCOLES", "MIERCOLES"), rest[len(d) + 1:]
    return bloque, None, rest


# ── Loaders (in-memory, no disk) ──────────────────────────────────────────────

def _load_parrilla(xlsx_bytes: bytes) -> Dict:
    """
    Returns dict with:
      sheets, col_issues, records, stats
    records: list of {playa, dia_orig, dia_new, tipo, dpn_raw}
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    sheets = wb.sheetnames

    col_issues = []
    records = []
    stats = {"regular": 0, "cancelada": 0, "especial": 0, "irregular": 0, "other": 0}

    # Find the best sheet: one that contains TIPO_SALIDA
    target_sheet = None
    for sh in sheets:
        ws = wb[sh]
        hdr = next(ws.iter_rows(values_only=True, max_row=1), ())
        if any(str(h or "").strip().upper() == "TIPO_SALIDA" for h in hdr):
            target_sheet = sh
            break

    if not target_sheet:
        col_issues.append(_issue(
            "error", "estructura",
            "No se encuentra ninguna hoja con columna TIPO_SALIDA",
            f"Hojas disponibles: {', '.join(sheets)}. "
            "La parrilla no tiene el formato esperado — no se podrá procesar.",
        ))
        return {"sheets": sheets, "col_issues": col_issues, "records": [], "stats": stats,
                "target_sheet": None, "has_playa_col": False, "has_bloques_sheet": False,
                "cancelado_masc": []}

    ws = wb[target_sheet]
    all_rows = list(ws.iter_rows(values_only=True))
    hdr = all_rows[0]
    col = {str(h or "").strip().upper(): i for i, h in enumerate(hdr) if h}

    # Check PLAYA column
    has_playa = "PLAYA" in col or "AGRUPACION_PLAYA" in col
    has_dpn   = "DIA_PLAYA_NEW" in col

    if not has_playa:
        if has_dpn:
            col_issues.append(_issue(
                "warning", "estructura",
                "Columna PLAYA no encontrada — se usará DIA_PLAYA_NEW",
                "Se esperaba columna PLAYA o AGRUPACION_PLAYA. "
                "El nombre de playa se extraerá automáticamente del valor de DIA_PLAYA_NEW "
                "(ej. LUNES_ESPANA_GUARROMAN → ESPANA_GUARROMAN).",
                autocorrected=True,
            ))
        else:
            col_issues.append(_issue(
                "error", "estructura",
                "No hay columna PLAYA ni DIA_PLAYA_NEW",
                f"Columnas encontradas: {', '.join(col.keys())}. "
                "No es posible identificar los destinos.",
            ))

    # Check BLOQUE column
    if "BLOQUE" not in col:
        col_issues.append(_issue(
            "warning", "estructura",
            "Columna BLOQUE no encontrada",
            "Sin esta columna, el bloque horario de las especiales se derivará "
            "del ID_CLUSTER — puede ser menos preciso.",
            autocorrected=True,
        ))

    # Check Resumen Bloques / bloques sheet
    has_bloques_sheet = "Resumen Bloques" in sheets
    bloques_alt = None
    if not has_bloques_sheet:
        # Look for any sheet with bloque/cluster columns
        for sh in sheets:
            if sh == target_sheet:
                continue
            ws_b = wb[sh]
            hdr_b = next(ws_b.iter_rows(values_only=True, max_row=1), ())
            hdr_b_up = [str(h or "").strip().upper() for h in hdr_b]
            if any("BLOQUE" in h or "CLUSTER" in h for h in hdr_b_up):
                bloques_alt = sh
                break
        if bloques_alt:
            col_issues.append(_issue(
                "warning", "estructura",
                f"No hay hoja 'Resumen Bloques' — se usará '{bloques_alt}'",
                "Los timings horarios de bloques se leerán desde esta hoja alternativa. "
                "Comprueba que tiene las columnas correctas.",
                autocorrected=True,
            ))
        else:
            col_issues.append(_issue(
                "error", "estructura",
                "No hay hoja 'Resumen Bloques' ni equivalente",
                "Sin los timings de bloques no se pueden detectar colisiones horarias "
                "en la asignación de rampas.",
            ))

    # Check SEMANA SANTA sheet (informative — S15 format doesn't need it)
    has_ss_sheet = any(
        "SEMANA" in s.upper() or "SANTA" in s.upper() for s in sheets
    )
    if not has_ss_sheet and target_sheet:
        col_issues.append(_issue(
            "info", "estructura",
            "No hay hoja 'SEMANA SANTA' — formato unificado detectado",
            f"Los datos de canceladas y especiales se leerán de la hoja '{target_sheet}'. "
            "Formato S15/unificado reconocido.",
            autocorrected=True,
        ))

    # Parse records
    tipo_idx = col.get("TIPO_SALIDA")
    dpn_idx  = col.get("DIA_PLAYA_NEW") or col.get("DIA_PLAYA")
    dpo_idx  = col.get("DIA_PLAYA_ORIGINAL")
    dsn_idx  = col.get("DIA_SALIDA_NEW")
    dso_idx  = col.get("DIA_SALIDA_ORIGINAL") or col.get("DIA_SALIDA")
    playa_idx = col.get("PLAYA") or col.get("AGRUPACION_PLAYA")
    bloque_idx = col.get("BLOQUE")
    idc_idx  = col.get("ID_CLUSTER")
    idcn_idx = col.get("ID_CLUSTER_NEW")

    cancelado_masc = []  # rows with CANCELADO_ (masculine) prefix

    for row in all_rows[1:]:
        def g(i):
            if i is None or i >= len(row) or row[i] is None:
                return ""
            s = str(row[i]).strip()
            return "" if s.startswith("=") or s == "#N/A" else s

        tipo   = g(tipo_idx).upper()
        dpn    = g(dpn_idx)   # DIA_PLAYA_NEW  → día nuevo + playa
        dpo    = g(dpo_idx)   # DIA_PLAYA_ORIGINAL → día original + playa
        dia_new  = g(dsn_idx).upper()
        dia_orig = g(dso_idx).upper()

        # Extract playa name.
        # For ESPECIAL DIA CAMBIO: use DIA_PLAYA_ORIGINAL as the authoritative source
        # because the GD is keyed on the ORIGINAL day+playa, not the new one.
        # CANCELADA: DIA_PLAYA_NEW already has CANCELADA_PLAYA format.
        if playa_idx is not None and g(playa_idx):
            playa = g(playa_idx)
        elif tipo == "ESPECIAL DIA CAMBIO" and dpo:
            # DIA_PLAYA_ORIGINAL = "MARTES_ESPANA_GUARROMAN" → playa = ESPANA_GUARROMAN
            playa = _extract_playa(dpo)
            # Also extract dia_orig from DIA_PLAYA_ORIGINAL if DIA_SALIDA_ORIGINAL is missing
            if not dia_orig:
                m = _DAY_PFX.match(dpo)
                if m:
                    dia_orig = m.group(1).upper()
        elif dpn:
            playa = _extract_playa(dpn)
            # Detect masculine CANCELADO_ prefix
            if dpn.upper().startswith("CANCELADO_"):
                cancelado_masc.append(dpn)
        else:
            playa = ""

        # For CANCELADA: also extract dia_orig from DIA_PLAYA_ORIGINAL if missing
        if tipo == "CANCELADA" and not dia_orig and dpo:
            m = _DAY_PFX.match(dpo)
            if m:
                dia_orig = m.group(1).upper()

        if not playa or not tipo:
            continue

        # Normalise tipo → kind
        if tipo in ("REGULAR", "HABITUAL", "ESPECIAL CUTOFF"):
            stats["regular"] += 1
            kind = "regular"
        elif tipo == "CANCELADA":
            stats["cancelada"] += 1
            kind = "cancelada"
        elif tipo == "ESPECIAL DIA CAMBIO":
            stats["especial"] += 1
            kind = "especial"
        elif tipo == "IRREGULAR":
            stats["irregular"] += 1
            kind = "irregular"
        else:
            stats["other"] += 1
            kind = "other"

        records.append({
            "playa": playa.upper(),
            "dia_orig": dia_orig,
            "dia_new": dia_new,
            "tipo": kind,
            "dpn_raw": dpn,
            "bloque": g(bloque_idx),
            "id_cluster": g(idc_idx),
            "id_cluster_new": g(idcn_idx),
        })

    if cancelado_masc:
        col_issues.append(_issue(
            "warning", "contenido",
            f"Prefijo CANCELADO_ (masculino) detectado en {len(cancelado_masc)} fila(s)",
            "Se usarán exactamente igual que CANCELADA_.",
            items=[dpn for dpn in cancelado_masc[:10]],
            autocorrected=True,
        ))

    return {
        "sheets": sheets,
        "target_sheet": target_sheet,
        "col_issues": col_issues,
        "records": records,
        "stats": stats,
        "has_playa_col": has_playa,
        "has_bloques_sheet": has_bloques_sheet or bool(bloques_alt),
        "cancelado_masc": cancelado_masc,
    }


def _load_gd_playas(xlsx_bytes: bytes) -> Dict:
    """
    Returns {(dia, playa): [elementos]} for all POSTEX sorter entries in GD.
    Also returns set of all playa names found.
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h or "").strip() for h in all_rows[0]]
    hdrs = {h.upper(): i for i, h in enumerate(hdr)}
    is_dxc = "Estado" in hdr or "Secuencia" in hdr

    if is_dxc:
        idx_desc  = hdrs.get("DESCRIPCIÓN GRUPOS DE DESTINO", hdrs.get("DESCRIPCION GRUPOS DE DESTINO"))
        idx_zona  = hdrs.get("TIPO DE ZONA")
        idx_elem  = hdrs.get("ELEMENTO")
    else:
        idx_desc, idx_zona, idx_elem = 2, 3, 6

    by_dia_playa = defaultdict(list)
    all_playas   = set()

    for row in all_rows[1:]:
        def g(i):
            if i is None or i >= len(row) or row[i] is None:
                return ""
            s = str(row[i]).strip()
            return "" if s.startswith("=") else s

        desc = g(idx_desc)
        zona = g(idx_zona).upper()
        elem = g(idx_elem)

        if zona != "POSTEX" or not _SORTER_RE.match(elem):
            continue

        bloque, dia, playa = _parse_gd_desc(desc)
        if not (bloque and dia and playa):
            continue

        all_playas.add(playa.upper())
        by_dia_playa[(dia.upper(), playa.upper())].append(elem)

    return {"by_dia_playa": dict(by_dia_playa), "all_playas": all_playas}


# ── Main validation ───────────────────────────────────────────────────────────

def validate(parrilla_bytes: bytes, gd_bytes: Optional[bytes] = None) -> List[Dict]:
    """
    Run all validations. Returns list of issue dicts.
    GD is optional — without it, cobertura checks are skipped.
    """
    issues = []

    # ── 1. Parse parrilla ─────────────────────────────────────────────────────
    par = _load_parrilla(parrilla_bytes)
    issues.extend(par["col_issues"])

    records  = par["records"]
    especiales = [r for r in records if r["tipo"] == "especial"]
    canceladas = [r for r in records if r["tipo"] == "cancelada"]

    # ── 2. Contenido: counts ──────────────────────────────────────────────────
    if not records:
        issues.append(_issue(
            "error", "contenido",
            "No se encontraron registros válidos en la parrilla",
            "Revisa que la hoja correcta está seleccionada y que las columnas "
            "TIPO_SALIDA y la columna de playa existen.",
        ))
        return issues

    issues.append(_issue(
        "ok", "contenido",
        f"{len(especiales)} especiales · {len(canceladas)} canceladas leídas",
        f"Total registros procesables: {len(records)} "
        f"(+{par['stats']['irregular']} irregulares ignoradas).",
    ))

    # Warn if zero especiales and zero canceladas — likely a read problem
    if len(especiales) == 0 and len(canceladas) == 0 and len(records) > 0:
        issues.append(_issue(
            "error", "contenido",
            "No se detectó ninguna especial ni cancelada",
            "Hay registros en la parrilla pero ninguno es ESPECIAL DIA CAMBIO ni CANCELADA. "
            "Es posible que la columna TIPO_SALIDA tenga valores distintos a los esperados.",
            items=list({r["tipo"] for r in records})[:10],
        ))

    # Especiales without new day
    esp_sin_dia = [r for r in especiales if not r["dia_new"]]
    if esp_sin_dia:
        issues.append(_issue(
            "warning", "contenido",
            f"{len(esp_sin_dia)} especiales sin DIA_SALIDA_NEW",
            "Estas especiales no tienen día de destino definido — se intentará derivar "
            "del ID_CLUSTER pero puede fallar.",
            items=[r["playa"] for r in esp_sin_dia[:10]],
        ))

    # Especiales without bloque and without ID_CLUSTER_NEW
    esp_sin_bloque = [
        r for r in especiales
        if not r["bloque"] or r["bloque"] in ("#N/A", "NO_BLOQUE", "")
        and not r["id_cluster_new"]
    ]
    if esp_sin_bloque:
        issues.append(_issue(
            "warning", "contenido",
            f"{len(esp_sin_bloque)} especiales sin BLOQUE ni ID_CLUSTER_NEW",
            "No se podrá determinar el bloque horario del nuevo día — "
            "la asignación de rampas podría ser incorrecta.",
            items=[r["playa"] for r in esp_sin_bloque[:10]],
        ))

    # ── 3. Cobertura: especiales vs GD ────────────────────────────────────────
    if gd_bytes is None:
        issues.append(_issue(
            "info", "cobertura",
            "GD no subido — validación de cobertura omitida",
            "Sube el fichero GRUPO_DESTINOS para verificar que todas las especiales "
            "tienen config en su día ORIGINAL (el script la usa para saber cuántas "
            "posiciones asignar en el nuevo día).",
        ))
        return issues

    gd = _load_gd_playas(gd_bytes)
    by_dia_playa = gd["by_dia_playa"]
    gd_playas    = gd["all_playas"]

    # Para cada especial, verificar que existe config en el GD del DÍA ORIGINAL.
    # El GD del día nuevo NO existe todavía — la aplicación lo crea.
    # Si no hay config en el día original, el script intenta fallback a otro día;
    # si no hay config en ningún día, no puede asignar posiciones.
    sin_config_ninguno = []   # playa no existe en GD en ningún día
    sin_config_orig    = []   # playa existe en GD pero no en el día original (usará fallback)
    ok_count = 0

    seen = set()
    for r in especiales:
        playa    = r["playa"]
        dia_orig = r["dia_orig"]
        if playa in seen:
            continue
        seen.add(playa)

        if playa not in gd_playas:
            sin_config_ninguno.append(r)
        elif dia_orig and (dia_orig, playa) not in by_dia_playa:
            other_days = sorted({d for (d, p) in by_dia_playa if p == playa})
            sin_config_orig.append({**r, "found_on_days": other_days})
        else:
            ok_count += 1

    if ok_count > 0:
        issues.append(_issue(
            "ok", "cobertura",
            f"{ok_count} especiales con config GD en el día original",
            "El script leerá sus posiciones actuales y las reasignará al nuevo día.",
        ))

    if sin_config_orig:
        items_detail = [
            f"{r['playa']}  (día original: {r['dia_orig'] or '?'} → encontrada en: {', '.join(r['found_on_days'][:3])})"
            for r in sin_config_orig[:10]
        ]
        issues.append(_issue(
            "warning", "cobertura",
            f"{len(sin_config_orig)} especiales sin config en el día original — se usará fallback",
            "Existen en el GD pero no en el día de origen de la parrilla. "
            "El script buscará el día con más posiciones y lo usará como fuente — "
            "revisa que el día original en la parrilla es correcto.",
            items=items_detail,
        ))

    if sin_config_ninguno:
        issues.append(_issue(
            "error", "cobertura",
            f"{len(sin_config_ninguno)} especiales SIN config en el GD en ningún día",
            "Estas playas aparecen como ESPECIAL DIA CAMBIO en la parrilla pero no tienen "
            "ninguna entrada POSTEX en el GD origen. Sin esa base no se pueden calcular "
            "qué posiciones asignar en el nuevo día — quedarán sin sort map.",
            items=[r["playa"] for r in sin_config_ninguno],
        ))

    # Check canceladas: warn if any cancelada IS in GD (REVISAR cases)
    canceladas_en_gd = []
    seen_c = set()
    for r in canceladas:
        playa = r["playa"]
        if playa in seen_c:
            continue
        seen_c.add(playa)
        if playa in gd_playas:
            canceladas_en_gd.append(r)

    if canceladas_en_gd:
        issues.append(_issue(
            "warning", "cobertura",
            f"{len(canceladas_en_gd)} canceladas que tienen config en GD (⚠ REVISAR)",
            "Estas playas están marcadas como CANCELADA en la parrilla pero tienen "
            "entradas en el GD origen. Se eliminarán del GD generado — confirma que es correcto.",
            items=[r["playa"] for r in canceladas_en_gd[:15]],
        ))

    # ── Zona consistency: parrilla E2 vs GD sorter elements ──────────────────
    issues.extend(validate_zona_consistency(parrilla_bytes, gd_bytes))

    return issues


# ── Severity summary ──────────────────────────────────────────────────────────

def summary(issues: List[Dict]) -> Dict:
    """Returns {errors, warnings, infos, oks} counts."""
    counts = {"error": 0, "warning": 0, "info": 0, "ok": 0}
    for iss in issues:
        counts[iss["severity"]] = counts.get(iss["severity"], 0) + 1
    return counts

# ── Post-generation sort map validation ───────────────────────────────────────

def _load_gd_output(xlsx_bytes: bytes) -> Dict:
    """
    Parse the GENERATED GD xlsx (output of process_parrilla.py).
    Returns:
      playas_por_dia: {dia: set of playa names} — all POSTEX sorter entries
      all_playas: set of all playa names across all days
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return {"playas_por_dia": {}, "all_playas": set()}

    hdr = [str(h or "").strip() for h in all_rows[0]]
    hdrs = {h.upper(): i for i, h in enumerate(hdr)}
    is_dxc = "Estado" in hdr or "Secuencia" in hdr

    if is_dxc:
        idx_desc = hdrs.get("DESCRIPCIÓN GRUPOS DE DESTINO",
                            hdrs.get("DESCRIPCION GRUPOS DE DESTINO"))
        idx_zona = hdrs.get("TIPO DE ZONA")
        idx_elem = hdrs.get("ELEMENTO")
    else:
        idx_desc, idx_zona, idx_elem = 2, 3, 6

    playas_por_dia: Dict[str, set] = defaultdict(set)
    all_playas: set = set()

    for row in all_rows[1:]:
        def g(i):
            if i is None or i >= len(row) or row[i] is None:
                return ""
            s = str(row[i]).strip()
            return "" if s.startswith("=") else s

        desc = g(idx_desc)
        zona = g(idx_zona).upper()
        elem = g(idx_elem)

        if zona != "POSTEX" or not _SORTER_RE.match(elem):
            continue

        # _CANCELADA_SOLO_W* entries are leftover positions from previous weeks.
        # Treat them as normal occupied slots — strip the suffix so playa name
        # parses correctly, but keep the entry in the map.
        if re.search(r"_CANCELADA_SOLO_W[^_\s]*", desc, re.IGNORECASE):
            desc = re.sub(r"_CANCELADA_SOLO_W[^_\s)]*", "", desc)

        bloque, dia, playa = _parse_gd_desc(desc)
        if not (dia and playa):
            continue

        dia = dia.upper()
        playa = playa.upper()
        # Strip (ESPECIAL...) suffix added by process_parrilla for new-day entries
        playa = re.sub(r"\s*\(ESPECIAL.*\)$", "", playa).strip()

        playas_por_dia[dia].add(playa)
        all_playas.add(playa)

    return {"playas_por_dia": dict(playas_por_dia), "all_playas": all_playas}


def validate_output(parrilla_bytes: bytes, gd_output_bytes: bytes) -> List[Dict]:
    """
    Post-generation validation: cross-check the generated GD against the parrilla.

    Checks:
      1. Especiales → must appear in sort map on the NEW day
      2. Canceladas → must NOT appear in sort map on ANY day
      3. Especiales → must NOT appear in sort map on the ORIGINAL day (moved away)

    Returns list of issue dicts (same format as validate()).
    """
    issues = []

    # Parse parrilla (reuse existing loader)
    par = _load_parrilla(parrilla_bytes)
    records  = par["records"]
    especiales = [r for r in records if r["tipo"] == "especial"]
    canceladas = [r for r in records if r["tipo"] == "cancelada"]

    if not records:
        issues.append(_issue(
            "error", "resultado",
            "No se pudieron leer registros de la parrilla para validar",
            "Revisa que el fichero de parrilla es correcto.",
        ))
        return issues

    # Parse generated GD output
    out = _load_gd_output(gd_output_bytes)
    playas_por_dia = out["playas_por_dia"]
    all_playas_out = out["all_playas"]

    # ── Check 1: Especiales deben estar en el día NUEVO ───────────────────────
    esp_ok, esp_falta = [], []
    seen = set()
    for r in especiales:
        playa   = r["playa"]
        dia_new = r["dia_new"]
        if playa in seen or not dia_new:
            continue
        seen.add(playa)
        playas_en_dia_new = playas_por_dia.get(dia_new, set())
        if playa in playas_en_dia_new:
            esp_ok.append(r)
        else:
            esp_falta.append(r)

    if esp_ok:
        issues.append(_issue(
            "ok", "resultado",
            f"{len(esp_ok)} especiales presentes en el sort map en el día nuevo",
            "Estas playas aparecen correctamente en el día de destino.",
        ))

    if esp_falta:
        # Separate E2/manual routes (expected to be absent) from real missing ones
        _E2_KNOWN = {"BOSNIA_CPT","CHIPRE_NORTE","INDONESIA","CHIPRE",
                     "BOSNIA_CPT_2","CHIPRE_NORTE_2","INDONESIA_CPT"}
        esp_falta_e2    = [r for r in esp_falta if r["playa"] in _E2_KNOWN]
        esp_falta_real  = [r for r in esp_falta if r["playa"] not in _E2_KNOWN]

        if esp_falta_e2:
            issues.append(_issue(
                "info", "resultado",
                f"{len(esp_falta_e2)} especiales E2/manual — no pasan por rampas del sorter",
                "Estas rutas usan elemento MAN/EXDOCK, no tienen posición física en el sort map. "
                "Es el comportamiento esperado.",
                items=[f"{r['playa']}  ({r['dia_orig']} → {r['dia_new']})" for r in esp_falta_e2],
            ))
        if esp_falta_real:
            issues.append(_issue(
                "error", "resultado",
                f"{len(esp_falta_real)} especiales NO encontradas en el sort map en el día nuevo",
                "Estaban en la parrilla como ESPECIAL DIA CAMBIO pero no aparecen "
                "en el día nuevo del GD generado. Revisa si tienen config en el GD origen.",
                items=[f"{r['playa']}  ({r['dia_orig']} → {r['dia_new']})" for r in esp_falta_real],
            ))

    # ── Check 2: Canceladas NO deben estar en el día CANCELADO ──────────────────
    # Una playa puede estar cancelada el LUNES pero tener salida REGULAR el JUEVES:
    # el check es por (dia_cancelada, playa), no "en ningún día".
    # Además, filtramos solo canceladas "reales": cuyo DPN empieza por CANCELADA_/CANCELADO_
    # (las que tienen DPN=DIA_PLAYA son especiales que también están marcadas como cancelada
    # en la parrilla pero corresponden a la salida movida, no a una eliminación real).
    can_ok, can_presentes = [], []
    dias_cancelada: Dict[str, set] = defaultdict(set)
    for r in canceladas:
        if r["dia_orig"] and r.get("dpn_raw",""):
            dpn_up = r["dpn_raw"].upper()
            if dpn_up.startswith("CANCELADA_") or dpn_up.startswith("CANCELADO_"):
                dias_cancelada[r["playa"]].add(r["dia_orig"])

    for playa, dias_can in dias_cancelada.items():
        dias_mal = []
        for dia in dias_can:
            if playa in playas_por_dia.get(dia, set()):
                dias_mal.append(dia)
        if dias_mal:
            can_presentes.append({"playa": playa, "dias_cancelada": sorted(dias_can),
                                   "dias_mal": sorted(dias_mal)})
        else:
            can_ok.append({"playa": playa})

    if can_ok:
        issues.append(_issue(
            "ok", "resultado",
            f"{len(can_ok)} canceladas correctamente eliminadas del sort map en su día",
            "No aparecen en el GD generado en los días en que fueron canceladas.",
        ))

    if can_presentes:
        items_detail = [
            f"{r['playa']}  (cancelada {'+'.join(r['dias_cancelada'])} — sigue en: {', '.join(r['dias_mal'])})"
            for r in can_presentes
        ]
        issues.append(_issue(
            "error", "resultado",
            f"{len(can_presentes)} canceladas que SIGUEN en el sort map en su día cancelado",
            "Estas playas están marcadas como CANCELADA en la parrilla para un día concreto "
            "pero siguen apareciendo en el GD generado en ese mismo día.",
            items=items_detail,
        ))

    # ── Check 3: Especiales NO deben estar en el día ORIGINAL ─────────────────
    # Build set of (dia, playa) that are confirmed cancelled (real CANCELADA_ rows)
    confirmed_cancelled_days: set = set()
    for r in canceladas:
        if r["dia_orig"] and r.get("dpn_raw",""):
            dpn_up = r["dpn_raw"].upper()
            if dpn_up.startswith("CANCELADA_") or dpn_up.startswith("CANCELADO_"):
                confirmed_cancelled_days.add((r["dia_orig"], r["playa"]))

    esp_still_orig_ok, esp_still_orig = [], []
    seen3 = set()
    for r in especiales:
        playa    = r["playa"]
        dia_orig = r["dia_orig"]
        if playa in seen3 or not dia_orig:
            continue
        seen3.add(playa)
        playas_en_dia_orig = playas_por_dia.get(dia_orig, set())
        if playa in playas_en_dia_orig:
            # Only flag if this day was NOT also explicitly cancelled
            # (a confirmed cancel means that entry was already removed correctly)
            if (dia_orig, playa) not in confirmed_cancelled_days:
                esp_still_orig.append(r)
            else:
                esp_still_orig_ok.append(r)
        else:
            esp_still_orig_ok.append(r)

    if esp_still_orig_ok:
        issues.append(_issue(
            "ok", "resultado",
            f"{len(esp_still_orig_ok)} especiales correctamente eliminadas del día original",
            "No aparecen en su día de origen en el GD generado.",
        ))

    if esp_still_orig:
        issues.append(_issue(
            "error", "resultado",
            f"{len(esp_still_orig)} especiales que SIGUEN en el día original",
            "Estas playas son ESPECIAL DIA CAMBIO pero todavía aparecen en su día "
            "original en el GD generado — deberían haberse eliminado al moverlas.",
            items=[f"{r['playa']}  (día orig: {r['dia_orig']})" for r in esp_still_orig],
        ))

    return issues



# ── Zona consistency check ────────────────────────────────────────────────────

def validate_zona_consistency(parrilla_bytes: bytes, gd_bytes: bytes) -> List[Dict]:
    """
    Check especiales where parrilla says zona=E2 but GD origen has real sorter elements.
    The GD always takes precedence — this generates a warning to fix the parrilla.
    """
    issues = []
    if gd_bytes is None:
        return issues

    # Build playa → has_sorter from GD origen
    from openpyxl import load_workbook as _lwb_z
    import re as _re_z

    wb_gd = _lwb_z(io.BytesIO(gd_bytes), read_only=True)
    ws_gd = wb_gd.active
    rows_gd = list(ws_gd.iter_rows(values_only=True))
    hdr_gd = [str(h or "").strip() for h in rows_gd[0]]
    is_dxc = "Estado" in hdr_gd or "Secuencia" in hdr_gd
    gi_desc = 1 if is_dxc else 2
    gi_zona = 2 if is_dxc else 3
    gi_elem = 5 if is_dxc else 6

    playa_has_sorter: dict = {}
    for row in rows_gd[1:]:
        desc = str(row[gi_desc] or "") if gi_desc < len(row) else ""
        zona = str(row[gi_zona] or "").upper() if gi_zona < len(row) else ""
        elem = str(row[gi_elem] or "") if gi_elem < len(row) else ""
        if zona != "POSTEX":
            continue
        _, _, playa = _parse_gd_desc(desc)
        if not playa:
            continue
        pu = playa.upper()
        if pu not in playa_has_sorter:
            playa_has_sorter[pu] = False
        if _SORTER_RE.match(elem) and not _E2_RE.match(elem):
            playa_has_sorter[pu] = True

    # Check parrilla especiales with zona=E2 against GD
    wb_p = _lwb_z(io.BytesIO(parrilla_bytes), read_only=True)
    target = next((s for s in wb_p.sheetnames
                   if any(str(h or "").strip().upper() == "TIPO_SALIDA"
                          for h in next(wb_p[s].iter_rows(values_only=True, max_row=1), ()))),
                  None)
    if not target:
        return issues

    rows_p = list(wb_p[target].iter_rows(values_only=True))
    col_p = {str(h or "").strip().upper(): i for i, h in enumerate(rows_p[0]) if h}

    import re as _re_z2
    _DAY_RE_Z = _re_z2.compile(
        r"^(?:DOMINGO|LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO)_(.+)$",
        _re_z2.IGNORECASE)

    def _gv_p(row, col):
        i = col_p.get(col)
        if i is None or i >= len(row) or row[i] is None: return ""
        s = str(row[i]).strip()
        return "" if s.startswith("=") or s == "#N/A" else s

    zona_e2_but_sorter = []
    seen = set()
    for r in rows_p[1:]:
        tipo = _gv_p(r, "TIPO_SALIDA").upper()
        if "ESPECIAL DIA CAMBIO" not in tipo:
            continue
        zona = _gv_p(r, "ZONA").upper()
        if zona != "E2":
            continue
        dpo = _gv_p(r, "DIA_PLAYA_ORIGINAL")
        _mz = _DAY_RE_Z.match(dpo)
        playa = _mz.group(1).strip().upper() if _mz else _gv_p(r, "AGRUPACION_PLAYA").upper()
        if not playa or playa in seen:
            continue
        seen.add(playa)
        if playa_has_sorter.get(playa) is True:
            zona_e2_but_sorter.append(playa)

    if zona_e2_but_sorter:
        issues.append(_issue(
            "warning", "contenido",
            f"{len(zona_e2_but_sorter)} especiales marcadas como E2 en parrilla pero con elementos sorter en GD",
            "La parrilla indica zona E2 pero el GD origen tiene posiciones de rampa reales. "
            "El script usará el GD (las tratará como sorter). Revisar la zona en la parrilla.",
            items=zona_e2_but_sorter,
        ))
    else:
        issues.append(_issue(
            "ok", "contenido",
            "Zonas de especiales consistentes entre parrilla y GD",
            "No se detectaron especiales marcadas como E2 en parrilla con elementos sorter en GD.",
        ))

    return issues
