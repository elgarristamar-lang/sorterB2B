# validate_parrilla.py — Version 0.02
# Validación de parrilla + GD antes de generar.
# Devuelve una lista de issues estructurados; no bloquea la generación.
#
# Cambios v0.02:
#   - Acepta parámetro sheet_name para usar la hoja seleccionada por el usuario
#   - Reconoce columna DIA_PLAYA (formato nuevo, equivalente a DIA_PLAYA_NEW)
#   - Mapea nuevos tipos: ESPECIAL DIA, ESPECIAL DIA+CUTOFF, ESPECIAL DIA + CUTOFF
#   - Mapea CANCELADA (MOVIDO...) como cancelada, BI SEMANAL como regular
#   - Deriva dia_new desde TURNO_REPARTO cuando no hay DIA_SALIDA_NEW
#   - Pasa sheet_name a validate_output y validate_zona_consistency
#
# Uso:
#   from validate_parrilla import validate
#   issues = validate(parrilla_bytes, gd_bytes, sheet_name="AGENDA S26 Sant Joan")

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

# Letra de bloque → día de la semana (usado para derivar dia_new desde TURNO_REPARTO)
_LETRA_DIA = {
    "D": "DOMINGO", "L": "LUNES",    "M": "MARTES",
    "X": "MIERCOLES", "J": "JUEVES", "V": "VIERNES", "S": "SABADO",
}

# Hojas base a excluir del selector de "hojas de evento"
_BASE_SHEETS = {"B2B", "B2C", "BLOQUES", "RESUMEN BLOQUES", "VOLUMENES", "APY",
                "RESUMEN BLOQUES", "BLOQUES S26", "BLOQUES 25 DE MAYO"}


def _extract_playa(dpn_val: str) -> str:
    """
    Extract playa name from DIA_PLAYA / DIA_PLAYA_NEW value.
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


def _derive_dia_new_from_turno(turno_val: str) -> str:
    """
    Derive the new expedition day from TURNO_REPARTO.
    TURNO_REPARTO format: 'D4', 'L4', 'V4', 'M1', 'X3', 'J5', etc.
    The first letter indicates the block day: D=DOMINGO, L=LUNES, M=MARTES, etc.
    For complex values like 'L4-V5', uses the first letter only.
    """
    if not turno_val:
        return ""
    letra = turno_val.strip()[0].upper()
    return _LETRA_DIA.get(letra, "")


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

def _load_parrilla(xlsx_bytes: bytes, sheet_name: Optional[str] = None) -> Dict:
    """
    Returns dict with:
      sheets, col_issues, records, stats
    records: list of {playa, dia_orig, dia_new, tipo, dpn_raw}

    sheet_name: if provided and valid, use that sheet; otherwise auto-detect.
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    sheets = wb.sheetnames

    col_issues = []
    records = []
    stats = {"regular": 0, "cancelada": 0, "especial": 0, "irregular": 0, "other": 0}

    # Determine which sheet to use
    target_sheet = None

    # 1. If caller specifies a sheet, use it if it has TIPO_SALIDA
    if sheet_name and sheet_name in sheets:
        ws_probe = wb[sheet_name]
        hdr_probe = next(ws_probe.iter_rows(values_only=True, max_row=1), ())
        if any(str(h or "").strip().upper() == "TIPO_SALIDA" for h in hdr_probe):
            target_sheet = sheet_name

    # 2. Auto-detect: find first non-base sheet with TIPO_SALIDA
    if not target_sheet:
        for sh in sheets:
            if sh.upper() in _BASE_SHEETS:
                continue
            ws = wb[sh]
            hdr = next(ws.iter_rows(values_only=True, max_row=1), ())
            if any(str(h or "").strip().upper() == "TIPO_SALIDA" for h in hdr):
                target_sheet = sh
                break

    # 3. Fallback: any sheet with TIPO_SALIDA
    if not target_sheet:
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

    # Warn if we used a different sheet than requested
    if sheet_name and sheet_name != target_sheet:
        col_issues.append(_issue(
            "warning", "estructura",
            f"Hoja '{sheet_name}' no encontrada o sin TIPO_SALIDA — usando '{target_sheet}'",
            f"La hoja seleccionada no existe o no tiene la columna requerida. "
            f"Se procesará '{target_sheet}'.",
            autocorrected=True,
        ))

    ws = wb[target_sheet]
    all_rows = list(ws.iter_rows(values_only=True))
    hdr = all_rows[0]
    col = {str(h or "").strip().upper(): i for i, h in enumerate(hdr) if h}

    # ── Check PLAYA column ────────────────────────────────────────────────────
    # Accepted: PLAYA, AGRUPACION_PLAYA (old format)
    #           DIA_PLAYA, DIA_PLAYA_NEW (new/unified format — DIA_PLAYA is standard)
    has_playa    = "PLAYA" in col or "AGRUPACION_PLAYA" in col
    has_dpn      = "DIA_PLAYA_NEW" in col or "DIA_PLAYA" in col

    if not has_playa and not has_dpn:
        col_issues.append(_issue(
            "error", "estructura",
            "No hay columna PLAYA, DIA_PLAYA ni DIA_PLAYA_NEW",
            f"Columnas encontradas: {', '.join(col.keys())}. "
            "No es posible identificar los destinos.",
        ))
    elif not has_playa and has_dpn:
        # DIA_PLAYA is the new standard — treat as info, not warning
        dpn_col = "DIA_PLAYA_NEW" if "DIA_PLAYA_NEW" in col else "DIA_PLAYA"
        col_issues.append(_issue(
            "info", "estructura",
            f"Formato unificado: usando columna {dpn_col}",
            f"La playa se extrae de {dpn_col} (ej. LUNES_ESPANA_GUARROMAN → ESPANA_GUARROMAN). "
            "Comportamiento correcto para el formato nuevo de parrilla.",
            autocorrected=True,
        ))

    # ── Check BLOQUE column ───────────────────────────────────────────────────
    if "BLOQUE" not in col:
        col_issues.append(_issue(
            "warning", "estructura",
            "Columna BLOQUE no encontrada",
            "Sin esta columna, el bloque horario de las especiales se derivará "
            "del ID_CLUSTER — puede ser menos preciso.",
            autocorrected=True,
        ))

    # ── Check Resumen Bloques / bloques sheet ─────────────────────────────────
    has_bloques_sheet = "Resumen Bloques" in sheets
    bloques_alt = None
    if not has_bloques_sheet:
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
                "info", "estructura",
                f"Hoja 'Resumen Bloques' no encontrada — se usará '{bloques_alt}'",
                "Los timings horarios de bloques se leerán desde esta hoja. "
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

    # ── Parse records ─────────────────────────────────────────────────────────
    tipo_idx   = col.get("TIPO_SALIDA")
    # Columna de playa: DIA_PLAYA > DIA_PLAYA_NEW > (legacy) DIA_PLAYA_ORIGINAL
    dpn_idx    = col.get("DIA_PLAYA") or col.get("DIA_PLAYA_NEW")
    dpo_idx    = col.get("DIA_PLAYA_ORIGINAL")           # old format only
    dsn_idx    = col.get("DIA_SALIDA_NEW")               # old format only
    dso_idx    = col.get("DIA_SALIDA_ORIGINAL") or col.get("DIA_SALIDA")
    playa_idx  = col.get("PLAYA") or col.get("AGRUPACION_PLAYA")
    bloque_idx = col.get("BLOQUE")
    idc_idx    = col.get("ID_CLUSTER")
    idcn_idx   = col.get("ID_CLUSTER_NEW")
    turno_idx  = col.get("TURNO_REPARTO")                # new format — encodes dia_new

    cancelado_masc = []  # rows with CANCELADO_ (masculine) prefix

    for row in all_rows[1:]:
        def g(i):
            if i is None or i >= len(row) or row[i] is None:
                return ""
            s = str(row[i]).strip()
            return "" if s.startswith("=") or s == "#N/A" else s

        tipo     = g(tipo_idx).upper().strip()
        dpn      = g(dpn_idx)    # DIA_PLAYA or DIA_PLAYA_NEW → día+playa
        dpo      = g(dpo_idx)    # DIA_PLAYA_ORIGINAL (old format)
        dia_new  = g(dsn_idx).upper()   # DIA_SALIDA_NEW (old format)
        dia_orig = g(dso_idx).upper()
        turno    = g(turno_idx)

        # Derive dia_new from TURNO_REPARTO when DIA_SALIDA_NEW is absent (new format)
        if not dia_new and turno:
            dia_new = _derive_dia_new_from_turno(turno)

        # ── Extract playa name ────────────────────────────────────────────────
        if playa_idx is not None and g(playa_idx):
            playa = g(playa_idx)
        elif tipo in ("ESPECIAL DIA CAMBIO",) and dpo:
            playa = _extract_playa(dpo)
            if not dia_orig:
                m = _DAY_PFX.match(dpo)
                if m:
                    dia_orig = m.group(1).upper()
        elif dpn:
            playa = _extract_playa(dpn)
            if dpn.upper().startswith("CANCELADO_"):
                cancelado_masc.append(dpn)
        else:
            playa = ""

        # For CANCELADA: also extract dia_orig from DIA_PLAYA_ORIGINAL if missing
        if tipo == "CANCELADA" and not dia_orig and dpo:
            m = _DAY_PFX.match(dpo)
            if m:
                dia_orig = m.group(1).upper()

        # dia_orig from DIA_PLAYA prefix when DIA_SALIDA_ORIGINAL is absent (new format)
        if not dia_orig and dpn:
            m = _DAY_PFX.match(dpn)
            if m:
                dia_orig = m.group(1).upper()

        if not playa or not tipo:
            continue

        # ── Normalize tipo → kind ─────────────────────────────────────────────
        # Order matters: check CANCELADA variants first, then ESPECIAL variants
        if tipo.startswith("CANCELADA"):
            # Handles: CANCELADA, CANCELADA (MOVIDO A S21), etc.
            stats["cancelada"] += 1
            kind = "cancelada"
        elif "ESPECIAL" in tipo and "DIA" in tipo:
            # Handles: ESPECIAL DIA CAMBIO, ESPECIAL DIA, ESPECIAL DIA+CUTOFF,
            #          ESPECIAL DIA + CUTOFF — all involve a day change
            stats["especial"] += 1
            kind = "especial"
        elif tipo in ("REGULAR", "HABITUAL", "ESPECIAL CUTOFF", "BI SEMANAL",
                      "IRREGULAR"):
            # ESPECIAL CUTOFF → only cutoff changes, not day → treat as regular
            # BI SEMANAL → normal cadence (may go twice a week) → treat as regular
            if tipo == "IRREGULAR":
                stats["irregular"] += 1
                kind = "irregular"
            else:
                stats["regular"] += 1
                kind = "regular"
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
        "has_playa_col": has_playa or has_dpn,
        "has_bloques_sheet": has_bloques_sheet or bool(bloques_alt),
        "cancelado_masc": cancelado_masc,
    }


def _load_gd_playas(xlsx_bytes: bytes) -> Dict:
    """
    Returns:
      by_dia_playa: {(dia, playa): [elementos]} for POSTEX sorter entries
      all_playas: set of playa names with sorter elements
      e2_playas: set of playa names that are E2/manual (MAN/EXDOCK elements only)
    """
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(xlsx_bytes), read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h or "").strip() for h in all_rows[0]]
    hdrs = {h.upper(): i for i, h in enumerate(hdr)}
    is_dxc = "Estado" in hdr or "Secuencia" in hdr

    if is_dxc:
        idx_desc = hdrs.get("DESCRIPCIÓN GRUPOS DE DESTINO", hdrs.get("DESCRIPCION GRUPOS DE DESTINO"))
        idx_zona = hdrs.get("TIPO DE ZONA")
        idx_elem = hdrs.get("ELEMENTO")
    else:
        idx_desc, idx_zona, idx_elem = 2, 3, 6

    by_dia_playa = defaultdict(list)
    all_playas   = set()
    playa_has_sorter: dict = {}   # playa → True if has sorter elem, False if only E2

    for row in all_rows[1:]:
        def g(i):
            if i is None or i >= len(row) or row[i] is None: return ""
            s = str(row[i]).strip()
            return "" if s.startswith("=") else s

        desc = g(idx_desc)
        zona = g(idx_zona).upper()
        elem = g(idx_elem)

        if zona != "POSTEX": continue

        bloque, dia, playa = _parse_gd_desc(desc)
        if not (dia and playa): continue

        pu = playa.upper()
        if pu not in playa_has_sorter:
            playa_has_sorter[pu] = False

        if _SORTER_RE.match(elem) and not _E2_RE.match(elem):
            playa_has_sorter[pu] = True
            all_playas.add(pu)
            by_dia_playa[(dia.upper(), pu)].append(elem)
        elif _E2_RE.match(elem):
            pass  # E2 element — tracked in playa_has_sorter as False

    e2_playas = {p for p, has_sorter in playa_has_sorter.items() if not has_sorter}

    return {"by_dia_playa": dict(by_dia_playa), "all_playas": all_playas,
            "e2_playas": e2_playas}


# ── Main validation ───────────────────────────────────────────────────────────

def validate(parrilla_bytes: bytes, gd_bytes: Optional[bytes] = None,
             sheet_name: Optional[str] = None) -> List[Dict]:
    """
    Run all validations. Returns list of issue dicts.
    sheet_name: the parrilla sheet selected by the user (e.g. "AGENDA S26 Sant Joan").
    GD is optional — without it, cobertura checks are skipped.
    """
    issues = []

    # ── 1. Parse parrilla ─────────────────────────────────────────────────────
    par = _load_parrilla(parrilla_bytes, sheet_name=sheet_name)
    issues.extend(par["col_issues"])

    records    = par["records"]
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
        tipo_vals = list({r["tipo"] for r in records})[:10]
        issues.append(_issue(
            "error", "contenido",
            "No se detectó ninguna especial ni cancelada",
            "Hay registros en la parrilla pero ninguno tiene TIPO_SALIDA de tipo "
            "ESPECIAL DIA (*) ni CANCELADA. "
            "Comprueba que la hoja seleccionada es la del evento especial (ej. AGENDA S26 ...), "
            "no la hoja base B2B.",
            items=tipo_vals,
        ))

    # Especiales without new day
    esp_sin_dia = [r for r in especiales if not r["dia_new"]]
    if esp_sin_dia:
        issues.append(_issue(
            "warning", "contenido",
            f"{len(esp_sin_dia)} especiales sin día nuevo identificado",
            "No se pudo determinar el nuevo día de expedición (ni de DIA_SALIDA_NEW "
            "ni de TURNO_REPARTO). Se intentará derivar del ID_CLUSTER pero puede fallar.",
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
    gd_e2_playas = gd.get("e2_playas", set())

    sin_config_ninguno = []
    sin_config_orig    = []
    ok_count = 0

    seen = set()
    for r in especiales:
        playa    = r["playa"]
        dia_orig = r["dia_orig"]
        if playa in seen:
            continue
        seen.add(playa)

        if playa in gd_e2_playas:
            sin_config_orig.append({**r, "found_on_days": ["E2/MANUAL"]})
        elif playa not in gd_playas:
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
        e2_in_orig     = [r for r in sin_config_orig if r.get("found_on_days") == ["E2/MANUAL"]]
        fallback_orig  = [r for r in sin_config_orig if r.get("found_on_days") != ["E2/MANUAL"]]

        if e2_in_orig:
            issues.append(_issue(
                "info", "cobertura",
                f"{len(e2_in_orig)} especiales con config E2/manual en GD — no pasan por sorter",
                "Tienen config en el GD origen pero con elementos MAN/EXDOCK (ruta E2). "
                "No se les asignará posición de rampa — es el comportamiento esperado.",
                items=[r["playa"] for r in e2_in_orig],
            ))
        if fallback_orig:
            items_detail = [
                f"{r['playa']}  (día original: {r['dia_orig'] or '?'} → encontrada en: {', '.join(r['found_on_days'][:3])})"
                for r in fallback_orig[:10]
            ]
            issues.append(_issue(
                "warning", "cobertura",
                f"{len(fallback_orig)} especiales sin config en el día original — se usará fallback",
                "Existen en el GD pero no en el día de origen de la parrilla. "
                "El script buscará el día con más posiciones y lo usará como fuente.",
                items=items_detail,
            ))

    if sin_config_ninguno:
        issues.append(_issue(
            "error", "cobertura",
            f"{len(sin_config_ninguno)} especiales SIN config en el GD en ningún día",
            "Estas playas aparecen como ESPECIAL DIA en la parrilla pero no tienen "
            "ninguna entrada POSTEX en el GD origen. Sin esa base no se pueden calcular "
            "qué posiciones asignar en el nuevo día — quedarán sin sort map.",
            items=[r["playa"] for r in sin_config_ninguno],
        ))

    # Canceladas que siguen en GD (REVISAR cases)
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
        _can_dias: dict = {}
        for _r in canceladas_en_gd:
            _p = _r["playa"]
            if _p not in _can_dias:
                _can_dias[_p] = []
            if _r["dia_orig"]:
                _can_dias[_p].append(_r["dia_orig"])
        _can_items = [
            f"{p}  ({', '.join(sorted(set(dias)))})" if dias else p
            for p, dias in sorted(_can_dias.items())
        ]
        issues.append(_issue(
            "warning", "cobertura",
            f"{len(_can_dias)} canceladas que tienen config en GD (⚠ REVISAR)",
            "Estas playas están marcadas como CANCELADA en la parrilla pero tienen "
            "entradas en el GD origen. Se eliminarán del GD generado — confirma que es correcto.",
            items=_can_items[:20],
        ))

    # Zona consistency check
    issues.extend(validate_zona_consistency(parrilla_bytes, gd_bytes,
                                            sheet_name=sheet_name))

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

        if re.search(r"_CANCELADA_SOLO_W[^_\s]*", desc, re.IGNORECASE):
            desc = re.sub(r"_CANCELADA_SOLO_W[^_\s)]*", "", desc)

        bloque, dia, playa = _parse_gd_desc(desc)
        if not (dia and playa):
            continue

        dia = dia.upper()
        playa = playa.upper()
        playa = re.sub(r"\s*\(ESPECIAL.*\)$", "", playa).strip()

        playas_por_dia[dia].add(playa)
        all_playas.add(playa)

    return {"playas_por_dia": dict(playas_por_dia), "all_playas": all_playas}


def validate_output(parrilla_bytes: bytes, gd_output_bytes: bytes,
                    sheet_name: Optional[str] = None) -> List[Dict]:
    """
    Post-generation validation: cross-check the generated GD against the parrilla.
    sheet_name: the parrilla sheet selected by the user.
    """
    issues = []

    par = _load_parrilla(parrilla_bytes, sheet_name=sheet_name)
    records    = par["records"]
    especiales = [r for r in records if r["tipo"] == "especial"]
    canceladas = [r for r in records if r["tipo"] == "cancelada"]

    if not records:
        issues.append(_issue(
            "error", "resultado",
            "No se pudieron leer registros de la parrilla para validar",
            "Revisa que el fichero de parrilla es correcto.",
        ))
        return issues

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
        if playa in playas_por_dia.get(dia_new, set()):
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
        _E2_KNOWN = {"BOSNIA_CPT","CHIPRE_NORTE","INDONESIA","CHIPRE",
                     "BOSNIA_CPT_2","CHIPRE_NORTE_2","INDONESIA_CPT"}
        esp_falta_e2   = [r for r in esp_falta if r["playa"] in _E2_KNOWN]
        esp_falta_real = [r for r in esp_falta if r["playa"] not in _E2_KNOWN]

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
                "Estaban en la parrilla como ESPECIAL DIA pero no aparecen "
                "en el día nuevo del GD generado. Revisa si tienen config en el GD origen.",
                items=[f"{r['playa']}  ({r['dia_orig']} → {r['dia_new']})" for r in esp_falta_real],
            ))

    # ── Check 2: Canceladas NO deben estar en su día cancelado ───────────────
    can_ok, can_presentes = [], []
    dias_cancelada: Dict[str, set] = defaultdict(set)
    for r in canceladas:
        if r["dia_orig"] and r.get("dpn_raw",""):
            dpn_up = r["dpn_raw"].upper()
            if dpn_up.startswith("CANCELADA_") or dpn_up.startswith("CANCELADO_"):
                dias_cancelada[r["playa"]].add(r["dia_orig"])
        # New format: CANCELADA without DPN prefix — use dia_orig directly
        elif r["tipo"] == "cancelada" and r["dia_orig"]:
            dias_cancelada[r["playa"]].add(r["dia_orig"])

    for playa, dias_can in dias_cancelada.items():
        dias_mal = [dia for dia in dias_can if playa in playas_por_dia.get(dia, set())]
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
    confirmed_cancelled_days: set = set()
    for r in canceladas:
        if r["dia_orig"]:
            confirmed_cancelled_days.add((r["dia_orig"], r["playa"]))

    esp_still_orig_ok, esp_still_orig = [], []
    seen3 = set()
    for r in especiales:
        playa    = r["playa"]
        dia_orig = r["dia_orig"]
        if playa in seen3 or not dia_orig:
            continue
        seen3.add(playa)
        if playa in playas_por_dia.get(dia_orig, set()):
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
            "Estas playas son ESPECIAL DIA pero todavía aparecen en su día "
            "original en el GD generado — deberían haberse eliminado al moverlas.",
            items=[f"{r['playa']}  (día orig: {r['dia_orig']})" for r in esp_still_orig],
        ))

    # ── Tabla resumen: parrilla vs sort map por día ───────────────────────────
    DAYS_ORDER = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]

    par_by_dia_new: dict  = defaultdict(list)
    par_by_dia_orig: dict = defaultdict(list)
    for r in especiales:
        if r["dia_new"]:
            par_by_dia_new[r["dia_new"]].append(r["playa"])
        if r["dia_orig"]:
            par_by_dia_orig[r["dia_orig"]].append(r["playa"])

    active_days = sorted(
        {r["dia_new"] for r in especiales if r["dia_new"]} |
        {r["dia_orig"] for r in especiales if r["dia_orig"]},
        key=lambda d: DAYS_ORDER.index(d) if d in DAYS_ORDER else 99
    )

    if active_days:
        table_rows = []
        for dia in active_days:
            esp_llegan   = sorted(set(par_by_dia_new.get(dia, [])))
            sm_presentes = playas_por_dia.get(dia, set())
            match        = [p for p in esp_llegan if p in sm_presentes]
            faltantes    = [p for p in esp_llegan if p not in sm_presentes]
            table_rows.append({
                "dia": dia, "par_llegan": len(esp_llegan),
                "sm": len(sm_presentes), "match": len(match), "faltantes": faltantes,
            })

        lines = []
        all_ok = True
        for row in table_rows:
            status = "OK" if not row["faltantes"] else "FALTA"
            if row["faltantes"]: all_ok = False
            line = f"{status}  {row['dia']:<12}  parrilla llegan: {row['par_llegan']}  sort map: {row['match']}/{row['par_llegan']}"
            if row["faltantes"]:
                line += f"  — falta: {', '.join(row['faltantes'])}"
            lines.append(line)

        issues.append(_issue(
            "ok" if all_ok else "warning",
            "resultado",
            "Resumen por dia: especiales parrilla vs sort map",
            "\n".join(lines),
        ))

    return issues


# ── Zona consistency check ────────────────────────────────────────────────────

def validate_zona_consistency(parrilla_bytes: bytes, gd_bytes: bytes,
                               sheet_name: Optional[str] = None) -> List[Dict]:
    """
    Check especiales where parrilla says zona=E2 but GD origen has real sorter elements.
    """
    issues = []
    if gd_bytes is None:
        return issues

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

    # Load parrilla with the selected sheet
    wb_p = _lwb_z(io.BytesIO(parrilla_bytes), read_only=True)

    # Use sheet_name if provided, else auto-detect
    target = None
    if sheet_name and sheet_name in wb_p.sheetnames:
        ws_probe = wb_p[sheet_name]
        hdr_probe = next(ws_probe.iter_rows(values_only=True, max_row=1), ())
        if any(str(h or "").strip().upper() == "TIPO_SALIDA" for h in hdr_probe):
            target = sheet_name
    if not target:
        target = next((s for s in wb_p.sheetnames
                       if s.upper() not in _BASE_SHEETS and
                       any(str(h or "").strip().upper() == "TIPO_SALIDA"
                           for h in next(wb_p[s].iter_rows(values_only=True, max_row=1), ()))),
                      None)
    if not target:
        target = next((s for s in wb_p.sheetnames
                       if any(str(h or "").strip().upper() == "TIPO_SALIDA"
                              for h in next(wb_p[s].iter_rows(values_only=True, max_row=1), ()))),
                      None)
    if not target:
        return issues

    rows_p = list(wb_p[target].iter_rows(values_only=True))
    col_p = {str(h or "").strip().upper(): i for i, h in enumerate(rows_p[0]) if h}

    _DAY_RE_Z = re.compile(
        r"^(?:DOMINGO|LUNES|MARTES|MIERCOLES|JUEVES|VIERNES|SABADO)_(.+)$",
        re.IGNORECASE)

    def _gv_p(row, col):
        i = col_p.get(col)
        if i is None or i >= len(row) or row[i] is None: return ""
        s = str(row[i]).strip()
        return "" if s.startswith("=") or s == "#N/A" else s

    zona_e2_but_sorter = []
    seen = set()
    for r in rows_p[1:]:
        tipo = _gv_p(r, "TIPO_SALIDA").upper()
        # Match all ESPECIAL DIA variants
        if not ("ESPECIAL" in tipo and "DIA" in tipo):
            continue
        zona = _gv_p(r, "ZONA").upper()
        if zona != "E2":
            continue
        # Try DIA_PLAYA_ORIGINAL first, then DIA_PLAYA (new format)
        dpo = _gv_p(r, "DIA_PLAYA_ORIGINAL") or _gv_p(r, "DIA_PLAYA")
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
            f"{len(zona_e2_but_sorter)} especiales con zona E2 en parrilla pero posiciones E3 en GD",
            "La parrilla indica zona E2 pero el GD origen tiene posiciones de rampa reales (E3). "
            "La zona en la parrilla es incorrecta — se procesará correctamente usando el GD.",
            items=zona_e2_but_sorter,
        ))

    return issues
