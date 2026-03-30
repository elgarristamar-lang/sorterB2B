# Version: 0.01
# sorter_map_excel_por_dia.py
# ---------------------------------------------------------
# Genera un Excel "Sorter Map" en formato slots:
# - 1 pestaña por día (DOMINGO..SABADO)
# - Cada día agrega sus bloques (ej DOMINGO = D0..D5)
# - Colores por bloque (repetibles entre días)
# - SOLO cuenta posiciones POSTEX (suelo)
# - Slots usados muestran el BLOQUE en la celda (J1, J2...)
# - Slots multi-bloque -> texto "J1+J2", color gris
# - Columna extra MULTI_BLOQUE por subrampa con detalle de posiciones multi + comentario
# - Incluye LEYENDA y tabla BLOQUES_DESTINOS (1 fila por destino)
# ---------------------------------------------------------

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional, Dict, Set, Tuple, List
from collections import defaultdict
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter


# -----------------------------
# Paths
# -----------------------------
import sys as _sys

def _parse_args():
    """sorter_map_por_dia.py <capacity.csv> <grupo_destinos.xlsx> <bloques_horarios.xlsx> <output.xlsx> [grupo_sheet]"""
    args = _sys.argv[1:]
    if len(args) < 4:
        print("Uso: python sorter_map_por_dia.py <ramp_capacity.csv> <grupo_destinos.xlsx> <bloques_horarios.xlsx> <output.xlsx> [hoja]")
        _sys.exit(1)
    return Path(args[0]), Path(args[1]), Path(args[2]), Path(args[3]), args[4] if len(args) > 4 else "Hoja1"

CAPACITY_CSV, GRUPO_XLSX, BLOQUES_XLSX, _OUTPUT_PATH_ARG, GRUPO_SHEET = _parse_args()

EXCLUDE_PREFIXES = ("R01", "R03")  # rampas no utilizables
MAX_POS_COL_WIDTH = 4
SUBRAMPA_COL_WIDTH = 11
MULTI_COL_WIDTH = 45  # columna extra


# -----------------------------
# Helpers columnas
# -----------------------------
def standardize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_col(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = {c.lower(): c for c in df.columns}
    for cand in candidates:
        key = cand.lower()
        if key in cols:
            return cols[key]
    return None


# -----------------------------
# Parsing
# -----------------------------
def parse_subramp_and_slot_from_elemento(value: object) -> Tuple[Optional[str], Optional[int]]:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None, None

    s = str(value).strip().upper()
    if not s:
        return None, None

    m = re.search(r"\b(R\d{2})\s*[_\- ]\s*([A-D])\s*-\s*(\d+)\b", s)
    if m:
        return f"{m.group(1)}{m.group(2)}", int(m.group(3))

    m2 = re.search(r"\b(R\d{2})([A-D])\s*-\s*(\d+)\b", s)
    if m2:
        return f"{m2.group(1)}{m2.group(2)}", int(m2.group(3))

    m3 = re.search(r"\b(R\d{2})([A-D])\b", s)
    if m3:
        return f"{m3.group(1)}{m3.group(2)}", None

    return None, None


def ramp_sort_key(subramp: str):
    subramp = subramp.strip().upper()
    m = re.match(r"^R(\d{2})([A-D])$", subramp)
    if not m:
        return (999, "Z", subramp)
    return (int(m.group(1)), m.group(2), subramp)


# -----------------------------
# Limpieza destino (agrupación playa)
# -----------------------------
def clean_desc_to_destino(desc: str) -> str:
    s = (desc or "").strip()
    s = re.sub(r"^\[[^\]]+\]\s*", "", s)  # quita [B2B]
    s = s.strip()
    s = re.sub(r"^(?:\d+)?BLO[A-Z]\d+[_\- ]+", "", s, flags=re.IGNORECASE)  # quita prefijo bloque
    return s.strip().upper()


# -----------------------------
# Filtro por bloque token (D0, L1, etc.)
# -----------------------------
def filter_rows_by_block(df: pd.DataFrame, block: str, desc_col: str) -> pd.DataFrame:
    s = df[desc_col].astype(str).fillna("")
    block = block.strip().upper()

    if "BLO" in block:
        mask_literal = s.str.contains(re.escape(block), case=False, na=False)
        if mask_literal.any():
            return df[mask_literal].copy()

    mask_priority = s.str.contains(fr"BLO{re.escape(block)}", case=False, na=False)
    if mask_priority.any():
        return df[mask_priority].copy()

    mask_fallback = s.str.contains(re.escape(block), case=False, na=False)
    return df[mask_fallback].copy()


# -----------------------------
# Loaders
# -----------------------------
def load_capacity(path: Path) -> Dict[str, int]:
    df = pd.read_csv(path, sep=None, engine="python")
    df = standardize_columns(df)

    col_ramp = find_col(df, ["RAMP", "RAMPA", "SUBRAMPA"])
    col_pallets = find_col(df, ["PALLETS", "CAPACITY", "CAPACIDAD", "POSICIONES", "PALLETS_CAPACITY"])

    if not col_ramp or not col_pallets:
        raise ValueError(f"CSV ramp_capacity inválido. Columnas: {list(df.columns)}")

    df[col_ramp] = df[col_ramp].astype(str).str.strip().str.upper()
    df[col_pallets] = pd.to_numeric(df[col_pallets], errors="coerce")

    cap_map: Dict[str, int] = {}
    for _, r in df.iterrows():
        sub = str(r[col_ramp]).strip().upper()
        if not sub:
            continue
        if any(sub.startswith(p) for p in EXCLUDE_PREFIXES):
            continue
        pallets = r[col_pallets]
        if pd.isna(pallets):
            continue
        cap_map[sub] = int(pallets)

    return cap_map


def load_grupo_destinos(path: Path, sheet: str) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=sheet)
    return standardize_columns(df)


def load_bloques_horarios(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path)
    df = standardize_columns(df)

    col_block = find_col(df, ["NUEVO BLOQUE", "BLOQUE", "BLOCK"])
    col_start_day = find_col(df, ["Día LIBERACIÓN BLOQUES", "DIA LIBERACIÓN BLOQUES", "DIA LIBERACION BLOQUES", "DIA LIBERACION"])
    col_start_time = find_col(df, ["Hora LIBERACIÓN BLOQUES", "HORA LIBERACIÓN BLOQUES", "HORA LIBERACION BLOQUES", "HORA LIBERACION"])
    col_end_day = find_col(df, ["Día DESACTIVACIÓN", "DIA DESACTIVACIÓN", "DIA DESACTIVACION"])
    col_end_time = find_col(df, ["Hora DESACTIVACIÓN", "HORA DESACTIVACIÓN", "HORA DESACTIVACION"])

    if not all([col_block, col_start_day, col_start_time, col_end_day, col_end_time]):
        raise ValueError(f"bloques_horarios inválido. Columnas: {list(df.columns)}")

    df[[col_block, col_start_day, col_start_time, col_end_day, col_end_time]] = df[
        [col_block, col_start_day, col_start_time, col_end_day, col_end_time]
    ].ffill()

    out = df.rename(columns={
        col_block: "BLOCK",
        col_start_day: "START_DAY",
        col_start_time: "START_TIME",
        col_end_day: "END_DAY",
        col_end_time: "END_TIME",
    })

    out["BLOCK"] = out["BLOCK"].astype(str).str.strip().str.upper()
    out["START_DAY"] = out["START_DAY"].astype(str).str.strip().str.upper()
    out["END_DAY"] = out["END_DAY"].astype(str).str.strip().str.upper()
    return out[["BLOCK", "START_DAY", "START_TIME", "END_DAY", "END_TIME"]].drop_duplicates(subset=["BLOCK"])


# -----------------------------
# Colores por bloque (reutilizables entre días)
# -----------------------------
PALETTE = [
    "BDD7EE",  # azul claro
    "C6E0B4",  # verde claro
    "F8CBAD",  # naranja claro
    "D9E1F2",  # lila claro
    "E7E6E6",  # gris claro
    "FFF2CC",  # amarillo claro
]

def build_block_color_map_for_day(day_code: str) -> Dict[str, str]:
    m: Dict[str, str] = {}
    for i in range(6):
        key = f"{day_code}{i}"
        m[key] = PALETTE[i % len(PALETTE)]
    return m



# -----------------------------
# Block timing helpers
# -----------------------------
_DAY_IDX = {"DOMINGO":0,"LUNES":1,"MARTES":2,"MIERCOLES":3,"JUEVES":4,"VIERNES":5,"SABADO":6}

def _time_to_min(t) -> int:
    import re as _re
    if hasattr(t, "hour"):
        return t.hour * 60 + t.minute
    m = _re.match(r"(\d{1,2}):(\d{2})", str(t).strip())
    return int(m.group(1))*60 + int(m.group(2)) if m else 0

def build_block_intervals(bloques_df: pd.DataFrame) -> Dict[str, Tuple[int,int]]:
    """Return {block_name: (start_week_min, end_week_min)} for overlap checks.
    Indexed both by full name (e.g. '3BLOM4') and short token (e.g. 'M4').
    """
    import re as _re2
    intervals: Dict[str, Tuple[int,int]] = {}
    for _, row in bloques_df.iterrows():
        block = str(row["BLOCK"]).strip().upper()
        sd = str(row["START_DAY"]).strip().upper()
        ed = str(row["END_DAY"]).strip().upper()
        if sd not in _DAY_IDX or ed not in _DAY_IDX:
            continue
        start = _DAY_IDX[sd] * 1440 + _time_to_min(row["START_TIME"])
        end   = _DAY_IDX[ed] * 1440 + _time_to_min(row["END_TIME"])
        if _DAY_IDX[ed] < _DAY_IDX[sd]:
            end += 7 * 1440
        elif _DAY_IDX[ed] == _DAY_IDX[sd] and end <= start:
            end += 1440
        intervals[block] = (start, end)
        # Also index by short token: "3BLOM4" → "M4", "2BLOL5" → "L5", etc.
        m = _re2.search(r"BLO([A-Z]\d+)$", block)
        if m:
            intervals[m.group(1)] = (start, end)
    return intervals

def blocks_overlap(a: Tuple[int,int], b: Tuple[int,int]) -> bool:
    return a[0] < b[1] and b[0] < a[1]

# -----------------------------
# Cálculo uso POSTEX por día (agregando bloques)
# -----------------------------
def compute_day_usage(
    grupo_df: pd.DataFrame,
    day_blocks: List[str],
    block_intervals: Optional[Dict[str, Tuple[int,int]]] = None,
) -> Tuple[Dict[str, Dict[str, Set[int]]], int]:
    warnings = 0

    desc_col = find_col(grupo_df, ["Descripción Grupos de destino", "Descripcion Grupos de destino", "DESCRIPCION GRUPOS DE DESTINO"])
    elem_col = find_col(grupo_df, ["Elemento", "ELEMENTO"])
    tipo_col = find_col(grupo_df, ["Tipo de zona", "TIPO DE ZONA", "TIPO_ZONA", "TIPO ZONA"])

    if not desc_col or not elem_col or not tipo_col:
        raise ValueError(f"Faltan columnas clave en GRUPO_DESTINOS. Columnas: {list(grupo_df.columns)}")

    usage: Dict[str, Dict[str, Set[int]]] = defaultdict(lambda: defaultdict(set))

    # Build set of day names active for each block token (from intervals)
    _ALL_DAY_NAMES = ["DOMINGO","LUNES","MARTES","MIERCOLES","JUEVES","VIERNES","SABADO"]

    for block_token in day_blocks:
        rows = filter_rows_by_block(grupo_df, block_token, desc_col)
        if rows.empty:
            continue

        # Determine which day this block belongs to (from block_intervals start_day)
        # Entries whose description mentions a DIFFERENT day belong to another day sheet
        # and should not appear here (e.g. "3BLOM2_LUNES_..." when processing MARTES)
        block_start_day = None
        if block_intervals:
            new_iv = block_intervals.get(block_token)
            if new_iv:
                # Reverse-lookup start day from interval start minute
                day_idx = new_iv[0] // 1440
                if 0 <= day_idx < len(_ALL_DAY_NAMES):
                    block_start_day = _ALL_DAY_NAMES[day_idx]

        for _, r in rows.iterrows():
            tipo = str(r[tipo_col]).strip().upper()
            if tipo != "POSTEX":
                continue

            # Skip entries that belong to a different day
            # e.g. "3BLOM2_LUNES_ESPANA_BENAVENTE" on MARTES day sheet
            if block_start_day:
                desc_str = str(r[desc_col]).upper()
                for other_day in _ALL_DAY_NAMES:
                    if other_day == block_start_day:
                        continue
                    # If desc explicitly names another day right after the block prefix, skip
                    if f"_{other_day}_" in desc_str and f"_{block_start_day}_" not in desc_str:
                        break
                else:
                    pass  # no break → ok to include
                # Simpler: if desc contains a day name that != block_start_day → skip
                has_other_day = any(
                    f"_{d}_" in desc_str
                    for d in _ALL_DAY_NAMES if d != block_start_day
                )
                has_own_day = f"_{block_start_day}_" in desc_str
                if has_other_day and not has_own_day:
                    continue  # this entry belongs to a different day

            sub, slot = parse_subramp_and_slot_from_elemento(r[elem_col])
            if sub is None or slot is None:
                v = r[elem_col]
                if v is not None and str(v).strip():
                    warnings += 1
                continue

            if any(sub.startswith(p) for p in EXCLUDE_PREFIXES):
                continue

            # Check timing compatibility with already-assigned blocks on this slot
            if block_intervals:
                new_iv = block_intervals.get(block_token)
                conflict = False
                for existing_block, existing_slots in usage.items():
                    if existing_block.startswith("_CONFLICT_"):
                        continue
                    if existing_block == block_token:  # same block, always compatible
                        continue
                    if slot not in existing_slots.get(sub, set()):
                        continue
                    ex_iv = block_intervals.get(existing_block)
                    if new_iv and ex_iv and blocks_overlap(new_iv, ex_iv):
                        conflict = True
                        break
                if conflict:
                    usage[f"_CONFLICT_{block_token}"][sub].add(slot)
                    continue

            usage[block_token][sub].add(slot)

    return usage, warnings


# -----------------------------
# Excel styles
# -----------------------------
def make_styles():
    bold = Font(bold=True)
    title_font = Font(bold=True, size=12)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    wrap_top = Alignment(vertical="top", wrap_text=True)
    thin = Side(style="thin", color="CCCCCC")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    return bold, title_font, center, left, wrap_top, border


# -----------------------------
# Day sheet writer (UPDATED)
# -----------------------------
def write_day_sheet(
    ws,
    day_name: str,
    day_code: str,
    cap_map: Dict[str, int],
    usage_by_block: Dict[str, Dict[str, Set[int]]],
    block_colors: Dict[str, str],
    bold, title_font, center, left, border
):
    subramps = sorted(cap_map.keys(), key=ramp_sort_key)
    max_pos = max(cap_map.values()) if cap_map else 14

    # Title
    ws["A1"] = f"SORTER MAP - {day_name} (agrega {day_code}0..{day_code}5) | slots POSTEX"
    ws["A1"].font = title_font
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2 + max_pos)  # + MULTI col
    ws.row_dimensions[1].height = 22

    # Header row
    ws["A2"] = "Subrampa"
    ws["A2"].font = bold
    ws["A2"].alignment = center
    ws["A2"].border = border

    for p in range(1, max_pos + 1):
        c = ws.cell(row=2, column=1 + p, value=f"{p:02d}")
        c.font = bold
        c.alignment = center
        c.border = border

    # New extra column
    multi_col_idx = 2 + max_pos
    h = ws.cell(row=2, column=multi_col_idx, value="MULTI_BLOQUE")
    h.font = bold
    h.alignment = center
    h.border = border

    # widths
    ws.column_dimensions["A"].width = SUBRAMPA_COL_WIDTH
    for p in range(1, max_pos + 1):
        ws.column_dimensions[get_column_letter(1 + p)].width = MAX_POS_COL_WIDTH
    ws.column_dimensions[get_column_letter(multi_col_idx)].width = MULTI_COL_WIDTH

    # slot_blocks[sub][slot] = {block_tokens}
    # _CONFLICT_ tokens mark slots where two overlapping blocks both claim the position
    slot_blocks: Dict[str, Dict[int, Set[str]]] = defaultdict(lambda: defaultdict(set))
    conflict_slots: Dict[str, Set[int]] = defaultdict(set)  # sub → set of conflicting slots
    for bt, per_sub in usage_by_block.items():
        is_conflict = bt.startswith("_CONFLICT_")
        for sub, slots in per_sub.items():
            for s in slots:
                if is_conflict:
                    conflict_slots[sub].add(s)
                else:
                    slot_blocks[sub][s].add(bt)

    row = 3
    for sub in subramps:
        cap = cap_map[sub]

        # Subramp cell
        name_cell = ws.cell(row=row, column=1, value=sub)
        name_cell.font = bold
        name_cell.alignment = left
        name_cell.border = border

        # base grid
        for p in range(1, max_pos + 1):
            c = ws.cell(row=row, column=1 + p, value="")
            c.border = border
            c.alignment = center
            if p > cap:
                c.fill = PatternFill("solid", fgColor="F2F2F2")

        # MULTI column base
        multi_cell = ws.cell(row=row, column=multi_col_idx, value="")
        multi_cell.border = border
        multi_cell.alignment = Alignment(horizontal="left", vertical="top", wrap_text=True)

        # paint used slots + put BLOCK text
        multi_details: List[str] = []  # "pos 02: J1+J2"
        sub_conflicts = conflict_slots.get(sub, set())

        for slot, blocks_here in sorted(slot_blocks.get(sub, {}).items(), key=lambda x: x[0]):
            if slot > cap:
                continue
            cell = ws.cell(row=row, column=1 + slot)
            blocks_sorted = sorted(blocks_here)

            if slot in sub_conflicts:
                # Timing conflict: two overlapping blocks claim this slot → red/orange
                cell.fill = PatternFill("solid", fgColor="FF9999")
                cell.value = "⚠" + "+".join(blocks_sorted)
                multi_details.append(f"pos {slot:02d}: CONFLICTO " + "+".join(blocks_sorted))
            elif len(blocks_sorted) == 1:
                b = blocks_sorted[0]
                cell.fill = PatternFill("solid", fgColor=block_colors.get(b, "FFFFFF"))
                cell.value = b
            else:
                # Multiple blocks but no timing conflict (compatible time windows)
                cell.fill = PatternFill("solid", fgColor="BFBFBF")
                cell.value = "+".join(blocks_sorted)
                multi_details.append(f"pos {slot:02d}: " + "+".join(blocks_sorted))

            # Also paint conflict slots not yet in slot_blocks
            cell.alignment = center
            cell.border = border
            cell.font = Font(bold=True, size=9)

        # Paint conflict-only slots (slot claimed only by conflicting blocks)
        for slot in sorted(sub_conflicts):
            if slot > cap: continue
            if slot in slot_blocks.get(sub, {}): continue  # already painted above
            cell = ws.cell(row=row, column=1 + slot)
            cell.fill = PatternFill("solid", fgColor="FF9999")
            cell.value = "⚠"
            cell.alignment = center
            cell.border = border
            cell.font = Font(bold=True, size=9)
            multi_details.append(f"pos {slot:02d}: CONFLICTO HORARIO")

        # rellenar columna MULTI_BLOQUE + comentario
        if multi_details:
            text = "; ".join(multi_details)
            multi_cell.value = text
            multi_cell.comment = Comment("\n".join(multi_details), f"{day_name}_multi")

        ws.row_dimensions[row].height = 18
        row += 1

    ws.freeze_panes = "B3"


# -----------------------------
# LEYENDA
# -----------------------------
def write_leyenda_sheet(ws, day_block_color_map: Dict[str, str], bold, title_font, center, border):
    ws["A1"] = "LEYENDA (colores por bloque)"
    ws["A1"].font = title_font

    ws["A3"] = "Bloque"
    ws["B3"] = "Color"
    ws["A3"].font = bold
    ws["B3"].font = bold

    r = 4
    for block_token in sorted(day_block_color_map.keys()):
        c1 = ws.cell(row=r, column=1, value=block_token)
        c1.border = border
        c1.alignment = center
        box = ws.cell(row=r, column=2, value="")
        box.fill = PatternFill("solid", fgColor=day_block_color_map[block_token])
        box.border = border
        r += 1

    ws["A10"] = "Notas:"
    ws["A10"].font = bold
    ws["A11"] = "• 1 celda = 1 posición suelo (POSTEX)"
    ws["A12"] = "• Si un slot lo usan 2 bloques del mismo día -> se pinta GRIS + texto 'J1+J2'"
    ws["A13"] = "• Columna MULTI_BLOQUE indica qué posiciones están en más de un bloque"
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 12


# -----------------------------
# BLOQUES_DESTINOS (igual que antes)
# -----------------------------
def write_bloques_destinos_sheet(
    ws,
    grupo_df: pd.DataFrame,
    bloques_df: pd.DataFrame,
    all_block_tokens: List[str],
    bold, title_font, center, wrap_top, border
) -> int:
    desc_col = find_col(grupo_df, ["Descripción Grupos de destino", "Descripcion Grupos de destino", "DESCRIPCION GRUPOS DE DESTINO"])
    elem_col = find_col(grupo_df, ["Elemento", "ELEMENTO"])
    tipo_col = find_col(grupo_df, ["Tipo de zona", "TIPO DE ZONA", "TIPO_ZONA", "TIPO ZONA"])

    if not desc_col or not elem_col or not tipo_col:
        raise ValueError(f"Faltan columnas clave en GRUPO_DESTINOS. Columnas: {list(grupo_df.columns)}")

    ws["A1"] = "BLOQUES_DESTINOS (1 fila por destino + posiciones POSTEX)"
    ws["A1"].font = title_font
    ws.merge_cells("A1:J1")

    headers = [
        "BLOCK_TOKEN", "BLOCK_SCHEDULE", "START_DAY", "START_TIME", "END_DAY", "END_TIME",
        "DESTINO_CLEAN", "POSTEX_POS_N", "POSTEX_POS_SUBRAMPA", "POSTEX_POS_LIST"
    ]
    for col, h in enumerate(headers, start=1):
        c = ws.cell(row=2, column=col, value=h)
        c.font = bold
        c.alignment = center
        c.border = border

    schedules = {}
    for _, r in bloques_df.iterrows():
        schedules[str(r["BLOCK"]).upper()] = (r["START_DAY"], r["START_TIME"], r["END_DAY"], r["END_TIME"])

    out_row = 3
    warnings = 0

    for token in all_block_tokens:
        rows = filter_rows_by_block(grupo_df, token, desc_col)
        if rows.empty:
            continue

        tmp = rows.copy()
        tmp["_DESTINO_CLEAN_"] = tmp[desc_col].astype(str).apply(clean_desc_to_destino)

        destino_groups: Dict[str, pd.DataFrame] = {}
        for destino, g in tmp.groupby("_DESTINO_CLEAN_"):
            destino_groups[destino] = g

        schedule_block_name = None
        for bname in schedules.keys():
            if re.search(fr"BLO{re.escape(token)}\b", bname, re.IGNORECASE):
                schedule_block_name = bname
                break

        start_day = start_time = end_day = end_time = ""
        if schedule_block_name and schedule_block_name in schedules:
            start_day, start_time, end_day, end_time = schedules[schedule_block_name]

        for destino_clean, g in sorted(destino_groups.items()):
            postex_slots: Set[Tuple[str, int]] = set()
            postex_subs: Set[str] = set()

            for _, rr in g.iterrows():
                if str(rr[tipo_col]).strip().upper() != "POSTEX":
                    continue
                sub, slot = parse_subramp_and_slot_from_elemento(rr[elem_col])
                if sub is None or slot is None:
                    v = rr[elem_col]
                    if v is not None and str(v).strip():
                        warnings += 1
                    continue
                if any(sub.startswith(p) for p in EXCLUDE_PREFIXES):
                    continue

                postex_slots.add((sub, slot))
                postex_subs.add(sub)

            pos_list = sorted(postex_slots, key=lambda x: (ramp_sort_key(x[0]), x[1]))
            pos_list_str = ", ".join([f"{s}-{p:02d}" for s, p in pos_list])
            pos_sub_str = ", ".join(sorted(postex_subs, key=ramp_sort_key))
            pos_n = len(pos_list)

            row_values = [
                token,
                schedule_block_name or "",
                start_day, str(start_time), end_day, str(end_time),
                destino_clean,
                pos_n,
                pos_sub_str,
                pos_list_str,
            ]

            for col, v in enumerate(row_values, start=1):
                c = ws.cell(row=out_row, column=col, value=v)
                c.border = border
                c.alignment = wrap_top

            ws.row_dimensions[out_row].height = 32
            out_row += 1

    widths = [12, 16, 12, 10, 12, 10, 36, 12, 26, 90]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A3"
    return warnings


# -----------------------------
# Main
# -----------------------------
DAY_SHEETS = [
    ("DOMINGO", "D"),
    ("LUNES", "L"),
    ("MARTES", "M"),
    ("MIERCOLES", "X"),
    ("JUEVES", "J"),
    ("VIERNES", "V"),
    ("SABADO", "S"),
]


def main():
    print("=== Generador SORTER_MAP Excel (1 pestaña por día) ===")

    cap_map = load_capacity(CAPACITY_CSV)
    grupo = load_grupo_destinos(GRUPO_XLSX, GRUPO_SHEET)
    bloques = load_bloques_horarios(BLOQUES_XLSX)

    wb = Workbook()
    wb.remove(wb.active)

    bold, title_font, center, left, wrap_top, border = make_styles()

    total_warnings = 0
    global_color_map: Dict[str, str] = {}

    block_intervals = build_block_intervals(bloques)

    for day_name, day_code in DAY_SHEETS:
        day_blocks = [f"{day_code}{i}" for i in range(6)]  # D0..D5, etc
        usage_by_block, warnings = compute_day_usage(grupo, day_blocks, block_intervals)
        total_warnings += warnings

        block_colors = build_block_color_map_for_day(day_code)
        global_color_map.update(block_colors)

        ws = wb.create_sheet(day_name)
        write_day_sheet(
            ws=ws,
            day_name=day_name,
            day_code=day_code,
            cap_map=cap_map,
            usage_by_block=usage_by_block,
            block_colors=block_colors,
            bold=bold, title_font=title_font, center=center, left=left, border=border
        )

    ws_leg = wb.create_sheet("LEYENDA")
    write_leyenda_sheet(ws_leg, global_color_map, bold, title_font, center, border)

    ws_tbl = wb.create_sheet("BLOQUES_DESTINOS")
    all_block_tokens = [f"{code}{i}" for _, code in DAY_SHEETS for i in range(6)]
    warnings_tbl = write_bloques_destinos_sheet(
        ws=ws_tbl,
        grupo_df=grupo,
        bloques_df=bloques,
        all_block_tokens=all_block_tokens,
        bold=bold, title_font=title_font, center=center, wrap_top=wrap_top, border=border
    )
    total_warnings += warnings_tbl

    out_path = _OUTPUT_PATH_ARG
    wb.save(out_path)

    print(f"✅ Generado: {out_path}")
    if total_warnings:
        print(f"⚠️ Warnings: {total_warnings} valores de 'Elemento' no parseables (ignorados).")


if __name__ == "__main__":
    main()