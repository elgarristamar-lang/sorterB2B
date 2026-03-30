# Version: 0.01
#!/usr/bin/env python3
"""
Herramienta de configuración del Sorter VDL B2B - MANGO
Procesa semanas especiales: asigna rampas y posiciones libres a destinos
que cambian de día de salida, respetando los horarios de bloque.

Admite dos formatos de GRUPO_DESTINOS:
  - Formato clásico (GRUPO_DESTINOS.xlsx)
  - Formato DXC export (resultadosConsulta*.xlsx): con columnas Estado, Secuencia, etc.

Uso:
    python process_parrilla.py <parrilla.xlsx> <grupo_destinos.xlsx> <ramp_capacity.csv> <sheet_name> [semana]

Ejemplo:
    python process_parrilla.py parrilla_de_salidas.xlsx resultadosConsulta.xlsx ramp_capacity.csv parrilla_test_s14 S14
"""

import sys, re, csv, os, shutil
from datetime import datetime
from collections import defaultdict

try:
    from openpyxl import load_workbook, Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: pip install openpyxl"); sys.exit(1)

BLOQUE_RE  = re.compile(r'^(\d+BLO[A-Z]\d+)_(.+)$')
DAYS       = ['DOMINGO','LUNES','MARTES','MIERCOLES','MIÉRCOLES','JUEVES','VIERNES','SABADO']
E2_RE      = re.compile(r'^(MAN|EXDOCK|DOCK)', re.IGNORECASE)
SORTER_RE  = re.compile(r'^R\d+')

# Rampas reserved for manipulado — never use for automatic especial assignment
EXCLUDED_RAMPAS = {
    'R03A', 'R03B', 'R03C', 'R03D',   # Rampa 3: manipulado exclusivo
}

# Minutes-from-start-of-week (Sunday 00:00)
DAY_MIN = {'DOMINGO':0,'LUNES':1440,'MARTES':2880,'MIERCOLES':4320,
           'MIÉRCOLES':4320,'JUEVES':5760,'VIERNES':7200,'SABADO':8640}

# Cluster-letter → day prefix mapping (L=Lunes, M=Martes, X=Miercoles, J=Jueves, V=Viernes, S=Sabado, D=Domingo)
CLUSTER_DAY = {'D':'DOMINGO','L':'LUNES','M':'MARTES','X':'MIERCOLES',
               'J':'JUEVES','V':'VIERNES','S':'SABADO'}


# ─── HELPERS ──────────────────────────────────────────────────────────────────

def parse_gd_desc(desc):
    if not desc: return None, None, None
    core = re.sub(r'^\[B2B\]\s*', '', str(desc)).strip()
    core = re.sub(r'\s+PARA BAJAR POR.*$', '', core)
    core = re.sub(r'_CANCELADA.*$', '', core)
    m = BLOQUE_RE.match(core)
    if not m: return None, None, None
    bloque, rest = m.group(1), m.group(2)
    for d in DAYS:
        if rest.upper().startswith(d + '_'):
            return bloque, d.replace('MIÉRCOLES','MIERCOLES'), rest[len(d)+1:]
    return bloque, None, rest

def safe_str(val):
    if val is None: return ''
    s = str(val).strip()
    return '' if s.startswith('=') or s == '#N/A' else s

def parse_rampa(elemento):
    m = re.match(r'^R0*(\d+)_([A-Z])-(\d+)$', str(elemento or ''))
    if not m: return None, None
    return f"R{int(m.group(1)):02d}{m.group(2)}", int(m.group(3))

def postex_elem(rampa, pos):
    return f"R{rampa[1:-1]}_{rampa[-1]}-{pos:02d}"

def is_sorter_elem(elem):
    return bool(SORTER_RE.match(str(elem or '')))

def is_e2_elem(elem):
    return bool(E2_RE.match(str(elem or '')))

def timing_to_min(s):
    """'LUNES_16:00' → minutes from Sunday 00:00"""
    if not s or '_' not in s: return None
    day, time = s.split('_', 1)
    h, m = map(int, time.split(':'))
    return DAY_MIN.get(day.upper(), 0) + h * 60 + m

def bloques_overlap(b1, b2, bloque_timings):
    """True if the two bloques have overlapping [liberacion, desactivacion] windows."""
    t1, t2 = bloque_timings.get(b1), bloque_timings.get(b2)
    if not t1 or not t2: return True  # conservative
    l1, d1 = timing_to_min(t1['lib']), timing_to_min(t1['desac'])
    l2, d2 = timing_to_min(t2['lib']), timing_to_min(t2['desac'])
    if None in (l1, d1, l2, d2): return True
    return not (d1 <= l2 or d2 <= l1)

def resolve_bloque_for_new_day(id_cluster_str, new_dia, bloque_timings):
    """
    Given 'L4-J4' and new_dia='LUNES', return the matching bloque code '2BLOL4'.
    Parses each cluster token, finds the one whose day matches new_dia, returns its bloque.
    Falls back to the first token if no exact match.
    """
    if not id_cluster_str: return None
    # Build reverse: cluster_code → bloque
    cluster_to_bloque = {v['cluster']: k for k, v in bloque_timings.items()}
    
    tokens = [t.strip() for t in id_cluster_str.split('-') if t.strip()]
    # Find token matching new_dia
    new_dia_norm = new_dia.upper().replace('MIÉRCOLES','MIERCOLES')
    for token in tokens:
        if not token: continue
        letter = token[0].upper()
        token_day = CLUSTER_DAY.get(letter)
        if token_day == new_dia_norm:
            bloque = cluster_to_bloque.get(token)
            if bloque: return bloque
    # Fallback: first valid token
    for token in tokens:
        bloque = cluster_to_bloque.get(token.strip())
        if bloque: return bloque
    return None


# ─── LOADERS ──────────────────────────────────────────────────────────────────

def load_capacity(path):
    cap = {}
    with open(path, newline='') as f:
        reader = csv.reader(f, delimiter=';')
        next(reader)
        for row in reader:
            if len(row) >= 2 and row[0].strip():
                try: cap[row[0].strip()] = int(row[1].strip())
                except ValueError: pass
    return cap

def load_bloque_timings(parrilla_path):
    """Load Resumen Bloques sheet → {bloque: {lib, cutoff, desac, cluster}}"""
    wb = load_workbook(parrilla_path, read_only=True)
    timings = {}
    if 'Resumen Bloques' not in wb.sheetnames:
        return timings
    ws = wb['Resumen Bloques']
    for r in ws.iter_rows(values_only=True):
        if r[0] and isinstance(r[0], str) and not r[0].startswith('=') and r[2]:
            timings[r[0]] = {
                'cluster': str(r[1]) if r[1] else '',
                'lib':     str(r[2]) if r[2] else '',
                'cutoff':  str(r[3]) if r[3] else '',
                'desac':   str(r[4]) if r[4] else '',
            }
    return timings

def load_grupo_destinos(path):
    wb = load_workbook(path, read_only=True)
    ws = wb.active
    all_rows = list(ws.iter_rows(values_only=True))
    file_header = all_rows[0]
    raw_data = all_rows[1:]

    hdrs = [str(h).strip() if h else '' for h in file_header]
    is_dxc = 'Estado' in hdrs or 'Secuencia' in hdrs

    if is_dxc:
        idx_grupo = hdrs.index('Grupo de destinos')
        idx_desc  = hdrs.index('Descripción Grupos de destino')
        idx_zona  = hdrs.index('Tipo de zona')
        idx_dest  = hdrs.index('Destino')
        idx_alm   = hdrs.index('Almacén')
        idx_elem  = hdrs.index('Elemento')
        idx_est   = hdrs.index('Estado') if 'Estado' in hdrs else None
        out_header = ('BLOQUE','Grupo de destinos','Descripción Grupos de destino',
                      'Tipo de zona','Destino','Almacén','Elemento')
    else:
        idx_grupo, idx_desc, idx_zona, idx_dest, idx_alm, idx_elem = 1, 2, 3, 4, 5, 6
        idx_est = None
        out_header = file_header

    print(f"  Formato: {'DXC export' if is_dxc else 'clásico'}")

    by_dia_playa = defaultdict(list)
    tagged = []

    for row in raw_data:
        grupo     = str(row[idx_grupo]) if row[idx_grupo] else ''
        desc      = str(row[idx_desc])  if row[idx_desc]  else ''
        tipo_zona = row[idx_zona] or ''
        destino   = str(row[idx_dest])  if row[idx_dest]  else ''
        almacen   = row[idx_alm]  or ''
        elemento  = str(row[idx_elem])  if row[idx_elem]  else ''
        estado    = str(row[idx_est])   if idx_est is not None and row[idx_est] else 'N/A'

        pb, pd, pp = parse_gd_desc(desc)
        entry = {
            'grupo': grupo, 'desc': desc, 'tipo_zona': tipo_zona,
            'destino': destino, 'almacen': almacen, 'elemento': elemento,
            'bloque': pb, 'dia': pd, 'playa': pp, 'estado': estado,
            'is_weekly': bool(pb and pd and pp),
            'is_sorter': is_sorter_elem(elemento),
            'is_e2': is_e2_elem(elemento),
            '_raw': (None, grupo, desc, tipo_zona, destino, almacen, elemento),
        }
        tagged.append(entry)
        if pb and pd and pp:
            by_dia_playa[(pd, pp)].append(entry)

    return out_header, tagged, by_dia_playa

def load_parrilla(path, sheet_name):
    wb = load_workbook(path, read_only=True)
    if sheet_name not in wb.sheetnames:
        raise ValueError(f"Hoja '{sheet_name}' no encontrada. Disponibles: {', '.join(wb.sheetnames)}")
    ws = wb[sheet_name]
    all_rows = list(ws.iter_rows(values_only=True))
    col = {str(h): i for i, h in enumerate(all_rows[0]) if h}
    records = []
    for row in all_rows[1:]:
        def g(c, d=''):
            i = col.get(c)
            return safe_str(row[i]) if i is not None and row[i] is not None else d
        r = {'dia_playa': g('DIA_PLAYA'), 'playa': g('PLAYA'),
             'dia_salida': g('DIA_SALIDA'), 'cutoff': g('CUTOFF'),
             'dia_salida_new': g('DIA_SALIDA_NEW'),
             'bloque': g('BLOQUE'), 'nomenclatura': g('NOMENCLATURA'),
             'tipo_salida': g('TIPO_SALIDA'), 'id_cluster': g('ID_CLUSTER')}
        if r['playa'] and r['tipo_salida']:
            records.append(r)
    return records


# ─── RAMPA ASSIGNMENT ─────────────────────────────────────────────────────────

def build_day_occ(tagged, dia, new_bloque, bloque_timings,
                   exclude_playas=None, run_occ=None):
    """
    Build {rampa: {pos: grupo}} for dia, merging:
    - positions from tagged GD (timing-filtered)
    - positions already assigned in this run (run_occ) — always conflicting
    """
    ex = exclude_playas or set()
    occ = defaultdict(dict)
    for e in tagged:
        if e['tipo_zona'] != 'POSTEX' or e['dia'] != dia: continue
        if e['playa'] in ex or not e['is_sorter']: continue
        existing_bloque = e['bloque'] or ''
        if new_bloque and existing_bloque and new_bloque in bloque_timings and existing_bloque in bloque_timings:
            if not bloques_overlap(new_bloque, existing_bloque, bloque_timings):
                continue
        r, p = parse_rampa(e['elemento'])
        if r and p:
            occ[r][p] = e['grupo']
    # Always add within-run assignments — no timing bypass
    for rampa, pos_map in (run_occ or {}).items():
        for pos, grupo in pos_map.items():
            occ[rampa][pos] = grupo
    return occ

def find_free_slots(occ, capacity, n_needed):
    all_rampas = sorted(capacity.keys(), key=lambda r: (int(r[1:-1]), r[-1]))
    rampa_free = []
    for rampa in all_rampas:
        if rampa in EXCLUDED_RAMPAS:
            continue   # Reserved for manipulado — skip
        cap = capacity[rampa]
        used = set(occ.get(rampa, {}).keys())
        free = sorted(p for p in range(1, cap+1) if p not in used)
        if free:
            rampa_free.append((rampa, free, len(used) == 0))
    rampa_free.sort(key=lambda x: (-int(x[2]), -len(x[1])))
    assigned, rem = [], n_needed
    for rampa, free, _ in rampa_free:
        if rem <= 0: break
        take = free[:rem]
        assigned.extend((rampa, p) for p in take)
        rem -= len(take)
    return assigned, rem

def playa_is_e2(playa, by_dia_playa):
    """
    Check if a playa is E2-only (all its elements across ALL days are MAN/EXDOCK).
    Returns True only if at least one entry exists and ALL of them are E2 elements.
    """
    all_entries = [e for entries in by_dia_playa.values()
                   for e in entries if e['playa'] == playa]
    if not all_entries:
        return False
    return all(e['is_e2'] or not e['elemento'] for e in all_entries)

def find_best_source_day(playa, orig_dia, by_dia_playa):
    """Find the day with most sorter POSTEX entries for this playa."""
    candidates = []
    for (d, p), entries in by_dia_playa.items():
        if p != playa: continue
        postex_sorter = [e for e in entries if e['tipo_zona'] == 'POSTEX' and e['is_sorter']]
        if postex_sorter:
            candidates.append((d, postex_sorter))
    if not candidates: return None, []
    day_order = ['DOMINGO','LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO']
    orig_idx = day_order.index(orig_dia) if orig_dia in day_order else 0
    candidates.sort(key=lambda x: (abs(day_order.index(x[0])-orig_idx) if x[0] in day_order else 99, -len(x[1])))
    return candidates[0]

def get_slot_structure(postex_entries, sorexp_entries):
    """
    Group POSTEX entries by their shared physical (rampa, pos) slot.
    Multiple destinos can share the same slot — e.g. MEXICO packs 63 store codes
    into a single sorter position. The slot is the true unit of assignment.
    Returns: slots list, sorexp_by_rampa dict, alm_postex str, alm_sorexp str.
    """
    pos_dests   = defaultdict(list)
    rampa_dests = defaultdict(list)
    for e in postex_entries:
        r, p = parse_rampa(e['elemento'])
        if r and p:
            pos_dests[(r, p)].append(e['destino'])
    for e in sorexp_entries:
        elem = e['elemento']
        if re.match(r'^R\d+[A-Z]$', elem):
            rampa_dests[elem].append(e['destino'])
    slots = [{'dests': v, 'orig_rampa': k[0], 'orig_pos': k[1]}
             for k, v in sorted(pos_dests.items())]
    alm_p = postex_entries[0]['almacen'] if postex_entries else 'CONRAM'
    alm_s = sorexp_entries[0]['almacen'] if sorexp_entries else 'SOREXP'
    return slots, dict(rampa_dests), alm_p, alm_s


def assign_especial(orig_dia, orig_playa, new_dia, raw_bloque, id_cluster,
                    by_dia_playa, tagged, capacity, bloque_timings, freed_in_new_dia,
                    run_occ_new_dia=None):
    """
    Assign slots from (orig_dia, orig_playa) to free sorter positions in new_dia,
    respecting bloque timing and preserving the slot→destinos structure.

    The assignment unit is a physical SLOT (rampa, pos) that may hold multiple
    destinos — e.g. MEXICO packs 63 store codes per slot. We find N free slots
    (where N = number of unique positions in the original config) and replicate
    the same destino grouping onto each new slot.
    """
    if playa_is_e2(orig_playa, by_dia_playa):
        return [], {'status':'E2_ROUTE','playa':orig_playa,'dia_orig':orig_dia,'dia_new':new_dia,
                    'msg':'Ruta E2/manual (MAN/EXDOCK) — no pasa por rampas del sorter'}

    new_bloque = None
    if raw_bloque and raw_bloque not in ('#N/A','NO_BLOQUE','?',''):
        new_bloque = raw_bloque
    if not new_bloque:
        new_bloque = resolve_bloque_for_new_day(id_cluster, new_dia, bloque_timings)

    orig        = by_dia_playa.get((orig_dia, orig_playa), [])
    postex_orig = [e for e in orig if e['tipo_zona'] == 'POSTEX' and e['is_sorter']]
    sorexp_orig = [e for e in orig if e['tipo_zona'] == 'SOREXP' and e['is_sorter']]
    source_day  = orig_dia

    if not postex_orig:
        best_day, fallback = find_best_source_day(orig_playa, orig_dia, by_dia_playa)
        if not best_day:
            all_p  = [e for v in by_dia_playa.values() for e in v if e['playa'] == orig_playa]
            status = 'E2_ROUTE' if all_p else 'NO_CONFIG'
            msg    = ('Ruta E2/manual — elementos no son rampas estándar' if all_p
                      else 'Sin configuración GD en ningún día')
            return [], {'status':status,'playa':orig_playa,'dia_orig':orig_dia,'dia_new':new_dia,'msg':msg}
        postex_orig = fallback
        sorexp_orig = [e for e in by_dia_playa.get((best_day, orig_playa),[])
                       if e['tipo_zona'] == 'SOREXP' and e['is_sorter']]
        source_day  = best_day

    # Build slot structure: N unique (rampa,pos) → list of destinos each
    slots, _, alm_postex, alm_sorexp = get_slot_structure(postex_orig, sorexp_orig)
    n_slots = len(slots)

    freed    = freed_in_new_dia & {e['playa'] for e in tagged if e['dia'] == new_dia}
    occ      = build_day_occ(tagged, new_dia, new_bloque, bloque_timings,
                             exclude_playas=freed, run_occ=run_occ_new_dia)
    assigned, unmet = find_free_slots(occ, capacity, n_slots)

    playa_tag = orig_playa.replace('ESPANA_','')[:6].replace('_','')
    new_grupo = f"ESP_{new_dia[:3]}_{playa_tag}"[:18]
    fb_note   = f" [datos {source_day}]" if source_day != orig_dia else ""
    new_desc  = f"[B2B] {new_bloque or '?'}_{new_dia}_{orig_playa} (ESPECIAL{fb_note})"

    rows_out, rampas_used = [], defaultdict(list)
    n_destinos_total = 0
    for slot, (new_rampa, new_pos) in zip(slots, assigned):
        for destino in slot['dests']:
            rows_out.append((None, new_grupo, new_desc, 'POSTEX',
                             destino, alm_postex, postex_elem(new_rampa, new_pos)))
            rows_out.append((None, new_grupo, new_desc, 'SOREXP',
                             destino, alm_sorexp, new_rampa))
        rampas_used[new_rampa].append(new_pos)
        n_destinos_total += len(slot['dests'])

    return rows_out, {
        'status':      'OK' if unmet == 0 else 'PARTIAL',
        'playa':        orig_playa,
        'dia_orig':     orig_dia,
        'dia_new':      new_dia,
        'bloque_new':   new_bloque or '?',
        'n_slots':      n_slots,
        'n_destinos':   n_destinos_total,
        'n_assigned':   n_slots - unmet,
        'n_unmet':      unmet,
        'grupo_new':    new_grupo,
        'desc_new':     new_desc,
        'source_day':   source_day,
        'rampas':       {r: sorted(p) for r, p in rampas_used.items()},
        'n_rows':       len(rows_out),
    }


# ─── PROCESS ──────────────────────────────────────────────────────────────────

def process(parrilla_records, tagged, by_dia_playa, capacity, bloque_timings):
    canceladas, especiales, habituales = {}, {}, {}

    for r in parrilla_records:
        playa, tipo = r['playa'], r['tipo_salida']
        dia_orig = r['dia_salida']
        dia_new  = r['dia_salida_new'] or dia_orig
        if tipo == 'CANCELADA':
            canceladas[(dia_orig, playa)] = r
        elif tipo == 'ESPECIAL DIA CAMBIO' and dia_new and dia_new != dia_orig:
            especiales[(dia_orig, playa)] = (dia_new, r)
        elif tipo in ('HABITUAL','REGULAR','ESPECIAL CUTOFF'):
            habituales[(dia_orig, playa)] = r

    freed_per_day = defaultdict(set)
    for (dia, playa) in canceladas: freed_per_day[dia].add(playa)
    for (dia_orig, playa) in especiales: freed_per_day[dia_orig].add(playa)

    output_rows = []
    rem_cancel = rem_moved = kept = kept_nw = 0

    for entry in tagged:
        if not entry['is_weekly']:
            output_rows.append(entry['_raw']); kept_nw += 1; continue
        key = (entry['dia'], entry['playa'])
        if key in canceladas: rem_cancel += 1; continue
        if key in especiales: rem_moved  += 1; continue
        output_rows.append(entry['_raw']); kept += 1

    # Accumulate slots assigned in this run per new_dia so subsequent
    # playas don't collide (e.g. MEXICO + MEXICO_2 both going to MIERCOLES)
    run_occ = defaultdict(dict)   # dia -> {rampa: {pos: grupo}}

    results, added = [], 0
    for (dia_orig, playa), (dia_new, record) in especiales.items():
        new_rows, info = assign_especial(
            dia_orig, playa, dia_new,
            record['bloque'], record['id_cluster'],
            by_dia_playa, tagged, capacity, bloque_timings,
            freed_per_day.get(dia_new, set()),
            run_occ.get(dia_new, {}),
        )
        # Register newly assigned positions so next playa in same run sees them
        for row in new_rows:
            if row[3] == 'POSTEX':
                import re as _re
                m = _re.match(r'^R0*(\d+)_([A-Z])-(\d+)$', str(row[6] or ''))
                if m:
                    rampa = f"R{int(m.group(1)):02d}{m.group(2)}"
                    pos   = int(m.group(3))
                    run_occ.setdefault(dia_new, {}).setdefault(rampa, {})[pos] = info.get('grupo_new','ESP')
        output_rows.extend(new_rows)
        added += len(new_rows)
        results.append(info)

    return output_rows, {
        'n_canceladas': len(canceladas), 'n_especiales': len(especiales),
        'rows_removed_cancelled': rem_cancel, 'rows_removed_moved': rem_moved,
        'rows_kept': kept, 'rows_kept_non_weekly': kept_nw,
        'rows_added_especial': added, 'total_output_rows': len(output_rows),
        'canceladas': canceladas, 'assignment_results': results,
    }


# ─── WRITE EXCEL ──────────────────────────────────────────────────────────────

def write_gd(output_rows, header, out_path):
    wb = Workbook(); ws = wb.active; ws.title = 'Hoja1'
    hf = PatternFill('solid', start_color='1A1A1A')
    for ci, w in enumerate([8,15,65,10,15,10,15], 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci, h in enumerate(header, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.font = Font(name='Arial', bold=True, size=10, color='FFFFFF')
        c.fill = hf; c.alignment = Alignment(horizontal='center')
    ws.row_dimensions[1].height = 18
    thin = Border(bottom=Side(style='thin', color='EBEBEB'))
    df = Font(name='Arial', size=9)
    for ri, row in enumerate(output_rows, 2):
        _, grupo, desc, tipo_zona, destino, almacen, elemento = row
        for ci, val in enumerate([f'=MID(B{ri},3,2)' if grupo else '', grupo, desc,
                                   tipo_zona, destino, almacen, elemento], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = df; c.border = thin
            if ci == 1: c.alignment = Alignment(horizontal='center')
    ws.freeze_panes = 'A2'
    wb.save(out_path)


# ─── WRITE HTML ───────────────────────────────────────────────────────────────

def write_html(summary, semana, out_path):
    _DAY_ORDER = ['DOMINGO','LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO']
    def _day_key(r): return (_DAY_ORDER.index(r['dia_new']) if r['dia_new'] in _DAY_ORDER else 99, r['playa'])

    ok      = sorted([r for r in summary['assignment_results'] if r['status'] == 'OK'],       key=_day_key)
    partial = sorted([r for r in summary['assignment_results'] if r['status'] == 'PARTIAL'],  key=_day_key)
    e2      = sorted([r for r in summary['assignment_results'] if r['status'] == 'E2_ROUTE'], key=_day_key)
    nocfg   = sorted([r for r in summary['assignment_results'] if r['status'] == 'NO_CONFIG'],key=_day_key)
    cancels = list(summary['canceladas'].items())
    n_warn  = len(partial) + len(nocfg)
    now     = datetime.now().strftime("%d/%m/%Y %H:%M")

    def b(cls, txt): return f'<span class="b {cls}">{txt}</span>'

    def cancel_rows():
        out, seen = '', defaultdict(list)
        for (dia, playa), _ in cancels: seen[dia].append(playa)
        for dia in ['DOMINGO','LUNES','MARTES','MIERCOLES','JUEVES','VIERNES','SABADO']:
            for playa in seen.get(dia, []):
                out += f'<tr><td>{b("bd",dia)}</td><td>{playa}</td>{b("bc","CANCELADA")}</td><td class="mt">—</td></tr>\n'
        return out

    def ok_rows():
        out = ''
        for r in ok:
            fb  = f' <em class="mt">(datos de {r["source_day"]})</em>' \
                  if r.get('source_day') and r['source_day'] != r['dia_orig'] else ''
            blq = f'<span class="mono">{r["bloque_new"]}</span>' if r.get('bloque_new') and r['bloque_new'] != '?' else '<span class="mt">—</span>'
            rstr = ' · '.join(f'{k}[{",".join(str(p) for p in v)}]' for k,v in sorted(r['rampas'].items()))
            out += (f'<tr><td>{b("bs",r["dia_new"])}</td><td>{r["playa"]}{fb}</td>'
                    f'<td>{b("bd",r["dia_orig"])}</td><td>{blq}</td>'
                    f'<td class="mono mt small">{rstr}</td></tr>\n')
        return out

    def warn_rows():
        out = ''
        for r in partial:
            out += (f'<tr><td>{b("bd",r["dia_orig"])}</td><td>{r["playa"]}</td>'
                    f'<td>{b("bw","→ "+r["dia_new"])}</td><td colspan="2" class="wt">'
                    f'⚠ Solo {r["n_assigned"]}/{r["n_destinos"]} posiciones libres</td></tr>\n')
        for r in e2:
            out += (f'<tr><td>{b("bd",r["dia_orig"])}</td><td>{r["playa"]}</td>'
                    f'<td>{b("bi","→ "+r["dia_new"])}</td><td colspan="2" class="it">'
                    f'🔀 Ruta E2/manual — no pasa por rampas del sorter</td></tr>\n')
        for r in nocfg:
            out += (f'<tr><td>{b("bd",r["dia_orig"])}</td><td>{r["playa"]}</td>'
                    f'<td>{b("bw","→ "+r["dia_new"])}</td><td colspan="2" class="wt">'
                    f'❌ Sin configuración GD — revisar con equipo</td></tr>\n')
        return out

    html = f'''<!DOCTYPE html>
<html lang="es"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Sorter VDL B2B — {semana}</title>
<style>
*,*::before,*::after{{box-sizing:border-box;margin:0;padding:0}}
body{{font-family:"Trebuchet MS",Arial,sans-serif;font-weight:300;background:#f0f0f0;color:#1a1a1a;font-size:13px;line-height:1.5}}
.rpt{{max-width:980px;margin:0 auto;background:#fff;min-height:100vh}}
.rh{{background:#000;color:#fff;padding:36px 48px 28px}}
.ey{{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#888;margin-bottom:10px}}
.rh h1{{font-size:26px;font-weight:300;letter-spacing:-.02em;margin-bottom:4px}}
.rh .sub{{font-size:13px;color:#888}}
.meta{{display:flex;gap:28px;margin-top:20px;padding-top:16px;border-top:1px solid #333}}
.mi label{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#666;display:block}}
.mi span{{font-size:12px;color:#ccc}}
.kpis{{background:#000;display:grid;grid-template-columns:repeat(5,1fr);border-top:1px solid #222}}
.kpi{{padding:18px 20px;border-right:1px solid #222}}.kpi:last-child{{border-right:none}}
.kv{{font-size:28px;font-weight:300;color:#fff;letter-spacing:-.03em;font-family:monospace;line-height:1;margin-bottom:4px}}
.kl{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#555}}
.kv.red{{color:#ef4444}}.kv.amb{{color:#f59e0b}}.kv.grn{{color:#22c55e}}
.cnt{{padding:0 48px 48px}}
.sec{{margin-top:34px}}
.sn{{font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:#999;margin-bottom:4px}}
.sec h2{{font-size:17px;font-weight:400;letter-spacing:-.01em;margin-bottom:3px}}
.sd{{font-size:12px;color:#888;margin-bottom:14px}}
.sb{{display:grid;grid-template-columns:repeat(3,1fr);gap:1px;background:#ebebeb;border:1px solid #ebebeb;border-radius:3px;overflow:hidden;margin-bottom:24px}}
.sc{{background:#fff;padding:14px 18px;text-align:center}}
.sv{{font-size:22px;font-weight:300;color:#1a1a1a;font-family:monospace;letter-spacing:-.02em}}
.sl{{font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:#999;margin-top:2px}}
table{{width:100%;border-collapse:collapse;font-size:12px}}
thead tr{{border-bottom:1px solid #1a1a1a}}
thead th{{text-align:left;font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:#999;padding:7px 10px 7px 0;font-weight:400}}
tbody tr{{border-bottom:1px solid #ebebeb}}tbody tr:last-child{{border-bottom:none}}
tbody td{{padding:8px 10px 8px 0;vertical-align:top}}
.mt{{color:#888}}.wt{{color:#b45309}}.it{{color:#1d4ed8}}
.mono{{font-family:monospace;font-size:11px}}.small{{font-size:11px}}em.mt{{font-style:normal}}
.b{{display:inline-block;font-size:9px;letter-spacing:.06em;text-transform:uppercase;padding:2px 6px;border-radius:2px;font-family:monospace;white-space:nowrap}}
.bd{{background:#f3f4f6;color:#374151}}.bc{{background:#fee2e2;color:#b91c1c}}
.bs{{background:#dbeafe;color:#1d4ed8}}.bw{{background:#fef3c7;color:#92400e}}
.bi{{background:#eff6ff;color:#1d4ed8}}
.ib{{border-left:3px solid #3b82f6;background:#eff6ff;padding:11px 15px;margin-bottom:10px;font-size:12px;color:#1e40af}}
.wb{{border-left:3px solid #f59e0b;background:#fffbeb;padding:11px 15px;margin-bottom:10px;font-size:12px}}
.wb strong{{color:#92400e}}
.rf{{background:#000;color:#555;font-size:9px;letter-spacing:.08em;text-transform:uppercase;padding:14px 48px;display:flex;justify-content:space-between;font-family:monospace}}
.np{{text-align:center;padding:22px 0 36px}}
.bp{{font-family:monospace;background:#000;color:#fff;border:none;padding:9px 28px;cursor:pointer;text-transform:uppercase;letter-spacing:.1em;font-size:10px}}
.bp:hover{{background:#333}}
@media print{{
  @page{{size:A4 portrait;margin:12mm 14mm 14mm 14mm}}
  *{{-webkit-print-color-adjust:exact!important;print-color-adjust:exact!important}}
  body{{background:#fff}}.rpt{{max-width:100%}}
  .rh,.kpis,.sec{{break-inside:avoid}}tr{{break-inside:avoid}}.np{{display:none!important}}
}}
</style></head><body>
<div class="rpt">
<div class="rh">
  <div class="ey">Mango · Logística · VDL B2B</div>
  <h1>Configuración Sorter — {semana}</h1>
  <div class="sub">Semana especial · asignación con control de solapamiento de bloques horarios</div>
  <div class="meta">
    <div class="mi"><label>Generado</label><span>{now}</span></div>
    <div class="mi"><label>Semana</label><span>{semana}</span></div>
    <div class="mi"><label>Filas output GD</label><span>{summary["total_output_rows"]:,}</span></div>
    <div class="mi"><label>Canceladas</label><span>{summary["n_canceladas"]}</span></div>
    <div class="mi"><label>Cambios de día</label><span>{summary["n_especiales"]}</span></div>
  </div>
</div>
<div class="kpis">
  <div class="kpi"><div class="kv red">{summary["n_canceladas"]}</div><div class="kl">Canceladas</div></div>
  <div class="kpi"><div class="kv grn">{len(ok)}</div><div class="kl">Asignadas OK</div></div>
  <div class="kpi"><div class="kv">{len(e2)}</div><div class="kl">Rutas E2</div></div>
  <div class="kpi"><div class="kv {"amb" if n_warn else "grn"}">{n_warn}</div><div class="kl">Revisar</div></div>
  <div class="kpi"><div class="kv">{summary["total_output_rows"]:,}</div><div class="kl">Filas GD</div></div>
</div>
<div class="cnt">

<div class="sec">
  <div class="sn">00 — Resumen</div>
  <h2>Impacto en el fichero GRUPO_DESTINOS</h2>
  <div class="sd">La asignación de posiciones respeta los horarios de bloque: dos destinos pueden compartir rampa si sus ventanas de liberación/desactivación no se solapan.</div>
  <div class="sb">
    <div class="sc"><div class="sv">{summary["rows_removed_cancelled"]:,}</div><div class="sl">Filas eliminadas</div></div>
    <div class="sc"><div class="sv">{summary["rows_added_especial"]:,}</div><div class="sl">Filas añadidas</div></div>
    <div class="sc"><div class="sv">{summary["total_output_rows"]:,}</div><div class="sl">Total filas output</div></div>
  </div>
</div>

<div class="sec">
  <div class="sn">01 — Salidas canceladas</div>
  <h2>Destinos eliminados del sorter esta semana</h2>
  <div class="sd">{summary["n_canceladas"]} salidas canceladas — {summary["rows_removed_cancelled"]:,} filas eliminadas.</div>
  <table>
    <thead><tr><th>Día original</th><th>Destino</th><th>Estado</th><th>Nota</th></tr></thead>
    <tbody>{cancel_rows()}</tbody>
  </table>
</div>

<div class="sec">
  <div class="sn">02 — Asignación automática de rampas (con control de horarios)</div>
  <h2>Destinos reasignados a nuevo día</h2>
  <div class="sd">{len(ok)} destinos asignados a posiciones libres en el nuevo día.
  El bloque se deriva del ID_CLUSTER del destino en el nuevo día.
  Si el origen tenía bloque desconocido (#N/A), se indica el bloque calculado.</div>
  {'<p class="mt" style="margin:10px 0">— Sin reasignaciones automáticas —</p>' if not ok else f"""
  <table>
    <thead><tr><th>Nuevo día</th><th>Destino</th><th>Día orig.</th><th>Bloque</th><th>Rampas y posiciones</th></tr></thead>
    <tbody>{ok_rows()}</tbody>
  </table>"""}
</div>

<div class="sec">
  <div class="sn">03 — Rutas E2, sin configuración y asignaciones parciales</div>
  <h2>Casos que requieren atención</h2>
  <div class="sd">
    {len(e2)} rutas E2 (MAN/EXDOCK — sin acción en sorter) ·
    {len(nocfg)} sin config GD · {len(partial)} parciales.
  </div>
  {'<p class="mt" style="margin:10px 0">✓ Sin casos pendientes.</p>' if not (n_warn + len(e2)) else
   ''.join([f'<div class="ib">🔀 <strong>{r["playa"]}</strong> ({r["dia_orig"]} → {r["dia_new"]}): {r["msg"]}</div>' for r in e2])
   + ''.join([f'<div class="wb">❌ <strong>{r["playa"]}</strong> ({r["dia_orig"]} → {r["dia_new"]}): {r["msg"]}</div>' for r in nocfg])
   + ''.join([f'<div class="wb">⚠ <strong>{r["playa"]}</strong>: Solo {r["n_assigned"]}/{r["n_destinos"]} posiciones en {r["dia_new"]}</div>' for r in partial])
   + f"""
  <table>
    <thead><tr><th>Día orig.</th><th>Destino</th><th>Nuevo día</th><th>Tipo</th><th>Detalle</th></tr></thead>
    <tbody>{warn_rows()}</tbody>
  </table>"""}
</div>

</div>
<div class="rf">
  <span>© Mango · Logística VDL B2B · Estrictamente confidencial</span>
  <span>Generado {datetime.now().strftime("%d/%m/%Y")}</span>
</div>
</div>
<div class="np"><button class="bp" onclick="window.print()">Imprimir / Exportar PDF</button></div>
</body></html>'''

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)


# ─── MAIN ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 5:
        print(__doc__); sys.exit(1)

    parrilla_path, gd_path, cap_path, sheet_name = sys.argv[1:5]
    semana = sys.argv[5] if len(sys.argv) > 5 else sheet_name.upper()

    out_dir  = '/home/claude'
    ts       = datetime.now().strftime('%Y%m%d_%H%M')
    gd_out   = os.path.join(out_dir, f'GRUPO_DESTINOS_{semana}_{ts}.xlsx')
    html_out = os.path.join(out_dir, f'resumen_sorter_{semana}_{ts}.html')

    print(f"Cargando capacidad: {cap_path}")
    capacity = load_capacity(cap_path)
    print(f"  → {len(capacity)} rampas")

    print(f"Cargando timings de bloques desde parrilla: {parrilla_path}")
    bloque_timings = load_bloque_timings(parrilla_path)
    print(f"  → {len(bloque_timings)} bloques con horario")

    print(f"Cargando GRUPO_DESTINOS: {gd_path}")
    gd_header, tagged, by_dia_playa = load_grupo_destinos(gd_path)
    print(f"  → {len(tagged):,} filas, {len(by_dia_playa)} combinaciones día×playa")

    print(f"Cargando parrilla '{sheet_name}': {parrilla_path}")
    parrilla = load_parrilla(parrilla_path, sheet_name)
    print(f"  → {len(parrilla)} destinos")

    print("Procesando y asignando rampas (con control horario de bloques)...")
    output_rows, summary = process(parrilla, tagged, by_dia_playa, capacity, bloque_timings)

    ok      = [r for r in summary['assignment_results'] if r['status'] == 'OK']
    partial = [r for r in summary['assignment_results'] if r['status'] == 'PARTIAL']
    e2      = [r for r in summary['assignment_results'] if r['status'] == 'E2_ROUTE']
    nocfg   = [r for r in summary['assignment_results'] if r['status'] == 'NO_CONFIG']

    print(f"\n{'='*65}")
    print(f"RESUMEN {semana}:")
    print(f"  Canceladas:             {summary['n_canceladas']} ({summary['rows_removed_cancelled']:,} filas)")
    print(f"  Cambios de día:         {summary['n_especiales']}")
    print(f"    → Asignados OK:       {len(ok)}")
    print(f"    → Rutas E2:           {len(e2)}")
    print(f"    → Sin config GD:      {len(nocfg)}")
    print(f"    → Parciales:          {len(partial)}")
    print(f"  Filas añadidas:         {summary['rows_added_especial']:,}")
    print(f"  Filas output total:     {summary['total_output_rows']:,}")
    print(f"{'='*65}")

    if ok:
        print("\nAsignaciones completadas:")
        for r in ok:
            fb   = f" [datos {r['source_day']}]" if r.get('source_day') and r['source_day'] != r['dia_orig'] else ''
            blq  = r.get('bloque_new','?')
            rstr = ', '.join(f"{k}({len(v)}p)" for k,v in sorted(r['rampas'].items()))
            print(f"  {r['playa']:42s} {r['dia_orig']}→{r['dia_new']} [{blq}]{fb}  {rstr}")

    if e2:
        print("\nRutas E2 (sin cambio en GD):")
        for r in e2:
            print(f"  {r['playa']:42s} {r['dia_orig']}→{r['dia_new']}  🔀 E2/manual")

    if nocfg:
        print("\n❌ Sin config GD:")
        for r in nocfg:
            print(f"  {r['playa']:42s} sin datos en ningún día")

    if partial:
        print("\n⚠ Parciales:")
        for r in partial:
            print(f"  {r['playa']:42s} {r['n_assigned']}/{r['n_destinos']} posiciones")

    print(f"\nEscribiendo GD  → {gd_out}")
    write_gd(output_rows, gd_header, gd_out)
    print(f"Escribiendo HTML → {html_out}")
    write_html(summary, semana, html_out)

    shutil.copy(gd_out,   f'/mnt/user-data/outputs/GRUPO_DESTINOS_{semana}.xlsx')
    shutil.copy(html_out, f'/mnt/user-data/outputs/resumen_sorter_{semana}.html')
    shutil.copy(__file__,  '/mnt/user-data/outputs/process_parrilla.py')
    print(f"\n✓ Listo.")

if __name__ == '__main__':
    main()
