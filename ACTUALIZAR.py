#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACTUALIZAR.py — Actualizador dashboards Grupo J&D
==================================================
Fuente de verdad: Cómo vamos Grupo Empresarial.xlsx (Desktop LOCAL)
Periodos CPP:     CALENDARIO MAESTRO (hoja en el mismo Excel)

Qué hace:
  1. Lee hojas JULIO y AGOSTO del Excel (datos diarios por modelo/studio)
  2. Lee CALENDARIO MAESTRO (períodos por plataforma)
  3. Reconstruye ALIADOS y ESTUDIO desde cero para Julio y Agosto
  4. Reconstruye GRUPO_PERIODOS / PERIODOS desde el calendario real
  5. Recalcula TOP_EST (por estudio) y TOP20G (grupo completo) para ambos meses
  6. Inyecta 'Agosto' en var MESES y var DMES de los 6 HTML
  7. Actualiza timestamp

FÓRMULA VERIFICADA: col_F4F(d) = 19 + (31-d)*5  [5 plats/día, LOCAL]
NO commit/push hasta validación del usuario.
"""

import openpyxl, re, base64, json, os
from datetime import datetime, date

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CV_PATH = os.path.join(SCRIPT_DIR, '..', 'Centro de Gestión Estratégica Grupo J&D',
                       'COMO VAMOS GRUPO', 'Cómo vamos Grupo Empresarial.xlsx')
DASHBOARDS = {
    'grupo':  os.path.join(SCRIPT_DIR, 'grupo579780',                     'index.html'),
    'erika':  os.path.join(SCRIPT_DIR, 'erika868527',                     'index.html'),
    'fabio':  os.path.join(SCRIPT_DIR, 'fabio473013',                     'index.html'),
    'fornax': os.path.join(SCRIPT_DIR, 'estudios', 'fornax-studios345929', 'index.html'),
    'gold':   os.path.join(SCRIPT_DIR, 'estudios', 'goldonline078939',     'index.html'),
    'cyv':    os.path.join(SCRIPT_DIR, 'estudios', 'cyv-studios837357',    'index.html'),
}

MES     = 'Julio'
MES_AGO = 'Agosto'
PLATS   = ['F4F', 'SC', 'CB', 'CAM', 'STR']

# Meses ordenados con días
ALL_MESES   = ['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto']
DMES_VALS   = {'Enero':31,'Febrero':28,'Marzo':31,'Abril':30,
               'Mayo':31,'Junio':30,'Julio':31,'Agosto':31}
MES_NUM     = {1:'Enero',2:'Febrero',3:'Marzo',4:'Abril',5:'Mayo',6:'Junio',
               7:'Julio',8:'Agosto',9:'Septiembre',10:'Octubre',11:'Noviembre',12:'Diciembre'}
MES_ABREV   = {'Ene':'Enero','Feb':'Febrero','Mar':'Marzo','Abr':'Abril','May':'Mayo',
               'Jun':'Junio','Jul':'Julio','Ago':'Agosto','Sep':'Septiembre',
               'Oct':'Octubre','Nov':'Noviembre','Dic':'Diciembre'}
MONTH_ESP   = {1:'Ene',2:'Feb',3:'Mar',4:'Abr',5:'May',6:'Jun',
               7:'Jul',8:'Ago',9:'Sep',10:'Oct',11:'Nov',12:'Dic'}

# Plataforma → columna en CALENDARIO MAESTRO (0-indexed, R4 header)
CAL_COL = {'F4F':0, 'Chaturbate':1, 'Stripchat':2, 'CamSoda':3, 'Streamate':4}
# Display name → internal key
CAL_PLAT_DISPLAY = {
    'F4F':        ('F4F','F4F'),
    'Chaturbate': ('Chaturbate','CB'),
    'Stripchat':  ('Stripchat','SC'),
    'CamSoda':    ('CamSoda','CAM'),
    'Streamate':  ('Streamate','STR'),
}

TODAY = date.today()


# ═══════════════════════════════════════════════════════════════════
# FÓRMULA DE COLUMNA
# ═══════════════════════════════════════════════════════════════════
def day_col(d):
    return 19 + (31 - d) * 5


# ═══════════════════════════════════════════════════════════════════
# BASE64 HELPERS
# ═══════════════════════════════════════════════════════════════════
def b64enc(obj):
    return base64.b64encode(
        json.dumps(obj, ensure_ascii=False).encode('utf-8')
    ).decode('ascii')

def b64dec(s):
    return json.loads(base64.b64decode(s).decode('utf-8'))

def get_var(html, varname):
    for q in ['"', "'"]:
        m = re.search(rf"var {varname}\s*=\s*_b64dec\({q}([^{q}]+){q}\)", html)
        if m:
            try:
                return b64dec(m.group(1)), q
            except Exception:
                pass
    return None, None

def set_var(html, varname, data, quote):
    new_b64 = b64enc(data)
    pattern = rf"(var {varname}\s*=\s*_b64dec\(){quote}[^{quote}]+{quote}(\))"
    result, n = re.subn(pattern, rf'\g<1>{quote}{new_b64}{quote}\g<2>', html, count=1)
    if n == 0:
        print(f"  ⚠  No se encontró var {varname} (quote={quote})")
    return result


# ═══════════════════════════════════════════════════════════════════
# INYECCIÓN DE MESES EN JS
# ═══════════════════════════════════════════════════════════════════
def patch_meses_dmes(html, meses_list):
    """
    Reemplaza var MESES=[...] y var DMES={...} en el HTML
    para incluir todos los meses en meses_list.
    """
    meses_js   = "var MESES=[" + ",".join(f"'{m}'" for m in meses_list) + "];"
    dmes_parts = ",".join(f"{m}:{DMES_VALS[m]}" for m in meses_list if m in DMES_VALS)
    dmes_js    = "var DMES={" + dmes_parts + "};"

    html = re.sub(r"var MESES=\[[^\]]*\];", meses_js, html, count=1)
    html = re.sub(r"var DMES=\{[^}]*\};",  dmes_js,  html, count=1)
    return html


# ═══════════════════════════════════════════════════════════════════
# TIMESTAMP
# ═══════════════════════════════════════════════════════════════════
def inject_timestamp(html, ts_str):
    badge = f'<span class="nav-update">Última actualización: {ts_str}</span>'
    css   = '.nav-update{display:block;font-size:9px;color:var(--muted);margin-top:1px}'
    if '.nav-update' not in html:
        html = html.replace('</style>', css + '\n</style>', 1)
    existing = re.search(r'<span class="nav-update">.*?</span>', html)
    if existing:
        html = html[:existing.start()] + badge + html[existing.end():]
    else:
        html = html.replace('</div></nav>', badge + '</div></nav>', 1)
    return html


# ═══════════════════════════════════════════════════════════════════
# LEER EXCEL — HOJA DE DATOS DIARIOS
# ═══════════════════════════════════════════════════════════════════
def read_cv_sheet(wb, sheet):
    """
    Lee una hoja del workbook en RAM (iter_rows) → evita timeouts.
    Devuelve: cv = {(model, studio): {day_int: {plat: float|None}}}, last_day
    """
    if sheet not in wb.sheetnames:
        print(f"  ⚠  Hoja '{sheet}' no encontrada.")
        return {}, 0
    ws = wb[sheet]
    data = tuple(ws.iter_rows(values_only=True))
    MAX_COL = ws.max_column
    print(f"  📋 [{sheet}] {len(data)}r × {MAX_COL}c")

    # Verify structure
    ok = str(data[2][18] or '') == '31' and 'lirt' in str(data[3][18] or '').lower()
    print(f"  📋 Col19: día={data[2][18]!r} plat={data[3][18]!r}  {'✅' if ok else '❌'}")

    IGNORAR_STUDIOS = {
        'CAMSODA','CAMSODA POR QUINCENA','CHATURBATE','FLIRT4FREE',
        'STREAMATE','STREAMATE POR QUINCENA','STRIPCHAT','STRIPCHAT POR QUINCENA',
        'TOTAL','TOTALES','CAMSODA POR QUINCENA','STREAMATE POR QUINCENA','STRIPCHAT POR QUINCENA',
    }
    cv = {}; last_day_excel = 0
    for ri in range(4, len(data)):
        row    = data[ri]
        model  = str(row[0] or '').strip()
        studio = str(row[1] or '').strip()
        if not model or not studio:
            continue
        if studio.upper() in IGNORAR_STUDIOS or studio.upper() in {s.upper() for s in IGNORAR_STUDIOS}:
            continue
        if model.upper() in {'TOTAL','TOTALES','MODELO'}:
            continue
        entry = {}
        for d in range(1, 32):
            c0 = day_col(d)
            if c0 > MAX_COL:
                break
            day_data = {}; has = False
            for pi, p in enumerate(PLATS):
                ci = c0 + pi - 1
                if ci >= len(row):
                    day_data[p] = None; continue
                v = row[ci]
                if isinstance(v, (int, float)) and float(v) > 0:
                    val = float(v)
                    if p == 'STR':          # Streamate: USD → créditos (÷ 0.05 = × 20)
                        val = val / 0.05
                    day_data[p] = val; has = True
                    last_day_excel = max(last_day_excel, d)
                else:
                    day_data[p] = None
            if has:
                entry[d] = day_data
        if entry:
            cv[(model, studio)] = entry
    print(f"  📅 Último día: {last_day_excel} | Modelos: {len(cv)}")
    return cv, last_day_excel


# ═══════════════════════════════════════════════════════════════════
# LEER CALENDARIO MAESTRO
# ═══════════════════════════════════════════════════════════════════
def read_calendario_maestro(wb):
    """
    Lee el CALENDARIO MAESTRO y devuelve:
    {
      'F4F':        [{'label':'27 Jul – 09 Ago','start':'27/07/2026','end':'09/08/2026'}, ...],
      'Chaturbate': [...],
      'Stripchat':  [...],
      'CamSoda':    [...],
      'Streamate':  [...],
    }
    Solo incluye periodos cuya fecha de INICIO <= hoy.
    """
    if 'CALENDARIO MAESTRO' not in wb.sheetnames:
        print("  ⚠  Hoja 'CALENDARIO MAESTRO' no encontrada.")
        return {}

    ws   = wb['CALENDARIO MAESTRO']
    rows = list(ws.iter_rows(values_only=True))

    # Row 4 (index 3) = headers, Row 5+ (index 4+) = period ranges
    col_order = ['F4F', 'Chaturbate', 'Stripchat', 'CamSoda', 'Streamate']
    result    = {k: [] for k in col_order}

    def parse_range(cell_str):
        """'27/07/2026 - 09/08/2026' → (start_date, end_date, start_str, end_str)"""
        if not cell_str or not isinstance(cell_str, str):
            return None
        cell_str = cell_str.strip()
        m = re.match(r'(\d{2}/\d{2}/\d{4})\s*[-–]\s*(\d{2}/\d{2}/\d{4})', cell_str)
        if not m:
            return None
        s_str, e_str = m.group(1), m.group(2)
        try:
            sd = datetime.strptime(s_str, '%d/%m/%Y').date()
            ed = datetime.strptime(e_str, '%d/%m/%Y').date()
            return sd, ed, s_str, e_str
        except ValueError:
            return None

    def fmt_label(sd, ed):
        """27/07/2026 → '27 Jul – 09 Ago'"""
        return f"{sd.day:02d} {MONTH_ESP[sd.month]} – {ed.day:02d} {MONTH_ESP[ed.month]}"

    for row in rows[4:]:  # skip header rows
        for ci, plat in enumerate(col_order):
            if ci >= len(row):
                continue
            cell_val = row[ci]
            if cell_val is None:
                continue
            parsed = parse_range(str(cell_val))
            if parsed is None:
                continue
            sd, ed, s_str, e_str = parsed
            # Only include periods whose START is <= today
            if sd > TODAY:
                continue
            result[plat].append({
                'label': fmt_label(sd, ed),
                'start': s_str,
                'end':   e_str,
                'models': {},
                'total':  0,
            })

    for plat, perds in result.items():
        print(f"  📅 {plat}: {len(perds)} periodos")
    return result


# ═══════════════════════════════════════════════════════════════════
# CÁLCULO DE TOTALES POR PERÍODO
# ═══════════════════════════════════════════════════════════════════
def _sum_period_from_aliados(all_mods_by_month, s_str, e_str, plat_key):
    """
    Suma producción de plat_key para el rango s_str→e_str usando
    all_mods_by_month = {mes_name: {model: {day_int: {plat: val}}}}.
    Devuelve {model: total}.
    """
    try:
        sd = datetime.strptime(s_str, '%d/%m/%Y').date()
        ed = datetime.strptime(e_str, '%d/%m/%Y').date()
    except ValueError:
        return {}

    new_models = {}
    # Iterate month by month within the period
    cur_month  = sd.replace(day=1)
    while cur_month <= ed.replace(day=1):
        mes_name = MES_NUM.get(cur_month.month)
        if mes_name and mes_name in all_mods_by_month:
            # Day range for this month within the period
            d_start = sd.day if (cur_month.year == sd.year and cur_month.month == sd.month) else 1
            import calendar as _cal
            last_in_month = _cal.monthrange(cur_month.year, cur_month.month)[1]
            d_end   = ed.day if (cur_month.year == ed.year and cur_month.month == ed.month) else last_in_month

            for model, days in all_mods_by_month[mes_name].items():
                for d in range(d_start, d_end + 1):
                    v = (days.get(d) or {}).get(plat_key) or 0
                    if v:
                        new_models[model] = new_models.get(model, 0) + v

        # Next month
        if cur_month.month == 12:
            cur_month = cur_month.replace(year=cur_month.year + 1, month=1)
        else:
            cur_month = cur_month.replace(month=cur_month.month + 1)

    return new_models


def _build_all_mods_by_month(aliados, aliados_key=None):
    """
    Construye {mes_name: {model: {day_int: {plat:val}}}} desde ALIADOS.
    Si aliados_key no es None, restringe a ese studio.
    """
    all_mods = {}
    scope = aliados if aliados_key is None else {aliados_key: aliados.get(aliados_key, {})}
    for ak, ainfo in scope.items():
        for mes, md in (ainfo.get('data') or {}).items():
            if mes not in all_mods:
                all_mods[mes] = {}
            for model, days in (md.get('modelos') or {}).items():
                if model not in all_mods[mes]:
                    all_mods[mes][model] = {}
                for dk, dv in days.items():
                    if dv:
                        all_mods[mes][model][int(dk)] = dv
    return all_mods


def build_grupo_periodos(cal, aliados, meta=None):
    """
    Construye GRUPO_PERIODOS completo desde el calendario y los datos de ALIADOS.
    Preserva periods históricos con sus totales originales si existen en el html.
    """
    plat_key_map = {'F4F':'F4F','Chaturbate':'CB','Stripchat':'SC','CamSoda':'CAM','Streamate':'STR'}
    all_mods_by_month = _build_all_mods_by_month(aliados)

    gp = {}
    for plat_display, perds in cal.items():
        plat_key = plat_key_map.get(plat_display, plat_display)
        out = []
        for p in perds:
            new_models = _sum_period_from_aliados(all_mods_by_month, p['start'], p['end'], plat_key)
            p['models'] = new_models
            p['total']  = sum(new_models.values())
            out.append(p)
        gp[plat_display] = out

    if meta:
        gp['_meta'] = meta
    else:
        gp['_meta'] = {'studio': 'Grupo Empresarial J&D', 'entity_label': 'Estudio',
                       'streamate_active': True}
    return gp


def build_studio_periodos(cal, aliados_entry):
    """
    Construye PERIODOS para un dashboard de estudio.
    """
    plat_key_map = {'F4F':'F4F','Chaturbate':'CB','Stripchat':'SC','CamSoda':'CAM','Streamate':'STR'}
    all_mods_by_month = {}
    for mes, md in (aliados_entry.get('data') or {}).items():
        all_mods_by_month[mes] = {}
        for model, days in (md.get('modelos') or {}).items():
            all_mods_by_month[mes][model] = {int(dk): dv for dk, dv in days.items() if dv}

    periodos = {}
    for plat_display, perds in cal.items():
        plat_key = plat_key_map.get(plat_display, plat_display)
        out = []
        for p in perds:
            new_models = _sum_period_from_aliados(all_mods_by_month, p['start'], p['end'], plat_key)
            # filter zero-models
            new_models = {k: v for k, v in new_models.items() if v > 0}
            out.append({
                'label':  p['label'],
                'start':  p['start'],
                'end':    p['end'],
                'models': new_models,
                'total':  sum(new_models.values()),
            })
        periodos[plat_display] = out
    return periodos


# ═══════════════════════════════════════════════════════════════════
# REBUILD ALIADOS
# ═══════════════════════════════════════════════════════════════════
def rebuild_aliados_from_excel(aliados, cv, key_map, mes, last_day):
    """Reconstruye datos diarios del ALIADOS para el mes. Crea el mes si no existe."""
    cleared = 0; updated = 0
    for ak, ainfo in aliados.items():
        cv_studios = key_map.get(ak)
        if cv_studios is None:
            continue
        data = ainfo.setdefault('data', {})
        if mes not in data:
            ref_mods = list((data.get(MES, {}).get('modelos') or {}).keys())
            data[mes] = {'dias': 0, 'modelos': {m: {} for m in ref_mods}}
            print(f"  ➕ Creado {mes} en ALIADOS[{ak}] ({len(ref_mods)} modelos)")

        md      = data[mes]
        modelos = md.get('modelos')
        if modelos is None:
            continue
        for model in list(modelos.keys()):
            modelos[model] = {}; cleared += 1
        for model in list(modelos.keys()):
            for studio in cv_studios:
                key = (model, studio)
                if key not in cv:
                    continue
                for d, day_vals in cv[key].items():
                    entry = {p: (day_vals.get(p) if day_vals.get(p, 0) else None) for p in PLATS}
                    modelos[model][str(d)] = entry; updated += 1
                break
        md['dias'] = last_day
    print(f"  🗑  [{mes}] Limpiados: {cleared} | Reconstruidos: {updated}")
    return aliados


# ═══════════════════════════════════════════════════════════════════
# REBUILD ESTUDIO
# ═══════════════════════════════════════════════════════════════════
def rebuild_estudio_from_excel(estudio, cv, cv_studio, last_day, mes=MES):
    """Actualiza ESTUDIO.data[mes] desde Excel. Crea el mes si no existe."""
    data = estudio.setdefault('data', {})
    if mes not in data:
        ref_mods = list((data.get(MES, {}).get('modelos') or {}).keys())
        data[mes] = {'dias': 0, 'modelos': {m: {} for m in ref_mods}}
        print(f"  ➕ Creado {mes} en ESTUDIO [{cv_studio}]")

    md      = data[mes]
    modelos = md.get('modelos') or {}
    cl = 0; up = 0
    for model in list(modelos.keys()):
        modelos[model] = {}; cl += 1
    for model in list(modelos.keys()):
        key = (model, cv_studio)
        if key not in cv:
            continue
        for d, day_vals in cv[key].items():
            entry = {p: (day_vals.get(p) if day_vals.get(p, 0) else None) for p in PLATS}
            modelos[model][str(d)] = entry; up += 1
    md['dias'] = last_day
    print(f"  ✓ ESTUDIO [{cv_studio}] {mes}: {cl} limpiados, {up} entradas | dias={last_day}")
    return estudio


# ═══════════════════════════════════════════════════════════════════
# TOP_EST y TOP20G
# ═══════════════════════════════════════════════════════════════════
def recalc_top_est(top_est, aliados_entry, mes=MES):
    md       = (aliados_entry.get('data') or {}).get(mes, {})
    all_mods = md.get('modelos') or {}
    totals = {}
    for model, days in all_mods.items():
        t = sum(sum((dv.get(p) or 0) for p in PLATS) for dv in days.values() if dv)
        if t > 0:
            totals[model] = t
    ranked   = sorted(totals.items(), key=lambda x: -x[1])
    new_list = [{'modelo': m} for m, _ in ranked]
    if top_est is None:
        return {mes: new_list}
    top_est[mes] = new_list
    return top_est

ALIADOS_DISPLAY = {
    'Fornax Studios':    'Fornax Studios',
    'agatha_studios_':   'Agatha Studio',
    'goldonline':        'Gold Online',
    'kama_studio':       'Kama Studio',
    'the_Room_studios':  'The Room Studio',
    'CyV Studios':       'CyV Studios',
    'Studio Levi':       'Studios Levi',
    'The Online Agency': 'The Online Agency',
    'Elite Cam House':   'Elite Cam House',
    'Studio RWB':        'Studios RWB',
    'PrestigeCam':       'Prestige Cam',
    'Atelier_glamour':   'Atelier Glamour',
    'Dejavu Studio':     'Dejavu Studio',
    'Dynasty_studio_':   'Dynasty Studio',
    'piscis_studio':     'Piscis Studio',
}

def recalc_top20g(top20g, grupo_aliados, mes=MES):
    all_totals = {}
    for ak, ainfo in grupo_aliados.items():
        studio_name = ALIADOS_DISPLAY.get(ak, ak)
        md = (ainfo.get('data') or {}).get(mes, {})
        for model, days in (md.get('modelos') or {}).items():
            t = sum(sum((dv.get(p) or 0) for p in PLATS) for dv in days.values() if dv)
            if t > 0:
                key = (model, studio_name)
                all_totals[key] = all_totals.get(key, 0) + t
    ranked   = sorted(all_totals.items(), key=lambda x: -x[1])[:20]
    new_list = [{'modelo': m, 'studio': s} for (m, s), _ in ranked]
    print(f"    TOP20G [{mes}]: {len(new_list)} | #1={new_list[0]['modelo'] if new_list else '—'}")
    if top20g is None:
        return {mes: new_list}
    top20g[mes] = new_list
    return top20g


# ═══════════════════════════════════════════════════════════════════
# MAPEOS ALIADOS KEY → CV STUDIOS
# ═══════════════════════════════════════════════════════════════════
GRUPO_MAP = {
    'Fornax Studios':    ['Fornax Studios'],
    'agatha_studios_':   ['Agatha Studio'],
    'goldonline':        ['Gold Online'],
    'kama_studio':       ['Kama Studio', 'GrupoJ&D'],
    'the_Room_studios':  ['The Room Studio', 'The Room Studios'],
    'CyV Studios':       ['CyV Studios'],
    'Studio Levi':       ['Studios Levi'],
    'The Online Agency': ['The Online Agency'],
    'Elite Cam House':   ['Elite Cam House'],
    'Studio RWB':        ['Studios RWB'],
    'PrestigeCam':       ['Prestige Cam'],
    'Atelier_glamour':   ['Atelier Glamour'],
    'Dejavu Studio':     ['Dejavu Studio'],
    'Dynasty_studio_':   [],
    'piscis_studio':     ['Piscis Studio'],
}
ERIKA_MAP = {
    'CyV Studios':       ['CyV Studios'],
    'Studio Levi':       ['Studios Levi'],
    'The Online Agency': ['The Online Agency'],
    'Elite Cam House':   ['Elite Cam House'],
    'Studio RWB':        ['Studios RWB'],
    'PrestigeCam':       ['Prestige Cam'],
    'Dulce Luna':        ['Fornax Studios'],
    'Liam Terrier':      ['Fornax Studios'],
    'Zac Levis':         ['Gold Online'],
    'Joel Souza':        ['Erika Noguera'],
    'William Gardener':  ['Erika Noguera'],
}
FABIO_MAP = {
    'Atelier_glamour':   ['Atelier Glamour'],
    'Dejavu Studio':     ['Dejavu Studio'],
    'Dynasty_studio_':   [],
    'piscis_studio':     ['Piscis Studio'],
    'Alice Steel':       ['Fabio Robledo'],
    'Eli Cortes':        ['Fabio Robledo'],
    'Evelyn Lovers':     ['Fabio Robledo'],
    'Jack Miller':       ['Fabio Robledo'],
    'Maximus Clark':     ['Fabio Robledo'],
    'Amanda Bond':       ['Fabio Robledo'],
    'Yessie Jacobs':     ['Fabio Robledo'],
    'Ana Black':         ['Fabio Robledo'],
    'Amadeus Studio':    ['Amadeus Studio'],
    'Black Card':        ['Black Card'],
    'Iridium Studio':    [],
    'Studio JGM':        ['Studio JGM'],
}


# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    now    = datetime.now()
    ts_str = now.strftime('%d/%m/%Y — %H:%M')
    meses_activos = ALL_MESES  # Enero → Agosto

    print('=' * 65)
    print(f'  ACTUALIZADOR J&D — {ts_str}')
    print(f'  MESES: {" | ".join(meses_activos)}')
    print(f'  Fórmula: day_col(d) = 19+(31-d)*5  [5 plats/día]')
    print('=' * 65)

    # ── Cargar Excel ──────────────────────────────────────────────
    print(f'\n📂 Cargando: {os.path.basename(CV_PATH)}')
    wb = openpyxl.load_workbook(CV_PATH, data_only=True)

    cv_jul, last_day_jul = read_cv_sheet(wb, 'JULIO')
    cv_ago, last_day_ago = read_cv_sheet(wb, 'AGOSTO')

    print(f'\n📅 JULIO: último día={last_day_jul} | AGOSTO: último día={last_day_ago}')

    if last_day_jul == 0:
        print("⚠  Sin datos en JULIO — verificar Excel."); return

    # ── Leer calendario de periodos ───────────────────────────────
    print('\n📆 Leyendo CALENDARIO MAESTRO...')
    cal = read_calendario_maestro(wb)

    # ── Helpers ───────────────────────────────────────────────────
    def load_dash(path, al_varname):
        with open(path, 'r', encoding='utf-8') as f:
            html = f.read()
        aliados, al_q = get_var(html, al_varname)
        if aliados is None:
            print(f"  ⚠  No se encontró {al_varname} en {os.path.basename(os.path.dirname(path))}")
        return html, aliados, al_q

    def save_dash(path, html):
        with open(path, 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"  💾 {os.path.relpath(path, SCRIPT_DIR)}")

    # ════════════════════════════════════════════════════════════
    # GRUPO EMPRESARIAL
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  GRUPO EMPRESARIAL')
    html, aliados, al_q = load_dash(DASHBOARDS['grupo'], 'ALIADOS')
    if aliados is not None:
        # Rebuild datos diarios
        aliados = rebuild_aliados_from_excel(aliados, cv_jul, GRUPO_MAP, MES,     last_day_jul)
        aliados = rebuild_aliados_from_excel(aliados, cv_ago, GRUPO_MAP, MES_AGO, last_day_ago)

        # Rebuild GRUPO_PERIODOS desde el calendario completo
        gp, gp_q = get_var(html, 'GRUPO_PERIODOS')
        meta = (gp or {}).get('_meta')
        gp_new = build_grupo_periodos(cal, aliados, meta)
        for plat, perds in gp_new.items():
            if plat == '_meta':
                continue
            last = (perds or [{}])[-1]
            print(f"  ✓ GP {plat}: {len(perds)} periodos | último={last.get('label')} → {last.get('total',0):.0f}")

        # TOP20G
        top20g, t20_q = get_var(html, 'TOP20')
        if top20g is not None:
            for m in meses_activos:
                top20g = recalc_top20g(top20g, aliados, m)

        # Patch HTML
        html = set_var(html, 'ALIADOS',        aliados, al_q)
        html = set_var(html, 'GRUPO_PERIODOS', gp_new, gp_q)
        if top20g is not None:
            html = set_var(html, 'TOP20', top20g, t20_q)
        html = patch_meses_dmes(html, meses_activos)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['grupo'], html)

    # ════════════════════════════════════════════════════════════
    # ERIKA NOGUERA
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  ERIKA NOGUERA')
    html, aliados_e, al_q = load_dash(DASHBOARDS['erika'], 'ALIADOS')
    if aliados_e is not None:
        aliados_e = rebuild_aliados_from_excel(aliados_e, cv_jul, ERIKA_MAP, MES,     last_day_jul)
        aliados_e = rebuild_aliados_from_excel(aliados_e, cv_ago, ERIKA_MAP, MES_AGO, last_day_ago)

        ep, ep_q = get_var(html, 'EXEC_PERIODOS')
        gp, gp_q = get_var(html, 'GRUPO_PERIODOS')
        if ep is not None:
            ep_meta = ep.get('_meta')
            ep_new  = build_grupo_periodos(cal, aliados_e, ep_meta)
            html = set_var(html, 'EXEC_PERIODOS', ep_new, ep_q)
        if gp is not None:
            # Load grupo aliados for full GRUPO_PERIODOS
            with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f:
                g_html = f.read()
            g_aliados, _ = get_var(g_html, 'ALIADOS')
            gp_meta = gp.get('_meta')
            gp_new  = build_grupo_periodos(cal, g_aliados or aliados_e, gp_meta)
            html = set_var(html, 'GRUPO_PERIODOS', gp_new, gp_q)

        html = set_var(html, 'ALIADOS', aliados_e, al_q)
        html = patch_meses_dmes(html, meses_activos)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['erika'], html)

    # ════════════════════════════════════════════════════════════
    # FABIO ROBLEDO
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  FABIO ROBLEDO')
    html, aliados_f, al_q = load_dash(DASHBOARDS['fabio'], 'ALIADOS')
    if aliados_f is not None:
        aliados_f = rebuild_aliados_from_excel(aliados_f, cv_jul, FABIO_MAP, MES,     last_day_jul)
        aliados_f = rebuild_aliados_from_excel(aliados_f, cv_ago, FABIO_MAP, MES_AGO, last_day_ago)

        ep, ep_q = get_var(html, 'EXEC_PERIODOS')
        if ep is not None:
            ep_meta = ep.get('_meta')
            ep_new  = build_grupo_periodos(cal, aliados_f, ep_meta)
            html = set_var(html, 'EXEC_PERIODOS', ep_new, ep_q)

        html = set_var(html, 'ALIADOS', aliados_f, al_q)
        html = patch_meses_dmes(html, meses_activos)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['fabio'], html)

    # ── Leer grupo actualizado (para TOP20G en estudios) ──────────
    with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f:
        g_html = f.read()
    g_aliados, _ = get_var(g_html, 'ALIADOS')

    def update_studio(dash_key, cv_studio_name, aliados_map_key):
        """Actualiza un dashboard de estudio."""
        print(f'\n▶  {aliados_map_key.upper()} (estudio)')
        with open(DASHBOARDS[dash_key], 'r', encoding='utf-8') as f:
            html = f.read()

        # ESTUDIO
        estudio, es_q = get_var(html, 'ESTUDIO')
        if estudio is not None:
            estudio = rebuild_estudio_from_excel(estudio, cv_jul, cv_studio_name, last_day_jul, MES)
            estudio = rebuild_estudio_from_excel(estudio, cv_ago, cv_studio_name, last_day_ago, MES_AGO)
            html = set_var(html, 'ESTUDIO', estudio, es_q)

        # PERIODOS (desde calendario)
        aliados_entry = (g_aliados or {}).get(aliados_map_key, {})
        per_new = build_studio_periodos(cal, aliados_entry)
        for plat, perds in per_new.items():
            last = (perds or [{}])[-1]
            if last.get('total', 0):
                print(f"    {plat}: {len(perds)} periodos | último → {last['total']:.0f}")
        pq = get_var(html, 'PERIODOS')[1]
        html = set_var(html, 'PERIODOS', per_new, pq)

        # TOP_EST
        top_est, te_q = get_var(html, 'TOP_EST')
        if top_est is not None:
            for m in meses_activos:
                top_est = recalc_top_est(top_est, aliados_entry, m)
            html = set_var(html, 'TOP_EST', top_est, te_q)
            print(f"  ✓ TOP_EST: " +
                  " | ".join(f"{m}={len((top_est or {}).get(m,[]))}" for m in meses_activos))

        # TOP20G
        if g_aliados is not None:
            top20g, t20_q = get_var(html, 'TOP20G')
            if top20g is not None:
                for m in meses_activos:
                    top20g = recalc_top20g(top20g, g_aliados, m)
                html = set_var(html, 'TOP20G', top20g, t20_q)

        html = patch_meses_dmes(html, meses_activos)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS[dash_key], 'w', encoding='utf-8') as f:
            f.write(html)
        print(f"  💾 {dash_key}")

    update_studio('fornax', 'Fornax Studios', 'Fornax Studios')
    update_studio('gold',   'Gold Online',    'goldonline')
    update_studio('cyv',    'CyV Studios',    'CyV Studios')

    # ── Resumen ───────────────────────────────────────────────────
    print(f'\n{"=" * 65}')
    print(f'  ✅ Rebuild completado — {ts_str}')
    print(f'  📅 JULIO días 1–{last_day_jul} | AGOSTO días 1–{last_day_ago}')
    print(f'  📆 Periodos CPP desde CALENDARIO MAESTRO')
    print(f'  🗓  MESES activos: {" | ".join(meses_activos)}')
    print(f'  ⚠  NO commit/push — esperar validación')
    print(f'{"=" * 65}\n')


if __name__ == '__main__':
    main()
