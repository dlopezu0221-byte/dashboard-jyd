#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACTUALIZAR.py — Actualizador dashboards Grupo J&D
==================================================
Lee el Excel "Cómo Vamos" (hoja JULIO), actualiza los dashboards con:
  - Nuevos datos diarios desde el Excel (días disponibles)
  - ALIADOS.dias = último día con datos en ALIADOS (preserva datos históricos)
  - GRUPO_PERIODOS / PERIODOS / EXEC_PERIODOS recalculados desde ALIADOS
  - Timestamp de generación en todos los dashboards

Uso: python ACTUALIZAR.py

NOTAS:
  - NO inventa datos: solo usa lo que está en el Excel o en ALIADOS
  - NO hace commit/push (debe ser autorizado manualmente)
  - Fórmula de columnas diarias: col_F4F(d) = 16 + (31-d)*4
    Plataformas por día: F4F(+0), SC(+1), CB(+2), CAM(+3) — sin STR en diario
"""

import openpyxl, re, base64, json, os
from datetime import datetime
from collections import defaultdict

# ═══════════════════════════════════════════════════════════════════
# RUTAS
# ═══════════════════════════════════════════════════════════════════
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CV_PATH = os.path.join(SCRIPT_DIR, '..', 'Centro de Gestión Estratégica Grupo J&D',
                       'COMO VAMOS GRUPO', 'Cómo vamos Grupo Empresarial.xlsx')
DASHBOARDS = {
    'grupo':  os.path.join(SCRIPT_DIR, 'grupo579780',                    'index.html'),
    'erika':  os.path.join(SCRIPT_DIR, 'erika868527',                    'index.html'),
    'fabio':  os.path.join(SCRIPT_DIR, 'fabio473013',                    'index.html'),
    'fornax': os.path.join(SCRIPT_DIR, 'estudios', 'fornax-studios345929','index.html'),
    'gold':   os.path.join(SCRIPT_DIR, 'estudios', 'goldonline078939',    'index.html'),
    'cyv':    os.path.join(SCRIPT_DIR, 'estudios', 'cyv-studios837357',   'index.html'),
}

MES    = 'Julio'
PLATS  = ['F4F', 'SC', 'CB', 'CAM', 'STR']
PLATS4 = ['F4F', 'SC', 'CB', 'CAM']          # sólo estos en columnas diarias

def day_col(d):
    """Columna F4F para el día d en la hoja JULIO (4 plats/día, inicio col 16)."""
    return 16 + (31 - d) * 4

# ═══════════════════════════════════════════════════════════════════
# HELPERS BASE64
# ═══════════════════════════════════════════════════════════════════
def b64enc(obj):
    return base64.b64encode(json.dumps(obj, ensure_ascii=False).encode('utf-8')).decode('ascii')

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
        print(f"  ⚠  No se encontró var {varname} con quote={quote}")
    return result

# ═══════════════════════════════════════════════════════════════════
# PASO 1 — LEER EXCEL (sólo para datos NUEVOS no inventar)
# ═══════════════════════════════════════════════════════════════════
def read_cv_excel(path, sheet='JULIO'):
    print(f"\n📂 Leyendo Excel: {os.path.basename(path)}")
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        print(f"  ⚠  Hoja '{sheet}' no encontrada. Hojas disponibles: {wb.sheetnames}")
        return {}, 0
    ws = wb[sheet]

    cv = {}   # (model, studio) → {day → {plat → float}}
    last_day_excel = 0

    for r in range(5, ws.max_row + 1):
        model  = str(ws.cell(r, 1).value or '').strip()
        studio = str(ws.cell(r, 2).value or '').strip()
        if not model or not studio:
            continue
        entry = {}
        for d in range(1, 32):
            c0 = day_col(d)
            if c0 > ws.max_column:
                break
            day_data = {}
            has = False
            for pi, p in enumerate(PLATS4):
                v = ws.cell(r, c0 + pi).value
                if v and isinstance(v, (int, float)) and float(v) > 0:
                    day_data[p] = float(v)
                    has = True
                    last_day_excel = max(last_day_excel, d)
                else:
                    day_data[p] = None
            if has:
                entry[d] = day_data
        if entry:
            key = (model, studio)
            cv[key] = entry

    print(f"  📅 Último día con datos en Excel: {last_day_excel}")
    print(f"  👥 Modelos encontrados: {len(cv)}")
    return cv, last_day_excel

# ═══════════════════════════════════════════════════════════════════
# PASO 2 — DETERMINAR ÚLTIMO DÍA REAL DEL ALIADOS
# ═══════════════════════════════════════════════════════════════════
def aliados_last_day(aliados, mes=MES):
    """Devuelve el día más alto con algún dato en cualquier modelo del ALIADOS."""
    last = 0
    for ak, ainfo in aliados.items():
        md = (ainfo.get('data') or {}).get(mes, {})
        for model, days in (md.get('modelos') or {}).items():
            for dk, dv in days.items():
                if dv and any(dv.values()):
                    try:
                        last = max(last, int(dk))
                    except (ValueError, TypeError):
                        pass
    return last

# ═══════════════════════════════════════════════════════════════════
# PASO 3 — ACTUALIZAR ALIADOS con nuevos datos Excel
#  (sólo SOBREESCRIBE si Excel tiene dato; no borra días existentes)
# ═══════════════════════════════════════════════════════════════════
def update_aliados_from_excel(aliados, cv, key_map, last_day, mes=MES):
    """
    key_map: {aliados_key → [cv_studio_name, ...]}
    Actualiza días 1..last_day desde Excel si el modelo aparece ahí.
    Preserva días con datos que no están en Excel (datos históricos).
    Siempre fija dias = last_day.
    """
    updated = 0
    for ak, ainfo in aliados.items():
        cv_studios = key_map.get(ak)
        if cv_studios is None:
            continue  # no mapeado
        md = (ainfo.get('data') or {}).get(mes)
        if not md:
            continue
        modelos = md.get('modelos') or {}
        for model in modelos:
            for studio in cv_studios:
                key = (model, studio)
                if key not in cv:
                    continue
                for d, day_vals in cv[key].items():
                    entry = {p: (day_vals.get(p) if day_vals.get(p, 0) else None) for p in PLATS}
                    modelos[model][str(d)] = entry
                    updated += 1
                break  # encontrado en primer studio
        md['dias'] = last_day
    print(f"  ✓ Excel→ALIADOS: {updated} entradas actualizadas")
    return aliados

# ═══════════════════════════════════════════════════════════════════
# PASO 4 — RECALCULAR PERIODOS DESDE ALIADOS
# ═══════════════════════════════════════════════════════════════════
def parse_day_from_date(date_str):
    """'27/07/2026' → 27 (para julio)."""
    try:
        parts = str(date_str).split('/')
        day, month = int(parts[0]), int(parts[1])
        return day if month == 7 else (31 if month > 7 else 0)
    except Exception:
        return 0

def recalc_periods(periods, aliados_mods, mes=MES):
    """
    periods: [{label, start, end, models, total, unit, has_activity}, ...]
    aliados_mods: dict {model → {day → {plat → val}}}  (todos los modelos del scope)
    Recalcula el ÚLTIMO período (el más reciente).
    """
    if not periods:
        return periods
    last = periods[-1]
    d_from = parse_day_from_date(last.get('start', ''))
    d_to   = parse_day_from_date(last.get('end',   ''))
    if d_from == 0:
        return periods

    for plat in PLATS:
        new_models = {}
        for model, days in aliados_mods.items():
            t = 0.0
            for d in range(d_from, d_to + 1):
                dv = days.get(str(d)) or days.get(d)
                t += (dv.get(plat) or 0) if dv else 0
            if t > 0:
                new_models[model] = t
        if new_models:
            # Update only the periods for this specific platform
            # We'll do this per-platform in the caller
            pass

    # Per-platform in GRUPO_PERIODOS / PERIODOS the structure is nested
    return periods  # handled in caller

def recalc_gp_last(gp, aliados, aliados_key=None, mes=MES):
    """
    Recalcula el ÚLTIMO periodo de cada plataforma en GRUPO_PERIODOS.
    Suma sobre TODOS los aliados (o el aliado específico si se da).
    """
    # Collect all modelos from all aliados (for Grupo scope)
    all_mods = {}  # model → {day → {plat → val}}
    scope = aliados if aliados_key is None else {aliados_key: aliados[aliados_key]}
    for ak, ainfo in scope.items():
        md = (ainfo.get('data') or {}).get(mes, {})
        for model, days in (md.get('modelos') or {}).items():
            if model not in all_mods:
                all_mods[model] = {}
            for dk, dv in days.items():
                if dv:
                    all_mods[model][dk] = dv

    for plat_name, plat_key in [('F4F','F4F'),('Stripchat','SC'),
                                  ('Chaturbate','CB'),('CamSoda','CAM'),('Streamate','STR')]:
        periods = gp.get(plat_name, [])
        if not periods:
            continue
        last = periods[-1]
        d_from = parse_day_from_date(last.get('start', ''))
        d_to   = parse_day_from_date(last.get('end',   ''))
        if d_from == 0:
            continue
        new_models = {}
        for model, days in all_mods.items():
            t = 0.0
            for d in range(d_from, d_to + 1):
                dv = days.get(str(d)) or days.get(d)
                t += (dv.get(plat_key) or 0) if dv else 0
            if t > 0:
                new_models[model] = t
        if new_models or last.get('models'):
            last['models'] = new_models if new_models else last.get('models', {})
            last['total']  = sum(new_models.values()) if new_models else last.get('total', 0)
    return gp

def recalc_studio_periodos(periodos, aliados_entry, mes=MES):
    """Recalcula PERIODOS para un dashboard de estudio."""
    md = (aliados_entry.get('data') or {}).get(mes, {})
    all_mods = md.get('modelos') or {}

    for plat_name, plat_key in [('F4F','F4F'),('Stripchat','SC'),
                                  ('Chaturbate','CB'),('CamSoda','CAM'),('Streamate','STR')]:
        periods = periodos.get(plat_name, [])
        if not periods:
            continue
        last = periods[-1]
        d_from = parse_day_from_date(last.get('start', ''))
        d_to   = parse_day_from_date(last.get('end',   ''))
        if d_from == 0:
            continue
        new_models = {}
        for model, days in all_mods.items():
            t = 0.0
            for d in range(d_from, d_to + 1):
                dv = days.get(str(d)) or days.get(d)
                t += (dv.get(plat_key) or 0) if dv else 0
            if t > 0:
                new_models[model] = t
        if new_models:
            last['models'] = new_models
            last['total']  = sum(new_models.values())
            print(f"    {plat_name}: {last.get('label')} → {last['total']:.0f}")
    return periodos

# ═══════════════════════════════════════════════════════════════════
# PASO 5 — TIMESTAMP
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
# MAPEOS ALIADOS KEY → CV STUDIOS
# ═══════════════════════════════════════════════════════════════════
GRUPO_MAP = {
    'Fornax Studios':   ['Fornax Studios'],
    'agatha_studios_':  ['Agatha Studio'],
    'goldonline':       ['Gold Online'],
    'kama_studio':      ['Kama Studio', 'GrupoJ&D'],
    'the_Room_studios': ['The Room Studio'],
    'CyV Studios':      ['CyV Studios'],
    'Studio Levi':      ['Studios Levi'],
    'The Online Agency':['The Online Agency'],
    'Elite Cam House':  ['Elite Cam House'],
    'Studio RWB':       ['Studios RWB'],
    'PrestigeCam':      ['Prestige Cam'],
    'Atelier_glamour':  ['Atelier Glamour'],
    'Dejavu Studio':    ['Dejavu Studio'],
    'Dynasty_studio_':  [],
    'piscis_studio':    ['Piscis Studio'],
}

ERIKA_MAP = {
    'CyV Studios':      ['CyV Studios'],
    'Studio Levi':      ['Studios Levi'],
    'The Online Agency':['The Online Agency'],
    'Elite Cam House':  ['Elite Cam House'],
    'Studio RWB':       ['Studios RWB'],
    'PrestigeCam':      ['Prestige Cam'],
    'Dulce Luna':       ['Fornax Studios'],
    'Liam Terrier':     ['Fornax Studios'],
    'Zac Levis':        ['Gold Online'],
    'Joel Souza':       ['Erika Noguera'],
    'William Gardener': ['Erika Noguera'],
}

FABIO_MAP = {
    'Atelier_glamour':  ['Atelier Glamour'],
    'Dejavu Studio':    ['Dejavu Studio'],
    'Dynasty_studio_':  [],
    'piscis_studio':    ['Piscis Studio'],
    'Alice Steel':      ['Fabio Robleo'],
    'Eli Cortes':       ['Fabio Robleo'],
    'Evelyn Lovers':    ['Fabio Robleo'],
    'Jack Miller':      ['Fabio Robleo'],
    'Maximus Clark':    ['Fabio Robleo'],
    'Amanda Bond':      ['Fabio Robleo'],
    'Yessie Jacobs':    ['Fabio Robleo'],
    'Ana Black':        ['Fabio Robleo'],
    'Amadeus Studio':   ['Amadeus Studio'],
    'Black Card':       ['Black Card'],
    'Iridium Studio':   [],
    'Studio JGM':       ['Studio JGM'],
}

# ═══════════════════════════════════════════════════════════════════
# MAIN
# ═══════════════════════════════════════════════════════════════════
def main():
    now = datetime.now()
    ts_str = now.strftime('%d/%m/%Y — %H:%M')
    print('=' * 60)
    print(f'  ACTUALIZADOR J&D — {ts_str}')
    print('=' * 60)

    # Leer Excel
    cv, last_day_excel = read_cv_excel(CV_PATH)

    def process(name, dash_path, al_varname, gp_varname, extra_vars=None):
        print(f'\n▶  {name}')
        with open(dash_path, 'r', encoding='utf-8') as f:
            html = f.read()
        aliados, al_q = get_var(html, al_varname)
        if aliados is None:
            print(f"  ⚠  No se encontró {al_varname}"); return

        # Determinar last_day: máximo de Excel y de lo que ya hay en ALIADOS
        al_last = aliados_last_day(aliados)
        last_day = max(last_day_excel, al_last)
        print(f"  📅 last_day Excel={last_day_excel} ALIADOS={al_last} → usando {last_day}")

        return html, aliados, al_q, last_day

    # ── GRUPO ──
    result = process('GRUPO EMPRESARIAL', DASHBOARDS['grupo'], 'ALIADOS', 'GRUPO_PERIODOS')
    if result:
        html, aliados, al_q, last_day = result
        key_map = GRUPO_MAP
        aliados = update_aliados_from_excel(aliados, cv, key_map, last_day)

        gp, gp_q = get_var(html, 'GRUPO_PERIODOS')
        gp = recalc_gp_last(gp, aliados)
        for plat in ['Stripchat','Chaturbate','CamSoda','Streamate','F4F']:
            last = (gp.get(plat) or [{}])[-1]
            if last.get('total'):
                print(f"  ✓ GP {plat}: {last.get('label')} → {last.get('total'):.0f}")

        html = set_var(html, 'ALIADOS',        aliados, al_q)
        html = set_var(html, 'GRUPO_PERIODOS', gp,      gp_q)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['grupo'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 grupo579780/index.html")

    # ── ERIKA ──
    result = process('ERIKA NOGUERA', DASHBOARDS['erika'], 'ALIADOS', 'EXEC_PERIODOS')
    if result:
        html, aliados, al_q, last_day = result
        aliados = update_aliados_from_excel(aliados, cv, ERIKA_MAP, last_day)

        ep, ep_q   = get_var(html, 'EXEC_PERIODOS')
        gp, gp_q   = get_var(html, 'GRUPO_PERIODOS')
        ep = recalc_gp_last(ep, aliados)
        if gp: gp = recalc_gp_last(gp, aliados)

        html = set_var(html, 'ALIADOS',       aliados, al_q)
        html = set_var(html, 'EXEC_PERIODOS', ep,      ep_q)
        if gp: html = set_var(html, 'GRUPO_PERIODOS', gp, gp_q)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['erika'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 erika868527/index.html")

    # ── FABIO ──
    result = process('FABIO ROBLEDO', DASHBOARDS['fabio'], 'ALIADOS', 'EXEC_PERIODOS')
    if result:
        html, aliados, al_q, last_day = result
        aliados = update_aliados_from_excel(aliados, cv, FABIO_MAP, last_day)

        ep, ep_q = get_var(html, 'EXEC_PERIODOS')
        ep = recalc_gp_last(ep, aliados)
        html = set_var(html, 'ALIADOS',       aliados, al_q)
        html = set_var(html, 'EXEC_PERIODOS', ep,      ep_q)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['fabio'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 fabio473013/index.html")

    # ── FORNAX STUDIO ──
    print('\n▶  FORNAX STUDIOS (dashboard estudio)')
    with open(DASHBOARDS['fornax'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        # Para studio dashboard, usar datos del ALIADOS Grupo para Fornax Studios
        with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
        g_aliados, _ = get_var(g2, 'ALIADOS')
        fx_entry = g_aliados.get('Fornax Studios', {})
        print('  Recalculando PERIODOS desde ALIADOS Fornax:')
        periodos = recalc_studio_periodos(periodos, fx_entry)
        html = set_var(html, 'PERIODOS', periodos, pq)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['fornax'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 fornax-studios345929/index.html")

    # ── GOLD ONLINE STUDIO ──
    print('\n▶  GOLD ONLINE (dashboard estudio)')
    with open(DASHBOARDS['gold'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
        g_aliados, _ = get_var(g2, 'ALIADOS')
        go_entry = g_aliados.get('goldonline', {})
        print('  Recalculando PERIODOS desde ALIADOS Gold Online:')
        periodos = recalc_studio_periodos(periodos, go_entry)
        html = set_var(html, 'PERIODOS', periodos, pq)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['gold'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 goldonline078939/index.html")

    # ── CYV STUDIO ──
    print('\n▶  CYV STUDIOS (dashboard estudio)')
    with open(DASHBOARDS['cyv'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        with open(DASHBOARDS['erika'], 'r', encoding='utf-8') as f: e2 = f.read()
        e_aliados, _ = get_var(e2, 'ALIADOS')
        if e_aliados is None:
            with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
            e_aliados, _ = get_var(g2, 'ALIADOS')
        cy_entry = e_aliados.get('CyV Studios', {})
        print('  Recalculando PERIODOS desde ALIADOS CyV:')
        periodos = recalc_studio_periodos(periodos, cy_entry)
        top_est_cy, te_q = get_var(html, 'TOP_EST')
        if top_est_cy is not None:
            top_est_cy = recalc_top_est(top_est_cy, cy_entry)
            html = set_var(html, 'TOP_EST', top_est_cy, te_q)
            print(f"  ✓ TOP_EST CyV: {len(top_est_cy.get(MES,[]))} modelos")
        html = set_var(html, 'PERIODOS', periodos, pq)
        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['cyv'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 cyv-studios837357/index.html")

    print(f'\n{"=" * 60}')
    print(f'  ✅ Actualización completada — timestamp: {ts_str}')
    print(f'  📌 días con datos en Excel: {last_day_excel}')
    print(f'  ⚠  SC/CB días 29-30: requieren actualización en Excel')
    print(f'{"=" * 60}\n')

if __name__ == '__main__':
    main()
