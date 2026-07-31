#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACTUALIZAR.py — Actualizador dashboards Grupo J&D
==================================================
RECONSTRUYE desde cero los datos diarios del ALIADOS cada vez que se ejecuta.

ESTRUCTURA VERIFICADA del Excel LOCAL (fuente de verdad):
  - Col 1 : nombre modelo
  - Col 2 : studio
  - Cols 3-7   : P1 (F4F, SC, CB, CAM, STR)
  - Cols 8-12  : P2
  - Cols 13-17 : TOTAL MES
  - Col 18 : nombre modelo (repetido)
  - Cols 19+   : datos diarios (5 plataformas por día)
    → día 31 : cols 19-23  (F4F, SC, CB, CAM, STR)
    → día 30 : cols 24-28
    → día 1  : cols 169-173
  Fórmula: col_F4F(d) = 19 + (31-d)*5

Uso: python ACTUALIZAR.py
"""

import openpyxl, re, base64, json, os
from datetime import datetime

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
# Todas las plataformas (tal como están en el ALIADOS de los HTML)
PLATS  = ['F4F', 'SC', 'CB', 'CAM', 'STR']

# ═══════════════════════════════════════════════════════════════════
# FÓRMULA VERIFICADA contra el Excel LOCAL (fuente de verdad)
# ═══════════════════════════════════════════════════════════════════
def day_col(d):
    """
    Columna F4F para el día d en el Excel LOCAL (5 plats/día, inicio col 19).
    Verificado:
      day_col(31) = 19  → col 19 row4='Flirt4free'   ✓
      day_col(30) = 24  → col 24 row3=30              ✓
      day_col(1)  = 169 → Isa Raven col 169 = 4578    ✓
    """
    return 19 + (31 - d) * 5

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
# PASO 1 — LEER EXCEL: FUENTE DE VERDAD
# ═══════════════════════════════════════════════════════════════════
def read_cv_excel(path, sheet='JULIO'):
    """
    Lee el Excel LOCAL y extrae datos diarios.
    - 5 plataformas por día (F4F, SC, CB, CAM, STR)
    - Fórmula: day_col(d) = 19 + (31-d)*5
    - SOLO registra valores explícitamente > 0
    - NO inventa ni interpola datos
    - USA iter_rows(values_only=True) para lectura en RAM (evita acceso
      celda-a-celda sobre mount de red que causaba timeout de 45s).
    """
    print(f"\n📂 Leyendo Excel: {os.path.basename(path)}")
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet not in wb.sheetnames:
        print(f"  ⚠  Hoja '{sheet}' no encontrada. Hojas: {wb.sheetnames}")
        return {}, 0
    ws = wb[sheet]

    # Cargar TODO el sheet en RAM de una sola pasada secuencial
    # data[r][c] → valor (0-indexed); equivale a ws.cell(r+1, c+1).value
    data = tuple(ws.iter_rows(values_only=True))
    MAX_COL = ws.max_column
    print(f"  📋 Hoja en RAM: {len(data)} filas × {MAX_COL} columnas")

    # Verificar estructura (fila 3 = encabezado días, fila 4 = plataformas)
    h_day  = data[2][18]  # row 3, col 19 → índices [2][18]
    h_plat = data[3][18]  # row 4, col 19
    ok = (str(h_day) == '31' and 'lirt' in str(h_plat or ''))
    print(f"  📋 Col 19: día={h_day!r}, plat={h_plat!r}  {'✅' if ok else '❌ ADVERTENCIA: estructura inesperada'}")

    cv = {}            # (model, studio) → {day(int) → {plat → float|None}}
    last_day_excel = 0

    # Estudios a ignorar (resúmenes, totales, plataformas)
    IGNORAR_STUDIOS = {
        'CAMSODA', 'CAMSODA POR QUINCENA', 'CHATURBATE', 'FLIRT4FREE',
        'STREAMATE', 'STREAMATE POR QUINCENA', 'STRIPCHAT', 'STRIPCHAT POR QUINCENA',
        'TOTAL', 'TOTALES', 'CAMSODA por quincena', 'STREAMATE por quincena',
        'STRIPCHAT por quincena',
    }

    for ri in range(4, len(data)):        # fila 5+ → índice 4+
        row    = data[ri]
        model  = str(row[0] or '').strip()  # col 1 → índice 0
        studio = str(row[1] or '').strip()  # col 2 → índice 1
        if not model or not studio:
            continue
        if studio.upper() in {s.upper() for s in IGNORAR_STUDIOS}:
            continue
        if model.upper() in {'TOTAL', 'TOTALES', 'MODELO'}:
            continue

        entry = {}
        for d in range(1, 32):
            c0 = day_col(d)           # 1-indexed column
            if c0 > MAX_COL:
                break
            day_data = {}
            has = False
            for pi, p in enumerate(PLATS):
                ci = c0 + pi - 1      # 0-indexed
                if ci >= len(row):
                    day_data[p] = None
                    continue
                v = row[ci]
                if isinstance(v, (int, float)) and float(v) > 0:
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
    print(f"  👥 Modelos con datos diarios: {len(cv)}")
    return cv, last_day_excel

# ═══════════════════════════════════════════════════════════════════
# PASO 2 — REBUILD ALIADOS DESDE CERO
# ═══════════════════════════════════════════════════════════════════
def rebuild_aliados_from_excel(aliados, cv, key_map, mes=MES):
    """
    RECONSTRUYE completamente los datos diarios del ALIADOS:
    1. BORRA todos los datos diarios del mes en curso
    2. Reconstruye SOLO desde Excel (fuente de verdad)
    3. días sin dato → {} (vacío, no inventado)
    4. dias = 31 (muestra el mes completo)

    key_map: {aliados_key → [cv_studio_name, ...]}
    """
    cleared = 0
    updated = 0

    for ak, ainfo in aliados.items():
        cv_studios = key_map.get(ak)
        if cv_studios is None:
            continue
        md = (ainfo.get('data') or {}).get(mes)
        if not md:
            continue
        modelos = md.get('modelos')
        if modelos is None:
            continue

        # ── PASO A: borrar TODOS los datos diarios del mes ──
        for model in list(modelos.keys()):
            modelos[model] = {}
            cleared += 1

        # ── PASO B: reconstruir desde Excel ──
        for model in list(modelos.keys()):
            found = False
            for studio in cv_studios:
                key = (model, studio)
                if key not in cv:
                    continue
                found = True
                for d, day_vals in cv[key].items():
                    entry = {p: (day_vals.get(p) if day_vals.get(p, 0) else None)
                             for p in PLATS}
                    modelos[model][str(d)] = entry
                    updated += 1
                break

        # dias = 31 (mes completo; días sin datos quedan vacíos)
        md['dias'] = 31

    print(f"  🗑  Modelos limpiados: {cleared}")
    print(f"  ✓  Entradas reconstruidas desde Excel: {updated}")
    return aliados

# ═══════════════════════════════════════════════════════════════════
# PASO 3 — RECALCULAR PERIODOS DESDE ALIADOS
# ═══════════════════════════════════════════════════════════════════
def parse_day_from_date(date_str):
    """'27/07/2026' → 27 (para julio)."""
    try:
        parts = str(date_str).split('/')
        day, month = int(parts[0]), int(parts[1])
        return day if month == 7 else (31 if month > 7 else 0)
    except Exception:
        return 0

def recalc_gp_last(gp, aliados, aliados_key=None, mes=MES):
    """Recalcula el ÚLTIMO periodo de cada plataforma en GRUPO_PERIODOS/EXEC_PERIODOS."""
    all_mods = {}
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

def recalc_top_est(top_est, aliados_entry, mes=MES):
    """Recalcula TOP_EST (ranking de modelos del estudio) desde ALIADOS."""
    md = (aliados_entry.get('data') or {}).get(mes, {})
    all_mods = md.get('modelos') or {}

    totals = {}
    for model, days in all_mods.items():
        t = sum(
            sum((dv.get(p) or 0) for p in PLATS)
            for dv in days.values() if dv
        )
        if t > 0:
            totals[model] = t

    ranked = sorted(totals.items(), key=lambda x: -x[1])
    # Formato correcto: {'modelo': m} — igual que meses anteriores
    new_list = [{'modelo': m} for m, t in ranked]

    if top_est is None:
        return {mes: new_list}
    top_est[mes] = new_list
    return top_est

def recalc_top20g(top20g, grupo_aliados, mes=MES):
    """Recalcula TOP20G (top 20 modelos del grupo) desde ALIADOS grupo.
    Formato: {'modelo': name, 'studio': display_name}"""
    all_totals = {}  # (model, studio_display) → total
    for ak, ainfo in grupo_aliados.items():
        studio_name = ALIADOS_DISPLAY.get(ak, ak)
        md = (ainfo.get('data') or {}).get(mes, {})
        for model, days in (md.get('modelos') or {}).items():
            t = sum(
                sum((dv.get(p) or 0) for p in PLATS)
                for dv in days.values() if dv
            )
            if t > 0:
                key = (model, studio_name)
                all_totals[key] = all_totals.get(key, 0) + t

    ranked = sorted(all_totals.items(), key=lambda x: -x[1])[:20]
    new_list = [{'modelo': m, 'studio': s} for (m, s), t in ranked]
    print(f"    TOP20G: {len(new_list)} modelos | #1 → {new_list[0]['modelo'] if new_list else 'N/A'}")

    if top20g is None:
        return {mes: new_list}
    top20g[mes] = new_list
    return top20g

# ═══════════════════════════════════════════════════════════════════
# PASO 4 — TIMESTAMP
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
# MAPEOS ALIADOS KEY → CV STUDIOS (Excel col 2)
# ═══════════════════════════════════════════════════════════════════
# Nombres de display para TOP20G (aliados key → nombre visible)
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
    'Alice Steel':       ['Fabio Robledo'],   # corregido: Robledo (no Robleo)
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
    now = datetime.now()
    ts_str = now.strftime('%d/%m/%Y — %H:%M')
    print('=' * 65)
    print(f'  ACTUALIZADOR J&D — {ts_str}')
    print(f'  MODO: REBUILD COMPLETO (limpia + reconstruye desde Excel)')
    print(f'  Fórmula: day_col(d) = 19 + (31-d)*5  [5 plats/día]')
    print('=' * 65)

    # ── Leer Excel (fuente de verdad) ──
    cv, last_day_excel = read_cv_excel(CV_PATH)
    if last_day_excel == 0:
        print("⚠  No hay datos diarios en el Excel. Verificar ruta y estructura.")
        return

    def load_dash(path, al_varname):
        with open(path, 'r', encoding='utf-8') as f:
            html = f.read()
        aliados, al_q = get_var(html, al_varname)
        if aliados is None:
            print(f"  ⚠  No se encontró {al_varname}")
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
        aliados = rebuild_aliados_from_excel(aliados, cv, GRUPO_MAP)

        gp, gp_q = get_var(html, 'GRUPO_PERIODOS')
        if gp:
            gp = recalc_gp_last(gp, aliados)
            for plat in ['F4F', 'Stripchat', 'Chaturbate', 'CamSoda', 'Streamate']:
                last = (gp.get(plat) or [{}])[-1]
                if last.get('total'):
                    print(f"  ✓ GP {plat}: {last.get('label')} → {last.get('total'):.0f}")

        html = set_var(html, 'ALIADOS',        aliados, al_q)
        html = set_var(html, 'GRUPO_PERIODOS', gp,      gp_q)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['grupo'], html)

    # ════════════════════════════════════════════════════════════
    # ERIKA NOGUERA
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  ERIKA NOGUERA')
    html, aliados, al_q = load_dash(DASHBOARDS['erika'], 'ALIADOS')
    if aliados is not None:
        aliados = rebuild_aliados_from_excel(aliados, cv, ERIKA_MAP)

        ep, ep_q = get_var(html, 'EXEC_PERIODOS')
        gp, gp_q = get_var(html, 'GRUPO_PERIODOS')
        if ep: ep = recalc_gp_last(ep, aliados)
        if gp: gp = recalc_gp_last(gp, aliados)

        html = set_var(html, 'ALIADOS',       aliados, al_q)
        if ep: html = set_var(html, 'EXEC_PERIODOS', ep, ep_q)
        if gp: html = set_var(html, 'GRUPO_PERIODOS', gp, gp_q)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['erika'], html)

    # ════════════════════════════════════════════════════════════
    # FABIO ROBLEDO
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  FABIO ROBLEDO')
    html, aliados, al_q = load_dash(DASHBOARDS['fabio'], 'ALIADOS')
    if aliados is not None:
        aliados = rebuild_aliados_from_excel(aliados, cv, FABIO_MAP)

        ep, ep_q = get_var(html, 'EXEC_PERIODOS')
        if ep: ep = recalc_gp_last(ep, aliados)

        html = set_var(html, 'ALIADOS',       aliados, al_q)
        if ep: html = set_var(html, 'EXEC_PERIODOS', ep, ep_q)
        html = inject_timestamp(html, ts_str)
        save_dash(DASHBOARDS['fabio'], html)

    # ════════════════════════════════════════════════════════════
    # FORNAX STUDIOS (dashboard de estudio)
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  FORNAX STUDIOS (estudio)')
    with open(DASHBOARDS['fornax'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
        g_aliados, _ = get_var(g2, 'ALIADOS')
        fx_entry = g_aliados.get('Fornax Studios', {})

        print('  Recalculando PERIODOS desde ALIADOS Fornax:')
        periodos = recalc_studio_periodos(periodos, fx_entry)
        html = set_var(html, 'PERIODOS', periodos, pq)

        top_est_fx, te_q = get_var(html, 'TOP_EST')
        if top_est_fx is not None:
            top_est_fx = recalc_top_est(top_est_fx, fx_entry)
            html = set_var(html, 'TOP_EST', top_est_fx, te_q)
            print(f"  ✓ TOP_EST Fornax: {len((top_est_fx or {}).get(MES, []))} modelos")

        top20g_fx, t20_q = get_var(html, 'TOP20G')
        if top20g_fx is not None:
            print('  Recalculando TOP20G:')
            top20g_fx = recalc_top20g(top20g_fx, g_aliados)
            html = set_var(html, 'TOP20G', top20g_fx, t20_q)

        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['fornax'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 fornax-studios345929/index.html")

    # ════════════════════════════════════════════════════════════
    # GOLD ONLINE (dashboard de estudio)
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  GOLD ONLINE (estudio)')
    with open(DASHBOARDS['gold'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
        g_aliados, _ = get_var(g2, 'ALIADOS')
        go_entry = g_aliados.get('goldonline', {})

        print('  Recalculando PERIODOS desde ALIADOS Gold Online:')
        periodos = recalc_studio_periodos(periodos, go_entry)
        html = set_var(html, 'PERIODOS', periodos, pq)

        top_est_go, te_q = get_var(html, 'TOP_EST')
        if top_est_go is not None:
            top_est_go = recalc_top_est(top_est_go, go_entry)
            html = set_var(html, 'TOP_EST', top_est_go, te_q)
            print(f"  ✓ TOP_EST Gold: {len((top_est_go or {}).get(MES, []))} modelos")

        top20g_go, t20_q = get_var(html, 'TOP20G')
        if top20g_go is not None:
            print('  Recalculando TOP20G:')
            top20g_go = recalc_top20g(top20g_go, g_aliados)
            html = set_var(html, 'TOP20G', top20g_go, t20_q)

        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['gold'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 goldonline078939/index.html")

    # ════════════════════════════════════════════════════════════
    # CYV STUDIOS (dashboard de estudio)
    # ════════════════════════════════════════════════════════════
    print(f'\n▶  CYV STUDIOS (estudio)')
    with open(DASHBOARDS['cyv'], 'r', encoding='utf-8') as f: html = f.read()
    periodos, pq = get_var(html, 'PERIODOS')
    if periodos:
        with open(DASHBOARDS['erika'], 'r', encoding='utf-8') as f: e2 = f.read()
        e_aliados, _ = get_var(e2, 'ALIADOS')
        if e_aliados is None:
            with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g2 = f.read()
            e_aliados, _ = get_var(g2, 'ALIADOS')
        cy_entry = e_aliados.get('CyV Studios', {})

        # Cargar grupo para TOP20G
        with open(DASHBOARDS['grupo'], 'r', encoding='utf-8') as f: g3 = f.read()
        g_aliados_cy, _ = get_var(g3, 'ALIADOS')

        print('  Recalculando PERIODOS desde ALIADOS CyV:')
        periodos = recalc_studio_periodos(periodos, cy_entry)
        html = set_var(html, 'PERIODOS', periodos, pq)

        top_est_cy, te_q = get_var(html, 'TOP_EST')
        if top_est_cy is not None:
            top_est_cy = recalc_top_est(top_est_cy, cy_entry)
            html = set_var(html, 'TOP_EST', top_est_cy, te_q)
            print(f"  ✓ TOP_EST CyV: {len((top_est_cy or {}).get(MES, []))} modelos")

        top20g_cy, t20_q = get_var(html, 'TOP20G')
        if top20g_cy is not None and g_aliados_cy is not None:
            print('  Recalculando TOP20G:')
            top20g_cy = recalc_top20g(top20g_cy, g_aliados_cy)
            html = set_var(html, 'TOP20G', top20g_cy, t20_q)

        html = inject_timestamp(html, ts_str)
        with open(DASHBOARDS['cyv'], 'w', encoding='utf-8') as f: f.write(html)
        print(f"  💾 cyv-studios837357/index.html")

    print(f'\n{"=" * 65}')
    print(f'  ✅ Rebuild completado — {ts_str}')
    print(f'  📅 Excel LOCAL: último día con datos = {last_day_excel}')
    print(f'  ⚠  NO se ha hecho commit/push. Revisar antes de publicar.')
    print(f'{"=" * 65}\n')

if __name__ == '__main__':
    main()
