import sys, json, base64, re, urllib.request, urllib.parse, importlib.util, os

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SUPA_URL   = "https://mqafwoaghwhnvorerjha.supabase.co"
SUPA_KEY   = ""

sp = os.path.join(SCRIPT_DIR, "secrets_local.py")
spec = importlib.util.spec_from_file_location("secrets_local", sp)
sm = importlib.util.module_from_spec(spec)
spec.loader.exec_module(sm)
SUPA_KEY = getattr(sm, "SUPA_SERVICE_KEY", "")
print(f"Clave: {'OK (' + SUPA_KEY[:10] + '...)' if SUPA_KEY else 'VACÍA'}")

ESTUDIOS = {
    "cyv-studios837357":    "estudios/cyv-studios837357/index.html",
    "fornax-studios345929": "estudios/fornax-studios345929/index.html",
    "goldonline078939":     "estudios/goldonline078939/index.html",
}

def fetch(estudio_id):
    url = (f"{SUPA_URL}/rest/v1/liquidaciones"
           f"?estudio_id=eq.{urllib.parse.quote(estudio_id)}"
           f"&select=*&order=fecha.desc")
    req = urllib.request.Request(url, headers={
        "apikey": SUPA_KEY,
        "Authorization": f"Bearer {SUPA_KEY}",
        "Content-Type": "application/json",
    })
    with urllib.request.urlopen(req, timeout=15) as r:
        rows = json.loads(r.read().decode())
    result = []
    for r in rows:
        extras = r.get("datos_extra") or {}
        if not extras.get("colillaGenerada"):
            continue
        obj = dict(extras)
        obj.update({
            "id": r.get("id"), "estudioId": r.get("estudio_id") or estudio_id,
            "estudioNombre": r.get("estudio_nombre") or estudio_id,
            "fecha": r.get("fecha") or "", "netoCOP": r.get("pago_estudio") or 0,
            "brutoCOP": r.get("valor_total") or 0, "pagoEstudio": r.get("pago_estudio") or 0,
            "estado": r.get("estado") or "Pendiente", "pagado": bool(r.get("pagado")),
            "colillaGenerada": True,
        })
        result.append(obj)
    return result

for eid, rel_path in ESTUDIOS.items():
    path = os.path.join(SCRIPT_DIR, rel_path)
    print(f"\n{eid}:", end=" ", flush=True)
    try:
        rows = fetch(eid)
        print(f"{len(rows)} colillas")
        b64 = base64.b64encode(json.dumps(rows, ensure_ascii=False).encode()).decode()
        with open(path, "r", encoding="utf-8") as f:
            html = f.read()
        new_html = re.sub(r"var COLILLAS_DATA_B64='[^']*'", f"var COLILLAS_DATA_B64='{b64}'", html, count=1)
        with open(path, "w", encoding="utf-8") as f:
            f.write(new_html)
        print(f"  Inyectado OK")
    except Exception as e:
        print(f"  ERROR: {e}")

print("\nListo.")
