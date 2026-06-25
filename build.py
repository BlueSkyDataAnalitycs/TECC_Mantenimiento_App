#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build.py — App de Mantenimiento TECC
Descarga los dos Google Sheets (Aseos de vehículos + Taller interno),
procesa los aseos, asigna valor por combinación y genera un index.html autónomo.

Uso:   python3 build.py
Salida: index.html  (doble clic para abrir)  +  datos.json
"""
import csv, io, json, sys, re, urllib.request, datetime, os

# ---- IDs de los Google Sheets (link-sharing: lector) ----
SHEET_ASEOS  = "1cTsB7riRbd7u0h3DsTdlItYnUPg_b-s2YcCXIhtx9HY"
SHEET_TALLER = "1jaxw_YLSa9KNh0F7LZWyTpaixUIGYQpvJloV_yOwvZ0"
GVIZ = "https://docs.google.com/spreadsheets/d/{id}/gviz/tq?tqx=out:csv"

# ---- Tabla de precios por combinación exacta de "Tipo de aseo" (COP) ----
# Editable también dentro de la app (módulo Ajustes). Estos son los valores por defecto.
PRECIOS = {
    "ASEO EXTERIOR, LAVADO DE LLANTAS + LLANTIL": 30000,
    "ASEO EXTERIOR, BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA, LAVADO DE LLANTAS + LLANTIL": 120000,
    "ASEO EXTERIOR, BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE LLANTAS + LLANTIL": 50000,
    "ASEO EXTERIOR, TRAPIADA INTERNA, LAVADO DE LLANTAS + LLANTIL": 40000,
    "ASEO EXTERIOR, BARRIDO INTERNO, LAVADO DE LLANTAS + LLANTIL": 40000,
    "ASEO EXTERIOR, BRILLADA": 150000,
    "BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA": 90000,
    "ASEO EXTERIOR, DESMANCHADA, BRILLADA, LAVADO DE LLANTAS + LLANTIL": 180000,
    "LAVADO DE LLANTAS + LLANTIL": 10000,
    "LAVADO DE COJINERIA": 70000,
    "ASEO EXTERIOR, BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA, DESMANCHADA, BRILLADA, LAVADO DE LLANTAS + LLANTIL": 270000,
    "ASEO EXTERIOR, DESMANCHADA": 40000,
    "BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA, LAVADO DE LLANTAS + LLANTIL": 100000,
    "BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA, DESMANCHADA": 110000,
    "ASEO EXTERIOR, BARRIDO INTERNO, LAVADO DE COJINERIA, LAVADO DE LLANTAS + LLANTIL": 110000,
    "DESMANCHADA": 20000,
    "ASEO EXTERIOR, BARRIDO INTERNO, TRAPIADA INTERNA, LAVADO DE COJINERIA": 110000,
}

# ---- Taller: contexto económico fijo (master prompt INDICADOR TALLER JULIAN) ----
TALLER_CONFIG = {
    "costoSemanal": 1875000,
    "tablaExterna": {  # mano de obra + servicio promedio (sin repuestos)
        "Diagnóstico / revisión sistema de carga": 100000,
        "Cambio de bombillos (incl. LED)": 60000,
        "Blower / ventilador de cabina": 80000,
        "Instalación / cambio de exploradoras": 40000,
        "Luces traseras / laterales": 70000,
        "Motor de arranque / reparación": 120000,
        "Direccionales / estacionarias / flasher": 60000,
        "Sensor de reversa": 50000,
        "Motoventiladores (cambio)": 140000,
        "Puerta eléctrica / corrediza": 150000,
        "Cambio / tensado de correas": 30000,
        "Tablero (bajar/armar/reparar)": 150000,
        "Farolas / faros": 70000,
        "Relés de motoventiladores": 40000,
        "Stop / plaqueta de stop": 30000,
        "Relés de luces": 40000,
        "Sistema limpiaparabrisas": 60000,
        "Fusibles": 40000,
        "Compresor A/A": 70000,
        "Instalación de rutero": 30000,
        "Arreglo compartimiento motor": 40000,
        "Escape (soporte / reparación)": 40000,
        "Sensor cigüeñal": 40000,
        # categorías añadidas a partir de los informes históricos del taller
        "Relés (general)": 40000,
        "Batería / bornes / puentes": 40000,
        "Dispositivo de velocidad / tacógrafo": 80000,
        "Elevavidrios (motor / switch)": 70000,
        "Frenos (válvula / ajuste)": 90000,
        "Cuchilla / accesorio de cabina": 30000,
        "Carrocería (babero / bómper)": 50000,
        "Instalación eléctrica / accesorios": 60000,
        "Arreglo de luces (general)": 50000,
        "Domicilio / traslado": 25000,
    },
}

# ---- Modelo de costo de salario (mensual, COP) — editable en la app (Ajustes) ----
# Costo total empleador 2026 ≈ $2.907.595 = SMMLV + auxilio de transporte + prestaciones/parafiscales.
# Asignación de días: NIDIA GASCA, CARLOS GASCA y TECC = 4 días/mes c/u; WG = días marcados como "Finca".
SALARIO_CONFIG = {
    "salarioBase": 1750905,    # SMMLV 2026
    "auxTransporte": 249095,   # auxilio de transporte 2026
    "prestaciones": 907595,    # prestaciones sociales + parafiscales
    "diaDivisor": 24,          # día base = costo total ÷ 24
    "diasFijos": 4,            # días/mes fijos para NIDIA, CARLOS, TECC
}


def fetch_csv(sheet_id):
    url = GVIZ.format(id=sheet_id)
    req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
    with urllib.request.urlopen(req, timeout=40) as r:
        raw = r.read().decode("utf-8", errors="replace")
    return list(csv.reader(io.StringIO(raw)))


def col_index(header, *needles):
    """Devuelve el índice de la primera columna cuyo encabezado contiene TODOS los needles."""
    for i, h in enumerate(header):
        hl = (h or "").lower()
        if all(n.lower() in hl for n in needles):
            return i
    return -1


def col_exact(header, name):
    """Índice de la columna cuyo encabezado es EXACTAMENTE name (evita coincidencias parciales)."""
    name = name.strip().lower()
    for i, h in enumerate(header):
        if (h or "").strip().lower() == name:
            return i
    return -1


def drive_id(s):
    """Extrae el ID de archivo de un enlace de Google Drive (formato Forms: open?id=ID)."""
    m = re.search(r"[-\w]{25,}", s or "")
    return m.group(0) if m else ""


def norm_veh(v):
    """Normaliza el N° de vehículo: O→0, quita .0 y ceros a la izquierda ('034'=='34')."""
    v = (v or "").strip().upper().replace("O", "0").replace(".0", "")
    if v.isdigit():
        v = str(int(v))
    return v


def parse_date(s):
    s = (s or "").strip()
    if not s:
        return None
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d",
                "%d-%m-%Y", "%m/%d/%Y %H:%M:%S"):
        try:
            return datetime.datetime.strptime(s, fmt)
        except ValueError:
            continue
    # último intento: solo la parte de fecha
    try:
        return datetime.datetime.strptime(s.split()[0], "%d/%m/%Y")
    except Exception:
        return None


def dur_min(ini, fin):
    def t(x):
        x = (x or "").strip().lower().replace("a. m.", "am").replace("p. m.", "pm")
        x = x.replace("a.m.", "am").replace("p.m.", "pm").replace(" ", "")
        for fmt in ("%I:%M:%S%p", "%I:%M%p", "%H:%M:%S", "%H:%M"):
            try:
                return datetime.datetime.strptime(x, fmt)
            except ValueError:
                continue
        return None
    a, b = t(ini), t(fin)
    if a and b:
        d = (b - a).total_seconds() / 60
        return round(d) if 0 < d < 24 * 60 else None
    return None


def aseo_parte_b(ts):
    """22-jun-2026 12:00 → parte B (formulario corregido). Antes → parte A (cruzado)."""
    d = parse_date(ts)
    if not d:
        return False
    day = datetime.datetime(d.year, d.month, d.day)
    cut = datetime.datetime(2026, 6, 22)
    if day < cut:
        return False
    if day > cut:
        return True
    hh = hora_de(ts)
    return hh is not None and hh >= 12


def parse_veh_b(q):
    """'LFL492 - 2' → ('LFL492','2'); '277' → ('','277')."""
    q = (q or "").strip()
    if not q:
        return "", ""
    parts = re.split(r"\s*[-–/]\s*", q)
    if len(parts) >= 2:
        return parts[0].strip().upper(), norm_veh(parts[-1])
    return "", norm_veh(q)


def process_aseos(rows):
    if not rows:
        return [], []
    h = rows[0]
    # Corte 22-jun-2026 mediodía: parte A (cruzado: vehículo en "Lugar de Aseo"/C, origen en "Vehiculo"/Q o "Lugar de lavado"/P)
    # vs parte B (corregido: vehículo en "Vehiculo"/Q como "PLACA - N° interno", origen en "Lugar de Aseo"/C).
    iVeh = col_index(h, "lugar de aseo")
    if iVeh < 0:
        iVeh = col_index(h, "numero interno")
    iOrigen = col_exact(h, "vehiculo")          # origen primario (incl. CASA CONDUCTOR / proveedor)
    iOrigen2 = col_index(h, "lugar de lavado")  # origen de respaldo (form viejo)
    iCC = col_index(h, "centro de costo")
    iCond = col_index(h, "conductor")
    iDia = col_index(h, "dia de lavado")
    iIni = col_index(h, "hora de inicio")
    iFin = col_index(h, "finaliz")
    iTipo = col_index(h, "tipo de aseo")
    iValExt = col_index(h, "valor", "externo")   # valor del aseo si lo hizo un proveedor externo
    # fotos del vehículo (subidas al Form → enlaces de Drive). Orden fijo = labels en template.html.
    foto_cols = [col_index(h, "foto delantera"), col_index(h, "izquierdo"), col_index(h, "derecho"),
                 col_index(h, "trasera"), col_index(h, "parte interna"), col_index(h, "foto interna")]
    recs = []
    months = set()
    seen = {}
    for row in rows[1:]:
        if iTipo < 0 or iTipo >= len(row):
            continue
        tipo = (row[iTipo] or "").strip()
        if not tipo:
            continue
        # fecha operativa: día de lavado; si falta, marca temporal (col 0)
        fecha = parse_date(row[iDia]) if (0 <= iDia < len(row) and row[iDia].strip()) else parse_date(row[0] if row else "")
        if not fecha:
            continue
        # resolver vehículo y origen según el corte del 22-jun
        if aseo_parte_b(row[0] if row else ""):
            qv = row[iOrigen].strip() if 0 <= iOrigen < len(row) else ""
            placa, num = parse_veh_b(qv)
            veh = num or norm_veh(qv)
            origen = row[iVeh].strip() if 0 <= iVeh < len(row) else ""
        else:
            placa = ""
            veh = norm_veh(row[iVeh].strip() if 0 <= iVeh < len(row) else "")
            origen = (row[iOrigen].strip() if 0 <= iOrigen < len(row) else "") \
                or (row[iOrigen2].strip() if 0 <= iOrigen2 < len(row) else "")
        fch = fecha.strftime("%Y-%m-%d")
        # id estable por-aseo (debe coincidir con processAseos del template.html)
        base_key = f"{fch}|{veh}|{tipo}"
        k = seen.get(base_key, 0)
        seen[base_key] = k + 1
        fot = [drive_id(row[i]) if 0 <= i < len(row) else "" for i in foto_cols]
        rec = {
            "id": f"{base_key}|{k}",
            "ym": fecha.strftime("%Y-%m"),
            "fecha": fch,
            "vehiculo": veh,
            "placa": placa,
            "centro": (row[iCC].strip() if 0 <= iCC < len(row) else "") or "(sin CC)",
            "conductor": (row[iCond].strip() if 0 <= iCond < len(row) else ""),
            "tipo": tipo,
            "durMin": dur_min(row[iIni] if 0 <= iIni < len(row) else "",
                              row[iFin] if 0 <= iFin < len(row) else ""),
            "lugar": (row[iOrigen2].strip() if 0 <= iOrigen2 < len(row) else ""),
            "lugarAseo": origen,   # origen resuelto (clasificación lavador/conductor/proveedor)
            "valExt": int("".join(c for c in (row[iValExt] if 0 <= iValExt < len(row) else "") if c.isdigit()) or 0),
        }
        if any(fot):
            rec["fot"] = fot
        recs.append(rec)
        months.add(fecha.strftime("%Y-%m"))
    return recs, sorted(months)


def process_taller(rows):
    if not rows:
        return {"headers": [], "records": []}
    h = rows[0]
    iId = col_index(h, "id del mantenimiento")
    iSis = col_index(h, "sistema")
    iEsp = col_index(h, "especifique")
    iEst = col_index(h, "estado")
    iPlaca = col_index(h, "placa")
    iHIni = col_index(h, "hora", "inicio")   # hora de inicio (formulario nuevo)
    recs = []
    for row in rows[1:]:
        if not any((c or "").strip() for c in row):
            continue
        fecha = parse_date(row[0] if row else "")
        rec = {
            "id": (row[iId].strip() if 0 <= iId < len(row) else ""),
            "sistema": (row[iSis].strip() if 0 <= iSis < len(row) else ""),
            "especifique": (row[iEsp].strip() if 0 <= iEsp < len(row) else ""),
            "estado": (row[iEst].strip() if 0 <= iEst < len(row) else ""),
            "placa": (row[iPlaca].strip() if 0 <= iPlaca < len(row) else ""),
            "fecha": fecha.strftime("%Y-%m-%d") if fecha else "",
        }
        hh = None
        ti = row[iHIni] if 0 <= iHIni < len(row) else ""
        mt = re.search(r"(\d{1,2}):(\d{2})", str(ti or ""))
        if mt:
            hh = int(mt.group(1))
            low = str(ti).lower()
            if "p" in low and hh < 12:
                hh += 12
            if "a" in low and hh == 12:
                hh = 0
        if hh is None:
            hh = hora_de(row[0] if row else "")
        if hh is not None and 0 <= hh < 24:
            rec["hora"] = hh
        recs.append(rec)
    return {"headers": h, "records": recs}


def hora_de(s):
    """Hora (0-23) de una marca temporal 'DD/MM/YYYY HH:MM:SS'; None si no hay."""
    m = re.search(r"\d{1,2}/\d{1,2}/\d{2,4}[ ,]+(\d{1,2}):(\d{2})(?::\d{2})?\s*(a\.?\s*m\.?|p\.?\s*m\.?|am|pm)?", str(s or ""), re.I)
    if not m:
        return None
    h = int(m.group(1))
    ap = (m.group(3) or "").lower().replace(".", "").replace(" ", "")
    if ap == "pm" and h < 12:
        h += 12
    if ap == "am" and h == 12:
        h = 0
    return h if 0 <= h < 24 else None


def main():
    here = os.path.dirname(os.path.abspath(__file__))
    print("· Descargando Sheet de aseos…")
    try:
        aseos_rows = fetch_csv(SHEET_ASEOS)
    except Exception as e:
        print("  ⚠️  No se pudo descargar aseos:", e); aseos_rows = []
    print("· Descargando Sheet de taller…")
    try:
        taller_rows = fetch_csv(SHEET_TALLER)
    except Exception as e:
        print("  ⚠️  No se pudo descargar taller:", e); taller_rows = []

    recs, months = process_aseos(aseos_rows)
    taller = process_taller(taller_rows)

    # Base de datos histórica del taller (informes ene–jun 2026). SIEMPRE presente.
    # Las intervenciones nuevas llegan por el Google Sheet y se unen a estas.
    taller_seed = []
    seed_path = os.path.join(here, "taller_seed.json")
    if os.path.exists(seed_path):
        with open(seed_path, encoding="utf-8") as f:
            taller_seed = json.load(f)
    taller_all = taller_seed + taller["records"]   # seed + lo que haya en el Sheet

    # Base de datos de vehículos (consolidada de los 2 Excel; SIN PII). Ver import_vehiculos.py
    vehiculos = []
    veh_path = os.path.join(here, "vehiculos.json")
    if os.path.exists(veh_path):
        with open(veh_path, encoding="utf-8") as f:
            vehiculos = json.load(f).get("vehiculos", [])

    data = {
        "generated_at": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
        "aseos": {"records": recs, "months": months},
        "taller": {"headers": taller.get("headers", []), "records": taller_all},
        "tallerSeed": taller_seed,
        "vehiculos": vehiculos,
        "precios": PRECIOS,
        "tallerConfig": TALLER_CONFIG,
        "salarioConfig": SALARIO_CONFIG,
    }

    with open(os.path.join(here, "datos.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    with open(os.path.join(here, "template.html"), encoding="utf-8") as f:
        tpl = f.read()
    html = tpl.replace("__DATA_JSON__", json.dumps(data, ensure_ascii=False))
    with open(os.path.join(here, "index.html"), "w", encoding="utf-8") as f:
        f.write(html)

    # resumen
    total = sum(PRECIOS.get(r["tipo"], 0) for r in recs)
    sin = sum(1 for r in recs if r["tipo"] not in PRECIOS)
    print("\n✅ Generado index.html")
    print(f"   Aseos: {len(recs)} registros · {len(months)} meses ({months[0] if months else '—'} → {months[-1] if months else '—'})")
    print(f"   Valor total histórico: ${total:,.0f}  | sin valor: {sin}")
    print(f"   Taller: {len(taller_all)} intervenciones ({len(taller_seed)} histórico + {len(taller['records'])} del Sheet)")


if __name__ == "__main__":
    main()
