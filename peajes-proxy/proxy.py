#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
proxy.py — Proxy seguro Flypass (FlyConnect) para el módulo Peajes de la app TECC.

La app es estática y PÚBLICA → los secretos no pueden vivir en ella. Este servicio:
  1) Guarda credenciales en .env (gitignored, NUNCA se sube).
  2) Se autentica contra AWS Cognito (InitiateAuth) y cachea el token.
  3) Consulta los movimientos de la empresa en Flypass (paginado) y devuelve JSON limpio.

Correr local:  pip install -r requirements.txt  &&  python3 proxy.py   → http://localhost:8088
Producción:    desplegar con HTTPS (Render/Railway/Cloud Run/VPS) y restringir CORS.

Flujo (según la documentación oficial de Flypass / F2X):
  Auth:  POST https://cognito-idp.us-east-1.amazonaws.com/
         headers: Content-Type: application/x-amz-json-1.1
                  X-Amz-Target: AWSCognitoIdentityProviderService.InitiateAuth
         body: {"AuthParameters":{"USERNAME":..,"PASSWORD":..},
                "AuthFlow":"USER_PASSWORD_AUTH","ClientId":..}
         → AuthenticationResult.AccessToken + .IdToken
  Movs:  GET {BASE}/api/v1/customers/{NIT}/wallet/movements/{transactionType}
              ?dateType=&startDate=&endDate=&page=&size=
         headers: X-Id-Token: <IdToken>,  Authorization: Bearer <AccessToken>
"""
import os, json, time, base64
from flask import Flask, jsonify, request
from flask_cors import CORS
import requests
try:
    from dotenv import load_dotenv; load_dotenv()
except ImportError:
    pass

CLIENT_ID = (os.getenv("FLYPASS_CLIENT_ID") or "").strip()
USERNAME  = (os.getenv("FLYPASS_USERNAME") or "").strip()
PASSWORD  = (os.getenv("FLYPASS_PASSWORD") or "").strip()
# NIT SIN dígito de verificación. Acepta FLYPASS_DOCUMENT o FLYPASS_DOCUMENT_NUMBER.
# Normalmente el NIT real viene DENTRO del token (custom:document_number), así que esto es opcional.
DOCUMENT  = (os.getenv("FLYPASS_DOCUMENT") or os.getenv("FLYPASS_DOCUMENT_NUMBER") or "").strip()
if DOCUMENT in ("__PENDIENTE__", "PENDIENTE"): DOCUMENT = ""
ENVIRON   = (os.getenv("FLYPASS_ENV") or "cert").strip().lower()    # cert | prod
PORT      = int(os.getenv("PORT", "8088"))
ALLOW_ORIGIN = os.getenv("ALLOW_ORIGIN", "*")                       # en prod: https://blueskydataanalitycs.github.io

COGNITO = ((os.getenv("FLYPASS_COGNITO_URL") or "https://cognito-idp.us-east-1.amazonaws.com").strip().rstrip("/")) + "/"
API_BASE = (os.getenv("FLYPASS_API_BASE") or "").strip().rstrip("/")   # si se define, manda sobre BASES[ENVIRON]
BASES = {
    "cert": "https://cert-api.flypass.com.co/company",
    "prod": "https://api.flypass.com.co/companyService",
}
VALID_TX = {"CONSUMPTION", "PAYMENT", "ADJUSTMENT", "COMMISSION", "ALL"}

app = Flask(__name__)
CORS(app, resources={r"/api/*": {"origins": ALLOW_ORIGIN}})

_tok = {"access": None, "id": None, "exp": 0, "doc": None}

def _jwt_claims(jwt):
    try:
        p = jwt.split(".")[1]; p += "=" * (-len(p) % 4)
        return json.loads(base64.urlsafe_b64decode(p))
    except Exception:
        return {}

def get_tokens():
    """Devuelve {access, id} cacheando hasta ~1 min antes de expirar."""
    if _tok["access"] and time.time() < _tok["exp"] - 60:
        return _tok
    if not (CLIENT_ID and USERNAME and PASSWORD):
        raise RuntimeError("Faltan FLYPASS_CLIENT_ID / FLYPASS_USERNAME / FLYPASS_PASSWORD en .env")
    body = {"AuthParameters": {"USERNAME": USERNAME, "PASSWORD": PASSWORD},
            "AuthFlow": "USER_PASSWORD_AUTH", "ClientId": CLIENT_ID}
    r = requests.post(COGNITO, data=json.dumps(body), timeout=30, headers={
        "Content-Type": "application/x-amz-json-1.1",
        "X-Amz-Target": "AWSCognitoIdentityProviderService.InitiateAuth",
    })
    r.raise_for_status()
    res = (r.json() or {}).get("AuthenticationResult", {})
    _tok["access"] = res.get("AccessToken")
    _tok["id"]     = res.get("IdToken")
    _tok["exp"]    = time.time() + int(res.get("ExpiresIn", 3600))
    if not _tok["access"] or not _tok["id"]:
        raise RuntimeError("Cognito no devolvió AccessToken/IdToken (revisa credenciales)")
    # El NIT habilitado viene DENTRO del token (custom:document_number). Así cert usa su
    # NIT de prueba y prod el NIT real, sin tener que cambiar la config.
    _tok["doc"] = _jwt_claims(_tok["id"]).get("custom:document_number") or DOCUMENT
    return _tok

def fetch_page(tx, date_type, start, end, page, size, retries=2):
    t = get_tokens()
    base = API_BASE or BASES.get(ENVIRON, BASES["cert"])
    url = f"{base}/api/v1/customers/{t['doc']}/wallet/movements/{tx}"
    params = {"dateType": date_type, "startDate": start, "endDate": end, "page": page, "size": size}
    headers = {"X-Id-Token": t["id"], "Authorization": "Bearer " + t["access"]}
    last = None
    for i in range(retries + 1):
        try:
            r = requests.get(url, params=params, timeout=75, headers=headers)
            if r.status_code != 504:        # 504 = timeout del backend de Flypass → reintentar
                return r
            last = r
        except requests.exceptions.RequestException as e:
            last = e
        time.sleep(1.5)
    if isinstance(last, Exception):
        raise last
    return last

@app.get("/health")
def health():
    return jsonify(ok=True, servicio="peajes-proxy", env=ENVIRON,
                   configurado=bool(CLIENT_ID and USERNAME and PASSWORD and DOCUMENT))

@app.get("/api/peajes")
def peajes():
    if not DOCUMENT:
        return jsonify(error="Falta FLYPASS_DOCUMENT (NIT sin DV) en .env"), 400
    tx = (request.args.get("transactionType") or "CONSUMPTION").upper()
    if tx not in VALID_TX:
        return jsonify(error=f"transactionType inválido. Usa: {sorted(VALID_TX)}"), 400
    date_type = request.args.get("dateType", "1")
    start = request.args.get("startDate"); end = request.args.get("endDate")
    if not start or not end:
        return jsonify(error="startDate y endDate son obligatorios (yyyyMMddHHmmss, máx 90 días)"), 400
    size = max(1, min(80, int(request.args.get("size", "80"))))
    # Paginar: traer todas las páginas (Flypass limita a 80 páginas/consulta)
    movimientos, page, pginfo = [], 0, {}
    try:
        while True:
            r = fetch_page(tx, date_type, start, end, page, size)
            if r.status_code != 200:
                return jsonify(error="Flypass respondió "+str(r.status_code), detalle=r.text[:800]), 502
            d = r.json() or {}
            code = str(d.get("code", ""))
            body = d.get("body") or {}
            movs = body.get("movements") or []
            movimientos.extend(movs)
            pginfo = body.get("paginationInfo") or {}
            if code not in ("000", "0003") and not movs:
                return jsonify(error="Flypass code "+code, message=d.get("message")), 502
            total_pages = int(pginfo.get("totalPages", 1) or 1)
            page += 1
            if page >= total_pages or not movs or page > 80:
                break
    except requests.exceptions.RequestException as e:
        return jsonify(error="No se pudo conectar con Flypass: "+str(e)), 502
    except Exception as e:
        return jsonify(error=str(e)), 500
    return jsonify(code="000", transactionType=tx, env=ENVIRON,
                   total=len(movimientos), movements=movimientos)

if __name__ == "__main__":
    print(f"▶ peajes-proxy en http://localhost:{PORT}  (env={ENVIRON}, NIT={'set' if DOCUMENT else 'FALTA'})")
    app.run(host="0.0.0.0", port=PORT, debug=True)
