#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
explore.py — Prueba rápida del flujo Flypass SIN levantar el servidor.
  1) Autentica contra Cognito y muestra los tokens + el NIT del token.
  2) Consulta movimientos ALL (últimos 90 días) para uno o varios NIT.

Uso:
  python3 explore.py                 # usa el NIT del token (custom:document_number)
  python3 explore.py 100000004 800126547   # prueba esos NIT explícitamente
"""
import os, json, time, base64, datetime, sys
import requests
from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

CLIENT_ID = (os.getenv("FLYPASS_CLIENT_ID") or "").strip()
USERNAME  = (os.getenv("FLYPASS_USERNAME") or "").strip()
PASSWORD  = (os.getenv("FLYPASS_PASSWORD") or "").strip()
COGNITO   = ((os.getenv("FLYPASS_COGNITO_URL") or "https://cognito-idp.us-east-1.amazonaws.com").rstrip("/")) + "/"
_DEFBASE  = {"cert": "https://cert-api.flypass.com.co/company",
             "prod": "https://api.flypass.com.co/companyService"}.get((os.getenv("FLYPASS_ENV") or "cert").lower())
BASE      = (os.getenv("FLYPASS_API_BASE") or _DEFBASE).rstrip("/")
docs_env  = (os.getenv("FLYPASS_DOCUMENT") or os.getenv("FLYPASS_DOCUMENT_NUMBER") or "").strip()
if docs_env in ("__PENDIENTE__", "PENDIENTE"): docs_env = ""

if not (CLIENT_ID and USERNAME and PASSWORD):
    sys.exit("⚠️  Faltan credenciales en .env (FLYPASS_CLIENT_ID/USERNAME/PASSWORD).")

def claims(jwt):
    p = jwt.split(".")[1]; p += "=" * (-len(p) % 4)
    return json.loads(base64.urlsafe_b64decode(p))

print("→ 1) Autenticando con Cognito…")
auth = requests.post(COGNITO, data=json.dumps({
    "AuthParameters": {"USERNAME": USERNAME, "PASSWORD": PASSWORD},
    "AuthFlow": "USER_PASSWORD_AUTH", "ClientId": CLIENT_ID}),
    headers={"Content-Type": "application/x-amz-json-1.1",
             "X-Amz-Target": "AWSCognitoIdentityProviderService.InitiateAuth"}, timeout=30)
print("   status:", auth.status_code)
if auth.status_code != 200:
    print(auth.text[:600]); sys.exit("❌ Falló la autenticación (revisa ClientId/usuario/clave).")
res = auth.json()["AuthenticationResult"]; acc = res["AccessToken"]; idt = res["IdToken"]
tokdoc = claims(idt).get("custom:document_number")
print("   AccessToken: OK | IdToken: OK | NIT habilitado por el token:", tokdoc)

docs = sys.argv[1:] or ([docs_env] if docs_env else [tokdoc])
end = datetime.datetime.now(); start = end - datetime.timedelta(days=90)
f = lambda d: d.strftime("%Y%m%d%H%M%S")
H = {"X-Id-Token": idt, "Authorization": "Bearer " + acc}
for doc in docs:
    print(f"\n→ 2) Movimientos ALL para NIT {doc} (últimos 90 días) en {BASE}…")
    for i in range(3):
        try:
            r = requests.get(f"{BASE}/api/v1/customers/{doc}/wallet/movements/ALL",
                params={"dateType": 1, "startDate": f(start), "endDate": f(end), "page": 0, "size": 50},
                headers=H, timeout=75)
            if r.status_code == 504:
                print("   504 (timeout del backend Flypass), reintentando…"); time.sleep(2); continue
            j = r.json(); movs = (j.get("body") or {}).get("movements") or []
            print(f"   HTTP {r.status_code} | code {j.get('code')} | {j.get('message')} | movimientos: {len(movs)}")
            if movs:
                print(json.dumps(movs[:2], indent=1, ensure_ascii=False)[:1300])
            break
        except Exception as e:
            print("   error:", e); break
