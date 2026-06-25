#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
explore.py — Probador rápido del flujo Flypass SIN levantar el servidor.
  1) Autentica contra Cognito y muestra si llegaron AccessToken/IdToken.
  2) Consulta una página de movimientos (CONSUMPTION = peajes) de los últimos 30 días.

Uso:
  cp .env.example .env   # y rellena tus credenciales
  pip install -r requirements.txt
  python3 explore.py
"""
import os, json, time, datetime, sys
import requests
try:
    from dotenv import load_dotenv; load_dotenv()
except ImportError:
    pass

CLIENT_ID = (os.getenv("FLYPASS_CLIENT_ID") or "").strip()
USERNAME  = (os.getenv("FLYPASS_USERNAME") or "").strip()
PASSWORD  = (os.getenv("FLYPASS_PASSWORD") or "").strip()
DOCUMENT  = (os.getenv("FLYPASS_DOCUMENT") or "").strip()
ENVIRON   = (os.getenv("FLYPASS_ENV") or "cert").strip().lower()
BASES = {"cert": "https://cert-api.flypass.com.co/company",
         "prod": "https://api.flypass.com.co/companyService"}

if not (CLIENT_ID and USERNAME and PASSWORD and DOCUMENT):
    sys.exit("⚠️  Faltan credenciales en .env (FLYPASS_CLIENT_ID/USERNAME/PASSWORD/DOCUMENT).")

print("→ 1) Autenticando con Cognito…")
auth = requests.post("https://cognito-idp.us-east-1.amazonaws.com/",
    data=json.dumps({"AuthParameters": {"USERNAME": USERNAME, "PASSWORD": PASSWORD},
                     "AuthFlow": "USER_PASSWORD_AUTH", "ClientId": CLIENT_ID}),
    headers={"Content-Type": "application/x-amz-json-1.1",
             "X-Amz-Target": "AWSCognitoIdentityProviderService.InitiateAuth"}, timeout=30)
print("   status:", auth.status_code)
if auth.status_code != 200:
    print(auth.text[:800]); sys.exit("❌ Falló la autenticación (revisa ClientId/usuario/clave).")
res = auth.json().get("AuthenticationResult", {})
access, idt = res.get("AccessToken"), res.get("IdToken")
print("   AccessToken:", "OK" if access else "FALTA", "| IdToken:", "OK" if idt else "FALTA",
      "| expira en", res.get("ExpiresIn"), "s")
if not (access and idt): sys.exit("❌ Cognito no devolvió los tokens esperados.")

end = datetime.datetime.now()
start = end - datetime.timedelta(days=30)
fmt = lambda d: d.strftime("%Y%m%d%H%M%S")
url = f"{BASES[ENVIRON]}/api/v1/customers/{DOCUMENT}/wallet/movements/CONSUMPTION"
params = {"dateType": 1, "startDate": fmt(start), "endDate": fmt(end), "page": 0, "size": 80}
print(f"\n→ 2) Movimientos (CONSUMPTION, últimos 30 días) en {ENVIRON}…")
print("   GET", url, params)
r = requests.get(url, params=params, timeout=60,
                 headers={"X-Id-Token": idt, "Authorization": "Bearer " + access})
print("   status:", r.status_code)
try:
    d = r.json()
    print("   code:", d.get("code"), "| message:", d.get("message"))
    body = d.get("body") or {}
    movs = body.get("movements") or []
    print("   movimientos en esta página:", len(movs), "| paginación:", body.get("paginationInfo"))
    print(json.dumps(movs[:2], indent=2, ensure_ascii=False)[:1500])
except Exception:
    print(r.text[:1500])
