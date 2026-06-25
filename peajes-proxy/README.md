# Proxy de Peajes (Flypass / FlyConnect) — TECC

La app TECC es estática y **pública** (GitHub Pages), así que las credenciales de Flypass
**no pueden vivir en ella** (ni el navegador podría llamarlas por CORS). Este proxy:

1. Guarda las credenciales en `.env` (gitignored, **nunca** se sube).
2. Se autentica contra **AWS Cognito** (`InitiateAuth`, `USER_PASSWORD_AUTH`) y cachea el token.
3. Consulta los **movimientos** de la empresa en Flypass (paginado) y le entrega a la app un JSON limpio.

## Puesta en marcha (local)

```bash
cd peajes-proxy
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env          # y pega tus credenciales (ClientID, usuario, contraseña)
python3 explore.py            # prueba rápida: autentica + trae 1 página de peajes
python3 proxy.py              # levanta el proxy en http://localhost:8088
```

En la app → pestaña **Peajes** → pega la URL del proxy (ej. `http://localhost:8088`) y consulta.

## Flujo real (documentación oficial Flypass / F2X)

- **Auth (Cognito):** `POST https://cognito-idp.us-east-1.amazonaws.com/`
  - headers: `Content-Type: application/x-amz-json-1.1`, `X-Amz-Target: AWSCognitoIdentityProviderService.InitiateAuth`
  - body: `{"AuthParameters":{"USERNAME":..,"PASSWORD":..},"AuthFlow":"USER_PASSWORD_AUTH","ClientId":..}`
  - respuesta: `AuthenticationResult.AccessToken` + `.IdToken`
- **Movimientos:** `GET {BASE}/api/v1/customers/{NIT}/wallet/movements/{transactionType}?dateType=&startDate=&endDate=&page=&size=`
  - headers: `X-Id-Token: <IdToken>`, `Authorization: Bearer <AccessToken>`
  - BASE cert (pruebas): `https://cert-api.flypass.com.co/company`
  - BASE prod: `https://api.flypass.com.co/companyService`
- **NIT** sin dígito de verificación. TECC = `800126547`.
- `transactionType`: `CONSUMPTION` (peajes) · `PAYMENT` (recargas) · `ADJUSTMENT` · `COMMISSION` · `ALL`.
- `dateType`: `0` fecha de aplicación · `1` fecha de movimiento. Fechas `yyyyMMddHHmmss`, **máx 90 días**. `size` ≤ 80.
- Cada movimiento trae `flyKeys` (placa `PLATE` y `TAG`), `attentionPoint` (punto de peaje), `direction`, `amount`.

## Endpoints del proxy

- `GET /health` — estado y si está configurado.
- `GET /api/peajes?transactionType=CONSUMPTION&dateType=1&startDate=YYYYMMDDHHMMSS&endDate=YYYYMMDDHHMMSS[&size=80]`
  - autentica, **pagina automáticamente** y devuelve `{code, total, movements:[...]}`.

## Despliegue (para que la app publicada lo use)

Súbelo a un host con **HTTPS** (Render / Railway / Cloud Run / VPS), define las variables de
entorno (las mismas del `.env`), y en `.env` pon `ALLOW_ORIGIN=https://blueskydataanalitycs.github.io`.
Luego pega esa URL pública en la pestaña Peajes.

> ⚠️ **Seguridad:** las credenciales que se compartieron por chat quedaron expuestas; conviene
> **rotarlas** con Flypass. Nunca las pongas en el repo ni en la app: solo en `.env` / variables
> de entorno del proxy.
