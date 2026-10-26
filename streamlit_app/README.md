# Mega Pipes Web Pilot — Streamlit Quotation Application

Enterprise internal quotation management and commercial cost-sheet interface built with Streamlit and Python.

## Architecture

```
[ Browser User ]
       │
       ▼
[ Streamlit App (streamlit_app/app.py) ]
       │
       │ HTTP / REST with Session Cookies & X-CSRF-Token
       ▼
[ Render FastAPI Backend (https://meagpipes-1.onrender.com) ]
       │
       ├─► PostgreSQL 16 Database (Quotations, Audit Logs, Snapshots)
       ├─► Redis Rate Limiter & Account Lockout
       └─► Authoritative calculation_engine.py & ExcelExportService
```

## Features

- **Session Authentication:** Authenticates against `POST /api/auth/login` and restores session via `GET /api/auth/me`. Complete per-user session isolation via `st.session_state`.
- **CSRF Protection:** Automatically extracts the `megapipes_csrf` cookie and injects `X-CSRF-Token` headers for mutating operations (`POST`, `PUT`).
- **Dashboard:** Search, filter, and inspect saved quotations with status badges (`DRAFT` vs `CALCULATED`).
- **Quotation Editor:** Comprehensive quotation header and multi-line pipe configuration (OD/ID, thickness, quantity, unit, 3LPE external, epoxy internal, commercial surcharges, and overrides).
- **Authoritative Calculations:** Triggers `POST /api/quotations/{id}/calculate` and visualizes exact server-side cost sheet snapshots (Grand Total, Subtotal, GST, Total Metres, Total MT, Logistics Trips).
- **Commercial Excel Export:** Direct binary streaming of password-free `.xlsx` quotation workbooks via `GET /api/quotations/{id}/export`.

---

## Local Development

1. **Install dependencies:**
   ```bash
   pip install -r streamlit_app/requirements.txt
   ```

2. **Run locally:**
   ```bash
   streamlit run streamlit_app/app.py
   ```

3. **Backend Target:**
   By default, the app targets the live production Render backend:
   `https://meagpipes-1.onrender.com`

   To target a local FastAPI backend instead:
   ```bash
   set MEGAPIPES_BACKEND_URL=http://localhost:8000
   streamlit run streamlit_app/app.py
   ```

---

## Deployment Options

### Option 1: Streamlit Community Cloud (Recommended & Free)

1. Navigate to [share.streamlit.io](https://share.streamlit.io).
2. Connect your GitHub account and select repository: `megaplast999-create/MeagPipes`.
3. Choose branch: `web-pilot-backend` (or `main`).
4. Set **Main file path**: `streamlit_app/app.py`.
5. Under Advanced Settings -> Secrets (Optional):
   ```toml
   BACKEND_URL = "https://meagpipes-1.onrender.com"
   ```
6. Click **Deploy!** Streamlit Cloud will automatically build and publish with a public `https://<app-name>.streamlit.app` URL.

### Option 2: Render Web Service

1. Create a new Web Service on Render pointing to `megaplast999-create/MeagPipes`.
2. Environment: `Python 3`.
3. Build Command: `pip install -r streamlit_app/requirements.txt`.
4. Start Command: `streamlit run streamlit_app/app.py --server.port $PORT --server.address 0.0.0.0`.
