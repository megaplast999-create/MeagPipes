# Mega Pipes Quotation System — Backend Scaffolding (Phase 1)

FastAPI + PostgreSQL + Redis backend for the Mega Pipes Quotation System Web Pilot, configured for Render deployment.

## Architecture

- **Web Framework:** FastAPI (async routes, Pydantic v2 schemas, OpenAPI/Swagger docs)
- **Database:** PostgreSQL with SQLAlchemy 2.0 ORM & Alembic migrations
- **Cache & Rate Limiting:** Redis with strict **Fail-Closed** policy (HTTP 503 if Redis is unavailable)
- **Authentication:** Single-user pilot session model with HTTP-only cookies, bcrypt hashing, 30-min idle timeout, and 8-hr absolute lifetime
- **Calculation Engine Adapter:** Thin adapter integrating `Quotation System/calculation_engine.py` with zero commercial calculation drift
- **Excel Export Service:** `openpyxl`-based generator producing fully editable, password-free `.xlsx` customer quotations (Offer + Synopsis)

## Directory Structure

```
backend/
  app/
    api/              # Route handlers (auth, quotations)
    core/             # Configuration, security, rate limiting, logging
    db/               # SQLAlchemy models and session management
    services/         # Calculation engine adapter and Excel export service
    main.py           # Application factory and middleware
  alembic/            # Database migration scripts
  tests/              # Security, parity, rate limit, and API test suite
  requirements.txt    # Pinned Python dependencies
  render.yaml         # Render Blueprint deployment definition
  alembic.ini         # Alembic database migration config
  .env.example        # Environment variable template (placeholders only)
  .gitignore          # Git exclusions for Python & local environments
```

## Local Development Setup

1. **Environment Setup:**
   ```bash
   cd backend
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   pip install -r requirements.txt
   cp .env.example .env
   ```

2. **Run Tests:**
   ```bash
   pytest tests/ -v
   ```

3. **Run API Server:**
   ```bash
   uvicorn app.main:app --reload --port 8000
   ```

4. **API Documentation:**
   - Swagger UI: `http://localhost:8000/docs`
   - ReDoc: `http://localhost:8000/redoc`

## Safety & Parity Principles

- **Commercial Calculation Parity:** Recalculations use the authoritative `calculation_engine.py`. Golden baselines (Approved CostSheet ₹17,61,17,500, Legacy ₹17,86,57,900, Standard Tender ₹17,88,01,050) are verified via automated tests.
- **Desktop Isolation:** No modifications to `DEV/Mega_Pipes_Quotation_DEV.xlsm`, VBA modules, or desktop backup files.
- **Fail-Closed Security:** Rate limiting rejects sensitive operations (HTTP 503) if Redis cannot be reached.
- **Password-Free Export:** Generated customer workbooks are unlocked, unpassworded, and free of external workbook links.
