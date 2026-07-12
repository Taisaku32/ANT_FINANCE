# Finanzas Streamlit — Repo Instructions (Claude Code)

## Purpose
Personal finance app to register and visualize incomes/expenses (fixed + one-time). Data entry is done via a small UI, and analytics/graphs are shown in a Streamlit dashboard.

## Repo layout
- `main.py`: data-entry UI (register transactions)
- `dashboard.py`: Streamlit dashboard (tables + charts + data entry). Dual mode:
  local (SQLite in `users/`) or cloud (Turso via `[usuarios]` secrets + login)
- `turso_db.py`: minimal Turso/libSQL HTTP client (Hrana v2 pipeline, sqlite3-like API)
- `migrar_a_turso.py`: one-time migration of a local user DB to Turso (local DB opened read-only)
- `DESPLIEGUE.md`: deploy guide (GitHub + Streamlit Community Cloud + Turso)
- `secrets.ejemplo.toml`: secrets template (real secrets go in Streamlit Cloud settings
  or `.streamlit/secrets.toml`, both never committed)
- `finanzas.db`: SQLite DB (root)

Per-user data (do not modify structure lightly):
- `users/<username>/finanzas.db`
- `users/<username>/finanzas.xlsx`
- `users/<username>/finanzas_gastos.xlsx`

Note: Entries created in `main.py` are also written/updated into the user Excel files.

## How to run
- Data entry: `python main.py`
- Dashboard: `streamlit run dashboard.py`

## Non-negotiables
- NEVER delete/reset any DB or Excel file.
- Keep both entrypoints working (`main.py` + `dashboard.py`).
- Schema changes must be safe migrations only (create tables / add columns). No data loss.
- UI labels remain in Spanish.
- `Monto` (amount) and `Fecha` (date) must stay required.

## Workflow expectations
- Before coding: inspect relevant files and summarize findings briefly.
- Implement changes in small, verifiable steps.
- Prefer minimal changes consistent with the existing style.
- If touching data export/import: preserve Excel compatibility (do not reorder/remove existing columns; append new ones to the right if needed).
