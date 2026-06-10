# CUBE Management Contract SaaS

Versione SaaS multi-azienda / multi-tenant.

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Avvio online

Usa il file `render.yaml` su Render.

## Login iniziale Super Admin

```text
username: superadmin
password: admin123
```

## Flusso SaaS

1. Il Super Admin accede.
2. Le aziende possono registrarsi dalla pagina pubblica.
3. Ogni azienda ha il proprio Admin Azienda.
4. Ogni azienda vede solo i propri dati.
5. Lo staff vede solo ciò che il ruolo consente.

## Ruoli

- Super Admin SaaS
- Admin Azienda
- Manager Operativo
- Gestione Finanziaria
- Operativo Avanzato
- Operativo Base

## Database

- PostgreSQL se è presente `DATABASE_URL`
- SQLite locale se `DATABASE_URL` non è presente

## Struttura

```text
app.py
requirements.txt
Dockerfile
render.yaml
Procfile
start.sh
docs/
migrations/
scripts/
tests/
uploads/
data/
```
