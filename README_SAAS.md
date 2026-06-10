# CUBE Management Contract SaaS — Multi-azienda

Versione SaaS multi-tenant.

## Cosa include

- Registrazione pubblica nuova azienda
- Super Admin SaaS
- Admin Azienda
- Staff con ruoli
- Isolamento dati per azienda tramite tenant_id
- CRM clienti
- Contratti
- Pagamenti/rate/acconti/saldi
- Lavori con data e orario facoltativo
- Documenti
- Feedback
- Fatture interne PDF
- Piani abbonamento
- PostgreSQL online tramite DATABASE_URL
- SQLite fallback locale

## Credenziali Super Admin iniziali

```text
username: superadmin
password: admin123
```

## Deploy Render

Il file render.yaml crea:

- database PostgreSQL
- web service Docker
- disco persistente per upload su `/app/uploads`

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run app.py
```
