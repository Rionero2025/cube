# CUBE Management Contract — Pacchetto unico tutto insieme

Questa versione è pensata per essere caricata su GitHub/Render come UNICO progetto.

Non devi separare sito e gestionale.

## Cosa contiene

Nel progetto root trovi direttamente:

- `app.py` — portale completo Streamlit
- `requirements.txt`
- `Dockerfile`
- `render.yaml`
- `Procfile`
- `start.sh`
- `.streamlit/config.toml`
- `docs/`
- `migrations/`
- `scripts/`
- `tests/`
- `public_site/` — sito statico opzionale incluso come materiale commerciale

## Dentro app.py ci sono già insieme

### Parte pubblica
- home page del programma
- spiegazione del gestionale
- pacchetti
- registrazione azienda
- prova gratuita 30 giorni
- login

### Parte SaaS
- Super Admin
- aziende SaaS
- utenti globali
- piani abbonamento
- dati SaaS
- Admin Azienda
- CRM clienti
- contratti
- pagamenti/rate
- lavori
- documenti
- feedback
- fatture interne
- staff e permessi

## Link diretti supportati

- `/?public_page=login`
- `/?public_page=plans`
- `/?public_page=register&plan=Starter`
- `/?public_page=register&plan=Professional`
- `/?public_page=register&plan=Business`

## Login iniziale Super Admin

```text
username: superadmin
password: admin123
```

## Deploy

Carica TUTTI i file della root su GitHub, poi su Render fai Manual Sync / Deploy latest commit.
