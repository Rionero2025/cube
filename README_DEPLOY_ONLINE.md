# CUBE Management Contract — Versione Online

Questa versione è pronta per essere pubblicata online.

## Avvio locale

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy online

Il pacchetto include:

- Dockerfile
- render.yaml
- Procfile
- .streamlit/config.toml
- start.sh

## Deploy consigliato su Render

1. Carica questi file in un repository GitHub.
2. Su Render crea un nuovo Web Service.
3. Collega il repository.
4. Render userà il Dockerfile.
5. Il comando di avvio è:
   `streamlit run app.py --server.address=0.0.0.0 --server.port=$PORT`

## Database

La versione attuale usa SQLite:

```text
data/cube_contracts_pro.db
```

Il file `render.yaml` include un disco persistente su `/app/data`.

## Credenziali iniziali

```text
username: admin
password: admin123
```

Cambia la password da Admin / Staff dopo il primo accesso.
