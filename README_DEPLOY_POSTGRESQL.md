# CUBE Management Contract — PostgreSQL

Questa versione supporta PostgreSQL online tramite variabile ambiente:

```text
DATABASE_URL
```

Se `DATABASE_URL` è presente, il programma usa PostgreSQL.
Se non è presente, usa SQLite locale in `data/cube_contracts_pro.db`.

## Deploy Render con PostgreSQL

Il file `render.yaml` crea:

- servizio web Docker
- database PostgreSQL `cube-postgres`
- variabile `DATABASE_URL` collegata automaticamente

## Avvio locale con SQLite

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Avvio locale con PostgreSQL

```bash
set DATABASE_URL=postgresql://user:password@host:5432/dbname
streamlit run app.py
```

## Credenziali iniziali

```text
username: admin
password: admin123
```

Dopo il primo accesso cambia la password da Admin / Staff.

## Nota dati

Il passaggio a PostgreSQL crea un nuovo database vuoto. Se hai dati nel vecchio SQLite, serve una procedura di migrazione dati separata.
