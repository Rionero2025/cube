
import os
import textwrap
import re
import json
import hmac
import hashlib
import base64
import sqlite3
from pathlib import Path
from datetime import date, datetime, timedelta
from typing import Any

import pandas as pd
import streamlit as st

try:
    import psycopg2
except Exception:
    psycopg2 = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.pdfgen import canvas
except Exception:
    canvas = None
    A4 = None
    cm = 28.35


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = Path(os.getenv("UPLOAD_DIR", str(BASE_DIR / "uploads")))
DATA_DIR.mkdir(exist_ok=True)
UPLOAD_DIR.mkdir(exist_ok=True)

DB_PATH = DATA_DIR / "cube_saas.db"
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()
IS_POSTGRES = bool(DATABASE_URL)

APP_NAME = "CUBE Management Contract SaaS"
DEFAULT_PRIMARY = "#0f6dd0"

ROLE_SUPER_ADMIN = "Super Admin SaaS"
ROLE_ADMIN = "Admin Azienda"
ROLE_MANAGER = "Manager Operativo"
ROLE_FINANCE = "Gestione Finanziaria"
ROLE_ADVANCED = "Operativo Avanzato"
ROLE_BASE = "Operativo Base"

ROLES = [ROLE_ADMIN, ROLE_MANAGER, ROLE_FINANCE, ROLE_ADVANCED, ROLE_BASE]
TENANT_STATUS = ["Attivo", "Sospeso", "In prova", "Scaduto", "Disattivato"]
CONTRACT_STATUS = ["Bozza", "Inviato", "Firmato", "Attivo", "Sospeso", "Scaduto", "Archiviato"]
PAYMENT_STATUS = ["Da pagare", "Acconto", "Pagata", "Scaduta", "Sollecitata", "Annullata"]


# ============================================================
# BASIC HELPERS
# ============================================================

def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")

def today_iso() -> str:
    return date.today().isoformat()

def safe(v: Any) -> str:
    import html
    return html.escape("" if v is None else str(v))

def money(v: Any) -> str:
    try:
        n = float(v or 0)
    except Exception:
        n = 0.0
    s = f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"€ {s}"

def parse_float(v: Any, default: float = 0.0) -> float:
    try:
        if isinstance(v, str):
            v = v.replace("€", "").replace(" ", "").replace(".", "").replace(",", ".")
        return float(v)
    except Exception:
        return default

def add_months(d: date, months: int) -> date:
    m = d.month - 1 + int(months)
    y = d.year + m // 12
    m = m % 12 + 1
    days = [31, 29 if y % 4 == 0 and (y % 100 != 0 or y % 400 == 0) else 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    return date(y, m, min(d.day, days[m - 1]))

def slug_filename(name: str) -> str:
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name or "file")
    return name[:160] or "file"

def tenant_upload_dir(tenant_id: int | None) -> Path:
    p = UPLOAD_DIR / ("global" if tenant_id is None else f"tenant_{tenant_id}")
    p.mkdir(parents=True, exist_ok=True)
    return p

def save_upload(upload, tenant_id: int | None, prefix: str = "") -> str | None:
    if not upload:
        return None
    folder = tenant_upload_dir(tenant_id)
    filename = f"{prefix}{datetime.now().strftime('%Y%m%d_%H%M%S')}_{slug_filename(upload.name)}"
    path = folder / filename
    path.write_bytes(upload.getbuffer())
    try:
        return str(path.relative_to(BASE_DIR))
    except Exception:
        return str(path)

def hash_password(password: str) -> str:
    raw = ("CUBE_SAAS|" + str(password or "")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()

def verify_password(password: str, password_hash: str | None) -> bool:
    if not password_hash:
        return False
    return hmac.compare_digest(hash_password(password), str(password_hash))

def logo_as_base64(path: str | None) -> str:
    if not path:
        return ""
    p = Path(path)
    if not p.is_absolute():
        p = BASE_DIR / p
    if not p.exists():
        return ""
    try:
        return base64.b64encode(p.read_bytes()).decode("utf-8")
    except Exception:
        return ""


# ============================================================
# DATABASE LAYER: SQLITE + POSTGRESQL
# ============================================================

def sql_params(q: str) -> str:
    return q.replace("?", "%s") if IS_POSTGRES else q

def conn():
    if IS_POSTGRES:
        if psycopg2 is None:
            raise RuntimeError("psycopg2-binary non installato.")
        return psycopg2.connect(DATABASE_URL)
    c = sqlite3.connect(DB_PATH)
    c.row_factory = sqlite3.Row
    return c

def execute(q: str, params: tuple = ()) -> int:
    if IS_POSTGRES:
        q2 = sql_params(q.strip().rstrip(";"))
        is_insert = q2.lower().lstrip().startswith("insert into")
        if is_insert and " returning " not in q2.lower():
            q2 += " RETURNING id"
        with conn() as c:
            with c.cursor() as cur:
                cur.execute(q2, params)
                new_id = 0
                if is_insert:
                    row = cur.fetchone()
                    if row:
                        new_id = int(row[0])
                c.commit()
                return new_id
    with conn() as c:
        cur = c.execute(q, params)
        c.commit()
        return int(cur.lastrowid or 0)

def read_df(q: str, params: tuple = ()) -> pd.DataFrame:
    with conn() as c:
        return pd.read_sql_query(sql_params(q), c, params=params)

def db_executescript_sqlite(script: str):
    with conn() as c:
        c.executescript(script)
        c.commit()

def db_executes_postgres(statements: list[str]):
    with conn() as c:
        with c.cursor() as cur:
            for s in statements:
                cur.execute(s)
        c.commit()


# ============================================================
# SCHEMA MULTI-TENANT
# ============================================================

def init_db():
    if IS_POSTGRES:
        init_postgres()
    else:
        init_sqlite()
    seed_platform()

def init_sqlite():
    db_executescript_sqlite("""
    CREATE TABLE IF NOT EXISTS tenants (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ragione_sociale TEXT NOT NULL,
        forma_giuridica TEXT,
        partita_iva TEXT,
        codice_fiscale TEXT,
        sede_legale TEXT,
        pec TEXT,
        codice_sdi TEXT,
        iban TEXT,
        telefono TEXT,
        email TEXT,
        logo_file TEXT,
        stato_account TEXT DEFAULT 'In prova',
        piano_abbonamento TEXT DEFAULT 'Starter',
        data_registrazione TEXT NOT NULL,
        note TEXT
    );

    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER,
        username TEXT NOT NULL,
        email TEXT,
        password_hash TEXT NOT NULL,
        nome TEXT NOT NULL,
        cognome TEXT,
        ruolo TEXT NOT NULL,
        stato TEXT DEFAULT 'Attivo',
        telefono TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS subscription_plans (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        nome TEXT NOT NULL,
        prezzo_mensile REAL DEFAULT 0,
        max_utenti INTEGER DEFAULT 1,
        max_clienti INTEGER DEFAULT 50,
        max_contratti INTEGER DEFAULT 100,
        funzioni_json TEXT,
        attivo INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS subscriptions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        piano_id INTEGER,
        stato TEXT DEFAULT 'Attivo',
        data_inizio TEXT,
        data_scadenza TEXT,
        metodo_pagamento TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS companies (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        nome TEXT NOT NULL,
        forma_giuridica TEXT,
        piva TEXT,
        cf TEXT,
        sede TEXT,
        pec TEXT,
        codice_sdi TEXT,
        iban TEXT,
        telefono TEXT,
        email TEXT,
        logo_file TEXT,
        note TEXT,
        is_default INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        ragione_sociale TEXT NOT NULL,
        forma_giuridica TEXT,
        partita_iva TEXT,
        codice_fiscale TEXT,
        rea TEXT,
        sede_legale TEXT,
        pec TEXT,
        codice_sdi TEXT,
        legale_rappresentante TEXT,
        telefono TEXT,
        email TEXT,
        codice_ateco TEXT,
        settore TEXT,
        stato_crm TEXT DEFAULT 'Attivo',
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS client_assignments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS contract_templates (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER,
        nome TEXT NOT NULL,
        descrizione TEXT,
        testo_base TEXT NOT NULL,
        attivo INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS contracts (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        company_id INTEGER,
        client_id INTEGER NOT NULL,
        template_id INTEGER,
        assigned_user_id INTEGER,
        titolo TEXT NOT NULL,
        tipo_contratto TEXT,
        data_firma TEXT,
        luogo_firma TEXT DEFAULT 'Napoli',
        data_decorrenza TEXT NOT NULL,
        data_scadenza TEXT NOT NULL,
        durata_mesi INTEGER NOT NULL,
        importo_totale REAL NOT NULL,
        iva_percentuale REAL NOT NULL,
        modalita_pagamento TEXT NOT NULL,
        foro_competente TEXT DEFAULT 'Napoli',
        stato TEXT NOT NULL,
        servizi_json TEXT,
        clausole_extra TEXT,
        note TEXT,
        file_docx TEXT,
        file_pdf TEXT,
        file_firmato TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS payments (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        contract_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        numero_rata INTEGER NOT NULL,
        data_scadenza TEXT NOT NULL,
        imponibile REAL NOT NULL,
        iva REAL NOT NULL,
        totale REAL NOT NULL,
        stato TEXT NOT NULL DEFAULT 'Da pagare',
        data_pagamento TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS payment_movements (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        payment_id INTEGER NOT NULL,
        importo_pagato REAL NOT NULL,
        tipo_movimento TEXT NOT NULL DEFAULT 'Acconto',
        data_pagamento TEXT NOT NULL,
        allegato_file TEXT,
        note TEXT,
        registrato_da_user_id INTEGER,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS work_logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        user_id INTEGER,
        data_lavoro TEXT NOT NULL,
        ora_lavoro TEXT,
        tipo_lavoro TEXT,
        titolo TEXT NOT NULL,
        descrizione TEXT,
        stato TEXT,
        allegato_file TEXT,
        note_interne TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS documents (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER,
        contract_id INTEGER,
        work_id INTEGER,
        tipo TEXT,
        titolo TEXT NOT NULL,
        file_path TEXT,
        note TEXT,
        uploaded_by_user_id INTEGER,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS feedback (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        user_id INTEGER,
        data_feedback TEXT NOT NULL,
        provenienza TEXT DEFAULT 'Cliente',
        valutazione INTEGER,
        testo_feedback TEXT,
        allegato_file TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    );

    CREATE TABLE IF NOT EXISTS invoices (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        tenant_id INTEGER NOT NULL,
        company_id INTEGER,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        payment_id INTEGER,
        work_id INTEGER,
        numero TEXT NOT NULL,
        anno INTEGER NOT NULL,
        data_fattura TEXT NOT NULL,
        scadenza TEXT,
        descrizione TEXT,
        imponibile REAL NOT NULL,
        iva_percentuale REAL NOT NULL,
        iva REAL NOT NULL,
        totale REAL NOT NULL,
        stato TEXT DEFAULT 'Bozza',
        file_pdf TEXT,
        note TEXT,
        emessa_da_user_id INTEGER,
        created_at TEXT NOT NULL
    );
    """)

def init_postgres():
    statements = []
    # Same schema with SERIAL
    sqlite_script = open(__file__, "r", encoding="utf-8").read()
    # explicit statements are safer
    table_sql = [
    """CREATE TABLE IF NOT EXISTS tenants (
        id SERIAL PRIMARY KEY,
        ragione_sociale TEXT NOT NULL,
        forma_giuridica TEXT,
        partita_iva TEXT,
        codice_fiscale TEXT,
        sede_legale TEXT,
        pec TEXT,
        codice_sdi TEXT,
        iban TEXT,
        telefono TEXT,
        email TEXT,
        logo_file TEXT,
        stato_account TEXT DEFAULT 'In prova',
        piano_abbonamento TEXT DEFAULT 'Starter',
        data_registrazione TEXT NOT NULL,
        note TEXT
    )""",
    """CREATE TABLE IF NOT EXISTS users (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER,
        username TEXT NOT NULL,
        email TEXT,
        password_hash TEXT NOT NULL,
        nome TEXT NOT NULL,
        cognome TEXT,
        ruolo TEXT NOT NULL,
        stato TEXT DEFAULT 'Attivo',
        telefono TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS subscription_plans (
        id SERIAL PRIMARY KEY,
        nome TEXT NOT NULL,
        prezzo_mensile REAL DEFAULT 0,
        max_utenti INTEGER DEFAULT 1,
        max_clienti INTEGER DEFAULT 50,
        max_contratti INTEGER DEFAULT 100,
        funzioni_json TEXT,
        attivo INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS subscriptions (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        piano_id INTEGER,
        stato TEXT DEFAULT 'Attivo',
        data_inizio TEXT,
        data_scadenza TEXT,
        metodo_pagamento TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS companies (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        nome TEXT NOT NULL,
        forma_giuridica TEXT,
        piva TEXT,
        cf TEXT,
        sede TEXT,
        pec TEXT,
        codice_sdi TEXT,
        iban TEXT,
        telefono TEXT,
        email TEXT,
        logo_file TEXT,
        note TEXT,
        is_default INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS clients (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        ragione_sociale TEXT NOT NULL,
        forma_giuridica TEXT,
        partita_iva TEXT,
        codice_fiscale TEXT,
        rea TEXT,
        sede_legale TEXT,
        pec TEXT,
        codice_sdi TEXT,
        legale_rappresentante TEXT,
        telefono TEXT,
        email TEXT,
        codice_ateco TEXT,
        settore TEXT,
        stato_crm TEXT DEFAULT 'Attivo',
        note TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS client_assignments (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        user_id INTEGER NOT NULL,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS contract_templates (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER,
        nome TEXT NOT NULL,
        descrizione TEXT,
        testo_base TEXT NOT NULL,
        attivo INTEGER DEFAULT 1,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS contracts (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        company_id INTEGER,
        client_id INTEGER NOT NULL,
        template_id INTEGER,
        assigned_user_id INTEGER,
        titolo TEXT NOT NULL,
        tipo_contratto TEXT,
        data_firma TEXT,
        luogo_firma TEXT DEFAULT 'Napoli',
        data_decorrenza TEXT NOT NULL,
        data_scadenza TEXT NOT NULL,
        durata_mesi INTEGER NOT NULL,
        importo_totale REAL NOT NULL,
        iva_percentuale REAL NOT NULL,
        modalita_pagamento TEXT NOT NULL,
        foro_competente TEXT DEFAULT 'Napoli',
        stato TEXT NOT NULL,
        servizi_json TEXT,
        clausole_extra TEXT,
        note TEXT,
        file_docx TEXT,
        file_pdf TEXT,
        file_firmato TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS payments (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        contract_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        numero_rata INTEGER NOT NULL,
        data_scadenza TEXT NOT NULL,
        imponibile REAL NOT NULL,
        iva REAL NOT NULL,
        totale REAL NOT NULL,
        stato TEXT NOT NULL DEFAULT 'Da pagare',
        data_pagamento TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS payment_movements (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        payment_id INTEGER NOT NULL,
        importo_pagato REAL NOT NULL,
        tipo_movimento TEXT NOT NULL DEFAULT 'Acconto',
        data_pagamento TEXT NOT NULL,
        allegato_file TEXT,
        note TEXT,
        registrato_da_user_id INTEGER,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS work_logs (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        user_id INTEGER,
        data_lavoro TEXT NOT NULL,
        ora_lavoro TEXT,
        tipo_lavoro TEXT,
        titolo TEXT NOT NULL,
        descrizione TEXT,
        stato TEXT,
        allegato_file TEXT,
        note_interne TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS documents (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER,
        contract_id INTEGER,
        work_id INTEGER,
        tipo TEXT,
        titolo TEXT NOT NULL,
        file_path TEXT,
        note TEXT,
        uploaded_by_user_id INTEGER,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS feedback (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        user_id INTEGER,
        data_feedback TEXT NOT NULL,
        provenienza TEXT DEFAULT 'Cliente',
        valutazione INTEGER,
        testo_feedback TEXT,
        allegato_file TEXT,
        note TEXT,
        created_at TEXT NOT NULL
    )""",
    """CREATE TABLE IF NOT EXISTS invoices (
        id SERIAL PRIMARY KEY,
        tenant_id INTEGER NOT NULL,
        company_id INTEGER,
        client_id INTEGER NOT NULL,
        contract_id INTEGER,
        payment_id INTEGER,
        work_id INTEGER,
        numero TEXT NOT NULL,
        anno INTEGER NOT NULL,
        data_fattura TEXT NOT NULL,
        scadenza TEXT,
        descrizione TEXT,
        imponibile REAL NOT NULL,
        iva_percentuale REAL NOT NULL,
        iva REAL NOT NULL,
        totale REAL NOT NULL,
        stato TEXT DEFAULT 'Bozza',
        file_pdf TEXT,
        note TEXT,
        emessa_da_user_id INTEGER,
        created_at TEXT NOT NULL
    )"""
    ]
    db_executes_postgres(table_sql)

def seed_platform():
    # Super admin
    admins = read_df("SELECT id FROM users WHERE ruolo=?", (ROLE_SUPER_ADMIN,))
    if admins.empty:
        execute(
            "INSERT INTO users (tenant_id,username,email,password_hash,nome,cognome,ruolo,stato,created_at) VALUES (?,?,?,?,?,?,?,?,?)",
            (None, "superadmin", "admin@cube.local", hash_password("admin123"), "Super", "Admin", ROLE_SUPER_ADMIN, "Attivo", now_iso())
        )

    # Plans: sincronizzazione piani ufficiali CUBE.
    # Viene eseguita anche se il database contiene vecchi piani, così aggiorna prezzi/limiti online.
    official_plans = [
        ("Free", 0.0, 1, 3, 3, ["CRM base", "3 contratti gestibili", "1 utente admin", "Prova del gestionale"]),
        ("Starter", 9.0, 2, 10, 10, ["CRM", "Contratti", "Documenti", "1 staff aggiuntivo"]),
        ("Professional", 29.0, 4, 30, 30, ["CRM", "Contratti", "Pagamenti", "Fatture", "3 staff aggiuntivi"]),
        ("Business", 49.0, 11, 100, 100, ["CRM", "Contratti", "Pagamenti", "Fatture", "10 staff aggiuntivi"]),
        ("Enterprise", 99.0, 999999, 999999, 999999, ["Tutto illimitato", "Supporto", "Personalizzazioni", "Staff illimitato"]),
    ]

    # Disattiva eventuali vecchi piani non ufficiali.
    official_names = tuple([p[0] for p in official_plans])
    try:
        existing = read_df("SELECT id,nome FROM subscription_plans")
        for _, oldp in existing.iterrows():
            if str(oldp["nome"]) not in official_names:
                execute("UPDATE subscription_plans SET attivo=0 WHERE id=?", (int(oldp["id"]),))
    except Exception:
        pass

    for nome, prezzo, utenti, clienti, contratti, funzioni in official_plans:
        found = read_df("SELECT id FROM subscription_plans WHERE nome=?", (nome,))
        if found.empty:
            execute(
                "INSERT INTO subscription_plans (nome,prezzo_mensile,max_utenti,max_clienti,max_contratti,funzioni_json,attivo,created_at) VALUES (?,?,?,?,?,?,?,?)",
                (nome, prezzo, utenti, clienti, contratti, json.dumps(funzioni), 1, now_iso())
            )
        else:
            execute(
                "UPDATE subscription_plans SET prezzo_mensile=?, max_utenti=?, max_clienti=?, max_contratti=?, funzioni_json=?, attivo=1 WHERE nome=?",
                (prezzo, utenti, clienti, contratti, json.dumps(funzioni), nome)
            )

    # Global contract template
    if read_df("SELECT id FROM contract_templates WHERE tenant_id IS NULL LIMIT 1").empty:
        execute(
            "INSERT INTO contract_templates (tenant_id,nome,descrizione,testo_base,attivo,created_at) VALUES (?,?,?,?,?,?)",
            (None, "Template consulenza strategica e operativa", "Template globale disponibile per tutte le aziende.", DEFAULT_TEMPLATE, 1, now_iso())
        )


# ============================================================
# SESSION / PERMISSIONS
# ============================================================

def current_user() -> dict | None:
    uid = st.session_state.get("user_id")
    if not uid:
        return None
    df = read_df("SELECT * FROM users WHERE id=?", (int(uid),))
    if df.empty:
        st.session_state.clear()
        return None
    return df.iloc[0].to_dict()

def current_tenant_id() -> int | None:
    u = current_user()
    if not u:
        return None
    return None if pd.isna(u.get("tenant_id")) else int(u["tenant_id"])

def current_tenant() -> dict | None:
    tid = current_tenant_id()
    if tid is None:
        return None
    df = read_df("SELECT * FROM tenants WHERE id=?", (tid,))
    return None if df.empty else df.iloc[0].to_dict()

def role() -> str:
    u = current_user()
    return "" if not u else str(u.get("ruolo") or "")

def is_super_admin() -> bool:
    return role() == ROLE_SUPER_ADMIN

def is_tenant_admin() -> bool:
    return role() == ROLE_ADMIN

def can_finance() -> bool:
    return role() in [ROLE_SUPER_ADMIN, ROLE_ADMIN, ROLE_FINANCE]

def can_manage_users() -> bool:
    return role() in [ROLE_SUPER_ADMIN, ROLE_ADMIN]

def can_manage_contracts() -> bool:
    return role() in [ROLE_SUPER_ADMIN, ROLE_ADMIN, ROLE_MANAGER]

def can_see_all_clients() -> bool:
    return role() in [ROLE_SUPER_ADMIN, ROLE_ADMIN, ROLE_MANAGER, ROLE_FINANCE, ROLE_ADVANCED]

def tenant_filter_clause(alias: str = "") -> tuple[str, tuple]:
    tid = current_tenant_id()
    prefix = f"{alias}." if alias else ""
    if is_super_admin():
        return "1=1", ()
    return f"{prefix}tenant_id=?", (tid,)

def visible_clients_df() -> pd.DataFrame:
    tid = current_tenant_id()
    if is_super_admin():
        return read_df("SELECT c.*, t.ragione_sociale tenant FROM clients c JOIN tenants t ON t.id=c.tenant_id ORDER BY c.id DESC")
    if can_see_all_clients():
        return read_df("SELECT * FROM clients WHERE tenant_id=? ORDER BY ragione_sociale", (tid,))
    # Base: solo clienti assegnati.
    return read_df("""
        SELECT DISTINCT c.*
        FROM clients c
        JOIN client_assignments a ON a.client_id=c.id
        WHERE c.tenant_id=? AND a.user_id=?
        ORDER BY c.ragione_sociale
    """, (tid, int(current_user()["id"])))

def users_for_tenant(active_only=True) -> pd.DataFrame:
    tid = current_tenant_id()
    if tid is None:
        return read_df("SELECT * FROM users ORDER BY id DESC")
    q = "SELECT * FROM users WHERE tenant_id=?"
    params = [tid]
    if active_only:
        q += " AND stato='Attivo'"
    q += " ORDER BY nome,cognome"
    return read_df(q, tuple(params))


# ============================================================
# UI
# ============================================================

def css():
    st.markdown(f"""
    <style>
    :root {{
      --cube-blue:#0f6dd0;
      --cube-blue2:#0b5bb3;
      --cube-navy:#061b3a;
      --cube-dark:#071527;
      --cube-text:#11243d;
      --cube-muted:#60728c;
      --cube-bg:#f4f8fd;
      --cube-soft:#eaf3ff;
      --cube-border:#dce8f5;
      --cube-shadow:0 18px 45px rgba(8, 34, 73, .10);
      --cube-radius:24px;
    }}

    .stApp {{
      background:
        radial-gradient(circle at 12% 6%, rgba(15,109,208,.13), transparent 28%),
        radial-gradient(circle at 88% 10%, rgba(15,109,208,.10), transparent 26%),
        linear-gradient(180deg,#ffffff 0%,#f4f8fd 34%,#eef5fc 100%);
      color:var(--cube-text);
    }}

    .block-container {{
      padding-top: 1.1rem !important;
      max-width: 1440px !important;
      padding-left: 2.4rem !important;
      padding-right: 2.4rem !important;
    }}

    [data-testid="stSidebar"] {{
      background:linear-gradient(180deg,#071527 0%,#0d2946 100%);
      border-right:1px solid rgba(255,255,255,.08);
    }}
    [data-testid="stSidebar"] * {{ color:white; }}

    /* Public navbar */
    .cube-topbar {{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:22px;
      min-height:72px;
      padding:10px 0 18px 0;
      margin-bottom:4px;
    }}
    .cube-brand {{
      display:flex;
      align-items:center;
      gap:12px;
      color:var(--cube-navy);
      font-weight:950;
      letter-spacing:-.02em;
      text-decoration:none;
    }}
    .cube-logo-box {{
      width:50px;
      height:50px;
      border-radius:14px;
      display:grid;
      place-items:center;
      background:linear-gradient(135deg,#f0f7ff,#dbeaff);
      border:1px solid #d2e3f7;
    }}
    .cube-logo-mark {{
      width:36px;
      height:42px;
      border-radius:14px;
      display:grid;
      place-items:center;
      color:white;
      background:linear-gradient(135deg,#1b8cff,#084aa5);
      box-shadow:0 12px 30px rgba(15,109,208,.26);
      font-size:1.1rem;
    }}
    .cube-brand-copy strong {{
      display:block;
      font-size:2rem;
      line-height:.9;
      letter-spacing:-.06em;
      color:#0a1f45;
    }}
    .cube-brand-copy small {{
      display:block;
      color:#394f6d;
      font-size:.9rem;
      margin-top:2px;
      font-weight:800;
      line-height:1.05;
    }}
    .cube-nav {{
      display:flex;
      gap:28px;
      align-items:center;
      color:#40536d;
      font-weight:850;
      font-size:.94rem;
      margin-left:auto;
      margin-right:10px;
    }}
    .cube-nav a {{ text-decoration:none; color:#40536d; }}
    .cube-nav span {{
      cursor:default;
    }}
    .cube-top-actions {{
      display:flex;
      gap:10px;
      align-items:center;
    }}

    /* Buttons */
    div.stButton > button,
    div.stDownloadButton > button,
    div[data-testid="stFormSubmitButton"] button {{
      border-radius:13px !important;
      min-height:42px !important;
      font-weight:900 !important;
      border:1px solid rgba(15,109,208,.18) !important;
      box-shadow:0 9px 22px rgba(15,109,208,.13) !important;
      transition:all .18s ease !important;
    }}
    div.stButton > button:hover,
    div.stDownloadButton > button:hover,
    div[data-testid="stFormSubmitButton"] button:hover {{
      transform:translateY(-1px);
      box-shadow:0 13px 28px rgba(15,109,208,.20) !important;
    }}

    /* Premium public landing */
    .premium-hero {{
      position:relative;
      overflow:hidden;
      border-radius:34px;
      padding:42px 38px 34px;
      background:
        radial-gradient(circle at 90% 12%, rgba(15,109,208,.17), transparent 30%),
        linear-gradient(135deg,#ffffff 0%,#edf6ff 52%,#e8f3ff 100%);
      border:1px solid var(--cube-border);
      box-shadow:var(--cube-shadow);
      margin-bottom:24px;
    }}
    .premium-hero-grid {{
      display:grid;
      grid-template-columns: .96fr 1.04fr;
      gap:44px;
      align-items:center;
    }}
    .premium-badge {{
      display:inline-flex;
      align-items:center;
      gap:8px;
      background:#eaf3ff;
      color:#0f53a5;
      border:1px solid #cfe2fa;
      border-radius:999px;
      padding:7px 12px;
      font-size:.83rem;
      font-weight:950;
      margin-bottom:18px;
    }}
    .premium-hero h1 {{
      font-size:clamp(2.4rem,4.1vw,4.8rem);
      line-height:.98;
      margin:0 0 18px;
      letter-spacing:-.055em;
      color:var(--cube-navy);
    }}
    .premium-hero h1 span {{
      color:var(--cube-blue);
    }}
    .premium-hero p {{
      color:#546982;
      font-size:1.08rem;
      line-height:1.72;
      max-width:700px;
      margin:0 0 24px;
    }}
    .premium-actions {{
      display:flex;
      gap:14px;
      flex-wrap:wrap;
      margin-bottom:22px;
    }}
    .html-btn {{
      display:inline-flex;
      align-items:center;
      justify-content:center;
      gap:9px;
      border-radius:14px;
      padding:14px 22px;
      font-weight:950;
      text-decoration:none;
      border:1px solid transparent;
    }}
    .html-btn.primary {{
      color:white;
      background:linear-gradient(135deg,#0f6dd0,#0b58b8);
      box-shadow:0 18px 38px rgba(15,109,208,.28);
    }}
    .html-btn.secondary {{
      color:#0f53a5;
      background:white;
      border-color:#b8d1ef;
      box-shadow:0 10px 25px rgba(8,34,73,.06);
    }}
    .html-btn.compact {{ padding:11px 20px; min-width:120px; }}

    .hero-micro {{
      display:flex;
      gap:22px;
      flex-wrap:wrap;
      color:#41556f;
      font-size:.88rem;
      font-weight:850;
    }}
    .hero-micro span {{
      display:inline-flex;
      gap:7px;
      align-items:center;
    }}

    .mockup {{
      position:relative;
      min-height:470px;
      padding-left:20px;
    }}
    .mock-browser {{
      background:white;
      border:1px solid #dce8f5;
      border-radius:22px;
      box-shadow:0 22px 55px rgba(8,34,73,.14);
      padding:16px;
      overflow:hidden;
    }}
    .browser-dots {{
      display:flex;
      gap:6px;
      margin-bottom:12px;
    }}
    .browser-dots span {{
      width:9px;
      height:9px;
      border-radius:999px;
      background:#d9e6f4;
    }}
    .mock-shell {{
      display:grid;
      grid-template-columns:74px 1fr;
      gap:14px;
    }}
    .mock-side {{
      background:linear-gradient(180deg,#08244a,#06336f);
      border-radius:18px;
      min-height:285px;
      padding:14px 9px;
    }}
    .mock-side i {{
      display:block;
      width:38px;
      height:8px;
      border-radius:999px;
      background:rgba(255,255,255,.26);
      margin:13px auto;
    }}
    .mock-side i:first-child {{
      height:34px;
      width:34px;
      background:#fff;
      opacity:.95;
    }}
    .mock-main h4 {{
      margin:4px 0 12px;
      color:var(--cube-navy);
    }}
    .mock-stats {{
      display:grid;
      grid-template-columns:repeat(4,1fr);
      gap:10px;
      margin-bottom:12px;
    }}
    .mock-stat {{
      border:1px solid #e2edf8;
      border-radius:14px;
      padding:10px;
      background:#fbfdff;
    }}
    .mock-stat small {{
      display:block;
      color:#6b7e96;
      font-weight:850;
      font-size:.68rem;
      margin-bottom:5px;
    }}
    .mock-stat strong {{
      color:#061b3a;
      font-size:1.05rem;
    }}
    .mock-panels {{
      display:grid;
      grid-template-columns:1.25fr .75fr;
      gap:12px;
    }}
    .mock-panel {{
      border:1px solid #e2edf8;
      background:#fbfdff;
      border-radius:16px;
      padding:12px;
      min-height:128px;
    }}
    .chart-line {{
      height:92px;
      border-radius:14px;
      background:
        linear-gradient(135deg, transparent 0 32%, rgba(15,109,208,.16) 32% 34%, transparent 34% 55%, rgba(15,109,208,.20) 55% 57%, transparent 57%),
        repeating-linear-gradient(0deg,#eef5fc 0 1px,transparent 1px 22px);
    }}
    .mock-list p {{
      display:flex;
      justify-content:space-between;
      gap:8px;
      font-size:.72rem;
      color:#536982;
      margin:8px 0;
    }}
    .float-card {{
      position:absolute;
      background:white;
      border:1px solid #dce8f5;
      border-radius:18px;
      box-shadow:0 18px 42px rgba(8,34,73,.12);
      padding:14px;
    }}
    .float-card.contracts {{
      left:-26px;
      bottom:5px;
      width:245px;
    }}
    .float-card.payments {{
      right:-24px;
      bottom:25px;
      width:230px;
    }}
    .donut {{
      width:96px;
      height:96px;
      border-radius:999px;
      margin:4px auto 8px;
      background:conic-gradient(#0f6dd0 0 62%,#20bf78 62% 82%,#ffb020 82% 100%);
      display:grid;
      place-items:center;
    }}
    .donut div {{
      width:62px;
      height:62px;
      border-radius:999px;
      background:white;
      display:grid;
      place-items:center;
      font-weight:950;
      color:var(--cube-navy);
      font-size:.82rem;
    }}

    .benefit-strip {{
      display:grid;
      grid-template-columns:repeat(4,1fr);
      gap:14px;
      padding:14px;
      border:1px solid var(--cube-border);
      border-radius:26px;
      background:rgba(255,255,255,.78);
      box-shadow:0 12px 28px rgba(8,34,73,.06);
      margin:22px 0;
    }}
    .benefit-item {{
      display:flex;
      gap:13px;
      align-items:center;
    }}
    .benefit-icon {{
      width:48px;
      height:48px;
      border-radius:16px;
      background:#eaf3ff;
      color:#0f6dd0;
      display:grid;
      place-items:center;
      font-size:1.45rem;
      flex:0 0 auto;
    }}
    .benefit-item b {{
      display:block;
      color:var(--cube-navy);
      margin-bottom:2px;
    }}
    .benefit-item small {{
      color:#60728c;
      line-height:1.35;
      display:block;
    }}

    .premium-section {{
      margin:30px 0;
    }}
    .premium-title {{
      display:flex;
      align-items:center;
      gap:14px;
      margin:26px 0 16px;
    }}
    .premium-title .iconbox {{
      width:48px;
      height:48px;
      border-radius:16px;
      display:grid;
      place-items:center;
      color:white;
      background:linear-gradient(135deg,#0f6dd0,#073b86);
      box-shadow:0 12px 28px rgba(15,109,208,.22);
    }}
    .premium-title h2 {{
      color:var(--cube-navy);
      font-size:1.75rem;
      margin:0;
      line-height:1.05;
      letter-spacing:-.03em;
    }}
    .premium-title p {{
      margin:4px 0 0;
      color:#60728c;
    }}

    .feature-grid {{
      display:grid;
      grid-template-columns:repeat(4,1fr);
      gap:16px;
    }}
    .feature-card-premium {{
      background:white;
      border:1px solid var(--cube-border);
      border-radius:22px;
      padding:20px;
      box-shadow:0 12px 30px rgba(8,34,73,.06);
      min-height:152px;
    }}
    .feature-card-premium .ficon {{
      width:48px;
      height:48px;
      border-radius:16px;
      background:#eef6ff;
      display:grid;
      place-items:center;
      color:#0f6dd0;
      font-size:1.45rem;
      margin-bottom:12px;
    }}
    .feature-card-premium b {{
      color:var(--cube-navy);
      display:block;
      margin-bottom:7px;
    }}
    .feature-card-premium p {{
      color:#60728c;
      line-height:1.45;
      font-size:.9rem;
      margin:0;
    }}

    .steps-row {{
      display:grid;
      grid-template-columns:repeat(4,1fr);
      gap:18px;
      align-items:stretch;
    }}
    .step-card {{
      background:white;
      border:1px solid var(--cube-border);
      border-radius:22px;
      padding:18px;
      box-shadow:0 12px 28px rgba(8,34,73,.06);
      position:relative;
    }}
    .step-num {{
      width:30px;
      height:30px;
      border-radius:999px;
      background:#0f6dd0;
      color:white;
      display:grid;
      place-items:center;
      font-weight:950;
      margin-bottom:10px;
    }}
    .step-card b {{
      color:var(--cube-navy);
    }}
    .step-card p {{
      color:#60728c;
      line-height:1.45;
      margin:6px 0 0;
      font-size:.9rem;
    }}

    .pricing-grid {{
      display:grid;
      grid-template-columns:repeat(5,1fr);
      gap:18px;
      align-items:stretch;
    }}
    .pricing-card {{
      position:relative;
      background:white;
      border:1px solid var(--cube-border);
      border-radius:24px;
      padding:22px 18px;
      box-shadow:0 12px 30px rgba(8,34,73,.07);
      text-align:center;
      min-height:346px;
      display:flex;
      flex-direction:column;
    }}
    .pricing-card.featured {{
      border:2px solid #0f6dd0;
      box-shadow:0 18px 44px rgba(15,109,208,.17);
      transform:translateY(-12px);
    }}
    .popular-ribbon {{
      position:absolute;
      top:-13px;
      left:50%;
      transform:translateX(-50%);
      background:#0f6dd0;
      color:white;
      border-radius:999px;
      padding:5px 14px;
      font-size:.74rem;
      font-weight:950;
      white-space:nowrap;
    }}
    .pricing-icon {{
      width:54px;
      height:54px;
      border-radius:18px;
      background:#eef6ff;
      color:#0f6dd0;
      display:grid;
      place-items:center;
      font-size:1.55rem;
      margin:0 auto 10px;
    }}
    .pricing-card h3 {{
      margin:0 0 8px;
      color:var(--cube-navy);
      font-size:1.15rem;
    }}
    .price {{
      font-size:2.05rem;
      color:#061b3a;
      font-weight:950;
      letter-spacing:-.04em;
      margin:8px 0 2px;
    }}
    .price small {{
      font-size:.78rem;
      color:#60728c;
      font-weight:800;
    }}
    .pricing-card p {{
      color:#60728c;
      font-size:.88rem;
      line-height:1.42;
      margin:6px 0 14px;
    }}
    .pricing-card ul {{
      list-style:none;
      padding:0;
      margin:0 0 16px;
      color:#203653;
      font-size:.87rem;
      line-height:1.7;
      text-align:left;
    }}
    .pricing-card li::before {{
      content:"✓";
      color:#0f6dd0;
      font-weight:950;
      margin-right:7px;
    }}
    .pricing-card .html-btn {{
      margin-top:auto;
      padding:10px 12px;
      width:100%;
    }}

    .final-cta {{
      margin:30px 0 0;
      border-radius:28px;
      background:
        radial-gradient(circle at 8% 50%, rgba(255,255,255,.22), transparent 18%),
        linear-gradient(135deg,#0f6dd0,#0742a0 70%,#062a66);
      color:white;
      padding:28px 34px;
      display:flex;
      justify-content:space-between;
      align-items:center;
      gap:24px;
      box-shadow:0 18px 44px rgba(15,109,208,.24);
    }}
    .final-cta h2 {{
      margin:0 0 6px;
      font-size:1.85rem;
      letter-spacing:-.035em;
    }}
    .final-cta p {{
      margin:0;
      color:#dcecff;
    }}

    .premium-footer {{
      background:#061b3a;
      color:#c9d7ea;
      border-radius:28px 28px 0 0;
      padding:34px;
      margin-top:24px;
      display:grid;
      grid-template-columns:1.4fr repeat(4,1fr);
      gap:24px;
    }}
    .premium-footer b {{
      color:white;
      display:block;
      margin-bottom:10px;
    }}
    .premium-footer p, .premium-footer small {{
      color:#c9d7ea;
      line-height:1.55;
    }}
    .premium-footer a {{
      display:block;
      color:#d8e7fa;
      text-decoration:none;
      margin:7px 0;
      font-size:.9rem;
    }}

    .card {{
      background:white;
      border:1px solid var(--cube-border);
      border-radius:20px;
      padding:18px;
      box-shadow:0 12px 30px rgba(8,34,73,.06);
    }}
    .card .label {{ color:#60728c; font-weight:900; font-size:.85rem; }}
    .card .value {{ color:var(--cube-navy); font-weight:950; font-size:1.35rem; }}

    div[data-testid="stMetric"] {{
      background:white;
      padding:16px;
      border-radius:18px;
      border:1px solid var(--cube-border);
      box-shadow:0 12px 28px rgba(8,34,73,.06);
    }}
    div[data-testid="stDataFrame"] {{
      background:white;
      border-radius:18px;
      border:1px solid var(--cube-border);
      padding:4px;
      overflow-x:auto!important;
      box-shadow:0 12px 28px rgba(8,34,73,.05);
    }}

    @media (max-width: 1200px) {{
      .pricing-grid {{ grid-template-columns:repeat(2,1fr); }}
      .feature-grid {{ grid-template-columns:repeat(2,1fr); }}
      .benefit-strip {{ grid-template-columns:repeat(2,1fr); }}
    }}

    @media (max-width: 780px) {{
      .block-container {{
        padding-left:.8rem!important;
        padding-right:.8rem!important;
      }}
      .cube-topbar {{
        flex-wrap:wrap;
      }}
      .cube-nav {{
        width:100%;
        overflow:auto;
        gap:14px;
        font-size:.85rem;
      }}
      .premium-hero {{
        padding:24px 18px;
        border-radius:24px;
      }}
      .premium-hero-grid {{
        grid-template-columns:1fr;
        gap:24px;
      }}
      .premium-hero h1 {{
        font-size:2.15rem;
      }}
      .mockup {{
        min-height:auto;
      }}
      .mock-stats, .mock-panels, .benefit-strip, .feature-grid, .steps-row, .pricing-grid, .premium-footer {{
        grid-template-columns:1fr;
      }}
      .float-card {{
        position:static;
        width:auto!important;
        margin-top:12px;
      }}
      .final-cta {{
        flex-direction:column;
        align-items:flex-start;
        padding:24px 18px;
      }}
      .html-btn {{
        width:100%;
      }}
      div[data-testid="stHorizontalBlock"] {{
        flex-direction:column!important;
      }}
      div[data-testid="stHorizontalBlock"] > div {{
        width:100%!important;
        min-width:100%!important;
        flex:1 1 100%!important;
      }}
    }}
    </style>
    """, unsafe_allow_html=True)


def section(icon, title, cap=""):
    st.markdown(f"<div class='section'><div class='ico'>{icon}</div><div><div class='ttl'>{safe(title)}</div>{('<div class=cap>'+safe(cap)+'</div>') if cap else ''}</div></div>", unsafe_allow_html=True)

def card(label, value):
    st.markdown(f"<div class='card'><div class='label'>{safe(label)}</div><div class='value'>{safe(value)}</div></div>", unsafe_allow_html=True)

def header():
    u = current_user()
    tenant = current_tenant()
    logo_file = tenant.get("logo_file") if tenant else None
    b64 = logo_as_base64(logo_file)
    logo = f"<img src='data:image/png;base64,{b64}' style='max-width:180px;height:auto;background:white;border-radius:14px;padding:10px;border:1px solid #e1e9f4;margin-bottom:10px'>" if b64 else ""
    subtitle = "Portale multi-azienda SaaS per CRM, contratti, lavori, pagamenti, fatture e staff."
    if tenant:
        subtitle = f"Area aziendale: {tenant.get('ragione_sociale')}. Dati isolati per tenant."
    st.markdown(logo + f"<div class='hero'><span class='badge'>☁️ SaaS multi-azienda</span><h1>{APP_NAME}</h1><p>{safe(subtitle)}</p></div>", unsafe_allow_html=True)

def sidebar():
    u = current_user()
    if u:
        st.sidebar.markdown(f"**👤 {safe(u.get('nome'))} {safe(u.get('cognome') or '')}**")
        st.sidebar.caption(role())
        if current_tenant():
            st.sidebar.caption(f"Azienda: {current_tenant().get('ragione_sociale')}")
        if st.sidebar.button("🚪 Esci"):
            st.session_state.clear()
            st.rerun()

    if is_super_admin():
        opts = {
            "🏠 Super Dashboard": "super_dashboard",
            "🏢 Aziende SaaS": "tenants",
            "👤 Utenti globali": "global_users",
            "💳 Piani": "plans",
            "🧩 Dati SaaS": "super_data",
            "📊 Log / Stato": "status",
        }
    else:
        opts = {"🏠 Dashboard": "dashboard", "👥 Clienti CRM": "clients", "🛠️ Lavori": "work", "📎 Documenti": "documents", "💬 Feedback": "feedback"}
        if can_manage_contracts():
            opts["📚 Contratti"] = "contracts"
            opts["📄 Template"] = "templates"
        if can_finance():
            opts["💶 Pagamenti"] = "payments"
            opts["🧾 Fatture"] = "invoices"
        if can_manage_users():
            opts["👤 Staff"] = "staff"
            opts["🏢 Dati azienda"] = "company"
    labels = list(opts.keys())
    selected = st.sidebar.radio("Menu", labels, label_visibility="collapsed")
    return opts[selected]


# ============================================================
# LOGIN + REGISTRATION
# ============================================================


# ============================================================
# PUBLIC WEBSITE + LOGIN + REGISTRATION
# ============================================================

def get_public_plans() -> pd.DataFrame:
    try:
        return read_df("""
            SELECT * FROM subscription_plans
            WHERE attivo=1
            ORDER BY CASE nome
                WHEN 'Free' THEN 1
                WHEN 'Starter' THEN 2
                WHEN 'Professional' THEN 3
                WHEN 'Business' THEN 4
                WHEN 'Enterprise' THEN 5
                ELSE 99
            END
        """)
    except Exception:
        return pd.DataFrame()

def set_public_page(page: str, plan_name: str | None = None):
    st.session_state["public_page"] = page
    if plan_name:
        st.session_state["selected_plan_name"] = plan_name

def public_topbar():
    st.markdown(textwrap.dedent("""
    <div class="cube-topbar">
      <div class="cube-brand">
        <div class="cube-logo-box">
          <div class="cube-logo-mark">⬢</div>
        </div>
        <div class="cube-brand-copy">
          <strong>CUBE</strong>
          <small>Management Contract</small>
        </div>
      </div>
      <div class="cube-nav">
        <a href="#funzionalita">Funzionalità</a>
        <a href="#come-funziona">Come funziona</a>
        <a href="#pacchetti">Prezzi</a>
        <a href="#faq">FAQ</a>
      </div>
      <div class="cube-top-actions">
        <a class="html-btn secondary compact" href="/?public_page=login">Accedi</a>
        <a class="html-btn primary compact" href="/?public_page=plans">Registrati</a>
      </div>
    </div>
    """), unsafe_allow_html=True)


def public_landing_page():
    css()
    public_topbar()

    hero_html = textwrap.dedent("""
    <div class="premium-hero">
      <div class="premium-hero-grid">
        <div>
          <div class="premium-badge">☁️ Multi-azienda · 30 giorni gratis</div>
          <h1>Il gestionale online per <span>contratti, clienti, lavori e pagamenti</span></h1>
          <p>
            CUBE Management Contract aiuta aziende, consulenti, agenzie e società di servizi a gestire
            CRM, contratti, scadenze, staff, documenti, rate, pagamenti e fatture interne in un unico
            sistema semplice e sicuro.
          </p>
          <div class="premium-actions">
            <a class="html-btn primary" href="/?public_page=plans">Prova gratis 30 giorni →</a>
            <a class="html-btn secondary" href="#funzionalita">Guarda le funzionalità ▶</a>
          </div>
          <div class="hero-micro">
            <span>✓ Nessuna carta di credito</span>
            <span>✓ Attivazione immediata</span>
            <span>✓ Assistenza dedicata</span>
          </div>
        </div>

        <div class="mockup">
          <div class="mock-browser">
            <div class="browser-dots"><span></span><span></span><span></span></div>
            <div class="mock-shell">
              <div class="mock-side"><i></i><i></i><i></i><i></i><i></i><i></i></div>
              <div class="mock-main">
                <h4>Dashboard</h4>
                <div class="mock-stats">
                  <div class="mock-stat"><small>Contratti attivi</small><strong>128</strong></div>
                  <div class="mock-stat"><small>Scadenze</small><strong>24</strong></div>
                  <div class="mock-stat"><small>Fatturato mese</small><strong>€48.750</strong></div>
                  <div class="mock-stat"><small>Pagamenti</small><strong>€32.100</strong></div>
                </div>
                <div class="mock-panels">
                  <div class="mock-panel"><b>Panoramica</b><div class="chart-line"></div></div>
                  <div class="mock-panel mock-list"><b>Scadenze prossime</b><p><span>Contratto consulenza</span><span>15 mag</span></p><p><span>Manutenzione</span><span>18 mag</span></p><p><span>Servizio assistenza</span><span>22 mag</span></p></div>
                </div>
              </div>
            </div>
          </div>
          <div class="float-card contracts"><b>Contratti attivi</b><div class="mock-list"><p><span>Consulenza strategica</span><span class="pill green">Attivo</span></p><p><span>Manutenzione annuale</span><span class="pill green">Attivo</span></p><p><span>Servizio supporto</span><span class="pill green">Attivo</span></p></div></div>
          <div class="float-card payments"><b>Pagamenti</b><div class="donut"><div>€32.100</div></div><small>Totale ricevuto</small></div>
        </div>
      </div>
    </div>

    <div class="benefit-strip" id="funzionalita">
      <div class="benefit-item"><div class="benefit-icon">🏢</div><div><b>Multi-azienda</b><small>Gestisci più realtà in modo sicuro.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">👥</div><div><b>Ruoli e permessi</b><small>Accessi staff granulari.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">📄</div><div><b>Contratti e pagamenti</b><small>Rate, acconti e saldi sempre chiari.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">☁️</div><div><b>Accesso cloud</b><small>Online, responsive e sempre disponibile.</small></div></div>
    </div>

    <div class="premium-title">
      <div class="iconbox">⚙️</div>
      <div><h2>Funzionalità principali</h2><p>Tutto ciò che serve per gestire clienti, contratti, lavori e incassi.</p></div>
    </div>
    <div class="feature-grid">
      <div class="feature-card-premium"><div class="ficon">👥</div><b>CRM clienti</b><p>Anagrafiche complete, note, storico attività e documenti sempre a portata di mano.</p></div>
      <div class="feature-card-premium"><div class="ficon">📚</div><b>Contratti e scadenze</b><p>Gestisci contratti, rinnovi, scadenze e alert per non perdere nulla.</p></div>
      <div class="feature-card-premium"><div class="ficon">📊</div><b>Lavori e report</b><p>Organizza lavori, attività e report dettagliati per cliente.</p></div>
      <div class="feature-card-premium"><div class="ficon">💳</div><b>Pagamenti e rate</b><p>Monitora rate, acconti, saldi e residui in modo semplice.</p></div>
      <div class="feature-card-premium"><div class="ficon">🧾</div><b>Fatture interne</b><p>Crea e archivia fatture interne, bozze, invii e storico.</p></div>
      <div class="feature-card-premium"><div class="ficon">🛡️</div><b>Staff e permessi</b><p>Definisci ruoli e accessi per ogni membro del team.</p></div>
      <div class="feature-card-premium"><div class="ficon">🏢</div><b>Multi-azienda SaaS</b><p>Ogni azienda ha dati isolati e gestibili da un unico portale.</p></div>
      <div class="feature-card-premium"><div class="ficon">📈</div><b>Dashboard operative</b><p>KPI e cruscotti per decisioni più rapide e consapevoli.</p></div>
    </div>

    <div class="premium-title" id="come-funziona">
      <div class="iconbox">🚀</div>
      <div><h2>Come funziona</h2><p>Dal piano alla gestione operativa in pochi passaggi.</p></div>
    </div>
    <div class="steps-row">
      <div class="step-card"><div class="step-num">1</div><b>Scegli il piano</b><p>Seleziona il pacchetto più adatto alle tue esigenze.</p></div>
      <div class="step-card"><div class="step-num">2</div><b>Attiva 30 giorni gratis</b><p>Prova tutte le funzionalità senza impegno.</p></div>
      <div class="step-card"><div class="step-num">3</div><b>Configura azienda</b><p>Imposta dati, logo, ruoli, staff e preferenze.</p></div>
      <div class="step-card"><div class="step-num">4</div><b>Gestisci tutto</b><p>Clienti, contratti, scadenze, pagamenti e lavori.</p></div>
    </div>

    <div class="premium-title" id="pacchetti">
      <div class="iconbox">💳</div>
      <div><h2>Scegli il pacchetto</h2><p>Tutti i piani includono 30 giorni di prova gratuita.</p></div>
    </div>
    """)
    st.markdown(hero_html, unsafe_allow_html=True)

    render_public_plan_cards()

    bottom_html = textwrap.dedent("""
    <div class="benefit-strip">
      <div class="benefit-item"><div class="benefit-icon">☁️</div><div><b>Multi-tenant SaaS</b><small>Piattaforma pensata per più aziende.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">👥</div><div><b>Accessi staff con ruoli</b><small>Permessi e responsabilità chiare.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">🔒</div><div><b>Dati azienda separati</b><small>Ogni azienda ha i propri dati isolati.</small></div></div>
      <div class="benefit-item"><div class="benefit-icon">📱</div><div><b>Responsive</b><small>Desktop, tablet e smartphone.</small></div></div>
    </div>

    <div class="final-cta">
      <div>
        <h2>Porta online la gestione della tua azienda</h2>
        <p>Semplifica contratti, clienti, lavori e pagamenti. Inizia oggi la prova gratuita di 30 giorni.</p>
      </div>
      <div class="premium-actions" style="margin:0">
        <a class="html-btn secondary" href="/?public_page=plans">Inizia ora →</a>
        <a class="html-btn primary" href="/?public_page=plans">Richiedi demo</a>
      </div>
    </div>

    <div class="premium-title" id="faq">
      <div class="iconbox">❓</div>
      <div><h2>FAQ</h2><p>Risposte rapide alle domande più frequenti.</p></div>
    </div>
    <div class="feature-grid">
      <div class="feature-card-premium"><b>È inclusa la prova gratuita?</b><p>Sì, tutti i piani hanno 30 giorni gratuiti.</p></div>
      <div class="feature-card-premium"><b>È multi-azienda?</b><p>Sì, ogni azienda ha un ambiente separato con dati isolati.</p></div>
      <div class="feature-card-premium"><b>Lo staff vede i pagamenti?</b><p>Solo se il ruolo assegnato prevede permessi finanziari.</p></div>
      <div class="feature-card-premium"><b>Funziona online?</b><p>Sì, è pensato per desktop, tablet e mobile.</p></div>
    </div>

    <div class="premium-footer">
      <div><b>CUBE Management Contract</b><p>Il gestionale online per aziende, consulenti e società di servizi.</p><small>© 2026 CUBE Management Contract</small></div>
      <div><b>Prodotto</b><a>Funzionalità</a><a>Prezzi</a><a>Integrazioni</a><a>Changelog</a></div>
      <div><b>Azienda</b><a>Chi siamo</a><a>Lavora con noi</a><a>Contatti</a><a>Partner</a></div>
      <div><b>Supporto</b><a>FAQ</a><a>Guide e tutorial</a><a>Assistenza</a><a>Stato servizio</a></div>
      <div><b>Legale</b><a>Termini di servizio</a><a>Privacy policy</a><a>Cookie policy</a><a>DPA</a></div>
    </div>
    """)
    st.markdown(bottom_html, unsafe_allow_html=True)


def render_public_plan_cards():
    plans = get_public_plans()
    if plans.empty:
        st.info("I pacchetti saranno disponibili a breve.")
        return

    order = ["Free", "Starter", "Professional", "Business", "Enterprise"]
    plan_icons = {
        "Free": "👤",
        "Starter": "🚀",
        "Professional": "⭐",
        "Business": "💼",
        "Enterprise": "👑",
    }
    plan_notes = {
        "Free": ["30 giorni gratuiti", "fino a 3 contratti da gestire"],
        "Starter": ["30 giorni gratuiti", "fino a 10 clienti/contratti", "+ 1 membro staff"],
        "Professional": ["30 giorni gratuiti", "fino a 30 clienti", "+ 3 membri staff"],
        "Business": ["30 giorni gratuiti", "fino a 100 clienti", "+ 10 membri staff"],
        "Enterprise": ["30 giorni gratuiti", "tutto illimitato"],
    }

    rows = []
    for name in order:
        df = plans[plans["nome"].astype(str) == name]
        if df.empty:
            continue
        p = df.iloc[0].to_dict()
        featured = " featured" if name == "Professional" else ""
        ribbon = "<div class='popular-ribbon'>Più scelto</div>" if name == "Professional" else ""
        price = money(float(p.get("prezzo_mensile") or 0))
        notes = plan_notes.get(name, [])
        subtitle = notes[0] if notes else ""
        bullet_lines = "<br>".join([safe(x) for x in notes[1:]])
        cta = "Prova gratis" if name == "Free" else "Scegli piano"
        rows.append(f"""
        <div class="pricing-card{featured}">
          {ribbon}
          <div class="pricing-icon">{plan_icons.get(name,'📦')}</div>
          <h3>{safe(name)}</h3>
          <div class="price">{price}<small> /mese</small></div>
          <p>{safe(subtitle)}</p>
          <ul><li>{bullet_lines.replace('<br>','</li><li>')}</li></ul>
          <a class="html-btn {'primary' if name == 'Professional' else 'secondary'}" href="/?public_page=register&plan={safe(name)}">{cta}</a>
        </div>
        """)

    st.markdown("<div class='pricing-grid'>" + "\n".join(rows) + "</div>", unsafe_allow_html=True)


def public_plans_page():
    css()
    public_topbar()
    st.markdown(textwrap.dedent("""
    <div class="premium-title" style="margin-top:8px">
      <div class="iconbox">💳</div>
      <div><h2>Scegli il pacchetto</h2><p>Tutti i piani includono 30 giorni di prova gratuita. Nessun vincolo, puoi cambiare piano quando vuoi.</p></div>
    </div>
    """), unsafe_allow_html=True)
    render_public_plan_cards()
    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("<div class='final-cta'><div><h2>Vuoi iniziare subito?</h2><p>Attiva il tuo account aziendale e configura il portale in pochi minuti.</p></div><div class='premium-actions' style='margin:0'><a class='html-btn secondary' href='/?public_page=home'>Torna alla home</a></div></div>", unsafe_allow_html=True)


def public_login_page():
    css()
    public_topbar()
    st.markdown("<br>", unsafe_allow_html=True)
    section("🔐", "Login area riservata", "Accedi al tuo spazio aziendale.")

    with st.form("login_form_public"):
        username = st.text_input("Username o email")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Accedi")

    if submit:
        df = read_df("""
            SELECT * FROM users
            WHERE stato='Attivo'
              AND (LOWER(username)=LOWER(?) OR LOWER(COALESCE(email,''))=LOWER(?))
            LIMIT 1
        """, (username.strip(), username.strip()))
        if df.empty:
            st.error("Utente non trovato o non attivo.")
        else:
            row = df.iloc[0].to_dict()
            if verify_password(password, row["password_hash"]):
                if row.get("tenant_id") is not None and not pd.isna(row.get("tenant_id")):
                    t = read_df("SELECT * FROM tenants WHERE id=?", (int(row["tenant_id"]),))
                    if t.empty or str(t.iloc[0]["stato_account"]) in ["Sospeso", "Disattivato", "Scaduto"]:
                        st.error("Account azienda non attivo. Contattare l'amministratore SaaS.")
                        return
                st.session_state["user_id"] = int(row["id"])
                st.rerun()
            else:
                st.error("Password errata.")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Non hai un account? Registrati", key="login_to_plans"):
            set_public_page("plans")
            st.rerun()
    with c2:
        if st.button("Torna alla home", key="login_to_home"):
            set_public_page("home")
            st.rerun()

def public_register_page():
    css()
    public_topbar()

    selected_plan = st.session_state.get("selected_plan_name", "")
    plans = get_public_plans()
    if not selected_plan and not plans.empty:
        selected_plan = str(plans.iloc[0]["nome"])

    section("🏢", "Registrazione azienda", "Crea il tuo spazio aziendale. La prova gratuita dura 30 giorni.")

    if plans.empty:
        st.error("Nessun piano disponibile.")
        return

    plan_names = plans["nome"].astype(str).tolist()
    if selected_plan not in plan_names:
        selected_plan = plan_names[0]

    plan_name = st.selectbox("Piano scelto", plan_names, index=plan_names.index(selected_plan))
    selected = plans[plans["nome"].astype(str) == plan_name].iloc[0].to_dict()

    st.success(f"Hai scelto il piano {plan_name}. I primi 30 giorni sono gratuiti.")

    with st.form("register_tenant_public"):
        st.markdown("### Dati azienda")
        c1, c2 = st.columns(2)
        with c1:
            ragione = st.text_input("Ragione sociale *")
            forma = st.text_input("Forma giuridica")
            piva = st.text_input("Partita IVA")
            cf = st.text_input("Codice fiscale")
            sede = st.text_area("Sede legale")
            logo = st.file_uploader("Logo azienda", type=["png", "jpg", "jpeg", "webp"])
        with c2:
            pec = st.text_input("PEC")
            sdi = st.text_input("Codice SDI")
            telefono = st.text_input("Telefono")
            email_azienda = st.text_input("Email azienda")

        st.markdown("### Account amministratore")
        a1, a2 = st.columns(2)
        with a1:
            admin_nome = st.text_input("Nome admin *")
            admin_cognome = st.text_input("Cognome admin")
            admin_email = st.text_input("Email admin *")
        with a2:
            admin_user = st.text_input("Username admin *")
            admin_password = st.text_input("Password admin *", type="password")
            privacy = st.checkbox("Confermo di voler attivare la prova gratuita di 30 giorni")

        reg = st.form_submit_button("Attiva prova gratuita")

    if reg:
        if not privacy:
            st.error("Devi confermare l'attivazione della prova gratuita.")
        elif not ragione.strip() or not admin_nome.strip() or not admin_email.strip() or not admin_user.strip() or not admin_password:
            st.error("Compila ragione sociale e dati admin obbligatori.")
        elif not read_df("SELECT id FROM users WHERE LOWER(username)=LOWER(?) OR LOWER(COALESCE(email,''))=LOWER(?)", (admin_user, admin_email)).empty:
            st.error("Username o email admin già esistenti.")
        else:
            logo_rel = save_upload(logo, None, "tenant_logo_") if logo else None
            tid = execute("""
                INSERT INTO tenants (ragione_sociale,forma_giuridica,partita_iva,codice_fiscale,sede_legale,pec,codice_sdi,telefono,email,logo_file,stato_account,piano_abbonamento,data_registrazione,note)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (ragione, forma, piva, cf, sede, pec, sdi, telefono, email_azienda, logo_rel, "In prova", plan_name, today_iso(), "Registrazione da sito pubblico con prova 30 giorni"))

            execute("""
                INSERT INTO companies (tenant_id,nome,forma_giuridica,piva,cf,sede,pec,codice_sdi,telefono,email,logo_file,is_default,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (tid, ragione, forma, piva, cf, sede, pec, sdi, telefono, email_azienda, logo_rel, 1, now_iso()))

            execute("""
                INSERT INTO users (tenant_id,username,email,password_hash,nome,cognome,ruolo,stato,telefono,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?)
            """, (tid, admin_user, admin_email, hash_password(admin_password), admin_nome, admin_cognome, ROLE_ADMIN, "Attivo", telefono, now_iso()))

            plan_id = int(selected["id"])
            trial_end = (date.today() + timedelta(days=30)).isoformat()
            execute("""
                INSERT INTO subscriptions (tenant_id,piano_id,stato,data_inizio,data_scadenza,metodo_pagamento,note,created_at)
                VALUES (?,?,?,?,?,?,?,?)
            """, (tid, plan_id, "In prova", today_iso(), trial_end, "Prova gratuita", f"Trial 30 giorni piano {plan_name}", now_iso()))

            st.success("Registrazione completata. Ora puoi accedere con l'utente admin creato.")
            set_public_page("login")

def public_router():
    # Permette link diretti tipo:
    # /?public_page=login
    # /?public_page=plans
    # /?public_page=register&plan=Professional
    try:
        qp = st.query_params
        qp_page = qp.get("public_page", None)
        qp_plan = qp.get("plan", None)
        if qp_page:
            st.session_state["public_page"] = str(qp_page)
        if qp_plan:
            st.session_state["selected_plan_name"] = str(qp_plan)
    except Exception:
        pass

    page = st.session_state.get("public_page", "home")
    if page == "login":
        public_login_page()
    elif page == "plans":
        public_plans_page()
    elif page == "register":
        public_register_page()
    else:
        public_landing_page()

# Backwards compatibility
def login_page():
    public_router()


# ============================================================
# SELECT HELPERS
# ============================================================

def select_client(label="Cliente", key="client"):
    df = visible_clients_df()
    if df.empty:
        st.info("Nessun cliente disponibile.")
        return None
    opts = {f"{r['ragione_sociale']} · P.IVA {r.get('partita_iva') or '-'} · ID {int(r['id'])}": int(r["id"]) for _, r in df.iterrows()}
    return opts[st.selectbox(label, list(opts.keys()), key=key)]

def select_user(label="Staff / Responsabile", key="user", include_none=True):
    df = users_for_tenant()
    opts = {}
    if include_none:
        opts["Non assegnato"] = None
    for _, r in df.iterrows():
        opts[f"{r['nome']} {r.get('cognome') or ''} · {r['ruolo']} · ID {int(r['id'])}"] = int(r["id"])
    return opts[st.selectbox(label, list(opts.keys()), key=key)]

def select_contract(label="Contratto", key="contract", include_all=False):
    tid = current_tenant_id()
    if is_super_admin():
        df = read_df("SELECT c.*, cl.ragione_sociale cliente FROM contracts c JOIN clients cl ON cl.id=c.client_id ORDER BY c.id DESC")
    elif can_manage_contracts() or can_finance():
        df = read_df("SELECT c.*, cl.ragione_sociale cliente FROM contracts c JOIN clients cl ON cl.id=c.client_id WHERE c.tenant_id=? ORDER BY c.id DESC", (tid,))
    else:
        df = read_df("""
            SELECT DISTINCT c.*, cl.ragione_sociale cliente
            FROM contracts c
            JOIN clients cl ON cl.id=c.client_id
            LEFT JOIN client_assignments a ON a.client_id=cl.id
            WHERE c.tenant_id=? AND (c.assigned_user_id=? OR a.user_id=?)
            ORDER BY c.id DESC
        """, (tid, int(current_user()["id"]), int(current_user()["id"])))
    if include_all:
        options = {"Tutti i contratti": None}
    else:
        options = {}
    for _, r in df.iterrows():
        options[f"ID {int(r['id'])} · {r['cliente']} · {r['titolo']}"] = int(r["id"])
    if not options:
        st.info("Nessun contratto disponibile.")
        return None
    return options[st.selectbox(label, list(options.keys()), key=key)]


# ============================================================
# SUPER ADMIN PAGES
# ============================================================

def page_super_dashboard():
    header()
    section("🏠", "Super Dashboard SaaS", "Panoramica globale del portale multi-azienda.")
    tenants = read_df("SELECT * FROM tenants")
    users = read_df("SELECT * FROM users")
    subs = read_df("SELECT * FROM subscriptions")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Aziende registrate", len(tenants))
    c2.metric("Utenti totali", len(users))
    c3.metric("Abbonamenti", len(subs))
    c4.metric("Aziende attive", len(tenants[tenants["stato_account"] == "Attivo"]) if not tenants.empty else 0)
    section("🏢", "Ultime aziende")
    if tenants.empty:
        st.info("Nessuna azienda registrata.")
    else:
        st.dataframe(tenants.sort_values("id", ascending=False).head(20), use_container_width=True, hide_index=True)


def page_tenants():
    header()
    section("🏢", "Aziende SaaS", "Il Super Admin può creare, modificare e gestire manualmente tutte le aziende registrate.")
    if not is_super_admin():
        st.error("Accesso riservato al Super Admin SaaS.")
        return

    tab_list, tab_create, tab_edit, tab_admin = st.tabs([
        "📋 Elenco aziende",
        "➕ Crea azienda",
        "✏️ Modifica azienda",
        "👤 Admin azienda"
    ])

    with tab_list:
        df = read_df("SELECT * FROM tenants ORDER BY id DESC")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Aziende totali", len(df))
        c2.metric("Attive", len(df[df["stato_account"] == "Attivo"]) if not df.empty else 0)
        c3.metric("In prova", len(df[df["stato_account"] == "In prova"]) if not df.empty else 0)
        c4.metric("Sospese/scadute", len(df[df["stato_account"].isin(["Sospeso","Scaduto","Disattivato"])]) if not df.empty else 0)

        if df.empty:
            st.info("Nessuna azienda registrata.")
        else:
            view = df.copy()
            st.dataframe(view, use_container_width=True, hide_index=True)

    with tab_create:
        st.subheader("Crea azienda manualmente")
        st.caption("Questa funzione serve quando vuoi registrare tu l'azienda, senza far usare il form pubblico al cliente.")

        with st.form("super_create_tenant"):
            c1, c2 = st.columns(2)
            with c1:
                ragione = st.text_input("Ragione sociale *")
                forma = st.text_input("Forma giuridica", placeholder="Es. S.r.l., S.r.l.s., S.a.s.")
                piva = st.text_input("Partita IVA")
                cf = st.text_input("Codice fiscale")
                sede = st.text_area("Sede legale")
                pec = st.text_input("PEC")
                sdi = st.text_input("Codice SDI")
                iban = st.text_input("IBAN")
            with c2:
                telefono = st.text_input("Telefono")
                email_azienda = st.text_input("Email azienda")
                stato = st.selectbox("Stato account", TENANT_STATUS, index=TENANT_STATUS.index("Attivo"))
                piani = read_df("SELECT * FROM subscription_plans WHERE attivo=1 ORDER BY prezzo_mensile")
                piano_nome = st.selectbox("Piano abbonamento", piani["nome"].tolist() if not piani.empty else ["Starter"])
                logo = st.file_uploader("Logo azienda", type=["png","jpg","jpeg","webp"])
                note = st.text_area("Note interne Super Admin")

            st.markdown("### Primo Admin Azienda")
            a1, a2 = st.columns(2)
            with a1:
                admin_nome = st.text_input("Nome admin *")
                admin_cognome = st.text_input("Cognome admin")
                admin_email = st.text_input("Email admin *")
            with a2:
                admin_username = st.text_input("Username admin *")
                admin_password = st.text_input("Password admin *", type="password")
                admin_tel = st.text_input("Telefono admin")

            submitted = st.form_submit_button("✅ Crea azienda + admin + abbonamento")

        if submitted:
            if not ragione.strip() or not admin_nome.strip() or not admin_email.strip() or not admin_username.strip() or not admin_password:
                st.error("Compila ragione sociale e dati obbligatori del primo admin.")
            elif not read_df("SELECT id FROM users WHERE LOWER(username)=LOWER(?) OR LOWER(COALESCE(email,''))=LOWER(?)", (admin_username, admin_email)).empty:
                st.error("Username o email admin già esistenti.")
            else:
                logo_rel = save_upload(logo, None, "tenant_logo_") if logo else None
                tid = execute("""
                    INSERT INTO tenants (ragione_sociale,forma_giuridica,partita_iva,codice_fiscale,sede_legale,pec,codice_sdi,iban,telefono,email,logo_file,stato_account,piano_abbonamento,data_registrazione,note)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (ragione, forma, piva, cf, sede, pec, sdi, iban, telefono, email_azienda, logo_rel, stato, piano_nome, today_iso(), note))

                execute("""
                    INSERT INTO companies (tenant_id,nome,forma_giuridica,piva,cf,sede,pec,codice_sdi,iban,telefono,email,logo_file,note,is_default,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (tid, ragione, forma, piva, cf, sede, pec, sdi, iban, telefono, email_azienda, logo_rel, note, 1, now_iso()))

                execute("""
                    INSERT INTO users (tenant_id,username,email,password_hash,nome,cognome,ruolo,stato,telefono,note,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?)
                """, (tid, admin_username, admin_email, hash_password(admin_password), admin_nome, admin_cognome, ROLE_ADMIN, "Attivo", admin_tel, "Creato manualmente dal Super Admin", now_iso()))

                piano_id = None
                if not piani.empty:
                    match = piani[piani["nome"] == piano_nome]
                    if not match.empty:
                        piano_id = int(match.iloc[0]["id"])

                execute("""
                    INSERT INTO subscriptions (tenant_id,piano_id,stato,data_inizio,data_scadenza,metodo_pagamento,note,created_at)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (tid, piano_id, stato if stato in ["Attivo","In prova"] else "Sospeso", today_iso(), add_months(date.today(), 1).isoformat(), "Manuale", "Creato dal Super Admin", now_iso()))

                st.success("Azienda creata manualmente con admin aziendale.")
                st.rerun()

    with tab_edit:
        df = read_df("SELECT * FROM tenants ORDER BY ragione_sociale")
        if df.empty:
            st.info("Nessuna azienda da modificare.")
            return

        opts = {f"ID {int(r['id'])} · {r['ragione_sociale']} · {r['stato_account']}": int(r["id"]) for _, r in df.iterrows()}
        tid = opts[st.selectbox("Seleziona azienda", list(opts.keys()), key="edit_tenant_select")]
        t = read_df("SELECT * FROM tenants WHERE id=?", (tid,)).iloc[0].to_dict()

        with st.form("super_edit_tenant"):
            c1, c2 = st.columns(2)
            with c1:
                ragione = st.text_input("Ragione sociale", t.get("ragione_sociale") or "")
                forma = st.text_input("Forma giuridica", t.get("forma_giuridica") or "")
                piva = st.text_input("Partita IVA", t.get("partita_iva") or "")
                cf = st.text_input("Codice fiscale", t.get("codice_fiscale") or "")
                sede = st.text_area("Sede legale", t.get("sede_legale") or "")
                pec = st.text_input("PEC", t.get("pec") or "")
                sdi = st.text_input("Codice SDI", t.get("codice_sdi") or "")
                iban = st.text_input("IBAN", t.get("iban") or "")
            with c2:
                telefono = st.text_input("Telefono", t.get("telefono") or "")
                email = st.text_input("Email", t.get("email") or "")
                stato = st.selectbox("Stato account", TENANT_STATUS, index=TENANT_STATUS.index(t.get("stato_account")) if t.get("stato_account") in TENANT_STATUS else 0)
                piano = st.text_input("Piano abbonamento", t.get("piano_abbonamento") or "")
                logo = st.file_uploader("Sostituisci logo", type=["png","jpg","jpeg","webp"])
                note = st.text_area("Note", t.get("note") or "")
            save_btn = st.form_submit_button("💾 Salva modifiche azienda")

        if save_btn:
            logo_rel = save_upload(logo, tid, "tenant_logo_") if logo else t.get("logo_file")
            execute("""
                UPDATE tenants SET ragione_sociale=?,forma_giuridica=?,partita_iva=?,codice_fiscale=?,sede_legale=?,pec=?,codice_sdi=?,iban=?,telefono=?,email=?,logo_file=?,stato_account=?,piano_abbonamento=?,note=?
                WHERE id=?
            """, (ragione, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, stato, piano, note, tid))

            # Sincronizza anche l'azienda principale interna del tenant.
            company = read_df("SELECT id FROM companies WHERE tenant_id=? AND is_default=1 ORDER BY id DESC LIMIT 1", (tid,))
            if company.empty:
                execute("""
                    INSERT INTO companies (tenant_id,nome,forma_giuridica,piva,cf,sede,pec,codice_sdi,iban,telefono,email,logo_file,note,is_default,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (tid, ragione, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, note, 1, now_iso()))
            else:
                execute("""
                    UPDATE companies SET nome=?,forma_giuridica=?,piva=?,cf=?,sede=?,pec=?,codice_sdi=?,iban=?,telefono=?,email=?,logo_file=?,note=?
                    WHERE id=? AND tenant_id=?
                """, (ragione, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, note, int(company.iloc[0]["id"]), tid))

            st.success("Azienda aggiornata.")
            st.rerun()

        st.markdown("### Azioni rapide")
        a1, a2, a3 = st.columns(3)
        if a1.button("✅ Attiva azienda", key=f"activate_{tid}"):
            execute("UPDATE tenants SET stato_account='Attivo' WHERE id=?", (tid,))
            st.success("Azienda attivata.")
            st.rerun()
        if a2.button("⏸️ Sospendi azienda", key=f"suspend_{tid}"):
            execute("UPDATE tenants SET stato_account='Sospeso' WHERE id=?", (tid,))
            st.warning("Azienda sospesa.")
            st.rerun()
        with a3:
            confirm = st.checkbox("Confermo disattivazione", key=f"confirm_disable_{tid}")
            if st.button("🚫 Disattiva", key=f"disable_{tid}", disabled=not confirm):
                execute("UPDATE tenants SET stato_account='Disattivato' WHERE id=?", (tid,))
                execute("UPDATE users SET stato='Disattivato' WHERE tenant_id=?", (tid,))
                st.warning("Azienda e utenti disattivati.")
                st.rerun()

    with tab_admin:
        st.subheader("Gestione admin/utenti azienda")
        tenants = read_df("SELECT * FROM tenants ORDER BY ragione_sociale")
        if tenants.empty:
            st.info("Nessuna azienda.")
            return
        opts = {f"ID {int(r['id'])} · {r['ragione_sociale']}": int(r["id"]) for _, r in tenants.iterrows()}
        tid = opts[st.selectbox("Azienda", list(opts.keys()), key="tenant_admin_users")]

        users = read_df("SELECT * FROM users WHERE tenant_id=? ORDER BY id DESC", (tid,))
        if users.empty:
            st.info("Nessun utente aziendale.")
        else:
            st.dataframe(users[["id","username","email","nome","cognome","ruolo","stato","telefono","created_at"]], use_container_width=True, hide_index=True)

        st.markdown("### Crea nuovo utente per questa azienda")
        with st.form("super_create_user_for_tenant"):
            c1, c2, c3 = st.columns(3)
            username = c1.text_input("Username *")
            email = c1.text_input("Email")
            password = c1.text_input("Password *", type="password")
            nome = c2.text_input("Nome *")
            cognome = c2.text_input("Cognome")
            telefono = c2.text_input("Telefono")
            ruolo = c3.selectbox("Ruolo", ROLES, index=0)
            stato = c3.selectbox("Stato", ["Attivo","Sospeso","Disattivato"], index=0)
            note = st.text_area("Note")
            if st.form_submit_button("Crea utente azienda"):
                if not username or not password or not nome:
                    st.error("Username, password e nome sono obbligatori.")
                elif not read_df("SELECT id FROM users WHERE LOWER(username)=LOWER(?)", (username,)).empty:
                    st.error("Username già esistente.")
                else:
                    execute("""
                        INSERT INTO users (tenant_id,username,email,password_hash,nome,cognome,ruolo,stato,telefono,note,created_at)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?)
                    """, (tid, username, email, hash_password(password), nome, cognome, ruolo, stato, telefono, note, now_iso()))
                    st.success("Utente creato.")
                    st.rerun()


def page_plans():
    header()
    section("💳", "Piani abbonamento", "Il Super Admin può creare, modificare, attivare e disattivare tutti i piani SaaS.")
    if not is_super_admin():
        st.error("Accesso riservato al Super Admin SaaS.")
        return

    tab_list, tab_create, tab_edit = st.tabs(["📋 Elenco piani", "➕ Crea piano", "✏️ Modifica piano"])

    with tab_list:
        df = read_df("SELECT * FROM subscription_plans ORDER BY prezzo_mensile")
        if df.empty:
            st.info("Nessun piano.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_create:
        with st.form("new_plan"):
            st.subheader("Crea nuovo piano")
            c1, c2, c3, c4 = st.columns(4)
            nome = c1.text_input("Nome piano")
            prezzo = c2.number_input("Prezzo mensile", min_value=0.0, value=0.0)
            utenti = c3.number_input("Max utenti", min_value=1, value=3)
            clienti = c4.number_input("Max clienti", min_value=1, value=50)
            contratti = st.number_input("Max contratti", min_value=1, value=100)
            funzioni = st.text_area("Funzioni, una per riga")
            attivo = st.checkbox("Piano attivo", value=True)
            if st.form_submit_button("Crea piano"):
                if not nome:
                    st.error("Nome piano obbligatorio.")
                else:
                    execute("INSERT INTO subscription_plans (nome,prezzo_mensile,max_utenti,max_clienti,max_contratti,funzioni_json,attivo,created_at) VALUES (?,?,?,?,?,?,?,?)",
                            (nome, prezzo, utenti, clienti, contratti, json.dumps([x.strip() for x in funzioni.splitlines() if x.strip()]), 1 if attivo else 0, now_iso()))
                    st.success("Piano creato.")
                    st.rerun()

    with tab_edit:
        df = read_df("SELECT * FROM subscription_plans ORDER BY prezzo_mensile")
        if df.empty:
            st.info("Nessun piano da modificare.")
            return

        opts = {f"ID {int(r['id'])} · {r['nome']} · {money(r['prezzo_mensile'])}/mese": int(r["id"]) for _, r in df.iterrows()}
        pid = opts[st.selectbox("Seleziona piano", list(opts.keys()), key="edit_plan_select")]
        p = df[df["id"] == pid].iloc[0].to_dict()
        try:
            funzioni_text = "\n".join(json.loads(p.get("funzioni_json") or "[]"))
        except Exception:
            funzioni_text = p.get("funzioni_json") or ""

        with st.form("edit_plan_form"):
            c1, c2, c3, c4 = st.columns(4)
            nome = c1.text_input("Nome piano", p.get("nome") or "")
            prezzo = c2.number_input("Prezzo mensile", min_value=0.0, value=float(p.get("prezzo_mensile") or 0))
            utenti = c3.number_input("Max utenti", min_value=1, value=int(p.get("max_utenti") or 1))
            clienti = c4.number_input("Max clienti", min_value=1, value=int(p.get("max_clienti") or 1))
            contratti = st.number_input("Max contratti", min_value=1, value=int(p.get("max_contratti") or 1))
            funzioni = st.text_area("Funzioni, una per riga", value=funzioni_text)
            attivo = st.checkbox("Piano attivo", value=bool(int(p.get("attivo") or 0)))
            if st.form_submit_button("💾 Salva modifiche piano"):
                execute("""
                    UPDATE subscription_plans SET nome=?,prezzo_mensile=?,max_utenti=?,max_clienti=?,max_contratti=?,funzioni_json=?,attivo=?
                    WHERE id=?
                """, (nome, prezzo, utenti, clienti, contratti, json.dumps([x.strip() for x in funzioni.splitlines() if x.strip()]), 1 if attivo else 0, pid))
                st.success("Piano aggiornato.")
                st.rerun()

        st.markdown("### Azioni piano")
        col1, col2 = st.columns(2)
        if col1.button("✅ Attiva", key=f"plan_active_{pid}"):
            execute("UPDATE subscription_plans SET attivo=1 WHERE id=?", (pid,))
            st.success("Piano attivato.")
            st.rerun()
        if col2.button("⛔ Disattiva", key=f"plan_inactive_{pid}"):
            execute("UPDATE subscription_plans SET attivo=0 WHERE id=?", (pid,))
            st.warning("Piano disattivato.")
            st.rerun()


def page_global_users():
    header()
    section("👤", "Utenti globali", "Il Super Admin può vedere, creare e modificare tutti gli utenti di tutte le aziende.")
    if not is_super_admin():
        st.error("Accesso riservato al Super Admin SaaS.")
        return

    df = read_df("""
        SELECT u.id, u.tenant_id, t.ragione_sociale tenant, u.username, u.email, u.nome, u.cognome, u.ruolo, u.stato, u.telefono, u.note, u.created_at
        FROM users u LEFT JOIN tenants t ON t.id=u.tenant_id
        ORDER BY u.id DESC
    """)
    if not df.empty:
        st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Modifica utente globale")
    if df.empty:
        return

    opts = {f"ID {int(r['id'])} · {r['username']} · {r['ruolo']} · {r.get('tenant') or 'Piattaforma'}": int(r["id"]) for _, r in df.iterrows()}
    uid = opts[st.selectbox("Utente", list(opts.keys()), key="global_user_edit")]
    u = read_df("SELECT * FROM users WHERE id=?", (uid,)).iloc[0].to_dict()

    tenants = read_df("SELECT id, ragione_sociale FROM tenants ORDER BY ragione_sociale")
    tenant_opts = {"Piattaforma / Super Admin": None}
    for _, r in tenants.iterrows():
        tenant_opts[f"ID {int(r['id'])} · {r['ragione_sociale']}"] = int(r["id"])

    current_tid = None if u.get("tenant_id") is None or pd.isna(u.get("tenant_id")) else int(u.get("tenant_id"))
    default_label = "Piattaforma / Super Admin"
    for label, val in tenant_opts.items():
        if val == current_tid:
            default_label = label
            break

    with st.form("global_user_edit_form"):
        c1, c2, c3 = st.columns(3)
        tenant_label = c1.selectbox("Azienda / tenant", list(tenant_opts.keys()), index=list(tenant_opts.keys()).index(default_label))
        username = c1.text_input("Username", u.get("username") or "")
        email = c1.text_input("Email", u.get("email") or "")
        nome = c2.text_input("Nome", u.get("nome") or "")
        cognome = c2.text_input("Cognome", u.get("cognome") or "")
        telefono = c2.text_input("Telefono", u.get("telefono") or "")
        possible_roles = [ROLE_SUPER_ADMIN] + ROLES
        ruolo = c3.selectbox("Ruolo", possible_roles, index=possible_roles.index(u.get("ruolo")) if u.get("ruolo") in possible_roles else 1)
        stato = c3.selectbox("Stato", ["Attivo","Sospeso","Disattivato"], index=["Attivo","Sospeso","Disattivato"].index(u.get("stato")) if u.get("stato") in ["Attivo","Sospeso","Disattivato"] else 0)
        password = st.text_input("Nuova password, lascia vuoto per non cambiarla", type="password")
        note = st.text_area("Note", u.get("note") or "")
        if st.form_submit_button("💾 Salva utente"):
            new_tid = tenant_opts[tenant_label]
            if ruolo == ROLE_SUPER_ADMIN:
                new_tid = None
            if password:
                execute("UPDATE users SET tenant_id=?,username=?,email=?,password_hash=?,nome=?,cognome=?,ruolo=?,stato=?,telefono=?,note=? WHERE id=?",
                        (new_tid, username, email, hash_password(password), nome, cognome, ruolo, stato, telefono, note, uid))
            else:
                execute("UPDATE users SET tenant_id=?,username=?,email=?,nome=?,cognome=?,ruolo=?,stato=?,telefono=?,note=? WHERE id=?",
                        (new_tid, username, email, nome, cognome, ruolo, stato, telefono, note, uid))
            st.success("Utente aggiornato.")
            st.rerun()

def page_status():
    header()
    section("📊", "Log / Stato", "Stato tecnico sintetico.")
    st.write("Database:", "PostgreSQL" if IS_POSTGRES else "SQLite locale")
    st.write("Upload directory:", str(UPLOAD_DIR))
    st.write("Current user:", current_user())


# ============================================================
# TENANT DASHBOARD + COMPANY
# ============================================================

def page_dashboard():
    header()
    tid = current_tenant_id()
    section("🏠", "Dashboard azienda", "Dati isolati della tua azienda.")
    clients = visible_clients_df()
    contracts = read_df("SELECT * FROM contracts WHERE tenant_id=?", (tid,))
    work = read_df("SELECT * FROM work_logs WHERE tenant_id=?", (tid,))
    docs = read_df("SELECT * FROM documents WHERE tenant_id=?", (tid,))
    pay = read_df("SELECT * FROM payments WHERE tenant_id=?", (tid,)) if can_finance() else pd.DataFrame()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Clienti", len(clients))
    c2.metric("Contratti", len(contracts))
    c3.metric("Lavori", len(work))
    c4.metric("Documenti", len(docs))
    if can_finance():
        c5, c6, c7 = st.columns(3)
        c5.metric("Totale rate", money(pay["totale"].sum() if not pay.empty else 0))
        c6.metric("Incassato", money(payment_paid_total_for_tenant(tid)))
        c7.metric("Residuo", money(payment_residue_total_for_tenant(tid)))
    section("👥", "Clienti recenti")
    if clients.empty:
        st.info("Nessun cliente.")
    else:
        st.dataframe(clients.head(20), use_container_width=True, hide_index=True)

def page_company():
    header()
    section("🏢", "Dati azienda", "Anagrafica, logo e dati fiscali visibili nei documenti.")
    tid = current_tenant_id()
    company = read_df("SELECT * FROM companies WHERE tenant_id=? AND is_default=1 ORDER BY id DESC LIMIT 1", (tid,))
    row = {} if company.empty else company.iloc[0].to_dict()
    with st.form("company_form"):
        c1, c2 = st.columns(2)
        nome = c1.text_input("Nome azienda *", row.get("nome", ""))
        forma = c1.text_input("Forma giuridica", row.get("forma_giuridica", "") or "")
        piva = c1.text_input("P.IVA", row.get("piva", "") or "")
        cf = c1.text_input("CF", row.get("cf", "") or "")
        pec = c1.text_input("PEC", row.get("pec", "") or "")
        sdi = c1.text_input("Codice SDI", row.get("codice_sdi", "") or "")
        sede = c2.text_area("Sede", row.get("sede", "") or "")
        iban = c2.text_input("IBAN", row.get("iban", "") or "")
        telefono = c2.text_input("Telefono", row.get("telefono", "") or "")
        email = c2.text_input("Email", row.get("email", "") or "")
        logo = st.file_uploader("Logo azienda", type=["png", "jpg", "jpeg", "webp"])
        note = st.text_area("Note", row.get("note", "") or "")
        if st.form_submit_button("Salva dati azienda"):
            if not nome:
                st.error("Nome azienda obbligatorio.")
            else:
                logo_rel = save_upload(logo, tid, "logo_") if logo else row.get("logo_file")
                if company.empty:
                    execute("INSERT INTO companies (tenant_id,nome,forma_giuridica,piva,cf,sede,pec,codice_sdi,iban,telefono,email,logo_file,note,is_default,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)",
                            (tid, nome, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, note, 1, now_iso()))
                else:
                    execute("UPDATE companies SET nome=?,forma_giuridica=?,piva=?,cf=?,sede=?,pec=?,codice_sdi=?,iban=?,telefono=?,email=?,logo_file=?,note=? WHERE id=? AND tenant_id=?",
                            (nome, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, note, row["id"], tid))
                execute("UPDATE tenants SET ragione_sociale=?,forma_giuridica=?,partita_iva=?,codice_fiscale=?,sede_legale=?,pec=?,codice_sdi=?,iban=?,telefono=?,email=?,logo_file=? WHERE id=?",
                        (nome, forma, piva, cf, sede, pec, sdi, iban, telefono, email, logo_rel, tid))
                st.success("Dati azienda salvati.")
                st.rerun()


# ============================================================
# STAFF / USERS
# ============================================================

def page_staff():
    header()
    section("👤", "Staff", "Utenti aziendali, livelli di accesso e permessi.")
    if not can_manage_users():
        st.error("Accesso non consentito.")
        return
    tid = current_tenant_id()
    df = users_for_tenant(active_only=False)
    st.dataframe(df[["id","username","email","nome","cognome","ruolo","stato","telefono","note"]] if not df.empty else df, use_container_width=True, hide_index=True)

    st.subheader("Crea / modifica utente")
    ids = [0] + df["id"].astype(int).tolist() if not df.empty else [0]
    choice = st.selectbox("Utente", ids, format_func=lambda x: "Nuovo utente" if x == 0 else f"ID {x} - {df[df.id==x].iloc[0]['nome']} {df[df.id==x].iloc[0]['cognome'] or ''}")
    row = {} if choice == 0 else df[df.id == choice].iloc[0].to_dict()
    with st.form("user_form"):
        c1, c2, c3 = st.columns(3)
        username = c1.text_input("Username *", row.get("username", "") or "")
        email = c1.text_input("Email", row.get("email", "") or "")
        nome = c2.text_input("Nome *", row.get("nome", "") or "")
        cognome = c2.text_input("Cognome", row.get("cognome", "") or "")
        ruolo = c3.selectbox("Ruolo / livello", ROLES, index=ROLES.index(row.get("ruolo")) if row.get("ruolo") in ROLES else 0)
        stato = c3.selectbox("Stato", ["Attivo", "Sospeso", "Disattivato"], index=["Attivo", "Sospeso", "Disattivato"].index(row.get("stato")) if row.get("stato") in ["Attivo", "Sospeso", "Disattivato"] else 0)
        telefono = st.text_input("Telefono", row.get("telefono", "") or "")
        password = st.text_input("Password nuova / iniziale", type="password")
        note = st.text_area("Note", row.get("note", "") or "")
        if st.form_submit_button("Salva utente"):
            if not username or not nome:
                st.error("Username e nome obbligatori.")
            elif choice == 0:
                if not password:
                    st.error("Password obbligatoria per nuovo utente.")
                else:
                    execute("INSERT INTO users (tenant_id,username,email,password_hash,nome,cognome,ruolo,stato,telefono,note,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                            (tid, username, email, hash_password(password), nome, cognome, ruolo, stato, telefono, note, now_iso()))
                    st.success("Utente creato.")
                    st.rerun()
            else:
                if password:
                    execute("UPDATE users SET username=?,email=?,password_hash=?,nome=?,cognome=?,ruolo=?,stato=?,telefono=?,note=? WHERE id=? AND tenant_id=?",
                            (username, email, hash_password(password), nome, cognome, ruolo, stato, telefono, note, choice, tid))
                else:
                    execute("UPDATE users SET username=?,email=?,nome=?,cognome=?,ruolo=?,stato=?,telefono=?,note=? WHERE id=? AND tenant_id=?",
                            (username, email, nome, cognome, ruolo, stato, telefono, note, choice, tid))
                st.success("Utente aggiornato.")
                st.rerun()


# ============================================================
# CLIENTS
# ============================================================

def page_clients():
    header()
    section("👥", "Clienti CRM", "Anagrafiche clienti isolate per azienda.")
    tid = current_tenant_id()
    tab_list, tab_edit = st.tabs(["Elenco / scheda", "Nuovo / modifica"])

    with tab_list:
        df = visible_clients_df()
        if df.empty:
            st.info("Nessun cliente.")
        else:
            st.dataframe(df, use_container_width=True, hide_index=True)
            cid = select_client("Apri cliente", "open_client")
            if cid:
                client_detail(cid)

    with tab_edit:
        if role() == ROLE_BASE:
            st.warning("Il livello Operativo Base non può creare o modificare clienti.")
            return
        df_all = read_df("SELECT * FROM clients WHERE tenant_id=? ORDER BY ragione_sociale", (tid,))
        ids = [0] + df_all["id"].astype(int).tolist() if not df_all.empty else [0]
        choice = st.selectbox("Cliente", ids, format_func=lambda x: "Nuovo cliente" if x == 0 else f"ID {x} - {df_all[df_all.id==x].iloc[0]['ragione_sociale']}")
        row = {} if choice == 0 else df_all[df_all.id == choice].iloc[0].to_dict()
        with st.form("client_form"):
            c1, c2 = st.columns(2)
            rag = c1.text_input("Ragione sociale *", row.get("ragione_sociale", "") or "")
            forma = c1.text_input("Forma giuridica", row.get("forma_giuridica", "") or "")
            piva = c1.text_input("P.IVA", row.get("partita_iva", "") or "")
            cf = c1.text_input("Codice fiscale", row.get("codice_fiscale", "") or "")
            rea = c1.text_input("REA", row.get("rea", "") or "")
            ateco = c1.text_input("ATECO", row.get("codice_ateco", "") or "")
            sede = c2.text_area("Sede legale", row.get("sede_legale", "") or "")
            pec = c2.text_input("PEC", row.get("pec", "") or "")
            sdi = c2.text_input("Codice SDI", row.get("codice_sdi", "") or "")
            legale = c2.text_input("Legale rappresentante", row.get("legale_rappresentante", "") or "")
            tel = c2.text_input("Telefono", row.get("telefono", "") or "")
            email = c2.text_input("Email", row.get("email", "") or "")
            settore = st.text_input("Settore", row.get("settore", "") or "")
            stato = st.selectbox("Stato CRM", ["Attivo", "Prospect", "In pausa", "Chiuso"], index=0)
            assigned = select_user("Assegna a membro staff", "assign_client", include_none=True)
            note = st.text_area("Note", row.get("note", "") or "")
            if st.form_submit_button("Salva cliente"):
                if not rag:
                    st.error("Ragione sociale obbligatoria.")
                elif choice == 0:
                    cid_new = execute("""
                        INSERT INTO clients (tenant_id,ragione_sociale,forma_giuridica,partita_iva,codice_fiscale,rea,sede_legale,pec,codice_sdi,legale_rappresentante,telefono,email,codice_ateco,settore,stato_crm,note,created_at)
                        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                    """, (tid, rag, forma, piva, cf, rea, sede, pec, sdi, legale, tel, email, ateco, settore, stato, note, now_iso()))
                    if assigned:
                        execute("INSERT INTO client_assignments (tenant_id,client_id,user_id,created_at) VALUES (?,?,?,?)", (tid, cid_new, assigned, now_iso()))
                    st.success("Cliente creato.")
                    st.rerun()
                else:
                    execute("""
                        UPDATE clients SET ragione_sociale=?,forma_giuridica=?,partita_iva=?,codice_fiscale=?,rea=?,sede_legale=?,pec=?,codice_sdi=?,legale_rappresentante=?,telefono=?,email=?,codice_ateco=?,settore=?,stato_crm=?,note=? WHERE id=? AND tenant_id=?
                    """, (rag, forma, piva, cf, rea, sede, pec, sdi, legale, tel, email, ateco, settore, stato, note, choice, tid))
                    if assigned:
                        execute("DELETE FROM client_assignments WHERE tenant_id=? AND client_id=? AND user_id=?", (tid, choice, assigned))
                        execute("INSERT INTO client_assignments (tenant_id,client_id,user_id,created_at) VALUES (?,?,?,?)", (tid, choice, assigned, now_iso()))
                    st.success("Cliente aggiornato.")
                    st.rerun()

def client_detail(cid: int):
    tid = current_tenant_id()
    c = read_df("SELECT * FROM clients WHERE id=? AND tenant_id=?", (cid, tid))
    if c.empty:
        st.error("Cliente non trovato.")
        return
    row = c.iloc[0].to_dict()
    st.subheader(row["ragione_sociale"])
    c1, c2, c3 = st.columns(3)
    c1.metric("P.IVA", row.get("partita_iva") or "-")
    c2.metric("PEC", row.get("pec") or "-")
    c3.metric("Stato", row.get("stato_crm") or "-")
    tabs = st.tabs(["Contratti", "Lavori", "Documenti", "Feedback"])
    with tabs[0]:
        contracts = read_df("SELECT id,titolo,data_decorrenza,data_scadenza,stato FROM contracts WHERE tenant_id=? AND client_id=? ORDER BY id DESC", (tid, cid))
        st.dataframe(contracts, use_container_width=True, hide_index=True)
    with tabs[1]:
        render_work_form(cid)
    with tabs[2]:
        render_documents(cid)
    with tabs[3]:
        render_feedback(cid)


# ============================================================
# CONTRACTS + PAYMENTS
# ============================================================

DEFAULT_TEMPLATE = """CONTRATTO DI CONSULENZA STRATEGICA E OPERATIVA

Tra {{AZIENDA}} e {{CLIENTE}}.

ART. 1 - OGGETTO
Il presente contratto disciplina lo svolgimento dei servizi selezionati nel gestionale.

ART. 2 - SERVIZI
{{SERVIZI}}

ART. 3 - DURATA
La durata è pari a {{DURATA_MESI}} mesi, con decorrenza {{DATA_DECORRENZA}} e scadenza {{DATA_SCADENZA}}.

ART. 4 - CORRISPETTIVO
Il corrispettivo complessivo è pari a {{IMPORTO}} oltre IVA, secondo la modalità di pagamento {{MODALITA_PAGAMENTO}}.

ART. 5 - FORO
Foro competente: {{FORO}}.
"""

def page_templates():
    header()
    section("📄", "Template contratti", "Template globali e aziendali.")
    tid = current_tenant_id()
    df = read_df("SELECT * FROM contract_templates WHERE tenant_id=? OR tenant_id IS NULL ORDER BY tenant_id NULLS FIRST, id DESC", (tid,)) if IS_POSTGRES else read_df("SELECT * FROM contract_templates WHERE tenant_id=? OR tenant_id IS NULL ORDER BY id DESC", (tid,))
    st.dataframe(df[["id","tenant_id","nome","descrizione","attivo"]], use_container_width=True, hide_index=True)
    with st.form("template_form"):
        nome = st.text_input("Nome template")
        descr = st.text_input("Descrizione")
        testo = st.text_area("Testo base", value=DEFAULT_TEMPLATE, height=360)
        if st.form_submit_button("Salva template aziendale"):
            execute("INSERT INTO contract_templates (tenant_id,nome,descrizione,testo_base,attivo,created_at) VALUES (?,?,?,?,?,?)", (tid, nome, descr, testo, 1, now_iso()))
            st.success("Template salvato.")
            st.rerun()

def page_contracts():
    header()
    section("📚", "Contratti", "Creazione, assegnazione staff, servizi e piano pagamenti.")
    if not can_manage_contracts():
        st.error("Accesso non consentito.")
        return
    tid = current_tenant_id()
    tab_list, tab_new = st.tabs(["Elenco", "Crea nuovo contratto"])

    with tab_list:
        df = read_df("""
            SELECT c.id, cl.ragione_sociale cliente, u.nome responsabile, c.titolo, c.data_decorrenza, c.data_scadenza, c.importo_totale, c.iva_percentuale, c.stato
            FROM contracts c
            JOIN clients cl ON cl.id=c.client_id
            LEFT JOIN users u ON u.id=c.assigned_user_id
            WHERE c.tenant_id=?
            ORDER BY c.id DESC
        """, (tid,))
        if df.empty:
            st.info("Nessun contratto.")
        else:
            v = df.copy()
            v["importo_totale"] = v["importo_totale"].apply(money)
            st.dataframe(v, use_container_width=True, hide_index=True)

    with tab_new:
        clients = read_df("SELECT * FROM clients WHERE tenant_id=? ORDER BY ragione_sociale", (tid,))
        if clients.empty:
            st.warning("Crea prima un cliente.")
            return
        with st.form("new_contract"):
            client_id = select_client("Cliente contratto", "contract_client")
            assigned = select_user("Responsabile staff", "contract_staff", include_none=True)
            title = st.text_input("Titolo", "CONTRATTO DI CONSULENZA STRATEGICA E OPERATIVA")
            c1, c2, c3 = st.columns(3)
            decor = c1.date_input("Decorrenza", value=date.today())
            durata = c2.number_input("Durata mesi", min_value=1, value=12)
            scad = add_months(decor, int(durata))
            c3.write(f"Scadenza: **{scad.isoformat()}**")
            importo = c1.number_input("Importo netto totale", min_value=0.0, value=1200.0)
            iva_pct = c2.number_input("IVA %", min_value=0.0, value=22.0)
            modalita = c3.selectbox("Modalità pagamento", ["Mensile", "Trimestrale", "Semestrale", "Annuale", "Unica soluzione"])
            stato = st.selectbox("Stato", CONTRACT_STATUS, index=CONTRACT_STATUS.index("Bozza"))
            foro = st.text_input("Foro competente", "Napoli")
            servizi = st.text_area("Servizi, uno per riga", "Gestione progetto\nConsulenza strategica\nReport mensile")
            note = st.text_area("Note")
            if st.form_submit_button("Crea contratto e piano rate"):
                cid = execute("""
                    INSERT INTO contracts (tenant_id,client_id,assigned_user_id,titolo,tipo_contratto,data_firma,luogo_firma,data_decorrenza,data_scadenza,durata_mesi,importo_totale,iva_percentuale,modalita_pagamento,foro_competente,stato,servizi_json,note,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (tid, client_id, assigned, title, "Consulenza", today_iso(), "Napoli", decor.isoformat(), scad.isoformat(), int(durata), float(importo), float(iva_pct), modalita, foro, stato, json.dumps([x.strip() for x in servizi.splitlines() if x.strip()]), note, now_iso()))
                if assigned and client_id:
                    execute("DELETE FROM client_assignments WHERE tenant_id=? AND client_id=? AND user_id=?", (tid, client_id, assigned))
                    execute("INSERT INTO client_assignments (tenant_id,client_id,user_id,created_at) VALUES (?,?,?,?)", (tid, client_id, assigned, now_iso()))
                generate_payment_plan(cid)
                st.success("Contratto creato con piano rate.")
                st.rerun()

def generate_payment_plan(contract_id: int):
    cdf = read_df("SELECT * FROM contracts WHERE id=?", (contract_id,))
    if cdf.empty:
        return
    c = cdf.iloc[0].to_dict()
    tid = int(c["tenant_id"])
    client_id = int(c["client_id"])
    execute("DELETE FROM payments WHERE contract_id=? AND tenant_id=?", (contract_id, tid))
    modalita = c["modalita_pagamento"]
    durata = int(c["durata_mesi"])
    if modalita == "Mensile":
        n = durata
        step = 1
    elif modalita == "Trimestrale":
        n = max(1, (durata + 2) // 3)
        step = 3
    elif modalita == "Semestrale":
        n = max(1, (durata + 5) // 6)
        step = 6
    elif modalita == "Annuale":
        n = max(1, (durata + 11) // 12)
        step = 12
    else:
        n = 1
        step = durata
    imponibile_rata = float(c["importo_totale"]) / n if n else float(c["importo_totale"])
    iva_rata = imponibile_rata * float(c["iva_percentuale"]) / 100
    decor = datetime.fromisoformat(c["data_decorrenza"]).date()
    for i in range(1, n + 1):
        due = add_months(decor, step * (i - 1))
        execute("""
            INSERT INTO payments (tenant_id,contract_id,client_id,numero_rata,data_scadenza,imponibile,iva,totale,stato,created_at)
            VALUES (?,?,?,?,?,?,?,?,?,?)
        """, (tid, contract_id, client_id, i, due.isoformat(), imponibile_rata, iva_rata, imponibile_rata + iva_rata, "Da pagare", now_iso()))


# ============================================================
# PAYMENTS
# ============================================================

def payment_paid(payment_id: int) -> float:
    df = read_df("SELECT COALESCE(SUM(importo_pagato),0) totale FROM payment_movements WHERE payment_id=?", (payment_id,))
    return float(df.iloc[0]["totale"] or 0)

def update_payment_status(payment_id: int):
    p = read_df("SELECT * FROM payments WHERE id=?", (payment_id,))
    if p.empty:
        return
    row = p.iloc[0].to_dict()
    paid = payment_paid(payment_id)
    total = float(row["totale"])
    if paid >= total - 0.01:
        stato = "Pagata"
        data_pagamento = today_iso()
    elif paid > 0:
        stato = "Acconto"
        data_pagamento = None
    elif row["data_scadenza"] < today_iso():
        stato = "Scaduta"
        data_pagamento = None
    else:
        stato = "Da pagare"
        data_pagamento = None
    execute("UPDATE payments SET stato=?, data_pagamento=? WHERE id=?", (stato, data_pagamento, payment_id))

def payments_df() -> pd.DataFrame:
    tid = current_tenant_id()
    if is_super_admin():
        return read_df("""
            SELECT p.*, cl.ragione_sociale cliente, c.titolo contratto,
                   COALESCE((SELECT SUM(m.importo_pagato) FROM payment_movements m WHERE m.payment_id=p.id),0) pagato
            FROM payments p JOIN clients cl ON cl.id=p.client_id JOIN contracts c ON c.id=p.contract_id
            ORDER BY p.data_scadenza
        """)
    return read_df("""
        SELECT p.*, cl.ragione_sociale cliente, c.titolo contratto,
               COALESCE((SELECT SUM(m.importo_pagato) FROM payment_movements m WHERE m.payment_id=p.id),0) pagato
        FROM payments p JOIN clients cl ON cl.id=p.client_id JOIN contracts c ON c.id=p.contract_id
        WHERE p.tenant_id=?
        ORDER BY p.data_scadenza
    """, (tid,))

def payment_paid_total_for_tenant(tid: int) -> float:
    df = read_df("SELECT COALESCE(SUM(importo_pagato),0) totale FROM payment_movements WHERE tenant_id=?", (tid,))
    return float(df.iloc[0]["totale"] or 0)

def payment_residue_total_for_tenant(tid: int) -> float:
    df = payments_df()
    if df.empty:
        return 0.0
    return float((df["totale"].astype(float) - df["pagato"].astype(float)).clip(lower=0).sum())

def page_payments():
    header()
    section("💶", "Pagamenti", "Rate, acconti, saldi, allegati e tracciamento operatore.")
    if not can_finance():
        st.error("Accesso non consentito.")
        return
    df = payments_df()
    if df.empty:
        st.info("Nessuna rata.")
        return
    df["residuo"] = (df["totale"].astype(float) - df["pagato"].astype(float)).clip(lower=0)
    years = ["Tutti"] + sorted({str(x)[:4] for x in df["data_scadenza"].astype(str) if len(str(x)) >= 4})
    months = ["Tutti"] + [f"{i:02d}" for i in range(1, 13)]
    c1, c2, c3 = st.columns(3)
    y = c1.selectbox("Anno", years)
    m = c2.selectbox("Mese", months)
    status = c3.selectbox("Stato", ["Tutti"] + PAYMENT_STATUS)
    view = df.copy()
    if y != "Tutti":
        view = view[view["data_scadenza"].astype(str).str[:4] == y]
    if m != "Tutti":
        view = view[view["data_scadenza"].astype(str).str[5:7] == m]
    if status != "Tutti":
        view = view[view["stato"] == status]
    summary1, summary2, summary3, summary4 = st.columns(4)
    summary1.metric("Totale rate", money(view["totale"].sum()))
    summary2.metric("Incassato", money(view["pagato"].sum()))
    summary3.metric("Residuo", money(view["residuo"].sum()))
    summary4.metric("Rate", len(view))
    show = view[["id","cliente","contratto","numero_rata","data_scadenza","imponibile","iva","totale","pagato","residuo","stato"]].copy()
    for col in ["imponibile","iva","totale","pagato","residuo"]:
        show[col] = show[col].apply(money)
    st.dataframe(show, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Registra pagamento / acconto")
    opts = {f"ID {r['id']} · {r['cliente']} · rata {r['numero_rata']} · residuo {money(r['residuo'])}": int(r["id"]) for _, r in view.iterrows()}
    if not opts:
        return
    pid = opts[st.selectbox("Rata", list(opts.keys()))]
    with st.form("pay_form"):
        c1, c2, c3 = st.columns(3)
        tipo = c1.selectbox("Tipo movimento", ["Acconto", "Saldo", "Rettifica"])
        imp = c2.number_input("Importo pagato", min_value=0.0, value=0.0)
        data_pag = c3.date_input("Data pagamento", value=date.today())
        allegato = st.file_uploader("Allega PDF/foto pagamento", type=["pdf","png","jpg","jpeg","webp"])
        note = st.text_area("Note")
        if st.form_submit_button("Salva movimento"):
            rel = save_upload(allegato, current_tenant_id(), "pagamento_") if allegato else None
            execute("""
                INSERT INTO payment_movements (tenant_id,payment_id,importo_pagato,tipo_movimento,data_pagamento,allegato_file,note,registrato_da_user_id,created_at)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (current_tenant_id(), pid, float(imp), tipo, data_pag.isoformat(), rel, note, int(current_user()["id"]), now_iso()))
            update_payment_status(pid)
            st.success("Movimento registrato.")
            st.rerun()

    st.subheader("Storico movimenti rata selezionata")
    mov = read_df("""
        SELECT m.id,m.tipo_movimento,m.data_pagamento,m.importo_pagato,u.nome operatore,m.allegato_file,m.note
        FROM payment_movements m LEFT JOIN users u ON u.id=m.registrato_da_user_id
        WHERE m.payment_id=? ORDER BY m.id DESC
    """, (pid,))
    if not mov.empty:
        mov["importo_pagato"] = mov["importo_pagato"].apply(money)
    st.dataframe(mov, use_container_width=True, hide_index=True)


# ============================================================
# WORK, DOCS, FEEDBACK
# ============================================================

def render_work_form(cid: int | None = None):
    tid = current_tenant_id()
    if cid is None:
        cid = select_client("Cliente lavoro", "work_client")
        if not cid:
            return
    contract_id = select_contract("Contratto collegato", f"work_contract_{cid}", include_all=True)
    with st.form(f"work_form_{cid}"):
        c1, c2, c3, c4 = st.columns(4)
        data_lav = c1.date_input("Data lavoro", value=date.today())
        use_time = c2.checkbox("Aggiungi orario")
        ora = c2.time_input("Orario", value=datetime.now().time().replace(second=0, microsecond=0)) if use_time else None
        tipo = c3.selectbox("Tipo lavoro", ["Sito web","Social","ADS","Grafica","Marketplace","Consulenza","Documenti","Video call","Altro"])
        stato = c4.selectbox("Stato", ["Da fare","In lavorazione","Completato","Consegnato","In attesa cliente","Bloccato"])
        titolo = st.text_input("Titolo lavoro")
        desc = st.text_area("Descrizione lavoro")
        note = st.text_area("Note interne")
        allegato = st.file_uploader("Allega PDF/foto/documento", type=["pdf","png","jpg","jpeg","docx","xlsx","csv"])
        if st.form_submit_button("Salva lavoro"):
            rel = save_upload(allegato, tid, "lavoro_") if allegato else None
            execute("""
                INSERT INTO work_logs (tenant_id,client_id,contract_id,user_id,data_lavoro,ora_lavoro,tipo_lavoro,titolo,descrizione,stato,allegato_file,note_interne,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (tid, cid, contract_id, int(current_user()["id"]), data_lav.isoformat(), ora.strftime("%H:%M") if ora else "", tipo, titolo, desc, stato, rel, note, now_iso()))
            st.success("Lavoro salvato.")
            st.rerun()

def page_work():
    header()
    section("🛠️", "Lavori", "Registro operativo dei lavori per cliente, con data e orario facoltativo.")
    render_work_form()
    st.divider()
    tid = current_tenant_id()
    df = read_df("""
        SELECT w.id, cl.ragione_sociale cliente, w.data_lavoro, w.ora_lavoro, w.tipo_lavoro, w.titolo, w.descrizione, w.stato, u.nome operatore, w.allegato_file
        FROM work_logs w JOIN clients cl ON cl.id=w.client_id LEFT JOIN users u ON u.id=w.user_id
        WHERE w.tenant_id=? ORDER BY w.data_lavoro DESC, w.id DESC
    """, (tid,))
    st.dataframe(df, use_container_width=True, hide_index=True)

def render_documents(cid: int | None = None):
    tid = current_tenant_id()
    if cid is None:
        cid = select_client("Cliente documento", "doc_client")
        if not cid:
            return
    contract_id = select_contract("Contratto collegato", f"doc_contract_{cid}", include_all=True)
    with st.form(f"doc_form_{cid}"):
        tipo = st.selectbox("Tipo documento", ["Contratto","Pagamento","Lavoro","Fattura","Feedback","Altro"])
        titolo = st.text_input("Titolo documento")
        file = st.file_uploader("Allega file", type=["pdf","png","jpg","jpeg","docx","xlsx","csv","txt"])
        note = st.text_area("Note")
        if st.form_submit_button("Salva documento"):
            rel = save_upload(file, tid, "documento_") if file else None
            execute("""
                INSERT INTO documents (tenant_id,client_id,contract_id,tipo,titolo,file_path,note,uploaded_by_user_id,created_at)
                VALUES (?,?,?,?,?,?,?,?,?)
            """, (tid, cid, contract_id, tipo, titolo, rel, note, int(current_user()["id"]), now_iso()))
            st.success("Documento salvato.")
            st.rerun()
    df = read_df("SELECT * FROM documents WHERE tenant_id=? AND client_id=? ORDER BY id DESC", (tid, cid))
    st.dataframe(df, use_container_width=True, hide_index=True)

def page_documents():
    header()
    section("📎", "Documenti", "Archivio documentale per cliente e contratto.")
    render_documents()

def render_feedback(cid: int | None = None):
    tid = current_tenant_id()
    if cid is None:
        cid = select_client("Cliente feedback", "feedback_client")
        if not cid:
            return
    contract_id = select_contract("Contratto collegato", f"feedback_contract_{cid}", include_all=True)
    with st.form(f"feedback_form_{cid}"):
        c1,c2,c3 = st.columns(3)
        data_fb = c1.date_input("Data feedback", value=date.today())
        provenienza = c2.selectbox("Provenienza", ["Cliente","WhatsApp","Email","Telefonata","Video call","Interno"])
        val = c3.selectbox("Valutazione", ["Non indicata",1,2,3,4,5])
        testo = st.text_area("Testo feedback")
        file = st.file_uploader("Allega file", type=["pdf","png","jpg","jpeg","docx","txt"])
        note = st.text_area("Note")
        if st.form_submit_button("Salva feedback"):
            rel = save_upload(file, tid, "feedback_") if file else None
            execute("""
                INSERT INTO feedback (tenant_id,client_id,contract_id,user_id,data_feedback,provenienza,valutazione,testo_feedback,allegato_file,note,created_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)
            """, (tid, cid, contract_id, int(current_user()["id"]), data_fb.isoformat(), provenienza, None if val=="Non indicata" else int(val), testo, rel, note, now_iso()))
            st.success("Feedback salvato.")
            st.rerun()
    df = read_df("SELECT * FROM feedback WHERE tenant_id=? AND client_id=? ORDER BY id DESC", (tid, cid))
    st.dataframe(df, use_container_width=True, hide_index=True)

def page_feedback():
    header()
    section("💬", "Feedback", "Feedback ricevuti dai clienti e note interne.")
    render_feedback()


# ============================================================
# INVOICES
# ============================================================

def generate_invoice_pdf(invoice_id: int) -> str | None:
    if canvas is None:
        return None
    df = read_df("""
        SELECT i.*, cl.ragione_sociale cliente, cl.partita_iva, cl.sede_legale, co.nome azienda, co.piva azienda_piva, co.sede azienda_sede
        FROM invoices i
        JOIN clients cl ON cl.id=i.client_id
        LEFT JOIN companies co ON co.id=i.company_id
        WHERE i.id=?
    """, (invoice_id,))
    if df.empty:
        return None
    r = df.iloc[0].to_dict()
    folder = tenant_upload_dir(int(r["tenant_id"]))
    path = folder / f"fattura_{slug_filename(r['numero'])}_{invoice_id}.pdf"
    c = canvas.Canvas(str(path), pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 16)
    c.drawString(2*cm, h-2*cm, f"Fattura {r['numero']}")
    c.setFont("Helvetica", 10)
    c.drawString(2*cm, h-3*cm, f"Emittente: {r.get('azienda') or '-'} - P.IVA {r.get('azienda_piva') or '-'}")
    c.drawString(2*cm, h-3.6*cm, f"Cliente: {r.get('cliente')} - P.IVA {r.get('partita_iva') or '-'}")
    c.drawString(2*cm, h-4.2*cm, f"Data: {r['data_fattura']} - Scadenza: {r.get('scadenza') or '-'}")
    c.line(2*cm, h-5*cm, w-2*cm, h-5*cm)
    c.drawString(2*cm, h-6*cm, "Descrizione")
    c.drawString(12*cm, h-6*cm, "Imponibile")
    c.drawString(15*cm, h-6*cm, "IVA")
    c.drawString(17*cm, h-6*cm, "Totale")
    c.setFont("Helvetica", 9)
    c.drawString(2*cm, h-6.8*cm, str(r.get("descrizione") or "")[:80])
    c.drawString(12*cm, h-6.8*cm, money(r["imponibile"]))
    c.drawString(15*cm, h-6.8*cm, money(r["iva"]))
    c.drawString(17*cm, h-6.8*cm, money(r["totale"]))
    c.setFont("Helvetica-Bold", 13)
    c.drawRightString(w-2*cm, h-8.5*cm, f"Totale fattura: {money(r['totale'])}")
    c.showPage()
    c.save()
    try:
        return str(path.relative_to(BASE_DIR))
    except Exception:
        return str(path)

def page_invoices():
    header()
    section("🧾", "Fatture interne", "Emissione interna PDF e tracciamento operatore.")
    if not can_finance():
        st.error("Accesso non consentito.")
        return
    tid = current_tenant_id()
    tab_new, tab_list = st.tabs(["Nuova fattura", "Archivio"])
    with tab_new:
        cid = select_client("Cliente fattura", "invoice_client")
        if not cid:
            return
        contract_id = select_contract("Contratto", "invoice_contract", include_all=True)
        company = read_df("SELECT * FROM companies WHERE tenant_id=? AND is_default=1 ORDER BY id DESC LIMIT 1", (tid,))
        company_id = None if company.empty else int(company.iloc[0]["id"])
        with st.form("invoice_form"):
            c1,c2,c3 = st.columns(3)
            numero = c1.text_input("Numero", f"{next_invoice_number(tid)}/{date.today().year}")
            data_f = c2.date_input("Data fattura", value=date.today())
            scad = c3.date_input("Scadenza", value=date.today()+timedelta(days=30))
            desc = st.text_area("Descrizione", "Canone/servizi professionali come da contratto.")
            imponibile = st.number_input("Imponibile", min_value=0.0, value=0.0)
            iva_pct = st.number_input("IVA %", min_value=0.0, value=22.0)
            stato = st.selectbox("Stato", ["Bozza","Pronta","Inviata","Pagata","Annullata"])
            note = st.text_area("Note")
            if st.form_submit_button("Crea fattura interna e PDF"):
                iva = imponibile * iva_pct / 100
                totale = imponibile + iva
                fid = execute("""
                    INSERT INTO invoices (tenant_id,company_id,client_id,contract_id,numero,anno,data_fattura,scadenza,descrizione,imponibile,iva_percentuale,iva,totale,stato,note,emessa_da_user_id,created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
                """, (tid, company_id, cid, contract_id, numero, date.today().year, data_f.isoformat(), scad.isoformat(), desc, imponibile, iva_pct, iva, totale, stato, note, int(current_user()["id"]), now_iso()))
                pdf = generate_invoice_pdf(fid)
                if pdf:
                    execute("UPDATE invoices SET file_pdf=? WHERE id=? AND tenant_id=?", (pdf, fid, tid))
                st.success("Fattura creata.")
                st.rerun()
    with tab_list:
        df = read_df("""
            SELECT i.id,i.numero,i.data_fattura,cl.ragione_sociale cliente,u.nome emessa_da,i.descrizione,i.imponibile,i.iva,i.totale,i.stato,i.file_pdf
            FROM invoices i JOIN clients cl ON cl.id=i.client_id LEFT JOIN users u ON u.id=i.emessa_da_user_id
            WHERE i.tenant_id=? ORDER BY i.id DESC
        """, (tid,))
        if not df.empty:
            for col in ["imponibile","iva","totale"]:
                df[col] = df[col].apply(money)
        st.dataframe(df, use_container_width=True, hide_index=True)

def next_invoice_number(tid: int) -> int:
    df = read_df("SELECT COUNT(*) n FROM invoices WHERE tenant_id=? AND anno=?", (tid, date.today().year))
    return int(df.iloc[0]["n"] or 0) + 1




# ============================================================
# SUPER ADMIN DATA CONTROL
# ============================================================

def page_super_data():
    header()
    section("🧩", "Dati SaaS", "Vista completa Super Admin su clienti, contratti, pagamenti, fatture, lavori, documenti e feedback.")
    if not is_super_admin():
        st.error("Accesso riservato al Super Admin SaaS.")
        return

    tab_clients, tab_contracts, tab_payments, tab_invoices, tab_work, tab_docs, tab_feedback = st.tabs([
        "Clienti", "Contratti", "Pagamenti", "Fatture", "Lavori", "Documenti", "Feedback"
    ])

    with tab_clients:
        df = read_df("""
            SELECT c.*, t.ragione_sociale tenant
            FROM clients c JOIN tenants t ON t.id=c.tenant_id
            ORDER BY c.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_contracts:
        df = read_df("""
            SELECT c.*, t.ragione_sociale tenant, cl.ragione_sociale cliente
            FROM contracts c
            JOIN tenants t ON t.id=c.tenant_id
            JOIN clients cl ON cl.id=c.client_id
            ORDER BY c.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_payments:
        df = read_df("""
            SELECT p.*, t.ragione_sociale tenant, cl.ragione_sociale cliente
            FROM payments p
            JOIN tenants t ON t.id=p.tenant_id
            JOIN clients cl ON cl.id=p.client_id
            ORDER BY p.data_scadenza DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_invoices:
        df = read_df("""
            SELECT i.*, t.ragione_sociale tenant, cl.ragione_sociale cliente
            FROM invoices i
            JOIN tenants t ON t.id=i.tenant_id
            JOIN clients cl ON cl.id=i.client_id
            ORDER BY i.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_work:
        df = read_df("""
            SELECT w.*, t.ragione_sociale tenant, cl.ragione_sociale cliente
            FROM work_logs w
            JOIN tenants t ON t.id=w.tenant_id
            JOIN clients cl ON cl.id=w.client_id
            ORDER BY w.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_docs:
        df = read_df("""
            SELECT d.*, t.ragione_sociale tenant
            FROM documents d JOIN tenants t ON t.id=d.tenant_id
            ORDER BY d.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_feedback:
        df = read_df("""
            SELECT f.*, t.ragione_sociale tenant, cl.ragione_sociale cliente
            FROM feedback f
            JOIN tenants t ON t.id=f.tenant_id
            JOIN clients cl ON cl.id=f.client_id
            ORDER BY f.id DESC
        """)
        st.dataframe(df, use_container_width=True, hide_index=True)

# ============================================================
# MAIN
# ============================================================

def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📘", layout="wide")
    init_db()
    css()
    if not current_user():
        login_page()
        return
    page = sidebar()
    if page == "super_dashboard": page_super_dashboard()
    elif page == "tenants": page_tenants()
    elif page == "global_users": page_global_users()
    elif page == "plans": page_plans()
    elif page == "super_data": page_super_data()
    elif page == "status": page_status()
    elif page == "dashboard": page_dashboard()
    elif page == "company": page_company()
    elif page == "staff": page_staff()
    elif page == "clients": page_clients()
    elif page == "contracts": page_contracts()
    elif page == "templates": page_templates()
    elif page == "payments": page_payments()
    elif page == "work": page_work()
    elif page == "documents": page_documents()
    elif page == "feedback": page_feedback()
    elif page == "invoices": page_invoices()

if __name__ == "__main__":
    main()
