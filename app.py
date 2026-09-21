import streamlit as st
import pandas as pd
import datetime
import os
import sqlite3

st.set_page_config(page_title="Gestionale Scenografia", page_icon="🎬", layout="wide")

# ==============================================================================
# DATABASE MANAGER
# ==============================================================================
DB_PATH = "gestionale_scenografia.db"

def get_connection():
    return sqlite3.connect(DB_PATH)

def init_db():
    with get_connection() as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS clienti (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, telefono TEXT, email TEXT, pec TEXT, indirizzo TEXT, piva TEXT, note TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS materiali (id INTEGER PRIMARY KEY AUTOINCREMENT, codice TEXT UNIQUE, nome TEXT NOT NULL, prezzo REAL NOT NULL, unita TEXT, sfrido REAL DEFAULT 10.0)''')
        c.execute('''CREATE TABLE IF NOT EXISTS lavorazioni (id INTEGER PRIMARY KEY AUTOINCREMENT, codice TEXT UNIQUE, nome TEXT NOT NULL, costo_orario REAL NOT NULL, unita TEXT, ore_um REAL DEFAULT 1.0)''')
        c.execute('''CREATE TABLE IF NOT EXISTS impostazioni (chiave TEXT PRIMARY KEY, valore REAL NOT NULL)''')
        c.execute('''CREATE TABLE IF NOT EXISTS preventivi (id INTEGER PRIMARY KEY AUTOINCREMENT, titolo TEXT NOT NULL, cliente TEXT NOT NULL, costo_diretto REAL, imponibile REAL, iva REAL, ivato REAL, km_trasporto REAL, costo_trasporto REAL, stato TEXT DEFAULT 'Bozza', revisione INTEGER DEFAULT 0, padre_id INTEGER DEFAULT 0, data_inizio TEXT, data_consegna TEXT, data_creazione DATETIME DEFAULT CURRENT_TIMESTAMP)''')
        c.execute('''CREATE TABLE IF NOT EXISTS voci_preventivo (id INTEGER PRIMARY KEY AUTOINCREMENT, preventivo_id INTEGER, tipo TEXT, nome TEXT, qta REAL, um TEXT, ore REAL, costo_base REAL, prezzo_vendita REAL)''')
        conn.commit()
    
    # Default settings
    defaults = {"iva": 22.0, "sfrido_generale": 10.0, "ricarico_materiali": 35.0, "ricarico_manodopera": 25.0, "spese_generali": 20.0, "markup_finale": 30.0, "costo_km": 0.90, "paga_falegname": 40.0}
    with get_connection() as conn:
        c = conn.cursor()
        for k, v in defaults.items():
            c.execute("INSERT OR IGNORE INTO impostazioni (chiave, valore) VALUES (?, ?)", (k, v))
        conn.commit()

init_db()

def get_imp():
    with get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT chiave, valore FROM impostazioni")
        return dict(c.fetchall())

# ==============================================================================
# INTERFACCIA STREAMLIT (TABS)
# ==============================================================================
st.title("🎬 Gestionale Preventivi Scenografici")

tab_dash, tab_prev, tab_cli, tab_mat, tab_lav, tab_imp = st.tabs(["📊 Dashboard", "📋 Preventivi", "👥 Clienti", "📦 Materiali", "🛠️ Lavorazioni", "⚙️ Impostazioni"])

# --- TAB DASHBOARD ---
with tab_dash:
    st.subheader("Dashboard & KPI Aziendali")
    with get_connection() as conn:
        df_prev = pd.read_sql("SELECT * FROM preventivi", conn)
    
    if not df_prev.empty:
        tot_p = len(df_prev)
        sum_imp = df_prev["imponibile"].sum()
        sum_costo = df_prev["costo_diretto"].sum()
        app_p = len(df_prev[df_prev["stato"] == "Approvato"])
        tasso = (app_p / tot_p * 100) if tot_p > 0 else 0
        margine = ((sum_imp - sum_costo) / sum_imp * 100) if sum_imp > 0 else 0
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Preventivi Totali", tot_p)
        c2.metric("Tasso Conversione", f"{tasso:.1f}%")
        c3.metric("Fatturato Pipeline", f"€ {sum_imp:,.2f}")
        c4.metric("Margine Medio Reale", f"{margine:.1f}%")
        
        st.markdown("### 🗓️ Prossimi Cantieri in Consegna")
        st.dataframe(df_prev[["id", "titolo", "cliente", "data_consegna", "stato", "ivato"]], use_container_width=True)
    else:
        st.info("Nessun preventivo inserito. Vai nella sezione Preventivi per crearne uno!")

# --- TAB CLIENTI ---
with tab_cli:
    st.subheader("Anagrafica Clienti")
    with st.expander("➕ Aggiungi Nuovo Cliente"):
        with st.form("form_cliente"):
            cnome = st.text_input("Nome Cliente*")
            ctel = st.text_input("Telefono")
            cemail = st.text_input("Email")
            cpec = st.text_input("PEC")
            cindirizzo = st.text_input("Indirizzo")
            cpiva = st.text_input("Partita IVA")
            cnote = st.text_area("Note")
            submitted = st.form_submit_button("Salva Cliente")
            if submitted and cnome:
                with get_connection() as conn:
                    conn.execute("INSERT INTO clienti (nome, telefono, email, pec, indirizzo, piva, note) VALUES (?,?,?,?,?,?,?)", (cnome, ctel, cemail, cpec, cindirizzo, cpiva, cnote))
                    conn.commit()
                st.success("Cliente aggiunto con successo!")
                st.rerun()

    with get_connection() as conn:
        df_cli = pd.read_sql("SELECT * FROM clienti", conn)
    st.dataframe(df_cli, use_container_width=True)

# --- TAB MATERIALI ---
with tab_mat:
    st.subheader("Listino Materiali")
    with st.expander("➕ Aggiungi Nuovo Materiale"):
        with st.form("form_mat"):
            m_cod = st.text_input("Codice", "MAT-001")
            m_nome = st.text_input("Nome Materiale*")
            m_prezzo = st.number_input("Prezzo Acquisto (€)", min_value=0.0, value=10.0)
            m_unita = st.text_input("U.M. (es. pz, m, kg)", "pz")
            m_sfrido = st.number_input("Sfrido %", value=10.0)
            if st.form_submit_button("Salva Materiale") and m_nome:
                with get_connection() as conn:
                    conn.execute("INSERT OR REPLACE INTO materiali (codice, nome, prezzo, unita, sfrido) VALUES (?,?,?,?,?)", (m_cod, m_nome, m_prezzo, m_unita, m_sfrido))
                    conn.commit()
                st.success("Materiale salvato!")
                st.rerun()
    with get_connection() as conn:
        st.dataframe(pd.read_sql("SELECT * FROM materiali", conn), use_container_width=True)

# --- TAB LAVORAZIONI ---
with tab_lav:
    st.subheader("Listino Lavorazioni Laboratorio")
    with st.expander("➕ Aggiungi Nuova Lavorazione"):
        with st.form("form_lav"):
            l_cod = st.text_input("Codice", "LAV-001")
            l_nome = st.text_input("Nome Lavorazione*")
            l_costo = st.number_input("Costo Orario (€/h)", min_value=0.0, value=40.0)
            l_unita = st.text_input("U.M.", "m²")
            l_ore = st.number_input("Ore per U.M.", value=1.0)
            if st.form_submit_button("Salva Lavorazione") and l_nome:
                with get_connection() as conn:
                    conn.execute("INSERT OR REPLACE INTO lavorazioni (codice, nome, costo_orario, unita, ore_um) VALUES (?,?,?,?,?)", (l_cod, l_nome, l_costo, l_unita, l_ore))
                    conn.commit()
                st.success("Lavorazione salvata!")
                st.rerun()
    with get_connection() as conn:
        st.dataframe(pd.read_sql("SELECT * FROM lavorazioni", conn), use_container_width=True)

# --- TAB IMPOSTAZIONI ---
with tab_imp:
    st.subheader("Parametri Generali e Ricarichi")
    imp = get_imp()
    with st.form("form_imp"):
        nuovi_valori = {}
        for k, v in imp.items():
            nuovi_valori[k] = st.number_input(k.replace("_", " ").capitalize(), value=float(v))
        if st.form_submit_button("Salva Impostazioni"):
            with get_connection() as conn:
                for k, v in nuovi_valori.items():
                    conn.execute("UPDATE impostazioni SET valore=? WHERE chiave=?", (v, k))
                conn.commit()
            st.success("Impostazioni aggiornate!")

# --- TAB PREVENTIVI ---
with tab_prev:
    st.subheader("Gestione Preventivi")
    with get_connection() as conn:
        df_p = pd.read_sql("SELECT id, titolo, cliente, imponibile, ivato, stato, data_consegna FROM preventivi", conn)
    if not df_p.empty:
        st.dataframe(df_p, use_container_width=True)
    else:
        st.info("Nessun preventivo presente.")
