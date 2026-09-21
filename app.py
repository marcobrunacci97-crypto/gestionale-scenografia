import streamlit as st
import pandas as pd
import datetime
import sqlite3

st.set_page_config(page_title="Gestionale Preventivi Scenografici", page_icon="🎬", layout="wide")

DB_PATH = "gestionale_scenografia_v2.db"

def get_connection():
    return sqlite3.connect(DB_PATH)

def init_db():
    with get_connection() as conn:
        c = conn.cursor()
        c.execute('''CREATE TABLE IF NOT EXISTS clienti (id INTEGER PRIMARY KEY AUTOINCREMENT, nome TEXT NOT NULL, telefono TEXT, email TEXT, pec TEXT, indirizzo TEXT, piva TEXT, note TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS materiali (id INTEGER PRIMARY KEY AUTOINCREMENT, codice TEXT UNIQUE, nome TEXT NOT NULL, prezzo REAL NOT NULL, unita TEXT, sfrido REAL DEFAULT 10.0)''')
        c.execute('''CREATE TABLE IF NOT EXISTS lavorazioni (id INTEGER PRIMARY KEY AUTOINCREMENT, codice TEXT UNIQUE, nome TEXT NOT NULL, costo_orario REAL NOT NULL, unita TEXT, ore_um REAL DEFAULT 1.0)''')
        c.execute('''CREATE TABLE IF NOT EXISTS impostazioni (chiave TEXT PRIMARY KEY, valore REAL NOT NULL, unita TEXT)''')
        c.execute('''CREATE TABLE IF NOT EXISTS preventivi (id INTEGER PRIMARY KEY AUTOINCREMENT, titolo TEXT NOT NULL, cliente TEXT NOT NULL, costo_diretto REAL, imponibile REAL, iva REAL, ivato REAL, km_trasporto REAL, costo_trasporto REAL, stato TEXT DEFAULT 'Bozza', revisione INTEGER DEFAULT 0, padre_id INTEGER DEFAULT 0, data_inizio TEXT, data_consegna TEXT, data_creazione DATETIME DEFAULT CURRENT_TIMESTAMP)''')
        c.execute('''CREATE TABLE IF NOT EXISTS voci_preventivo (id INTEGER PRIMARY KEY AUTOINCREMENT, preventivo_id INTEGER, tipo TEXT, nome TEXT, qta REAL, um TEXT, ore REAL, costo_base REAL, prezzo_vendita REAL)''')
        conn.commit()
    
    defaults = {
        "iva": (22.0, "%"), 
        "sfrido_generale": (10.0, "%"), 
        "ricarico_materiali": (35.0, "%"), 
        "ricarico_manodopera": (25.0, "%"),
        "spese_generali": (20.0, "%"), 
        "markup_finale": (30.0, "%"), 
        "costo_km": (0.90, "€/km"), 
        "paga_falegname": (40.0, "€/h"), 
        "paga_fabbro": (45.0, "€/h"),
        "paga_pittore": (40.0, "€/h"), 
        "paga_elettricista": (40.0, "€/h"),
        "paga_amministrazione": (25.0, "€/h"), 
        "paga_titolare": (40.0, "€/h")
    }
    
    with get_connection() as conn:
        c = conn.cursor()
        for k, v in defaults.items():
            c.execute("INSERT OR IGNORE INTO impostazioni (chiave, valore, unita) VALUES (?, ?, ?)", (k, v[0], v[1]))
        
        c.execute("SELECT COUNT(*) FROM clienti")
        if c.fetchone()[0] == 0:
            c.execute("""
                INSERT INTO clienti (nome, telefono, email, pec, indirizzo, piva, note) 
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (
                "Cinecittà S.p.A.", 
                "+39 06 722911", 
                "info@cinecittaspa.it", 
                "cinecittaspa@legalmail.it", 
                "Via Tuscolana 1055, 00173 Roma (RM)", 
                "06385451003", 
                "Studi di Cinecittà - Cliente principale pre-caricato"
            ))

        c.execute("SELECT COUNT(*) FROM materiali")
        if c.fetchone()[0] == 0:
            mat_iniziali = [
                ("MAT-001", "Pannello multistrato pioppo 10mm", 18.50, "m²", 10.0),
                ("MAT-002", "Listello abete 20x20mm", 2.20, "m", 15.0),
                ("MAT-003", "Pannello MDF 19mm", 28.00, "m²", 10.0),
                ("MAT-004", "Viti truciolari 4x40 (confezione)", 12.50, "pz", 5.0)
            ]
            c.executemany("INSERT INTO materiali (codice, nome, prezzo, unita, sfrido) VALUES (?, ?, ?, ?, ?)", mat_iniziali)

        c.execute("SELECT COUNT(*) FROM lavorazioni")
        if c.fetchone()[0] == 0:
            lav_iniziali = [
                ("LAV-001", "Taglio e squadratura legnami", 40.0, "m²", 1.0),
                ("LAV-002", "Assemblaggio struttura scenografica", 40.0, "h", 1.0),
                ("LAV-003", "Stuccatura e levigatura", 40.0, "m²", 1.5),
                ("LAV-004", "Decorazione e pittura scenografica", 45.0, "m²", 1.2)
            ]
            c.executemany("INSERT INTO lavorazioni (codice, nome, costo_orario, unita, ore_um) VALUES (?, ?, ?, ?, ?)", lav_iniziali)
            
        conn.commit()

init_db()

def get_imp_dict():
    with get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT chiave, valore FROM impostazioni")
        return dict(c.fetchall())

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
        st.info("Nessun preventivo inserito.")

# --- TAB PREVENTIVI ---
with tab_prev:
    st.subheader("Gestione Preventivi")
    
    with st.expander("➕ Crea Nuovo Preventivo"):
        with get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome FROM clienti")
            clienti = [r[0] for r in c.fetchall()]
            c.execute("SELECT nome, prezzo, unita, sfrido FROM materiali")
            l_mat = c.fetchall()
            c.execute("SELECT nome, costo_orario, unita, ore_um FROM lavorazioni")
            l_lav = c.fetchall()

        if not clienti:
            st.warning("⚠️ Inserisci prima almeno un cliente nella sezione 'Clienti'.")
        else:
            p_titolo = st.text_input("Titolo Progetto*")
            p_cliente = st.selectbox("Cliente*", clienti)
            
            c_i1, c_i2 = st.columns(2)
            p_inizio = c_i1.text_input("Inizio Lavori", datetime.date.today().strftime("%d/%m/%Y"))
            p_consegna = c_i2.text_input("Consegna", (datetime.date.today() + datetime.timedelta(days=15)).strftime("%d/%m/%Y"))

            st.markdown("#### Voci del Preventivo")
            if "voci_temp" not in st.session_state:
                st.session_state.voci_temp = []

            col_m1, col_m2, col_m3 = st.columns([3, 1, 1])
            sel_mat = col_m1.selectbox("Seleziona Materiale", [m[0] for m in l_mat] if l_mat else ["Nessun materiale"])
            qta_mat = col_m2.number_input("Q.tà Mat.", min_value=0.1, value=1.0)
            if col_m3.button("Aggiungi Materiale") and l_mat:
                m_obj = next((m for m in l_mat if m[0] == sel_mat), None)
                if m_obj:
                    costo = m_obj[1] * qta_mat * (1 + m_obj[3]/100.0)
                    imp_val = get_imp_dict()
                    vendita = costo * (1 + imp_val.get("ricarico_materiali", 35)/100.0)
                    st.session_state.voci_temp.append({"tipo": "Materiale", "nome": m_obj[0], "qta": qta_mat, "um": m_obj[2], "ore": 0.0, "costo_base": costo, "prezzo": vendita})
                    st.success(f"Aggiunto: {m_obj[0]}")

            col_l1, col_l2, col_l3, col_l4 = st.columns([2, 1, 1, 1])
            sel_lav = col_l1.selectbox("Seleziona Lavorazione", [l[0] for l in l_lav] if l_lav else ["Nessuna lavorazione"])
            qta_lav = col_l2.number_input("Q.tà Lav.", min_value=0.1, value=1.0)
            ore_lav = col_l3.number_input("Ore stimate", min_value=0.1, value=1.0)
            if col_l4.button("Aggiungi Lav.") and l_lav:
                l_obj = next((l for l in l_lav if l[0] == sel_lav), None)
                if l_obj:
                    costo = ore_lav * l_obj[1]
                    imp_val = get_imp_dict()
                    vendita = costo * (1 + imp_val.get("ricarico_manodopera", 25)/100.0)
                    st.session_state.voci_temp.append({"tipo": "Lavorazione", "nome": l_obj[0], "qta": qta_lav, "um": l_obj[2], "ore": ore_lav, "costo_base": costo, "prezzo": vendita})
                    st.success(f"Aggiunta lavorazione: {l_obj[0]}")

            if st.session_state.voci_temp:
                st.dataframe(pd.DataFrame(st.session_state.voci_temp), use_container_width=True)
                if st.button("🗑️ Svuota Voci"):
                    st.session_state.voci_temp = []
                    st.rerun()

            st.markdown("#### Cantiere & Logistica")
            c_ext1, c_ext2 = st.columns(2)
            ore_falegname = c_ext1.number_input("Ore Falegname in Cantiere", min_value=0.0, value=0.0)
            km_trasporto = c_ext2.number_input("Km Trasporto A/R", min_value=0.0, value=0.0)

            if st.button("💾 Calcola e Salva Preventivo Definitivo"):
                if not p_titolo:
                    st.error("Inserisci il titolo del progetto.")
                else:
                    imp = get_imp_dict()
                    voci = st.session_state.voci_temp
                    
                    c_mat = sum(v["costo_base"] for v in voci if v["tipo"] == "Materiale")
                    v_mat = sum(v["prezzo"] for v in voci if v["tipo"] == "Materiale")
                    c_lav = sum(v["costo_base"] for v in voci if v["tipo"] == "Lavorazione")
                    v_lav = sum(v["prezzo"] for v in voci if v["tipo"] == "Lavorazione")

                    c_fal = ore_falegname * imp.get("paga_falegname", 40)
                    v_fal = c_fal * (1 + imp.get("ricarico_manodopera", 25)/100.0)

                    c_tr = km_trasporto * imp.get("costo_km", 0.90)
                    markup = 1 + (imp.get("markup_finale", 30)/100.0)
                    v_tr = c_tr * markup

                    c_dir = c_mat + c_lav + c_fal + c_tr
                    c_aziendale = c_dir * (1 + imp.get("spese_generali", 20)/100.0)

                    v_base = v_mat + v_lav + v_fal + v_tr
                    imponibile = v_base * markup
                    iva = imponibile * (imp.get("iva", 22)/100.0)
                    ivato = imponibile + iva

                    with get_connection() as conn:
                        cursor = conn.cursor()
                        cursor.execute("""
                            INSERT INTO preventivi (titolo, cliente, costo_diretto, imponibile, iva, ivato, km_trasporto, costo_trasporto, stato, data_inizio, data_consegna)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Bozza', ?, ?)
                        """, (p_titolo, p_cliente, c_aziendale, imponibile, iva, ivato, km_trasporto, c_tr, p_inizio, p_consegna))
                        pid = cursor.lastrowid

                        for v in voci:
                            cursor.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                                      (pid, v["tipo"], v["nome"], v["qta"], v["um"], v["ore"], v["costo_base"], v["prezzo"]))
                        if ore_falegname > 0:
                            cursor.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, 'Installazione / Falegname', ?, 'h', ?, ?, ?)",
                                      (pid, "Manodopera Cantiere", ore_falegname, ore_falegname, c_fal, v_fal * markup))
                        if km_trasporto > 0:
                            cursor.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, 'Trasporto e Logistica A/R', ?, 'km', 0, ?, ?)",
                                      (pid, "Logistica", km_trasporto, c_tr, v_tr))
                        conn.commit()

                    st.session_state.voci_temp = []
                    st.success(f"Preventivo #{pid} salvato con successo!")
                    st.rerun()

    with get_connection() as conn:
        df_p = pd.read_sql("SELECT id, titolo, cliente, imponibile, ivato, stato, data_consegna FROM preventivi ORDER BY id DESC", conn)
    if not df_p.empty:
        st.dataframe(df_p, use_container_width=True)
    else:
        st.info("Nessun preventivo salvato.")

# --- TAB CLIENTI (Modificabile direttamente) ---
with tab_cli:
    st.subheader("Anagrafica Clienti")
    st.info("✏️ Puoi modificare direttamente i campi cliccando sulle celle della tabella sottostante o aggiungere nuove righe in fondo. Clicca il pulsante per salvare.")
    
    with get_connection() as conn:
        df_clienti = pd.read_sql("SELECT * FROM clienti", conn)
    
    edited_clienti = st.data_editor(df_clienti, num_rows="dynamic", key="editor_clienti", use_container_width=True)
    
    if st.button("💾 Salva Modifiche Clienti"):
        with get_connection() as conn:
            conn.execute("DELETE FROM clienti")
            edited_clienti.to_sql("clienti", conn, if_exists="append", index=False)
            conn.commit()
        st.success("Anagrafica clienti aggiornata con successo!")
        st.rerun()

# --- TAB MATERIALI (Modificabile direttamente) ---
with tab_mat:
    st.subheader("Listino Materiali")
    st.info("✏️ Modifica i prezzi, i nomi, le unità di misura o lo sfrido direttamente nella tabella. Puoi anche aggiungere o rimuovere materiali.")
    
    with get_connection() as conn:
        df_mat = pd.read_sql("SELECT * FROM materiali", conn)
        
    edited_mat = st.data_editor(df_mat, num_rows="dynamic", key="editor_mat", use_container_width=True)
    
    if st.button("💾 Salva Modifiche Materiali"):
        with get_connection() as conn:
            conn.execute("DELETE FROM materiali")
            edited_mat.to_sql("materiali", conn, if_exists="append", index=False)
            conn.commit()
        st.success("Listino materiali aggiornato con successo!")
        st.rerun()

# --- TAB LAVORAZIONI (Modificabile direttamente) ---
with tab_lav:
    st.subheader("Listino Lavorazioni Laboratorio")
    st.info("✏️ Modifica i costi orari, le unità di misura o le ore stimate direttamente nella tabella.")
    
    with get_connection() as conn:
        df_lav = pd.read_sql("SELECT * FROM lavorazioni", conn)
        
    edited_lav = st.data_editor(df_lav, num_rows="dynamic", key="editor_lav", use_container_width=True)
    
    if st.button("💾 Salva Modifiche Lavorazioni"):
        with get_connection() as conn:
            conn.execute("DELETE FROM lavorazioni")
            edited_lav.to_sql("lavorazioni", conn, if_exists="append", index=False)
            conn.commit()
        st.success("Listino lavorazioni aggiornato con successo!")
        st.rerun()

# --- TAB IMPOSTAZIONI ---
with tab_imp:
    st.subheader("Parametri Generali e Ricarichi")
    with get_connection() as conn:
        c = conn.cursor()
        c.execute("SELECT chiave, valore, unita FROM impostazioni")
        rows = c.fetchall()
        
    with st.form("form_imp"):
        nuovi_valori = {}
        for r in rows:
             chiave, valore_corrente, um = r[0], r[1], r[2] if len(r) > 2 and r[2] else ""
             etichetta = f"{chiave.replace('_', ' ').capitalize()} ({um})" if um else chiave.replace('_', ' ').capitalize()
             nuovi_valori[chiave] = st.number_input(etichetta, value=float(valore_corrente))
             
        if st.form_submit_button("Salva Impostazioni"):
            with get_connection() as conn:
                for k, v in nuovi_valori.items():
                    conn.execute("UPDATE impostazioni SET valore=? WHERE chiave=?", (v, k))
                conn.commit()
            st.success("Impostazioni aggiornate con successo!")
