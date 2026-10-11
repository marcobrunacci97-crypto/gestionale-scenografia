
import os
import shutil
import sqlite3
import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd

# Importazione di ReportLab per la generazione PDF professionale
try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


# ==============================================================================
# 1. DATABASE MANAGER
# ==============================================================================
class DatabaseManager:
    def __init__(self, db_path="gestionale_scenografia.db"):
        self.db_path = db_path
        self.esegui_backup_automatico()
        self.init_db()

    def get_connection(self):
        return sqlite3.connect(self.db_path)

    def esegui_backup_automatico(self):
        if os.path.exists(self.db_path):
            backup_dir = "backups"
            os.makedirs(backup_dir, exist_ok=True)
            data_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_file = os.path.join(backup_dir, f"gestionale_backup_{data_str}.db")
            try:
                shutil.copy2(self.db_path, backup_file)
            except Exception as e:
                print(f"[BACKUP] Errore: {e}")

    def ripristina_backup(self, filepath_backup):
        try:
            shutil.copy2(filepath_backup, self.db_path)
            return True
        except Exception as e:
            return False

    def init_db(self):
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute('''
                CREATE TABLE IF NOT EXISTS clienti (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    nome TEXT NOT NULL,
                    telefono TEXT,
                    email TEXT,
                    pec TEXT,
                    indirizzo TEXT,
                    piva TEXT,
                    note TEXT
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS materiali (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codice TEXT UNIQUE,
                    nome TEXT NOT NULL,
                    prezzo REAL NOT NULL,
                    unita TEXT,
                    sfrido REAL DEFAULT 10.0
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS lavorazioni (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    codice TEXT UNIQUE,
                    nome TEXT NOT NULL,
                    costo_orario REAL NOT NULL,
                    unita TEXT,
                    ore_um REAL DEFAULT 1.0
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS impostazioni (
                    chiave TEXT PRIMARY KEY,
                    valore REAL NOT NULL,
                    unita TEXT
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS preventivi (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    titolo TEXT NOT NULL,
                    cliente TEXT NOT NULL,
                    costo_diretto REAL,
                    imponibile REAL,
                    iva REAL,
                    ivato REAL,
                    km_trasporto REAL,
                    costo_trasporto REAL,
                    stato TEXT DEFAULT 'Bozza',
                    revisione INTEGER DEFAULT 0,
                    padre_id INTEGER DEFAULT 0,
                    data_inizio TEXT,
                    data_consegna TEXT,
                    data_creazione DATETIME DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS voci_preventivo (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    preventivo_id INTEGER,
                    tipo TEXT,
                    nome TEXT,
                    qta REAL,
                    um TEXT,
                    ore REAL,
                    costo_base REAL,
                    prezzo_vendita REAL,
                    FOREIGN KEY(preventivo_id) REFERENCES preventivi(id) ON DELETE CASCADE
                )
            ''')
            c.execute('''
                CREATE TABLE IF NOT EXISTS operatori_preventivo (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    preventivo_id INTEGER,
                    ruolo TEXT,
                    ore REAL,
                    costo REAL,
                    FOREIGN KEY(preventivo_id) REFERENCES preventivi(id) ON DELETE CASCADE
                )
            ''')

            # Verifica e aggiunta colonna 'unita' se manca
            c.execute("PRAGMA table_info(impostazioni)")
            cols = [col[1] for col in c.fetchall()]
            if "unita" not in cols:
                try:
                    c.execute("ALTER TABLE impostazioni ADD COLUMN unita TEXT")
                except Exception:
                    pass

            conn.commit()
        self.popola_impostazioni_default()

    def popola_impostazioni_default(self):
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
        with self.get_connection() as conn:
            c = conn.cursor()
            for key, (val, um) in defaults.items():
                c.execute("INSERT OR IGNORE INTO impostazioni (chiave, valore, unita) VALUES (?, ?, ?)", (key, val, um))
                c.execute("UPDATE impostazioni SET unita=? WHERE chiave=? AND (unita IS NULL OR unita='')", (um, key))
            conn.commit()

    def carica_da_excel_se_vuoto(self, file_path="Listino_Allestimenti_Scenografici.xlsx"):
        if not os.path.exists(file_path):
            return
        with self.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*) FROM materiali")
            if c.fetchone()[0] == 0:
                try:
                    xls = pd.ExcelFile(file_path)
                    if "Listino_Materiali" in xls.sheet_names:
                        df_mat = pd.read_excel(xls, sheet_name="Listino_Materiali")
                        for _, row in df_mat.iterrows():
                            c.execute(
                                "INSERT OR IGNORE INTO materiali (codice, nome, unita, prezzo, sfrido) VALUES (?, ?, ?, ?, ?)",
                                (str(row.get("Codice", "")), str(row.get("Materiale", "")), str(row.get("U.M.", "pz")),
                                 float(row.get("Costo acquisto (€)", 0.0)), float(row.get("Sfrido %", 0.10)) * 100)
                            )
                    if "Listino_Lavorazioni" in xls.sheet_names:
                        df_lav = pd.read_excel(xls, sheet_name="Listino_Lavorazioni")
                        for _, row in df_lav.iterrows():
                            c.execute(
                                "INSERT OR IGNORE INTO lavorazioni (codice, nome, unita, costo_orario, ore_um) VALUES (?, ?, ?, ?, ?)",
                                (str(row.get("Codice", "")), str(row.get("Lavorazione", "")), str(row.get("U.M.", "m²")),
                                 float(row.get("Costo orario (€)", 40.0)), float(row.get("Ore/U.M.", 1.0)))
                            )
                    conn.commit()
                except Exception as e:
                    print(f"Errore import: {e}")


def parse_float_positivo(valore_str, nome_campo, consenti_zero=True):
    val_clean = valore_str.strip().replace(",", ".")
    if not val_clean:
        raise ValueError(f"Il campo '{nome_campo}' non può essere vuoto.")
    try:
        val = float(val_clean)
    except ValueError:
        raise ValueError(f"Il valore in '{nome_campo}' non è un numero valido.")
    if not consenti_zero and val <= 0:
        raise ValueError(f"Il campo '{nome_campo}' deve essere maggiore di zero.")
    if val < 0:
        raise ValueError(f"Il campo '{nome_campo}' non può essere negativo.")
    return val


# ==============================================================================
# 2. TAB CLIENTI
# ==============================================================================
class TabClienti(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.setup_ui()

    def setup_ui(self):
        frame_top = ttk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_top, text="+ Nuovo Cliente", command=self.modal_nuovo_cliente).pack(side="left", padx=5)
        ttk.Button(frame_top, text="✏️ Modifica", command=self.modal_modifica_cliente).pack(side="left", padx=5)
        ttk.Button(frame_top, text="🗑️ Elimina", command=self.elimina_cliente).pack(side="left", padx=5)

        ttk.Label(frame_top, text="🔍 Cerca:").pack(side="left", padx=(20, 5))
        self.e_ricerca = ttk.Entry(frame_top, width=25)
        self.e_ricerca.pack(side="left", padx=5)
        self.e_ricerca.bind("<KeyRelease>", lambda e: self.aggiorna_tabella())

        cols = ("ID", "Nome", "Telefono", "Email", "PEC", "Indirizzo", "P.IVA", "Note")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100 if col == "ID" else 130)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.aggiorna_tabella()

    def aggiorna_tabella(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        q = f"%{self.e_ricerca.get().strip()}%"
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, nome, telefono, email, pec, indirizzo, piva, note FROM clienti WHERE nome LIKE ? OR email LIKE ? OR piva LIKE ?", (q, q, q))
            for row in c.fetchall():
                self.tree.insert("", "end", values=row)

    def modal_nuovo_cliente(self, cliente_id=None):
        win = tk.Toplevel(self)
        win.title("Gestione Cliente")
        win.geometry("400x480")
        fields = ["Nome (Obbligatorio)*", "Telefono", "Email", "PEC", "Indirizzo", "Partita IVA", "Note"]
        entries = {}
        c_data = {}
        if cliente_id:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT nome, telefono, email, pec, indirizzo, piva, note FROM clienti WHERE id=?", (cliente_id,))
                row = c.fetchone()
                if row:
                    c_data = dict(zip(["nome", "telefono", "email", "pec", "indirizzo", "piva", "note"], row))
        keys_map = {"Nome (Obbligatorio)*": "nome", "Telefono": "telefono", "Email": "email", "PEC": "pec", "Indirizzo": "indirizzo", "Partita IVA": "piva", "Note": "note"}

        for f in fields:
            ttk.Label(win, text=f).pack(anchor="w", padx=20, pady=(5, 0))
            e = ttk.Entry(win, width=40)
            if cliente_id and keys_map[f] in c_data:
                e.insert(0, str(c_data[keys_map[f]] or ""))
            e.pack(padx=20, pady=(0, 5))
            entries[f] = e

        def salva():
            nome = entries["Nome (Obbligatorio)*"].get().strip()
            if not nome:
                messagebox.showerror("Errore", "Nome obbligatorio!")
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if cliente_id:
                    c.execute("UPDATE clienti SET nome=?, telefono=?, email=?, pec=?, indirizzo=?, piva=?, note=? WHERE id=?",
                              (nome, entries["Telefono"].get(), entries["Email"].get(), entries["PEC"].get(), entries["Indirizzo"].get(), entries["Partita IVA"].get(), entries["Note"].get(), cliente_id))
                else:
                    c.execute("INSERT INTO clienti (nome, telefono, email, pec, indirizzo, piva, note) VALUES (?, ?, ?, ?, ?, ?, ?)",
                              (nome, entries["Telefono"].get(), entries["Email"].get(), entries["PEC"].get(), entries["Indirizzo"].get(), entries["Partita IVA"].get(), entries["Note"].get()))
                conn.commit()
            self.aggiorna_tabella()
            win.destroy()

        ttk.Button(win, text="Salva", command=salva).pack(pady=15)

    def modal_modifica_cliente(self):
        sel = self.tree.selection()
        if not sel: return
        self.modal_nuovo_cliente(cliente_id=self.tree.item(sel[0], "values")[0])

    def elimina_cliente(self):
        sel = self.tree.selection()
        if not sel: return
        cid = self.tree.item(sel[0], "values")[0]
        if messagebox.askyesno("Conferma", "Eliminare il cliente?"):
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM clienti WHERE id=?", (cid,))
                conn.commit()
            self.aggiorna_tabella()


# ==============================================================================
# 3. TAB MATERIALI
# ==============================================================================
class TabMateriali(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.setup_ui()

    def setup_ui(self):
        frame_top = ttk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_top, text="+ Nuovo Materiale", command=self.modal_nuovo_materiale).pack(side="left", padx=5)
        ttk.Button(frame_top, text="✏️ Modifica", command=self.modal_modifica_materiale).pack(side="left", padx=5)
        ttk.Button(frame_top, text="🗑️ Elimina", command=self.elimina_materiale).pack(side="left", padx=5)

        ttk.Label(frame_top, text="🔍 Cerca:").pack(side="left", padx=(20, 5))
        self.e_ricerca = ttk.Entry(frame_top, width=25)
        self.e_ricerca.pack(side="left", padx=5)
        self.e_ricerca.bind("<KeyRelease>", lambda e: self.aggiorna_tabella())

        cols = ("ID", "Codice", "Nome Materiale", "Prezzo Acquisto (€)", "U.M.", "Sfrido %")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80 if col == "ID" else 130)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.aggiorna_tabella()

    def aggiorna_tabella(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        q = f"%{self.e_ricerca.get().strip()}%"
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, codice, nome, prezzo, unita, sfrido FROM materiali WHERE nome LIKE ? OR codice LIKE ?", (q, q))
            for row in c.fetchall():
                self.tree.insert("", "end", values=(row[0], row[1], row[2], f"€ {row[3]:.2f}", row[4], f"{row[5]:.0f}%"))

    def modal_nuovo_materiale(self, mat_id=None):
        win = tk.Toplevel(self)
        win.title("Gestione Materiale")
        win.geometry("380x380")
        m_data = {}
        if mat_id:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT codice, nome, prezzo, unita, sfrido FROM materiali WHERE id=?", (mat_id,))
                row = c.fetchone()
                if row: m_data = dict(zip(["codice", "nome", "prezzo", "unita", "sfrido"], row))

        ttk.Label(win, text="Codice").pack(anchor="w", padx=20, pady=(5, 0))
        e_cod = ttk.Entry(win, width=35); e_cod.pack(padx=20); e_cod.insert(0, m_data.get("codice", "MAT-001"))
        ttk.Label(win, text="Nome*").pack(anchor="w", padx=20, pady=(5, 0))
        e_nome = ttk.Entry(win, width=35); e_nome.pack(padx=20); 
        if "nome" in m_data: e_nome.insert(0, m_data["nome"])
        ttk.Label(win, text="Prezzo Acquisto (€)*").pack(anchor="w", padx=20, pady=(5, 0))
        e_prezzo = ttk.Entry(win, width=35); e_prezzo.pack(padx=20); 
        if "prezzo" in m_data: e_prezzo.insert(0, str(m_data["prezzo"]))
        ttk.Label(win, text="U.M. (es. m², m, pz, kg)").pack(anchor="w", padx=20, pady=(5, 0))
        e_unita = ttk.Entry(win, width=35); e_unita.pack(padx=20); e_unita.insert(0, m_data.get("unita", "m²"))
        ttk.Label(win, text="Sfrido %").pack(anchor="w", padx=20, pady=(5, 0))
        e_sfrido = ttk.Entry(win, width=35); e_sfrido.pack(padx=20); e_sfrido.insert(0, str(m_data.get("sfrido", 10.0)))

        def salva():
            try:
                prezzo = parse_float_positivo(e_prezzo.get(), "Prezzo", False)
                sfrido = parse_float_positivo(e_sfrido.get(), "Sfrido")
            except ValueError as e:
                messagebox.showerror("Errore", str(e))
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if mat_id:
                    c.execute("UPDATE materiali SET codice=?, nome=?, prezzo=?, unita=?, sfrido=? WHERE id=?",
                              (e_cod.get(), e_nome.get(), prezzo, e_unita.get(), sfrido, mat_id))
                else:
                    c.execute("INSERT INTO materiali (codice, nome, prezzo, unita, sfrido) VALUES (?, ?, ?, ?, ?)",
                              (e_cod.get(), e_nome.get(), prezzo, e_unita.get(), sfrido))
                conn.commit()
            self.aggiorna_tabella()
            win.destroy()

        ttk.Button(win, text="Salva", command=salva).pack(pady=15)

    def modal_modifica_materiale(self):
        sel = self.tree.selection()
        if not sel: return
        self.modal_nuovo_materiale(mat_id=self.tree.item(sel[0], "values")[0])

    def elimina_materiale(self):
        sel = self.tree.selection()
        if not sel: return
        mid = self.tree.item(sel[0], "values")[0]
        if messagebox.askyesno("Conferma", "Eliminare materiale?"):
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM materiali WHERE id=?", (mid,))
                conn.commit()
            self.aggiorna_tabella()


# ==============================================================================
# 4. TAB LAVORAZIONI
# ==============================================================================
class TabLavorazioni(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.setup_ui()

    def setup_ui(self):
        frame_top = ttk.Frame(self)
        frame_top.pack(fill="x", padx=10, pady=5)
        ttk.Button(frame_top, text="+ Nuova Lavorazione", command=self.modal_nuova_lavorazione).pack(side="left", padx=5)
        ttk.Button(frame_top, text="✏️ Modifica", command=self.modal_modifica_lavorazione).pack(side="left", padx=5)
        ttk.Button(frame_top, text="🗑️ Elimina", command=self.elimina_lavorazione).pack(side="left", padx=5)

        cols = ("ID", "Codice", "Nome Lavorazione", "Costo Orario (€/h)", "U.M.", "Ore/U.M.")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=80 if col == "ID" else 130)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.aggiorna_tabella()

    def aggiorna_tabella(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, codice, nome, costo_orario, unita, ore_um FROM lavorazioni")
            for row in c.fetchall():
                self.tree.insert("", "end", values=(row[0], row[1], row[2], f"€ {row[3]:.2f}", row[4], f"{row[5]:.2f}h"))

    def modal_nuova_lavorazione(self, lav_id=None):
        win = tk.Toplevel(self)
        win.title("Gestione Lavorazione")
        win.geometry("380x380")
        l_data = {}
        if lav_id:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT codice, nome, costo_orario, unita, ore_um FROM lavorazioni WHERE id=?", (lav_id,))
                row = c.fetchone()
                if row: l_data = dict(zip(["codice", "nome", "costo_orario", "unita", "ore_um"], row))

        ttk.Label(win, text="Codice").pack(anchor="w", padx=20, pady=(5,0))
        e_cod = ttk.Entry(win, width=35); e_cod.pack(padx=20); e_cod.insert(0, l_data.get("codice", "LAV-001"))
        ttk.Label(win, text="Nome*").pack(anchor="w", padx=20, pady=(5,0))
        e_nome = ttk.Entry(win, width=35); e_nome.pack(padx=20); 
        if "nome" in l_data: e_nome.insert(0, l_data["nome"])
        ttk.Label(win, text="Costo Orario (€/h)*").pack(anchor="w", padx=20, pady=(5,0))
        e_costo = ttk.Entry(win, width=35); e_costo.pack(padx=20); 
        if "costo_orario" in l_data: e_costo.insert(0, str(l_data["costo_orario"]))
        ttk.Label(win, text="U.M. (es. h, m², pz)").pack(anchor="w", padx=20, pady=(5,0))
        e_unita = ttk.Entry(win, width=35); e_unita.pack(padx=20); e_unita.insert(0, l_data.get("unita", "h"))
        ttk.Label(win, text="Ore per U.M.").pack(anchor="w", padx=20, pady=(5,0))
        e_ore = ttk.Entry(win, width=35); e_ore.pack(padx=20); e_ore.insert(0, str(l_data.get("ore_um", 1.0)))

        def salva():
            try:
                costo = parse_float_positivo(e_costo.get(), "Costo", False)
                ore = parse_float_positivo(e_ore.get(), "Ore", False)
            except ValueError as e:
                messagebox.showerror("Errore", str(e))
                return
            with self.db.get_connection() as conn:
                c = conn.cursor()
                if lav_id:
                    c.execute("UPDATE lavorazioni SET codice=?, nome=?, costo_orario=?, unita=?, ore_um=? WHERE id=?",
                              (e_cod.get(), e_nome.get(), costo, e_unita.get(), ore, lav_id))
                else:
                    c.execute("INSERT INTO lavorazioni (codice, nome, costo_orario, unita, ore_um) VALUES (?, ?, ?, ?, ?)",
                              (e_cod.get(), e_nome.get(), costo, e_unita.get(), ore))
                conn.commit()
            self.aggiorna_tabella()
            win.destroy()

        ttk.Button(win, text="Salva", command=salva).pack(pady=15)

    def modal_modifica_lavorazione(self):
        sel = self.tree.selection()
        if not sel: return
        self.modal_nuova_lavorazione(lav_id=self.tree.item(sel[0], "values")[0])

    def elimina_lavorazione(self):
        sel = self.tree.selection()
        if not sel: return
        lid = self.tree.item(sel[0], "values")[0]
        if messagebox.askyesno("Conferma", "Eliminare lavorazione?"):
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM lavorazioni WHERE id=?", (lid,))
                conn.commit()
            self.aggiorna_tabella()


# ==============================================================================
# 5. TAB IMPOSTAZIONI
# ==============================================================================
class TabImpostazioni(ttk.Frame):
    def __init__(self, parent, db, callback):
        super().__init__(parent)
        self.db = db
        self.callback = callback
        self.setup_ui()

    def setup_ui(self):
        frame = ttk.LabelFrame(self, text="Parametri Generali, Paga Oraria e Ricarichi")
        frame.pack(fill="x", padx=15, pady=10)
        self.entries = {}
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT chiave, valore, unita FROM impostazioni")
            imp = c.fetchall()

        for i, row in enumerate(imp):
            k = row[0]
            v = row[1]
            um = row[2] if len(row) > 2 and row[2] else ""

            label_text = f"{k.replace('_', ' ').capitalize()} ({um})" if um else k.replace('_', ' ').capitalize()
            ttk.Label(frame, text=label_text).grid(row=i, column=0, sticky="w", padx=10, pady=2)
            e = ttk.Entry(frame, width=15)
            e.insert(0, str(v))
            e.grid(row=i, column=1, sticky="w", padx=10, pady=2)
            self.entries[k] = e

        ttk.Button(frame, text="💾 Salva Impostazioni", command=self.salva).grid(row=len(imp), column=1, pady=10, sticky="w")

        f_help = ttk.LabelFrame(self, text="ℹ️ Guida e Formule di Calcolo")
        f_help.pack(fill="x", padx=15, pady=10)

        help_testo = (
            "• Sfrido Materiali: applicato al costo base di ciascun materiale\n"
            "  (Costo Base = Prezzo Acquisto × Q.tà × (1 + Sfrido / 100))\n\n"
            "• Ricarico Materiali: applicato al costo del materiale con sfrido\n"
            "  (Prezzo Vendita Materiale = Costo Base × (1 + Ricarico Materiali / 100))\n\n"
            "• Ricarico Manodopera: applicato al costo orario di lavorazioni di laboratorio e personale di cantiere\n"
            "  (Prezzo Vendita Manodopera = Costo Orario × Ore × (1 + Ricarico Manodopera / 100))"
        )

        lbl_help = ttk.Label(f_help, text=help_testo, font=("Helvetica", 9), justify="left")
        lbl_help.pack(padx=10, pady=8, anchor="w")

    def salva(self):
        try:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                for k, ent in self.entries.items():
                    val = parse_float_positivo(ent.get(), k)
                    c.execute("UPDATE impostazioni SET valore=? WHERE chiave=?", (val, k))
                conn.commit()
            messagebox.showinfo("Salvato", "Impostazioni aggiornate!")
            self.callback()
        except ValueError as e:
            messagebox.showerror("Errore", str(e))


# ==============================================================================
# 6. TAB DASHBOARD
# ==============================================================================
class TabDashboard(ttk.Frame):
    def __init__(self, parent, db):
        super().__init__(parent)
        self.db = db
        self.setup_ui()

    def setup_ui(self):
        ttk.Label(self, text="📊 Dashboard & KPI Aziendali", font=("Helvetica", 14, "bold")).pack(anchor="w", padx=15, pady=10)
        
        self.f_kpi = ttk.Frame(self)
        self.f_kpi.pack(fill="x", padx=15, pady=5)

        f_scad = ttk.LabelFrame(self, text="🗓️ Prossimi Cantieri in Consegna")
        f_scad.pack(fill="both", expand=True, padx=15, pady=10)

        cols = ("ID", "Progetto", "Cliente", "Data Consegna", "Stato", "Importo Ivato")
        self.tree = ttk.Treeview(f_scad, columns=cols, show="headings", height=6)
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120)
        self.tree.pack(fill="both", expand=True, padx=5, pady=5)

        f_help = ttk.LabelFrame(self, text="ℹ️ Guida Indicatori (Cosa sono e come si calcolano)")
        f_help.pack(fill="x", padx=15, pady=10)

        help_testo = (
            "• Tasso di Conversione:\n"
            "  - Cos'è: La percentuale di preventivi che si trasformano in contratti vinti.\n"
            "  - Come si calcola: (Preventivi con stato 'Approvato' ÷ Preventivi Totali) × 100.\n\n"
            "• Fatturato Pipeline (Imponibile):\n"
            "  - Cos'è: Il valore economico complessivo (al netto dell'IVA) di tutti i preventivi presenti nel gestionale.\n"
            "  - Come si calcola: La somma degli importi imponibili di tutti i preventivi salvati.\n\n"
            "• Margine Medio Reale:\n"
            "  - Cos'è: Il guadagno percentuale effettivo dell'officina rispetto al costo di produzione industriale sostenuto.\n"
            "  - Come si calcola: [(Totale Imponibile - Totale Costi Diretti / Aziendali) ÷ Totale Imponibile] × 100."
        )
        
        lbl_help = ttk.Label(f_help, text=help_testo, font=("Helvetica", 9), justify="left")
        lbl_help.pack(padx=10, pady=8, anchor="w")

        self.aggiorna()

    def aggiorna(self):
        for w in self.f_kpi.winfo_children(): w.destroy()
        for item in self.tree.get_children(): self.tree.delete(item)

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT COUNT(*), SUM(imponibile), SUM(costo_diretto) FROM preventivi")
            tot_p, sum_imp, sum_costo = c.fetchone()
            tot_p = tot_p or 0
            sum_imp = sum_imp or 0.0
            sum_costo = sum_costo or 0.0

            c.execute("SELECT COUNT(*) FROM preventivi WHERE stato='Approvato'")
            app_p = c.fetchone()[0] or 0

            tasso = (app_p / tot_p * 100) if tot_p > 0 else 0.0
            margine = ((sum_imp - sum_costo) / sum_imp * 100) if sum_imp > 0 else 0.0

            c.execute("SELECT id, titolo, cliente, data_consegna, stato, ivato FROM preventivi WHERE data_consegna IS NOT NULL AND data_consegna != '' ORDER BY data_consegna ASC")
            for r in c.fetchall():
                self.tree.insert("", "end", values=(r[0], r[1], r[2], r[3], r[4], f"€ {r[5]:.2f}"))

        kpis = [
            ("Preventivi Totali", f"{tot_p}", "#2B6CB0"),
            ("Tasso di Conversione Approvati", f"{tasso:.1f}%", "#2F855A"),
            ("Fatturato Pipeline (Imponibile)", f"€ {sum_imp:,.2f}", "#D69E2E"),
            ("Margine Medio Reale", f"{margine:.1f}%", "#805AD5")
        ]
        for i, (tit, val, col) in enumerate(kpis):
            box = ttk.LabelFrame(self.f_kpi, text=tit)
            box.grid(row=0, column=i, padx=5, sticky="nsew")
            self.f_kpi.columnconfigure(i, weight=1)
            ttk.Label(box, text=val, font=("Helvetica", 13, "bold"), foreground=col).pack(padx=10, pady=10)


# ==============================================================================
# 7. TAB PREVENTIVI & ESPORTAZIONE PDF PROFESSIONALE
# ==============================================================================
class TabPreventivi(ttk.Frame):
    STATI = ["Bozza", "Inviato", "Approvato", "Rifiutato", "In Lavorazione", "Completato"]

    def __init__(self, parent, db, callback_dash):
        super().__init__(parent)
        self.db = db
        self.callback_dash = callback_dash
        self.setup_ui()

    def get_imp(self):
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT chiave, valore FROM impostazioni")
            return dict(c.fetchall())

    def setup_ui(self):
        f_top = ttk.Frame(self)
        f_top.pack(fill="x", padx=10, pady=5)
        ttk.Button(f_top, text="+ Nuovo Preventivo", command=lambda: self.modal_nuovo_preventivo()).pack(side="left", padx=5)
        ttk.Button(f_top, text="📋 Modifica / Revisiona", command=self.apri_revisione).pack(side="left", padx=5)
        ttk.Button(f_top, text="🔍 Dettaglio", command=self.dettaglio).pack(side="left", padx=5)
        ttk.Button(f_top, text="📦 Ordine Materiali", command=self.ordine_materiali).pack(side="left", padx=5)
        ttk.Button(f_top, text="📄 PDF", command=self.esporta_pdf).pack(side="left", padx=5)
        ttk.Button(f_top, text="📊 Excel", command=self.esporta_excel).pack(side="left", padx=5)
        ttk.Button(f_top, text="🔄 Stato", command=self.cambia_stato).pack(side="left", padx=5)
        ttk.Button(f_top, text="🗑️ Elimina", command=self.elimina).pack(side="left", padx=5)

        cols = ("ID", "Titolo", "Rev.", "Cliente", "Imponibile", "Ivato", "Stato", "Consegna")
        self.tree = ttk.Treeview(self, columns=cols, show="headings")
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=50 if col == "Rev." else 120)
        self.tree.pack(fill="both", expand=True, padx=10, pady=5)
        self.tree.bind("<Double-1>", lambda e: self.dettaglio())
        self.aggiorna()

    def aggiorna(self):
        for item in self.tree.get_children(): self.tree.delete(item)
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, titolo, revisione, cliente, imponibile, ivato, stato, data_consegna FROM preventivi ORDER BY id DESC")
            for r in c.fetchall():
                self.tree.insert("", "end", values=(r[0], r[1], f"Rev.{r[2]}", r[3], f"€ {r[4]:.2f}", f"€ {r[5]:.2f}", r[6], r[7] or "-"))
        self.callback_dash()

    def cambia_stato(self):
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        win = tk.Toplevel(self)
        win.title("Cambia Stato")
        win.geometry("260x130")
        cb = ttk.Combobox(win, values=self.STATI, state="readonly", width=20)
        cb.pack(pady=15); cb.set(self.tree.item(sel[0], "values")[6])
        def salva():
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("UPDATE preventivi SET stato=? WHERE id=?", (cb.get(), pid))
                conn.commit()
            win.destroy(); self.aggiorna()
        ttk.Button(win, text="Aggiorna", command=salva).pack()

    def apri_revisione(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Seleziona", "Seleziona un preventivo da revisionare.")
            return
        pid = self.tree.item(sel[0], "values")[0]
        self.modal_nuovo_preventivo(preventivo_id_da_copiare=pid)

    def modal_nuovo_preventivo(self, preventivo_id_da_copiare=None):
        imp = self.get_imp()
        win = tk.Toplevel(self)
        win.title("Composizione / Revisione Preventivo Scenografico")
        win.geometry("1020x920")

        f_top = ttk.LabelFrame(win, text="Dati Generali")
        f_top.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(f_top, text="Titolo Progetto*:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        e_titolo = ttk.Entry(f_top, width=30); e_titolo.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(f_top, text="Cliente*:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome FROM clienti")
            clienti = [r[0] for r in c.fetchall()]
        combo_cl = ttk.Combobox(f_top, values=clienti, width=25); combo_cl.grid(row=0, column=3, padx=5, pady=5)

        ttk.Label(f_top, text="Inizio Lavori:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        e_in = ttk.Entry(f_top, width=15); e_in.insert(0, datetime.date.today().strftime("%d/%m/%Y")); e_in.grid(row=1, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(f_top, text="Consegna:").grid(row=1, column=2, padx=5, pady=5, sticky="w")
        e_co = ttk.Entry(f_top, width=15); e_co.insert(0, (datetime.date.today() + datetime.timedelta(days=15)).strftime("%d/%m/%Y")); e_co.grid(row=1, column=3, padx=5, pady=5, sticky="w")

        f_voci = ttk.LabelFrame(win, text="Voci Inserite nel Preventivo")
        f_voci.pack(fill="both", expand=True, padx=10, pady=5)
        
        cols = ("Tipo", "Nome", "Q.tà", "U.M.", "Ore", "Costo", "Vendita")
        tree_v = ttk.Treeview(f_voci, columns=cols, show="headings", height=6)
        for col in cols: tree_v.heading(col, text=col); tree_v.column(col, width=110)
        tree_v.pack(fill="both", expand=True, padx=5, pady=5)

        voci_list = []

        def rimuovi_voce_selezionata():
            sel_items = tree_v.selection()
            if not sel_items:
                messagebox.showwarning("Seleziona", "Seleziona una voce da eliminare.")
                return
            for item in sel_items:
                idx = tree_v.index(item)
                del voci_list[idx]
                tree_v.delete(item)

        ttk.Button(f_voci, text="🗑️ Rimuovi Voce Selezionata", command=rimuovi_voce_selezionata).pack(pady=5, anchor="e", padx=5)

        num_rev = 0
        padre_id = 0
        if preventivo_id_da_copiare:
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("SELECT titolo, cliente, km_trasporto, revisione FROM preventivi WHERE id=?", (preventivo_id_da_copiare,))
                p_old = c.fetchone()
                if p_old:
                    e_titolo.insert(0, p_old[0])
                    combo_cl.set(p_old[1])
                    num_rev = p_old[3] + 1
                    padre_id = preventivo_id_da_copiare

                c.execute("SELECT tipo, nome, qta, um, ore, costo_base, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=?", (preventivo_id_da_copiare,))
                for vr in c.fetchall():
                    if vr[0] not in ["Logistica", "Manodopera / Operatori", "Manodopera Cantiere"]:
                        voci_list.append({"tipo": vr[0], "nome": vr[1], "qta": vr[2], "um": vr[3], "ore": vr[4], "costo_base": vr[5], "prezzo": vr[6]})
                        tree_v.insert("", "end", values=(vr[0], vr[1], vr[2], vr[3], f"{vr[4]}h" if vr[4]>0 else "-", f"€ {vr[5]:.2f}", f"€ {vr[6]:.2f}"))

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome, prezzo, unita, sfrido FROM materiali")
            l_mat = c.fetchall()

        f_mat = ttk.LabelFrame(win, text="Aggiungi Materiale")
        f_mat.pack(fill="x", padx=10, pady=5)
        cb_m = ttk.Combobox(f_mat, values=[m[0] for m in l_mat], width=30); cb_m.grid(row=0, column=0, padx=5)
        e_qm = ttk.Entry(f_mat, width=8); e_qm.insert(0, "1"); e_qm.grid(row=0, column=2, padx=5)
        lbl_mat_um = ttk.Label(f_mat, text="U.M.: -", font=("Helvetica", 9, "bold"))
        lbl_mat_um.grid(row=0, column=3, padx=10)

        def update_mat_um(event=None):
            idx = cb_m.current()
            if idx >= 0:
                lbl_mat_um.config(text=f"U.M.: {l_mat[idx][2]}")
        cb_m.bind("<<ComboboxSelected>>", update_mat_um)

        def add_m():
            idx = cb_m.current()
            if idx < 0: return
            m = l_mat[idx]
            qta = parse_float_positivo(e_qm.get(), "Q.tà", False)
            costo = m[1] * qta * (1 + m[3]/100.0)
            vendita = costo * (1 + imp.get("ricarico_materiali", 35)/100.0)
            voci_list.append({"tipo": "Materiale", "nome": m[0], "qta": qta, "um": m[2], "ore": 0.0, "costo_base": costo, "prezzo": vendita})
            tree_v.insert("", "end", values=("Materiale", m[0], qta, m[2], "-", f"€ {costo:.2f}", f"€ {vendita:.2f}"))
        ttk.Button(f_mat, text="+ Aggiungi", command=add_m).grid(row=0, column=4, padx=10)

        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome, costo_orario, unita, ore_um FROM lavorazioni")
            l_lav = c.fetchall()

        f_lav = ttk.LabelFrame(win, text="Aggiungi Lavorazione Laboratorio")
        f_lav.pack(fill="x", padx=10, pady=5)
        cb_l = ttk.Combobox(f_lav, values=[l[0] for l in l_lav], width=30); cb_l.grid(row=0, column=0, padx=5)
        e_ql = ttk.Entry(f_lav, width=8); e_ql.insert(0, "1"); e_ql.grid(row=0, column=2, padx=5)
        e_ol = ttk.Entry(f_lav, width=8); e_ol.grid(row=0, column=4, padx=5)
        lbl_lav_um = ttk.Label(f_lav, text="U.M.: -", font=("Helvetica", 9, "bold"))
        lbl_lav_um.grid(row=0, column=5, padx=10)

        def update_lav_um(event=None):
            idx = cb_l.current()
            if idx >= 0:
                lbl_lav_um.config(text=f"U.M.: {l_lav[idx][2]}")
        cb_l.bind("<<ComboboxSelected>>", update_lav_um)

        def add_l():
            idx = cb_l.current()
            if idx < 0: return
            l = l_lav[idx]
            qta = parse_float_positivo(e_ql.get(), "Q.tà", False)
            ore = parse_float_positivo(e_ol.get(), "Ore", False)
            costo = ore * l[1]
            vendita = costo * (1 + imp.get("ricarico_manodopera", 25)/100.0)
            voci_list.append({"tipo": "Lavorazione", "nome": l[0], "qta": qta, "um": l[2], "ore": ore, "costo_base": costo, "prezzo": vendita})
            tree_v.insert("", "end", values=("Lavorazione", l[0], qta, l[2], f"{ore}h", f"€ {costo:.2f}", f"€ {vendita:.2f}"))
        ttk.Button(f_lav, text="+ Aggiungi", command=add_l).grid(row=0, column=6, padx=10)

        f_ext = ttk.LabelFrame(win, text="Cantiere, Trasporto & Personale Operativo")
        f_ext.pack(fill="x", padx=10, pady=5)

        ttk.Label(f_ext, text="Ore Falegname:").grid(row=0, column=0, padx=5, pady=3, sticky="e")
        e_of = ttk.Entry(f_ext, width=8); e_of.insert(0, "0"); e_of.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(f_ext, text="Ore Fabbro:").grid(row=0, column=2, padx=5, pady=3, sticky="e")
        e_ofab = ttk.Entry(f_ext, width=8); e_ofab.insert(0, "0"); e_ofab.grid(row=0, column=3, padx=5, pady=3)

        ttk.Label(f_ext, text="Ore Pittore:").grid(row=0, column=4, padx=5, pady=3, sticky="e")
        e_opit = ttk.Entry(f_ext, width=8); e_opit.insert(0, "0"); e_opit.grid(row=0, column=5, padx=5, pady=3)

        ttk.Label(f_ext, text="Ore Elettricista:").grid(row=1, column=0, padx=5, pady=3, sticky="e")
        e_oele = ttk.Entry(f_ext, width=8); e_oele.insert(0, "0"); e_oele.grid(row=1, column=1, padx=5, pady=3)

        ttk.Label(f_ext, text="Ore Amministrazione:").grid(row=1, column=2, padx=5, pady=3, sticky="e")
        e_oamm = ttk.Entry(f_ext, width=8); e_oamm.insert(0, "0"); e_oamm.grid(row=1, column=3, padx=5, pady=3)

        ttk.Label(f_ext, text="Ore Titolare:").grid(row=1, column=4, padx=5, pady=3, sticky="e")
        e_otit = ttk.Entry(f_ext, width=8); e_otit.insert(0, "0"); e_otit.grid(row=1, column=5, padx=5, pady=3)

        ttk.Label(f_ext, text="Km Trasporto A/R:").grid(row=2, column=0, padx=5, pady=5, sticky="e")
        e_km = ttk.Entry(f_ext, width=8); e_km.insert(0, "0"); e_km.grid(row=2, column=1, padx=5, pady=5)

        def salva():
            titolo = e_titolo.get().strip()
            cliente = combo_cl.get().strip()
            if not titolo or not cliente:
                messagebox.showerror("Errore", "Titolo e Cliente obbligatori!")
                return
            try:
                km = parse_float_positivo(e_km.get(), "Km Trasporto")
                ore_fal = parse_float_positivo(e_of.get(), "Ore Falegname")
                ore_fab = parse_float_positivo(e_ofab.get(), "Ore Fabbro")
                ore_pit = parse_float_positivo(e_opit.get(), "Ore Pittore")
                ore_ele = parse_float_positivo(e_oele.get(), "Ore Elettricista")
                ore_amm = parse_float_positivo(e_oamm.get(), "Ore Amministrazione")
                ore_tit = parse_float_positivo(e_otit.get(), "Ore Titolare")
            except ValueError as e:
                messagebox.showerror("Errore", str(e))
                return

            c_mat = sum(v["costo_base"] for v in voci_list if v["tipo"] == "Materiale")
            v_mat = sum(v["prezzo"] for v in voci_list if v["tipo"] == "Materiale")
            c_lav = sum(v["costo_base"] for v in voci_list if v["tipo"] == "Lavorazione")
            v_lav = sum(v["prezzo"] for v in voci_list if v["tipo"] == "Lavorazione")

            c_fal = ore_fal * imp.get("paga_falegname", 40)
            c_fab = ore_fab * imp.get("paga_fabbro", 45)
            c_pit = ore_pit * imp.get("paga_pittore", 40)
            c_ele = ore_ele * imp.get("paga_elettricista", 40)
            c_amm = ore_amm * imp.get("paga_amministrazione", 25)
            c_tit = ore_tit * imp.get("paga_titolare", 40)

            c_operai = c_fal + c_fab + c_pit + c_ele + c_amm + c_tit
            v_operai = c_operai * (1 + imp.get("ricarico_manodopera", 25)/100.0)

            c_tr = km * imp.get("costo_km", 0.90)
            markup = 1 + (imp.get("markup_finale", 30)/100.0)
            v_tr = c_tr * markup

            c_dir = c_mat + c_lav + c_operai + c_tr
            c_aziendale = c_dir * (1 + imp.get("spese_generali", 20)/100.0)

            v_base = v_mat + v_lav + v_operai + v_tr
            imponibile = v_base * markup
            iva = imponibile * (imp.get("iva", 22)/100.0)
            totale = imponibile + iva

            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("""
                    INSERT INTO preventivi (titolo, cliente, costo_diretto, imponibile, iva, ivato, km_trasporto, costo_trasporto, stato, revisione, padre_id, data_inizio, data_consegna)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Bozza', ?, ?, ?, ?)
                """, (titolo, cliente, c_aziendale, imponibile, iva, totale, km, c_tr, num_rev, padre_id, e_in.get(), e_co.get()))
                pid = c.lastrowid

                for v in voci_list:
                    c.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
                              (pid, v["tipo"], v["nome"], v["qta"], v["um"], v["ore"], v["costo_base"], v["prezzo"]))

                operai_map = [
                    ("Falegname (Cantiere/Opera)", ore_fal, c_fal),
                    ("Fabbro (Cantiere/Opera)", ore_fab, c_fab),
                    ("Pittore (Cantiere/Opera)", ore_pit, c_pit),
                    ("Elettricista (Cantiere/Opera)", ore_ele, c_ele),
                    ("Amministrazione", ore_amm, c_amm),
                    ("Titolare / Direzione", ore_tit, c_tit),
                ]

                for ruolo_nome, ore_val, costo_val in operai_map:
                    if ore_val > 0:
                        prezzo_val = costo_val * (1 + imp.get("ricarico_manodopera", 25)/100.0) * markup
                        c.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, ?, ?, 'h', ?, ?, ?)",
                                  (pid, "Manodopera / Operatori", ruolo_nome, ore_val, ore_val, costo_val, prezzo_val))

                if km > 0:
                    c.execute("INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita) VALUES (?, ?, 'Trasporto e Logistica A/R', ?, 'km', 0, ?, ?)",
                              (pid, "Logistica", km, c_tr, v_tr))
                conn.commit()

            messagebox.showinfo("Successo", f"Preventivo #{pid} (Rev.{num_rev}) salvato con successo!")
            win.destroy(); self.aggiorna()

        ttk.Button(win, text="💾 Calcola e Salva Preventivo", command=salva).pack(pady=10)

    def dettaglio(self):
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, titolo, cliente, imponibile, iva, ivato, stato, data_consegna, revisione FROM preventivi WHERE id=?", (pid,))
            p = c.fetchone()
            c.execute("SELECT tipo, nome, qta, um, ore, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=?", (pid,))
            voci = c.fetchall()

        win = tk.Toplevel(self)
        win.title(f"Dettaglio Preventivo #{p[0]} (Rev.{p[8]})")
        win.geometry("800x600")
        ttk.Label(win, text=f"Progetto: {p[1]} | Cliente: {p[2]} | Stato: {p[6]}", font=("Helvetica", 11, "bold")).pack(padx=10, pady=10, anchor="w")
        
        cols = ("Tipo", "Nome", "Q.tà", "U.M.", "Ore", "Prezzo Vendita")
        tree = ttk.Treeview(win, columns=cols, show="headings")
        for col in cols: tree.heading(col, text=col); tree.column(col, width=120)
        tree.pack(fill="both", expand=True, padx=10, pady=5)
        for v in voci:
            tree.insert("", "end", values=(v[0], v[1], v[2], v[3], f"{v[4]}h" if v[4]>0 else "-", f"€ {v[5]:.2f}"))
        ttk.Label(win, text=f"Imponibile: € {p[3]:.2f} | Totale Ivato: € {p[5]:.2f}", font=("Helvetica", 11, "bold"), foreground="green").pack(pady=15)

    def ordine_materiali(self):
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT nome, qta, um FROM voci_preventivo WHERE preventivo_id=? AND tipo='Materiale'", (pid,))
            rows = c.fetchall()
        if not rows:
            messagebox.showinfo("Info", "Nessun materiale in questo preventivo.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=f"Ordine_Materiali_Prev_{pid}.xlsx")
        if not path: return
        df = pd.DataFrame(rows, columns=["Materiale", "Quantità", "U.M."])
        df.to_excel(path, index=False)
        messagebox.showinfo("Completato", f"Ordine salvato in:\n{path}")

    def esporta_pdf(self):
        if not REPORTLAB_AVAILABLE:
            messagebox.showerror("Errore", "ReportLab non installato!")
            return
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT id, titolo, cliente, imponibile, iva, ivato, data_consegna FROM preventivi WHERE id=?", (pid,))
            p = c.fetchone()
            c.execute("SELECT tipo, nome, qta, um, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=?", (pid,))
            voci = c.fetchall()

        path = filedialog.asksaveasfilename(
            defaultextension=".pdf", 
            filetypes=[("PDF", "*.pdf")], 
            initialfile=f"Preventivo_{pid}_{p[1].replace(' ', '_')}.pdf"
        )
        if not path: return

        # Configurazione documento con margini idonei
        doc = SimpleDocTemplate(
            path,
            pagesize=A4,
            rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
        )
        styles = getSampleStyleSheet()

        # Stili personalizzati con wrapping testo automatico
        title_style = ParagraphStyle(
            'DocTitle',
            parent=styles['Title'],
            fontSize=18,
            leading=22,
            textColor=colors.HexColor('#1A365D'),
            alignment=0,
            spaceAfter=6
        )

        subtitle_style = ParagraphStyle(
            'SubTitle',
            parent=styles['Normal'],
            fontSize=10,
            leading=14,
            textColor=colors.HexColor('#4A5568'),
            spaceAfter=15
        )

        cell_style = ParagraphStyle(
            'CellText',
            parent=styles['Normal'],
            fontSize=9,
            leading=12,
            textColor=colors.HexColor('#2D3748')
        )

        cell_header_style = ParagraphStyle(
            'CellHeader',
            parent=styles['Normal'],
            fontSize=9,
            leading=12,
            fontName='Helvetica-Bold',
            textColor=colors.white
        )

        cell_right = ParagraphStyle('CellRight', parent=cell_style, alignment=2)
        cell_center = ParagraphStyle('CellCenter', parent=cell_style, alignment=1)
        header_right = ParagraphStyle('HeaderRight', parent=cell_header_style, alignment=2)
        header_center = ParagraphStyle('HeaderCenter', parent=cell_header_style, alignment=1)

        elements = []

        # Intestazione Preventivo
        elements.append(Paragraph(f"<b>Preventivo #{p[0]} - {p[1]}</b>", title_style))
        elements.append(Paragraph(f"<b>Cliente:</b> {p[2]} &nbsp;&nbsp;|&nbsp;&nbsp; <b>Data Consegna:</b> {p[6] or 'N/D'}", subtitle_style))
        elements.append(Spacer(1, 10))

        # Costruzione Dati Tabella con Paragrafi per evitare sovrapposizioni
        table_data = [
            [
                Paragraph("<b>Tipo</b>", cell_header_style),
                Paragraph("<b>Descrizione</b>", cell_header_style),
                Paragraph("<b>Q.tà</b>", header_center),
                Paragraph("<b>U.M.</b>", header_center),
                Paragraph("<b>Prezzo (€)</b>", header_right)
            ]
        ]

        for v in voci:
            table_data.append([
                Paragraph(str(v[0]), cell_style),
                Paragraph(str(v[1]), cell_style),
                Paragraph(f"{v[2]:.1f}", cell_center),
                Paragraph(str(v[3]), cell_center),
                Paragraph(f"€ {v[4]:.2f}", cell_right)
            ])

        # Larghezze colonne bilanciate (Totale = 523pt per pagina A4 con margini)
        col_widths = [125, 218, 45, 45, 90]
        t = Table(table_data, colWidths=col_widths, repeatRows=1)

        t_style = [
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
        ]

        # Alternanza colore righe
        for i in range(1, len(table_data)):
            if i % 2 == 0:
                t_style.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F7FAFC')))

        t.setStyle(TableStyle(t_style))
        elements.append(t)
        elements.append(Spacer(1, 20))

        # Sezione Totali
        totali_data = [
            [Paragraph("<b>Totale Imponibile:</b>", cell_right), Paragraph(f"€ {p[3]:.2f}", cell_right)],
            [Paragraph("<b>IVA (22%):</b>", cell_right), Paragraph(f"€ {p[4]:.2f}", cell_right)],
            [
                Paragraph("<b>Totale Ivato:</b>", ParagraphStyle('TotBold', parent=cell_right, fontName='Helvetica-Bold', fontSize=10)),
                Paragraph(f"<b>€ {p[5]:.2f}</b>", ParagraphStyle('TotBoldR', parent=cell_right, fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor('#2F855A')))
            ]
        ]
        totali_table = Table(totali_data, colWidths=[383, 140])
        totali_table.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LINEABOVE', (0, 2), (-1, 2), 1, colors.HexColor('#2B6CB0')),
        ]))

        elements.append(totali_table)

        doc.build(elements)
        messagebox.showinfo("PDF Generato", "Il PDF è stato generato con successo!")

    def esporta_excel(self):
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        with self.db.get_connection() as conn:
            c = conn.cursor()
            c.execute("SELECT tipo, nome, qta, um, ore, costo_base, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=?", (pid,))
            voci = c.fetchall()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if not path: return
        pd.DataFrame(voci, columns=["Tipo", "Nome", "Q.tà", "U.M.", "Ore", "Costo", "Vendita"]).to_excel(path, index=False)
        messagebox.showinfo("Excel", "Esportato con successo!")

    def elimina(self):
        sel = self.tree.selection()
        if not sel: return
        pid = self.tree.item(sel[0], "values")[0]
        if messagebox.askyesno("Conferma", "Eliminare il preventivo?"):
            with self.db.get_connection() as conn:
                c = conn.cursor()
                c.execute("DELETE FROM preventivi WHERE id=?", (pid,))
                conn.commit()
            self.aggiorna()


# ==============================================================================
# 8. APPLICAZIONE PRINCIPALE
# ==============================================================================
class GestionaleScenografia:
    def __init__(self, root):
        self.root = root
        self.root.title("Gestionale Preventivi Scenografici")
        self.root.geometry("1200x880")

        self.db = DatabaseManager()
        self.db.carica_da_excel_se_vuoto()

        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_dash = TabDashboard(self.notebook, self.db)
        self.tab_prev = TabPreventivi(self.notebook, self.db, self.tab_dash.aggiorna)
        self.tab_cli = TabClienti(self.notebook, self.db)
        self.tab_mat = TabMateriali(self.notebook, self.db)
        self.tab_lav = TabLavorazioni(self.notebook, self.db)
        self.tab_imp = TabImpostazioni(self.notebook, self.db, self.tab_dash.aggiorna)

        self.notebook.add(self.tab_dash, text="📊 Dashboard")
        self.notebook.add(self.tab_prev, text="📋 Preventivi")
        self.notebook.add(self.tab_cli, text="👥 Clienti")
        self.notebook.add(self.tab_mat, text="📦 Materiali")
        self.notebook.add(self.tab_lav, text="🛠️ Lavorazioni")
        self.notebook.add(self.tab_imp, text="⚙️ Impostazioni")


if __name__ == "__main__":
    root = tk.Tk()
    app = GestionaleScenografia(root)
    root.mainloop()
