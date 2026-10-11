
import datetime
import hmac
import io
import os
import re
import shutil
import sqlite3
from contextlib import contextmanager
from xml.sax.saxutils import escape

import pandas as pd
import streamlit as st

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
    from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

    REPORTLAB_AVAILABLE = True
except ImportError:  # il PDF viene disattivato, il resto dell'app funziona
    REPORTLAB_AVAILABLE = False


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.environ.get("GESTIONALE_DB", os.path.join(BASE_DIR, "gestionale_scenografia.db"))
BACKUP_DIR = os.path.join(os.path.dirname(DB_PATH), "backups")
MAX_BACKUP_LOCALI = 20

STATI = ["Bozza", "Inviato", "Approvato", "Rifiutato", "In Lavorazione", "Completato"]
STATI_VINTI = ("Approvato", "In Lavorazione", "Completato")

TIPO_MATERIALE = "Materiale"
TIPO_LAVORAZIONE = "Lavorazione"
TIPO_OPERATORI = "Manodopera / Operatori"
TIPO_LOGISTICA = "Logistica"

# (chiave, etichetta, nome della voce nel preventivo, chiave della paga oraria, paga di default)
RUOLI = [
    ("falegname", "Falegname", "Falegname (Cantiere/Opera)", "paga_falegname", 40.0),
    ("fabbro", "Fabbro", "Fabbro (Cantiere/Opera)", "paga_fabbro", 45.0),
    ("pittore", "Pittore", "Pittore (Cantiere/Opera)", "paga_pittore", 40.0),
    ("elettricista", "Elettricista", "Elettricista (Cantiere/Opera)", "paga_elettricista", 40.0),
    ("amministrazione", "Amministrazione", "Amministrazione", "paga_amministrazione", 25.0),
    ("titolare", "Titolare / Direzione", "Titolare / Direzione", "paga_titolare", 40.0),
]

IMPOSTAZIONI_DEFAULT = {
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
    "paga_titolare": (40.0, "€/h"),
}

PAGINA_DASHBOARD = "📊 Dashboard"
PAGINA_PREVENTIVI = "📋 Preventivi"
PAGINA_COMPOSER = "📝 Nuovo preventivo"
PAGINA_CLIENTI = "👥 Clienti"
PAGINA_MATERIALI = "📦 Materiali"
PAGINA_LAVORAZIONI = "🛠️ Lavorazioni"
PAGINA_IMPOSTAZIONI = "⚙️ Impostazioni"
PAGINA_DATI = "💾 Dati e backup"
PAGINE = [
    PAGINA_DASHBOARD, PAGINA_PREVENTIVI, PAGINA_COMPOSER, PAGINA_CLIENTI,
    PAGINA_MATERIALI, PAGINA_LAVORAZIONI, PAGINA_IMPOSTAZIONI, PAGINA_DATI,
]


# ==============================================================================
# UTILITÀ
# ==============================================================================
def eur(valore):
    """Formatta un importo all'italiana: 1.234,56 €"""
    try:
        testo = f"{float(valore):,.2f}"
    except (TypeError, ValueError):
        testo = "0.00"
    return "€ " + testo.replace(",", "X").replace(".", ",").replace("X", ".")


def num(valore):
    """Numero senza zeri inutili (2.0 -> 2, 2.5 -> 2.5)."""
    try:
        return f"{float(valore):g}"
    except (TypeError, ValueError):
        return str(valore)


def parse_data(valore):
    """Accetta date ISO (2026-10-11) o italiane (11/10/2026). Ritorna date o None."""
    if not valore:
        return None
    if isinstance(valore, datetime.datetime):
        return valore.date()
    if isinstance(valore, datetime.date):
        return valore
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.datetime.strptime(str(valore).strip()[:10], fmt).date()
        except ValueError:
            continue
    return None


def fmt_data(valore):
    d = parse_data(valore)
    return d.strftime("%d/%m/%Y") if d else "-"


def nome_file_sicuro(testo):
    return re.sub(r"[^\w\-]+", "_", str(testo)).strip("_") or "preventivo"


# ==============================================================================
# LISTINO DI PARTENZA (da Listino_Allestimenti_Scenografici.xlsx)
# Caricato automaticamente quando le tabelle materiali / lavorazioni sono vuote.
# ==============================================================================
# (codice, nome, U.M., prezzo di acquisto €, sfrido %)
LISTINO_MATERIALI = [
    ('LEG-001', 'Multistrato pioppo 4-6 mm', 'm²', 9, 15),
    ('LEG-002', 'Multistrato pioppo 10 mm', 'm²', 12, 15),
    ('LEG-003', 'Multistrato pioppo 15 mm', 'm²', 16, 15),
    ('LEG-004', 'Multistrato betulla 12-15 mm', 'm²', 24, 12),
    ('LEG-005', 'Multistrato betulla 18 mm', 'm²', 32, 12),
    ('LEG-006', 'OSB 9-12 mm', 'm²', 7, 15),
    ('LEG-007', 'OSB 15-18 mm', 'm²', 9.5, 15),
    ('LEG-008', 'MDF 6 mm', 'm²', 7, 12),
    ('LEG-009', 'MDF 10 mm', 'm²', 9.5, 12),
    ('LEG-010', 'MDF 19 mm', 'm²', 16, 12),
    ('LEG-011', 'Truciolare nobilitato', 'm²', 15, 10),
    ('LEG-012', 'Listello abete', 'ml', 3, 10),
    ('LEG-013', 'Listello strutturale in legno', 'ml', 4, 10),
    ('LEG-014', 'Pedana in legno - materiale', 'm²', 22, 10),
    ('FER-001', 'Tubolare quadro ferro 20x20', 'ml', 3, 10),
    ('FER-002', 'Tubolare quadro ferro 30x30', 'ml', 4.5, 10),
    ('FER-003', 'Tubolare quadro ferro 40x40', 'ml', 6, 10),
    ('FER-004', 'Tubolare quadro ferro 50x50', 'ml', 8, 10),
    ('FER-005', 'Tubolare rettangolare ferro', 'ml', 6, 10),
    ('FER-006', 'Tubo tondo ferro', 'ml', 5, 10),
    ('FER-007', 'Angolare ferro', 'ml', 4, 10),
    ('FER-008', 'Piattina ferro', 'ml', 3.5, 10),
    ('FER-009', 'Lamiera ferro', 'm²', 25, 10),
    ('FER-010', 'Profilo alluminio', 'ml', 12, 8),
    ('FER-011', 'Lamiera alluminio', 'm²', 38, 8),
    ('ED-001', 'Lastra cartongesso standard', 'm²', 6.5, 10),
    ('ED-002', 'Lastra cartongesso ignifuga', 'm²', 11, 10),
    ('ED-003', 'Profilo metallico cartongesso', 'ml', 2.5, 10),
    ('ED-004', 'Stucco cartongesso', 'kg', 1.8, 10),
    ('ED-005', 'Rasante', 'kg', 1.5, 10),
    ('ED-006', 'Primer/fondo', 'L', 6, 10),
    ('SC-001', 'Polistirolo EPS', 'm³', 180, 12),
    ('SC-002', 'Polistirene estruso XPS', 'm³', 260, 12),
    ('SC-003', 'Poliuretano espanso', 'm³', 450, 12),
    ('SC-004', 'Gommapiuma', 'm³', 350, 12),
    ('SC-005', 'Forex/PVC espanso 3-5 mm', 'm²', 14, 10),
    ('SC-006', 'Plexiglass', 'm²', 45, 8),
    ('SC-007', 'Policarbonato', 'm²', 38, 8),
    ('COL-001', 'Pittura murale', 'L', 5.5, 10),
    ('COL-002', 'Smalto', 'L', 12, 10),
    ('COL-003', 'Vernice trasparente', 'L', 14, 10),
    ('COL-004', 'Resina', 'kg', 14, 10),
    ('COL-005', 'Stucco decorativo', 'kg', 5, 10),
    ('TES-001', 'Moquette standard', 'm²', 8, 10),
    ('TES-002', 'Moquette premium', 'm²', 16, 10),
    ('TES-003', 'Tessuto scenografico', 'ml', 15, 10),
    ('TES-004', 'Velluto', 'ml', 28, 10),
    ('TES-005', 'Lycra', 'ml', 14, 10),
    ('TES-006', 'TNT', 'ml', 3, 10),
    ('TES-007', 'Tela stampabile', 'm²', 18, 8),
    ('GRA-001', 'PVC banner', 'm²', 8, 8),
    ('GRA-002', 'Vinile adesivo', 'm²', 12, 8),
    ('GRA-003', 'Forex stampato', 'm²', 22, 8),
    ('GRA-004', 'Dibond stampato', 'm²', 35, 8),
    ('GRA-005', 'Canvas stampato', 'm²', 24, 8),
    ('GRA-006', 'Pellicola trasparente', 'm²', 14, 8),
    ('ELE-001', 'Strip LED', 'ml', 12, 5),
    ('ELE-002', 'Profilo alluminio LED', 'ml', 10, 5),
    ('ELE-003', 'Alimentatore LED', 'pz', 25, 3),
    ('ELE-004', 'Faretto LED', 'pz', 30, 3),
    ('ELE-005', 'Cavo elettrico', 'ml', 1.8, 5),
    ('ELE-006', 'Multipresa/ciabatta', 'pz', 12, 3),
    ('ELE-007', 'Quadro elettrico', 'pz', 120, 3),
    ('FER-012', 'Viti e minuteria strutturale', 'kg', 6, 8),
    ('FER-013', 'Bulloneria', 'kg', 8, 8),
    ('FER-014', "Cavo d'acciaio", 'ml', 2.5, 5),
    ('FER-015', 'Catena', 'ml', 5, 5),
    ('MON-001', 'Ruote industriali', 'pz', 18, 3),
    ('MON-002', 'Piedino regolabile', 'pz', 8, 3),
    ('MON-003', 'Piastra/base metallica', 'pz', 18, 5),
    ('FIN-001', 'Carta abrasiva', 'pz', 1.2, 5),
    ('FIN-002', 'Nastro carta', 'pz', 3, 5),
    ('FIN-003', 'Silicone/adesivo', 'pz', 6, 5),
    ('FIN-004', 'Colla strutturale', 'kg', 15, 8),
    ('FIN-005', 'Schiuma poliuretanica', 'pz', 8, 8),
    ('FIN-006', 'Telo protettivo', 'm²', 2.5, 5),
    ('FIN-007', 'Biadesivo professionale', 'ml', 2.5, 5),
    ('FIN-008', 'Velcro professionale', 'ml', 4.5, 5),
    ('FIN-009', 'Profilo PVC finitura', 'ml', 3.5, 5),
    ('FIN-010', 'Angolare PVC', 'ml', 3, 5),
    ('FIN-011', 'Lamiera zincata', 'm²', 18, 8),
    ('FIN-012', 'Rete metallica', 'm²', 12, 8),
    ('FIN-013', 'Tessuto oscurante', 'm²', 10, 10),
    ('FIN-014', 'Pannello fonoassorbente', 'm²', 28, 8),
    ('FIN-015', 'Materiale ignifugo spray', 'L', 18, 8),
]

# (codice, nome, U.M., costo orario €/h, ore per U.M.)
LISTINO_LAVORAZIONI = [
    ('LAV-001', 'Taglio pannello', 'm²', 40, 0.15),
    ('LAV-002', 'Taglio pannello sagomato', 'pezzo', 40, 0.5),
    ('LAV-003', 'Assemblaggio struttura in legno', 'm²', 40, 0.5),
    ('LAV-004', 'Assemblaggio mobile/bancone', 'pezzo', 45, 4),
    ('LAV-005', 'Saldatura struttura semplice', 'ml', 45, 0.3),
    ('LAV-006', 'Saldatura struttura complessa', 'kg', 45, 0.15),
    ('LAV-007', 'Montaggio struttura metallica', 'm²', 45, 0.8),
    ('LAV-008', 'Verniciatura ferro', 'm²', 40, 0.45),
    ('LAV-009', 'Taglio CNC', 'm²', 55, 0.25),
    ('LAV-010', 'Lavorazione polistirolo CNC', 'm³', 50, 1),
    ('LAV-011', 'Rivestimento indurente polistirolo', 'm²', 40, 0.4),
    ('LAV-012', 'Rasatura elemento scenografico', 'm²', 40, 0.35),
    ('LAV-013', 'Cartongesso parete standard', 'm²', 40, 1),
    ('LAV-014', 'Cartongesso doppia lastra', 'm²', 40, 1.35),
    ('LAV-015', 'Cartongesso curvo', 'm²', 45, 1.5),
    ('LAV-016', 'Controsoffitto cartongesso', 'm²', 40, 1.1),
    ('LAV-017', 'Velette/cassonetti cartongesso', 'ml', 42, 0.9),
    ('LAV-018', 'Stuccatura', 'm²', 38, 0.2),
    ('LAV-019', 'Carteggiatura', 'm²', 38, 0.15),
    ('LAV-020', 'Primer/fondo applicazione', 'm²', 38, 0.12),
    ('LAV-021', 'Pittura semplice', 'm²', 38, 0.2),
    ('LAV-022', 'Pittura multicolore', 'm²', 40, 0.35),
    ('LAV-023', 'Effetto cemento', 'm²', 45, 0.5),
    ('LAV-024', 'Effetto marmo', 'm²', 50, 0.8),
    ('LAV-025', 'Effetto legno', 'm²', 45, 0.6),
    ('LAV-026', "Trompe-l'œil", 'm²', 55, 1.8),
    ('LAV-027', 'Decorazione artistica', 'm²', 50, 1.2),
    ('LAV-028', 'Patinatura/invecchiamento', 'm²', 45, 0.5),
    ('LAV-029', 'Applicazione tessuto', 'm²', 38, 0.25),
    ('LAV-030', 'Cucitura tessuto', 'ml', 35, 0.15),
    ('LAV-031', 'Applicazione adesivo', 'm²', 38, 0.2),
    ('LAV-032', 'Montaggio LED', 'ml', 42, 0.2),
    ('LAV-033', 'Installazione elettrica', 'ora', 45, 1),
    ('LAV-034', 'Cablaggio', 'ml', 45, 0.15),
    ('LAV-035', 'Montaggio generale', 'ora', 40, 1),
    ('LAV-036', 'Smontaggio generale', 'ora', 40, 1),
    ('LAV-037', 'Facchinaggio', 'ora', 30, 1),
    ('LAV-038', 'Preparazione laboratorio', 'ora', 35, 1),
    ('LAV-039', 'Progettazione scenografica', 'ora', 60, 1),
    ('LAV-040', 'Disegno tecnico CAD', 'ora', 55, 1),
    ('LAV-041', 'Modellazione 3D', 'ora', 60, 1),
    ('LAV-042', 'Render', 'pezzo', 60, 2),
    ('LAV-043', 'Grafica esecutiva', 'ora', 50, 1),
    ('LAV-044', 'Project management', 'ora', 55, 1),
    ('LAV-045', 'Imballaggio', 'pezzo', 35, 0.5),
    ('LAV-046', 'Carico/scarico', 'ora', 30, 1),
    ('LAV-047', 'Trasporto furgone', 'viaggio', 0, 1),
    ('LAV-048', 'Trasporto camion', 'viaggio', 0, 1),
    ('LAV-049', 'Stoccaggio', 'm³/mese', 0, 1),
    ('LAV-050', 'Smaltimento', 'm³', 0, 1),
]


# ==============================================================================
# 1. DATABASE
# ==============================================================================
class DatabaseManager:
    def __init__(self, db_path=DB_PATH):
        self.db_path = db_path
        os.makedirs(os.path.dirname(os.path.abspath(db_path)), exist_ok=True)
        self.backup_automatico()
        self.init_db()
        self.carica_listino_iniziale()

    # ---- accesso ----------------------------------------------------------
    @contextmanager
    def conn(self):
        con = sqlite3.connect(self.db_path)
        con.row_factory = sqlite3.Row
        con.execute("PRAGMA foreign_keys = ON")  # necessario per ON DELETE CASCADE
        try:
            yield con
            con.commit()
        except Exception:
            con.rollback()
            raise
        finally:
            con.close()

    def query(self, sql, params=()):
        with self.conn() as con:
            return [dict(r) for r in con.execute(sql, params).fetchall()]

    def execute(self, sql, params=()):
        with self.conn() as con:
            return con.execute(sql, params).lastrowid

    # ---- backup -----------------------------------------------------------
    def backup_automatico(self):
        if not os.path.exists(self.db_path):
            return
        try:
            os.makedirs(BACKUP_DIR, exist_ok=True)
            stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            shutil.copy2(self.db_path, os.path.join(BACKUP_DIR, f"gestionale_backup_{stamp}.db"))
            vecchi = sorted(f for f in os.listdir(BACKUP_DIR) if f.endswith(".db"))
            for f in vecchi[:-MAX_BACKUP_LOCALI]:
                os.remove(os.path.join(BACKUP_DIR, f))
        except Exception as e:  # il backup non deve mai bloccare l'app
            print(f"[BACKUP] Errore: {e}")

    def esporta_bytes(self):
        """Copia coerente del database, da scaricare."""
        try:
            with self.conn() as con:
                return con.serialize()
        except Exception:
            with open(self.db_path, "rb") as f:
                return f.read()

    def ripristina_da_bytes(self, dati):
        """Sostituisce il database con quello caricato. Ritorna (ok, messaggio)."""
        tmp = self.db_path + ".upload"
        try:
            with open(tmp, "wb") as f:
                f.write(dati)
            try:
                con = sqlite3.connect(tmp)
                tabelle = {r[0] for r in con.execute("SELECT name FROM sqlite_master WHERE type='table'")}
                con.close()
            except sqlite3.DatabaseError:
                return False, "Il file caricato non è un database valido."
            mancanti = {"preventivi", "voci_preventivo", "materiali", "clienti"} - tabelle
            if mancanti:
                return False, "Il file non sembra un database del gestionale (mancano: " + ", ".join(sorted(mancanti)) + ")."
            self.backup_automatico()
            os.replace(tmp, self.db_path)
            self.init_db()
            return True, "Database ripristinato."
        except Exception as e:
            return False, f"Errore durante il ripristino: {e}"
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

    # ---- struttura --------------------------------------------------------
    def init_db(self):
        with self.conn() as con:
            c = con.cursor()
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
            c.execute("PRAGMA table_info(impostazioni)")
            if "unita" not in [col[1] for col in c.fetchall()]:
                c.execute("ALTER TABLE impostazioni ADD COLUMN unita TEXT")

            for chiave, (valore, unita) in IMPOSTAZIONI_DEFAULT.items():
                c.execute("INSERT OR IGNORE INTO impostazioni (chiave, valore, unita) VALUES (?, ?, ?)",
                          (chiave, valore, unita))
                c.execute("UPDATE impostazioni SET unita=? WHERE chiave=? AND (unita IS NULL OR unita='')",
                          (unita, chiave))

    # ---- impostazioni -----------------------------------------------------
    def get_imp(self):
        return {r["chiave"]: r["valore"] for r in self.query("SELECT chiave, valore FROM impostazioni")}

    # ---- listino Excel ----------------------------------------------------
    def importa_listino(self, sorgente):
        """Importa materiali e lavorazioni da un file Excel (percorso o file caricato).
        I codici già presenti vengono ignorati. Ritorna (n_materiali, n_lavorazioni)."""
        xls = pd.ExcelFile(sorgente)
        n_mat = n_lav = 0

        def testo(riga, chiave, default=""):
            v = riga.get(chiave)
            return default if pd.isna(v) else str(v).strip()

        def numero(riga, chiave, default):
            v = pd.to_numeric(riga.get(chiave), errors="coerce")
            return default if pd.isna(v) else float(v)

        with self.conn() as con:
            if "Listino_Materiali" in xls.sheet_names:
                for _, r in pd.read_excel(xls, sheet_name="Listino_Materiali").iterrows():
                    nome = testo(r, "Materiale")
                    if not nome:
                        continue
                    sfrido = numero(r, "Sfrido %", 0.10)
                    sfrido = sfrido * 100 if sfrido < 1 else sfrido  # 0.10 oppure 10 -> 10 %
                    codice = testo(r, "Codice") or None
                    if codice is None and con.execute("SELECT 1 FROM materiali WHERE nome=?", (nome,)).fetchone():
                        continue  # senza codice il vincolo UNIQUE non protegge: confronto per nome
                    cur = con.execute(
                        "INSERT OR IGNORE INTO materiali (codice, nome, unita, prezzo, sfrido) VALUES (?, ?, ?, ?, ?)",
                        (codice, nome, testo(r, "U.M.", "pz"), numero(r, "Costo acquisto (€)", 0.0), sfrido))
                    n_mat += cur.rowcount
            if "Listino_Lavorazioni" in xls.sheet_names:
                for _, r in pd.read_excel(xls, sheet_name="Listino_Lavorazioni").iterrows():
                    nome = testo(r, "Lavorazione")
                    if not nome:
                        continue
                    codice = testo(r, "Codice") or None
                    if codice is None and con.execute("SELECT 1 FROM lavorazioni WHERE nome=?", (nome,)).fetchone():
                        continue
                    cur = con.execute(
                        "INSERT OR IGNORE INTO lavorazioni (codice, nome, unita, costo_orario, ore_um) VALUES (?, ?, ?, ?, ?)",
                        (codice, nome, testo(r, "U.M.", "m²"),
                         numero(r, "Costo orario (€)", 40.0), numero(r, "Ore/U.M.", 1.0)))
                    n_lav += cur.rowcount
        return n_mat, n_lav

    def carica_listino_iniziale(self):
        """Se le tabelle sono vuote, le riempie con il listino incorporato nel programma."""
        with self.conn() as con:
            if con.execute("SELECT COUNT(*) FROM materiali").fetchone()[0] == 0:
                con.executemany(
                    "INSERT OR IGNORE INTO materiali (codice, nome, unita, prezzo, sfrido) VALUES (?, ?, ?, ?, ?)",
                    [(c, n, u, p, s) for c, n, u, p, s in LISTINO_MATERIALI])
            if con.execute("SELECT COUNT(*) FROM lavorazioni").fetchone()[0] == 0:
                con.executemany(
                    "INSERT OR IGNORE INTO lavorazioni (codice, nome, unita, costo_orario, ore_um) VALUES (?, ?, ?, ?, ?)",
                    [(c, n, u, co, o) for c, n, u, co, o in LISTINO_LAVORAZIONI])

    # ---- preventivi -------------------------------------------------------
    def salva_preventivo(self, titolo, cliente, inizio, consegna, km, calc, revisione, padre_id):
        costo_trasporto = sum(r["costo_base"] for r in calc["righe"] if r["tipo"] == TIPO_LOGISTICA)
        with self.conn() as con:
            cur = con.execute(
                """INSERT INTO preventivi (titolo, cliente, costo_diretto, imponibile, iva, ivato, km_trasporto,
                                           costo_trasporto, stato, revisione, padre_id, data_inizio, data_consegna)
                   VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'Bozza', ?, ?, ?, ?)""",
                (titolo, cliente, calc["costo_aziendale"], calc["imponibile"], calc["iva"], calc["ivato"], km,
                 costo_trasporto, revisione, padre_id,
                 inizio.isoformat() if inizio else None, consegna.isoformat() if consegna else None))
            pid = cur.lastrowid
            for r in calc["righe"]:
                con.execute(
                    """INSERT INTO voci_preventivo (preventivo_id, tipo, nome, qta, um, ore, costo_base, prezzo_vendita)
                       VALUES (?, ?, ?, ?, ?, ?, ?, ?)""",
                    (pid, r["tipo"], r["nome"], r["qta"], r["um"], r["ore"], r["costo_base"], r["prezzo"]))
        return pid


@st.cache_resource
def get_db():
    return DatabaseManager()


# ==============================================================================
# 2. CALCOLO DEL PREVENTIVO
# ==============================================================================
def calcola(voci, ore_ruoli, km, imp):
    """Calcola righe e totali.

    Ogni riga ha un prezzo di vendita già comprensivo di ricarico di categoria
    (materiali / manodopera) e di markup finale: la somma delle righe coincide
    quindi con l'imponibile mostrato nel PDF.
    """
    def g(chiave, default):
        return float(imp.get(chiave, default))

    markup = 1 + g("markup_finale", 30.0) / 100
    ric_mat = 1 + g("ricarico_materiali", 35.0) / 100
    ric_man = 1 + g("ricarico_manodopera", 25.0) / 100

    righe = []
    for v in voci:
        ric = ric_mat if v["tipo"] == TIPO_MATERIALE else ric_man
        righe.append({**v, "prezzo": round(v["costo_base"] * ric * markup, 2)})

    for chiave, _, nome_voce, chiave_paga, paga_default in RUOLI:
        ore = float(ore_ruoli.get(chiave) or 0)
        if ore > 0:
            costo = ore * g(chiave_paga, paga_default)
            righe.append({"tipo": TIPO_OPERATORI, "nome": nome_voce, "qta": ore, "um": "h", "ore": ore,
                          "costo_base": costo, "prezzo": round(costo * ric_man * markup, 2)})

    if km and km > 0:
        costo = km * g("costo_km", 0.90)
        righe.append({"tipo": TIPO_LOGISTICA, "nome": "Trasporto e Logistica A/R", "qta": km, "um": "km",
                      "ore": 0.0, "costo_base": costo, "prezzo": round(costo * markup, 2)})

    costo_diretto = sum(r["costo_base"] for r in righe)
    costo_aziendale = costo_diretto * (1 + g("spese_generali", 20.0) / 100)
    imponibile = round(sum(r["prezzo"] for r in righe), 2)
    iva_perc = g("iva", 22.0)
    iva = round(imponibile * iva_perc / 100, 2)
    ivato = round(imponibile + iva, 2)
    margine = (imponibile - costo_aziendale) / imponibile * 100 if imponibile > 0 else 0.0
    return {"righe": righe, "costo_diretto": costo_diretto, "costo_aziendale": costo_aziendale,
            "imponibile": imponibile, "iva_perc": iva_perc, "iva": iva, "ivato": ivato, "margine": margine}


# ==============================================================================
# 3. ESPORTAZIONI (PDF / EXCEL)
# ==============================================================================
def genera_pdf(db, pid):
    p = db.query("SELECT * FROM preventivi WHERE id=?", (pid,))[0]
    voci = db.query("SELECT tipo, nome, qta, um, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=? ORDER BY id", (pid,))

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36,
                            title=f"Preventivo {pid}")
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('DocTitle', parent=styles['Title'], fontSize=18, leading=22,
                                 textColor=colors.HexColor('#1A365D'), alignment=0, spaceAfter=6)
    subtitle_style = ParagraphStyle('SubTitle', parent=styles['Normal'], fontSize=10, leading=14,
                                    textColor=colors.HexColor('#4A5568'), spaceAfter=15)
    cell_style = ParagraphStyle('CellText', parent=styles['Normal'], fontSize=9, leading=12,
                                textColor=colors.HexColor('#2D3748'))
    cell_header = ParagraphStyle('CellHeader', parent=styles['Normal'], fontSize=9, leading=12,
                                 fontName='Helvetica-Bold', textColor=colors.white)
    cell_right = ParagraphStyle('CellRight', parent=cell_style, alignment=2)
    cell_center = ParagraphStyle('CellCenter', parent=cell_style, alignment=1)
    header_right = ParagraphStyle('HeaderRight', parent=cell_header, alignment=2)
    header_center = ParagraphStyle('HeaderCenter', parent=cell_header, alignment=1)

    rev = f" (Rev.{p['revisione']})" if p["revisione"] else ""
    elements = [
        Paragraph(f"<b>Preventivo #{p['id']}{rev} - {escape(p['titolo'])}</b>", title_style),
        Paragraph(f"<b>Cliente:</b> {escape(p['cliente'])} &nbsp;&nbsp;|&nbsp;&nbsp; "
                  f"<b>Data consegna:</b> {fmt_data(p['data_consegna'])}", subtitle_style),
        Spacer(1, 10),
    ]

    table_data = [[
        Paragraph("<b>Tipo</b>", cell_header), Paragraph("<b>Descrizione</b>", cell_header),
        Paragraph("<b>Q.tà</b>", header_center), Paragraph("<b>U.M.</b>", header_center),
        Paragraph("<b>Prezzo</b>", header_right),
    ]]
    for v in voci:
        table_data.append([
            Paragraph(escape(str(v["tipo"] or "")), cell_style),
            Paragraph(escape(str(v["nome"] or "")), cell_style),
            Paragraph(num(v["qta"] or 0), cell_center),
            Paragraph(escape(str(v["um"] or "")), cell_center),
            Paragraph(eur(v["prezzo_vendita"] or 0), cell_right),
        ])

    t = Table(table_data, colWidths=[125, 218, 45, 45, 90], repeatRows=1)
    t_style = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2B6CB0')),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#E2E8F0')),
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor('#CBD5E0')),
    ]
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            t_style.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor('#F7FAFC')))
    t.setStyle(TableStyle(t_style))
    elements += [t, Spacer(1, 20)]

    imponibile, iva, ivato = p["imponibile"] or 0, p["iva"] or 0, p["ivato"] or 0
    iva_perc = round(iva / imponibile * 100, 2) if imponibile else 22
    tot_bold = ParagraphStyle('TotBold', parent=cell_right, fontName='Helvetica-Bold', fontSize=10)
    tot_bold_g = ParagraphStyle('TotBoldR', parent=tot_bold, textColor=colors.HexColor('#2F855A'))
    totali = Table([
        [Paragraph("<b>Totale imponibile:</b>", cell_right), Paragraph(eur(imponibile), cell_right)],
        [Paragraph(f"<b>IVA ({num(iva_perc)}%):</b>", cell_right), Paragraph(eur(iva), cell_right)],
        [Paragraph("<b>Totale ivato:</b>", tot_bold), Paragraph(f"<b>{eur(ivato)}</b>", tot_bold_g)],
    ], colWidths=[383, 140])
    totali.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'RIGHT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
        ('LINEABOVE', (0, 2), (-1, 2), 1, colors.HexColor('#2B6CB0')),
    ]))
    elements.append(totali)
    doc.build(elements)
    return buf.getvalue()


def _excel_bytes(df, nome_foglio):
    from openpyxl.utils import get_column_letter
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=nome_foglio)
        ws = writer.sheets[nome_foglio]
        for i, col in enumerate(df.columns, 1):
            larghezza = max([len(str(col))] + [len(str(x)) for x in df[col]]) + 2
            ws.column_dimensions[get_column_letter(i)].width = min(larghezza, 60)
    return buf.getvalue()


def genera_excel(db, pid):
    voci = db.query("SELECT tipo, nome, qta, um, ore, costo_base, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=? ORDER BY id", (pid,))
    df = pd.DataFrame(voci).rename(columns={"tipo": "Tipo", "nome": "Nome", "qta": "Q.tà", "um": "U.M.",
                                            "ore": "Ore", "costo_base": "Costo", "prezzo_vendita": "Vendita"})
    return _excel_bytes(df, "Preventivo")


def genera_ordine_materiali(db, pid):
    righe = db.query("SELECT nome, qta, um FROM voci_preventivo WHERE preventivo_id=? AND tipo=? ORDER BY id",
                     (pid, TIPO_MATERIALE))
    if not righe:
        return None
    df = pd.DataFrame(righe).rename(columns={"nome": "Materiale", "qta": "Quantità", "um": "U.M."})
    return _excel_bytes(df, "Ordine materiali")


# ==============================================================================
# 4. STATO DELLA BOZZA (pagina "Nuovo preventivo")
# ==============================================================================
DRAFT_WIDGET_KEYS = (
    ["d_titolo", "d_cliente", "d_inizio", "d_consegna", "d_km",
     "d_mat_sel", "d_mat_qta", "d_lav_sel", "d_lav_qta", "d_lav_ore", "d_rimuovi"]
    + [f"d_ore_{r[0]}" for r in RUOLI]
)


def reset_draft():
    ss = st.session_state
    oggi = datetime.date.today()
    ss["d_titolo"] = ""
    ss["d_cliente"] = None
    ss["d_inizio"] = oggi
    ss["d_consegna"] = oggi + datetime.timedelta(days=15)
    ss["d_km"] = 0.0
    ss["d_mat_sel"] = None
    ss["d_mat_qta"] = 1.0
    ss["d_lav_sel"] = None
    ss["d_lav_qta"] = 1.0
    ss["d_lav_ore"] = 1.0
    ss["d_rimuovi"] = []
    for r in RUOLI:
        ss[f"d_ore_{r[0]}"] = 0.0
    ss["d_voci"] = []
    ss["d_padre"] = 0
    ss["d_rev"] = 0


def ensure_draft():
    if "d_voci" not in st.session_state:
        reset_draft()


def keep_draft_alive():
    """Streamlit cancella il valore dei widget non disegnati: riassegnandolo a ogni
    esecuzione la bozza sopravvive quando si cambia pagina."""
    for k in DRAFT_WIDGET_KEYS:
        if k in st.session_state:
            st.session_state[k] = st.session_state[k]


def carica_revisione(pid):
    """Callback: copia un preventivo nella bozza come nuova revisione."""
    db = get_db()
    p = db.query("SELECT * FROM preventivi WHERE id=?", (pid,))
    if not p:
        return
    p = p[0]
    reset_draft()
    ss = st.session_state
    oggi = datetime.date.today()
    ss["d_titolo"] = p["titolo"]
    ss["d_cliente"] = p["cliente"]
    ss["d_padre"] = pid
    ss["d_rev"] = (p["revisione"] or 0) + 1
    ss["d_inizio"] = parse_data(p["data_inizio"]) or oggi
    ss["d_consegna"] = parse_data(p["data_consegna"]) or (oggi + datetime.timedelta(days=15))
    ss["d_km"] = float(p["km_trasporto"] or 0)
    per_nome = {r[2]: r[0] for r in RUOLI}
    for v in db.query("SELECT * FROM voci_preventivo WHERE preventivo_id=? ORDER BY id", (pid,)):
        if v["tipo"] in (TIPO_MATERIALE, TIPO_LAVORAZIONE):
            ss["d_voci"].append({"tipo": v["tipo"], "nome": v["nome"], "qta": v["qta"], "um": v["um"],
                                 "ore": v["ore"] or 0.0, "costo_base": v["costo_base"] or 0.0})
        elif v["tipo"] == TIPO_OPERATORI and v["nome"] in per_nome:
            ss[f"d_ore_{per_nome[v['nome']]}"] = float(v["ore"] or 0)
    ss["pagina"] = PAGINA_COMPOSER


def nuova_bozza():
    reset_draft()
    st.session_state["pagina"] = PAGINA_COMPOSER


def rimuovi_voci_selezionate():
    ss = st.session_state
    da_rimuovere = set(ss.get("d_rimuovi", []))
    ss["d_voci"] = [v for i, v in enumerate(ss["d_voci"]) if i not in da_rimuovere]
    ss["d_rimuovi"] = []


def salva_bozza():
    """Callback del pulsante 'Calcola e salva'."""
    ss = st.session_state
    db = get_db()
    titolo = (ss.get("d_titolo") or "").strip()
    cliente = (ss.get("d_cliente") or "")
    cliente = cliente.strip() if isinstance(cliente, str) else ""
    if not titolo or not cliente:
        ss["_err_comp"] = "Titolo e Cliente sono obbligatori."
        return
    inizio, consegna = ss.get("d_inizio"), ss.get("d_consegna")
    if inizio and consegna and consegna < inizio:
        ss["_err_comp"] = "La data di consegna è precedente all'inizio lavori."
        return
    ore = {r[0]: ss.get(f"d_ore_{r[0]}", 0.0) for r in RUOLI}
    km = float(ss.get("d_km") or 0)
    calc = calcola(ss["d_voci"], ore, km, db.get_imp())
    if not calc["righe"] or calc["imponibile"] <= 0:
        ss["_err_comp"] = "Il preventivo è vuoto: aggiungi almeno una voce, delle ore o dei km."
        return
    rev, padre = ss["d_rev"], ss["d_padre"]
    pid = db.salva_preventivo(titolo, cliente, inizio, consegna, km, calc, rev, padre)
    ss["_flash"] = f"Preventivo #{pid} (Rev.{rev}) salvato con successo."
    reset_draft()
    ss["pagina"] = PAGINA_PREVENTIVI


# ==============================================================================
# 5. PAGINE
# ==============================================================================
def mostra_flash():
    msg = st.session_state.pop("_flash", None)
    if msg:
        st.success(msg)


def pagina_dashboard(db):
    st.header("📊 Dashboard e KPI aziendali")
    rows = db.query(
        """SELECT p.id, p.titolo, p.cliente, p.data_consegna, p.stato, p.imponibile, p.ivato, p.costo_diretto
           FROM preventivi p
           WHERE NOT EXISTS (SELECT 1 FROM preventivi r WHERE r.padre_id = p.id)"""
    )  # solo l'ultima revisione di ciascun preventivo
    totali = len(rows)
    vinti = sum(1 for r in rows if r["stato"] in STATI_VINTI)
    attivi = [r for r in rows if r["stato"] != "Rifiutato"]
    sum_imp = sum(r["imponibile"] or 0 for r in attivi)
    sum_costo = sum(r["costo_diretto"] or 0 for r in attivi)
    tasso = vinti / totali * 100 if totali else 0.0
    margine = (sum_imp - sum_costo) / sum_imp * 100 if sum_imp > 0 else 0.0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Preventivi totali", totali)
    c2.metric("Tasso di conversione", f"{tasso:.1f}%")
    c3.metric("Fatturato pipeline (imponibile)", eur(sum_imp))
    c4.metric("Margine medio reale", f"{margine:.1f}%")

    st.subheader("🗓️ Prossimi cantieri in consegna")
    oggi = datetime.date.today()
    prossimi = []
    for r in rows:
        d = parse_data(r["data_consegna"])
        if d and r["stato"] not in ("Rifiutato", "Completato"):
            giorni = (d - oggi).days
            scadenza = f"tra {giorni} gg" if giorni > 0 else ("oggi" if giorni == 0 else f"in ritardo di {-giorni} gg")
            prossimi.append((d, {"ID": r["id"], "Progetto": r["titolo"], "Cliente": r["cliente"],
                                 "Consegna": d.strftime("%d/%m/%Y"), "Scadenza": scadenza,
                                 "Stato": r["stato"], "Importo ivato": eur(r["ivato"] or 0)}))
    if prossimi:
        st.dataframe(pd.DataFrame([x[1] for x in sorted(prossimi, key=lambda x: x[0])]), hide_index=True)
    else:
        st.info("Nessun cantiere in consegna.")

    with st.expander("ℹ️ Guida indicatori (cosa sono e come si calcolano)"):
        st.markdown(
            "Gli indicatori considerano **solo l'ultima revisione** di ciascun preventivo, "
            "così le revisioni non vengono contate più volte.\n\n"
            "**Tasso di conversione** – percentuale di preventivi che diventano contratti. "
            "(Preventivi *Approvato*, *In Lavorazione* o *Completato* ÷ Preventivi totali) × 100.\n\n"
            "**Fatturato pipeline (imponibile)** – somma degli imponibili (IVA esclusa) dei preventivi "
            "non rifiutati.\n\n"
            "**Margine medio reale** – [(Imponibile − Costi aziendali) ÷ Imponibile] × 100, sui preventivi non "
            "rifiutati. I costi aziendali sono i costi diretti maggiorati delle spese generali."
        )


# ---- CRUD generico (clienti, materiali, lavorazioni) ------------------------
def campo(k, label, tipo="text", **kw):
    return {"k": k, "label": label, "tipo": tipo, **kw}


def pagina_crud(db, cfg):
    tabella = cfg["tabella"]
    campi = cfg["campi"]
    st.header(cfg["titolo"])

    q = st.text_input("🔍 Cerca", key=f"cerca_{tabella}")
    sql = f"SELECT * FROM {tabella}"
    params = ()
    if q.strip():
        sql += " WHERE " + " OR ".join(f"{c} LIKE ?" for c in cfg["ricerca"])
        params = tuple(f"%{q.strip()}%" for _ in cfg["ricerca"])
    rows = db.query(sql + " ORDER BY id DESC", params)

    if rows:
        df = pd.DataFrame([{h: (fmt(r[c]) if fmt else ("" if r[c] is None else r[c]))
                            for c, h, fmt in cfg["colonne"]} for r in rows])
        st.dataframe(df, hide_index=True)
    else:
        st.info("Nessun elemento trovato.")

    def render_campi(prefisso, valori):
        out = {}
        for c in campi:
            key = f"{prefisso}_{c['k']}"
            iniziale = (valori or {}).get(c["k"])
            if c["tipo"] == "float":
                base = iniziale if iniziale is not None else c.get("default", 0.0)
                out[c["k"]] = st.number_input(c["label"], min_value=0.0, value=float(base),
                                              step=c.get("step", 0.5), format="%.2f", key=key)
            else:
                base = iniziale if iniziale is not None else c.get("default", "")
                if c["tipo"] == "area":
                    out[c["k"]] = st.text_area(c["label"], value=str(base), key=key)
                else:
                    out[c["k"]] = st.text_input(c["label"], value=str(base), key=key)
        return out

    def valida(valori):
        errori = []
        for c in campi:
            v = valori[c["k"]]
            if c.get("obbl") and not str(v).strip():
                errori.append(f"Il campo «{c['label'].rstrip(' *')}» è obbligatorio.")
            if c.get("pos") and float(v) <= 0:
                errori.append(f"Il campo «{c['label'].rstrip(' *')}» deve essere maggiore di zero.")
        return errori

    def codice_libero():
        n = (db.query(f"SELECT COALESCE(MAX(id), 0) + 1 AS n FROM {tabella}")[0]["n"])
        while db.query(f"SELECT 1 FROM {tabella} WHERE codice=?", (f"{cfg['prefisso']}-{n:03d}",)):
            n += 1
        return f"{cfg['prefisso']}-{n:03d}"

    def salva(rid, valori):
        valori = {k: (v.strip() if isinstance(v, str) else v) for k, v in valori.items()}
        if cfg.get("prefisso") and not valori.get("codice"):
            valori["codice"] = codice_libero()
        cols = [c["k"] for c in campi]
        try:
            if rid:
                db.execute(f"UPDATE {tabella} SET " + ", ".join(f"{c}=?" for c in cols) + " WHERE id=?",
                           tuple(valori[c] for c in cols) + (rid,))
            else:
                db.execute(f"INSERT INTO {tabella} ({', '.join(cols)}) VALUES ({', '.join('?' for _ in cols)})",
                           tuple(valori[c] for c in cols))
        except sqlite3.IntegrityError:
            st.error("Esiste già un elemento con lo stesso codice.")
            return False
        return True

    with st.expander(f"➕ Nuovo {cfg['singolare']}", expanded=not rows and not q.strip()):
        with st.form(f"form_nuovo_{tabella}", clear_on_submit=True):
            valori = render_campi(f"nuovo_{tabella}", None)
            invia = st.form_submit_button("Salva")
        if invia:
            errori = valida(valori)
            if errori:
                for e in errori:
                    st.error(e)
            elif salva(None, valori):
                st.session_state["_flash"] = f"{cfg['singolare'].capitalize()} salvato."
                st.rerun()

    if rows:
        with st.expander(f"✏️ Modifica o elimina un {cfg['singolare']}"):
            per_id = {r["id"]: r for r in rows}
            rid = st.selectbox("Seleziona", list(per_id), key=f"sel_{tabella}",
                               format_func=lambda i: cfg["label"](per_id[i]))
            with st.form(f"form_mod_{tabella}_{rid}"):
                valori = render_campi(f"mod_{tabella}_{rid}", per_id[rid])
                invia = st.form_submit_button("Salva modifiche")
            if invia:
                errori = valida(valori)
                if errori:
                    for e in errori:
                        st.error(e)
                elif salva(rid, valori):
                    st.session_state["_flash"] = "Modifiche salvate."
                    st.rerun()
            st.divider()
            conferma = st.checkbox("Confermo di voler eliminare l'elemento selezionato", key=f"conf_{tabella}_{rid}")
            if st.button("🗑️ Elimina", disabled=not conferma, key=f"del_{tabella}_{rid}"):
                db.execute(f"DELETE FROM {tabella} WHERE id=?", (rid,))
                st.session_state["_flash"] = "Elemento eliminato."
                st.rerun()


def config_crud(db):
    imp = db.get_imp()
    return {
        PAGINA_CLIENTI: dict(
            tabella="clienti", titolo="👥 Clienti", singolare="cliente",
            campi=[campo("nome", "Nome *", obbl=True), campo("telefono", "Telefono"), campo("email", "Email"),
                   campo("pec", "PEC"), campo("indirizzo", "Indirizzo"), campo("piva", "Partita IVA"),
                   campo("note", "Note", "area")],
            colonne=[("id", "ID", None), ("nome", "Nome", None), ("telefono", "Telefono", None),
                     ("email", "Email", None), ("pec", "PEC", None), ("indirizzo", "Indirizzo", None),
                     ("piva", "P.IVA", None), ("note", "Note", None)],
            ricerca=["nome", "email", "piva", "telefono"],
            label=lambda r: r["nome"],
        ),
        PAGINA_MATERIALI: dict(
            tabella="materiali", titolo="📦 Materiali", singolare="materiale", prefisso="MAT",
            campi=[campo("codice", "Codice (vuoto = automatico)"), campo("nome", "Nome *", obbl=True),
                   campo("prezzo", "Prezzo acquisto (€) *", "float", pos=True, default=1.0),
                   campo("unita", "U.M. (es. m², m, pz, kg)", default="m²"),
                   campo("sfrido", "Sfrido %", "float", step=1.0, default=float(imp.get("sfrido_generale", 10.0)))],
            colonne=[("id", "ID", None), ("codice", "Codice", None), ("nome", "Nome materiale", None),
                     ("prezzo", "Prezzo acquisto", eur), ("unita", "U.M.", None),
                     ("sfrido", "Sfrido", lambda v: f"{v:.0f}%")],
            ricerca=["nome", "codice"],
            label=lambda r: f"{r['codice'] or '-'} · {r['nome']}",
        ),
        PAGINA_LAVORAZIONI: dict(
            tabella="lavorazioni", titolo="🛠️ Lavorazioni", singolare="lavorazione", prefisso="LAV",
            campi=[campo("codice", "Codice (vuoto = automatico)"), campo("nome", "Nome *", obbl=True),
                   campo("costo_orario", "Costo orario (€/h) *", "float", default=40.0),
                   campo("unita", "U.M. (es. h, m², pz)", default="h"),
                   campo("ore_um", "Ore per U.M. *", "float", pos=True, step=0.25, default=1.0)],
            colonne=[("id", "ID", None), ("codice", "Codice", None), ("nome", "Nome lavorazione", None),
                     ("costo_orario", "Costo orario", lambda v: eur(v) + "/h"), ("unita", "U.M.", None),
                     ("ore_um", "Ore/U.M.", lambda v: f"{v:.2f} h")],
            ricerca=["nome", "codice"],
            label=lambda r: f"{r['codice'] or '-'} · {r['nome']}",
        ),
    }


# ---- Impostazioni -----------------------------------------------------------
def pagina_impostazioni(db):
    st.header("⚙️ Impostazioni")
    rows = db.query("SELECT chiave, valore, unita FROM impostazioni ORDER BY rowid")
    st.caption("Parametri generali, paga oraria e ricarichi.")
    with st.form("form_impostazioni"):
        cols = st.columns(2)
        valori = {}
        for i, r in enumerate(rows):
            etichetta = r["chiave"].replace("_", " ").capitalize() + (f" ({r['unita']})" if r["unita"] else "")
            valori[r["chiave"]] = cols[i % 2].number_input(etichetta, min_value=0.0, value=float(r["valore"]),
                                                            step=0.5, format="%.2f", key=f"imp_{r['chiave']}")
        salva = st.form_submit_button("💾 Salva impostazioni")
    if salva:
        with db.conn() as con:
            for k, v in valori.items():
                con.execute("UPDATE impostazioni SET valore=? WHERE chiave=?", (v, k))
        st.session_state["_flash"] = "Impostazioni aggiornate."
        st.rerun()

    with st.expander("ℹ️ Guida e formule di calcolo"):
        st.markdown(
            "- **Costo materiale** = Prezzo acquisto × Q.tà × (1 + Sfrido / 100)\n"
            "- **Costo lavorazione / manodopera** = Costo orario × Ore\n"
            "- **Costo trasporto** = Km × Costo al km\n"
            "- **Prezzo di vendita di una riga** = Costo × (1 + Ricarico di categoria / 100) × (1 + Markup finale / 100). "
            "Il ricarico è quello dei materiali per i materiali, quello della manodopera per lavorazioni e personale; "
            "il trasporto riceve solo il markup finale.\n"
            "- **Imponibile** = somma dei prezzi di vendita delle righe\n"
            "- **IVA** = Imponibile × IVA %\n"
            "- **Spese generali**: non entrano nel prezzo, servono a calcolare il costo aziendale e quindi il margine reale "
            "(Costi diretti × (1 + Spese generali / 100)).\n"
            "- **Sfrido generale** è il valore proposto quando si crea un nuovo materiale."
        )


# ---- Preventivi (elenco e azioni) ---------------------------------------------
def pagina_preventivi(db):
    st.header("📋 Preventivi")
    st.button("➕ Nuovo preventivo", on_click=nuova_bozza)
    rows = db.query("SELECT * FROM preventivi ORDER BY id DESC")
    if not rows:
        st.info("Non ci sono ancora preventivi. Creane uno con «Nuovo preventivo».")
        return

    st.dataframe(pd.DataFrame([{
        "ID": r["id"], "Titolo": r["titolo"], "Rev.": f"Rev.{r['revisione'] or 0}", "Cliente": r["cliente"],
        "Imponibile": eur(r["imponibile"] or 0), "Ivato": eur(r["ivato"] or 0), "Stato": r["stato"],
        "Consegna": fmt_data(r["data_consegna"]),
    } for r in rows]), hide_index=True)

    st.divider()
    per_id = {r["id"]: r for r in rows}
    pid = st.selectbox("Seleziona un preventivo", list(per_id), key="prev_sel",
                       format_func=lambda i: f"#{i} · {per_id[i]['titolo']} · Rev.{per_id[i]['revisione'] or 0} · {per_id[i]['cliente']}")
    p = per_id[pid]

    st.button("📋 Modifica / Revisiona (crea una nuova revisione)", on_click=carica_revisione, args=(pid,))

    st.subheader(f"🔍 Dettaglio #{p['id']} (Rev.{p['revisione'] or 0}) · {p['titolo']}")
    st.caption(f"Cliente: {p['cliente']} · Stato: {p['stato']} · Inizio: {fmt_data(p['data_inizio'])} · "
               f"Consegna: {fmt_data(p['data_consegna'])}")
    voci = db.query("SELECT tipo, nome, qta, um, ore, prezzo_vendita FROM voci_preventivo WHERE preventivo_id=? ORDER BY id", (pid,))
    if voci:
        st.dataframe(pd.DataFrame([{
            "Tipo": v["tipo"], "Nome": v["nome"], "Q.tà": num(v["qta"] or 0), "U.M.": v["um"],
            "Ore": f"{num(v['ore'])} h" if v["ore"] else "-", "Prezzo di vendita": eur(v["prezzo_vendita"] or 0),
        } for v in voci]), hide_index=True)
    c1, c2, c3 = st.columns(3)
    c1.metric("Imponibile", eur(p["imponibile"] or 0))
    c2.metric("IVA", eur(p["iva"] or 0))
    c3.metric("Totale ivato", eur(p["ivato"] or 0))

    st.subheader("📥 Esporta")
    nome_base = f"Preventivo_{pid}_{nome_file_sicuro(p['titolo'])}"
    e1, e2, e3 = st.columns(3)
    if REPORTLAB_AVAILABLE:
        e1.download_button("📄 Scarica PDF", genera_pdf(db, pid), file_name=f"{nome_base}.pdf",
                           mime="application/pdf", key=f"pdf_{pid}")
    else:
        e1.error("ReportLab non installato: PDF non disponibile.")
    e2.download_button("📊 Scarica Excel", genera_excel(db, pid), file_name=f"{nome_base}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"xls_{pid}")
    ordine = genera_ordine_materiali(db, pid)
    if ordine:
        e3.download_button("📦 Ordine materiali", ordine, file_name=f"Ordine_Materiali_Prev_{pid}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"ord_{pid}")
    else:
        e3.caption("Nessun materiale in questo preventivo.")

    st.subheader("🔄 Stato")
    s1, s2 = st.columns([3, 1], vertical_alignment="bottom")
    nuovo_stato = s1.selectbox("Stato del preventivo", STATI, index=STATI.index(p["stato"]) if p["stato"] in STATI else 0,
                               key=f"stato_{pid}_{p['stato']}")
    if s2.button("Aggiorna stato", key=f"btn_stato_{pid}"):
        db.execute("UPDATE preventivi SET stato=? WHERE id=?", (nuovo_stato, pid))
        st.session_state["_flash"] = f"Stato del preventivo #{pid} aggiornato a «{nuovo_stato}»."
        st.rerun()

    with st.expander("🗑️ Elimina preventivo"):
        conferma = st.checkbox("Confermo di voler eliminare questo preventivo", key=f"conf_del_prev_{pid}")
        if st.button("Elimina definitivamente", disabled=not conferma, key=f"del_prev_{pid}"):
            db.execute("DELETE FROM preventivi WHERE id=?", (pid,))
            st.session_state["_flash"] = f"Preventivo #{pid} eliminato."
            st.rerun()


# ---- Composizione preventivo --------------------------------------------------
def pagina_composer(db):
    ss = st.session_state
    imp = db.get_imp()
    st.header("📝 Composizione / revisione preventivo")
    if ss["d_padre"]:
        st.info(f"Stai creando la **Rev.{ss['d_rev']}** del preventivo #{ss['d_padre']}. "
                "Modifica quello che serve e salva: il preventivo originale resta invariato.")
        st.button("Annulla revisione e parti da zero", on_click=nuova_bozza)

    # -- dati generali
    clienti = [r["nome"] for r in db.query("SELECT nome FROM clienti ORDER BY nome")]
    if ss.get("d_cliente") and ss["d_cliente"] not in clienti:
        clienti.append(ss["d_cliente"])  # cliente della revisione non più in anagrafica
    with st.container(border=True):
        st.subheader("Dati generali")
        g1, g2 = st.columns(2)
        g1.text_input("Titolo progetto *", key="d_titolo")
        g2.selectbox("Cliente *", clienti, index=None, placeholder="Scegli un cliente…", key="d_cliente")
        if not clienti:
            g2.caption("Nessun cliente in anagrafica: aggiungilo dalla pagina «Clienti».")
        g3, g4 = st.columns(2)
        g3.date_input("Inizio lavori", key="d_inizio", format="DD/MM/YYYY")
        g4.date_input("Consegna", key="d_consegna", format="DD/MM/YYYY")

    # -- materiali
    materiali = {m["id"]: m for m in db.query("SELECT * FROM materiali ORDER BY nome")}
    with st.container(border=True):
        st.subheader("Aggiungi materiale")
        if not materiali:
            st.caption("Il listino materiali è vuoto: aggiungi i materiali dalla pagina «Materiali».")
        else:
            m1, m2, m3 = st.columns([3, 1, 1], vertical_alignment="bottom")
            sel = m1.selectbox("Materiale", list(materiali), index=None, placeholder="Scegli un materiale…",
                               key="d_mat_sel",
                               format_func=lambda i: f"{materiali[i]['nome']} ({materiali[i]['unita']}) · {eur(materiali[i]['prezzo'])}")
            m2.number_input("Quantità", min_value=0.01, step=1.0, format="%.2f", key="d_mat_qta")
            if m3.button("➕ Aggiungi", key="btn_add_mat"):
                if sel is None:
                    st.warning("Scegli prima un materiale.")
                else:
                    m = materiali[sel]
                    qta = float(ss["d_mat_qta"])
                    costo = m["prezzo"] * qta * (1 + (m["sfrido"] or 0) / 100.0)
                    ss["d_voci"].append({"tipo": TIPO_MATERIALE, "nome": m["nome"], "qta": qta, "um": m["unita"],
                                         "ore": 0.0, "costo_base": costo})

    # -- lavorazioni
    lavorazioni = {l["id"]: l for l in db.query("SELECT * FROM lavorazioni ORDER BY nome")}

    def aggiorna_ore_lavorazione():
        lid = ss.get("d_lav_sel")
        if lid in lavorazioni:
            ss["d_lav_ore"] = round(float(ss["d_lav_qta"]) * float(lavorazioni[lid]["ore_um"] or 1.0), 2)

    with st.container(border=True):
        st.subheader("Aggiungi lavorazione di laboratorio")
        if not lavorazioni:
            st.caption("Il listino lavorazioni è vuoto: aggiungile dalla pagina «Lavorazioni».")
        else:
            l1, l2, l3, l4 = st.columns([3, 1, 1, 1], vertical_alignment="bottom")
            sel_l = l1.selectbox("Lavorazione", list(lavorazioni), index=None, placeholder="Scegli una lavorazione…",
                                 key="d_lav_sel", on_change=aggiorna_ore_lavorazione,
                                 format_func=lambda i: f"{lavorazioni[i]['nome']} ({lavorazioni[i]['unita']}) · {eur(lavorazioni[i]['costo_orario'])}/h")
            l2.number_input("Quantità", min_value=0.01, step=1.0, format="%.2f", key="d_lav_qta",
                            on_change=aggiorna_ore_lavorazione)
            l3.number_input("Ore totali", min_value=0.0, step=0.5, format="%.2f", key="d_lav_ore",
                            help="Proposte come Quantità × Ore/U.M. della lavorazione; puoi modificarle.")
            if l4.button("➕ Aggiungi", key="btn_add_lav"):
                if sel_l is None:
                    st.warning("Scegli prima una lavorazione.")
                elif float(ss["d_lav_ore"]) <= 0:
                    st.warning("Le ore devono essere maggiori di zero.")
                else:
                    l = lavorazioni[sel_l]
                    ore = float(ss["d_lav_ore"])
                    ss["d_voci"].append({"tipo": TIPO_LAVORAZIONE, "nome": l["nome"], "qta": float(ss["d_lav_qta"]),
                                         "um": l["unita"], "ore": ore, "costo_base": ore * l["costo_orario"]})

    # -- voci inserite
    with st.container(border=True):
        st.subheader("Voci inserite")
        voci = ss["d_voci"]
        if voci:
            calc_voci = calcola(voci, {}, 0, imp)["righe"]
            st.dataframe(pd.DataFrame([{
                "N.": i + 1, "Tipo": r["tipo"], "Nome": r["nome"], "Q.tà": num(r["qta"]), "U.M.": r["um"],
                "Ore": f"{num(r['ore'])} h" if r["ore"] else "-", "Costo": eur(r["costo_base"]),
                "Vendita": eur(r["prezzo"]),
            } for i, r in enumerate(calc_voci)]), hide_index=True)
            st.multiselect("Voci da rimuovere", list(range(len(voci))), key="d_rimuovi",
                           format_func=lambda i: f"{i + 1} · {voci[i]['nome']}")
            st.button("🗑️ Rimuovi voci selezionate", on_click=rimuovi_voci_selezionate)
        else:
            st.caption("Nessuna voce inserita.")

    # -- personale e trasporto
    with st.container(border=True):
        st.subheader("Cantiere, trasporto e personale operativo")
        cols = st.columns(3)
        for i, (chiave, etichetta, _, _, _) in enumerate(RUOLI):
            cols[i % 3].number_input(f"Ore {etichetta}", min_value=0.0, step=0.5, format="%.2f", key=f"d_ore_{chiave}")
        st.number_input("Km trasporto A/R", min_value=0.0, step=1.0, format="%.1f", key="d_km")

    # -- riepilogo
    ore_ruoli = {r[0]: ss.get(f"d_ore_{r[0]}", 0.0) for r in RUOLI}
    calc = calcola(voci, ore_ruoli, float(ss.get("d_km") or 0), imp)
    with st.container(border=True):
        st.subheader("Riepilogo")
        r1, r2, r3, r4 = st.columns(4)
        r1.metric("Imponibile", eur(calc["imponibile"]))
        r2.metric(f"IVA ({num(calc['iva_perc'])}%)", eur(calc["iva"]))
        r3.metric("Totale ivato", eur(calc["ivato"]))
        r4.metric("Margine reale", f"{calc['margine']:.1f}%")
        if calc["righe"]:
            with st.expander("Dettaglio righe del preventivo"):
                st.dataframe(pd.DataFrame([{
                    "Tipo": r["tipo"], "Descrizione": r["nome"], "Q.tà": num(r["qta"]), "U.M.": r["um"],
                    "Costo": eur(r["costo_base"]), "Prezzo": eur(r["prezzo"]),
                } for r in calc["righe"]]), hide_index=True)
        err = ss.pop("_err_comp", None)
        if err:
            st.error(err)
        st.button("💾 Calcola e salva preventivo", type="primary", on_click=salva_bozza)


# ---- Dati e backup ------------------------------------------------------------
def pagina_dati(db):
    st.header("💾 Dati e backup")
    st.warning(
        "**Attenzione:** su Streamlit Community Cloud il disco non è permanente. Se l'app viene riavviata, "
        "ripubblicata o resta a lungo inattiva, i dati inseriti possono andare persi. "
        "**Scarica un backup a fine giornata** e, se serve, ripristinalo da qui sotto."
    )
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    st.download_button("⬇️ Scarica backup completo (database)", db.esporta_bytes(),
                       file_name=f"gestionale_scenografia_{stamp}.db", mime="application/octet-stream")

    ver = st.session_state.setdefault("_up_ver", 0)
    st.subheader("Ripristina da backup")
    f = st.file_uploader("Carica un file .db (anche quello creato dalla versione desktop)", type=["db", "sqlite", "sqlite3"],
                         key=f"up_db_{ver}")
    if f is not None:
        st.caption("Il database attuale verrà sostituito (ne viene fatta prima una copia di sicurezza).")
        ok_ripr = st.checkbox("Confermo il ripristino", key=f"conf_ripr_{ver}")
        if st.button("Ripristina", disabled=not ok_ripr):
            ok, msg = db.ripristina_da_bytes(f.getvalue())
            if ok:
                st.session_state["_up_ver"] = ver + 1
                st.session_state["_flash"] = msg
                st.rerun()
            else:
                st.error(msg)

    st.subheader("Importa listino da Excel")
    st.caption("Fogli attesi: «Listino_Materiali» e «Listino_Lavorazioni». I codici già presenti vengono saltati.")
    x = st.file_uploader("Carica il listino (.xlsx)", type=["xlsx"], key=f"up_xls_{ver}")
    if x is not None and st.button("Importa listino"):
        try:
            n_mat, n_lav = db.importa_listino(x)
            st.session_state["_up_ver"] = ver + 1
            st.session_state["_flash"] = f"Importati {n_mat} materiali e {n_lav} lavorazioni."
            st.rerun()
        except Exception as e:
            st.error(f"Impossibile leggere il file: {e}")


# ==============================================================================
# 6. APPLICAZIONE
# ==============================================================================
def check_password():
    """Se nei Secrets è definita APP_PASSWORD, l'app richiede la password."""
    try:
        password = st.secrets["APP_PASSWORD"]
    except Exception:
        return True
    if st.session_state.get("_auth"):
        return True
    st.title("🔒 Gestionale Preventivi")
    with st.form("login"):
        inserita = st.text_input("Password", type="password")
        entra = st.form_submit_button("Entra")
    if entra:
        if hmac.compare_digest(inserita.encode(), str(password).encode()):
            st.session_state["_auth"] = True
            st.rerun()
        else:
            st.error("Password errata.")
    return False


def main():
    st.set_page_config(page_title="Gestionale Preventivi Scenografici", page_icon="🎭", layout="wide")
    if not check_password():
        st.stop()

    db = get_db()
    keep_draft_alive()
    ensure_draft()

    with st.sidebar:
        st.title("🎭 Gestionale Preventivi")
        st.caption("Scenografia cinema e teatro")
        st.radio("Menu", PAGINE, key="pagina", label_visibility="collapsed")

    mostra_flash()
    pagina = st.session_state.get("pagina", PAGINA_DASHBOARD)
    if pagina == PAGINA_DASHBOARD:
        pagina_dashboard(db)
    elif pagina == PAGINA_PREVENTIVI:
        pagina_preventivi(db)
    elif pagina == PAGINA_COMPOSER:
        pagina_composer(db)
    elif pagina in (PAGINA_CLIENTI, PAGINA_MATERIALI, PAGINA_LAVORAZIONI):
        pagina_crud(db, config_crud(db)[pagina])
    elif pagina == PAGINA_IMPOSTAZIONI:
        pagina_impostazioni(db)
    elif pagina == PAGINA_DATI:
        pagina_dati(db)


if __name__ == "__main__":
    main()
