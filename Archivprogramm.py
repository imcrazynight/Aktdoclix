import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import sqlite3
import os
import sys
import re
import csv
import shutil
import json
import subprocess
import platform

# --- KONFIGURATION ---
DB_NAME = "Aktdoclix.db"
BACKUP_NAME = "archiv_v46_sicherung.db"
SETTINGS_FILE = "settings.json"

class ArchivApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Aktdoclix")

        # --- SYSTEM INITIALISIERUNG ---
        self.init_system()

        # --- FENSTER AUF LINKE HÄLFTE SETZEN ---
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        new_width = int(screen_width / 2)
        self.root.geometry(f"{new_width}x{screen_height-70}+0+0")
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(3, weight=1)

        # PFAD PRÜFUNG
        if getattr(sys, 'frozen', False):
            application_path = os.path.dirname(sys.executable)
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))

        self.base_path = os.path.join(application_path, "Datenbank")
        self.settings_path = os.path.join(application_path, SETTINGS_FILE)
        
        if not os.path.exists(self.base_path):
            try: os.makedirs(self.base_path)
            except Exception as e: messagebox.showerror("Fehler", f"Konnte DB-Ordner nicht erstellen: {e}")
            
        self.var_clear_fields = tk.BooleanVar(value=True) 
        self.var_details = tk.BooleanVar(value=False) 

        # --- EINSTELLUNGEN LADEN ---
        self.settings = self.load_settings()
        self.custom_history = self.settings["custom_btn_history"]
        self.custom_btn_text = self.custom_history[0] if self.custom_history else "Kassentagebuch"
        self.lagerorte = self.settings["lagerorte"]
        self.kat_map = self.settings["kategorien"]

        # --- LOGIK ---
        self.titel_logik = {
            "Gemeinde": "Gemeinderechnung Oberschwaningen",
            "Kirche": "Kirchenrechnung Oberschwaningen",
            "Schule": "Schulrechnung Oberschwaningen",
            "Armenpflege": "Armenpflegerechnung Oberschwaningen",
            "Sonderakten": "Rechnung (Sonderakte)"
        }
        
        # Timer Variable für die Suche
        self.search_timer = None

        self.setup_ui()
        self.on_kat_change(None)
        self.load_data()

    def init_system(self):
        """Erstellt Backup und initialisiert DB"""
        if os.path.exists(DB_NAME):
            try: shutil.copy2(DB_NAME, BACKUP_NAME)
            except: pass

        self.run_query('''CREATE TABLE IF NOT EXISTS akten 
                 (id INTEGER PRIMARY KEY, signatur TEXT, titel TEXT, zeitraum TEXT, 
                  typ TEXT, anzahl INTEGER, kategorie TEXT, unterkat TEXT, 
                  zustand TEXT, schlagworte TEXT, notizen TEXT, pfad TEXT, lagerort TEXT)''')

    def run_query(self, sql, params=(), fetch=False, commit=True):
        """Hilfsmethode für Datenbankzugriffe (Context Manager)"""
        try:
            with sqlite3.connect(DB_NAME) as conn:
                c = conn.cursor()
                c.execute(sql, params)
                if commit:
                    conn.commit()
                if fetch:
                    return c.fetchall()
                return c.lastrowid
        except sqlite3.Error as e:
            messagebox.showerror("Datenbankfehler", f"SQL Fehler: {e}")
            return [] if fetch else None

    def setup_ui(self):
        # --- TURBO LEISTE ---
        self.turbo_frame = tk.Frame(self.root, bg="#e3f2fd", pady=10)
        self.turbo_frame.pack(side=tk.TOP, fill=tk.X)
        
        tk.Label(self.turbo_frame, text="SCHNELL-EINGABE:", bg="#e3f2fd", font=("Arial", 10, "bold")).pack(side=tk.LEFT, padx=10)
        
        self.create_turbo_btn(self.turbo_frame, "RECHNUNG", "#4caf50", self.click_rechnung)
        self.create_turbo_btn(self.turbo_frame, "BELEGE", "#00897b", self.click_beleg)
        
        self.btn_custom = self.create_turbo_btn(self.turbo_frame, self.custom_btn_text, "#b3e5fc", self.click_custom, fg="black")
        self.btn_custom.bind("<Button-3>", self.open_custom_menu) 

        self.create_turbo_btn(self.turbo_frame, "DUPLIKAT", "#795548", self.click_duplikat)

        # Jalousie-Button
        self.btn_toggle = tk.Button(self.turbo_frame, text="🔼 Maske ausblenden", 
                                    command=self.toggle_input_mask, bg="#e3f2fd", relief="flat")
        self.btn_toggle.pack(side=tk.RIGHT, padx=10)

        # --- EINGABEBEREICH ---
        self.input_frame = tk.LabelFrame(self.root, text=" Neue Akte erfassen ", padx=10, pady=10)
        self.input_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5)
        
        # Grid Config
        for i in range(4): self.input_frame.columnconfigure(i, weight=0)
        self.input_frame.columnconfigure(4, weight=1)

        # Zeile 0
        tk.Label(self.input_frame, text="Kategorie:").grid(row=0, column=0, sticky="w")
        self.combo_kat = ttk.Combobox(self.input_frame, values=list(self.kat_map.keys()), width=30)
        self.combo_kat.current(0)
        self.combo_kat.grid(row=0, column=1, sticky="w", padx=5)
        self.combo_kat.bind("<KeyRelease>", self.on_kat_type_live) 
        self.combo_kat.bind("<<ComboboxSelected>>", self.on_kat_change)
        self.combo_kat.bind("<Return>", lambda e: self.ent_titel.focus_set())

        tk.Label(self.input_frame, text="Signatur:").grid(row=0, column=2, sticky="w", padx=(10, 2))
        self.ent_sig = tk.Entry(self.input_frame, width=20)
        self.ent_sig.grid(row=0, column=3, sticky="w")

        # Zeile 1
        tk.Label(self.input_frame, text="Titel:").grid(row=1, column=0, sticky="w")
        self.ent_titel = ttk.Combobox(self.input_frame, values=[], width=50) 
        self.ent_titel.grid(row=1, column=1, sticky="w", padx=5, pady=5)
        self.ent_titel.bind("<Return>", lambda e: self.ent_zeit.focus_set())

        tk.Label(self.input_frame, text="Zeitraum:").grid(row=1, column=2, sticky="w", padx=(10, 2))
        self.ent_zeit = tk.Entry(self.input_frame, width=20, bg="#fff3e0")
        self.ent_zeit.grid(row=1, column=3, sticky="w")
        self.ent_zeit.bind("<FocusOut>", self.auto_complete_year)
        self.ent_zeit.bind("<Return>", self.auto_complete_year) 

        # Zeile 2
        tk.Label(self.input_frame, text="Art:").grid(row=2, column=0, sticky="w")
        art_frame = tk.Frame(self.input_frame)
        art_frame.grid(row=2, column=1, sticky="w", padx=5)
        
        self.combo_typ = ttk.Combobox(art_frame, values=["Einzelheft", "Buch", "Sammelband", "Ordner", "Karten und Pläne", "Urkunden"], state="readonly", width=15)
        self.combo_typ.set("Einzelheft")
        self.combo_typ.pack(side=tk.LEFT)
        self.combo_typ.bind("<<ComboboxSelected>>", self.toggle_fields)
        self.combo_typ.bind("<Return>", self.on_typ_enter)

        self.lbl_anzahl = tk.Label(art_frame, text="Anzahl:")
        self.ent_anzahl = tk.Entry(art_frame, width=5)
        self.ent_anzahl.bind("<Return>", lambda e: self.combo_zustand.focus_set())

        tk.Label(self.input_frame, text="Zustand:").grid(row=2, column=2, sticky="w", padx=(10, 2))
        self.combo_zustand = ttk.Combobox(self.input_frame, 
                                          values=["Stabil", "Leicht beschädigt", "Stark beschädigt", "Nicht benutzbar"], 
                                          state="readonly", width=20)
        self.combo_zustand.set("Stabil")
        self.combo_zustand.grid(row=2, column=3, sticky="w")
        self.combo_zustand.bind("<Return>", lambda e: self.combo_lager.focus_set())

        # Zeile 3
        tk.Label(self.input_frame, text="Lagerort:").grid(row=3, column=0, sticky="w")
        self.combo_lager = ttk.Combobox(self.input_frame, values=self.lagerorte, width=30)
        self.combo_lager.grid(row=3, column=1, sticky="w", padx=5, pady=5)
        self.combo_lager.bind("<Return>", self.on_lager_enter)

        self.lbl_sonder = tk.Label(self.input_frame, text="Unterkategorie:")
        self.sonder_themen = ["Infrastruktur", "Militär & Krieg", "Flurbereinigung", "Verwaltung", "Wirtschaft", "Justiz", "Sonstiges"]
        self.combo_sonder = ttk.Combobox(self.input_frame, values=self.sonder_themen, state="readonly", width=20)
        self.combo_sonder.bind("<Return>", lambda e: self.txt_notizen.focus_set())
        
        # Zeile 4 - Details
        self.chk_details = tk.Checkbutton(self.input_frame, text="Detail: Ort/Personen", 
                                          variable=self.var_details, command=self.toggle_details_input)
        self.chk_details.grid(row=4, column=0, sticky="w", padx=0, pady=5)

        self.detail_frame = tk.Frame(self.input_frame)
        tk.Label(self.detail_frame, text="Ort:").pack(side=tk.LEFT, padx=(5, 2))
        self.ent_ort = tk.Entry(self.detail_frame, width=20)
        self.ent_ort.pack(side=tk.LEFT, padx=2)
        tk.Label(self.detail_frame, text="Personen:").pack(side=tk.LEFT, padx=(10, 2))
        self.ent_pers = tk.Entry(self.detail_frame, width=20)
        self.ent_pers.pack(side=tk.LEFT, padx=2)

        # Zeile 5: PFAD
        tk.Label(self.input_frame, text="Pfad (Auto):").grid(row=5, column=0, sticky="w")
        self.ent_pfad = tk.Entry(self.input_frame, width=65, bg="#f0f0f0") 
        self.ent_pfad.grid(row=5, column=1, columnspan=2, sticky="w", padx=5, pady=5)
        tk.Button(self.input_frame, text="Pfad manuell", command=self.browse_folder).grid(row=5, column=3, sticky="w")

        # Zeile 6: Notizen
        tk.Label(self.input_frame, text="Notizen:").grid(row=6, column=0, sticky="nw")
        self.txt_notizen = tk.Text(self.input_frame, height=3, width=85)
        self.txt_notizen.grid(row=6, column=1, columnspan=3, sticky="w", padx=5, pady=5)

        # Checkbox & Speichern
        self.chk_clear = tk.Checkbutton(self.input_frame, text="Felder nach Speichern leeren", variable=self.var_clear_fields)
        self.chk_clear.grid(row=7, column=1, sticky="w", pady=5)

        btn_add = tk.Button(self.input_frame, text="AKTE SPEICHERN", command=self.add_entry, bg="#1565c0", fg="white", font=("Arial", 11, "bold"))
        btn_add.grid(row=7, column=3, pady=10, sticky="e")

        # --- LISTE ---
        search_frame = tk.Frame(self.root, padx=10, pady=5, bg="#fff9c4")
        search_frame.pack(side=tk.TOP, fill=tk.X)
        tk.Label(search_frame, text="Intelligente Suche:", bg="#fff9c4").pack(side=tk.LEFT)
        self.ent_search = tk.Entry(search_frame)
        self.ent_search.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        
        self.ent_search.bind("<KeyRelease>", self.start_search_timer)
        
        btn_export = tk.Button(search_frame, text="✈ Exportieren", command=self.open_export_window, bg="#ff9800", fg="black", font=("Arial", 10, "bold"))
        btn_export.pack(side=tk.LEFT, padx=5)

        list_frame = tk.Frame(self.root)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # NEUE SPALTE: "Scan"
        cols = ("ID", "Signatur", "Titel", "Zeit", "Lagerort", "Zustand", "Scan")
        self.tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        
        # Spaltenkonfiguration (Breiten angepasst)
        widths = [30, 110, 300, 80, 150, 120, 50]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor=tk.CENTER if col=="Scan" else tk.W)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree.tag_configure("ok", foreground="black")
        self.tree.tag_configure("warnung", foreground="black") 
        self.tree.tag_configure("kritisch", foreground="red")
        self.tree.tag_configure("gesperrt", foreground="gray") 

        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)

        self.lbl_status = tk.Label(self.root, text="System bereit", bd=1, relief=tk.SUNKEN, anchor=tk.W, fg="blue")
        self.lbl_status.pack(side=tk.BOTTOM, fill=tk.X)

    def toggle_input_mask(self):
        """Schaltet den Eingabebereich an und aus (Jalousie)"""
        if self.input_frame.winfo_viewable():
            self.input_frame.pack_forget()
            self.btn_toggle.config(text="🔽 Maske zeigen")
        else:
            self.input_frame.pack(side=tk.TOP, fill=tk.X, padx=10, pady=5, after=self.turbo_frame)
            self.btn_toggle.config(text="🔼 Maske ausblenden")

    def create_turbo_btn(self, parent, text, bg, cmd, fg="white"):
        btn = tk.Button(parent, text=text, bg=bg, fg=fg, font=("Arial", 10, "bold"), command=cmd)
        btn.pack(side=tk.LEFT, padx=5)
        return btn

    def open_file_safe(self, path):
        if not path or not os.path.exists(path):
            messagebox.showwarning("Fehler", "Pfad existiert nicht oder ist leer.")
            return
        try:
            if platform.system() == 'Windows': os.startfile(path)
            elif platform.system() == 'Darwin': subprocess.call(('open', path))
            else: subprocess.call(('xdg-open', path))
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte Ordner nicht öffnen: {e}")

    # --- SETTINGS MANAGEMENT ---
    def load_settings(self):
        defaults = {
            "custom_btn_history": ["Kassentagebuch", "Tagebuch", "Planungsbuch"],
            "lagerorte": ["Staatsarchiv Nürnberg", "Gemeindearchiv Unterschwaningen", "Schlossarchiv Dennenlohe", "Ortsarchiv Oberschwaningen"],
            "kategorien": {
                "Gemeinde": "Ge.Obs.", "Kirche": "Ki.Obs.", "Schule": "Schul.Obs.", 
                "Armenpflege": "Armenpf.Obs.", "Sonderakten": "Sonder."
            }
        }
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, "r", encoding="utf-8") as f:
                    return {**defaults, **json.load(f)}
            except: pass
        return defaults

    def save_settings(self):
        try:
            with open(self.settings_path, "w", encoding="utf-8") as f:
                json.dump({
                    "custom_btn_history": self.custom_history,
                    "lagerorte": self.lagerorte,
                    "kategorien": self.kat_map
                }, f, indent=4, ensure_ascii=False)
        except Exception as e: print(f"Settings Fehler: {e}")

    # --- BUTTON LOGIK ---
    def open_custom_menu(self, event):
        win = tk.Toplevel(self.root)
        win.title("Begriff wählen")
        win.geometry(f"300x130+{event.x_root}+{event.y_root}")
        tk.Label(win, text="Wählen oder neu eingeben:", font=("Arial", 9)).pack(pady=(10, 2))
        cb = ttk.Combobox(win, values=self.custom_history, width=30)
        cb.set(self.custom_btn_text)
        cb.pack(padx=10, pady=5); cb.focus_set()

        def confirm(event=None):
            val = cb.get().strip()
            if val:
                self.custom_btn_text = val
                self.btn_custom.config(text=val)
                if val not in self.custom_history:
                    self.custom_history.append(val)
                    self.save_settings()
                win.destroy()
        btn_ok = tk.Button(win, text="Übernehmen", bg="#b3e5fc", command=confirm)
        btn_ok.pack(pady=5)
        win.bind('<Return>', confirm); cb.bind('<Return>', confirm)

    def click_custom(self): self.turbo_logic(self.custom_btn_text)
    def click_rechnung(self): self.turbo_logic("Rechnung")
    def click_beleg(self): self.turbo_logic("Belege", prefix="Belege zur")
    def click_duplikat(self): self.turbo_logic("Duplikat", prefix="Duplikat zur")

    def turbo_logic(self, type_key, prefix=None):
        kat = self.combo_kat.get()
        base = self.titel_logik.get(kat, "Rechnung")
        if prefix:
             parts = base.split(" ")
             titel = f"{prefix} {parts[0]}"
        elif type_key == self.custom_btn_text:
             parts = base.split(" ")
             ort = parts[-1] if len(parts) > 1 else ""
             titel = f"{type_key} {ort}".strip()
        else:
             titel = base
        self.turbo_fill(titel)

    def turbo_fill(self, titel):
        self.ent_titel.set(titel)
        self.combo_typ.set("Einzelheft")
        self.ent_zeit.focus_set()
        self.ent_zeit.select_range(0, tk.END)

    # --- UI INTERAKTION ---
    def on_lager_enter(self, event):
        if self.combo_sonder.winfo_viewable(): self.combo_sonder.focus_set()
        else: self.txt_notizen.focus_set()
    
    def on_typ_enter(self, event):
        target = self.ent_anzahl if self.combo_typ.get() == "Sammelband" else self.combo_zustand
        target.focus_set()

    def auto_complete_year(self, event):
        text = self.ent_zeit.get().strip()
        if re.match(r'^\d{4}$', text):
            try:
                year = int(text)
                self.ent_zeit.delete(0, tk.END)
                self.ent_zeit.insert(0, f"{year}-{year+1}")
            except: pass
        self.combo_typ.focus_set()

    def toggle_details_input(self):
        if self.var_details.get(): self.detail_frame.grid(row=4, column=1, columnspan=4, sticky="w", padx=5)
        else: self.detail_frame.grid_remove()

    def toggle_fields(self, event=None):
        if self.combo_typ.get() == "Sammelband":
            self.lbl_anzahl.pack(side=tk.LEFT, padx=(10,5)); self.ent_anzahl.pack(side=tk.LEFT)
        else:
            self.lbl_anzahl.pack_forget(); self.ent_anzahl.pack_forget()
        
        if self.combo_kat.get() == "Sonderakten":
            self.lbl_sonder.grid(row=3, column=2, sticky="w", padx=(10, 2))
            self.combo_sonder.grid(row=3, column=3, sticky="w", padx=2)
        else:
            self.lbl_sonder.grid_forget(); self.combo_sonder.grid_forget()

    # --- SIGNATUR GENERIERUNG ---
    def on_kat_type_live(self, event):
        val = self.combo_kat.get()
        if val in self.kat_map:
            prefix = self.kat_map[val]
        else:
            prefix = (val[:3] + ".") if len(val) >= 3 else "Div."
        self.update_sig_field(prefix)

    def on_kat_change(self, event):
        prefix = self.kat_map.get(self.combo_kat.get(), "")
        self.update_sig_field(prefix)
        self.toggle_fields()

    def update_sig_field(self, prefix):
        next_sig = self.get_next_signature(prefix)
        self.ent_sig.delete(0, tk.END)
        self.ent_sig.insert(0, next_sig)

    def get_next_signature(self, prefix):
        if not prefix: prefix = "Div." 
        rows = self.run_query("SELECT signatur FROM akten WHERE signatur LIKE ? ORDER BY length(signatur) DESC, signatur DESC LIMIT 1", (prefix + "%",), fetch=True)
        if rows:
            match = re.search(r'(\d+)$', rows[0][0])
            if match: return f"{prefix}{int(match.group(1)) + 1:05d}"
        return f"{prefix}00001"

    # --- LISTE ACTIONS (Rechtsklick) ---
    def show_context_menu(self, event):
        item = self.tree.identify_row(event.y)
        if not item: return
        self.tree.selection_set(item)
        menu = tk.Menu(self.root, tearoff=0)
        
        # Rechtsklick: Nur noch Ordner & Kopieren
        menu.add_command(label="📂 Ordner öffnen", command=self.open_folder_from_list)
        menu.add_command(label="📋 Signatur kopieren", command=self.copy_sig_to_clipboard)
        
        menu.post(event.x_root, event.y_root)

    def open_folder_from_list(self):
        item = self.tree.selection()
        if not item: return
        item_id = self.tree.item(item)['values'][0]
        res = self.run_query("SELECT pfad FROM akten WHERE id=?", (item_id,), fetch=True)
        if res and res[0][0]: 
            self.open_file_safe(res[0][0])
        else:
            messagebox.showinfo("Info", "Kein Pfad hinterlegt.")

    def copy_sig_to_clipboard(self):
        item = self.tree.selection()
        if item:
            sig = self.tree.item(item)['values'][1]
            self.root.clipboard_clear(); self.root.clipboard_append(sig)

    def on_double_click(self, event):
        # LINKS-KLICK (Doppel) öffnet IMMER die gesperrte Detailansicht
        self.open_edit_window(None)

    def browse_folder(self):
        d = filedialog.askdirectory(initialdir=self.base_path)
        if d: self.ent_pfad.delete(0, tk.END); self.ent_pfad.insert(0, d)

    # --- HAUPTFUNKTION: SPEICHERN ---
    def add_entry(self):
        sig = self.ent_sig.get().strip()
        titel = self.ent_titel.get().strip()
        kat_input = self.combo_kat.get().strip()
        lager_input = self.combo_lager.get().strip()

        if not sig or not titel: return messagebox.showwarning("Fehler", "Signatur und Titel sind Pflichtfelder!")
        
        if lager_input and lager_input not in self.lagerorte:
            self.lagerorte.append(lager_input)
            self.combo_lager['values'] = self.lagerorte
            self.save_settings()

        if kat_input and kat_input not in self.kat_map:
            suggestion = kat_input[:3] + "."
            new_prefix = simpledialog.askstring("Neue Kategorie", f"Kürzel für '{kat_input}'?", initialvalue=suggestion, parent=self.root)
            prefix = new_prefix.strip() if new_prefix else suggestion
            self.kat_map[kat_input] = prefix
            self.combo_kat['values'] = list(self.kat_map.keys())
            self.save_settings()
            
            match = re.search(r'(\d+)$', sig)
            if match:
                sig = f"{prefix}{match.group(1)}"
                self.ent_sig.delete(0, tk.END); self.ent_sig.insert(0, sig)

        new_folder_path = os.path.join(self.base_path, sig)
        if not os.path.exists(new_folder_path):
            try: os.makedirs(new_folder_path)
            except OSError as e:
                messagebox.showerror("Fehler", f"Konnte Ordner nicht erstellen: {e}")
                return 
        
        self.ent_pfad.delete(0, tk.END); self.ent_pfad.insert(0, new_folder_path)

        notiz = self.txt_notizen.get("1.0", tk.END).strip()
        zusatz = []
        if self.ent_ort.get(): zusatz.append(f"Ort: {self.ent_ort.get()}")
        if self.ent_pers.get(): zusatz.append(f"Personen: {self.ent_pers.get()}")
        if zusatz: notiz = (notiz + "\n" + " | ".join(zusatz)).strip()

        sql = '''INSERT INTO akten (signatur, titel, zeitraum, typ, anzahl, kategorie, unterkat, zustand, schlagworte, notizen, pfad, lagerort) 
                 VALUES (?,?,?,?,?,?,?,?,?,?,?,?)'''
        params = (sig, titel, self.ent_zeit.get(), self.combo_typ.get(), self.ent_anzahl.get() or 1, 
                  kat_input, self.combo_sonder.get(), self.combo_zustand.get(), "", notiz, self.ent_pfad.get(), lager_input)
        
        self.run_query(sql, params)
        
        self.on_kat_change(None)
        
        if self.var_clear_fields.get():
            self.ent_titel.set(""); self.ent_zeit.delete(0, tk.END)
            self.txt_notizen.delete("1.0", tk.END); self.ent_pfad.delete(0, tk.END); self.combo_sonder.set("")
            self.ent_ort.delete(0, tk.END); self.ent_pers.delete(0, tk.END)
            self.var_details.set(False); self.toggle_details_input()
            self.combo_kat.focus_set()
        else: self.ent_titel.focus_set()
        
        self.load_data()
        self.lbl_status.config(text=f"Gespeichert: {sig}", fg="green")

    # --- VERBESSERTE SUCHE ---
    def start_search_timer(self, event):
        if self.search_timer:
            self.root.after_cancel(self.search_timer)
        self.search_timer = self.root.after(300, self.load_data)

    def load_data(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        term = self.ent_search.get().lower().strip()
        
        sql = "SELECT id, signatur, titel, zeitraum, kategorie, unterkat, zustand, pfad, schlagworte, notizen, lagerort, anzahl FROM akten"
        params = []
        
        if term:
            sql += """ WHERE (
                        lower(titel) LIKE ? OR 
                        lower(signatur) LIKE ? OR 
                        zeitraum LIKE ? OR 
                        lower(notizen) LIKE ? OR 
                        lower(lagerort) LIKE ? OR 
                        lower(unterkat) LIKE ? OR
                        lower(zustand) LIKE ? OR
                        lower(typ) LIKE ?
                       )"""
            wildcard = f"%{term}%"
            params = [wildcard] * 8
        
        # --- UPDATE: SORTIERUNG UMGEDREHT ---
        sql += " ORDER BY id DESC"

        rows = self.run_query(sql, params, fetch=True)
        
        for row in rows:
            zustand = row[6] if row[6] else "Unbekannt"
            pfad = row[7]
            
            tag = "ok"
            if "Leicht" in zustand: tag = "warnung"
            elif "Stark" in zustand: tag = "kritisch"
            elif "Nicht" in zustand: tag = "gesperrt"
            
            disp_zustand = f"{'🔴' if tag=='kritisch' else '🟡' if tag=='warnung' else '🟢'} {zustand}"
            
            # --- DATEI CHECK LOGIK ---
            scan_status = "⚠️"
            if pfad and os.path.exists(pfad):
                try:
                    # Prüfen ob Dateien drin sind (versteckte Dateien werden ignoriert)
                    files = [f for f in os.listdir(pfad) if not f.startswith('.')]
                    if len(files) > 0:
                        scan_status = "✅" # Voll
                    else:
                        scan_status = "⚪" # Leer
                except:
                    scan_status = "❓" # Fehler beim Zugriff
            
            # Werte in Tree einfügen (NEU: scan_status am Ende)
            self.tree.insert("", tk.END, values=(row[0], row[1], row[2], row[3], row[10], disp_zustand, scan_status), tags=(tag,))

    def open_export_window(self):
        exp_win = tk.Toplevel(self.root); exp_win.title("Export")
        exp_win.geometry("300x200")
        
        tk.Label(exp_win, text="Export Einstellungen", font=("Arial", 11, "bold")).pack(pady=10)
        
        var = tk.StringVar(value="all")
        tk.Radiobutton(exp_win, text="Alles exportieren", variable=var, value="all").pack(anchor="w", padx=20)
        tk.Radiobutton(exp_win, text="Nur aktuelle Kategorie:", variable=var, value="cat").pack(anchor="w", padx=20)
        
        cb = ttk.Combobox(exp_win, values=list(self.kat_map.keys()), state="readonly")
        cb.pack(padx=40, pady=5); cb.current(0)
        
        def go():
            f = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
            if not f: return
            
            sql = "SELECT * FROM akten"
            params = ()
            if var.get() == "cat":
                sql += " WHERE kategorie=?"
                params = (cb.get(),)
            
            rows = self.run_query(sql, params, fetch=True)
            
            try:
                with open(f, 'w', newline='', encoding='utf-8-sig') as file:
                    w = csv.writer(file, delimiter=';')
                    w.writerow(["ID","Sig","Titel","Zeit","Typ","Anz","Kat","SubKat","Zustand","Tags","Note","Pfad","Ort"])
                    w.writerows(rows)
                messagebox.showinfo("Erfolg", "Daten wurden exportiert."); exp_win.destroy()
            except Exception as e:
                messagebox.showerror("Fehler", f"Export fehlgeschlagen: {e}")

        tk.Button(exp_win, text="Export Starten", bg="#ff9800", command=go).pack(pady=15)

    # --- EDIT WINDOW (MIT SICHERHEITS-MODUS & NAVIGATIONS-SPRUNG & INTELLIGENTER SUCHE) ---
    def open_edit_window(self, event, specific_id=None):
        if specific_id:
            iid = specific_id
        else:
            item = self.tree.selection()
            if not item: return
            iid = self.tree.item(item)['values'][0]

        row = self.run_query("SELECT * FROM akten WHERE id=?", (iid,), fetch=True)
        if not row: return
        row = row[0]
        
        ew = tk.Toplevel(self.root)
        ew.title(f"Details: {row[1]}") # Start-Titel
        ew.geometry("700x750")
        ew.transient(self.root) 
        ew.grab_set() 
        
        ew.columnconfigure(1, weight=1)
        
        fields = {} 
        input_widgets = [] # Liste zum Sperren/Entsperren

        def mk_row(idx, label, val):
            tk.Label(ew, text=label).grid(row=idx, column=0, sticky="e", pady=5)
            e = tk.Entry(ew)
            e.insert(0, val if val else "")
            e.grid(row=idx, column=1, sticky="ew", padx=10)
            input_widgets.append(e) # Merken
            return e

        fields['id'] = mk_row(0, "ID:", str(row[0]))
        fields['sig'] = mk_row(1, "Signatur:", row[1])
        fields['titel'] = mk_row(2, "Titel:", row[2])
        fields['zeit'] = mk_row(3, "Zeitraum:", row[3])
        
        tk.Label(ew, text="Lagerort:").grid(row=4, column=0, sticky="e")
        fields['lager'] = ttk.Combobox(ew, values=self.lagerorte)
        fields['lager'].set(row[12] if len(row)>12 and row[12] else "")
        fields['lager'].grid(row=4, column=1, sticky="ew", padx=10)
        input_widgets.append(fields['lager']) # Merken
        
        tk.Label(ew, text="Unterkategorie:").grid(row=5, column=0, sticky="e")
        fields['sub'] = ttk.Combobox(ew, values=self.sonder_themen)
        fields['sub'].set(row[7] if row[7] else "")
        fields['sub'].grid(row=5, column=1, sticky="ew", padx=10)
        input_widgets.append(fields['sub']) # Merken

        tk.Label(ew, text="Zustand:").grid(row=6, column=0, sticky="e")
        fields['zustand'] = ttk.Combobox(ew, values=["Stabil", "Leicht beschädigt", "Stark beschädigt", "Nicht benutzbar"])
        fields['zustand'].set(row[8] if row[8] else "")
        fields['zustand'].grid(row=6, column=1, sticky="ew", padx=10)
        input_widgets.append(fields['zustand']) # Merken

        fields['anz'] = mk_row(7, "Anzahl:", str(row[5] or 1))

        tk.Label(ew, text="Notiz:").grid(row=8, column=0, sticky="ne")
        fields['notiz'] = tk.Text(ew, height=5)
        fields['notiz'].insert("1.0", row[10] if row[10] else "")
        fields['notiz'].grid(row=8, column=1, sticky="ew", padx=10)
        # Hinweis: Textfeld wird separat gesperrt

        fields['pfad'] = mk_row(9, "Pfad:", row[11])

        # --- ÄHNLICHE AKTEN ANZEIGE (INTERAKTIV & SCHLAU) ---
        tk.Label(ew, text="🔗 Ähnliche Akten (±5 Jahre vom Startjahr):", font=("Arial", 10, "bold"), fg="#1565c0").grid(row=10, column=0, columnspan=2, pady=(20, 5), sticky="w", padx=10)
        rel_frame = tk.Frame(ew)
        rel_frame.grid(row=11, column=0, columnspan=2, sticky="nsew", padx=10, pady=5)
        
        rel_list = ttk.Treeview(rel_frame, columns=("ID", "Sig", "Titel", "Zeit"), displaycolumns=("Sig", "Titel", "Zeit"), show="headings", height=5)
        
        rel_list.heading("Sig", text="Signatur"); rel_list.column("Sig", width=100)
        rel_list.heading("Titel", text="Titel"); rel_list.column("Titel", width=350)
        rel_list.heading("Zeit", text="Zeit"); rel_list.column("Zeit", width=80)
        rel_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        def on_related_double_click(event):
            selection = rel_list.selection()
            if not selection: return
            vals = rel_list.item(selection[0])['values']
            related_id = vals[0]
            ew.destroy()
            self.open_edit_window(None, specific_id=related_id)

        rel_list.bind("<Double-1>", on_related_double_click)

        # HELFER: Erstes Jahr extrahieren
        def get_start_year(text):
            match = re.search(r'\d{4}', str(text))
            if match: return int(match.group(0))
            return None

        # LOGIK: +/- 5 JAHRE VOM STARTJAHR
        my_start = get_start_year(row[3]) # row[3] ist Zeitraum
        
        if my_start:
             limit_low = my_start - 5
             limit_high = my_start + 5
             
             # Kandidaten holen (Gleiche Kategorie, nicht ich selbst)
             candidates = self.run_query("SELECT id, signatur, titel, zeitraum FROM akten WHERE kategorie=? AND id!=?", (row[6], iid), fetch=True)
             
             for cand in candidates:
                 cand_start = get_start_year(cand[3])
                 if cand_start:
                     # Ist das Startjahr des Kandidaten im Fenster?
                     if limit_low <= cand_start <= limit_high:
                         rel_list.insert("", tk.END, values=(cand[0], cand[1], cand[2], cand[3]))
        
        bf = tk.Frame(ew, pady=20); bf.grid(row=12, column=0, columnspan=2)

        # --- INTERNE FUNKTIONEN ---
        def save():
            try: new_id = int(fields['id'].get().strip())
            except ValueError: return messagebox.showerror("Fehler", "Die ID muss eine Zahl sein!")

            if new_id != iid:
                if self.run_query("SELECT id FROM akten WHERE id=?", (new_id,), fetch=True):
                    return messagebox.showerror("Fehler", "ID vergeben!")
            
            # Pfad-Rename Logik
            original_sig = row[1]
            new_sig = fields['sig'].get().strip()
            current_path = fields['pfad'].get().strip()
            if new_sig != original_sig and current_path and os.path.exists(current_path):
                if os.path.basename(current_path) == original_sig:
                    try:
                        new_path = os.path.join(os.path.dirname(current_path), new_sig)
                        os.rename(current_path, new_path)
                        fields['pfad'].delete(0, tk.END); fields['pfad'].insert(0, new_path)
                    except: pass

            sql = "UPDATE akten SET id=?, signatur=?, titel=?, zeitraum=?, lagerort=?, unterkat=?, zustand=?, anzahl=?, notizen=?, pfad=? WHERE id=?"
            params = (new_id, new_sig, fields['titel'].get(), fields['zeit'].get(), fields['lager'].get(), 
                      fields['sub'].get(), fields['zustand'].get(), fields['anz'].get(), 
                      fields['notiz'].get("1.0", tk.END).strip(), fields['pfad'].get(), iid)
            self.run_query(sql, params)
            self.load_data()
            ew.destroy()

        def delete():
            if messagebox.askyesno("Löschen?", "Eintrag wirklich löschen?"):
                self.run_query("DELETE FROM akten WHERE id=?", (iid,))
                self.load_data(); ew.destroy()

        # --- SICHERHEITS-LOGIK ---
        def enable_edit_mode():
            """Schaltet alle Felder frei"""
            ew.title(f"BEARBEITEN: {row[1]}")
            for w in input_widgets:
                if isinstance(w, ttk.Combobox):
                    w.config(state="readonly")
                else:
                    w.config(state="normal")
            
            fields['notiz'].config(state="normal", bg="white")
            
            btn_edit.pack_forget() 
            btn_save.pack(side=tk.LEFT, padx=10) 
            
        # 1. Startzustand: ALLES GESPERRT
        for w in input_widgets:
            if isinstance(w, ttk.Combobox):
                w.config(state="disabled")
            else:
                w.config(state="readonly")
        
        fields['notiz'].config(state="disabled", bg="#f0f0f0")

        # 2. Buttons
        btn_edit = tk.Button(bf, text="✏ Bearbeiten / Freigeben", bg="#ff9800", fg="black", font=("Arial", 10, "bold"), command=enable_edit_mode)
        btn_edit.pack(side=tk.LEFT, padx=10)

        btn_save = tk.Button(bf, text="💾 SPEICHERN", bg="#4caf50", fg="white", font=("Arial", 10, "bold"), command=save)
        
        tk.Button(bf, text="📂 Ordner öffnen", bg="#2196f3", fg="white", command=lambda: self.open_file_safe(fields['pfad'].get())).pack(side=tk.LEFT, padx=10)
        tk.Button(bf, text="🗑 Löschen", bg="#f44336", fg="white", command=delete).pack(side=tk.LEFT, padx=10)

if __name__ == "__main__":
    root = tk.Tk() 
    app = ArchivApp(root)
    root.mainloop()