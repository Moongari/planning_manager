import calendar as py_calendar

import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk, filedialog, simpledialog
import datetime
import sqlite3

from auth import get_db_connection, get_shift_max_load
from exporter import export_to_excel
from tkcalendar import DateEntry

# --- Constantes et données ---
SHIFTS = [
    ("07:00", "15:15"),
    ("15:15", "23:00"),
    ("23:00", "07:00")
]

MAX_SHIFT_LOADS = {
    0: 240,
    1: 300,
    2: 180
}

PERSONS = ["Alice", "Bob", "Charlie", "David"]
PERSON_COLORS = {
    "Alice": "#9c99ff",
    "Bob": "#99ff99",
    "Charlie": "#9999ff",
    "David": "#ffcc99"
}

TASK_NAMES = [
    "Maintenance",
    "Réparation",
    "Inspection",
    "Nettoyage",
    "Test",
    "Préparation",
    "Contrôle qualité",
    "Chambre",
    "Cuisine",
    "Inventaire"
]

DURATION_OPTIONS = [15, 30, 45, 60, 90, 120, 180]

# --- Fonctions utilitaires ---
def center_window(win, w, h):
    win.update_idletasks()
    ws = win.winfo_screenwidth()
    hs = win.winfo_screenheight()
    x = (ws // 2) - (w // 2)
    y = (hs // 2) - (h // 2)
    win.geometry(f"{w}x{h}+{x}+{y}")

def format_date(d):
    return d.strftime("%Y-%m-%d")

# --- Gestion des tâches ---
class TaskManager:
    def __init__(self):
        self.tasks = []

    def add_task(self, task):
        self.tasks.append(task)

    def remove_task(self, task):
        if task in self.tasks:
            self.tasks.remove(task)

    def get_tasks(self):
        return self.tasks

    def get_tasks_for(self, person, date, shift):
        return [t for t in self.tasks if t['assigned_to'] == (person, date, shift)]

# --- Calculs charges et surcharges ---
def calculate_loads(tasks):
    loads = {}
    for t in tasks:
        key = t['assigned_to']
        loads[key] = loads.get(key, 0) + t['duration']
    return loads


def detect_overloads(tasks):
        # Charger dynamiquement les max_load depuis la base
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT id, max_load FROM shifts")
            max_by_shift = {row[0]: row[1] for row in cursor.fetchall()}

        loads = calculate_loads(tasks)
        overloads = {}
        for (person, date, shift), dur in loads.items():
            max_load = max_by_shift.get(shift, 240)  # fallback 240 si shift inconnu
            if dur > max_load:
                overloads[(person, date, shift)] = dur
        return overloads


# --- Fenêtre principale ---
class PlanningApp(ctk.CTk):
    def __init__(self,username=None,role=None):
        super().__init__()
        self.on_export_button_click = None
        self.username = username
        self.role = role
        self.title(f"Planning - Connecté en tant que {self.username} role : {role}")
        #self.title("Planning Manager")
        self.geometry("1200x700")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        # Initialiser TaskManager
        self.task_manager = TaskManager()

        # Initialiser SQLite
        self.conn = sqlite3.connect("planning_manager.db")
        self.cursor = self.conn.cursor()

        # Créer tables (si nécessaire)
        self.create_tables()

        # Charger personnes depuis la BDD
        self.load_persons_from_db()  # 🔁 on récupère self.persons et self.person_colors
        self.load_rest_days_from_db() # 🔁 on récupère self.name et self.date jour de repos

        # Initialiser jours de repos
        self.rest_days = {p: [] for p in self.persons}  # ✅ basé sur self.persons chargé depuis la BDD

        # Charger tâches depuis la BDD
        self.load_tasks_from_db()

        # Initialiser mois courant
        today = datetime.date.today()
        self.current_year = today.year
        self.current_month = today.month

        self.rest_days = {p: [] for p in self.persons}

        # Initialiser SQLite
        self.conn = sqlite3.connect("planning_manager.db")  # Modifie le chemin si besoin
        self.cursor = self.conn.cursor()
        self.create_tables()
        self.load_tasks_from_db()

        # Navigation haut
        nav_frame = ctk.CTkFrame(self)
        nav_frame.pack(fill="x", padx=10, pady=5)

        self.btn_prev = ctk.CTkButton(nav_frame, text="<", width=30, command=self.prev_month)
        self.btn_prev.pack(side="left")

        self.lbl_month = ctk.CTkLabel(nav_frame, text="Juin 2025", width=150, text_color="white")
        self.lbl_month.pack(side="left", padx=10, pady=10)

        self.btn_next = ctk.CTkButton(nav_frame, text=">", width=30, command=self.next_month)
        self.btn_next.pack(side="left")

        self.btn_add_task = ctk.CTkButton(nav_frame, text="Ajouter tâche", command=self.open_add_task_window)
        self.btn_add_task.pack(side="right", padx=10)

        self.btn_manage_task_types = ctk.CTkButton(nav_frame, text="Gérer types tâches",
                                                   command=self.open_manage_task_types_window)
        self.btn_manage_task_types.pack(side="right", padx=10)

        self.btn_show_overloads = ctk.CTkButton(nav_frame, text="Voir surcharges", command=self.show_overloads)
        self.btn_show_overloads.pack(side="right", padx=10)

        self.canvas = tk.Canvas(self, bg="#222222", highlightthickness=0)
        self.canvas.pack(fill="both", expand=True, padx=10, pady=5)
        self.h_scrollbar = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview)
        self.h_scrollbar.pack(side="bottom", fill="x")

        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)
        self.inner_frame = ctk.CTkFrame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.btn_manage_rest = ctk.CTkButton(nav_frame, text="Gérer repos", command=self.open_manage_rest_window)
        self.btn_manage_rest.pack(side="right", padx=10)

        self.btn_export_excel = ctk.CTkButton(
            nav_frame,
            text="Exporter Excel",
            command=self.on_export_button_click_handler  # Appelle une fonction que tu vas ajouter juste après
        )
        self.btn_export_excel.pack(side="right", padx=10)

        self.refresh_planning_table()



    def create_tables(self):
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                person TEXT NOT NULL,
                date TEXT NOT NULL,
                shift INTEGER NOT NULL,
                duration INTEGER NOT NULL
            )
        """)
        # Table des personnes
        self.cursor.execute("""
              CREATE TABLE IF NOT EXISTS persons (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  name TEXT UNIQUE NOT NULL,
                  color TEXT NOT NULL
              )
          """)
        # Table des Repos
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS rest_days (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                person TEXT NOT NULL,
                date TEXT NOT NULL,
                UNIQUE(person, date)
            )
        """)
        # Table des Task
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS task_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL
            )
        """)
        # Création initiale (si jamais la table n'existe pas)
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS task_types (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                default_duration INTEGER DEFAULT 60
            )
        """)
        self.conn.commit()

        # Ajout colonne default_duration si elle n'existe pas encore
        try:
            self.cursor.execute("ALTER TABLE task_types ADD COLUMN default_duration INTEGER DEFAULT 60")
            self.conn.commit()
        except sqlite3.OperationalError as e:
            # Cette erreur arrive si la colonne existe déjà
            print("La colonne default_duration existe déjà ou autre erreur :", e)

        self.conn.commit()

        self.cursor.execute("SELECT COUNT(*) FROM task_types")
        if self.cursor.fetchone()[0] == 0:
            default_tasks = [
                ("Maintenance",), ("Réparation",), ("Inspection",),
                ("Nettoyage",), ("Test",), ("Préparation",),
                ("Contrôle qualité",), ("Chambre",), ("Cuisine",), ("Inventaire",)
            ]
            self.cursor.executemany("INSERT INTO task_types (name) VALUES (?)", default_tasks)

        self.cursor.execute("SELECT COUNT(*) FROM persons")
        if self.cursor.fetchone()[0] == 0:
            default_persons = [
                ("Alice", "#9c99ff"),
                ("Bob", "#99ff99"),
                ("Charlie", "#9999ff"),
                ("David", "#ffcc99")
            ]
            self.cursor.executemany("INSERT INTO persons (name, color) VALUES (?, ?)", default_persons)

        self.conn.commit()

    def load_tasks_from_db(self):
        self.task_manager.tasks.clear()  # vider avant de recharger
        self.cursor.execute("SELECT id, name, person, date, shift, duration FROM tasks")
        rows = self.cursor.fetchall()
        for id_, name, person, date_str, shift, duration in rows:
            date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            task_data = {
                'id': id_,  # on garde l'id ici
                'name': name,
                'duration': duration,
                'assigned_to': (person, date, shift)
            }
            self.task_manager.add_task(task_data)

    def load_persons_from_db(self):
        self.cursor.execute("SELECT name, color FROM persons")
        rows = self.cursor.fetchall()
        self.persons = [row[0] for row in rows]
        self.person_colors = {row[0]: row[1] for row in rows}


    def load_rest_days_from_db(self):
        self.rest_days = {}
        self.cursor.execute("SELECT person, date FROM rest_days")
        rows = self.cursor.fetchall()
        for person, date_str in rows:
            date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            self.rest_days.setdefault(person, []).append(date)
        for dates in self.rest_days.values():
            dates.sort()

    def load_task_names(self):
        self.cursor.execute("SELECT name FROM task_types ORDER BY name")
        rows = self.cursor.fetchall()
        return [r[0] for r in rows]

    def show_task_details_popup(self, person, date, shift):
        win = ctk.CTkToplevel(self)
        win.title(f"Tâches de {person} le {format_date(date)} (Shift {shift})")
        center_window(win, 450, 400)
        # 👉 Mettre la fenêtre au premier plan
        win.attributes('-topmost', True)
        win.focus_force()


        tasks = self.task_manager.get_tasks_for(person, date, shift)
        total_duration = sum(t['duration'] for t in tasks)
        #max_load = MAX_SHIFT_LOADS.get(shift, 240)
        max_load = get_shift_max_load(shift)
        if max_load is None:
            max_load = 240  # Valeur par défaut si non trouvée


        if total_duration > max_load:
            alert_text = f"Surcharge détectée : {total_duration} min / {max_load} min ⚠️"
            alert_color = "#ff4d4d"
        else:
            alert_text = f"Charge : {total_duration} min / {max_load} min"
            alert_color = "#228822"

        alert_label = ctk.CTkLabel(win, text=alert_text, fg_color=alert_color, text_color="black", corner_radius=8,
                                   width=400, height=40)
        alert_label.pack(pady=10)

        if tasks:
            # Affichage dans une liste avec bouton modifier par tâche
            for i, t in enumerate(tasks):
                frame = ctk.CTkFrame(win)
                frame.pack(fill="x", padx=10, pady=3)
                ctk.CTkLabel(frame, text=f"{t['name']} - {t['duration']} min", anchor="w").pack(side="left", fill="x",
                                                                                                expand=True)
                btn_edit = ctk.CTkButton(frame, text="Modifier", width=70,
                                         command=lambda task=t: self.open_edit_task_window(task, win))
                btn_edit.pack(side="right", padx=5)
        else:
            ctk.CTkLabel(win, text="Aucune tâche.", fg_color="#444").pack(pady=10)

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=10)

    def on_export_button_click_handler(self):
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            title="Enregistrer sous"
        )
        if filename:
            try:
                export_to_excel(
                    self.persons,
                    self.task_manager,
                    self.rest_days,
                    datetime.date(self.current_year, self.current_month, 1),  # début du mois affiché
                    filename
                )
                messagebox.showinfo("Export réussi", f"Exportation réussie :\n{filename}")
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'export :\n{e}")

    def open_edit_task_window(self, task, parent_window):
        win = ctk.CTkToplevel(self)
        win.title("Modifier tâche")
        center_window(win, 350, 370)
        # 👉 Mettre la fenêtre au premier plan
        win.attributes('-topmost', True)
        win.focus_force()
        # 👉 Bloquer les interactions avec la fenêtre principale (optionnel mais utile)
        win.grab_set()
        # Champs pré-remplis
        current_name = task['name']
        current_duration = task['duration']
        current_person, current_date, current_shift = task['assigned_to']

        # Nom tâche
        ctk.CTkLabel(win, text="Nom tâche:").pack(pady=5)
        name_var = tk.StringVar(value=current_name)
        name_entry = ctk.CTkEntry(win, textvariable=name_var)
        name_entry.pack()

        # Personne
        ctk.CTkLabel(win, text="Personne:").pack(pady=5)
        person_var = tk.StringVar(value=current_person)
        person_combo = ttk.Combobox(win, values=self.persons, textvariable=person_var, state="readonly")
        person_combo.pack()

        # Date
        ctk.CTkLabel(win, text="Date (YYYY-MM-DD):").pack(pady=5)
        date_var = tk.StringVar(value=format_date(current_date))
        date_entry = ctk.CTkEntry(win, textvariable=date_var)
        date_entry.pack()

        # Shift
        ctk.CTkLabel(win, text="Shift (0,1,2):").pack(pady=5)
        shift_var = tk.IntVar(value=current_shift)
        shift_combo = ttk.Combobox(win, values=[0, 1, 2], textvariable=shift_var, state="readonly")
        shift_combo.pack()

        # Durée
        ctk.CTkLabel(win, text="Durée (min):").pack(pady=5)
        duration_var = tk.IntVar(value=current_duration)
        duration_combo = ttk.Combobox(win, values=DURATION_OPTIONS, textvariable=duration_var, state="readonly")
        duration_combo.pack()

        def save_edit():
            new_name = name_var.get().strip()
            new_person = person_var.get()
            try:
                new_date = datetime.datetime.strptime(date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide")
                return
            new_shift = shift_var.get()
            new_duration = duration_var.get()

            # Vérifier surcharge avec les tâches existantes sans compter la tâche actuelle
            tasks = self.task_manager.get_tasks()
            total = new_duration
            for t in tasks:
                if t == task:
                    continue
                if t['assigned_to'] == (new_person, new_date, new_shift):
                    total += t['duration']
            #max_load = MAX_SHIFT_LOADS.get(new_shift, 240)
            max_load = get_shift_max_load(new_shift)
            if total > max_load:
                messagebox.showerror("Erreur", "Surcharge détectée, modification refusée")
                return

            # Mise à jour en base de données par id (pas avec LIMIT 1)
            self.cursor.execute("""
                UPDATE tasks SET name=?, person=?, date=?, shift=?, duration=?
                WHERE id=?
            """, (
                new_name, new_person, format_date(new_date), new_shift, new_duration, task['id']
            ))
            self.conn.commit()

            # Mise à jour dans la liste mémoire
            task['name'] = new_name
            task['duration'] = new_duration
            task['assigned_to'] = (new_person, new_date, new_shift)

            self.refresh_planning_table()
            parent_window.destroy()
            win.destroy()

        btn_save = ctk.CTkButton(win, text="Enregistrer", command=save_edit)
        btn_save.pack(pady=10)

    def refresh_planning_table(self):
        import datetime
        from tkinter import messagebox

        mois_nom = py_calendar.month_name[self.current_month] # ex: "Juin"
        mois_str = f"{mois_nom} {self.current_year}"
        self.lbl_month.configure(text=mois_str)

        # Supprime les anciens widgets
        for widget in self.inner_frame.winfo_children():
            widget.destroy()

        # Calcul du nombre de jours dans le mois courant
        if self.current_month == 12:
            next_month = datetime.date(self.current_year + 1, 1, 1)
        else:
            next_month = datetime.date(self.current_year, self.current_month + 1, 1)
        last_day_of_month = (next_month - datetime.timedelta(days=1)).day

        # Affichage des en-têtes : jours du mois
        for col, day in enumerate(range(1, last_day_of_month + 1)):
            lbl = ctk.CTkLabel(self.inner_frame, text=str(day), width=70, fg_color="#888", corner_radius=5)
            lbl.grid(row=0, column=col + 1, padx=1, pady=1)

        # Affichage des lignes : personnes et boutons
        for row, person in enumerate(self.persons):
            # Nom personne à gauche
            lbl = ctk.CTkLabel(self.inner_frame, text=person, width=100, fg_color="#470", corner_radius=5,
                               text_color="black")
            lbl.grid(row=row + 1, column=0, padx=1, pady=1)

            for col, day in enumerate(range(1, last_day_of_month + 1)):
                date = datetime.date(self.current_year, self.current_month, day)

                # Vérifier si jour de repos
                rest_dates = self.rest_days.get(person, [])
                if date in rest_dates:
                    # Bouton rouge "Repos"
                    btn = ctk.CTkButton(
                        self.inner_frame,
                        text="Repos",
                        width=70,
                        height=30,
                        fg_color="#a33",
                        text_color="white",
                        command=lambda p=person, d=date: messagebox.showinfo("Jour de repos",
                                                                             f"{p} est en repos le {d.strftime('%d/%m/%Y')}")
                    )
                else:
                    # Chercher les tâches pour ce jour, cette personne et shift 0
                    tasks = [t for t in self.task_manager.get_tasks() if t['assigned_to'] == (person, date, 0)]
                    btn_text = f"{len(tasks)} tâche(s)" if tasks else ""
                    default_color = self.person_colors.get(person, "#450")

                    # Détecter surcharge pour ce triplet
                    overloads = detect_overloads(self.task_manager.get_tasks())
                    is_overloaded = (person, date, 0) in overloads

                    # Couleur spéciale si surcharge
                    fg_color = "#d00" if is_overloaded else default_color
                    text_color = "white" if is_overloaded else "black"
                    btn_text = f"{btn_text} ⚠️" if is_overloaded else btn_text

                    btn = ctk.CTkButton(
                        self.inner_frame,
                        text=btn_text,
                        width=70,
                        height=30,
                        fg_color=fg_color,
                        text_color=text_color,
                        command=lambda p=person, d=date, s=0: self.show_task_details_popup(p, d, s)
                    )

                btn.grid(row=row + 1, column=col + 1, padx=2, pady=2)

    def load_task_types(self):
        self.cursor.execute("SELECT name, default_duration FROM task_types ORDER BY name")
        rows = self.cursor.fetchall()
        return {name: default_duration for name, default_duration in rows}

    def prev_month(self):
        self.current_month -= 1
        if self.current_month < 1:
            self.current_month = 12
            self.current_year -= 1
        self.refresh_planning_table()

    def next_month(self):
        self.current_month += 1
        if self.current_month > 12:
            self.current_month = 1
            self.current_year += 1
        self.refresh_planning_table()


# fenetre task
    def open_add_task_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Ajouter tâche")
        center_window(win, 350, 350)
        win.attributes('-topmost', True)
        win.focus_force()
        win.grab_set()

        ctk.CTkLabel(win, text="Nom tâche:").pack(pady=5)
        name_var = tk.StringVar()
        task_types = self.load_task_types()
        task_names = list(task_types.keys())
        name_combo = ttk.Combobox(win, values=task_names, textvariable=name_var, state="readonly")
        name_combo.pack()

        ctk.CTkLabel(win, text="Personne:").pack(pady=5)
        person_var = tk.StringVar()
        person_combo = ttk.Combobox(win, values=self.persons, textvariable=person_var, state="readonly")
        person_combo.pack()

        ctk.CTkLabel(win, text="Date:").pack(pady=5)

        date_entry = DateEntry(win, date_pattern="yyyy-mm-dd", font=("Arial", 12))
        date_entry.pack(pady=5)

        ctk.CTkLabel(win, text="Shift (0,1,2):").pack(pady=5)
        shift_var = tk.IntVar(value=0)
        shift_combo = ttk.Combobox(win, values=[0, 1, 2], textvariable=shift_var, state="readonly")
        shift_combo.pack()

        ctk.CTkLabel(win, text="Durée (min):").pack(pady=5)
        duration_var = tk.IntVar(value=15)
        duration_combo = ttk.Combobox(win, values=DURATION_OPTIONS, textvariable=duration_var, state="readonly")
        duration_combo.pack()

        # Met à jour la durée quand on change la tâche
        def on_task_selected(event):
            selected_task = name_var.get()
            default_dur = task_types.get(selected_task, 15)
            duration_var.set(default_dur)

        name_combo.bind("<<ComboboxSelected>>", on_task_selected)

        def add_task_action():
            name = name_var.get().strip()
            person = person_var.get()
            #date_var = date_entry.get_date()  # Cela retourne un objet datetime.date

            try:
                date = date = date_entry.get_date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide")
                return
            shift = shift_var.get()
            duration = duration_var.get()

            # Vérifier surcharge
            tasks = self.task_manager.get_tasks()
            total = duration
            for t in tasks:
                if t['assigned_to'] == (person, date, shift):
                    total += t['duration']
            max_load = get_shift_max_load(shift)
            if max_load is None:
                max_load = 240  # Valeur par défaut si non trouvée
            if total > max_load:
                messagebox.showerror("Erreur", "Surcharge détectée, tâche non ajoutée")
                return

            # Ajout en base de données
            self.cursor.execute("""
                INSERT INTO tasks (name, person, date, shift, duration) VALUES (?, ?, ?, ?, ?)
            """, (name, person, format_date(date), shift, duration))
            self.conn.commit()
            task_id = self.cursor.lastrowid

            task_data = {
                'id': task_id,
                'name': name,
                'duration': duration,
                'assigned_to': (person, date, shift)
            }
            self.task_manager.add_task(task_data)

            self.refresh_planning_table()
            win.destroy()

        btn_add = ctk.CTkButton(win, text="Ajouter", command=add_task_action)
        btn_add.pack(pady=10)

        # Sélectionne la première tâche par défaut pour initialiser la durée
        if task_names:
            name_var.set(task_names[0])
            duration_var.set(task_types[task_names[0]])

    # gestion des jours de repos windows
    def open_manage_rest_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Gestion des jours de repos")
        center_window(win, 400, 450)
        win.attributes('-topmost', True)
        win.focus_force()
        win.grab_set()

        # Style pour Combobox
        style = ttk.Style()
        style.theme_use("default")
        style.configure("CustomCombobox.TCombobox",
                        font=("Arial", 20),
                        padding=5)

        # Sélecteur de personne
        person_var = tk.StringVar(value=self.persons[0])
        person_combo = ttk.Combobox(win,
                                    values=self.persons,
                                    textvariable=person_var,
                                    state="readonly",
                                    style="CustomCombobox.TCombobox")
        person_combo.pack(pady=10)

        # Sélecteur de date (calendrier interactif)
        date_label = ctk.CTkLabel(win, text="Choisissez une date :", font=("Arial", 14))
        date_label.pack()
        date_entry = DateEntry(win, date_pattern="yyyy-mm-dd", font=("Arial", 14))
        date_entry.pack(pady=5)

        # Liste des jours de repos
        rest_listbox = tk.Listbox(win, font=("Arial", 13))
        rest_listbox.pack(expand=True, fill="both", padx=10, pady=10)

        def refresh_rest_list():
            rest_listbox.delete(0, tk.END)
            p = person_var.get()
            for d in sorted(self.rest_days.get(p, [])):
                rest_listbox.insert(tk.END, format_date(d))

        def add_rest_day():
            try:
                date_str = date_entry.get()
                if not date_str:
                    return
                d = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide")
                return

            p = person_var.get()
            if not p:
                messagebox.showerror("Erreur", "Veuillez sélectionner une personne")
                return

            if d not in self.rest_days.get(p, []):
                try:
                    self.cursor.execute("INSERT INTO rest_days (person, date) VALUES (?, ?)", (p, date_str))
                    self.conn.commit()
                    self.rest_days.setdefault(p, []).append(d)
                    self.rest_days[p].sort()
                    refresh_rest_list()
                    self.refresh_planning_table()
                except sqlite3.IntegrityError:
                    messagebox.showinfo("Info", "Ce jour est déjà enregistré comme jour de repos.")
            else:
                messagebox.showinfo("Info", "Ce jour est déjà un jour de repos pour cette personne.")

        def remove_selected_rest():
            p = person_var.get()
            sel = rest_listbox.curselection()
            if not sel:
                return
            idx = sel[0]
            d = self.rest_days[p][idx]
            self.cursor.execute("DELETE FROM rest_days WHERE person = ? AND date = ?", (p, format_date(d)))
            self.conn.commit()
            del self.rest_days[p][idx]
            refresh_rest_list()
            self.refresh_planning_table()

        # Boutons
        btn_frame = ctk.CTkFrame(win)
        btn_frame.pack(pady=10)

        btn_add = ctk.CTkButton(btn_frame, text="Ajouter repos", command=add_rest_day)
        btn_add.pack(side="left", padx=10)

        btn_remove = ctk.CTkButton(btn_frame, text="Supprimer repos", command=remove_selected_rest)
        btn_remove.pack(side="left", padx=10)

        person_combo.bind("<<ComboboxSelected>>", lambda e: refresh_rest_list())
        refresh_rest_list()

    def open_manage_task_types_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Gérer les types de tâches")
        center_window(win, 400, 400)
        win.grab_set()
        win.attributes('-topmost', True)
        win.focus_force()

        # Liste des tâches
        task_listbox = tk.Listbox(win)
        task_listbox.pack(expand=True, fill="both", padx=10, pady=10)

        def refresh_task_list():
            task_listbox.delete(0, tk.END)
            task_names = self.load_task_names()
            for t in task_names:
                task_listbox.insert(tk.END, t)

        def add_task_type():
            new_task = simpledialog.askstring("Ajouter type tâche", "Nom du nouveau type de tâche :", parent=win)
            if new_task:
                try:
                    self.cursor.execute("INSERT INTO task_types (name) VALUES (?)", (new_task,))
                    self.conn.commit()
                    refresh_task_list()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Erreur", "Ce type de tâche existe déjà.", parent=win)

        def remove_task_type():
            sel = task_listbox.curselection()
            if not sel:
                return
            task_to_remove = task_listbox.get(sel[0])
            if messagebox.askyesno("Confirmer", f"Supprimer le type de tâche '{task_to_remove}' ?", parent=win):
                self.cursor.execute("DELETE FROM task_types WHERE name = ?", (task_to_remove,))
                self.conn.commit()
                refresh_task_list()

        btn_frame = ctk.CTkFrame(win)
        btn_frame.pack(fill="x", pady=5)

        btn_add = ctk.CTkButton(btn_frame, text="Ajouter", command=add_task_type)
        btn_add.pack(side="left", padx=10)

        btn_remove = ctk.CTkButton(btn_frame, text="Supprimer", command=remove_task_type)
        btn_remove.pack(side="left", padx=10)

        refresh_task_list()

    def show_overloads(self):
        overloads = detect_overloads(self.task_manager.get_tasks())
        if not overloads:
            messagebox.showinfo("Surcharges", "Aucune surcharge détectée.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Surcharges")
        center_window(win, 400, 300)

        for (person, date, shift), duration in overloads.items():
            max_load = get_shift_max_load(shift)
            if max_load is None:
                max_load = 240  # Valeur par défaut si non trouvée
            lbl = ctk.CTkLabel(win, text=f"{person} le {format_date(date)} (Shift {shift}): {duration} min / {max_load} min", fg_color="#ff5555", corner_radius=5)
            lbl.pack(pady=5, padx=10, fill="x")

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=10)

    import openpyxl

    def export_to_excel(persons, task_manager, rest_days, start_date, filename):
        from openpyxl import Workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Planning"

        # En-têtes
        ws.append(["Nom", "Date", "Shift", "Tâche", "Durée (min)", "Repos ?"])

        for person in persons:
            for day_offset in range(31):  # Exporte jusqu'à 31 jours
                date = start_date + datetime.timedelta(days=day_offset)
                tasks = task_manager.get_tasks_for(person, date, 0)
                is_rest = date in rest_days.get(person, [])
                for task in tasks:
                    ws.append([
                        person,
                        date.strftime("%Y-%m-%d"),
                        0,
                        task['name'],
                        task['duration'],
                        "Oui" if is_rest else "Non"
                    ])

        wb.save(filename)


