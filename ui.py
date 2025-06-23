import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk, filedialog, simpledialog
import datetime
import sqlite3

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
    "Contrôle qualité"
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

def detect_overloads(tasks, max_by_shift=MAX_SHIFT_LOADS):
    loads = calculate_loads(tasks)
    overloads = {}
    for (person, date, shift), dur in loads.items():
        max_load = max_by_shift.get(shift, 240)
        if dur > max_load:
            overloads[(person, date, shift)] = dur
    return overloads

# --- Fenêtre principale ---
class PlanningApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Planning Manager")
        self.geometry("1200x700")
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.persons = PERSONS
        self.person_colors = PERSON_COLORS
        self.shifts = SHIFTS
        self.task_manager = TaskManager()

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

        self.lbl_month = ctk.CTkLabel(nav_frame, text="", width=150)
        self.lbl_month.pack(side="left", padx=10)

        self.btn_next = ctk.CTkButton(nav_frame, text=">", width=30, command=self.next_month)
        self.btn_next.pack(side="left")

        self.btn_add_task = ctk.CTkButton(nav_frame, text="Ajouter tâche", command=self.open_add_task_window)
        self.btn_add_task.pack(side="right", padx=10)

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

        self.btn_export_excel = ctk.CTkButton(nav_frame, text="Exporter Excel", command=self.export_to_excel)
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

    def show_task_details_popup(self, person, date, shift):
        win = ctk.CTkToplevel(self)
        win.title(f"Tâches de {person} le {format_date(date)} (Shift {shift})")
        center_window(win, 450, 400)

        tasks = self.task_manager.get_tasks_for(person, date, shift)
        total_duration = sum(t['duration'] for t in tasks)
        max_load = MAX_SHIFT_LOADS.get(shift, 240)

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

    def open_edit_task_window(self, task, parent_window):
        win = ctk.CTkToplevel(self)
        win.title("Modifier tâche")
        center_window(win, 350, 370)

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
            max_load = MAX_SHIFT_LOADS.get(new_shift, 240)
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
                    color = self.person_colors.get(person, "#450")
                    btn_text = f"{len(tasks)} tâche(s)" if tasks else ""
                    btn = ctk.CTkButton(
                        self.inner_frame,
                        text=btn_text,
                        width=70,
                        height=30,
                        fg_color=color,
                        command=lambda p=person, d=date, s=0: self.show_task_details_popup(p, d, s)
                    )

                btn.grid(row=row + 1, column=col + 1, padx=2, pady=2)

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

    def open_add_task_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Ajouter tâche")
        center_window(win, 350, 300)

        ctk.CTkLabel(win, text="Nom tâche:").pack(pady=5)
        name_var = tk.StringVar()
        name_combo = ttk.Combobox(win, values=TASK_NAMES, textvariable=name_var)
        name_combo.pack()

        ctk.CTkLabel(win, text="Personne:").pack(pady=5)
        person_var = tk.StringVar(value=self.persons[0])
        person_combo = ttk.Combobox(win, values=self.persons, textvariable=person_var, state="readonly")
        person_combo.pack()

        ctk.CTkLabel(win, text="Date (YYYY-MM-DD):").pack(pady=5)
        date_var = tk.StringVar(value=format_date(datetime.date.today()))
        date_entry = ctk.CTkEntry(win, textvariable=date_var)
        date_entry.pack()

        ctk.CTkLabel(win, text="Shift (0,1,2):").pack(pady=5)
        shift_var = tk.IntVar(value=0)
        shift_combo = ttk.Combobox(win, values=[0, 1, 2], textvariable=shift_var, state="readonly")
        shift_combo.pack()

        ctk.CTkLabel(win, text="Durée (min):").pack(pady=5)
        duration_var = tk.IntVar(value=15)
        duration_combo = ttk.Combobox(win, values=DURATION_OPTIONS, textvariable=duration_var, state="readonly")
        duration_combo.pack()

        def add_task_action():
            name = name_var.get().strip()
            person = person_var.get()
            try:
                date = datetime.datetime.strptime(date_var.get(), "%Y-%m-%d").date()
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
            max_load = MAX_SHIFT_LOADS.get(shift, 240)
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

    def open_manage_rest_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Gestion des jours de repos")
        center_window(win, 400, 400)

        person_var = tk.StringVar(value=self.persons[0])
        person_combo = ttk.Combobox(win, values=self.persons, textvariable=person_var, state="readonly")
        person_combo.pack(pady=10)

        rest_listbox = tk.Listbox(win)
        rest_listbox.pack(expand=True, fill="both", padx=10, pady=10)

        def refresh_rest_list():
            rest_listbox.delete(0, tk.END)
            p = person_var.get()
            for d in self.rest_days.get(p, []):
                rest_listbox.insert(tk.END, format_date(d))

        def add_rest_day():
            try:
                date_str = simpledialog.askstring("Ajouter jour repos", "Date (YYYY-MM-DD):")
                if date_str is None:
                    return
                d = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide")
                return

            p = person_var.get()  # Accès direct à person_var (pas self.person_var)
            if not p:
                messagebox.showerror("Erreur", "Veuillez sélectionner une personne")
                return

            if p not in self.rest_days:
                self.rest_days[p] = []

            if d not in self.rest_days[p]:
                self.rest_days[p].append(d)
                self.rest_days[p].sort()
                refresh_rest_list()
                self.refresh_planning_table()
            else:
                messagebox.showinfo("Info", "Ce jour est déjà un jour de repos pour cette personne.")

        def remove_selected_rest():
            sel = rest_listbox.curselection()
            if not sel:
                return
            p = person_var.get()
            idx = sel[0]
            del self.rest_days[p][idx]
            refresh_rest_list()
            self.refresh_planning_table()

        btn_add = ctk.CTkButton(win, text="Ajouter repos", command=add_rest_day)
        btn_add.pack(side="left", padx=10, pady=5)

        btn_remove = ctk.CTkButton(win, text="Supprimer repos", command=remove_selected_rest)
        btn_remove.pack(side="left", padx=10, pady=5)

        person_combo.bind("<<ComboboxSelected>>", lambda e: refresh_rest_list())
        refresh_rest_list()

    def show_overloads(self):
        overloads = detect_overloads(self.task_manager.get_tasks())
        if not overloads:
            messagebox.showinfo("Surcharges", "Aucune surcharge détectée.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Surcharges")
        center_window(win, 400, 300)

        for (person, date, shift), duration in overloads.items():
            max_load = MAX_SHIFT_LOADS.get(shift, 240)
            lbl = ctk.CTkLabel(win, text=f"{person} le {format_date(date)} (Shift {shift}): {duration} min / {max_load} min", fg_color="#ff5555", corner_radius=5)
            lbl.pack(pady=5, padx=10, fill="x")

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=10)

    def export_to_excel(self):
        # Placeholder pour export, à compléter si tu veux
        messagebox.showinfo("Export Excel", "Fonction d'export Excel non implémentée.")

