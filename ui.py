import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import datetime

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

        self.inner_frame = ctk.CTkFrame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")

        self.inner_frame.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))

        self.btn_manage_rest = ctk.CTkButton(nav_frame, text="Gérer repos", command=self.open_manage_rest_window)
        self.btn_manage_rest.pack(side="right", padx=10)

        self.btn_export_excel = ctk.CTkButton(nav_frame, text="Exporter Excel", command=self.export_to_excel)
        self.btn_export_excel.pack(side="right", padx=10)

        self.refresh_planning_table()

    def show_task_details_popup(self, person, date, shift):
        win = ctk.CTkToplevel(self)
        win.title(f"Tâches de {person} le {format_date(date)} (Shift {shift})")
        center_window(win, 400, 400)

        tasks = self.task_manager.get_tasks_for(person, date, shift)
        total_duration = sum(t['duration'] for t in tasks)
        max_load = MAX_SHIFT_LOADS.get(shift, 240)

        if total_duration > max_load:
            alert_text = f"Surcharge détectée : {total_duration} min / {max_load} min ⚠️"
            alert_color = "#ff4d4d"
        else:
            alert_text = f"Charge : {total_duration} min / {max_load} min"
            alert_color = "#228822"


        alert_label = ctk.CTkLabel(win, text=alert_text, fg_color=alert_color, text_color="black", corner_radius=8, width=300, height=40)

        alert_label.pack(pady=10)

        if tasks:
            for t in tasks:
                txt = f"{t['name']} - {t['duration']} min"
                ctk.CTkLabel(win, text=txt, justify="left", anchor="w").pack(fill="x", padx=10, pady=2)
        else:
            ctk.CTkLabel(win, text="Aucune tâche.", fg_color="#444").pack(pady=10)

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=10)

    def open_manage_rest_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Gestion des jours de repos")
        center_window(win, 400, 400)

        person_var = tk.StringVar()
        date_var = tk.StringVar()

        ctk.CTkLabel(win, text="Personne :").pack(pady=5)
        combo_person = ctk.CTkComboBox(win, values=self.persons, variable=person_var)
        combo_person.pack()

        ctk.CTkLabel(win, text="Date (AAAA-MM-JJ) :").pack(pady=5)
        entry_date = ctk.CTkEntry(win, textvariable=date_var)
        entry_date.pack()

        def add_rest_day():
            person = person_var.get()
            try:
                date = datetime.datetime.strptime(date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide (AAAA-MM-JJ)")
                return
            self.rest_days.setdefault(person, []).append(date)
            messagebox.showinfo("Info", f"{person} est en repos le {format_date(date)}.")
            self.refresh_planning_table()

        def remove_rest_day():
            person = person_var.get()
            try:
                date = datetime.datetime.strptime(date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide (AAAA-MM-JJ)")
                return
            if date in self.rest_days.get(person, []):
                self.rest_days[person].remove(date)
                messagebox.showinfo("Info", f"Repos supprimé pour {person} le {format_date(date)}.")
                self.refresh_planning_table()
            else:
                messagebox.showwarning("Info", f"{person} n'est pas en repos ce jour-là.")

        btn_add = ctk.CTkButton(win, text="Ajouter repos", command=add_rest_day)
        btn_add.pack(pady=5)

        btn_remove = ctk.CTkButton(win, text="Supprimer repos", command=remove_rest_day)
        btn_remove.pack(pady=5)

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=15)

    def refresh_planning_table(self):
        month_name = datetime.date(self.current_year, self.current_month, 1).strftime("%B %Y")
        self.lbl_month.configure(text=month_name.capitalize())

        for w in self.inner_frame.winfo_children():
            w.destroy()

        if self.current_month == 12:
            days_in_month = 31
        else:
            days_in_month = (datetime.date(self.current_year, self.current_month + 1, 1) - datetime.timedelta(days=1)).day

        ctk.CTkLabel(self.inner_frame, text="", width=60, height=30).grid(row=0, column=0, sticky="nsew", padx=1, pady=1)
        for d in range(1, days_in_month + 1):
            date = datetime.date(self.current_year, self.current_month, d)
            weekday = date.strftime("%a")
            ctk.CTkLabel(self.inner_frame, text=f"{d}\n{weekday}", width=60, height=30, fg_color="#444444", corner_radius=5).grid(row=0, column=d, sticky="nsew", padx=1, pady=1)

        row = 1
        for person in self.persons:
            for shift_idx, (start, end) in enumerate(self.shifts):
                person_shift_text = f"{person}\nShift {shift_idx}\n{start}-{end}"
                ctk.CTkLabel(self.inner_frame, text=person_shift_text, width=60, height=70, fg_color=self.person_colors.get(person, "#666666")).grid(row=row, column=0, sticky="nsew", padx=1, pady=1)

                for d in range(1, days_in_month + 1):
                    date = datetime.date(self.current_year, self.current_month, d)
                    bg_color = "#222222"
                    if date in self.rest_days.get(person, []):
                        bg_color = "#555555"

                    tasks = self.task_manager.get_tasks_for(person, date, shift_idx)
                    if tasks:
                        text = "\n".join([f"{t['name']} ({t['duration']}min)" for t in tasks])
                        bg_color = self.person_colors.get(person, "#666666")
                    else:
                        text = ""

                    lbl_cell = ctk.CTkLabel(self.inner_frame, text=text, width=60, height=70, fg_color=bg_color, corner_radius=5, justify="left", wraplength=55, text_color="white" if bg_color != "#555555" else "lightgray")
                    if tasks:
                        lbl_cell.bind("<Button-1>",
                                      lambda e, p=person, dt=date, s=shift_idx: self.show_task_details_popup(p, dt, s))
                    lbl_cell.grid(row=row, column=d, sticky="nsew", padx=1, pady=1)

                row += 1

    def prev_month(self):
        if self.current_month == 1:
            self.current_month = 12
            self.current_year -= 1
        else:
            self.current_month -= 1
        self.refresh_planning_table()

    def next_month(self):
        if self.current_month == 12:
            self.current_month = 1
            self.current_year += 1
        else:
            self.current_month += 1
        self.refresh_planning_table()

    def open_add_task_window(self):
        win = ctk.CTkToplevel(self)
        win.title("Ajouter tâche")
        center_window(win, 400, 400)

        name_var = tk.StringVar()
        duration_var = tk.IntVar(value=30)
        person_var = tk.StringVar()
        date_var = tk.StringVar()
        shift_var = tk.IntVar(value=0)

        ctk.CTkLabel(win, text="Nom tâche :").pack(pady=5)
        combo_name = ctk.CTkComboBox(win, values=TASK_NAMES, variable=name_var)
        combo_name.pack()

        ctk.CTkLabel(win, text="Durée (min) :").pack(pady=5)
        combo_dur = ctk.CTkComboBox(win, values=[str(d) for d in DURATION_OPTIONS], variable=duration_var)
        combo_dur.pack()

        ctk.CTkLabel(win, text="Personne :").pack(pady=5)
        combo_person = ctk.CTkComboBox(win, values=self.persons, variable=person_var)
        combo_person.pack()

        ctk.CTkLabel(win, text="Date (AAAA-MM-JJ) :").pack(pady=5)
        entry_date = ctk.CTkEntry(win, textvariable=date_var)
        entry_date.pack()

        ctk.CTkLabel(win, text="Shift :").pack(pady=5)
        combo_shift = ctk.CTkComboBox(win, values=[str(i) for i in range(len(self.shifts))], variable=shift_var)
        combo_shift.pack()

        def add_task_action():
            name = name_var.get()
            try:
                duration = int(duration_var.get())
            except Exception:
                messagebox.showerror("Erreur", "Durée invalide")
                return

            person = person_var.get()
            try:
                date = datetime.datetime.strptime(date_var.get(), "%Y-%m-%d").date()
            except ValueError:
                messagebox.showerror("Erreur", "Date invalide (AAAA-MM-JJ)")
                return

            shift = shift_var.get()
            if not name or not person:
                messagebox.showerror("Erreur", "Nom tâche et personne obligatoires")
                return

            task = {
                'name': name,
                'duration': duration,
                'assigned_to': (person, date, int(shift))
            }
            self.task_manager.add_task(task)
            self.refresh_planning_table()
            win.destroy()

        btn_add = ctk.CTkButton(win, text="Ajouter", command=add_task_action)
        btn_add.pack(pady=15)

    def show_overloads(self):
        overloads = detect_overloads(self.task_manager.get_tasks())
        if not overloads:
            messagebox.showinfo("Surcharges", "Aucune surcharge détectée.")
            return

        win = ctk.CTkToplevel(self)
        win.title("Surcharges détectées")
        center_window(win, 500, 400)

        for (person, date, shift), dur in overloads.items():
            max_load = MAX_SHIFT_LOADS.get(shift, 240)
            txt = f"{person} le {format_date(date)} Shift {shift}: {dur} min / {max_load} min"
            ctk.CTkLabel(win, text=txt, fg_color="#ff4d4d", corner_radius=8, width=450, height=30).pack(pady=3)

        ctk.CTkButton(win, text="Fermer", command=win.destroy).pack(pady=10)

    def export_to_excel(self):
        from openpyxl import Workbook
        from openpyxl.utils import get_column_letter

        filepath = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Fichiers Excel", "*.xlsx")],
            title="Enregistrer planning sous..."
        )
        if not filepath:
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "Planning Tâches"

        headers = ["Personne", "Date", "Shift", "Nom tâche", "Durée (min)"]
        ws.append(headers)

        tasks = self.task_manager.get_tasks()
        for t in tasks:
            person, date, shift = t['assigned_to']
            row = [person, format_date(date), shift, t['name'], t['duration']]
            ws.append(row)

        # Ajuster largeur colonnes
        for col_idx, col_title in enumerate(headers, start=1):
            max_length = len(col_title)
            for cell in ws[get_column_letter(col_idx)]:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = max_length + 2

        try:
            wb.save(filepath)
            messagebox.showinfo("Succès", f"Export réussi vers:\n{filepath}")
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible d'enregistrer le fichier:\n{e}")

