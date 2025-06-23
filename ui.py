import customtkinter as ctk
import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import datetime
import openpyxl

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

        self.refresh_planning_table()

    def show_task_details_popup(self, person, date, shift):
        win = ctk.CTkToplevel(self)
        win.title(f"Tâches de {person} le {format_date(date)} (Shift {shift})")
        center_window(win, 400, 400)

        tasks = self.task_manager.get_tasks_for(person, date, shift)
        total_duration = sum(t['duration'] for t in tasks)
        max_load = MAX_SHIFT_LOADS.get(shift, 240)

        # ✅ Message d'alerte si surcharge
        if total_duration > max_load:
            alert_text = f"Surcharge détectée : {total_duration} min / {max_load} min ⚠️"
            alert_color = "#ff4d4d"
        else:
            alert_text = f"Charge : {total_duration} min / {max_load} min"
            alert_color = "#228822"

        alert_label = ctk.CTkLabel(win, text=alert_text, fg_color=alert_color, corner_radius=8, width=300, height=40)
        alert_label.pack(pady=10)

        # ✅ Liste des tâches
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
                    # 👉 Ajouter un clic si des tâches existent
                    if tasks:
                        lbl_cell.bind("<Button-1>",
                                      lambda e, p=person, d=date, s=shift_idx: self.show_task_details_popup(p, d, s))

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
        AddTaskWindow(self)

    def show_overloads(self):
        overloads = detect_overloads(self.task_manager.get_tasks())
        win = ctk.CTkToplevel(self)
        win.title("Surcharges détectées")
        center_window(win, 400, 300)

        if overloads:
            shift_names = ["Matin", "Après-midi", "Soir"]
            for (person, date, shift), total in overloads.items():
                txt = f"{person} le {format_date(date)} ({shift_names[shift]}): {total} min"
                ctk.CTkLabel(win, text=txt, justify="left").pack(anchor="w", padx=10, pady=5)
        else:
            ctk.CTkLabel(win, text="Aucune surcharge détectée.", justify="center").pack(expand=True, pady=20)

class AddTaskWindow(ctk.CTkToplevel):
    def __init__(self, master, task=None):
        super().__init__(master)
        self.master = master
        self.task = task
        self.title("Ajouter tâche")
        center_window(self, 360, 400)

        ctk.CTkLabel(self, text="Nom tâche:").pack()
        self.combo_name = ctk.CTkComboBox(self, values=TASK_NAMES)
        self.combo_name.pack()

        ctk.CTkLabel(self, text="Personne:").pack()
        self.combo_person = ctk.CTkComboBox(self, values=self.master.persons)
        self.combo_person.pack()

        ctk.CTkLabel(self, text="Date (AAAA-MM-JJ):").pack()
        self.entry_date = ctk.CTkEntry(self)
        self.entry_date.pack()

        ctk.CTkLabel(self, text="Shift (0, 1, 2):").pack()
        self.combo_shift = ctk.CTkComboBox(self, values=["0", "1", "2"])
        self.combo_shift.pack()

        ctk.CTkLabel(self, text="Durée (min):").pack()
        self.combo_duration = ctk.CTkComboBox(self, values=[str(d) for d in DURATION_OPTIONS])
        self.combo_duration.pack()

        self.btn_add = ctk.CTkButton(self, text="Ajouter", command=self.add_task)
        self.btn_add.pack(pady=10)

    def add_task(self):
        name = self.combo_name.get()
        person = self.combo_person.get()
        date_str = self.entry_date.get().strip()
        shift = int(self.combo_shift.get())
        duration = int(self.combo_duration.get())

        try:
            date = datetime.datetime.strptime(date_str, "%Y-%m-%d").date()
        except ValueError:
            messagebox.showerror("Erreur", "Date invalide (AAAA-MM-JJ)")
            return

        # 🚨 Vérification si c'est un jour de repos
        if date in self.master.rest_days.get(person, []):
            messagebox.showerror("Erreur",
                                 f"{person} est en repos le {format_date(date)}. Impossible d'ajouter une tâche.")
            return

        task_data = {"name": name, "assigned_to": (person, date, shift), "duration": duration}
        self.master.task_manager.add_task(task_data)
        self.master.refresh_planning_table()
        self.destroy()


if __name__ == "__main__":
    app = PlanningApp()
    app.mainloop()
