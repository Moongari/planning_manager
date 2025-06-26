import tkinter as tk
import customtkinter as ctk
from tkinter import messagebox
from ui import PlanningApp  # Ton application principale

# Exemple simplifié (tu peux remplacer par vérification en base)
USERS = {
    "admin": "1234",
    "user": "pass"
}

class LoginWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Connexion")
        self.geometry("300x200")
        self.eval('tk::PlaceWindow . center')
        self.resizable(False, False)

        ctk.CTkLabel(self, text="Nom d'utilisateur").pack(pady=5)
        self.username_var = tk.StringVar()
        ctk.CTkEntry(self, textvariable=self.username_var).pack()

        ctk.CTkLabel(self, text="Mot de passe").pack(pady=5)
        self.password_var = tk.StringVar()
        ctk.CTkEntry(self, textvariable=self.password_var, show="*").pack()

        ctk.CTkButton(self, text="Se connecter", command=self.check_login).pack(pady=15)

    def check_login(self):
        username = self.username_var.get()
        password = self.password_var.get()

        if USERS.get(username) == password:
            self.destroy()
            self.open_main_app()
        else:
            messagebox.showerror("Erreur", "Identifiants incorrects")

    def open_main_app(self):
        app = PlanningApp()
        app.mainloop()

if __name__ == "__main__":
    login = LoginWindow()
    login.mainloop()
