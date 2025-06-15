import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
from cryptography.fernet import Fernet, InvalidToken
import os
import modules.login as lg

ENCRYPTION_KEY = os.getenv("ENCRYPTION_KEY")
fernet = Fernet(ENCRYPTION_KEY)

# Ensure the encryption key is loaded
if ENCRYPTION_KEY is None:
    raise ValueError("No encryption key found in environment variables.")
class LoginFrame(ctk.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.username_entry = ctk.CTkEntry(self, placeholder_text="Nombre de usuario")
        self.username_entry.pack(pady=10)
        self.password_entry = ctk.CTkEntry(self, placeholder_text="Contraseña", show="*")
        self.password_entry.pack(pady=10)
        self.remember_var = tk.BooleanVar()  # Variable to track the checkbox state
        self.remember_checkbox = ctk.CTkCheckBox(self, text="Recordar mis credenciales", variable=self.remember_var)
        self.remember_checkbox.pack(pady=5)
        self.login_button = ctk.CTkButton(self, text="Login", command=self.authenticate)
        self.login_button.pack(pady=10)
        self.user_access = None
        # Load saved credentials if available
        self.load_credentials()

    def save_credentials(self):
        if self.remember_var.get():
            encrypted_username = fernet.encrypt(self.username_entry.get().encode())
            encrypted_password = fernet.encrypt(self.password_entry.get().encode())
            with open("credentials.txt", "wb") as f:
                f.write(encrypted_username + b"," + encrypted_password)

    def load_credentials(self):
        try:
            with open("credentials.txt", "rb") as f:
                data = f.read()
                encrypted_username, encrypted_password = data.split(b",")
                self.username = fernet.decrypt(encrypted_username).decode()
                self.password = fernet.decrypt(encrypted_password).decode()
                self.username_entry.insert(0, self.username)
                self.password_entry.insert(0, self.password)
        except FileNotFoundError:
            pass
        except (ValueError, InvalidToken):
            messagebox.showerror("Error", "Unable to decrypt credentials. Please enter the correct password.")

    def authenticate(self):
        username = self.username_entry.get()
        password = self.password_entry.get()

        # Store user access configuration
        self.user_access = lg.authenticate_user(username, password)

        if lg.authenticate_user(username, password):
            messagebox.showinfo("Login Exitoso", "Bienvenido!")
            self.save_credentials()  # Save credentials before showing the app frame
            self.master.show_app_frame(self.user_access)
        else:
            messagebox.showerror("Login Fallido", "Credenciales invalidas o acceso denegado.")
