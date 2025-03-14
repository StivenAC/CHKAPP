import customtkinter as ctk
from tksheet import Sheet
from tkinter import ttk, messagebox
import modules.pdf as t
from openpyxl.styles import PatternFill
class MyFrame(ctk.CTkFrame):
    def __init__(self, tabview, master, load_data_func, create_record_func=None, edit_record_func=None, export_record_func=None, **kwargs):
        super().__init__(master, **kwargs)
        self.tabview = tabview
        # Store data management functions
        self.load_data_func = load_data_func
        self.create_record_func = create_record_func
        self.edit_record_func = edit_record_func
        self.export_record_func = export_record_func  

        # Original and filtered data storage
        self.original_data = []
        self.filtered_data = []

        # Create sheet
        self.sheet = Sheet(self)
        self.sheet.pack(fill="both", expand=True)

        # Enable row selection
        self.sheet.enable_bindings(
            (
                "single_select",
                "row_select",
                "arrowkeys",
                "column_width_resize",
                "row_width_resize",
                "double_click_row_resize",
                "column_select",
            )
        )
        # Create buttons
        self.button_frame = ctk.CTkFrame(self)
        self.button_frame.pack(fill="x", padx=10, pady=5)
        # Create filter entry
        self.filter_entry = ctk.CTkEntry(
            self.button_frame, placeholder_text="Buscar por Numero De control o Fecha")
        self.filter_entry.pack(padx=10, pady=10, fill="x")
        self.filter_entry.bind("<Return>", self.filter_data)

        self.create_record_button = ctk.CTkButton(
            self.button_frame, text="Crear Registro", command=self.handle_create_record
        )
        self.create_record_button.pack(side="left", padx=5)

        self.edit_record_button = ctk.CTkButton(
            self.button_frame, text="Editar Registro", command=self.handle_edit_record
        )
        self.edit_record_button.pack(side="left", padx=5)

        self.reload_button = ctk.CTkButton(self.button_frame, text="Refrescar", command=self.reload_data)
        self.reload_button.pack(side="left", padx=5)
        
        self.export_button = ctk.CTkButton( self.button_frame, text="Exportar PDF", command=self.export_pdf)
        self.export_button.pack(side="left", padx=5)

        # Initial data load
        self.load_data()

    def handle_create_record(self):
        """Handle record creation with custom function"""
        if self.create_record_func:
            self.create_record_func(self)
        else:
            self.show_default_create_dialog()

    def handle_edit_record(self):
        """Handle record editing with custom function"""
        if self.edit_record_func:
            self.edit_record_func(self)
        else:
            self.show_default_edit_dialog()

    def show_default_create_dialog(self):
        """Default create record dialog"""
        create_window = ctk.CTkToplevel(self)
        create_window.title("Crear Nuevo Registro")
        create_window.geometry("400x300")

    def show_default_export_dialog(self):
        """Default export record dialog"""
        create_window = ctk.CTkToplevel(self)
        create_window.title("exportar Registro")
        create_window.geometry("400x300")

    def show_default_edit_dialog(self, record_data):
        """Default edit record dialog"""
        edit_window = ctk.CTkToplevel(self)
        edit_window.title("Editar Registro")
        edit_window.geometry("400x300")

    def reload_data(self):
        """Reload the data"""
        # Clear existing data
        self.sheet.set_sheet_data([])

        # Load updated data from the database
        self.load_data()

    def load_data(self):
        """Initial data loading"""
        self.load_data_func(self)

    def filter_data(self, event):
        """Filter the data based on the input"""
        search_text = self.filter_entry.get().lower()
        if search_text:
            self.filtered_data = [
                row
                for row in self.original_data
                if search_text in str(row[0]).lower() or  # Search by "# OP"
                search_text in str(row[1]).lower()  # Search by "CODIGO ITEM"
            ]
            self.sheet.set_sheet_data(self.filtered_data)
        else:
            self.filtered_data = self.original_data
            self.sheet.set_sheet_data(self.filtered_data)

    def export_pdf(self):
        """Abrir la ventana para llenar el PDF"""
        active_tab = self.tabview.get()
        if active_tab == "Calidad":
            pdf_window = t.PDFFillerWindow(self)  # Pasar `self` como ventana principal
        else:
            return

    
