import customtkinter as ctk
from tkinter import messagebox
import pyodbc
import os
from dotenv import load_dotenv
from pathlib import Path
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfString, PdfObject
import tempfile
import subprocess
import sys
import modules.query as qry
import tkinter as tk

# Load environment variables

load_dotenv()

# Define PDF path constants
PDF_FILENAME = "Formato_de_aprobación_act.pdf"
PDF_PATH = Path(__file__).parent.parent / 'PDF' / PDF_FILENAME
if hasattr(sys, "_MEIPASS"):
    icon_path = os.path.join(sys._MEIPASS, "..","Assets", "icon.ico")
    theme_path = os.path.join(sys._MEIPASS, "..","themes", "lavender.json")
else:
    icon_path = os.path.join(os.path.dirname(__file__), "..","Assets", "icon.ico")
    theme_path = os.path.join(os.path.dirname(__file__), "..","themes", "lavender.json")

class DatabaseConnection:
    def __init__(self, db_number=1):
        prefix = f"DB{db_number}_"
        self.connection_string = (
            f"DRIVER={os.getenv(prefix + 'DRIVER')};"
            f"SERVER={os.getenv(prefix + 'SERVER')};"
            f"DATABASE={os.getenv(prefix + 'DATABASE')};"
            f"UID={os.getenv(prefix + 'UID')};"
            f"PWD={os.getenv(prefix + 'PWD')};"
        )

    def __enter__(self):
        self.conn = pyodbc.connect(self.connection_string)
        return self.conn

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.conn.close()

class PDFFillerWindow(ctk.CTkToplevel):
    
    def __init__(self, master=None):
        if os.path.exists(theme_path):
            ctk.set_default_color_theme(theme_path)
        super().__init__(master)
        self.title("Llenado de Formato de Aprobación")
        self.geometry("400x300")
        self.master = master 
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self.create_input_frame()
        self.create_button_frame()
        self.create_debug_frame()
        self.master.protocol("WM_DELETE_WINDOW", self.close_both_windows)
        self.protocol("WM_DELETE_WINDOW", self.close_window) 
        if icon_path:
            self.iconbitmap(icon_path)
        if not PDF_PATH.exists():
            messagebox.showerror(
                "Error",
                f"PDF no encontrado: {PDF_FILENAME}\n"
                "Asegúrese de que el PDF esté en el mismo directorio que este script.",
       # Hacer la ventana modal

      )
        self.transient(master)  # Asociar con la ventana principal
        self.grab_set()         # Deshabilitar la interacción con otras ventanas
        self.protocol("WM_DELETE_WINDOW", self.close_window)  # Manejo al cerrar ventana        
    def close_window(self):
        """Cerrar la ventana secundaria."""  # Liberar la interacción con otras ventanas
        self.destroy()
    def close_both_windows(self):
        """Cerrar ambas ventanas (principal y secundaria)."""
        if self.winfo_exists():
            self.destroy()  # Cierra la ventana secundaria primero
        if self.master and self.master.winfo_exists():
            self.master.destroy()
    def create_input_frame(self):
        input_frame = ctk.CTkFrame(self)
        input_frame.grid(row=0, column=0, padx=20, pady=20, sticky="ew")
        ctk.CTkLabel(input_frame, text="Número de Control:", font=("Roboto", 14)).grid(
            row=0, column=0, padx=10, pady=10
        )
        self.id_entry = ctk.CTkEntry(input_frame, placeholder_text="Ingrese número", width=200)
        self.id_entry.grid(row=0, column=1, padx=10, pady=10)
  
    def create_button_frame(self):
        button_frame = ctk.CTkFrame(self)
        button_frame.grid(row=1, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkButton(
            button_frame,
            text="Llenar PDF",
            command=self.fill_pdf,
            font=("Roboto", 14),
            #fg_color="#2B7539",
            #hover_color="#1E5C2C",
        ).pack(pady=10, padx=20, fill="x")

    def create_debug_frame(self):
        debug_frame = ctk.CTkFrame(self)
        debug_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        ctk.CTkButton(
            debug_frame,
            text="Verificar Campos del PDF",
            command=self.check_pdf_fields,
            font=("Roboto", 12),
            fg_color="#666666",
            hover_color="#444444",
        ).pack(pady=5, padx=20, fill="x")

    def check_pdf_fields(self):
        try:
            pdf = PdfReader(PDF_PATH)
            fields = []
            for page in pdf.pages:
                if page["/Annots"]:
                    for annot in page["/Annots"]:
                        if annot.get("/T"):
                            field_name = annot["/T"][1:-1] if annot["/T"].startswith("(") else annot["/T"]
                            fields.append(field_name)
            if fields:
                messagebox.showinfo("Campos Encontrados", f"Campos en el PDF:\n\n{', '.join(fields)}")
            else:
                messagebox.showwarning("Sin Campos", "No se encontraron campos en el PDF")
        except Exception as e:
            messagebox.showerror("Error", f"Error al leer campos: {str(e)}")
            
    def get_data_from_sql(self, control_number):
        try:
            with DatabaseConnection(1) as conn:
                cursor = conn.cursor()
                
                cursor.execute(qry.pdf_query, (control_number,))
                result = cursor.fetchone()
                return result  # Return the full row instead of just first column
        except pyodbc.Error as e:
            messagebox.showerror("Error de Base de Datos", f"Error al acceder a la base de datos: {str(e)}")
            return None

    def update_field_value(self, annotation, value):
        """Update a form field with the given value"""
        pdf_string = PdfString.encode(value)
        # Update the field value
        annotation.update(PdfDict(V=pdf_string, AS=pdf_string, Ff=0, AP=PdfDict(N=PdfDict(Tx=PdfDict()))))
        # Set field states using PdfName
        annotation[PdfName("F")] = 4
        annotation[PdfName("DA")] = PdfString("(/Helv 0 Tf 0 g)")
        annotation[PdfName("Q")] = 0
        # Set appearance stream
        if not annotation["/AP"]["/N"]["/Tx"].stream:
            annotation["/AP"]["/N"]["/Tx"].stream = f"BT\n/Helv 12 Tf\n1 0 0 1 2 2 Tm\n({value}) Tj\nET"

    def fill_pdf(self):
        try:
            control_number = self.id_entry.get()
            if not control_number:
                messagebox.showwarning("Error de Entrada", "Por favor ingrese un número de control")
                return

            result = self.get_data_from_sql(control_number)
           
            if result is None:
                messagebox.showwarning("No encontrado", "No se encontró información para este número de control")
                return
          
            for index, valor in enumerate(result):
             if index in {50,23,27,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,51,52,54,55,56,57,58,60}:
                 continue
             if valor is not None:
                valor = str(valor).strip()
             if valor is None or valor == "":
                
                 messagebox.showwarning("Datos incompletos", "El número de control tiene datos nulos o vacíos")
                
                 return
 
            # Field mapping: PDF field name -> SQL column index
            field_mapping = {
                "Texto 1": 4,  # Producto
                "Texto 2": 1,  # Fecha_Soli 
                "Texto 3#0": 3,  # PT
                "Texto 4": 59,  # NIT
                "Texto 5": 5,  # Nombre Cliente
                "Texto 6": 5,  # RAZON SOCIAL
                "Texto 7": 100,  # MARCA
                "Txt 34": 28,  # CAJA PLEGADIZA
                "Texto 1175": 2,   #N COTIZACION 
                "Texto 3#1": 3, #PT
                "Texto 17": 6, # NSOC
                "Texto 15": 0, # N_CONTROL
                "Texto 11": 13, #CONTENIDO_ML
                "Texto 20": 25, #COD-ETIQUETA
                "Texto 22": 9, #DESC-ENVASE
                "Texto 23": 8, #COD-ENVASE
                "Texto 26": 11, #COLOR-ENVASE
                "Texto 31": 15, #DESC-TAPA
                "Texto 32": 14, #COD-TAPA
                "Texto 36": 19, #COD-FOIL
                "Texto 94": 20, #DESCRIPCION FOIL
                "Texto 38": 22, #COD-FUNDA/BANDA
                "Texto 42": 49, #DIMEN-FUNDA/BANDA
                "Texto 41": 53, #OBS-FUNDA/BANDA
                "Texto 78": 17, #COLOR-TAPA
                "Txt 40": 26, #ETIQUETA
                "Txt 44": 12, #MAT-ENVASE
                "Txt 45": 10, #SUMIN-ENVASE
                "Txt 47": 26, #ETIQUETA
                "Txt 52": 18, #MATERIAL-TAPA 
                "Txt 53": 16, #SUMIN-TAPA
                "Txt 57": 21, #TIPO-FOIL
                "Txt 58": 22, #FUNDA
                "Txt 59": 24, #UBIC-BANDA
                "Texto 66": 47, #CAJA-TERCIARIA
                "Texto 16": 7, #NUMERO DE FRAGANCIA
                "Texto 1075": 60, #NUMERO DE VERSION APROBADA
                "Texto 12": 58, #ITEM SOFSIN
                "Texto 14": 61, #COMERCIAL
                "Txt 36": 62, #TITULAR NSOC
                "Txt 37": 63, #USO EXCLUSIVO
                "Txt 41": 64, #TIPO DE ENVASE
                "Txt 42": 65, #TIPO CUELLO ENVASE
                "Txt 43": 66, #TIPO ENVASE SELLADO
                "Txt 48": 67, #TIPO DE TAPA
                "Txt 49": 68, #TIPO DE VALVULA
                "Txt 50": 69, #SOBRETAPA
                "Txt 51": 70, #ESTILO DE TAPA
                "Txt 54": 71, #ACABADO DE CASQUETE
                "Txt 55": 72, #COLOR DE CASQUETE
                "Texto 65": 73, #UNIDADES CAJA
                "Texto 68": 74, #DIMENSIONES PARRILLA
                "Txt 39": 77, #ENVASE DECORADO
                "Txt 38": 76, #FORMA COMESTICA
                "Texto 19": 78, #cODIGO ENVADE DECORADO
                "Texto 21": 79, #DESCRIPCION ENVADE DECORADO
                "Txt 46": 80,#TAPA DECORADA
                "Texto 28": 81,#CODIGO TAPA DECORADA
                "Txt 47": 83,#ETIQUETA TAPA DECORADA
                "Texto 29": 84,#CODIGO ETIQUETA TAPA DECORADA
                "Texto 30": 82,#DESCRIPCION TAPA DECORADA
                "Txt 56": 85, #LLEVA FOIL
                "Txt 58": 86, #LLEVA TERMOENCOGIBLE
                "Txt 60": 87, #LLEVA ESCURRIDOR
                "Texto 39": 89, #DESCRIPCION ESCURRIDOR
                "Texto 40": 88, #CODIGO ESCURRIDOR
                "Txt 61": 90, #LLEVA CAPILAR
                "Texto 41": 91, #OBSERVACION CAPILAR
                "Texto 42": 92, #DIMENSIONES CAPILAR
                "Txt 62": 93, #LLEVA ACCESORIOS
                "Texto 71": 94, #DESCRIPCION ACCESORIOS
                "Txt 89": 75, #CAJA IMPRESA
                "Txt 90": 95, #CAJA CORRUGADA
                "Texto 76": 48, #TIPO DE CAJA CORRUGADA
                "Txt 91": 96, #LLEVA PARRILLA
                "Txt 92": 97, #ROTULO
                "Texto 69": 98, #TEXTO ROTULO
                "Texto 67": 99, #cLAVE
                
               
            }

            template_pdf = PdfReader(PDF_PATH)
            template_pdf.Root.AcroForm.update(PdfDict(NeedAppearances=PdfObject("true")))
            
            fields_updated = False
            for page in template_pdf.pages:
                if page["/Annots"]:
                    for annot in page["/Annots"]:
                        if annot.get("/T"):
                            field_name = annot["/T"][1:-1] if annot["/T"].startswith("(") else annot["/T"]
                            if field_name in field_mapping:
                                value = result[field_mapping[field_name]]
                                if value is not None:
                                    self.update_field_value(annot, str(value))
                                    fields_updated = True

            if not fields_updated:
                messagebox.showwarning("Sin actualización", "No se actualizó ningún campo")
                return

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                output_path = temp_pdf.name
                PdfWriter().write(output_path, template_pdf)

            try:
                if os.name == "nt":
                    os.startfile(output_path)
                elif os.name == "posix":
                    subprocess.run(["open" if sys.platform == "darwin" else "xdg-open", str(output_path)])
                else:
                    messagebox.showinfo("Éxito", "PDF prellenado listo, pero no se pudo abrir automáticamente.")
            except Exception as open_error:
                messagebox.showwarning("Advertencia", "PDF prellenado listo, pero no se pudo abrir")
        except Exception as e:
            messagebox.showerror("Error", f"Error al llenar el PDF: {str(e)}")
            print(f"Detailed error: {str(e)}")
def close_window(self):
        """Cerrar la ventana principal."""
        self.destroy()