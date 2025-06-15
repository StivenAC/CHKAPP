import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import pyodbc
from tksheet import Sheet
import sys
import os
import json
from dotenv import load_dotenv
from ldap3 import Server, Connection, ALL, NTLM, SUBTREE
from cryptography.fernet import Fernet, InvalidToken
import logging
import modules.gui_frame as gf
import modules.extra_functions as ef
import modules.f_combobox as f_combobox
import modules.login as lg
import modules.query as qry
from pypdf import PdfReader, PdfWriter
import modules.pdf as pdf
import modules.headers as hd
# Get the correct path to the icon
if hasattr(sys, "_MEIPASS"):
    icon_path = os.path.join(sys._MEIPASS, "..","Assets", "icon.ico")
    theme_path = os.path.join(sys._MEIPASS, "..","themes", "lavender.json")
else:
    icon_path = os.path.join(os.path.dirname(__file__), "..","Assets", "icon.ico")
    theme_path = os.path.join(os.path.dirname(__file__), "..","themes", "lavender.json")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme(theme_path)

logging.basicConfig(filename="auth.log", level=logging.INFO)

# Check if .env file exists in the current directory
env_file_path = ".env" if os.path.isfile(".env") else "_internal/.env"

# Load .env file
load_dotenv(env_file_path)

# AD settings
AD_SERVER = os.getenv("AD_SERVER")
AD_DOMAIN = os.getenv("AD_DOMAIN")
AD_USER = os.getenv("AD_USER")
AD_PASSWORD = os.getenv("AD_PASSWORD")
ALLOWED_GROUPS = os.getenv("ALLOWED_GROUPS")
ALLOWED_USERS = os.getenv("ALLOWED_USERS")
ACCESS_CONFIG = os.getenv("ACCESS_CONFIG", {})
# Retrieve the encryption key from the .env file

class MyTabView(ctk.CTkTabview):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        
        # Tab configurations
        self.tab_configs = {
            "Desarrollo": {
                "load_func": self.load_data_Desarrollo,
                "create_func": self.create_record_Desarrollo,
                "edit_func": self.edit_record_Desarrollo,
            },
            "Costos": {
                "load_func": self.load_data_Costos,
                "create_func": self.create_record_Costos,
                "edit_func": self.edit_record_Costos,
            },
            "Calidad": {
                "load_func": self.load_data_Calidad,
                "create_func": self.create_record_Calidad,
                "edit_func": self.edit_record_Calidad,
            },
            "Produccion": {
                "load_func": self.load_data_Produccion,
                "create_func": self.create_record_Produccion,
                "edit_func": self.edit_record_Produccion,
            },
        }
        # Create tabs and frames
        self.frames = {}
        for tab_name, config in self.tab_configs.items():
            self.add(tab_name)
            self.frames[tab_name] = gf.MyFrame(
                master=self.tab(tab_name),
                tabview=self,
                load_data_func=config["load_func"],
                create_record_func=config.get("create_func"),
                edit_record_func=config.get("edit_func"),
                export_record_func=config.get("export_pdf"),
            )
            self.frames[tab_name].pack(fill="both", expand=True)
    
    def load_data_Desarrollo(self, frame):
        cursor = None
        conn = None
        try:
            conn_str = (
                f"DRIVER={os.getenv('DB1_DRIVER')};"
                f"SERVER={os.getenv('DB1_SERVER')};"
                f"DATABASE={os.getenv('DB1_DATABASE')};"
                f"UID={os.getenv('DB1_UID')};"
                f"PWD={os.getenv('DB1_PWD')}"
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            
            cursor.execute(qry.load_query_des)
            data = cursor.fetchall()
            frame.sheet.headers(hd.headers_load_des)

            # Convert data to list of lists with string values
            formatted_data = []
            for row in data:
                 # Convert null values to empty string
                formatted_row = [str(value) if value is not None else "" for value in row]
                if any(value == "" for value in formatted_row):
                    formatted_row.append("X")  
                else:
                    formatted_row.append("✓")  
                formatted_data.append(formatted_row)
            
            frame.original_data = formatted_data
            frame.sheet.set_sheet_data(formatted_data)

        except pyodbc.Error as e:
            print(f"An error occurred while loading data: {str(e)}", file=sys.stderr)
            messagebox.showerror(title="Error", message=f"No se pudo cargar los datos: {str(e)}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()
    # Placeholder methods for create and edit
    def create_record_Desarrollo(self, frame):
        crear_dialog = ctk.CTkToplevel(self)
        crear_dialog.title("Crear Nuevo Registro - Desarrollo")
        crear_dialog.after(201, lambda :crear_dialog.iconbitmap(icon_path))
        crear_dialog.geometry("400x500")

        # Diccionario para guardar entradas
        entries = {}
        scrollable_frame = ef.ScrollableFrame(crear_dialog)
        scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)
        # Crear entradas para cada campo
        for campo in hd.campos_load_des:
            
            frame_campo = ctk.CTkFrame(scrollable_frame)
            frame_campo.pack(pady=5, padx=20, fill="x")

            label = ctk.CTkLabel(frame_campo, text=campo, width=150, anchor="w")
            label.pack(side="left", padx=(0, 10))

            entry = ctk.CTkEntry(frame_campo, width=300)
            entry.pack(side="left", expand=True, fill="x")
            entries[campo] = entry

        def guardar():
            """Guardar el nuevo registro"""
            try:
                # Preparar conexión a la base de datos
                conn_str = (
                    f"DRIVER={os.getenv('DB1_DRIVER')};"
                    f"SERVER={os.getenv('DB1_SERVER')};"
                    f"DATABASE={os.getenv('DB1_DATABASE')};"
                    f"UID={os.getenv('DB1_UID')};"
                    f"PWD={os.getenv('DB1_PWD')}"
                )
                conn = pyodbc.connect(conn_str)
                cursor = conn.cursor()

                # Preparar consulta de inserción
                

                # Recoger valores
                valores = [entries[campo].get() for campo in hd.campos_load_des]
                valores = [texto.upper() for texto in valores]
                
                #Validations of empty fields
                if valores[0]=="" and valores[1]=="":
                    messagebox.showerror(title="Campo obligatorio", message="los campos Fecha de solicitud y Numero de Control son obligatorios")
                    return
                if valores[1]=="":
                    messagebox.showerror(title="Campo obligatorio", message="El campo Fecha de solicitud es obligatorio")
                    return
                if valores[0]=="":
                    messagebox.showerror(title="Campo obligatorio", message="El campo Numero de Control es obligatorio")
                    return
                
                # Ejecutar inserción
                
                cursor.execute(qry.create_query_des, valores)
                conn.commit()

                messagebox.showinfo(title="Éxito", message="Registro creado correctamente")
                crear_dialog.destroy()
                frame.load_data()

            except Exception as e:
                messagebox.showerror(title="Error", message=f"No se pudo crear el registro: {str(e)}", icon="cancel")
            finally:
                if "conn" in locals():
                    conn.close()

        # Botón de guardar
        guardar_btn = ctk.CTkButton(crear_dialog, text="Guardar", command=guardar)
        guardar_btn.pack(pady=20)

    def edit_record_Desarrollo(self, frame):
        try:
            selected_rows = frame.sheet.get_selected_rows()

            if not selected_rows:
                messagebox.showerror(title="Error", message="Seleccione un registro para editar")
                return

            # Get the data of the selected row from filtered or original data
            selected_data = next(iter(selected_rows))
            row_data = frame.sheet.get_row_data(selected_data)
            N_control = row_data[0]  # Assuming N_control is in the first column (index 0)

            if not N_control:
                messagebox.showerror("Error", "No se pudo obtener el N_control del registro seleccionado.")
                logging.error("El N_control es nulo o no válido.")
                return

            logging.info(f"N_control seleccionado: {N_control}")

            # Create edit dialog window
            editar_dialog = ctk.CTkToplevel(frame)
            editar_dialog.title("Editar Registro")
            editar_dialog.after(201, lambda :editar_dialog.iconbitmap(icon_path))
            editar_dialog.geometry("400x500")

            # Fields (excluding N_control)
            campos = ["Numero Cotizacion", "PT", "Producto", "Cliente", "Notificacion Sanitaria","Estado Muestra","Diseno Desarrollo","Numero de Fragancia","NIT","Comercial","Marca"]

            # Dictionary to store entry widgets
            entries = {}
            # Create the scrollable frame
            scrollable_frame = ef.ScrollableFrame(editar_dialog)
            scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)

            # Create a frame for the input fields
            input_frame = ctk.CTkFrame(scrollable_frame)
            input_frame.pack(fill="x", expand=True)
            # Create input fields for each column
            for i, campo in enumerate(campos, start=2):
                field_frame = ctk.CTkFrame(input_frame)
                field_frame.pack(pady=5, padx=20, fill="x")

                label = ctk.CTkLabel(field_frame, text=campo, width=150, anchor="w")
                label.pack(side="left", padx=(0, 10))
       
                if campo == "":  
                    opciones = ["Sí", "No"]  
                    entry = ctk.CTkComboBox(field_frame, values=opciones, width=300)
                    entry.pack(side="left", expand=True, fill="x")
                       
                    valor_actual = row_data[i]
                    if valor_actual in opciones:
                        entry.set(valor_actual)  
                    else:
                        entry.set(opciones[0])  
                else:
                    entry = ctk.CTkEntry(field_frame, width=300)
                    entry.pack(side="left", expand=True, fill="x")                   
                    
                    valor_actual = row_data[i]
                    if valor_actual:  
                        entry.insert(0, valor_actual)

                entries[campo] = entry

            def actualizar():
                """Update the selected record"""
                try:
                    logging.info(f"Inicio de actualización para N_control: {N_control}")
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    # Collect values
                    valores = [entries[campo].get() for campo in campos] + [N_control]
                    valores = [texto.upper() for texto in valores]

                    # Execute update
                    cursor.execute(qry.query_update_des, valores)
                    conn.commit()

                    logging.info(f"Registro {N_control} actualizado exitosamente.")
                    messagebox.showinfo("Éxito", "Registro actualizado correctamente")
                    editar_dialog.destroy()

                    # Refresh data in the frame
                    frame.load_data()

                except Exception as e:
                    logging.error(f"Error al actualizar registro {N_control}: {e}")
                    messagebox.showerror("Error", f"No se pudo actualizar el registro: {str(e)}")
                finally:
                    if "conn" in locals():
                        conn.close()

            # Update button
            actualizar_btn = ctk.CTkButton(input_frame, text="Actualizar", command=actualizar)
            actualizar_btn.pack(pady=20)

        except Exception as e:
            logging.error(f"Error en el proceso de edición: {e}")
            messagebox.showerror("Error", f"Ocurrió un problema al intentar editar el registro: {str(e)}")

 
    def load_data_Costos(self, frame):
        cursor = None  # Initialize cursor with None
        conn = None  # Initialize conn with None
        try:
            conn_str = (
                f"DRIVER={os.getenv('DB1_DRIVER')};"
                f"SERVER={os.getenv('DB1_SERVER')};"
                f"DATABASE={os.getenv('DB1_DATABASE')};"
                f"UID={os.getenv('DB1_UID')};"
                f"PWD={os.getenv('DB1_PWD')}"
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
      
            cursor.execute(qry.load_query_costos)
            data = cursor.fetchall()        
            frame.sheet.headers(hd.headers_load_cos)
            # Convert data to list of lists with string values
            formatted_data = []
            for row in data:
                 # Convert null values to empty string
                formatted_row = [str(value) if value is not None else "" for value in row]
               
                if any(value == "" for value in formatted_row):
                    formatted_row.append("X")  
                else:
                    formatted_row.append("✓")  
                formatted_data.append(formatted_row)
            
            frame.original_data = formatted_data
            frame.sheet.set_sheet_data(formatted_data)

        except pyodbc.Error as e:
            print(f"An error occurred while loading data: {str(e)}", file=sys.stderr)
            messagebox.showerror(title="Error", message=f"No se pudo cargar los datos: {str(e)}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Placeholder methods for create and edit
    def create_record_Costos(self, frame):
        return;
    def edit_record_Costos(self, frame):
        try:
            selected_rows = frame.sheet.get_selected_rows()

            if not selected_rows:
                messagebox.showerror(title="Error", message="Seleccione un registro para editar")
                return

            # Get the data of the selected row from filtered or original data
            selected_data = next(iter(selected_rows))
            row_data = frame.sheet.get_row_data(selected_data)
            N_control = row_data[0]  # Assuming N_control is in the first column (index 0)

            if not N_control:
                messagebox.showerror("Error", "No se pudo obtener el N_control del registro seleccionado.")
                logging.error("El N_control es nulo o no válido.")
                return

            logging.info(f"N_control seleccionado: {N_control}")

            # Create edit dialog window
            editar_dialog = ctk.CTkToplevel(frame)
            editar_dialog.title("Editar Registro")
            editar_dialog.after(201, lambda :editar_dialog.iconbitmap(icon_path))
            editar_dialog.geometry("400x500")
            # Dictionary to store entry/combobox widgets
            entries = {}
            # Function to fetch item data from database
            def fetch_item_data():
                try:
                    conn_str = (
                        f"DRIVER={os.getenv('DB2_DRIVER')};"
                        f"SERVER={os.getenv('DB2_SERVER')};"
                        f"DATABASE={os.getenv('DB2_DATABASE')};"
                        f"UID={os.getenv('DB2_UID')};"
                        f"PWD={os.getenv('DB2_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    # Query to fetch item codes and descriptions
                    query = """
                    SELECT in_items.itecodigo AS [Codigo item], 
                        in_items.itedesccort AS [Descripcion item] 
                    FROM dbo.in_items 
                    WHERE in_items.itecompania = '01' AND 
                        (in_items.itecodigo LIKE 'ME%')
                    """
                    cursor.execute(query)
                    return cursor.fetchall()
                except Exception as e:
                    logging.error(f"Error fetching item data: {e}")
                    messagebox.showerror("Error", f"No se pudieron cargar los datos de elementos: {str(e)}")
                    return []
                finally:
                    if 'conn' in locals():
                        conn.close()

            # Create scrollable frame
            scrollable_frame = ef.ScrollableFrame(editar_dialog)
            scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)

            # Create input frame
            input_frame = ctk.CTkFrame(scrollable_frame)
            input_frame.pack(fill="x", expand=True)
            
            # Fetch item data once
            item_data = fetch_item_data()

            # Create a dictionary for Codigo and Descripcion
            item_dict = {row[0]: row[1] for row in item_data}
            item_codes = list(item_dict.keys())

            # Create a combined list for combobox display
            item_combined_list = [f"{codigo} - {descripcion}" for codigo, descripcion in item_dict.items()]

            # Specific fields to use the combined list
            special_fields = ["Descripcion Envase No Decorado","Descripcion Envase Decorado", "Descripcion Tapa No Decorado","Descripcion Tapa Decorada", "Foil", "Banda", "Caja Plegadiza","Descripcion Escurridor"]
            
            for i, campo in enumerate(hd.campos_update_cost, start=5):
                field_frame = ctk.CTkFrame(input_frame)
                field_frame.pack(pady=5, padx=20, fill="x")

                label = ctk.CTkLabel(field_frame, text=campo, width=150, anchor="w")
                label.pack(side="left", padx=(0, 10))

                # Prellenar el valor actual desde row_data
                valor_actual = row_data[i] if i < len(row_data) else ""

                if campo in special_fields:
                    # Crear combobox para campos especiales
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=item_combined_list)
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in item_combined_list:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")

                    # Funcionalidad de filtrado
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [
                            f"{codigo} - {descripcion}"
                            for codigo, descripcion in item_dict.items()
                            if input_text in codigo.lower() or input_text in descripcion.lower()
                        ]
                        combobox.configure(values=filtered_values)

                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                elif campo in hd.desplegables_cost:
                    # Crear combobox para campos en desplegables
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=hd.desplegables_cost[campo])
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in hd.desplegables_cost[campo]:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")
                    
                    # Funcionalidad de filtrado
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [value for value in hd.desplegables_cost[campo] if input_text in value.lower()]
                        combobox.configure(values=filtered_values)

                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                elif campo.startswith("Codigo"):
                    # Crear combobox para campos que comienzan con "Codigo"
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=item_codes)
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in item_codes:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")
                    
                    # Funcionalidad de filtrado para campos de Codigo
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [codigo for codigo in item_codes if input_text in codigo.lower()]
                        combobox.configure(values=filtered_values)
                    
                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                else:
                    # Entrada regular para campos no especiales y no de Codigo
                    entry = ctk.CTkEntry(field_frame, width=300)
                    entry.pack(side="left", expand=True, fill="x")
                    entry.insert(0, valor_actual)
                    entries[campo] = entry
            def actualizar():
                """Update the selected record"""
                try:
                    logging.info(f"Inicio de actualización para N_control: {N_control}")
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()
                    # Collect values in the correct order
                    valores = []
                    for campo in hd.ordered_campos_cost:
                        if campo in entries:
                            valores.append(entries[campo].get())
                        else:
                            valores.append("")  # Default to empty string if the field is not present

                    valores.append(N_control)  # Add N_control at the end
                    valores = [texto.upper() for texto in valores]  # Convert all values to uppercase

                    # Execute update query
                    cursor.execute(qry.update_query_costos, valores)
                    conn.commit()

                    logging.info(f"Registro {N_control} actualizado exitosamente.")
                    messagebox.showinfo("Éxito", "Registro actualizado correctamente")
                    editar_dialog.destroy()

                    # Refresh data in the frame
                    frame.load_data()

                except Exception as e:
                    logging.error(f"Error al actualizar registro {N_control}: {e}")
                    messagebox.showerror("Error", f"No se pudo actualizar el registro: {str(e)}")
                finally:
                    if "conn" in locals():
                        conn.close()

            # Update button
            actualizar_btn = ctk.CTkButton(input_frame, text="Actualizar", command=actualizar)
            actualizar_btn.pack(pady=20)

        except Exception as e:
            logging.error(f"Error en el proceso de edición: {e}")
            messagebox.showerror("Error", f"Ocurrió un problema al intentar editar el registro: {str(e)}")

    def load_data_Calidad(self, frame):
        cursor = None
        conn = None
        try:
            conn_str = (
                f"DRIVER={os.getenv('DB1_DRIVER')};"
                f"SERVER={os.getenv('DB1_SERVER')};"
                f"DATABASE={os.getenv('DB1_DATABASE')};"
                f"UID={os.getenv('DB1_UID')};"
                f"PWD={os.getenv('DB1_PWD')}"
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute(qry.load_query_calidad)
            data = cursor.fetchall()
            
            frame.sheet.headers(hd.headers_load_cal)

            # Convert data to list of lists with string values
            formatted_data = []
            for row in data:
                # Convert null values to empty string
                formatted_row = [str(value) if value is not None else "" for value in row]
                if any(value == "" for value in formatted_row):
                    formatted_row.append("X")  
                else:
                    formatted_row.append("✓")  
                formatted_data.append(formatted_row)
              
            frame.original_data = formatted_data
            frame.sheet.set_sheet_data(formatted_data)

        except pyodbc.Error as e:
            print(f"An error occurred while loading data: {str(e)}", file=sys.stderr)
            messagebox.showerror(title="Error", message=f"No se pudo cargar los datos: {str(e)}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Placeholder methods for create and edit
    def create_record_Calidad(self, frame):
       return
    def edit_record_Calidad(self, frame):
        try:
            selected_rows = frame.sheet.get_selected_rows()

            if not selected_rows:
                messagebox.showerror(title="Error", message="Seleccione un registro para editar")
                return

            selected_data = next(iter(selected_rows))
            row_data = frame.sheet.get_row_data(selected_data)
            N_control = row_data[0]

            if not N_control:
                messagebox.showerror("Error", "No se pudo obtener el N_control del registro seleccionado.")
                logging.error("El N_control es nulo o no válido.")
                return

            logging.info(f"N_control seleccionado: {N_control}")

            editar_dialog = ctk.CTkToplevel(frame)
            editar_dialog.title("Editar Registro")
            editar_dialog.after(201, lambda: editar_dialog.iconbitmap(icon_path))
            editar_dialog.geometry("400x500")

            entries = {}

            scrollable_frame = ef.ScrollableFrame(editar_dialog)
            scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)

            input_frame = ctk.CTkFrame(scrollable_frame)
            input_frame.pack(fill="x", expand=True)

            for i, campo in enumerate(hd.campos_update_cal, start=5):
                field_frame = ctk.CTkFrame(input_frame)
                field_frame.pack(pady=5, padx=20, fill="x")

                label = ctk.CTkLabel(field_frame, text=campo, width=150, anchor="w")
                label.pack(side="left", padx=(0, 10))
                dictionary= { "Titular NSOC": ["CLIENTE", "GB LAB","CLIENTE INT"],
                            "Uso Exclusivo": ["SI", "NO"],
                            "Forma Cosmetica": ["EMULSION", "SOLUCION", "GEL", "SUSPENSION"]}
                if campo in dictionary:
                    # Crear un ComboBox para "Titular Nsoc"
                    options= dictionary[campo]
                    combo = ctk.CTkComboBox(field_frame, values=options, width=300)
                    combo.pack(side="left", expand=True, fill="x")
                    # Pre-fill current value from row_data
                    combo.set(row_data[i])  
                    entries[campo] = combo
                else:
                    entry = ctk.CTkEntry(field_frame, width=300)
                    entry.pack(side="left", expand=True, fill="x")
                    valor_actual = row_data[i]
                    entry.insert(0, valor_actual)
                    entries[campo] = entry

            def actualizar():
                try:
                    logging.info(f"Inicio de actualización para N_control: {N_control}")
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    valores = [entries[campo].get() for campo in hd.campos_update_cal] + [N_control]
                    valores = [texto.upper() for texto in valores]

                    cursor.execute(qry.update_query_calidad, valores)
                    conn.commit()

                    logging.info(f"Registro {N_control} actualizado exitosamente.")
                    messagebox.showinfo("Éxito", "Registro actualizado correctamente")
                    editar_dialog.destroy()
                    frame.load_data()

                except Exception as e:
                    logging.error(f"Error al actualizar registro {N_control}: {e}")
                    messagebox.showerror("Error", f"No se pudo actualizar el registro: {str(e)}")
                finally:
                    if "conn" in locals():
                        conn.close()

            actualizar_btn = ctk.CTkButton(input_frame, text="Actualizar", command=actualizar)
            actualizar_btn.pack(pady=20)

        except Exception as e:
            logging.error(f"Error en el proceso de edición: {e}")
            messagebox.showerror("Error", f"Ocurrió un problema al intentar editar el registro: {str(e)}")
    def load_data_Produccion(self, frame):
        cursor = None
        conn = None
        try:
            conn_str = (
                f"DRIVER={os.getenv('DB1_DRIVER')};"
                f"SERVER={os.getenv('DB1_SERVER')};"
                f"DATABASE={os.getenv('DB1_DATABASE')};"
                f"UID={os.getenv('DB1_UID')};"
                f"PWD={os.getenv('DB1_PWD')}"
            )
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
           
            cursor.execute(qry.load_query_produccion)
            data = cursor.fetchall()
            
            frame.sheet.headers(hd.headers_load_prod)
            # Convert data to list of lists with string values
            formatted_data = []
            for row in data:
                # Convert null values to empty string
                formatted_row = [str(value) if value is not None else "" for value in row]
               
                if any(value == "" for value in formatted_row):
                    formatted_row.append("X") 
                else:
                    formatted_row.append("✓") 
                formatted_data.append(formatted_row)

            frame.original_data = formatted_data
            frame.sheet.set_sheet_data(formatted_data)

        except pyodbc.Error as e:
            print(f"An error occurred while loading data: {str(e)}", file=sys.stderr)
            messagebox.showerror(title="Error", message=f"No se pudo cargar los datos: {str(e)}")

        finally:
            if cursor:
                cursor.close()
            if conn:
                conn.close()

    # Placeholder methods for create and edit
    def create_record_Produccion(self, frame):
       return
    def edit_record_Produccion(self, frame):
        try:
            selected_rows = frame.sheet.get_selected_rows()

            if not selected_rows:
                messagebox.showerror(title="Error", message="Seleccione un registro para editar")
                return

            # Get the data of the selected row from filtered or original data
            selected_data = next(iter(selected_rows))
            row_data = frame.sheet.get_row_data(selected_data)
            N_control = row_data[0]  # Assuming N_control is in the first column (index 0)

            if not N_control:
                messagebox.showerror("Error", "No se pudo obtener el N_control del registro seleccionado.")
                logging.error("El N_control es nulo o no válido.")
                return

            logging.info(f"N_control seleccionado: {N_control}")

            # Create edit dialog window
            editar_dialog = ctk.CTkToplevel(frame)
            editar_dialog.title("Editar Registro")
            editar_dialog.after(201, lambda :editar_dialog.iconbitmap(icon_path))
            editar_dialog.geometry("400x500")
            # Fields (excluding N_control)
            
            def fetch_item_data():
                try:
                    conn_str = (
                        f"DRIVER={os.getenv('DB2_DRIVER')};"
                        f"SERVER={os.getenv('DB2_SERVER')};"
                        f"DATABASE={os.getenv('DB2_DATABASE')};"
                        f"UID={os.getenv('DB2_UID')};"
                        f"PWD={os.getenv('DB2_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    # Query to fetch item codes and descriptions
                    query = """
                    SELECT in_items.itecodigo AS [Codigo item], 
                        in_items.itedesccort AS [Descripcion item] 
                    FROM dbo.in_items 
                    WHERE in_items.itecompania = '01' AND 
                        (in_items.itecodigo LIKE 'ME%')
                    """
                    cursor.execute(query)
                    return cursor.fetchall()
                except Exception as e:
                    logging.error(f"Error fetching item data: {e}")
                    messagebox.showerror("Error", f"No se pudieron cargar los datos de elementos: {str(e)}")
                    return []
                finally:
                    if 'conn' in locals():
                        conn.close()
            # Dictionary to store entry widgets
            entries = {}
            # Create the scrollable frame
            scrollable_frame = ef.ScrollableFrame(editar_dialog)
            scrollable_frame.pack(pady=10, padx=20, fill="both", expand=True)
            item_data = fetch_item_data()
            # Create a dictionary for Codigo and Descripcion
            item_dict = {row[0]: row[1] for row in item_data}
            item_codes = list(item_dict.keys())
            # Create a frame for the input fields
            
            item_combined_list = [f"{codigo} - {descripcion}" for codigo, descripcion in item_dict.items()]

            # Specific fields to use the combined list
            special_fields = ["Caja Corrugada"]

            input_frame = ctk.CTkFrame(scrollable_frame)
            input_frame.pack(fill="x", expand=True)
            # Create input fields for each column
            desplegables = {
                "Rotulo": ["SI", "NO"],
                "Caja Impresa": ["SI", "NO"],
                "Lleva Parrilla": ["SI", "NO"]
            }
            for i, campo in enumerate(hd.campos_update_prod, start=5):
                field_frame = ctk.CTkFrame(input_frame)
                field_frame.pack(pady=5, padx=20, fill="x")

                label = ctk.CTkLabel(field_frame, text=campo, width=150, anchor="w")
                label.pack(side="left", padx=(0, 10))

                # Prellenar el valor actual desde row_data
                valor_actual = row_data[i] if i < len(row_data) else ""

                if campo in special_fields:
                    # Crear combobox para campos especiales
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=item_combined_list)
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in item_combined_list:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")

                    # Funcionalidad de filtrado
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [
                            f"{codigo} - {descripcion}"
                            for codigo, descripcion in item_dict.items()
                            if input_text in codigo.lower() or input_text in descripcion.lower()
                        ]
                        combobox.configure(values=filtered_values)

                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                elif campo in desplegables:
                    # Crear combobox para campos en desplegables
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=desplegables[campo])
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in desplegables[campo]:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")
                    
                    # Funcionalidad de filtrado
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [value for value in desplegables[campo] if input_text in value.lower()]
                        combobox.configure(values=filtered_values)

                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                elif campo.startswith("Codigo"):
                    # Crear combobox para campos que comienzan con "Codigo"
                    combobox = ctk.CTkComboBox(field_frame, width=300, values=item_codes)
                    combobox.pack(side="left", expand=True, fill="x")
                    
                    if valor_actual and valor_actual in item_codes:
                        combobox.set(valor_actual)
                    else:
                        combobox.set("")
                    
                    # Funcionalidad de filtrado para campos de Codigo
                    def on_combobox_input(event, combobox=combobox):
                        input_text = combobox.get().lower()
                        filtered_values = [codigo for codigo in item_codes if input_text in codigo.lower()]
                        combobox.configure(values=filtered_values)
                    
                    combobox.bind("<KeyRelease>", on_combobox_input)
                    entries[campo] = combobox
                
                else:
                    # Entrada regular para campos no especiales y no de Codigo
                    entry = ctk.CTkEntry(field_frame, width=300)
                    entry.pack(side="left", expand=True, fill="x")
                    entry.insert(0, valor_actual)
                    entries[campo] = entry

            def actualizar():
                """Update the selected record"""
                try:
                    logging.info(f"Inicio de actualización para N_control: {N_control}")
                    conn_str = (
                        f"DRIVER={os.getenv('DB1_DRIVER')};"
                        f"SERVER={os.getenv('DB1_SERVER')};"
                        f"DATABASE={os.getenv('DB1_DATABASE')};"
                        f"UID={os.getenv('DB1_UID')};"
                        f"PWD={os.getenv('DB1_PWD')}"
                    )
                    conn = pyodbc.connect(conn_str)
                    cursor = conn.cursor()

                    # Collect values
                    valores = [entries[campo].get() for campo in hd.campos_update_prod] + [N_control]
                    valores = [texto.upper() for texto in valores]

                    # Execute update
                    cursor.execute(qry.edit_query_produccion, valores)
                    conn.commit()

                    logging.info(f"Registro {N_control} actualizado exitosamente.")
                    messagebox.showinfo("Éxito", "Registro actualizado correctamente")
                    editar_dialog.destroy()
                    
                    # Refresh data in the frame
                    frame.load_data()

                except Exception as e:
                    logging.error(f"Error al actualizar registro {N_control}: {e}")
                    messagebox.showerror("Error", f"No se pudo actualizar el registro: {str(e)}")
                finally:
                    if "conn" in locals():
                        conn.close()

            # Update button
            actualizar_btn = ctk.CTkButton(input_frame, text="Actualizar", command=actualizar)
            actualizar_btn.pack(pady=20)

        except Exception as e:
            logging.error(f"Error en el proceso de edición: {e}")
            messagebox.showerror("Error", f"Ocurrió un problema al intentar editar el registro: {str(e)}")
