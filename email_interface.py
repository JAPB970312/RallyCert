# email_interface.py
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from email_sender import EmailSender
import threading
import re

class EmailSenderInterface:
    def __init__(self, parent):
        self.parent = parent
        self.email_sender = None
        self.excel_data = None
        self.pdf_folder = ""
        
        self.create_interface()
        self.setup_text_formatting()
    
    def create_interface(self):
        """Crea la interfaz para el envío de correos"""
        # Crear ventana secundaria
        self.window = tk.Toplevel(self.parent)
        self.window.title("Envío de Constancias por Correo - Outlook/Office365")
        self.window.geometry("800x900")
        self.window.configure(bg='#f0f0f0')
        
        # Frame principal
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        title_label = ttk.Label(main_frame, text="Envío de Constancias por Correo - Outlook/Office365", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Información de configuración Outlook
        outlook_info_frame = ttk.LabelFrame(main_frame, text="📧 Configuración Outlook/Office365", padding="10")
        outlook_info_frame.grid(row=1, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Información importante para Outlook
        info_text = """Para enviar desde Outlook/Office365:
• Usa tu correo institucional (@unison.mx) o personal
• Usa tu CONTRASEÑA NORMAL (no aplicación)
• Asegúrate de tener autenticación en dos pasos DESACTIVADA
• Dominios soportados: @unison.mx, @outlook.com, @hotmail.com, etc."""
        
        info_label = ttk.Label(outlook_info_frame, text=info_text, justify=tk.LEFT, 
                              background='#e7f3ff', padding=10)
        info_label.grid(row=0, column=0, columnspan=2, sticky='we')
        
        # Configuración de correo remitente
        email_frame = ttk.LabelFrame(main_frame, text="Configuración de Correo", padding="10")
        email_frame.grid(row=2, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Correo electrónico
        ttk.Label(email_frame, text="Correo electrónico:").grid(row=0, column=0, sticky='w', pady=5)
        self.email_entry = ttk.Entry(email_frame, width=40)
        self.email_entry.grid(row=0, column=1, sticky='we', pady=5, padx=(10, 0))
        self.email_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        
        # Placeholder con ejemplo
        self.email_entry.insert(0, "tu_correo@unison.mx")
        
        # Contraseña (con opción para mostrar/ocultar)
        ttk.Label(email_frame, text="Contraseña:").grid(row=1, column=0, sticky='w', pady=5)
        self.password_frame = ttk.Frame(email_frame)
        self.password_frame.grid(row=1, column=1, sticky='we', pady=5, padx=(10, 0))
        
        self.password_entry = ttk.Entry(self.password_frame, width=40, show="*")
        self.password_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.password_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        
        self.show_password_var = tk.BooleanVar()
        self.show_password_btn = ttk.Checkbutton(self.password_frame, text="👁", 
                                                variable=self.show_password_var,
                                                command=self.toggle_password_visibility)
        self.show_password_btn.pack(side=tk.RIGHT, padx=(5, 0))
        
        # Información sobre contraseña
        pass_info = ttk.Label(email_frame, text="Usa tu contraseña normal (no contraseña de aplicación)", 
                             foreground='gray', font=('Arial', 8))
        pass_info.grid(row=2, column=1, sticky='w', pady=(0, 5))
        
        # Nombre del remitente
        ttk.Label(email_frame, text="Nombre del remitente:").grid(row=3, column=0, sticky='w', pady=5)
        self.sender_name_entry = ttk.Entry(email_frame, width=40)
        self.sender_name_entry.grid(row=3, column=1, sticky='we', pady=5, padx=(10, 0))
        self.sender_name_entry.insert(0, "Universidad de Sonora - Sistema de Constancias")
        self.sender_name_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        
        # Botón de prueba de conexión
        self.test_btn = ttk.Button(email_frame, text="Probar Conexión", 
                                  command=self.test_connection)
        self.test_btn.grid(row=4, column=0, columnspan=2, pady=10)
        
        # Selección de archivos
        files_frame = ttk.LabelFrame(main_frame, text="Archivos", padding="10")
        files_frame.grid(row=3, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Carpeta de PDFs
        ttk.Label(files_frame, text="Carpeta de PDFs:").grid(row=0, column=0, sticky='w', pady=5)
        self.pdf_folder_entry = ttk.Entry(files_frame, width=30)
        self.pdf_folder_entry.grid(row=0, column=1, sticky='we', pady=5, padx=(10, 0))
        self.pdf_folder_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        ttk.Button(files_frame, text="Seleccionar", 
                  command=self.select_pdf_folder).grid(row=0, column=2, padx=(5, 0))
        
        # Archivo Excel
        ttk.Label(files_frame, text="Archivo Excel:").grid(row=1, column=0, sticky='w', pady=5)
        self.excel_file_entry = ttk.Entry(files_frame, width=30)
        self.excel_file_entry.grid(row=1, column=1, sticky='we', pady=5, padx=(10, 0))
        self.excel_file_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        ttk.Button(files_frame, text="Seleccionar", 
                  command=self.select_excel_file).grid(row=1, column=2, padx=(5, 0))
        
        # Configuración de columnas
        columns_frame = ttk.LabelFrame(main_frame, text="Configuración de Columnas", padding="10")
        columns_frame.grid(row=4, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Menús desplegables para columnas
        ttk.Label(columns_frame, text="Columna Nombre:").grid(row=0, column=0, sticky='w', pady=5)
        self.name_column_combo = ttk.Combobox(columns_frame, state="readonly", width=25)
        self.name_column_combo.grid(row=0, column=1, sticky='we', pady=5, padx=(10, 0))
        self.name_column_combo.bind('<<ComboboxSelected>>', lambda e: self.validate_form())
        
        ttk.Label(columns_frame, text="Columna Correo:").grid(row=1, column=0, sticky='w', pady=5)
        self.email_column_combo = ttk.Combobox(columns_frame, state="readonly", width=25)
        self.email_column_combo.grid(row=1, column=1, sticky='we', pady=5, padx=(10, 0))
        self.email_column_combo.bind('<<ComboboxSelected>>', lambda e: self.validate_form())
        
        ttk.Label(columns_frame, text="Columna Archivo PDF:").grid(row=2, column=0, sticky='w', pady=5)
        self.filename_column_combo = ttk.Combobox(columns_frame, state="readonly", width=25)
        self.filename_column_combo.grid(row=2, column=1, sticky='we', pady=5, padx=(10, 0))
        self.filename_column_combo.bind('<<ComboboxSelected>>', lambda e: self.validate_form())
        
        # Contenido del correo CON BARRA DE HERRAMIENTAS DE FORMATO
        content_frame = ttk.LabelFrame(main_frame, text="Contenido del Correo", padding="10")
        content_frame.grid(row=5, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Barra de herramientas de formato
        format_toolbar = ttk.Frame(content_frame)
        format_toolbar.grid(row=0, column=0, columnspan=2, sticky='we', pady=(0, 5))
        
        # Botones de formato
        self.btn_bold = ttk.Button(format_toolbar, text="𝐁", width=3, 
                                  command=lambda: self.format_text("bold"))
        self.btn_bold.pack(side=tk.LEFT, padx=2)
        
        self.btn_italic = ttk.Button(format_toolbar, text="𝐼", width=3,
                                   command=lambda: self.format_text("italic"))
        self.btn_italic.pack(side=tk.LEFT, padx=2)
        
        self.btn_underline = ttk.Button(format_toolbar, text="𝑈", width=3,
                                      command=lambda: self.format_text("underline"))
        self.btn_underline.pack(side=tk.LEFT, padx=2)
        
        # Separador
        ttk.Separator(format_toolbar, orient='vertical').pack(side=tk.LEFT, padx=10, fill='y')
        
        # Justificación
        ttk.Label(format_toolbar, text="Alineación:").pack(side=tk.LEFT, padx=(0, 2))
        self.justification_var = tk.StringVar(value="left")
        ttk.Radiobutton(format_toolbar, text="⬅", value="left", 
                       variable=self.justification_var,
                       command=lambda: self.format_text("justify_left")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(format_toolbar, text="⬌", value="center", 
                       variable=self.justification_var,
                       command=lambda: self.format_text("justify_center")).pack(side=tk.LEFT, padx=2)
        ttk.Radiobutton(format_toolbar, text="➡", value="right", 
                       variable=self.justification_var,
                       command=lambda: self.format_text("justify_right")).pack(side=tk.LEFT, padx=2)
        
        # Separador
        ttk.Separator(format_toolbar, orient='vertical').pack(side=tk.LEFT, padx=10, fill='y')
        
        # Configuración de párrafo
        ttk.Label(format_toolbar, text="Interlineado:").pack(side=tk.LEFT, padx=(0, 2))
        self.line_spacing_combo = ttk.Combobox(format_toolbar, 
                                              values=["Simple", "1.15", "1.5", "Doble"], 
                                              width=8, state="readonly")
        self.line_spacing_combo.set("1.15")
        self.line_spacing_combo.pack(side=tk.LEFT, padx=2)
        self.line_spacing_combo.bind('<<ComboboxSelected>>', self.apply_line_spacing)
        
        ttk.Label(format_toolbar, text="Sangría:").pack(side=tk.LEFT, padx=(10, 2))
        self.indent_combo = ttk.Combobox(format_toolbar, 
                                        values=["Ninguna", "Pequeña", "Mediana", "Grande"], 
                                        width=8, state="readonly")
        self.indent_combo.set("Ninguna")
        self.indent_combo.pack(side=tk.LEFT, padx=2)
        self.indent_combo.bind('<<ComboboxSelected>>', self.apply_indentation)
        
        # Asunto
        ttk.Label(content_frame, text="Asunto:").grid(row=1, column=0, sticky='w', pady=5)
        self.subject_entry = ttk.Entry(content_frame, width=50)
        self.subject_entry.grid(row=1, column=1, sticky='we', pady=5, padx=(10, 0))
        self.subject_entry.insert(0, "Constancia de Participación - Universidad de Sonora")
        self.subject_entry.bind('<KeyRelease>', lambda e: self.validate_form())
        
        # Cuerpo del correo con mejor formato
        ttk.Label(content_frame, text="Cuerpo del mensaje:").grid(row=2, column=0, sticky='nw', pady=5)
        
        # Frame para el área de texto con scrollbar
        text_frame = ttk.Frame(content_frame)
        text_frame.grid(row=2, column=1, sticky='nsew', pady=5, padx=(10, 0))
        
        # Configurar grid weights para expansión
        content_frame.rowconfigure(2, weight=1)
        content_frame.columnconfigure(1, weight=1)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        
        # Área de texto con formato mejorado
        self.body_text = tk.Text(text_frame, width=50, height=8, wrap=tk.WORD,
                                font=('Arial', 10), spacing1=2, spacing2=1, spacing3=1,
                                padx=10, pady=10)
        
        # Scrollbar para el cuerpo del mensaje
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=scrollbar.set)
        
        # Empaquetar widgets
        self.body_text.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')
        
        # Texto predeterminado con mejor formato
        default_body = """Estimado/a {nombre},

Le hacemos llegar su constancia de participación emitida por la Universidad de Sonora. Este documento certifica su asistencia y participación en nuestro evento académico.

Agradecemos su valiosa contribución y esperamos contar con su participación en futuras actividades.

Saludos cordiales,
Departamento de Constancias
Universidad de Sonora"""
        
        self.body_text.insert('1.0', default_body)
        self.body_text.bind('<KeyRelease>', lambda e: self.validate_form())
        
        # Atajos de teclado
        self.body_text.bind('<Control-b>', lambda e: self.format_text("bold"))
        self.body_text.bind('<Control-i>', lambda e: self.format_text("italic"))
        self.body_text.bind('<Control-u>', lambda e: self.format_text("underline"))
        
        # Información sobre placeholders
        placeholder_info = ttk.Label(content_frame, 
                                   text="💡 Placeholders disponibles: {nombre}, {Nombre}, {fecha} | Use Ctrl+B para negrita",
                                   foreground='gray', font=('Arial', 8))
        placeholder_info.grid(row=3, column=1, sticky='w', pady=(5, 0))
        
        # Botones de acción
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        self.send_btn = ttk.Button(buttons_frame, text="Enviar Correos", 
                                  command=self.start_sending_emails, state='disabled')
        self.send_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(buttons_frame, text="Cancelar", 
                  command=self.window.destroy).pack(side=tk.LEFT)
        
        # Barra de progreso
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=7, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Etiqueta de estado
        self.status_label = ttk.Label(main_frame, text="Esperando configuración...", foreground='blue')
        self.status_label.grid(row=8, column=0, columnspan=2)
        
        # Área de log detallado
        log_frame = ttk.LabelFrame(main_frame, text="Log de Ejecución", padding="10")
        log_frame.grid(row=9, column=0, columnspan=2, sticky='we', pady=(0, 10))
        
        # Text area para log
        self.log_text = tk.Text(log_frame, height=8, width=70)
        self.log_text.grid(row=0, column=0, sticky='nsew')
        
        # Scrollbar para el log
        log_scrollbar = ttk.Scrollbar(log_frame, orient='vertical', command=self.log_text.yview)
        log_scrollbar.grid(row=0, column=1, sticky='ns')
        self.log_text.configure(yscrollcommand=log_scrollbar.set)
        
        # Configurar grid weights
        for frame in [outlook_info_frame, email_frame, files_frame, columns_frame, content_frame, log_frame]:
            frame.columnconfigure(1, weight=1)
        
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
    
    def setup_text_formatting(self):
        """Configura los estilos de texto"""
        # Configurar tags para formato
        self.body_text.tag_configure("bold", font=('Arial', 10, 'bold'))
        self.body_text.tag_configure("italic", font=('Arial', 10, 'italic'))
        self.body_text.tag_configure("underline", font=('Arial', 10, 'underline'),
                                   underline=True)
        
        # Tags para justificación
        self.body_text.tag_configure("justify_left", justify='left')
        self.body_text.tag_configure("justify_center", justify='center')
        self.body_text.tag_configure("justify_right", justify='right')
    
    def toggle_password_visibility(self):
        """Alterna entre mostrar y ocultar la contraseña"""
        if self.show_password_var.get():
            self.password_entry.config(show="")
        else:
            self.password_entry.config(show="*")
    
    def format_text(self, format_type):
        """Aplica formato al texto seleccionado"""
        try:
            if not self.body_text.tag_ranges("sel"):
                return  # No hay texto seleccionado
            
            start, end = self.body_text.index("sel.first"), self.body_text.index("sel.last")
            
            if format_type == "bold":
                # Alternar negrita
                current_tags = self.body_text.tag_names(start)
                if "bold" in current_tags:
                    self.body_text.tag_remove("bold", start, end)
                else:
                    self.body_text.tag_add("bold", start, end)
                    
            elif format_type == "italic":
                # Alternar itálica
                current_tags = self.body_text.tag_names(start)
                if "italic" in current_tags:
                    self.body_text.tag_remove("italic", start, end)
                else:
                    self.body_text.tag_add("italic", start, end)
                    
            elif format_type == "underline":
                # Alternar subrayado
                current_tags = self.body_text.tag_names(start)
                if "underline" in current_tags:
                    self.body_text.tag_remove("underline", start, end)
                else:
                    self.body_text.tag_add("underline", start, end)
                    
            elif format_type.startswith("justify"):
                # Aplicar justificación
                alignment = format_type.replace("justify_", "")
                self.body_text.tag_configure("justify", justify=alignment)
                self.body_text.tag_add("justify", "1.0", "end")
                
        except Exception as e:
            print(f"Error aplicando formato: {e}")
    
    def apply_line_spacing(self, event=None):
        """Aplica el interlineado seleccionado"""
        try:
            spacing_map = {
                "Simple": 1.0,
                "1.15": 1.15, 
                "1.5": 1.5,
                "Doble": 2.0
            }
            spacing = spacing_map.get(self.line_spacing_combo.get(), 1.15)
            self.body_text.configure(spacing2=int((spacing - 1.0) * 10))
        except Exception as e:
            print(f"Error aplicando interlineado: {e}")
    
    def apply_indentation(self, event=None):
        """Aplica la sangría seleccionada"""
        try:
            indent_map = {
                "Ninguna": 0,
                "Pequeña": 20,
                "Mediana": 40,
                "Grande": 60
            }
            indent_px = indent_map.get(self.indent_combo.get(), 0)
            self.body_text.configure(padx=indent_px)
        except Exception as e:
            print(f"Error aplicando sangría: {e}")
    
    def select_pdf_folder(self):
        """Selecciona la carpeta que contiene los PDFs"""
        folder = filedialog.askdirectory(title="Seleccionar carpeta de PDFs")
        if folder:
            self.pdf_folder_entry.delete(0, tk.END)
            self.pdf_folder_entry.insert(0, folder)
            self.pdf_folder = folder
            self.validate_form()
    
    def select_excel_file(self):
        """Selecciona el archivo Excel con los datos"""
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file_path:
            self.excel_file_entry.delete(0, tk.END)
            self.excel_file_entry.insert(0, file_path)
            self.load_excel_columns(file_path)
    
    def load_excel_columns(self, excel_path):
        """Carga las columnas del archivo Excel en los combobox"""
        try:
            # Verificar que el archivo existe
            if not os.path.exists(excel_path):
                messagebox.showerror("Error", f"El archivo {excel_path} no existe")
                return
            
            # Cargar el Excel
            self.excel_data = pd.read_excel(excel_path)
            
            # Verificar que el DataFrame no esté vacío
            if len(self.excel_data) == 0:
                messagebox.showwarning("Advertencia", "El archivo Excel está vacío")
                return
            
            columns = self.excel_data.columns.tolist()
            
            # Actualizar comboboxes
            self.name_column_combo['values'] = columns
            self.email_column_combo['values'] = columns
            self.filename_column_combo['values'] = columns
            
            # Seleccionar automáticamente columnas comunes
            common_names = ['nombre', 'name', 'participante', 'estudiante', 'alumno']
            common_emails = ['email', 'correo', 'mail', 'e-mail']
            common_files = ['archivo', 'filename', 'pdf', 'constancia', 'documento']
            
            for col in columns:
                col_lower = col.lower()
                if any(name in col_lower for name in common_names) and not self.name_column_combo.get():
                    self.name_column_combo.set(col)
                if any(email in col_lower for email in common_emails) and not self.email_column_combo.get():
                    self.email_column_combo.set(col)
                if any(file in col_lower for file in common_files) and not self.filename_column_combo.get():
                    self.filename_column_combo.set(col)
            
            # Mostrar información del archivo cargado
            self.status_label.config(text=f"✅ Excel cargado: {len(self.excel_data)} registros", 
                                   foreground='green')
            
            self.validate_form()
            
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo cargar el archivo Excel: {str(e)}")
            self.excel_data = None
            self.validate_form()
    
    def validate_email_format(self, email):
        """Valida el formato básico de un email"""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None
    
    def is_outlook_email(self, email):
        """Verifica si el email es de Outlook/Office365 o dominios personalizados"""
        outlook_domains = [
            'outlook.com', 'hotmail.com', 'live.com', 
            'office365.com', 'microsoft.com',
            # Agregar dominios personalizados que usen Office365
            'unison.mx'  # ← DOMINIO AGREGADO
        ]
        domain = email.split('@')[-1].lower() if '@' in email else ''
        return any(domain.endswith(outlook_domain) for outlook_domain in outlook_domains)
    
    def validate_form(self):
        """Valida que el formulario esté completo para habilitar el envío"""
        try:
            email = self.email_entry.get().strip()
            
            required_fields = [
                email,
                self.password_entry.get().strip(),
                self.pdf_folder_entry.get().strip(),
                self.excel_file_entry.get().strip(),
                self.name_column_combo.get(),
                self.email_column_combo.get(),
                self.filename_column_combo.get(),
                self.subject_entry.get().strip(),
                self.body_text.get('1.0', 'end-1c').strip()
            ]
            
            # Verificar que todos los campos requeridos estén llenos
            if not all(required_fields):
                self.send_btn.config(state='disabled')
                return False
            
            # Verificar que sea un email válido
            if email and not self.validate_email_format(email):
                self.status_label.config(text="⚠️ Formato de email inválido", 
                                       foreground='orange')
                self.send_btn.config(state='disabled')
                return False
            
            # Verificar que el Excel esté cargado y tenga datos
            if self.excel_data is None or len(self.excel_data) == 0:
                self.send_btn.config(state='disabled')
                return False
            
            # Verificar que las columnas seleccionadas existan
            selected_columns = [
                self.name_column_combo.get(),
                self.email_column_combo.get(),
                self.filename_column_combo.get()
            ]
            
            for col in selected_columns:
                if col not in self.excel_data.columns:
                    self.send_btn.config(state='disabled')
                    return False
            
            # Si todo está bien, habilitar el botón
            self.send_btn.config(state='normal')
            self.status_label.config(text="✅ Configuración lista", foreground='green')
            return True
            
        except Exception as e:
            self.send_btn.config(state='disabled')
            return False
    
    def test_connection(self):
        """Prueba la conexión con el servidor SMTP"""
        email = self.email_entry.get().strip()
        password = self.password_entry.get().strip()
        
        if not email or not password:
            messagebox.showwarning("Advertencia", "Por favor ingresa el correo y contraseña")
            return
        
        # Validar formato de email
        if not self.validate_email_format(email):
            messagebox.showwarning("Advertencia", "Por favor ingresa un formato de email válido")
            return
        
        self.test_btn.config(state='disabled', text="Probando conexión...")
        self.status_label.config(text="Probando conexión...", foreground='blue')
        self.update_log("🔗 Iniciando prueba de conexión...")
        
        # Ejecutar en hilo separado para no bloquear la interfaz
        def test_thread():
            try:
                # Crear instancia temporal para la prueba
                temp_sender = EmailSender({}, pd.DataFrame(), "")
                success, message = temp_sender.test_connection(email, password)
                self.window.after(0, self.connection_test_result, success, message)
            except Exception as e:
                self.window.after(0, self.connection_test_result, False, f"Error inesperado: {str(e)}")
        
        threading.Thread(target=test_thread, daemon=True).start()
    
    def connection_test_result(self, success, message):
        """Muestra el resultado de la prueba de conexión"""
        self.test_btn.config(state='normal', text="Probar Conexión")
        if success:
            self.status_label.config(text="✓ Conexión exitosa", foreground='green')
            self.update_log("✅ Conexión SMTP exitosa")
            messagebox.showinfo("Éxito", message)
        else:
            self.status_label.config(text="✗ Error de conexión", foreground='red')
            self.update_log("❌ Error en conexión SMTP")
            
            # Mensaje específico para problemas comunes
            enhanced_message = message + "\n\nPara cuentas institucionales (@unison.mx) verifica:\n" \
                                       "• Tu contraseña es correcta\n" \
                                       "• La cuenta está activa\n" \
                                       "• No hay bloqueos de seguridad"
            
            messagebox.showerror("Error", enhanced_message)
    
    def start_sending_emails(self):
        """Inicia el proceso de envío de correos en un hilo separado"""
        if not self.validate_selections():
            return
        
        # Confirmar envío
        confirm = messagebox.askyesno(
            "Confirmar envío",
            f"¿Estás seguro de que deseas enviar {len(self.excel_data)} correos?\n\n"
            "Asegúrate de:\n"
            "1. Tener conexión a internet estable\n"
            "2. Usar contraseña correcta\n"
            "3. Revisar el contenido del mensaje"
        )
        
        if not confirm:
            return
        
        # Cambiar botón a "Cancelar" durante el envío
        self.send_btn.config(text="Cancelar Envío", command=self.stop_sending)
        self.progress.start()
        self.status_label.config(text="Iniciando envío de correos...", foreground='blue')
        self.update_log("🚀 Iniciando proceso de envío de correos...")
        
        # Limpiar log anterior
        self.log_text.delete('1.0', tk.END)
        
        # Ejecutar envío en hilo separado
        threading.Thread(target=self.send_emails_thread, daemon=True).start()
    
    def validate_selections(self):
        """Valida las selecciones antes del envío"""
        # Verificar que las columnas seleccionadas existan en los datos
        selected_columns = [
            self.name_column_combo.get(),
            self.email_column_combo.get(),
            self.filename_column_combo.get()
        ]
        
        for col in selected_columns:
            if col not in self.excel_data.columns:
                messagebox.showerror("Error", f"La columna '{col}' no existe en el archivo Excel")
                return False
        
        # Verificar que la carpeta de PDFs existe
        if not os.path.exists(self.pdf_folder_entry.get()):
            messagebox.showerror("Error", "La carpeta de PDFs no existe")
            return False
        
        return True
    
    def send_emails_thread(self):
        """Hilo para el envío de correos (ejecución en segundo plano)"""
        try:
            # Preparar configuración
            config = {
                'email': self.email_entry.get().strip(),
                'password': self.password_entry.get().strip(),
                'sender_name': self.sender_name_entry.get().strip(),
                'subject': self.subject_entry.get().strip(),
                'body': self.body_text.get('1.0', 'end-1c').strip(),
                'name_column': self.name_column_combo.get(),
                'email_column': self.email_column_combo.get(),
                'filename_column': self.filename_column_combo.get()
            }
            
            # Filtrar datos para solo incluir filas con información completa
            filtered_data = self.excel_data.dropna(subset=[
                config['name_column'],
                config['email_column'],
                config['filename_column']
            ])
            
            self.update_log(f"📊 Procesando {len(filtered_data)} registros válidos")
            self.update_log(f"📧 Enviando desde: {config['email']}")
            
            # Crear instancia de EmailSender con los parámetros correctos
            self.email_sender = EmailSender(config, filtered_data, self.pdf_folder_entry.get())
            
            # Conectar señales para actualizar la interfaz
            self.email_sender.progress.connect(self.update_progress)
            self.email_sender.log.connect(self.update_log)
            self.email_sender.finished.connect(self.sending_complete)
            
            # Iniciar el envío
            self.email_sender.start()
            
        except Exception as e:
            self.window.after(0, self.sending_error, str(e))
    
    def update_progress(self, value):
        """Actualiza la barra de progreso"""
        # Cambiar a modo determinada para mostrar progreso real
        self.progress.config(mode='determinate')
        self.progress['value'] = value
        self.window.update_idletasks()
    
    def update_log(self, message):
        """Actualiza el log en la interfaz"""
        self.status_label.config(text=message)
        
        # Agregar al log detallado con timestamp
        import datetime
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)  # Auto-scroll al final
        
        # Forzar actualización de la interfaz
        self.window.update_idletasks()
    
    def sending_complete(self, results):
        """Muestra los resultados del envío"""
        self.progress.stop()
        self.progress.config(mode='indeterminate')
        
        # Restaurar el botón a su estado original
        self.send_btn.config(text="Enviar Correos", command=self.start_sending_emails, state='normal')
        
        # Manejar el resultado
        if results.startswith("error:"):
            # Es un error
            error_message = results.replace("error:", "").strip()
            self.status_label.config(text="✗ Error en el envío", foreground='red')
            self.update_log(f"❌ Error final: {error_message}")
            messagebox.showerror("Error", f"Ocurrió un error durante el envío:\n{error_message}")
        else:
            # Es un resultado exitoso
            self.status_label.config(text=results, foreground='green')
            self.update_log(f"✅ Proceso finalizado: {results}")
            
            # Mostrar mensaje de éxito
            if "completado" in results.lower() or "exitosa" in results.lower():
                messagebox.showinfo("Éxito", f"Envío completado\n\n{results}")
            elif "cancelado" in results.lower():
                messagebox.showinfo("Envío cancelado", results)
            else:
                messagebox.showinfo("Proceso completado", results)
    
    def sending_error(self, error_message):
        """Muestra error en el envío"""
        self.progress.stop()
        self.send_btn.config(text="Enviar Correos", command=self.start_sending_emails, state='normal')
        self.status_label.config(text="✗ Error en el envío", foreground='red')
        self.update_log(f"❌ Error crítico: {error_message}")
        messagebox.showerror("Error", f"Ocurrió un error durante el envío:\n{error_message}")
    
    def stop_sending(self):
        """Detiene el envío de correos"""
        if self.email_sender and self.email_sender.isRunning():
            self.email_sender.stop()
            self.status_label.config(text="⏹️ Cancelando envío...", foreground='orange')
            self.update_log("⏹️ Solicitando cancelación del envío...")