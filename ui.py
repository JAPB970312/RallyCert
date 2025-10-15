# ui.py
import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QSpinBox, QCheckBox, QProgressBar, QMessageBox, QTextEdit, QComboBox, QFormLayout,
    QTabWidget, QSlider, QInputDialog, QScrollArea, QDialog, QLineEdit, QToolBar, QTextBrowser,
    QSizePolicy, QGridLayout, QFrame
)
from PyQt6.QtGui import QPixmap, QImage, QIcon, QTextCursor, QTextCharFormat, QTextBlockFormat, QTextFormat, QFont
from PyQt6.QtCore import Qt, QRegularExpression, QSize, pyqtSignal
from datetime import datetime

# Importaciones de nuestros módulos
from resource_manager import resource_path
from data_handler import get_excel_data
from worker import Worker
from document_processor import get_processor, PdfProcessor

# Importaciones de las nuevas mejoras
from validator import DocumentValidator
from template_library import TemplateLibrary, TemplateCategory
from performance_optimizer import PerformanceOptimizer

class ModernButton(QPushButton):
    """Botón moderno con efectos hover"""
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setMinimumHeight(35)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        # Estilo claro para botones
        self.setStyleSheet("""
            ModernButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-weight: bold;
                font-size: 14px;
            }
            ModernButton:hover {
                background-color: #357abd;
            }
            ModernButton:pressed {
                background-color: #2d6da3;
            }
            ModernButton:disabled {
                background-color: #cccccc;
                color: #666666;
            }
        """)

class ModernLineEdit(QLineEdit):
    """Campo de texto moderno"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(35)
        self.setStyleSheet("""
            ModernLineEdit {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
                selection-background-color: #4a90e2;
            }
            ModernLineEdit:focus {
                border-color: #4a90e2;
                background-color: #f8f9fa;
            }
        """)

class ModernComboBox(QComboBox):
    """Combo box moderno"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(35)
        self.setStyleSheet("""
            ModernComboBox {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            ModernComboBox:focus {
                border-color: #4a90e2;
            }
            ModernComboBox::drop-down {
                border: none;
            }
            ModernComboBox QAbstractItemView {
                background-color: white;
                border: 1px solid #e1e5e9;
                selection-background-color: #4a90e2;
                selection-color: white;
            }
        """)

class ModernLabel(QLabel):
    """Etiqueta moderna"""
    def __init__(self, text="", parent=None):
        super().__init__(text, parent)
        self.setWordWrap(True)
        self.setStyleSheet("""
            ModernLabel {
                background-color: transparent;
                color: #333333;
                font-size: 14px;
                padding: 2px;
            }
        """)

class SectionWidget(QFrame):
    """Widget de sección con título y contenido"""
    def __init__(self, title, parent=None):
        super().__init__(parent)
        self.setFrameStyle(QFrame.Shape.NoFrame)
        self.setStyleSheet("""
            SectionWidget {
                background-color: #ffffff;
                border-radius: 8px;
                padding: 12px;
                margin: 5px;
                border: 1px solid #e1e5e9;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)
        
        title_label = ModernLabel(title)
        title_label.setStyleSheet("""
            font-weight: bold;
            font-size: 13px;
            color: #2c3e50;
            padding-bottom: 8px;
            border-bottom: 1px solid #dee2e6;
            margin-bottom: 8px;
        """)
        layout.addWidget(title_label)
        
        self.content_widget = QWidget()
        self.content_layout = QVBoxLayout(self.content_widget)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.content_widget)
    
    def addWidget(self, widget):
        self.content_layout.addWidget(widget)
    
    def addLayout(self, layout):
        self.content_layout.addLayout(layout)

class EmailSenderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📤 Enviar Constancias por Correo")
        self.setMinimumSize(800, 700)
        
        # Agregar botones de minimizar y maximizar
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowMinMaxButtonsHint)
        
        self.email_sender = None
        self.setup_ui()

    def setup_ui(self):
        # FONDO CLARO MODERNO
        self.setStyleSheet("""
            QDialog {
                background-color: #f8f9fa;
                color: #333333;
            }
            QLabel {
                background-color: transparent;
                color: #333333;
            }
            QLineEdit, QComboBox, QSpinBox {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border-color: #4a90e2;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2d6da3;
            }
            QTextEdit {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            QProgressBar {
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                background-color: white;
                text-align: center;
                color: #333333;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #4a90e2;
                border-radius: 4px;
            }
            QCheckBox {
                spacing: 8px;
                font-size: 14px;
                color: #333333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #6c757d;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #4a90e2;
                border-color: #4a90e2;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # Crear scroll area para contenido
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        content_widget = QWidget()
        content_layout = QVBoxLayout(content_widget)
        content_layout.setSpacing(15)

        # Configuración de correo
        email_section = SectionWidget("📧 Configuración de Correo")
        email_form = QFormLayout()
        email_form.setVerticalSpacing(10)
        
        self.email_entry = ModernLineEdit()
        self.email_entry.setPlaceholderText("ejemplo@gmail.com")
        self.email_entry.textChanged.connect(self.validate_form)
        email_form.addRow("Correo electrónico:", self.email_entry)

        self.password_entry = ModernLineEdit()
        self.password_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_entry.setPlaceholderText("Contraseña o contraseña de aplicación")
        self.password_entry.textChanged.connect(self.validate_form)
        email_form.addRow("Contraseña:", self.password_entry)

        self.show_password = QCheckBox("Mostrar contraseña")
        self.show_password.toggled.connect(self.toggle_password_visibility)
        email_form.addRow("", self.show_password)

        self.sender_name_entry = ModernLineEdit()
        self.sender_name_entry.setText("Generador de Constancias")
        self.sender_name_entry.setPlaceholderText("Nombre del remitente")
        self.sender_name_entry.textChanged.connect(self.validate_form)
        email_form.addRow("Nombre del remitente:", self.sender_name_entry)
        
        email_section.addLayout(email_form)
        content_layout.addWidget(email_section)

        # Selección de archivos
        files_section = SectionWidget("📁 Selección de Archivos")
        files_form = QFormLayout()
        files_form.setVerticalSpacing(10)

        self.pdf_folder_entry = ModernLineEdit()
        self.pdf_folder_entry.setPlaceholderText("Carpeta con los PDFs")
        self.btn_select_pdf = ModernButton("Seleccionar Carpeta")
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_folder_entry)
        pdf_layout.addWidget(self.btn_select_pdf)
        files_form.addRow("Carpeta de PDFs:", pdf_layout)

        self.excel_file_entry = ModernLineEdit()
        self.excel_file_entry.setPlaceholderText("Archivo Excel con datos")
        self.btn_select_excel = ModernButton("Seleccionar Excel")
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_file_entry)
        excel_layout.addWidget(self.btn_select_excel)
        files_form.addRow("Archivo Excel:", excel_layout)
        
        files_section.addLayout(files_form)
        content_layout.addWidget(files_section)

        # Configuración de columnas
        columns_section = SectionWidget("🔗 Configuración de Columnas")
        columns_form = QFormLayout()
        columns_form.setVerticalSpacing(10)

        self.name_column_combo = ModernComboBox()
        self.name_column_combo.currentTextChanged.connect(self.validate_form)
        columns_form.addRow("Columna Nombre:", self.name_column_combo)

        self.email_column_combo = ModernComboBox()
        self.email_column_combo.currentTextChanged.connect(self.validate_form)
        columns_form.addRow("Columna Correo:", self.email_column_combo)

        self.filename_column_combo = ModernComboBox()
        self.filename_column_combo.currentTextChanged.connect(self.validate_form)
        columns_form.addRow("Columna Archivo PDF:", self.filename_column_combo)
        
        columns_section.addLayout(columns_form)
        content_layout.addWidget(columns_section)

        # Contenido del correo
        content_section = SectionWidget("📝 Contenido del Correo")
        
        # Asunto
        subject_layout = QHBoxLayout()
        subject_label = ModernLabel("Asunto:")
        self.subject_entry = ModernLineEdit()
        self.subject_entry.setText("Constancia de Participación")
        self.subject_entry.textChanged.connect(self.validate_form)
        subject_layout.addWidget(subject_label)
        subject_layout.addWidget(self.subject_entry)
        content_section.addLayout(subject_layout)

        # Cuerpo del mensaje
        body_label = ModernLabel("Cuerpo del mensaje:")
        content_section.addWidget(body_label)

        self.body_text = QTextEdit()
        self.body_text.setMinimumHeight(200)
        
        # TEXTO PREDETERMINADO MEJORADO
        default_body = """<div style="font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.4; color: #000000; background-color: #ffffff;">
    <p style="color: #000000; margin: 12px 0;">Estimado/a <strong style="color: #000000;">{nombre}</strong>,</p>
    
    <p style="color: #000000; text-align: justify; margin: 12px 0 12px 20px;">
        Le hacemos llegar su <strong style="color: #000000;">constancia de participación</strong> emitida por la 
        <strong style="color: #000000;">Universidad de Sonora</strong>. Este documento certifica su asistencia y 
        participación en nuestro evento académico.
    </p>
    
    <p style="color: #000000; text-align: justify; margin: 12px 0 12px 20px;">
        <strong style="color: #000000;">
            Agradecemos su valiosa contribución y esperamos contar con su participación en futuras actividades.
        </strong>
    </p>
    
    <p style="color: #000000; text-align: justify; margin: 12px 0;">
        Saludos cordiales,<br>
        <strong style="color: #000000;">Departamento de Constancias</strong><br>
        Universidad de Sonora
    </p>
</div>"""
        
        self.body_text.setHtml(default_body)
        self.body_text.textChanged.connect(self.validate_form)

        content_section.addWidget(self.body_text)

        content_layout.addWidget(content_section)
        content_layout.addStretch()

        scroll_area.setWidget(content_widget)
        main_layout.addWidget(scroll_area)

        # Botones de acción
        buttons_layout = QHBoxLayout()
        self.btn_test = ModernButton("🔍 Probar Conexión")
        self.btn_send = ModernButton("📤 Enviar Correos")
        self.btn_cancel = ModernButton("❌ Cancelar")

        self.btn_test.setStyleSheet("background-color: #17a2b8;")
        self.btn_send.setStyleSheet("background-color: #28a745;")
        self.btn_cancel.setStyleSheet("background-color: #dc3545;")

        buttons_layout.addWidget(self.btn_test)
        buttons_layout.addWidget(self.btn_send)
        buttons_layout.addWidget(self.btn_cancel)
        main_layout.addLayout(buttons_layout)

        # Barra de progreso y estado
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)

        self.status_label = ModernLabel("Listo para configurar el envío")
        self.status_label.setStyleSheet("""
            padding: 12px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #495057;
            font-size: 14px;
        """)
        main_layout.addWidget(self.status_label)

        # Conectar señales
        self.btn_select_pdf.clicked.connect(self.select_pdf_folder)
        self.btn_select_excel.clicked.connect(self.select_excel_file)
        self.btn_test.clicked.connect(self.test_connection)
        self.btn_send.clicked.connect(self.start_sending)
        self.btn_cancel.clicked.connect(self.cancel_operation)

        self.btn_send.setEnabled(False)

    def toggle_password_visibility(self, checked):
        if checked:
            self.password_entry.setEchoMode(QLineEdit.EchoMode.Normal)
        else:
            self.password_entry.setEchoMode(QLineEdit.EchoMode.Password)

    def select_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Seleccionar carpeta de PDFs")
        if folder:
            self.pdf_folder_entry.setText(folder)
            self.validate_form()

    def select_excel_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar archivo Excel", "", "Excel (*.xlsx *.xls)"
        )
        if file_path:
            self.excel_file_entry.setText(file_path)
            self.load_excel_columns(file_path)
            self.validate_form()

    def load_excel_columns(self, excel_path):
        try:
            from data_handler import get_excel_data
            columns, data = get_excel_data(excel_path)
            
            self.name_column_combo.clear()
            self.email_column_combo.clear()
            self.filename_column_combo.clear()
            
            self.name_column_combo.addItems(columns)
            self.email_column_combo.addItems(columns)
            self.filename_column_combo.addItems(columns)
            
            # Seleccionar automáticamente columnas comunes
            common_names = ['nombre', 'name', 'participante', 'estudiante', 'alumno', 'nom']
            common_emails = ['email', 'correo', 'mail', 'e-mail', 'correo_electronico']
            common_files = ['archivo', 'filename', 'pdf', 'constancia', 'certificado', 'documento']
            
            for i, col in enumerate(columns):
                col_lower = col.lower()
                if any(common in col_lower for common in common_names) and self.name_column_combo.currentIndex() == -1:
                    self.name_column_combo.setCurrentIndex(i)
                if any(common in col_lower for common in common_emails) and self.email_column_combo.currentIndex() == -1:
                    self.email_column_combo.setCurrentIndex(i)
                if any(common in col_lower for common in common_files) and self.filename_column_combo.currentIndex() == -1:
                    self.filename_column_combo.setCurrentIndex(i)
            
            self.status_label.setText("✅ Excel cargado correctamente")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #d4edda;
                border: 1px solid #c3e6cb;
                border-radius: 6px;
                color: #155724;
            """)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar el archivo Excel: {str(e)}")
            self.status_label.setText("❌ Error al cargar Excel")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #f8d7da;
                border: 1px solid #f5c6cb;
                border-radius: 6px;
                color: #721c24;
            """)

    def validate_form(self):
        """Valida que todos los campos requeridos estén completos"""
        required_fields = [
            self.email_entry.text().strip(),
            self.password_entry.text().strip(),
            self.pdf_folder_entry.text().strip(),
            self.excel_file_entry.text().strip(),
            self.name_column_combo.currentText(),
            self.email_column_combo.currentText(),
            self.filename_column_combo.currentText(),
            self.subject_entry.text().strip(),
            self.body_text.toPlainText().strip()
        ]
        
        is_complete = all(required_fields)
        self.btn_send.setEnabled(is_complete)
        
        if is_complete:
            self.status_label.setText("✅ Formulario completo. Listo para enviar.")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #d4edda;
                border: 1px solid #c3e6cb;
                border-radius: 6px;
                color: #155724;
            """)
        else:
            self.status_label.setText("ℹ️ Complete todos los campos para habilitar el envío")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #fff3cd;
                border: 1px solid #ffeaa7;
                border-radius: 6px;
                color: #856404;
            """)

    def test_connection(self):
        email = self.email_entry.text().strip()
        password = self.password_entry.text().strip()
        
        if not email or not password:
            QMessageBox.warning(self, "Campos requeridos", "Por favor ingrese el correo y contraseña.")
            return
        
        self.status_label.setText("🔄 Probando conexión con el servidor SMTP...")
        self.status_label.setStyleSheet("""
            padding: 12px;
            background-color: #fff3cd;
            border: 1px solid #ffeaa7;
            border-radius: 6px;
            color: #856404;
        """)
        self.btn_test.setEnabled(False)
        
        # Importar y probar conexión real
        try:
            from email_sender import EmailSender
            self.email_sender = EmailSender({}, [], "")
            success, message = self.email_sender.test_connection(email, password)
            
            if success:
                self.status_label.setText(message)
                self.status_label.setStyleSheet("""
                    padding: 12px;
                    background-color: #d4edda;
                    border: 1px solid #c3e6cb;
                    border-radius: 6px;
                    color: #155724;
                """)
                QMessageBox.information(self, "✅ Conexión Exitosa", 
                                      f"Conexión establecida correctamente con:\n\n"
                                      f"Correo: {email}\n"
                                      f"Servidor: SMTP\n\n"
                                      f"Ahora puede proceder con el envío de constancias.")
            else:
                self.status_label.setText(message)
                self.status_label.setStyleSheet("""
                    padding: 12px;
                    background-color: #f8d7da;
                    border: 1px solid #f5c6cb;
                    border-radius: 6px;
                    color: #721c24;
                """)
                QMessageBox.critical(self, "❌ Error de Conexión", 
                                   f"No se pudo establecer la conexión:\n\n{message}\n\n"
                                   f"Para Gmail, asegúrese de:\n"
                                   f"1. Activar verificación en 2 pasos\n"
                                   f"2. Usar una contraseña de aplicación\n"
                                   f"3. No usar contraseña principal")
                
        except Exception as e:
            error_msg = f"Error inesperado: {str(e)}"
            self.status_label.setText(error_msg)
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #f8d7da;
                border: 1px solid #f5c6cb;
                border-radius: 6px;
                color: #721c24;
            """)
            QMessageBox.critical(self, "Error", error_msg)
        
        finally:
            self.btn_test.setEnabled(True)

    def start_sending(self):
        # Validar campos requeridos
        if not self.validate_sending():
            return
        
        # Confirmar envío
        total_records = self.get_total_records()
        if total_records == 0:
            QMessageBox.warning(self, "Sin datos", "No se encontraron registros en el Excel.")
            return
        
        confirm_msg = f"""
¿Está seguro de que desea enviar los correos?

📧 Correo remitente: {self.email_entry.text()}
📁 Carpeta PDFs: {os.path.basename(self.pdf_folder_entry.text())}
📊 Archivo Excel: {os.path.basename(self.excel_file_entry.text())} ({total_records} registros)
📝 Columnas mapeadas:
   • Nombre: {self.name_column_combo.currentText()}
   • Correo: {self.email_column_combo.currentText()}
   • Archivo PDF: {self.filename_column_combo.currentText()}

¿Continuar con el envío?
        """
        
        reply = QMessageBox.question(self, "Confirmar envío", confirm_msg, 
                                   QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        
        if reply == QMessageBox.StandardButton.Yes:
            self.execute_sending()

    def validate_sending(self):
        """Valida todos los campos antes del envío"""
        if not self.email_entry.text().strip():
            QMessageBox.warning(self, "Campo requerido", "Ingrese el correo electrónico del remitente.")
            return False
        
        if not self.password_entry.text().strip():
            QMessageBox.warning(self, "Campo requerido", "Ingrese la contraseña del correo.")
            return False
        
        if not self.pdf_folder_entry.text().strip():
            QMessageBox.warning(self, "Campo requerido", "Seleccione la carpeta de PDFs.")
            return False
        
        if not os.path.exists(self.pdf_folder_entry.text()):
            QMessageBox.warning(self, "Carpeta no existe", "La carpeta de PDFs no existe.")
            return False
        
        if not self.excel_file_entry.text().strip():
            QMessageBox.warning(self, "Campo requerido", "Seleccione el archivo Excel.")
            return False
        
        if not os.path.exists(self.excel_file_entry.text()):
            QMessageBox.warning(self, "Archivo no existe", "El archivo Excel no existe.")
            return False
        
        if self.name_column_combo.currentText() == "":
            QMessageBox.warning(self, "Campo requerido", "Seleccione la columna para el nombre.")
            return False
        
        if self.email_column_combo.currentText() == "":
            QMessageBox.warning(self, "Campo requerido", "Seleccione la columna para el correo.")
            return False
        
        if self.filename_column_combo.currentText() == "":
            QMessageBox.warning(self, "Campo requerido", "Seleccione la columna para el nombre del archivo PDF.")
            return False
        
        return True

    def get_total_records(self):
        """Obtiene el número total de registros en el Excel"""
        try:
            from data_handler import get_excel_data
            _, data = get_excel_data(self.excel_file_entry.text())
            return len(data)
        except:
            return 0

    def execute_sending(self):
        """Ejecuta el envío real de correos"""
        try:
            from data_handler import get_excel_data
            from email_sender import EmailSender
            
            # Cargar datos del Excel
            columns, excel_data = get_excel_data(self.excel_file_entry.text())
            
            # Configuración para el envío
            config = {
                'email': self.email_entry.text().strip(),
                'password': self.password_entry.text().strip(),
                'sender_name': self.sender_name_entry.text().strip(),
                'subject': self.subject_entry.text().strip(),
                'body': self.body_text.toHtml() if self.body_text.toHtml().strip() else self.body_text.toPlainText().strip(),
                'name_column': self.name_column_combo.currentText(),
                'email_column': self.email_column_combo.currentText(),
                'filename_column': self.filename_column_combo.currentText()
            }
            
            # Crear y configurar el enviador de correos
            self.email_sender = EmailSender(config, excel_data, self.pdf_folder_entry.text())
            
            # Conectar señales
            self.email_sender.progress.connect(self.progress_bar.setValue)
            self.email_sender.log.connect(self.update_status)
            self.email_sender.finished.connect(self.on_sending_finished)
            
            # Configurar interfaz
            self.btn_send.setEnabled(False)
            self.btn_test.setEnabled(False)
            self.btn_cancel.setEnabled(True)
            self.progress_bar.setVisible(True)
            self.progress_bar.setValue(0)
            
            # Iniciar envío
            self.email_sender.start()
            self.status_label.setText("🚀 Iniciando envío de correos...")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #cce7ff;
                border: 1px solid #b3d9ff;
                border-radius: 6px;
                color: #004085;
            """)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error al iniciar el envío: {str(e)}")
            self.reset_interface()

    def update_status(self, message):
        """Actualiza el estado con mensajes del proceso de envío"""
        self.status_label.setText(message)

    def on_sending_finished(self, message):
        """Maneja la finalización del envío"""
        if message.startswith("error:"):
            QMessageBox.critical(self, "Error", message[6:])
            self.status_label.setText("❌ Error en el envío")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #f8d7da;
                border: 1px solid #f5c6cb;
                border-radius: 6px;
                color: #721c24;
            """)
        else:
            QMessageBox.information(self, "Proceso Completado", message)
            self.status_label.setText("✅ " + message)
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #d4edda;
                border: 1px solid #c3e6cb;
                border-radius: 6px;
                color: #155724;
            """)
        
        self.reset_interface()

    def cancel_operation(self):
        """Cancela la operación en curso"""
        if self.email_sender and self.email_sender.isRunning():
            self.email_sender.stop()
            self.status_label.setText("⏹️ Cancelando envío...")
            self.status_label.setStyleSheet("""
                padding: 12px;
                background-color: #fff3cd;
                border: 1px solid #ffeaa7;
                border-radius: 6px;
                color: #856404;
            """)
        else:
            self.reject()

    def reset_interface(self):
        """Restablece la interfaz a su estado inicial"""
        self.btn_send.setEnabled(True)
        self.btn_test.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress_bar.setVisible(False)
        self.progress_bar.setValue(0)

class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RallyCert - Generador de Constancias")
        self.setGeometry(100, 100, 1400, 900)
        self.setMinimumSize(1200, 700)
        self.setWindowIcon(QIcon(resource_path('assets/icon.ico')))

        # ESTILOS MODERNOS CON FONDO CLARO
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f8f9fa;
                color: #333333;
            }
            QWidget {
                background-color: transparent;
                color: #333333;
            }
            QPushButton {
                background-color: #4a90e2;
                color: white;
                border: none;
                border-radius: 6px;
                padding: 10px 16px;
                font-weight: bold;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #357abd;
            }
            QPushButton:pressed {
                background-color: #2d6da3;
            }
            QLineEdit, QComboBox, QSpinBox {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border-color: #4a90e2;
            }
            QTextEdit {
                background-color: white;
                color: #333333;
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                padding: 8px;
                font-size: 14px;
            }
            QTabWidget::pane {
                border: 2px solid #e1e5e9;
                border-radius: 8px;
                background-color: white;
            }
            QTabBar::tab {
                background-color: #e9ecef;
                color: #6c757d;
                padding: 12px 20px;
                margin-right: 2px;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                font-weight: bold;
            }
            QTabBar::tab:selected {
                background-color: white;
                color: #4a90e2;
                border-bottom: 2px solid #4a90e2;
            }
            QProgressBar {
                border: 2px solid #e1e5e9;
                border-radius: 6px;
                background-color: white;
                text-align: center;
                color: #333333;
                height: 20px;
            }
            QProgressBar::chunk {
                background-color: #4a90e2;
                border-radius: 4px;
            }
            QScrollArea {
                border: none;
                background-color: transparent;
            }
            QCheckBox {
                spacing: 8px;
                font-size: 14px;
                color: #333333;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border-radius: 3px;
                border: 2px solid #6c757d;
                background-color: white;
            }
            QCheckBox::indicator:checked {
                background-color: #4a90e2;
                border-color: #4a90e2;
            }
        """)

        # Estado de la aplicación
        self.template_path = ""
        self.excel_data = []
        self.excel_columns = []

        # Inicializar sistemas mejorados
        self.validator = DocumentValidator()
        self.template_library = TemplateLibrary()
        self.performance_optimizer = PerformanceOptimizer()

        self.setup_ui()

    def setup_ui(self):
        """Configura la interfaz de usuario moderna"""
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # LAYOUT PRINCIPAL
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(0)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # --- BANNER SUPERIOR ---
        banner_container = QWidget()
        banner_container.setStyleSheet("background-color: white;")
        banner_container.setFixedHeight(120)
        banner_layout = QVBoxLayout(banner_container)
        banner_layout.setContentsMargins(20, 10, 20, 10)
        
        self.banner_label = QLabel()
        banner_pixmap = QPixmap(resource_path('assets/Banner.png')) 
        if not banner_pixmap.isNull():
            banner_pixmap = banner_pixmap.scaledToHeight(80, Qt.TransformationMode.SmoothTransformation)
            self.banner_label.setPixmap(banner_pixmap)
        else:
            self.banner_label.setText("RallyCert - Generador de Constancias")
            self.banner_label.setStyleSheet("""
                color: #f7ead8; 
                font-weight: bold; 
                font-size: 24px;
                background-color: transparent;
            """)
        
        self.banner_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        banner_layout.addWidget(self.banner_label)
        main_layout.addWidget(banner_container)
        
        # --- CONTENIDO PRINCIPAL ---
        content_container = QWidget()
        content_layout = QHBoxLayout(content_container)
        content_layout.setSpacing(15)
        content_layout.setContentsMargins(15, 15, 15, 15)
        
        # --- PANEL DE CONTROL IZQUIERDO (40%) ---
        control_container = QWidget()
        control_container.setMaximumWidth(500)
        control_layout = QVBoxLayout(control_container)
        control_layout.setSpacing(10)
        
        # Scroll area para controles
        control_scroll = QScrollArea()
        control_scroll.setWidgetResizable(True)
        control_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        control_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        control_content = QWidget()
        control_scroll_layout = QVBoxLayout(control_content)
        control_scroll_layout.setSpacing(15)
        control_scroll_layout.setContentsMargins(5, 5, 5, 5)
        
        # Secciones del panel de control
        sections = [
            ("🎨 Plantillas Predefinidas", self.create_template_section()),
            ("📄 Cargar Plantilla", self.create_template_load_section()),
            ("📊 Cargar Participantes", self.create_excel_section()),
            ("🔗 Asignar Columnas", self.create_mapping_section()),
            ("🎯 Estilo de Texto", self.create_style_section()),
            ("🔐 Firma Digital", self.create_signature_section()),
            ("🏷️ Leyenda de Validación", self.create_validation_section()),
            ("✅ Validación", self.create_validation_check_section()),
            ("⚙️ Configuración", self.create_config_section()),
            ("🚀 Acciones", self.create_actions_section())
        ]
        
        for title, widget in sections:
            section = SectionWidget(title)
            section.addWidget(widget)
            control_scroll_layout.addWidget(section)
        
        control_scroll_layout.addStretch()
        control_scroll.setWidget(control_content)
        control_layout.addWidget(control_scroll)
        
        # --- PANEL DE PREVISUALIZACIÓN DERECHO (60%) ---
        preview_container = QWidget()
        preview_layout = QVBoxLayout(preview_container)
        preview_layout.setSpacing(10)
        
        # Tabs para previsualización y logs
        preview_tabs = QTabWidget()
        preview_tabs.setStyleSheet("""
            QTabWidget::pane {
                border: 2px solid #e1e5e9;
                border-radius: 8px;
                background-color: white;
            }
        """)
        
        # Pestaña de Previsualización (más pequeña)
        preview_widget = QWidget()
        preview_widget_layout = QVBoxLayout(preview_widget)
        preview_widget_layout.setContentsMargins(10, 10, 10, 10)
        
        preview_title = ModernLabel("👁️ Previsualización en Tiempo Real")
        preview_title.setStyleSheet("font-weight: bold; font-size: 14px; color: #2c3e50;")
        preview_widget_layout.addWidget(preview_title)
        
        self.preview_label = QLabel("Cargue una plantilla PDF para ver la previsualización.")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setStyleSheet("""
            border: 2px dashed #dee2e6; 
            background-color: #f8f9fa; 
            border-radius: 8px; 
            padding: 20px; 
            font-size: 14px; 
            color: #6c757d;
            min-height: 300px;
        """)
        self.preview_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        preview_widget_layout.addWidget(self.preview_label)
        
        # Pestaña de Log
        log_widget = QWidget()
        log_layout = QVBoxLayout(log_widget)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        log_title = ModernLabel("📝 Registro de Actividad")
        log_title.setStyleSheet("font-weight: bold; font-size: 14px; color: #2c3e50;")
        log_layout.addWidget(log_title)
        
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("""
            font-family: 'Consolas', 'Monaco', monospace; 
            font-size: 12px; 
            color: #2c3e50; 
            background-color: white;
            border: 1px solid #e1e5e9;
            border-radius: 6px;
        """)
        log_layout.addWidget(self.log_box)
        
        preview_tabs.addTab(preview_widget, "👁️ Previsualización")
        preview_tabs.addTab(log_widget, "📝 Registro")
        
        preview_layout.addWidget(preview_tabs)
        
        # Añadir paneles al layout principal
        content_layout.addWidget(control_container, 1)  # 40% del espacio
        content_layout.addWidget(preview_container, 2)  # 60% del espacio
        
        main_layout.addWidget(content_container, 1)

    def create_template_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        
        self.template_preset_combo = ModernComboBox()
        self.load_template_presets()
        self.template_preset_combo.currentTextChanged.connect(self.apply_template_preset)
        layout.addWidget(self.template_preset_combo)
        
        return widget

    def create_template_load_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        self.btn_load_template = ModernButton("Seleccionar Plantilla (.pdf, .docx, .pptx)")
        self.btn_load_template.clicked.connect(self.load_template)
        layout.addWidget(self.btn_load_template)
        
        self.lbl_template_path = ModernLabel("Ningún archivo seleccionado.")
        self.lbl_template_path.setStyleSheet("""
            padding: 8px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #6c757d;
            font-size: 13px;
        """)
        layout.addWidget(self.lbl_template_path)
        
        return widget

    def create_excel_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        self.btn_load_excel = ModernButton("Seleccionar Excel (.xlsx)")
        self.btn_load_excel.clicked.connect(self.load_excel)
        layout.addWidget(self.btn_load_excel)
        
        self.lbl_excel_path = ModernLabel("Ningún archivo seleccionado.")
        self.lbl_excel_path.setStyleSheet("""
            padding: 8px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #6c757d;
            font-size: 13px;
        """)
        layout.addWidget(self.lbl_excel_path)
        
        return widget

    def create_mapping_section(self):
        widget = QWidget()
        layout = QFormLayout(widget)
        layout.setVerticalSpacing(10)
        
        self.combo_text1 = ModernComboBox()
        self.combo_text2 = ModernComboBox()
        self.combo_filename = ModernComboBox()
        
        self.combo_text1.setEnabled(False)
        self.combo_text2.setEnabled(False)
        self.combo_filename.setEnabled(False)
        
        layout.addRow("{{TEXT_1}} (Nombre):", self.combo_text1)
        layout.addRow("{{TEXT_2}} (Título/Evento):", self.combo_text2)
        layout.addRow("Columna para nombre de archivo:", self.combo_filename)
        
        return widget

    def create_style_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(10)
        
        # Estilo para TEXT_1
        text1_group = QWidget()
        text1_layout = QFormLayout(text1_group)
        text1_layout.setVerticalSpacing(8)
        
        self.font_combo_1 = self._get_font_combo()
        self.font_size_spin_1 = QSpinBox()
        self.font_size_spin_1.setRange(8, 72)
        self.font_size_spin_1.setValue(24)
        self.bold_check_1 = QCheckBox("Negrita")
        self.bold_check_1.setChecked(True)
        
        text1_layout.addRow("Fuente:", self.font_combo_1)
        text1_layout.addRow("Tamaño:", self.font_size_spin_1)
        text1_layout.addRow("", self.bold_check_1)
        
        # Estilo para TEXT_2
        text2_group = QWidget()
        text2_layout = QFormLayout(text2_group)
        text2_layout.setVerticalSpacing(8)
        
        self.font_combo_2 = self._get_font_combo()
        self.font_size_spin_2 = QSpinBox()
        self.font_size_spin_2.setRange(8, 72)
        self.font_size_spin_2.setValue(18)
        self.bold_check_2 = QCheckBox("Negrita")
        
        text2_layout.addRow("Fuente:", self.font_combo_2)
        text2_layout.addRow("Tamaño:", self.font_size_spin_2)
        text2_layout.addRow("", self.bold_check_2)
        
        layout.addWidget(ModernLabel("{{TEXT_1}} (Nombre):"))
        layout.addWidget(text1_group)
        layout.addWidget(ModernLabel("{{TEXT_2}} (Título/Evento):"))
        layout.addWidget(text2_group)
        
        # Conectar señales de manera segura
        self.font_combo_1.currentTextChanged.connect(self.update_preview)
        self.font_size_spin_1.valueChanged.connect(self.update_preview)
        self.bold_check_1.stateChanged.connect(self.update_preview)
        self.font_combo_2.currentTextChanged.connect(self.update_preview)
        self.font_size_spin_2.valueChanged.connect(self.update_preview)
        self.bold_check_2.stateChanged.connect(self.update_preview)
        
        return widget

    def create_signature_section(self):
        """Crea la sección para controlar la firma digital y QR"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        # Checkbox para habilitar/deshabilitar firma digital
        self.signature_checkbox = QCheckBox("🔐 Habilitar Firma Digital y Código QR")
        self.signature_checkbox.setChecked(True)  # Por defecto activado
        self.signature_checkbox.setStyleSheet("""
            QCheckBox {
                font-weight: bold;
                font-size: 14px;
                color: #2c3e50;
                padding: 8px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
            QCheckBox::indicator:checked {
                background-color: #28a745;
                border: 2px solid #218838;
            }
            QCheckBox::indicator:unchecked {
                background-color: #dc3545;
                border: 2px solid #c82333;
            }
        """)
        layout.addWidget(self.signature_checkbox)
        
        # Información sobre la función
        info_label = ModernLabel("Cuando está activado: Se inserta código QR con firma digital\nCuando está desactivado: Se genera constancia sin QR ni firma")
        info_label.setStyleSheet("""
            padding: 8px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #495057;
            font-size: 12px;
        """)
        layout.addWidget(info_label)
        
        return widget

    def create_validation_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        validation_layout = QHBoxLayout()
        self.validation_text_entry = ModernLineEdit()
        self.validation_text_entry.setPlaceholderText("Texto de validación personalizado")
        self.validation_text_entry.setText("Validado por Rally de la Niñez Científica y EXPO STEM, Universidad de Sonora")
        validation_layout.addWidget(self.validation_text_entry)
        
        self.btn_apply_validation_text = ModernButton("💾 Aplicar")
        self.btn_apply_validation_text.setStyleSheet("background-color: #6c757d;")
        self.btn_apply_validation_text.clicked.connect(self.apply_validation_text)
        validation_layout.addWidget(self.btn_apply_validation_text)
        
        layout.addLayout(validation_layout)
        
        self.validation_status_label = ModernLabel("Leyenda actual: Validado por Rally de la Niñez Científica y EXPO STEM, Universidad de Sonora")
        self.validation_status_label.setStyleSheet("""
            padding: 8px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #495057;
            font-size: 12px;
        """)
        layout.addWidget(self.validation_status_label)
        
        return widget

    def create_validation_check_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        self.btn_validate = ModernButton("🔍 Validar Configuración")
        self.btn_validate.setStyleSheet("background-color: #17a2b8;")
        self.btn_validate.clicked.connect(self.validate_configuration)
        layout.addWidget(self.btn_validate)
        
        self.validation_label = ModernLabel("Estado: Sin validar")
        self.validation_label.setStyleSheet("""
            padding: 12px;
            background-color: #e9ecef;
            border-radius: 6px;
            color: #6c757d;
            font-size: 13px;
            min-height: 60px;
        """)
        layout.addWidget(self.validation_label)
        
        return widget

    def create_config_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        self.export_mode_combo = ModernComboBox()
        self.export_mode_combo.addItems(["Individual", "Un solo PDF combinado"])
        layout.addWidget(ModernLabel("Modo de exportación:"))
        layout.addWidget(self.export_mode_combo)
        
        return widget

    def create_actions_section(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        
        self.btn_send_email = ModernButton("📤 Enviar Constancias por Correo")
        self.btn_send_email.setStyleSheet("background-color: #007bff;")
        self.btn_send_email.clicked.connect(self.open_email_sender)
        layout.addWidget(self.btn_send_email)
        
        self.btn_generate = ModernButton("🚀 Generar Constancias")
        self.btn_generate.setStyleSheet("""
            background-color: #28a745; 
            font-size: 16px;
            padding: 12px;
        """)
        self.btn_generate.clicked.connect(self.start_generation)
        layout.addWidget(self.btn_generate)
        
        self.btn_cancel = ModernButton("❌ Cancelar")
        self.btn_cancel.setStyleSheet("background-color: #dc3545;")
        self.btn_cancel.clicked.connect(self.cancel_generation)
        layout.addWidget(self.btn_cancel)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)
        
        return widget

    def _get_font_combo(self):
        combo = ModernComboBox()
        fonts_dir = resource_path(os.path.join('assets', 'fonts'))
        if os.path.exists(fonts_dir):
            fonts = [os.path.splitext(f)[0] for f in os.listdir(fonts_dir) if f.lower().endswith(('.ttf', '.otf'))]
            if fonts:
                combo.addItems(sorted(fonts))
            else:
                combo.addItems(["Arial", "Times New Roman", "Courier New", "Georgia", "Verdana"])
        else:
            combo.addItems(["Arial", "Times New Roman", "Courier New", "Georgia", "Verdana"])
        return combo

    def load_template_presets(self):
        self.template_preset_combo.clear()
        self.template_preset_combo.addItem("-- Seleccionar plantilla predefinida --", None)
        for preset in self.template_library.get_all_presets():
            self.template_preset_combo.addItem(f"🎓 {preset.name}", preset.id)
        self.template_preset_combo.addItem("-- Guardar configuración actual --", "save_current")

    def apply_template_preset(self):
        preset_id = self.template_preset_combo.currentData()
        if not preset_id or preset_id == "save_current":
            return
        preset = self.template_library.get_preset(preset_id)
        if preset:
            for placeholder, font_config in preset.recommended_fonts.items():
                if placeholder == 'NOMBRE':
                    self.font_combo_1.setCurrentText(font_config['family'])
                    self.font_size_spin_1.setValue(font_config.get('size', 24))
                    self.bold_check_1.setChecked(font_config.get('bold', False))
                elif placeholder in ['CURSO', 'EVENTO']:
                    self.font_combo_2.setCurrentText(font_config['family'])
                    self.font_size_spin_2.setValue(font_config.get('size', 18))
                    self.bold_check_2.setChecked(font_config.get('bold', False))
            self.log_message(f"✅ Plantilla '{preset.name}' aplicada")
            self.update_preview()

    def load_template(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Plantilla", "", 
            "Archivos Soportados (*.pdf *.docx *.pptx);;PDF (*.pdf);;Word (*.docx);;PowerPoint (*.pptx)"
        )
        if path:
            self.template_path = path
            self.lbl_template_path.setText(os.path.basename(path))
            self.log_message(f"📄 Plantilla cargada: {os.path.basename(path)}")
            self.update_preview()

    def load_excel(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Seleccionar Archivo Excel", "", "Excel (*.xlsx *.xls);;Todos los archivos (*)"
        )
        if path:
            try:
                self.excel_columns, self.excel_data = get_excel_data(path)
                self.lbl_excel_path.setText(os.path.basename(path))
                self.combo_text1.clear()
                self.combo_text2.clear()
                self.combo_text1.addItems(self.excel_columns)
                self.combo_text2.addItems(self.excel_columns)
                self.combo_text1.setEnabled(True)
                self.combo_text2.setEnabled(True)
                self.log_message(f"📊 Lista cargada con {len(self.excel_data)} registros.")
                self.log_message(f"📋 Columnas detectadas: {', '.join(self.excel_columns)}")
            
                # Actualizar selector de columna para nombre de archivo
                self.combo_filename.clear()
                self.combo_filename.addItems(self.excel_columns)
                self.combo_filename.setEnabled(True)
                try:
                    # Por defecto seleccionar la misma columna que TEXT_1 si existe
                    self.combo_filename.setCurrentIndex(self.combo_text1.currentIndex())
                except:
                    pass
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error al cargar Excel: {str(e)}")
                self.log_message(f"❌ Error: {e}")

    def _get_font_info(self, index):
        if index == 1:
            return {
                'family': self.font_combo_1.currentText(),
                'size': self.font_size_spin_1.value(),
                'bold': self.bold_check_1.isChecked(),
            }
        elif index == 2:
            return {
                'family': self.font_combo_2.currentText(),
                'size': self.font_size_spin_2.value(),
                'bold': self.bold_check_2.isChecked(),
            }
        return {'family': 'Arial', 'size': 12, 'bold': False}

    def _get_font_map(self):
        return {
            "{{TEXT_1}}": self._get_font_info(1),
            "{{TEXT_2}}": self._get_font_info(2)
        }

    def update_preview(self):
        if not self.template_path or not self.template_path.lower().endswith('.pdf'):
            if self.template_path:
                self.preview_label.setText("La previsualización en tiempo real solo está disponible para plantillas PDF.\n\nPara DOCX/PPTX, use la validación para verificar la configuración.")
            else:
                self.preview_label.setText("Cargue una plantilla PDF para ver la previsualización.")
            return

        try:
            processor = get_processor(self.template_path)
            if not isinstance(processor, PdfProcessor):
                return

            font_map = self._get_font_map()
            data_map = {
                "{{TEXT_1}}": "María González López",
                "{{TEXT_2}}": "PROYECTO: Desarrollo Sostenible"
            }
            
            pix_data = processor.get_preview_pixmap(data_map, font_map)
            if pix_data:
                img = QImage(pix_data.samples, pix_data.width, pix_data.height, 
                           pix_data.stride, QImage.Format.Format_RGB888)
                pixmap = QPixmap.fromImage(img)
                # Redimensionar a la mitad del tamaño original
                scaled_pixmap = pixmap.scaled(
                    self.preview_label.width() // 2, 
                    self.preview_label.height() // 2,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.preview_label.setPixmap(scaled_pixmap)
        except Exception as e:
            self.preview_label.setText(f"Error en previsualización:\n{str(e)}")

    def apply_validation_text(self):
        """Aplica la leyenda de validación personalizada"""
        validation_text = self.validation_text_entry.text().strip()
        if not validation_text:
            QMessageBox.warning(self, "Campo vacío", "Por favor ingrese un texto para la leyenda de validación.")
            return
        
        try:
            from signature import set_validation_text
            set_validation_text(validation_text)
            self.validation_status_label.setText(f"Leyenda actual: {validation_text}")
            self.log_message(f"🏷️ Leyenda de validación actualizada: {validation_text}")
            
            QMessageBox.information(self, "Leyenda Aplicada", 
                                  f"La leyenda de validación ha sido actualizada:\n\n{validation_text}\n\n"
                                  f"Esta leyenda aparecerá en todas las constancias generadas a partir de ahora.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo aplicar la leyenda: {str(e)}")

    def validate_configuration(self):
        self.validation_label.setText("🔄 Validando configuración...")
        self.validation_label.setStyleSheet("padding: 12px; border: 1px solid #FFEAA7; border-radius: 6px; background-color: #FFF3CD; color: #856404;")
        
        validation_results = []
        
        # Validar plantilla
        if self.template_path:
            template_validation = self.validator.validate_template(self.template_path)
            if template_validation['is_valid']:
                validation_results.append("✅ Plantilla válida")
                if template_validation['placeholders_found']:
                    validation_results.append(f"📝 Placeholders: {', '.join(template_validation['placeholders_found'])}")
            else:
                validation_results.append("❌ Plantilla inválida")
                for error in template_validation['errors']:
                    validation_results.append(f"   • {error}")
        else:
            validation_results.append("❌ No hay plantilla cargada")
    
        # Validar datos Excel
        if hasattr(self, 'excel_data') and self.excel_data:
            validation_results.append(f"✅ Excel válido ({len(self.excel_data)} registros)")
            validation_results.append(f"📋 Columnas: {', '.join(self.excel_columns)}")
        else:
            validation_results.append("❌ No hay datos de Excel cargados")
        
        # Validar mapeo
        if self.combo_text1.currentText() and self.combo_text2.currentText():
            validation_results.append("✅ Mapeo de columnas configurado")
        else:
            validation_results.append("❌ Mapeo de columnas incompleto")
        
        # Validar fuentes
        available_fonts = [self.font_combo_1.itemText(i) for i in range(self.font_combo_1.count())]
        font_validation = self.validator.validate_fonts(self._get_font_map(), available_fonts)
        if font_validation['is_valid']:
            validation_results.append("✅ Fuentes disponibles")
        else:
            validation_results.append("⚠️ Problemas con fuentes")
            for warning in font_validation['warnings']:
                validation_results.append(f"   • {warning}")
        
        # Mostrar resultados
        result_text = "\n".join(validation_results)
        has_errors = any("❌" in result for result in validation_results)
        
        if not has_errors:
            self.validation_label.setStyleSheet("padding: 12px; border: 1px solid #C3E6CB; border-radius: 6px; background-color: #D4EDDA; color: #155724;")
            self.validation_label.setText("✅ Configuración válida\n" + result_text)
        else:
            self.validation_label.setStyleSheet("padding: 12px; border: 1px solid #F5C6CB; border-radius: 6px; background-color: #F8D7DA; color: #721c24;")
            self.validation_label.setText("❌ Problemas encontrados\n" + result_text)
        
        self.log_message("🔍 Validación completada")

    def start_generation(self):
        if not self.template_path:
            QMessageBox.warning(self, "Archivos Faltantes", "Seleccione una plantilla primero.")
            return
        
        if not hasattr(self, 'excel_data') or not self.excel_data:
            QMessageBox.warning(self, "Archivos Faltantes", "Cargue un archivo Excel primero.")
            return
        
        if not self.combo_text1.currentText() or not self.combo_text2.currentText():
            QMessageBox.warning(self, "Configuración Incompleta", "Seleccionen las columnas para ambos placeholders.")
            return
        
        output_dir = QFileDialog.getExistingDirectory(self, "Seleccionar Carpeta de Destino")
        if not output_dir:
            return

        self.performance_optimizer.optimize_memory()

        font_map = self._get_font_map()
        placeholder_map = {
            "{{TEXT_1}}": self.combo_text1.currentText(),
            "{{TEXT_2}}": self.combo_text2.currentText()
        }
        export_mode = self.export_mode_combo.currentText()
        
        # Obtener estado de la firma digital
        enable_signature = self.signature_checkbox.isChecked()
        
        # Columna seleccionada para nombrar archivos (si está habilitada)
        filename_column = self.combo_filename.currentText() if getattr(self, 'combo_filename', None) and self.combo_filename.isEnabled() else None

        self.btn_generate.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setValue(0)

        # Pasar el parámetro enable_signature al worker
        self.worker = Worker(self.template_path, self.excel_data, output_dir, font_map, 
                            placeholder_map, export_mode, filename_column, enable_signature)
    
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log.connect(self.log_message)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.start()
        
        mode_text = "con firma digital" if enable_signature else "sin firma digital"
        self.log_message(f"🚀 Iniciando generación de constancias {mode_text}...")

    def cancel_generation(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.stop()
            self.log_message("⏹️ Enviando señal de cancelación...")
            self.btn_cancel.setEnabled(False)

    def on_generation_finished(self, message):
        self.btn_generate.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        if "error" in message.lower():
            QMessageBox.critical(self, "Error", message)
            self.log_message(f"❌ {message}")
        else:
            QMessageBox.information(self, "Proceso Finalizado", message)
            self.log_message(f"🎉 {message}")
        self.performance_optimizer.optimize_memory()

    def log_message(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.append(f"[{timestamp}] {message}")
        self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())

    def open_email_sender(self):
        """Abre el diálogo para enviar constancias por correo"""
        try:
            self.email_dialog = EmailSenderDialog(self)
            self.email_dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir el envío de correos: {str(e)}")

    def resizeEvent(self, event):
        """Redimensiona el banner cuando cambia el tamaño de la ventana"""
        super().resizeEvent(event)
        # Actualizar banner al nuevo tamaño
        banner_pixmap = QPixmap(resource_path('assets/Banner.png'))
        if not banner_pixmap.isNull():
            # Escalar al 90% del ancho de la ventana
            new_width = int(self.width() * 0.9)
            scaled_pixmap = banner_pixmap.scaledToWidth(new_width, Qt.TransformationMode.SmoothTransformation)
            self.banner_label.setPixmap(scaled_pixmap)
        
        self.update_preview()

def main():
    app = QApplication(sys.argv)
    
    # Configurar estilo de la aplicación
    app.setStyle('Fusion')
    
    window = App()
    window.show()
    
    sys.exit(app.exec())

if __name__ == "__main__":
    main()
