# ui.py
import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QLabel,
    QFileDialog, QSpinBox, QCheckBox, QProgressBar, QMessageBox, QTextEdit, QComboBox, QFormLayout,
    QTabWidget, QSlider, QInputDialog, QScrollArea, QDialog, QLineEdit, QToolBar, QTextBrowser
)
from PyQt6.QtGui import QPixmap, QImage, QIcon, QTextCursor, QTextCharFormat, QTextBlockFormat, QTextFormat, QFont
from PyQt6.QtCore import Qt, QRegularExpression, QSize
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

class EmailSenderDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("📤 Enviar Constancias por Correo")
        self.setGeometry(200, 200, 800, 900)
        self.email_sender = None
        self.setup_ui()

    def setup_ui(self):
        # FONDO BLANCO PARA TODO EL DIÁLOGO
        self.setStyleSheet("""
            QDialog {
                background-color: #ffffff;
                color: #000000;
            }
            QLabel {
                background-color: transparent;
                color: #000000;
            }
            QLineEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
            }
            QComboBox {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
            }
            QSpinBox {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
            }
            QCheckBox {
                background-color: transparent;
                color: #000000;
            }
            QCheckBox::indicator {
                width: 15px;
                height: 15px;
            }
            QCheckBox::indicator:unchecked {
                border: 1px solid #cccccc;
                background-color: #ffffff;
            }
            QCheckBox::indicator:checked {
                border: 1px solid #007bff;
                background-color: #007bff;
            }
            QPushButton {
                background-color: #f8f9fa;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
            QTextEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px;
                font-family: Arial, sans-serif;
                font-size: 11pt;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                background-color: #ffffff;
                text-align: center;
                color: #000000;
            }
            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 3px;
            }
            QToolBar {
                background-color: #f8f9fa;
                border: 1px solid #dee2e6;
                spacing: 3px;
                padding: 5px;
            }
            QScrollArea {
                background-color: #ffffff;
                border: none;
            }
            QScrollBar:vertical {
                background-color: #f8f9fa;
                width: 15px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background-color: #cccccc;
                border-radius: 7px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #aaaaaa;
            }
        """)

        layout = QVBoxLayout(self)

        # Configuración de correo remitente
        email_frame = QWidget()
        email_frame.setStyleSheet("background-color: #ffffff;")
        email_layout = QFormLayout(email_frame)
        
        self.email_entry = QLineEdit()
        self.email_entry.setPlaceholderText("ejemplo@gmail.com")
        self.email_entry.textChanged.connect(self.validate_form)
        email_layout.addRow("Correo electrónico:", self.email_entry)

        self.password_entry = QLineEdit()
        self.password_entry.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_entry.setPlaceholderText("Ingrese su contraseña o contraseña de aplicación")
        self.password_entry.textChanged.connect(self.validate_form)
        email_layout.addRow("Contraseña:", self.password_entry)

        # Checkbox para mostrar contraseña
        self.show_password = QCheckBox("Mostrar contraseña")
        self.show_password.toggled.connect(self.toggle_password_visibility)
        email_layout.addRow("", self.show_password)

        self.sender_name_entry = QLineEdit()
        self.sender_name_entry.setText("Generador de Constancias")
        self.sender_name_entry.setPlaceholderText("Nombre que aparecerá como remitente")
        self.sender_name_entry.textChanged.connect(self.validate_form)
        email_layout.addRow("Nombre del remitente:", self.sender_name_entry)

        layout.addWidget(email_frame)

        # Selección de archivos
        files_frame = QWidget()
        files_frame.setStyleSheet("background-color: #ffffff;")
        files_layout = QFormLayout(files_frame)

        self.pdf_folder_entry = QLineEdit()
        self.pdf_folder_entry.setPlaceholderText("Seleccione la carpeta con los PDFs")
        self.btn_select_pdf = QPushButton("Seleccionar Carpeta")
        pdf_layout = QHBoxLayout()
        pdf_layout.addWidget(self.pdf_folder_entry)
        pdf_layout.addWidget(self.btn_select_pdf)
        files_layout.addRow("Carpeta de PDFs:", pdf_layout)

        self.excel_file_entry = QLineEdit()
        self.excel_file_entry.setPlaceholderText("Seleccione el archivo Excel con los datos")
        self.btn_select_excel = QPushButton("Seleccionar Excel")
        excel_layout = QHBoxLayout()
        excel_layout.addWidget(self.excel_file_entry)
        excel_layout.addWidget(self.btn_select_excel)
        files_layout.addRow("Archivo Excel:", excel_layout)

        layout.addWidget(files_frame)

        # Configuración de columnas
        columns_frame = QWidget()
        columns_frame.setStyleSheet("background-color: #ffffff;")
        columns_layout = QFormLayout(columns_frame)

        self.name_column_combo = QComboBox()
        self.name_column_combo.currentTextChanged.connect(self.validate_form)
        columns_layout.addRow("Columna Nombre:", self.name_column_combo)

        self.email_column_combo = QComboBox()
        self.email_column_combo.currentTextChanged.connect(self.validate_form)
        columns_layout.addRow("Columna Correo:", self.email_column_combo)

        self.filename_column_combo = QComboBox()
        self.filename_column_combo.currentTextChanged.connect(self.validate_form)
        columns_layout.addRow("Columna Archivo PDF:", self.filename_column_combo)

        # Información sobre múltiples PDFs
        info_label = QLabel("💡 Para múltiples PDFs, separar nombres por comas\nEjemplo: certificado1, certificado2")
        info_label.setStyleSheet("color: #000000; font-size: 10pt; margin-top: 10px; background-color: transparent;")
        columns_layout.addRow("", info_label)

        layout.addWidget(columns_frame)

        # Contenido del correo con barra de herramientas de formato
        content_frame = QWidget()
        content_frame.setStyleSheet("background-color:#ffffff;")
        content_layout = QVBoxLayout(content_frame)

        # Barra de herramientas de formato
        format_toolbar = QToolBar()
        format_toolbar.setIconSize(QSize(16, 16))
        
        # Botones de formato de texto
        self.btn_bold = QPushButton()
        self.btn_bold.setIcon(QIcon(resource_path("assets/icons/bold.png")))
        self.btn_bold.setToolTip("Negrita (Ctrl+B)")
        self.btn_bold.setFixedSize(30, 30)
        self.btn_bold.setStyleSheet("color: #000000;")
        self.btn_bold.clicked.connect(self.toggle_bold)
        format_toolbar.addWidget(self.btn_bold)

        self.btn_italic = QPushButton()
        self.btn_italic.setIcon(QIcon(resource_path("assets/icons/italic.png")))
        self.btn_italic.setToolTip("Itálica (Ctrl+I)")
        self.btn_italic.setFixedSize(30, 30)
        self.btn_italic.setStyleSheet("color: #000000;")
        self.btn_italic.clicked.connect(self.toggle_italic)
        format_toolbar.addWidget(self.btn_italic)

        self.btn_underline = QPushButton()
        self.btn_underline.setIcon(QIcon(resource_path("assets/icons/underline.png")))
        self.btn_underline.setToolTip("Subrayado (Ctrl+U)")
        self.btn_underline.setFixedSize(30, 30)
        self.btn_underline.setStyleSheet("color: #000000;")
        self.btn_underline.clicked.connect(self.toggle_underline)
        format_toolbar.addWidget(self.btn_underline)

        # Separador
        format_toolbar.addSeparator()

        # Alineación - CORREGIDO: texto negro visible
        self.btn_align_left = QPushButton()
        self.btn_align_left.setIcon(QIcon(resource_path("assets/icons/align_left.png")))
        self.btn_align_left.setToolTip("Alinear a la izquierda")
        self.btn_align_left.setFixedSize(30, 30)
        self.btn_align_left.setStyleSheet("color: #000000;")
        self.btn_align_left.clicked.connect(lambda: self.set_alignment(Qt.AlignmentFlag.AlignLeft))
        format_toolbar.addWidget(self.btn_align_left)

        self.btn_align_center = QPushButton()
        self.btn_align_center.setIcon(QIcon(resource_path("assets/icons/align_center.png")))
        self.btn_align_center.setToolTip("Centrar texto")
        self.btn_align_center.setFixedSize(30, 30)
        self.btn_align_center.setStyleSheet("color: #000000;")
        self.btn_align_center.clicked.connect(lambda: self.set_alignment(Qt.AlignmentFlag.AlignCenter))
        format_toolbar.addWidget(self.btn_align_center)

        self.btn_align_right = QPushButton()
        self.btn_align_right.setIcon(QIcon(resource_path("assets/icons/align_right.png")))
        self.btn_align_right.setToolTip("Alinear a la derecha")
        self.btn_align_right.setFixedSize(30, 30)
        self.btn_align_right.setStyleSheet("color: #000000;")
        self.btn_align_right.clicked.connect(lambda: self.set_alignment(Qt.AlignmentFlag.AlignRight))
        format_toolbar.addWidget(self.btn_align_right)

        self.btn_align_justify = QPushButton()
        self.btn_align_justify.setIcon(QIcon(resource_path("assets/icons/align_justify.png")))
        self.btn_align_justify.setToolTip("Justificar texto")
        self.btn_align_justify.setFixedSize(30, 30)
        self.btn_align_justify.setStyleSheet("color: #000000;")
        self.btn_align_justify.clicked.connect(lambda: self.set_alignment(Qt.AlignmentFlag.AlignJustify))
        format_toolbar.addWidget(self.btn_align_justify)

        # Separador
        format_toolbar.addSeparator()

        # Interlineado
        interlineado_layout = QHBoxLayout()
        interlineado_label = QLabel("Interlineado:")
        interlineado_label.setStyleSheet("color: #000000; background-color: transparent;")
        interlineado_layout.addWidget(interlineado_label)
        self.line_spacing_combo = QComboBox()
        self.line_spacing_combo.addItems(["1.0", "1.15", "1.5", "2.0"])
        self.line_spacing_combo.setCurrentText("1.15")
        self.line_spacing_combo.currentTextChanged.connect(self.apply_line_spacing)
        interlineado_layout.addWidget(self.line_spacing_combo)

        # Sangría
        sangria_label = QLabel("Sangría:")
        sangria_label.setStyleSheet("color: #000000; background-color: transparent;")
        interlineado_layout.addWidget(sangria_label)
        self.indent_combo = QComboBox()
        self.indent_combo.addItems(["0px", "10px", "20px", "30px", "40px"])
        self.indent_combo.setCurrentText("0px")
        self.indent_combo.currentTextChanged.connect(self.apply_indentation)
        interlineado_layout.addWidget(self.indent_combo)

        interlineado_widget = QWidget()
        interlineado_widget.setLayout(interlineado_layout)
        interlineado_widget.setStyleSheet("background-color: transparent;")
        format_toolbar.addWidget(interlineado_widget)

        content_layout.addWidget(format_toolbar)

        # Asunto
        subject_layout = QHBoxLayout()
        subject_label = QLabel("Asunto:")
        subject_label.setStyleSheet("color: #000000; background-color: transparent;")
        subject_layout.addWidget(subject_label)
        self.subject_entry = QLineEdit()
        self.subject_entry.setText("Constancia de Participación")
        self.subject_entry.setPlaceholderText("Asunto del correo electrónico")
        self.subject_entry.textChanged.connect(self.validate_form)
        subject_layout.addWidget(self.subject_entry)
        content_layout.addLayout(subject_layout)

        # Cuerpo del mensaje
        body_label = QLabel("Cuerpo del mensaje:")
        body_label.setStyleSheet("color: #000000; background-color: transparent;")
        content_layout.addWidget(body_label)

        self.body_text = QTextEdit()
        self.body_text.setMinimumHeight(200)
        
        # TEXTO PREDETERMINADO MEJORADO - COLORES VISIBLES
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

        # Conectar señales de cambio de formato
        self.body_text.cursorPositionChanged.connect(self.update_format_buttons)
        self.body_text.selectionChanged.connect(self.update_format_buttons)

        content_layout.addWidget(self.body_text)

        # Información sobre placeholders
        placeholder_info = QLabel("💡 Placeholders disponibles: {nombre}, {Nombre}, {fecha} | Atajos: Ctrl+B, Ctrl+I, Ctrl+U")
        placeholder_info.setStyleSheet("color: #000000; font-size: 9pt; margin-top: 5px; background-color: transparent;")
        content_layout.addWidget(placeholder_info)

        layout.addWidget(content_frame)

        # Botones de acción
        buttons_layout = QHBoxLayout()
        self.btn_test = QPushButton("🔍 Probar Conexión")
        self.btn_send = QPushButton("📤 Enviar Correos")
        self.btn_cancel = QPushButton("❌ Cancelar")

        # Botones especiales con colores específicos
        self.btn_test.setStyleSheet("""
            QPushButton {
                background-color: #17a2b8; 
                color: white; 
                font-weight: bold; 
                padding: 8px;
                border: none;
            }
            QPushButton:hover {
                background-color: #138496;
            }
        """)
        self.btn_send.setStyleSheet("""
            QPushButton {
                background-color: #28a745; 
                color: white; 
                font-weight: bold; 
                padding: 8px;
                border: none;
            }
            QPushButton:hover {
                background-color: #218838;
            }
        """)
        self.btn_cancel.setStyleSheet("""
            QPushButton {
                background-color: #dc3545; 
                color: white; 
                font-weight: bold; 
                padding: 8px;
                border: none;
            }
            QPushButton:hover {
                background-color: #c82333;
            }
        """)

        buttons_layout.addWidget(self.btn_test)
        buttons_layout.addWidget(self.btn_send)
        buttons_layout.addWidget(self.btn_cancel)

        layout.addLayout(buttons_layout)

        # Barra de progreso
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        # Etiqueta de estado
        self.status_label = QLabel("Listo para configurar el envío")
        self.status_label.setStyleSheet("""
            padding: 10px; 
            background-color: #f8f9fa; 
            border: 1px solid #dee2e6; 
            border-radius: 4px;
            color: #000000;
            font-weight: normal;
        """)
        self.status_label.setWordWrap(True)
        layout.addWidget(self.status_label)

        # Conectar señales
        self.btn_select_pdf.clicked.connect(self.select_pdf_folder)
        self.btn_select_excel.clicked.connect(self.select_excel_file)
        self.btn_test.clicked.connect(self.test_connection)
        self.btn_send.clicked.connect(self.start_sending)
        self.btn_cancel.clicked.connect(self.cancel_operation)

        # Configurar atajos de teclado
        self.setup_shortcuts()

        # Inicializar estado de botones
        self.btn_send.setEnabled(False)

    def setup_shortcuts(self):
        """Configura atajos de teclado para formato"""
        # Los atajos se manejarán en keyPressEvent

    def keyPressEvent(self, event):
        """Maneja atajos de teclado para formato de texto"""
        if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            if event.key() == Qt.Key.Key_B:
                self.toggle_bold()
                event.accept()
                return
            elif event.key() == Qt.Key.Key_I:
                self.toggle_italic()
                event.accept()
                return
            elif event.key() == Qt.Key.Key_U:
                self.toggle_underline()
                event.accept()
                return
        super().keyPressEvent(event)

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
                padding: 10px; 
                background-color: #d4edda; 
                border: 1px solid #c3e6cb; 
                border-radius: 4px; 
                color: #155724;
                font-weight: normal;
            """)
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo cargar el archivo Excel: {str(e)}")
            self.status_label.setText("❌ Error al cargar Excel")
            self.status_label.setStyleSheet("""
                padding: 10px; 
                background-color: #f8d7da; 
                border: 1px solid #f5c6cb; 
                border-radius: 4px; 
                color: #721c24;
                font-weight: normal;
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
                padding: 10px; 
                background-color: #d4edda; 
                border: 1px solid #c3e6cb; 
                border-radius: 4px; 
                color: #155724;
                font-weight: normal;
            """)
        else:
            self.status_label.setText("ℹ️ Complete todos los campos para habilitar el envío")
            self.status_label.setStyleSheet("""
                padding: 10px; 
                background-color: #fff3cd; 
                border: 1px solid #ffeaa7; 
                border-radius: 4px; 
                color: #856404;
                font-weight: normal;
            """)

    # MÉTODOS DE FORMATO DE TEXTO
    def toggle_bold(self):
        """Alterna formato negrita en el texto seleccionado"""
        cursor = self.body_text.textCursor()
        if cursor.hasSelection():
            format = QTextCharFormat()
            font_weight = QFont.Weight.Normal if cursor.charFormat().fontWeight() == QFont.Weight.Bold else QFont.Weight.Bold
            format.setFontWeight(font_weight)
            cursor.mergeCharFormat(format)
        else:
            # Alternar formato para el texto que se escribirá
            current_format = self.body_text.currentCharFormat()
            new_weight = QFont.Weight.Normal if current_format.fontWeight() == QFont.Weight.Bold else QFont.Weight.Bold
            new_format = QTextCharFormat()
            new_format.setFontWeight(new_weight)
            self.body_text.setCurrentCharFormat(new_format)

    def toggle_italic(self):
        """Alterna formato itálica en el texto seleccionado"""
        cursor = self.body_text.textCursor()
        if cursor.hasSelection():
            format = QTextCharFormat()
            format.setFontItalic(not cursor.charFormat().fontItalic())
            cursor.mergeCharFormat(format)
        else:
            current_format = self.body_text.currentCharFormat()
            new_format = QTextCharFormat()
            new_format.setFontItalic(not current_format.fontItalic())
            self.body_text.setCurrentCharFormat(new_format)

    def toggle_underline(self):
        """Alterna formato subrayado en el texto seleccionado"""
        cursor = self.body_text.textCursor()
        if cursor.hasSelection():
            format = QTextCharFormat()
            format.setFontUnderline(not cursor.charFormat().fontUnderline())
            cursor.mergeCharFormat(format)
        else:
            current_format = self.body_text.currentCharFormat()
            new_format = QTextCharFormat()
            new_format.setFontUnderline(not current_format.fontUnderline())
            self.body_text.setCurrentCharFormat(new_format)

    def set_alignment(self, alignment):
        """Establece la alineación del párrafo"""
        cursor = self.body_text.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.Document)
        
        block_format = QTextBlockFormat()
        block_format.setAlignment(alignment)
        cursor.mergeBlockFormat(block_format)

    def apply_line_spacing(self, spacing):
        """Aplica el interlineado seleccionado"""
        cursor = self.body_text.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.Document)
        
        block_format = QTextBlockFormat()
        line_height = float(spacing)
        block_format.setLineHeight(line_height * 100, QTextBlockFormat.LineHeightTypes.ProportionalHeight)
        cursor.mergeBlockFormat(block_format)

    def apply_indentation(self, indent):
        """Aplica la sangría seleccionada"""
        cursor = self.body_text.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.Document)
        
        block_format = QTextBlockFormat()
        indent_pixels = int(indent.replace('px', ''))
        block_format.setLeftMargin(indent_pixels)
        cursor.mergeBlockFormat(block_format)

    def update_format_buttons(self):
        """Actualiza el estado de los botones de formato según la selección actual"""
        cursor = self.body_text.textCursor()
        format = cursor.charFormat()
        
        # Actualizar botones de formato de caracteres
        self.btn_bold.setStyleSheet(
            "font-weight: bold; background-color: #e9ecef; color: #000000;" if format.fontWeight() == QFont.Weight.Bold else "color: #000000; background-color: transparent;"
        )
        self.btn_italic.setStyleSheet(
            "font-style: italic; background-color: #e9ecef; color: #000000;" if format.fontItalic() else "color: #000000; background-color: transparent;"
        )
        self.btn_underline.setStyleSheet(
            "text-decoration: underline; background-color: #e9ecef; color: #000000;" if format.fontUnderline() else "color: #000000; background-color: transparent;"
        )
        
        # Actualizar botones de alineación - CORREGIDO: texto siempre negro
        block_format = cursor.blockFormat()
        alignment = block_format.alignment()
        
        self.btn_align_left.setStyleSheet(
            "background-color: #e9ecef; color: #000000;" if alignment == Qt.AlignmentFlag.AlignLeft else "color: #000000; background-color: transparent;"
        )
        self.btn_align_center.setStyleSheet(
            "background-color: #e9ecef; color: #000000;" if alignment == Qt.AlignmentFlag.AlignCenter else "color: #000000; background-color: transparent;"
        )
        self.btn_align_right.setStyleSheet(
            "background-color: #e9ecef; color: #000000;" if alignment == Qt.AlignmentFlag.AlignRight else "color: #000000; background-color: transparent;"
        )
        self.btn_align_justify.setStyleSheet(
            "background-color: #e9ecef; color: #000000;" if alignment == Qt.AlignmentFlag.AlignJustify else "color: #000000; background-color: transparent;"
        )

    def test_connection(self):
        email = self.email_entry.text().strip()
        password = self.password_entry.text().strip()
        
        if not email or not password:
            QMessageBox.warning(self, "Campos requeridos", "Por favor ingrese el correo y contraseña.")
            return
        
        self.status_label.setText("🔄 Probando conexión con el servidor SMTP...")
        self.status_label.setStyleSheet("""
            padding: 10px; 
            background-color: #fff3cd; 
            border: 1px solid #ffeaa7; 
            border-radius: 4px; 
            color: #856404;
            font-weight: normal;
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
                    padding: 10px; 
                    background-color: #d4edda; 
                    border: 1px solid #c3e6cb; 
                    border-radius: 4px; 
                    color: #155724;
                    font-weight: normal;
                """)
                QMessageBox.information(self, "✅ Conexión Exitosa", 
                                      f"Conexión establecida correctamente con:\n\n"
                                      f"Correo: {email}\n"
                                      f"Servidor: SMTP\n\n"
                                      f"Ahora puede proceder con el envío de constancias.")
            else:
                self.status_label.setText(message)
                self.status_label.setStyleSheet("""
                    padding: 10px; 
                    background-color: #f8d7da; 
                    border: 1px solid #f5c6cb; 
                    border-radius: 4px; 
                    color: #721c24;
                    font-weight: normal;
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
                padding: 10px; 
                background-color: #f8d7da; 
                border: 1px solid #f5c6cb; 
                border-radius: 4px; 
                color: #721c24;
                font-weight: normal;
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
                padding: 10px; 
                background-color: #cce7ff; 
                border: 1px solid #b3d9ff; 
                border-radius: 4px; 
                color: #004085;
                font-weight: normal;
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
                padding: 10px; 
                background-color: #f8d7da; 
                border: 1px solid #f5c6cb; 
                border-radius: 4px; 
                color: #721c24;
                font-weight: normal;
            """)
        else:
            QMessageBox.information(self, "Proceso Completado", message)
            self.status_label.setText("✅ " + message)
            self.status_label.setStyleSheet("""
                padding: 10px; 
                background-color: #d4edda; 
                border: 1px solid #c3e6cb; 
                border-radius: 4px; 
                color: #155724;
                font-weight: normal;
            """)
        
        self.reset_interface()

    def cancel_operation(self):
        """Cancela la operación en curso"""
        if self.email_sender and self.email_sender.isRunning():
            self.email_sender.stop()
            self.status_label.setText("⏹️ Cancelando envío...")
            self.status_label.setStyleSheet("""
                padding: 10px; 
                background-color: #fff3cd; 
                border: 1px solid #ffeaa7; 
                border-radius: 4px; 
                color: #856404;
                font-weight: normal;
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

# Clase App principal con fondo blanco
class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("RallyCert")
        self.setGeometry(100, 100, 1200, 800)
        self.setMinimumSize(1100, 700)
        self.setWindowIcon(QIcon(resource_path('assets/icon.ico')))

        # FONDO BLANCO PARA LA VENTANA PRINCIPAL
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
                color: #000000;
            }
            QWidget {
                background-color: #ffffff;
                color: #000000;
            }
            QLabel {
                background-color: transparent;
                color: #000000;
            }
            QPushButton {
                background-color: #f8f9fa;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 8px 16px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #e9ecef;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
            }
            QLineEdit, QComboBox, QSpinBox {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
                padding: 5px;
            }
            QTextEdit {
                background-color: #ffffff;
                color: #000000;
                border: 1px solid #cccccc;
                border-radius: 4px;
            }
            QTabWidget::pane {
                border: 1px solid #cccccc;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #f8f9fa;
                color: #000000;
                padding: 8px 16px;
                border: 1px solid #cccccc;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                border-bottom: 1px solid #ffffff;
            }
            QScrollArea {
                background-color: #ffffff;
                border: none;
            }
            QProgressBar {
                border: 1px solid #cccccc;
                border-radius: 4px;
                background-color: #ffffff;
                text-align: center;
                color: #000000;
            }
            QProgressBar::chunk {
                background-color: #007bff;
                border-radius: 3px;
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

    def open_email_sender(self):
        """Abre el diálogo para enviar constancias por correo"""
        try:
            self.email_dialog = EmailSenderDialog(self)
            self.email_dialog.exec()
        except Exception as e:
            QMessageBox.critical(self, "Error", f"No se pudo abrir el envío de correos: {str(e)}")

    def setup_ui(self):
        """Configura la interfaz de usuario completa"""
        central_widget = QWidget()
        central_widget.setStyleSheet("background-color: #ffffff;")
        self.setCentralWidget(central_widget)
        
        # LAYOUT PRINCIPAL VERTICAL - Banner arriba, contenido abajo
        main_layout = QVBoxLayout(central_widget)
        
        # --- BANNER SUPERIOR CENTRADO ---
        banner_container = QWidget()
        banner_container.setStyleSheet("background-color: #ffffff;")
        banner_layout = QVBoxLayout(banner_container)
        banner_layout.setContentsMargins(0, 0, 0, 0)
        
        self.banner_label = QLabel()
        banner_pixmap = QPixmap(resource_path('assets/Banner.png')) 
        if not banner_pixmap.isNull():
            banner_pixmap = banner_pixmap.scaledToWidth(1000, Qt.TransformationMode.SmoothTransformation)
            self.banner_label.setPixmap(banner_pixmap)
        else:
            self.banner_label.setText("Banner no encontrado")
            self.banner_label.setStyleSheet("color: #000000; font-weight: bold; padding: 20px; background-color: transparent;")
        
        self.banner_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.banner_label.setStyleSheet("margin-bottom: 0px; padding: 0px; background-color: transparent;")
        
        banner_layout.addWidget(self.banner_label)
        main_layout.addWidget(banner_container)
        
        # --- CONTENIDO PRINCIPAL (Panel izquierdo + derecho) ---
        content_widget = QWidget()
        content_widget.setStyleSheet("background-color: #ffffff;")
        content_layout = QHBoxLayout(content_widget)
        content_layout.setContentsMargins(10, 0, 10, 10)
        
        # --- PANEL DE CONTROL (Izquierda) ---
        control_scroll = QScrollArea()
        control_scroll.setWidgetResizable(True)
        control_scroll.setMaximumWidth(500)
        control_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        control_widget = QWidget()
        control_widget.setStyleSheet("background-color: #ffffff;")
        control_panel = QVBoxLayout(control_widget)
        control_scroll.setWidget(control_widget)

        # 1. Plantillas Predefinidas
        control_panel.addWidget(self._create_section_label("🎨 Plantillas Predefinidas"))
        self.template_preset_combo = QComboBox()
        self.template_preset_combo.setMinimumHeight(30)
        self.load_template_presets()
        self.template_preset_combo.currentTextChanged.connect(self.apply_template_preset)
        control_panel.addWidget(self.template_preset_combo)

        # 2. Plantilla
        control_panel.addWidget(self._create_section_label("📄 Cargar Plantilla"))
        self.btn_load_template = QPushButton("Seleccionar Plantilla (.pdf, .docx, .pptx)")
        self.btn_load_template.setMinimumHeight(35)
        self.btn_load_template.clicked.connect(self.load_template)
        self.lbl_template_path = QLabel("Ningún archivo seleccionado.")
        self.lbl_template_path.setWordWrap(True)
        self.lbl_template_path.setStyleSheet("padding: 5px; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; color: #000000;")
        control_panel.addWidget(self.btn_load_template)
        control_panel.addWidget(self.lbl_template_path)

        # 3. Excel
        control_panel.addWidget(self._create_section_label("📊 Cargar Participantes"))
        self.btn_load_excel = QPushButton("Seleccionar Excel (.xlsx)")
        self.btn_load_excel.setMinimumHeight(35)
        self.btn_load_excel.clicked.connect(self.load_excel)
        self.lbl_excel_path = QLabel("Ningún archivo seleccionado.")
        self.lbl_excel_path.setWordWrap(True)
        self.lbl_excel_path.setStyleSheet("padding: 5px; background-color: #f8f9fa; border: 1px solid #dee2e6; border-radius: 4px; color: #000000;")
        control_panel.addWidget(self.btn_load_excel)
        control_panel.addWidget(self.lbl_excel_path)
        
        # 4. Mapeo de Columnas
        control_panel.addWidget(self._create_section_label("🔗 Asignar Columnas"))
        form_layout = QFormLayout()
        self.combo_text1 = QComboBox()
        self.combo_text2 = QComboBox()
        self.combo_text1.setEnabled(False)
        self.combo_text2.setEnabled(False)
        self.combo_text1.setMinimumHeight(30)
        self.combo_text2.setMinimumHeight(30)
        form_layout.addRow("{{TEXT_1}} (Nombre):", self.combo_text1)
        form_layout.addRow("{{TEXT_2}} (Título/Evento):", self.combo_text2)
        
        self.combo_filename = QComboBox()
        self.combo_filename.setEnabled(False)
        self.combo_filename.setMinimumHeight(30)
        form_layout.addRow("Columna para nombre de archivo:", self.combo_filename)
        control_panel.addLayout(form_layout)

        # 5. Estilo de Fuente - CONTROLES INDEPENDIENTES
        control_panel.addWidget(self._create_section_label("🎯 Estilo de Texto"))
        
        # TEXT_1 (Nombre) - CONFIGURACIÓN INDEPENDIENTE
        control_panel.addWidget(QLabel("{{TEXT_1}} (Nombre):"))
        name_style_layout = QHBoxLayout()
        
        self.font_combo_1 = self._get_font_combo()
        self.font_size_spin_1 = QSpinBox()
        self.font_size_spin_1.setRange(8, 72)
        self.font_size_spin_1.setValue(24)
        self.font_size_spin_1.setMinimumHeight(30)
        self.bold_check_1 = QCheckBox("Negrita")
        self.bold_check_1.setChecked(True)
        
        name_style_layout.addWidget(QLabel("Fuente:"))
        name_style_layout.addWidget(self.font_combo_1)
        name_style_layout.addWidget(QLabel("Tamaño:"))
        name_style_layout.addWidget(self.font_size_spin_1)
        name_style_layout.addWidget(self.bold_check_1)
        control_panel.addLayout(name_style_layout)
        
        # TEXT_2 (Título) - CONFIGURACIÓN INDEPENDIENTE
        control_panel.addWidget(QLabel("{{TEXT_2}} (Título/Evento):"))
        title_style_layout = QHBoxLayout()
        
        self.font_combo_2 = self._get_font_combo()
        self.font_size_spin_2 = QSpinBox()
        self.font_size_spin_2.setRange(8, 72)
        self.font_size_spin_2.setValue(18)
        self.font_size_spin_2.setMinimumHeight(30)
        self.bold_check_2 = QCheckBox("Negrita")
        
        title_style_layout.addWidget(QLabel("Fuente:"))
        title_style_layout.addWidget(self.font_combo_2)
        title_style_layout.addWidget(QLabel("Tamaño:"))
        title_style_layout.addWidget(self.font_size_spin_2)
        title_style_layout.addWidget(self.bold_check_2)
        control_panel.addLayout(title_style_layout)

        # Conectar eventos de estilo
        self.font_combo_1.currentTextChanged.connect(self.update_preview)
        self.font_size_spin_1.valueChanged.connect(self.update_preview)
        self.bold_check_1.stateChanged.connect(self.update_preview)
        self.font_combo_2.currentTextChanged.connect(self.update_preview)
        self.font_size_spin_2.valueChanged.connect(self.update_preview)
        self.bold_check_2.stateChanged.connect(self.update_preview)

        # 6. Validación
        control_panel.addWidget(self._create_section_label("✅ Validación"))
        self.btn_validate = QPushButton("🔍 Validar Configuración")
        self.btn_validate.setMinimumHeight(35)
        self.btn_validate.setStyleSheet("background-color: #17a2b8; color: white; font-weight: bold;")
        self.btn_validate.clicked.connect(self.validate_configuration)
        self.validation_label = QLabel("Estado: Sin validar")
        self.validation_label.setWordWrap(True)
        self.validation_label.setMinimumHeight(60)
        self.validation_label.setStyleSheet("padding: 8px; border: 1px solid #ccc; border-radius: 4px; background-color: #f8f9fa; font-size: 10pt; color: #000000;")
        control_panel.addWidget(self.btn_validate)
        control_panel.addWidget(self.validation_label)

        # 7. Exportación
        control_panel.addWidget(self._create_section_label("⚙️ Configuración"))
        self.export_mode_combo = QComboBox()
        self.export_mode_combo.addItems(["Individual", "Un solo PDF combinado"])
        self.export_mode_combo.setMinimumHeight(30)
        control_panel.addWidget(QLabel("Modo de exportación:"))
        control_panel.addWidget(self.export_mode_combo)

        # 8. Acciones
        control_panel.addWidget(self._create_section_label("🚀 Acciones"))
        
        # Botón para enviar por correo (NUEVO)
        self.btn_send_email = QPushButton("📤 Enviar Constancias por Correo")
        self.btn_send_email.setMinimumHeight(40)
        self.btn_send_email.setStyleSheet("background-color: #007bff; color: white; font-weight: bold;")
        self.btn_send_email.clicked.connect(self.open_email_sender)
        control_panel.addWidget(self.btn_send_email)

        # Botón para generar constancias (ORIGINAL)
        self.btn_generate = QPushButton("🚀 Generar Constancias")
        self.btn_generate.setMinimumHeight(45)
        self.btn_generate.setStyleSheet("background-color: #28a745; color: white; font-size: 14pt; font-weight: bold; border-radius: 6px;")
        self.btn_generate.clicked.connect(self.start_generation)
        
        self.btn_cancel = QPushButton("⏹️ Cancelar")
        self.btn_cancel.setMinimumHeight(35)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setStyleSheet("background-color: #dc3545; color: white; font-weight: bold; border-radius: 6px;")
        self.btn_cancel.clicked.connect(self.cancel_generation)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimumHeight(20)
        
        control_panel.addWidget(self.btn_generate)
        control_panel.addWidget(self.btn_cancel)
        control_panel.addWidget(self.progress_bar)

        control_panel.addStretch()

        # --- PANEL DE PREVISUALIZACIÓN (Derecha) ---
        preview_tab = QTabWidget()
        preview_tab.setMinimumWidth(600)

        # Pestaña de Previsualización
        preview_widget = QWidget()
        preview_widget.setStyleSheet("background-color: #ffffff;")
        preview_layout = QVBoxLayout(preview_widget)
        preview_label_title = QLabel("👁️ Previsualización en Tiempo Real")
        preview_label_title.setStyleSheet("color: #000000; background-color: transparent;")
        preview_layout.addWidget(preview_label_title)
        self.preview_label = QLabel("Cargue una plantilla PDF para ver la previsualización.")
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(400)
        self.preview_label.setStyleSheet("border: 2px solid #ccc; background-color: #f8f9fa; border-radius: 8px; padding: 20px; font-size: 11pt; color: #000000;")
        preview_layout.addWidget(self.preview_label, 1)

        # Pestaña de Log
        log_widget = QWidget()
        log_widget.setStyleSheet("background-color: #ffffff;")
        log_layout = QVBoxLayout(log_widget)
        log_label_title = QLabel("📝 Registro de Actividad")
        log_label_title.setStyleSheet("color: #000000; background-color: transparent;")
        log_layout.addWidget(log_label_title)
        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("font-family: 'Consolas', 'Monaco', monospace; font-size: 10pt; color: #000000; background-color: #ffffff;")
        log_layout.addWidget(self.log_box)

        preview_tab.addTab(preview_widget, "👁️ Previsualización")
        preview_tab.addTab(log_widget, "📝 Registro")

        # Añadir paneles al contenido principal
        content_layout.addWidget(control_scroll, 1)
        content_layout.addWidget(preview_tab, 2)
        
        # Añadir contenido al layout principal
        main_layout.addWidget(content_widget, 1)

    def _create_section_label(self, text):
        label = QLabel(text)
        label.setStyleSheet("font-weight: bold; margin-top: 15px; margin-bottom: 8px; font-size: 12pt; color: #000000; background-color: transparent;")
        return label
    
    def _get_font_combo(self):
        combo = QComboBox()
        combo.setMinimumHeight(30)
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
                scaled_pixmap = pixmap.scaled(
                    self.preview_label.width() - 40, 
                    self.preview_label.height() - 40,
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                self.preview_label.setPixmap(scaled_pixmap)
        except Exception as e:
            self.preview_label.setText(f"Error en previsualización:\n{str(e)}")

    def validate_configuration(self):
        self.validation_label.setText("🔄 Validando configuración...")
        self.validation_label.setStyleSheet("padding: 8px; border: 1px solid #FFEAA7; border-radius: 4px; background-color: #FFF3CD; color: #856404;")
        
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
            self.validation_label.setStyleSheet("padding: 8px; border: 1px solid #C3E6CB; border-radius: 4px; background-color: #D4EDDA; color: #155724;")
            self.validation_label.setText("✅ Configuración válida\n" + result_text)
        else:
            self.validation_label.setStyleSheet("padding: 8px; border: 1px solid #F5C6CB; border-radius: 4px; background-color: #F8D7DA; color: #721c24;")
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
            QMessageBox.warning(self, "Configuración Incompleta", "Seleccione las columnas para ambos placeholders.")
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
        # Columna seleccionada para nombrar archivos (si está habilitada)
        filename_column = self.combo_filename.currentText() if getattr(self, 'combo_filename', None) and self.combo_filename.isEnabled() else None

        self.btn_generate.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setValue(0)

        self.worker = Worker(self.template_path, self.excel_data, output_dir, font_map, placeholder_map, export_mode, filename_column)

        
        self.worker.progress.connect(self.progress_bar.setValue)
        self.worker.log.connect(self.log_message)
        self.worker.finished.connect(self.on_generation_finished)
        self.worker.start()
        self.log_message("🚀 Iniciando generación de constancias...")


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

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec())