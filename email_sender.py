# email_sender.py
import smtplib
import os
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
import glob
from datetime import datetime
from PyQt6.QtCore import QThread, pyqtSignal
import re

class EmailSender(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal(str)
    
    def __init__(self, config, excel_data, pdf_folder):
        super().__init__()
        self.config = config
        self.excel_data = excel_data
        self.pdf_folder = pdf_folder
        self.is_running = True
        
        # CONVERSIÓN SEGURA A DATAFRAME - CORRECCIÓN DEL ERROR
        self._convert_to_dataframe()
        
        # Configuración SMTP para diferentes proveedores
        self.smtp_config = {
            'gmail.com': {
                'server': 'smtp.gmail.com',
                'port': 587
            },
            'outlook.com': {
                'server': 'smtp.office365.com', 
                'port': 587
            },
            'hotmail.com': {
                'server': 'smtp.office365.com',
                'port': 587
            },
            'yahoo.com': {
                'server': 'smtp.mail.yahoo.com',
                'port': 587
            },
            'live.com': {
                'server': 'smtp.office365.com',
                'port': 587
            }
        }

    def _convert_to_dataframe(self):
        """Convierte los datos de Excel a DataFrame si es necesario"""
        try:
            if isinstance(self.excel_data, list):
                # Si es una lista, convertir a DataFrame
                self.excel_data = pd.DataFrame(self.excel_data)
                self.log.emit("📊 Datos convertidos de lista a DataFrame")
            elif isinstance(self.excel_data, pd.DataFrame):
                # Ya es DataFrame, no hacer nada
                pass
            else:
                # Tipo desconocido, intentar conversión
                self.excel_data = pd.DataFrame(self.excel_data)
                self.log.emit("📊 Datos convertidos a DataFrame")
        except Exception as e:
            self.log.emit(f"⚠️ Error en conversión: {str(e)}")
            # Crear DataFrame vacío para evitar errores
            self.excel_data = pd.DataFrame()

    def get_smtp_config(self, email):
        """Obtiene configuración SMTP basada en el dominio del email"""
        try:
            domain = email.split('@')[-1].lower()
            
            # Dominios que usan Office365/Outlook
            outlook_domains = [
                'outlook.com', 'hotmail.com', 'live.com',
                'office365.com', 'microsoft.com',
                'unison.mx'  # Dominio institucional agregado
            ]
            
            # Si es dominio de Outlook o personalizado que usa Outlook
            if any(domain.endswith(outlook_domain) for outlook_domain in outlook_domains):
                return {
                    'server': 'smtp.office365.com',
                    'port': 587
                }
            
            return self.smtp_config.get(domain, {
                'server': 'smtp.gmail.com',
                'port': 587
            })
        except:
            return {'server': 'smtp.gmail.com', 'port': 587}

    def run(self):
        """Ejecuta el envío de correos en un hilo separado"""
        try:
            results = self.send_emails()
            self.finished.emit(results)
        except Exception as e:
            self.finished.emit(f"error: {str(e)}")

    def stop(self):
        """Detiene el envío de correos"""
        self.is_running = False
        self.log.emit("⏹️ Cancelando envío...")

    def send_emails(self):
        """Envía correos electrónicos con constancias adjuntas"""
        # VERIFICAR QUE SEA DATAFRAME - CORRECCIÓN CLAVE
        if not isinstance(self.excel_data, pd.DataFrame):
            self.excel_data = pd.DataFrame(self.excel_data)
            
        total_emails = len(self.excel_data)
        if total_emails == 0:
            return "error: No hay datos para enviar"

        success_count = 0
        failed_count = 0
        errors = []

        try:
            # Obtener configuración SMTP
            smtp_config = self.get_smtp_config(self.config['email'])
            
            self.log.emit(f"🔗 Conectando a {smtp_config['server']}:{smtp_config['port']}")
            
            # Configurar servidor SMTP
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
            server.starttls()  # Usar TLS para seguridad
            server.login(self.config['email'], self.config['password'])
            
            self.log.emit(f"✅ Conexión exitosa. Enviando desde: {self.config['email']}")
            self.log.emit(f"📊 Total de correos a enviar: {total_emails}")
            
            # CORRECCIÓN: Verificar que podemos usar iterrows()
            if not hasattr(self.excel_data, 'iterrows'):
                return "error: Los datos no son un DataFrame válido de Pandas"
            
            # Enviar correos a cada participante
            for index, row in self.excel_data.iterrows():
                if not self.is_running:
                    break
                    
                try:
                    # VERIFICAR COLUMNAS EXISTENTES
                    if (self.config['name_column'] not in row.index or 
                        self.config['email_column'] not in row.index or 
                        self.config['filename_column'] not in row.index):
                        error_msg = f"❌ Columnas no encontradas en fila {index}"
                        self.log.emit(error_msg)
                        errors.append(error_msg)
                        failed_count += 1
                        continue
                        
                    participant_name = str(row[self.config['name_column']])
                    participant_email = str(row[self.config['email_column']])
                    pdf_filenames = str(row[self.config['filename_column']])
                    
                    # Calcular progreso
                    progress = int((index + 1) / total_emails * 100)
                    self.progress.emit(progress)
                    
                    self.log.emit(f"📧 Procesando: {participant_name} -> {participant_email}")
                    
                    # Buscar archivos PDF
                    pdf_paths = self._find_pdf_files(self.pdf_folder, pdf_filenames)
                    
                    if not pdf_paths:
                        error_msg = f"❌ PDFs no encontrados: {pdf_filenames}"
                        self.log.emit(error_msg)
                        errors.append(error_msg)
                        failed_count += 1
                        continue
                    
                    # Crear y enviar mensaje
                    msg = self._create_email_message(
                        self.config['email'],
                        self.config['sender_name'],
                        participant_email,
                        participant_name,
                        self.config['subject'],
                        self.config['body'],
                        pdf_paths
                    )
                    
                    server.send_message(msg)
                    self.log.emit(f"✅ Enviado a: {participant_name}")
                    success_count += 1
                    
                except Exception as e:
                    participant_name = row.get(self.config['name_column'], 'Desconocido') if hasattr(row, 'get') else 'Desconocido'
                    error_msg = f"❌ Error con {participant_name}: {str(e)}"
                    self.log.emit(error_msg)
                    errors.append(error_msg)
                    failed_count += 1
            
            server.quit()
            
            # Resultado final
            if self.is_running:
                message = f"🎉 Envío completado: {success_count} exitosos, {failed_count} fallidos"
                if errors and failed_count > 0:
                    message += f"\n\nErrores encontrados:\n" + "\n".join(errors[:3])
                    if len(errors) > 3:
                        message += f"\n... y {len(errors) - 3} errores más"
            else:
                message = f"⏹️ Envío cancelado: {success_count} enviados antes de cancelar"
                
            return message
            
        except smtplib.SMTPAuthenticationError as e:
            return f"error: Error de autenticación. Verifique correo y contraseña. Detalles: {str(e)}"
        except smtplib.SMTPException as e:
            return f"error: Error SMTP: {str(e)}"
        except Exception as e:
            return f"error: Error general: {str(e)}"
    
    def _find_pdf_files(self, pdf_folder: str, filenames: str) -> list:
        """Busca múltiples archivos PDF en la carpeta especificada"""
        pdf_paths = []
        
        if not filenames or pd.isna(filenames):
            return pdf_paths
            
        # Separar por comas, punto y coma o saltos de línea
        separators = [',', ';', '\n', '\t']
        filename_list = [filenames.strip()]
        
        for separator in separators:
            if separator in filenames:
                filename_list = [f.strip() for f in filenames.split(separator) if f.strip()]
                break
        
        for filename in filename_list:
            if not filename:
                continue
                
            patterns = [
                filename,
                f"{filename}.pdf",
                f"{filename.replace(' ', '_')}.pdf",
                f"{filename.replace(' ', '')}.pdf",
                f"*{filename}*.pdf",
                f"*{filename.replace(' ', '_')}*.pdf",
                f"*{filename.replace(' ', '')}*.pdf"
            ]
            
            found = False
            for pattern in patterns:
                full_pattern = os.path.join(pdf_folder, pattern)
                matches = glob.glob(full_pattern)
                
                for match in matches:
                    if os.path.isfile(match) and match.lower().endswith('.pdf'):
                        if match not in pdf_paths:
                            pdf_paths.append(match)
                            found = True
                            self.log.emit(f"   📄 Encontrado: {os.path.basename(match)}")
                
                if found:
                    break
        
        return pdf_paths

    def _create_email_message(self, from_email: str, sender_name: str, to_email: str, 
                            participant_name: str, subject: str, body: str, pdf_paths: list) -> MIMEMultipart:
        """Crea el mensaje de correo electrónico con múltiples adjuntos"""
        msg = MIMEMultipart('mixed')
        msg['From'] = formataddr((sender_name, from_email))
        msg['To'] = to_email
        msg['Subject'] = subject
        
        # Personalizar cuerpo del mensaje con placeholders
        personalized_body = self._personalize_body(body, participant_name)
        
        # Convertir HTML de Qt a HTML estándar para correos
        clean_html = self._convert_qt_html_to_standard_html(personalized_body)
        
        # Crear parte alternativa para HTML y texto plano
        msg_alternative = MIMEMultipart('alternative')
        msg.attach(msg_alternative)
        
        # Parte de texto plano (fallback para clientes que no soportan HTML)
        plain_text = self._html_to_plain_text(clean_html)
        msg_alternative.attach(MIMEText(plain_text, 'plain', 'utf-8'))
        
        # Parte HTML
        msg_alternative.attach(MIMEText(clean_html, 'html', 'utf-8'))
        
        # Adjuntar múltiples PDFs
        for pdf_path in pdf_paths:
            try:
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_attachment = MIMEApplication(pdf_file.read(), _subtype='pdf')
                    pdf_name = os.path.basename(pdf_path)
                    pdf_attachment.add_header('Content-Disposition', 'attachment', 
                                            filename=pdf_name)
                    msg.attach(pdf_attachment)
                    self.log.emit(f"   📎 Adjuntado: {pdf_name}")
            except Exception as e:
                self.log.emit(f"   ⚠️ Error adjuntando {pdf_path}: {str(e)}")
        
        return msg

    def _convert_qt_html_to_standard_html(self, qt_html: str) -> str:
        """Convierte HTML de Qt (qrichtext) a HTML estándar para correos"""
        try:
            # Si no es HTML de Qt, retornar tal cual
            if 'qrichtext' not in qt_html:
                return qt_html
                
            # Reemplazar estilos específicos de Qt por CSS estándar
            clean_html = qt_html
            
            # Remover metatags específicos de Qt
            clean_html = clean_html.replace('<meta name="qrichtext" content="1" />', '')
            clean_html = clean_html.replace('<meta charset="utf-8" />', '<meta charset="utf-8">')
            
            # Simplificar estilos CSS
            clean_html = clean_html.replace('font-family:\'Segoe UI\';', 'font-family: Arial, sans-serif;')
            clean_html = clean_html.replace('font-size:9pt;', 'font-size: 11pt;')
            clean_html = clean_html.replace('font-weight:700;', 'font-weight: bold;')
            clean_html = clean_html.replace('font-weight:400;', 'font-weight: normal;')
            
            # Simplificar estilos de párrafo
            clean_html = clean_html.replace('-qt-block-indent:0;', '')
            clean_html = clean_html.replace('text-indent:0px;', '')
            clean_html = clean_html.replace('line-height:115%;', 'line-height: 1.4;')
            
            # Reemplazar márgenes específicos
            clean_html = clean_html.replace('margin-top:12px; margin-bottom:12px; margin-left:0px; margin-right:0px;', 
                                          'margin: 12px 0;')
            clean_html = clean_html.replace('margin-top:12px; margin-bottom:12px; margin-left:20px; margin-right:0px;', 
                                          'margin: 12px 0 12px 20px;')
            
            # Agregar estilos adicionales para mejor compatibilidad
            style_additions = """
            <style type="text/css">
                body { 
                    font-family: Arial, sans-serif; 
                    font-size: 11pt; 
                    line-height: 1.4;
                    color: #333333;
                    margin: 0;
                    padding: 20px;
                    background-color: #ffffff;
                }
                p { 
                    margin: 12px 0;
                    line-height: 1.4;
                }
                .justify {
                    text-align: justify;
                }
                .bold {
                    font-weight: bold;
                }
                .indent {
                    margin-left: 20px;
                }
            </style>
            """
            
            # Insertar estilos en el head si existe, sino crear head
            if '<head>' in clean_html:
                clean_html = clean_html.replace('</head>', style_additions + '</head>')
            else:
                # Si no hay head, crear uno básico
                clean_html = f'<html><head>{style_additions}</head><body>{clean_html}</body></html>'
                
            return clean_html
            
        except Exception as e:
            # En caso de error, retornar HTML limpio básico
            self.log.emit(f"⚠️ Error convirtiendo HTML: {str(e)}")
            return f"""
            <!DOCTYPE html>
            <html>
            <head>
                <meta charset="utf-8">
                <style>
                    body {{ font-family: Arial, sans-serif; font-size: 11pt; line-height: 1.4; color: #333; }}
                    p {{ margin: 12px 0; }}
                    .bold {{ font-weight: bold; }}
                </style>
            </head>
            <body>
                {self._html_to_plain_text(qt_html)}
            </body>
            </html>
            """

    def _html_to_plain_text(self, html: str) -> str:
        """Convierte HTML a texto plano para clients que no soportan HTML"""
        try:
            # Remover tags HTML básicos
            text = re.sub(r'<br\s*/?>', '\n', html)
            text = re.sub(r'<p[^>]*>', '\n', text)
            text = re.sub(r'</p>', '\n', text)
            text = re.sub(r'<[^>]+>', '', text)
            
            # Decodificar entidades HTML básicas
            text = text.replace('&nbsp;', ' ')
            text = text.replace('&amp;', '&')
            text = text.replace('&lt;', '<')
            text = text.replace('&gt;', '>')
            
            # Limpiar espacios múltiples y saltos de línea
            text = re.sub(r'\n\s*\n', '\n\n', text)
            text = re.sub(r'[ \t]+', ' ', text)
            
            return text.strip()
            
        except Exception as e:
            # Fallback: extraer texto entre tags body
            if '<body' in html and '</body>' in html:
                body_content = html.split('<body')[1].split('</body>')[0]
                body_content = body_content.split('>', 1)[1] if '>' in body_content else body_content
                return body_content
            return "Constancia de participación - Universidad de Sonora"

    def _personalize_body(self, body: str, participant_name: str) -> str:
        """Personaliza el cuerpo del mensaje con placeholders"""
        personalized = body
        
        # Reemplazar placeholders manteniendo el formato HTML
        personalized = personalized.replace('{nombre}', participant_name)
        personalized = personalized.replace('{Nombre}', participant_name.title())
        personalized = personalized.replace('{fecha}', datetime.now().strftime('%d/%m/%Y'))
        personalized = personalized.replace('{FECHA}', datetime.now().strftime('%d/%m/%Y'))
        
        return personalized

    def test_connection(self, email: str, password: str):
        """Prueba la conexión con el servidor SMTP"""
        try:
            smtp_config = self.get_smtp_config(email)
            
            self.log.emit(f"🔗 Probando conexión con {smtp_config['server']}:{smtp_config['port']}")
            
            server = smtplib.SMTP(smtp_config['server'], smtp_config['port'])
            server.starttls()
            server.login(email, password)
            server.quit()
            
            return True, f"✅ Conexión exitosa con {smtp_config['server']}"
            
        except smtplib.SMTPAuthenticationError as e:
            return False, f"❌ Error de autenticación. Verifique:\n• Correo y contraseña correctos\n• Para Outlook: Use contraseña normal\n• Detalles: {str(e)}"
        
        except smtplib.SMTPConnectError as e:
            return False, f"❌ Error de conexión. Verifique:\n• Su conexión a internet\n• El firewall no bloquea la aplicación\n• Detalles: {str(e)}"
        
        except Exception as e:
            return False, f"❌ Error: {str(e)}"