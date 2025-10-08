# document_processor.py
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pptx import Presentation
from pptx.util import Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor
import comtypes.client
import os
import tempfile
import time
from abc import ABC, abstractmethod

class BaseProcessor(ABC):
    def __init__(self, template_path):
        self.template_path = template_path
        self.doc = None
        self.temp_files = []  # Lista para rastrear archivos temporales

    @abstractmethod
    def process(self, data_map: dict, font_map: dict):
        pass

    @abstractmethod
    def save_as_pdf(self, output_path: str):
        pass

    def _get_clean_filename(self, text: str) -> str:
        return "".join(c for c in str(text) if c.isalnum() or c in (' ', '-', '_')).rstrip()

    def _cleanup_temp_files(self):
        """Limpia todos los archivos temporales creados"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
            except Exception as e:
                print(f"⚠️ No se pudo eliminar archivo temporal {temp_file}: {e}")
        self.temp_files = []

class PdfProcessor(BaseProcessor):
    def __init__(self, template_path):
        super().__init__(template_path)
        self.doc = fitz.open(template_path)

    def _get_text_fit_info(self, rect, text, font_info):
        font_name = self._get_pdf_font(font_info['family'], font_info['bold'])
        max_width = rect.width * 0.98
        current_font_size = font_info['size']
        tw_initial = fitz.get_text_length(text, fontname=font_name, fontsize=current_font_size)

        final_font_size = current_font_size
        
        if tw_initial > max_width:
            while final_font_size > 6: 
                tw = fitz.get_text_length(text, fontname=font_name, fontsize=final_font_size)
                if tw <= max_width:
                    break
                final_font_size -= 1
        
        tw = fitz.get_text_length(text, fontname=font_name, fontsize=final_font_size)
        x_insert = rect.x0 + (rect.width - tw) / 2
        font_ascender = 0.8
        y_insert = rect.y0 + (rect.height - (final_font_size * font_ascender)) / 2 + final_font_size

        return x_insert, y_insert, final_font_size, font_name

    def process(self, data_map: dict, font_map: dict):
        self.doc = fitz.open(self.template_path)
        for page in self.doc:
            for placeholder, value in data_map.items():
                text_instances = page.search_for(placeholder)
                font_info = font_map.get(placeholder, {'family': 'Arial', 'size': 12, 'bold': False})
                for inst in text_instances:
                    page.add_redact_annot(inst)
                    page.apply_redactions()
                    x_insert, y_insert, final_size, font_name = self._get_text_fit_info(inst, value, font_info)
                    page.insert_text((x_insert, y_insert), value, fontsize=final_size, fontname=font_name, color=(0, 0, 0))

    def save_as_pdf(self, output_path: str):
        try:
            self.doc.save(output_path, garbage=4, deflate=True, clean=True)
        finally:
            self.doc.close()
            # Limpiar archivos temporales (aunque PDF no crea muchos)
            self._cleanup_temp_files()

    def get_preview_pixmap(self, data_map: dict, font_map: dict):
        temp_doc = fitz.open(self.template_path)
        try:
            page = temp_doc.load_page(0)
            for placeholder, value in data_map.items():
                text_instances = page.search_for(placeholder)
                font_info = font_map.get(placeholder, {'family': 'Arial', 'size': 12, 'bold': False})
                if text_instances:
                    inst = text_instances[0]
                    page.add_redact_annot(inst)
                    page.apply_redactions()
                    x_insert, y_insert, final_size, font_name = self._get_text_fit_info(inst, value, font_info)
                    page.insert_text((x_insert, y_insert), value, fontsize=final_size, fontname=font_name, color=(0,0,0))

            pix = page.get_pixmap()
            return pix
        finally:
            temp_doc.close()

    def _get_pdf_font(self, family, bold):
        family_lower = family.lower()
        if "arial" in family_lower or "helvetica" in family_lower: 
            return "helv-bold" if bold else "helv"
        if "times" in family_lower: 
            return "tibo" if bold else "timo"
        if "courier" in family_lower: 
            return "cobo" if bold else "cour"
        return "helv-bold" if bold else "helv"

class OfficeProcessor(BaseProcessor):
    def _convert_to_pdf_with_com(self, input_path: str, output_path: str, app_name: str, format_type: int):
        app = None
        doc = None
        max_retries = 3
        retry_delay = 2
        
        for attempt in range(max_retries):
            try:
                app = comtypes.client.CreateObject(app_name)
                app.Visible = False
                
                if "powerpoint" in app_name.lower():
                    doc = app.Presentations.Open(input_path)
                else:
                    doc = app.Documents.Open(input_path)
                    
                doc.SaveAs(output_path, FileFormat=format_type)
                break  # Éxito, salir del loop de reintentos
                
            except Exception as e:
                print(f"⚠️ Intento {attempt + 1} fallado: {e}")
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    raise e
            finally:
                # Cerrar documentos y aplicación
                if doc:
                    try:
                        doc.Close(False)
                    except:
                        pass
                if app:
                    try:
                        app.Quit()
                    except:
                        pass

class DocxProcessor(OfficeProcessor):
    def process(self, data_map: dict, font_map: dict): 
        self.doc = Document(self.template_path)
        
        # Obtener configuración de fuentes específica para cada placeholder
        font_info_1 = font_map.get("{{TEXT_1}}", {'family': 'Arial', 'size': 24, 'bold': False})
        font_info_2 = font_map.get("{{TEXT_2}}", {'family': 'Arial', 'size': 18, 'bold': False})

        for placeholder, value in data_map.items():
            for paragraph in self.doc.paragraphs:
                if placeholder in paragraph.text:
                    paragraph.text = ""
                    run = paragraph.add_run(str(value))
                    font = run.font
                    
                    # Aplicar configuración específica según el placeholder
                    if placeholder == "{{TEXT_1}}":
                        font.name = font_info_1['family']
                        font.size = Pt(font_info_1['size'])
                        font.bold = font_info_1['bold']
                    elif placeholder == "{{TEXT_2}}":
                        font.name = font_info_2['family']
                        font.size = Pt(font_info_2['size'])
                        font.bold = font_info_2['bold']
                    
                    font.color.rgb = RGBColor(0, 0, 0)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

    def save_as_pdf(self, output_path: str):
        temp_dir = os.path.dirname(output_path)
        
        # Usar nombre temporal más seguro
        temp_docx_path = os.path.join(temp_dir, f"temp_docx_{os.getpid()}_{int(time.time())}.docx")
        self.temp_files.append(temp_docx_path)
        
        try:
            # 1. Guardar el DOCX modificado
            self.doc.save(temp_docx_path)
            
            # 2. Convertir a PDF
            self._convert_to_pdf_with_com(
                os.path.abspath(temp_docx_path),
                os.path.abspath(output_path),
                'Word.Application',
                17  # wdFormatPDF
            )
            
        finally:
            # LIMPIEZA MEJORADA - Intentar múltiples veces
            self._cleanup_with_retry()

    def _cleanup_with_retry(self, max_retries=5, retry_delay=1):
        """Limpia archivos temporales con reintentos"""
        for attempt in range(max_retries):
            try:
                self._cleanup_temp_files()
                break  # Éxito
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    print(f"❌ No se pudieron eliminar algunos archivos temporales: {e}")

class PptxProcessor(OfficeProcessor):
    def process(self, data_map: dict, font_map: dict): 
        self.doc = Presentation(self.template_path)
        
        # Obtener configuración de fuentes específica para cada placeholder
        font_info_1 = font_map.get("{{TEXT_1}}", {'family': 'Arial', 'size': 24, 'bold': False})
        font_info_2 = font_map.get("{{TEXT_2}}", {'family': 'Arial', 'size': 18, 'bold': False})

        for placeholder, value in data_map.items():
            for slide in self.doc.slides:
                for shape in slide.shapes:
                    if not shape.has_text_frame:
                        continue
                    for paragraph in shape.text_frame.paragraphs:
                        if placeholder in paragraph.text:
                            paragraph.text = str(value)
                            if paragraph.runs:
                                font = paragraph.runs[0].font
                                
                                # Aplicar configuración específica según el placeholder
                                if placeholder == "{{TEXT_1}}":
                                    font.name = font_info_1['family']
                                    font.size = PptxPt(font_info_1['size'])
                                    font.bold = font_info_1['bold']
                                elif placeholder == "{{TEXT_2}}":
                                    font.name = font_info_2['family']
                                    font.size = PptxPt(font_info_2['size'])
                                    font.bold = font_info_2['bold']
                                
                                font.color.rgb = PptxRGBColor(0, 0, 0)

    def save_as_pdf(self, output_path: str):
        temp_dir = os.path.dirname(output_path)
        temp_pptx_path = os.path.join(temp_dir, f"temp_pptx_{os.getpid()}_{int(time.time())}.pptx")
        self.temp_files.append(temp_pptx_path)

        try:
            self.doc.save(temp_pptx_path)
            self._convert_to_pdf_with_com(
                os.path.abspath(temp_pptx_path),
                os.path.abspath(output_path),
                'Powerpoint.Application',
                32  # ppSaveAsPDF
            )
        finally:
            # LIMPIEZA MEJORADA
            self._cleanup_with_retry()

    def _cleanup_with_retry(self, max_retries=5, retry_delay=1):
        """Limpia archivos temporales con reintentos"""
        for attempt in range(max_retries):
            try:
                self._cleanup_temp_files()
                break  # Éxito
            except Exception as e:
                if attempt < max_retries - 1:
                    time.sleep(retry_delay)
                else:
                    print(f"❌ No se pudieron eliminar algunos archivos temporales: {e}")

def get_processor(template_path: str):
    ext = os.path.splitext(template_path)[1].lower()
    if ext == '.pdf':
        return PdfProcessor(template_path)
    elif ext == '.docx':
        return DocxProcessor(template_path)
    elif ext == '.pptx':
        return PptxProcessor(template_path)
    else:
        raise ValueError(f"Tipo de archivo no soportado: {ext}")