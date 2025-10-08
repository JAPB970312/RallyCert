# worker.py
import os
import fitz
import glob
from PyQt6.QtCore import QThread, pyqtSignal
from document_processor import get_processor

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, template_path, excel_data, output_dir, font_map, placeholder_map, export_mode, filename_column=None):
        super().__init__()
        self.template_path = template_path
        self.excel_data = excel_data
        self.output_dir = output_dir
        self.font_map = font_map
        self.placeholder_map = placeholder_map
        self.filename_column = filename_column
        self.export_mode = export_mode
        self.is_cancelled = False

    def run(self):
        try:
            total_files = len(self.excel_data)
            self.log.emit(f"Iniciando generación de {total_files} constancias...")
            
            # Procesamiento simple (sin multiprocessing)
            self.run_single_thread(total_files)

        except Exception as e:
            self.finished.emit(f"Ocurrió un error crítico: {e}")

    def run_single_thread(self, total_files):
        """Procesamiento en hilo único para todos los formatos"""
        
        combined_doc = fitz.open() if self.export_mode == "Un solo PDF combinado" else None
        success_count = 0
        temp_files_to_cleanup = []  # Lista para archivos temporales
        
        try:
            used_filenames = set()
            for i, record in enumerate(self.excel_data):
                if self.is_cancelled:
                    self.log.emit("Proceso cancelado por el usuario.")
                    break
                
                try:
                    # Crear procesador para cada documento
                    processor = get_processor(self.template_path)
                    data_map = {
                        placeholder: record.get(column_name, '') 
                        for placeholder, column_name in self.placeholder_map.items()
                    }
                    
                    
                    # Determinar el nombre base usando la columna seleccionada por el usuario si está disponible
                    if getattr(self, 'filename_column', None):
                        name_for_file = record.get(self.filename_column, '') or data_map.get('{{TEXT_1}}', f'Constancia_{i+1}')
                    else:
                        name_for_file = data_map.get('{{TEXT_1}}', f'Constancia_{i+1}')
                    # Normalizar el nombre (usar el método del procesador si está disponible)
                    try:
                        clean_name = processor._get_clean_filename(name_for_file)
                    except Exception:
                        clean_name = "".join(c for c in str(name_for_file) if c.isalnum() or c in (' ', '-', '_')).rstrip()
                    if not clean_name:
                        clean_name = f'Constancia_{i+1}'
                    # Asegurar unicidad: si ya existe, añadir sufijo incremental "(1)", "(2)", ...
                    base_name = clean_name
                    count = 0
                    candidate = base_name
                    while True:
                        output_candidate = os.path.join(self.output_dir, f"{candidate}.pdf")
                        if not os.path.exists(output_candidate) and candidate not in used_filenames:
                            break
                        count += 1
                        candidate = f"{base_name} ({count})"
                    used_filenames.add(candidate)
                    output_filename = os.path.join(self.output_dir, f"{candidate}.pdf")
                    processor.process(data_map, self.font_map)
                    processor.save_as_pdf(output_filename)

                    
                    if self.export_mode == "Un solo PDF combinado":
                        with fitz.open(output_filename) as temp_doc:
                            combined_doc.insert_pdf(temp_doc)
                        # Marcar para limpieza
                        temp_files_to_cleanup.append(output_filename)

                    success_count += 1
                    self.log.emit(f"✅ ({i+1}/{total_files}) Generada para: {name_for_file}")
                    self.progress.emit(int(((i + 1) / total_files) * 100))
                    
                except Exception as e:
                    self.log.emit(f"❌ Error en registro {i+1}: {str(e)}")
            
            if not self.is_cancelled:
                if self.export_mode == "Un solo PDF combinado":
                    final_path = os.path.join(self.output_dir, "Constancias_Combinadas.pdf")
                    combined_doc.save(final_path)
                    combined_doc.close()
                    self.finished.emit(f"¡Proceso completado! Se generó 1 PDF combinado con {success_count} constancias.")
                else:
                    self.finished.emit(f"¡Proceso completado! Se generaron {success_count} de {total_files} constancias.")
            else:
                self.finished.emit("Proceso detenido por el usuario.")
                
        finally:
            # LIMPIEZA FINAL DE ARCHIVOS TEMPORALES
            self._cleanup_temp_files(temp_files_to_cleanup)
            
            # Cerrar documentos combinados si existen
            if combined_doc:
                combined_doc.close()
        
    def _cleanup_temp_files(self, temp_files):
        """Limpia archivos temporales con manejo de errores"""
        cleaned_count = 0
        for temp_file in temp_files:
            try:
                if os.path.exists(temp_file):
                    os.remove(temp_file)
                    cleaned_count += 1
            except Exception as e:
                self.log.emit(f"⚠️ No se pudo eliminar temporal: {os.path.basename(temp_file)}")
        
        if cleaned_count > 0:
            self.log.emit(f"🧹 Se limpiaron {cleaned_count} archivos temporales")
        
    def stop(self):
        self.is_cancelled = True