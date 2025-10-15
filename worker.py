# worker.py
import os
import fitz
import glob
from PyQt6.QtCore import QThread, pyqtSignal
from document_processor import get_processor
from signature import sign_and_embed, ensure_keys, PRIVATE_KEY_PATH, PUBLIC_KEY_PATH
from datetime import datetime

class Worker(QThread):
    progress = pyqtSignal(int)
    finished = pyqtSignal(str)
    log = pyqtSignal(str)

    def __init__(self, template_path, excel_data, output_dir, font_map, placeholder_map, export_mode, filename_column=None, enable_signature=True):
        super().__init__()
        self.template_path = template_path
        self.excel_data = excel_data
        self.output_dir = output_dir
        self.font_map = font_map
        self.placeholder_map = placeholder_map
        self.filename_column = filename_column
        self.export_mode = export_mode
        self.enable_signature = enable_signature  # NUEVO PARÁMETRO
        self.is_cancelled = False

        # Ensure keys exist (solo si la firma está habilitada)
        if self.enable_signature:
            try:
                ensure_keys(PRIVATE_KEY_PATH, PUBLIC_KEY_PATH)
            except Exception as e:
                # no queremos detener el hilo por fallo en keys; lo registramos
                print(f"[signature] Warning: no se pudieron generar/validar llaves: {e}")

    def run(self):
        try:
            total_files = len(self.excel_data)
            mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
            self.log.emit(f"Iniciando generación de {total_files} constancias {mode_text}...")
            self.run_single_thread(total_files)
        except Exception as e:
            self.finished.emit(f"Ocurrió un error crítico: {e}")

    def run_single_thread(self, total_files):
        combined_doc = fitz.open() if self.export_mode == "Un solo PDF combinado" else None
        success_count = 0
        temp_files_to_cleanup = []

        try:
            used_filenames = set()
            for i, record in enumerate(self.excel_data):
                if self.is_cancelled:
                    self.log.emit("Proceso cancelado por el usuario.")
                    break

                try:
                    processor = get_processor(self.template_path)
                    data_map = {
                        placeholder: record.get(column_name, '') 
                        for placeholder, column_name in self.placeholder_map.items()
                    }

                    if getattr(self, 'filename_column', None):
                        name_for_file = record.get(self.filename_column, '') or data_map.get('{{TEXT_1}}', f'Constancia_{i+1}')
                    else:
                        name_for_file = data_map.get('{{TEXT_1}}', f'Constancia_{i+1}')

                    try:
                        clean_name = processor._get_clean_filename(name_for_file)
                    except Exception:
                        clean_name = "".join(c for c in str(name_for_file) if c.isalnum() or c in (' ', '-', '_')).rstrip()

                    if not clean_name:
                        clean_name = f'Constancia_{i+1}'

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

                    # Generar documento y guardarlo como PDF
                    processor.process(data_map, self.font_map)
                    processor.save_as_pdf(output_filename)

                    # --- FIRMAR Y EMBEDIR AUTOMÁTICAMENTE (SOLO SI ESTÁ HABILITADO) ---
                    if self.enable_signature:
                        cert_data = {
                            "nombre": data_map.get("{{TEXT_1}}", ""),
                            "evento": data_map.get("{{TEXT_2}}", ""),
                            "folio": f"RALLY-{int(i+1):06d}",
                        }
                        try:
                            # Firmar el documento (sobrescribe el mismo archivo)
                            metadata = sign_and_embed(output_filename, output_filename, cert_data)
                            self.log.emit(f"🔐 Firma añadida: {os.path.basename(output_filename)}")
                        except Exception as e:
                            self.log.emit(f"⚠️ No se pudo firmar {os.path.basename(output_filename)}: {e}")
                    else:
                        self.log.emit(f"📄 Generado sin firma: {os.path.basename(output_filename)}")

                    # Si estamos combinando PDFs, insertar después de procesar
                    if self.export_mode == "Un solo PDF combinado":
                        with fitz.open(output_filename) as temp_doc:
                            combined_doc.insert_pdf(temp_doc)
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
                    
                    # Agregar firma al PDF combinado si está habilitado
                    if self.enable_signature:
                        try:
                            cert_data = {
                                "documento": "Constancias Combinadas",
                                "total_constancias": success_count,
                                "folio": f"COMBINADO-{datetime.now().strftime('%Y%m%d')}",
                            }
                            metadata = sign_and_embed(final_path, final_path, cert_data)
                            self.log.emit(f"🔐 Firma añadida al documento combinado")
                        except Exception as e:
                            self.log.emit(f"⚠️ No se pudo firmar documento combinado: {e}")
                    
                    mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
                    self.finished.emit(f"¡Proceso completado! Se generó 1 PDF combinado {mode_text} con {success_count} constancias.")
                else:
                    mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
                    self.finished.emit(f"¡Proceso completado! Se generaron {success_count} de {total_files} constancias {mode_text}.")
            else:
                self.finished.emit("Proceso detenido por el usuario.")

        finally:
            # limpieza de temporales
            for temp in temp_files_to_cleanup:
                try:
                    if os.path.exists(temp):
                        os.remove(temp)
                except Exception:
                    pass
            if combined_doc:
                combined_doc.close()
