# worker.py (modificado y corregido)
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

    def __init__(self, template_path, excel_data, output_dir, font_map, placeholder_map, export_mode, filename_column=None, enable_signature=True, enable_folio=True, folio_column=None, folio_font_map=None):
        super().__init__()
        self.template_path = template_path
        self.excel_data = excel_data
        self.output_dir = output_dir
        self.font_map = font_map
        self.placeholder_map = placeholder_map
        self.filename_column = filename_column
        self.export_mode = export_mode
        self.enable_signature = enable_signature
        self.enable_folio = enable_folio
        self.folio_column = folio_column
        self.folio_font_map = folio_font_map or {}
        self.is_cancelled = False

        # Ensure keys exist (solo si la firma está habilitada)
        if self.enable_signature:
            try:
                ensure_keys(PRIVATE_KEY_PATH, PUBLIC_KEY_PATH)
            except Exception as e:
                print(f"[signature] Warning: no se pudieron generar/validar llaves: {e}")

    def run(self):
        try:
            total_files = len(self.excel_data)
            mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
            folio_text = "con folio" if self.enable_folio else "sin folio"
            self.log.emit(f"Iniciando generación de {total_files} constancias {mode_text} {folio_text}...")
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

                    # AGREGAR FOLIO AL DATA_MAP SI ESTÁ HABILITADO
                    if self.enable_folio and self.folio_column:
                        folio_value = record.get(self.folio_column, '')
                        if folio_value:
                            data_map["{{FOLIO}}"] = str(folio_value)
                        else:
                            data_map["{{FOLIO}}"] = f"FOLIO-{i+1:06d}"

                    # COMBINAR FONT_MAP CON FOLIO_FONT_MAP
                    combined_font_map = self.font_map.copy()
                    if self.enable_folio and self.folio_font_map:
                        combined_font_map["{{FOLIO}}"] = self.folio_font_map

                    if getattr(self, 'filename_column', None) and self.filename_column:
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
                    processor.process(data_map, combined_font_map)
                    processor.save_as_pdf(output_filename)

                    # --- FIRMAR Y EMBEDIR AUTOMÁTICAMENTE (SOLO SI ESTÁ HABILITADO) ---
                    if self.enable_signature:
                        # Preparar datos para el código QR con soporte para caracteres especiales
                        cert_data = {
                            "nombre": data_map.get("{{TEXT_1}}", ""),
                            "evento": data_map.get("{{TEXT_2}}", ""),
                            "folio": data_map.get("{{FOLIO}}", f"RALLY-{int(i+1):06d}"),
                            "fecha_emision": datetime.now().strftime('%Y-%m-%d'),
                            "institucion": "Universidad de Sonora"
                        }
                        
                        try:
                            # Firmar el documento (sobrescribe el mismo archivo)
                            metadata = sign_and_embed(output_filename, output_filename, cert_data)
                            folio_display = data_map.get("{{FOLIO}}", "N/A")
                            self.log.emit(f"🔐 Firma añadida: {os.path.basename(output_filename)} (Folio: {folio_display})")
                        except Exception as e:
                            self.log.emit(f"⚠️ No se pudo firmar {os.path.basename(output_filename)}: {e}")
                    else:
                        folio_display = data_map.get("{{FOLIO}}", "N/A") if self.enable_folio else "N/A"
                        self.log.emit(f"📄 Generado sin firma: {os.path.basename(output_filename)} (Folio: {folio_display})")

                    # Si estamos combinando PDFs, insertar después de procesar
                    if self.export_mode == "Un solo PDF combinado":
                        with fitz.open(output_filename) as temp_doc:
                            combined_doc.insert_pdf(temp_doc)
                        temp_files_to_cleanup.append(output_filename)

                    success_count += 1
                    
                    # Log con información del folio
                    folio_info = f" | Folio: {data_map.get('{{FOLIO}}', 'N/A')}" if self.enable_folio else ""
                    self.log.emit(f"✅ ({i+1}/{total_files}) Generada para: {name_for_file}{folio_info}")
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
                                "fecha_emision": datetime.now().strftime('%Y-%m-%d'),
                                "institucion": "Universidad de Sonora"
                            }
                            metadata = sign_and_embed(final_path, final_path, cert_data)
                            self.log.emit(f"🔐 Firma añadida al documento combinado")
                        except Exception as e:
                            self.log.emit(f"⚠️ No se pudo firmar documento combinado: {e}")
                    
                    mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
                    folio_text = "con folio" if self.enable_folio else "sin folio"
                    self.finished.emit(f"¡Proceso completado! Se generó 1 PDF combinado {mode_text} {folio_text} con {success_count} constancias.")
                else:
                    mode_text = "con firma digital" if self.enable_signature else "sin firma digital"
                    folio_text = "con folio" if self.enable_folio else "sin folio"
                    self.finished.emit(f"¡Proceso completado! Se generaron {success_count} de {total_files} constancias {mode_text} {folio_text}.")
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

    def stop(self):
        """Detiene la generación de manera segura"""
        self.is_cancelled = True
        self.log.emit("⏹️ Cancelando proceso...")
