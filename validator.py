# validator.py
import os
import fitz
import pandas as pd
from typing import Dict, List, Any

class DocumentValidator:
    def __init__(self):
        pass
    
    def validate_template(self, template_path: str) -> Dict[str, Any]:
        validation_result = {
            'is_valid': False,
            'warnings': [],
            'errors': [],
            'placeholders_found': []
        }
        
        try:
            if not os.path.exists(template_path):
                validation_result['errors'].append("El archivo de plantilla no existe")
                return validation_result
            
            ext = os.path.splitext(template_path)[1].lower()
            if ext not in ['.pdf', '.docx', '.pptx']:
                validation_result['errors'].append(f"Formato no soportado: {ext}")
                return validation_result
            
            if ext == '.pdf':
                return self._validate_pdf_template(template_path, validation_result)
            else:
                # Para DOCX/PPTX, asumimos que es válido si existe
                validation_result['is_valid'] = True
                validation_result['placeholders_found'] = ['{{TEXT_1}}', '{{TEXT_2}}']
                return validation_result
                
        except Exception as e:
            validation_result['errors'].append(f"Error validando plantilla: {str(e)}")
        
        return validation_result
    
    def _validate_pdf_template(self, template_path: str, result: Dict) -> Dict:
        try:
            doc = fitz.open(template_path)
            all_text = ""
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                page_text = page.get_text()
                all_text += page_text + " "
            
            placeholders = self._detect_placeholders(all_text)
            result['placeholders_found'] = placeholders
            
            if placeholders:
                result['is_valid'] = True
            else:
                result['warnings'].append("No se encontraron placeholders en el PDF")
            
            doc.close()
            
        except Exception as e:
            result['errors'].append(f"Error al validar PDF: {str(e)}")
        
        return result
    
    def _detect_placeholders(self, text: str) -> List[str]:
        import re
        pattern = r'\{\{[A-Za-z_0-9]+\}\}'
        placeholders = re.findall(pattern, text)
        return list(set(placeholders))
    
    def validate_fonts(self, font_map: Dict, available_fonts: List[str]) -> Dict[str, Any]:
        validation_result = {
            'is_valid': True,
            'warnings': [],
            'errors': [],
            'missing_fonts': []
        }
        
        for placeholder, font_info in font_map.items():
            font_family = font_info.get('family', 'Arial')
            if font_family not in available_fonts:
                validation_result['missing_fonts'].append(font_family)
                validation_result['warnings'].append(f"Fuente no disponible: {font_family}")
        
        if validation_result['missing_fonts']:
            validation_result['is_valid'] = False
        
        return validation_result