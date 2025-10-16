# folio_manager.py (nuevo archivo)
"""
Módulo para gestionar folios alfanuméricos en las constancias
"""

import re
import random
import string
from typing import Optional, Dict, Any

class FolioManager:
    def __init__(self):
        self.folio_cache = {}
        self.custom_patterns = {}
    
    def generate_folio(self, base_data: Dict[str, Any], pattern: str = "RALLY-{counter:06d}") -> str:
        """
        Genera un folio alfanumérico basado en un patrón
        
        Args:
            base_data: Diccionario con datos base para el folio
            pattern: Patrón para generar el folio. Puede incluir:
                    {counter} - Contador numérico
                    {random} - Caracteres aleatorios
                    {name_initials} - Iniciales del nombre
                    {date} - Fecha actual
                    Campos del base_data entre llaves
        """
        try:
            # Reemplazar campos del base_data
            for key, value in base_data.items():
                if isinstance(value, str):
                    placeholder = f"{{{key}}}"
                    if placeholder in pattern:
                        # Tomar primeras letras si es un nombre largo
                        if 'initials' in key.lower() or 'nombre' in key.lower():
                            initials = ''.join([word[0].upper() for word in value.split()[:2] if word])
                            pattern = pattern.replace(placeholder, initials)
                        else:
                            pattern = pattern.replace(placeholder, value[:10])  # Limitar longitud
            
            # Generar componentes dinámicos
            if "{counter}" in pattern:
                counter_value = self._get_next_counter()
                pattern = pattern.replace("{counter}", str(counter_value))
            
            if "{counter:06d}" in pattern:
                counter_value = self._get_next_counter()
                pattern = pattern.replace("{counter:06d}", f"{counter_value:06d}")
            
            if "{random}" in pattern:
                random_chars = ''.join(random.choices(string.ascii_uppercase + string.digits, k=6))
                pattern = pattern.replace("{random}", random_chars)
            
            if "{date}" in pattern:
                from datetime import datetime
                current_date = datetime.now().strftime("%Y%m%d")
                pattern = pattern.replace("{date}", current_date)
            
            # Limpiar caracteres no permitidos
            folio = self._clean_folio(pattern)
            
            # Guardar en cache para evitar duplicados
            self.folio_cache[folio] = True
            
            return folio
            
        except Exception as e:
            # Fallback: folio simple
            counter = self._get_next_counter()
            return f"RALLY-{counter:06d}"
    
    def _get_next_counter(self) -> int:
        """Obtiene el siguiente número de contador"""
        if not hasattr(self, '_counter'):
            self._counter = 0
        self._counter += 1
        return self._counter
    
    def _clean_folio(self, folio: str) -> str:
        """Limpia el folio de caracteres no deseados"""
        # Permitir letras, números, guiones y underscores
        cleaned = re.sub(r'[^\w\-_]', '', folio)
        return cleaned.upper()
    
    def validate_folio(self, folio: str) -> bool:
        """Valida que el folio tenga formato correcto"""
        if not folio or len(folio) < 3 or len(folio) > 50:
            return False
        
        # Verificar que solo contenga caracteres permitidos
        if not re.match(r'^[\w\-_]+$', folio):
            return False
        
        return True
    
    def extract_folio_components(self, folio: str) -> Dict[str, str]:
        """Extrae componentes del folio para análisis"""
        components = {
            'original': folio,
            'numeric_part': '',
            'alpha_part': '',
            'separator': ''
        }
        
        # Extraer parte numérica
        numbers = re.findall(r'\d+', folio)
        if numbers:
            components['numeric_part'] = numbers[0]
        
        # Extraer parte alfabética
        letters = re.findall(r'[A-Za-z]+', folio)
        if letters:
            components['alpha_part'] = letters[0]
        
        # Identificar separador
        separators = re.findall(r'[^A-Za-z0-9]', folio)
        if separators:
            components['separator'] = separators[0]
        
        return components

# Instancia global
folio_manager = FolioManager()