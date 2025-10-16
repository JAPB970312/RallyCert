# style_manager.py (nuevo archivo)
"""
Módulo para gestionar estilos de texto en las constancias
"""

from PyQt6.QtGui import QColor, QFont
from PyQt6.QtCore import QObject, pyqtSignal
import json

class StyleManager(QObject):
    style_changed = pyqtSignal(dict)
    
    def __init__(self):
        super().__init__()
        self.default_styles = {
            'font_family': 'Arial',
            'font_size': 12,
            'font_color': '#000000',
            'bold': False,
            'italic': False
        }
        self.current_styles = self.default_styles.copy()
        self.folio_styles = {
            'font_family': 'Arial',
            'font_size': 10,
            'font_color': '#6c757d',
            'bold': True,
            'italic': False
        }
    
    def update_style(self, element: str, style_type: str, value):
        """Actualiza un estilo específico"""
        if element == 'main':
            self.current_styles[style_type] = value
        elif element == 'folio':
            self.folio_styles[style_type] = value
        
        self.style_changed.emit(self.get_all_styles())
    
    def get_style(self, element: str) -> dict:
        """Obtiene los estilos para un elemento específico"""
        if element == 'main':
            return self.current_styles.copy()
        elif element == 'folio':
            return self.folio_styles.copy()
        else:
            return self.default_styles.copy()
    
    def get_all_styles(self) -> dict:
        """Obtiene todos los estilos"""
        return {
            'main': self.current_styles.copy(),
            'folio': self.folio_styles.copy()
        }
    
    def set_font_family(self, element: str, font_family: str):
        """Establece la familia de fuente"""
        self.update_style(element, 'font_family', font_family)
    
    def set_font_size(self, element: str, font_size: int):
        """Establece el tamaño de fuente"""
        self.update_style(element, 'font_size', font_size)
    
    def set_font_color(self, element: str, color: QColor):
        """Establece el color de fuente"""
        if isinstance(color, QColor):
            color_str = color.name()
        else:
            color_str = str(color)
        self.update_style(element, 'font_color', color_str)
    
    def set_bold(self, element: str, bold: bool):
        """Establece negrita"""
        self.update_style(element, 'bold', bold)
    
    def set_italic(self, element: str, italic: bool):
        """Establece cursiva"""
        self.update_style(element, 'italic', italic)
    
    def reset_styles(self):
        """Restablece los estilos a los valores por defecto"""
        self.current_styles = self.default_styles.copy()
        self.folio_styles = {
            'font_family': 'Arial',
            'font_size': 10,
            'font_color': '#6c757d',
            'bold': True,
            'italic': False
        }
        self.style_changed.emit(self.get_all_styles())
    
    def save_styles(self, file_path: str):
        """Guarda los estilos en un archivo JSON"""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(self.get_all_styles(), f, indent=2)
        except Exception as e:
            print(f"Error guardando estilos: {e}")
    
    def load_styles(self, file_path: str):
        """Carga los estilos desde un archivo JSON"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                styles = json.load(f)
                self.current_styles = styles.get('main', self.default_styles.copy())
                self.folio_styles = styles.get('folio', self.folio_styles.copy())
                self.style_changed.emit(self.get_all_styles())
        except Exception as e:
            print(f"Error cargando estilos: {e}")

# Instancia global
style_manager = StyleManager()