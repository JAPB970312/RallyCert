# template_library.py
import os
import json
from typing import Dict, List, Optional
from dataclasses import dataclass
from enum import Enum

class TemplateCategory(Enum):
    ACADEMIC = "académico"
    CERTIFICATE = "certificado"
    RECOGNITION = "reconocimiento"

@dataclass
class TemplatePreset:
    id: str
    name: str
    category: TemplateCategory
    description: str
    placeholders: List[str]
    recommended_fonts: Dict[str, Dict]
    default_styles: Dict[str, Dict]

class TemplateLibrary:
    def __init__(self, templates_dir: str = "template_presets"):
        self.templates_dir = templates_dir
        self.presets: Dict[str, TemplatePreset] = {}
        self.load_presets()
    
    def load_presets(self):
        """Carga las plantillas predefinidas"""
        self.presets = {
            'certificado_academico': TemplatePreset(
                id='certificado_academico',
                name='Certificado Académico',
                category=TemplateCategory.ACADEMIC,
                description='Para certificados de cursos y programas académicos',
                placeholders=['{{NOMBRE}}', '{{CURSO}}'],
                recommended_fonts={
                    'NOMBRE': {'family': 'Times New Roman', 'size': 36, 'bold': True},
                    'CURSO': {'family': 'Arial', 'size': 18, 'bold': False}
                },
                default_styles={
                    '{{NOMBRE}}': {'size': 36, 'bold': True},
                    '{{CURSO}}': {'size': 18, 'bold': False}
                }
            ),
            
            'reconocimiento_participacion': TemplatePreset(
                id='reconocimiento_participacion',
                name='Reconocimiento por Participación',
                category=TemplateCategory.RECOGNITION,
                description='Para reconocer participación en eventos y conferencias',
                placeholders=['{{NOMBRE}}', '{{EVENTO}}'],
                recommended_fonts={
                    'NOMBRE': {'family': 'Arial', 'size': 32, 'bold': True},
                    'EVENTO': {'family': 'Calibri', 'size': 16, 'bold': False}
                },
                default_styles={
                    '{{NOMBRE}}': {'size': 32, 'bold': True},
                    '{{EVENTO}}': {'size': 16, 'bold': False}
                }
            ),
            
            'constancia_ponente': TemplatePreset(
                id='constancia_ponente',
                name='Constancia para Ponente',
                category=TemplateCategory.RECOGNITION,
                description='Para certificar participación como ponente',
                placeholders=['{{NOMBRE}}', '{{PONENCIA}}'],
                recommended_fonts={
                    'NOMBRE': {'family': 'Georgia', 'size': 28, 'bold': True},
                    'PONENCIA': {'family': 'Arial', 'size': 14, 'bold': False}
                },
                default_styles={
                    '{{NOMBRE}}': {'size': 28, 'bold': True},
                    '{{PONENCIA}}': {'size': 14, 'bold': False}
                }
            )
        }
    
    def get_preset(self, preset_id: str) -> Optional[TemplatePreset]:
        return self.presets.get(preset_id)
    
    def get_all_presets(self) -> List[TemplatePreset]:
        return list(self.presets.values())
    
    def save_custom_preset(self, name: str, font_map: Dict, placeholder_map: Dict):
        """Guarda una configuración personalizada como nuevo preset"""
        preset_id = name.lower().replace(' ', '_')
        custom_preset = TemplatePreset(
            id=preset_id,
            name=name,
            category=TemplateCategory.CERTIFICATE,
            description='Configuración personalizada guardada por el usuario',
            placeholders=list(placeholder_map.keys()),
            recommended_fonts=font_map,
            default_styles=font_map
        )
        
        self.presets[preset_id] = custom_preset