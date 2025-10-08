# analytics.py
import logging
from datetime import datetime

class Analytics:
    def __init__(self):
        self.setup_logging()
        self.stats = {
            'documents_generated': 0,
            'errors_encountered': 0,
            'average_generation_time': 0,
            'templates_used': set()
        }
    
    def log_generation(self, template_type, record_count, duration):
        """Registro de actividad para analytics"""
        self.stats['documents_generated'] += record_count
        self.stats['templates_used'].add(template_type)
        
        logging.info(f"Generated {record_count} documents from {template_type} in {duration:.2f}s")