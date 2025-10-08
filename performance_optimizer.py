# performance_optimizer.py
import gc
import os

class PerformanceOptimizer:
    def __init__(self):
        pass
    
    def optimize_memory(self):
        """Optimiza el uso de memoria"""
        gc.collect()
    
    def clear_caches(self):
        """Limpia caches internos"""
        try:
            import fitz
            fitz.TOOLS.mupdf_clean()
        except:
            pass