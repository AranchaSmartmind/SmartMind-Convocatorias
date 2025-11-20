"""
Procesador de Cronograma
Extrae información de fechas del cronograma
"""
import pandas as pd
import io
from datetime import datetime
from typing import Dict


class CronogramaProcessor:
    """Procesa archivos de cronograma para extraer fechas"""
    
    def __init__(self):
        self.df = None
        self.fecha_inicio = None
        self.fecha_fin = None
    
    def cargar_cronograma(self, file_bytes: bytes) -> Dict:
        """
        Carga y procesa el cronograma
        
        Args:
            file_bytes: Bytes del archivo Excel
            
        Returns:
            Dict con fechas extraídas
        """
        try:
            # Leer sin encabezados
            self.df = pd.read_excel(io.BytesIO(file_bytes), header=None)
            
            # Extraer fechas
            self._extraer_fechas()
            
            return {
                'fecha_inicio': self.fecha_inicio.strftime('%d/%m/%Y') if self.fecha_inicio else '',
                'fecha_fin': self.fecha_fin.strftime('%d/%m/%Y') if self.fecha_fin else '',
                'success': True
            }
            
        except Exception as e:
            raise Exception(f"Error al procesar cronograma: {str(e)}")
    
    def _extraer_fechas(self):
        """Extrae fechas de inicio y fin del cronograma"""
        
        if self.df is None:
            return
        
        fechas_inicio = []
        fechas_fin = []
        
        # Las fechas están en columnas 8-9 (fecha inicio/fin teoría)
        # Empezar desde fila 11 (después del encabezado en fila 10)
        
        for i in range(11, min(30, self.df.shape[0])):
            # Columna 8: Fecha inicio
            if pd.notna(self.df.iloc[i, 8]):
                val = self.df.iloc[i, 8]
                if isinstance(val, datetime):
                    fechas_inicio.append(val)
            
            # Columna 9: Fecha fin
            if pd.notna(self.df.iloc[i, 9]):
                val = self.df.iloc[i, 9]
                if isinstance(val, datetime):
                    fechas_fin.append(val)
        
        # Tomar primera fecha de inicio y última de fin
        if fechas_inicio:
            self.fecha_inicio = min(fechas_inicio)
        
        if fechas_fin:
            self.fecha_fin = max(fechas_fin)