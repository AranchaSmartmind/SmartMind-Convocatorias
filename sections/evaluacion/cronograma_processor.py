"""
Procesador de Cronograma
Extrae información de fechas Y MÓDULOS del cronograma
"""
import pandas as pd
import io
import re
from datetime import datetime
from typing import Dict, List


class CronogramaProcessor:
    """Procesa archivos de cronograma para extraer fechas y módulos"""
    
    def __init__(self):
        self.df = None
        self.fecha_inicio = None
        self.fecha_fin = None
        self.modulos = []
    
    def cargar_cronograma(self, file_bytes: bytes) -> Dict:
        """
        Carga y procesa el cronograma
        
        Args:
            file_bytes: Bytes del archivo Excel
            
        Returns:
            Dict con fechas y módulos extraídos
        """
        try:

            self.df = pd.read_excel(io.BytesIO(file_bytes), header=None)
            
            self._extraer_fechas()
            
            self._extraer_modulos(file_bytes)
            
            return {
                'fecha_inicio': self.fecha_inicio.strftime('%d/%m/%Y') if self.fecha_inicio else '',
                'fecha_fin': self.fecha_fin.strftime('%d/%m/%Y') if self.fecha_fin else '',
                'modulos': self.modulos,
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
        
        for i in range(11, min(30, self.df.shape[0])):
            if pd.notna(self.df.iloc[i, 8]):
                val = self.df.iloc[i, 8]
                if isinstance(val, datetime):
                    fechas_inicio.append(val)
            
            if pd.notna(self.df.iloc[i, 9]):
                val = self.df.iloc[i, 9]
                if isinstance(val, datetime):
                    fechas_fin.append(val)
        
        if fechas_inicio:
            self.fecha_inicio = min(fechas_inicio)
        
        if fechas_fin:
            self.fecha_fin = max(fechas_fin)
    
    def _extraer_modulos(self, file_bytes: bytes):
        """
        Extrae información de módulos profesionales desde la hoja 'Calculos_UF'
        
        Args:
            file_bytes: Bytes del archivo Excel
        """
        try:
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Calculos_UF', header=None)
            
            modulos = []
            
            for idx, row in df.iterrows():
                texto = str(row[0])
                
                if pd.notna(row[0]) and texto.startswith('MF'):
                    match = re.match(r'(MF\d+_\d+)\s+(.*)', texto)
                    if match:
                        codigo = match.group(1)
                        nombre = match.group(2).strip()
                        horas = None
                        for i in range(1, 10):
                            if idx + i < len(df):
                                siguiente = df.iloc[idx + i]
                                
                                if (pd.isna(siguiente[0]) and 
                                    pd.isna(siguiente[1]) and 
                                    pd.notna(siguiente[2])):
                                    try:
                                        horas = float(siguiente[2])
                                        break
                                    except:
                                        pass
                                
                                if pd.notna(siguiente[0]) and str(siguiente[0]).startswith('MF'):
                                    break
                        
                        if horas:
                            modulos.append({
                                'codigo': codigo,
                                'nombre': nombre,
                                'horas': int(horas)
                            })
            
            self.modulos = modulos
            print(f" Módulos extraídos del cronograma: {len(modulos)}")
            for mod in modulos:
                print(f"  - {mod['codigo']}: {mod['nombre']} ({mod['horas']}h)")
            
        except Exception as e:
            print(f" No se pudieron extraer módulos del cronograma: {e}")
            self.modulos = []