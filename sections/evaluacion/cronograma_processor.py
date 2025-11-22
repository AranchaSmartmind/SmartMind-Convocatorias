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
            # Leer sin encabezados
            self.df = pd.read_excel(io.BytesIO(file_bytes), header=None)
            
            # Extraer fechas
            self._extraer_fechas()
            
            # NUEVO: Extraer módulos
            self._extraer_modulos(file_bytes)
            
            return {
                'fecha_inicio': self.fecha_inicio.strftime('%d/%m/%Y') if self.fecha_inicio else '',
                'fecha_fin': self.fecha_fin.strftime('%d/%m/%Y') if self.fecha_fin else '',
                'modulos': self.modulos,  # <-- NUEVO
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
    
    def _extraer_modulos(self, file_bytes: bytes):
        """
        Extrae información de módulos profesionales desde la hoja 'Calculos_UF'
        
        Args:
            file_bytes: Bytes del archivo Excel
        """
        try:
            # Leer la hoja 'Calculos_UF' sin encabezados
            df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='Calculos_UF', header=None)
            
            modulos = []
            
            # Buscar líneas que empiecen con "MF" (código de módulo)
            for idx, row in df.iterrows():
                texto = str(row[0])
                
                # Si la celda contiene un código de módulo (ej: MF0969_1)
                if pd.notna(row[0]) and texto.startswith('MF'):
                    # Extraer código y nombre del módulo
                    match = re.match(r'(MF\d+_\d+)\s+(.*)', texto)
                    if match:
                        codigo = match.group(1)
                        nombre = match.group(2).strip()
                        
                        # Buscar las horas en las filas siguientes
                        # Las horas suelen estar unas filas más abajo en la columna 2 (índice 2)
                        horas = None
                        for i in range(1, 10):  # Buscar en las siguientes 10 filas
                            if idx + i < len(df):
                                siguiente = df.iloc[idx + i]
                                
                                # Si encontramos una celda con valor numérico en columna 2
                                # y las columnas 0 y 1 están vacías, es probablemente el total de horas
                                if (pd.isna(siguiente[0]) and 
                                    pd.isna(siguiente[1]) and 
                                    pd.notna(siguiente[2])):
                                    try:
                                        horas = float(siguiente[2])
                                        break
                                    except:
                                        pass
                                
                                # Si encontramos otro módulo, detenemos la búsqueda
                                if pd.notna(siguiente[0]) and str(siguiente[0]).startswith('MF'):
                                    break
                        
                        # Solo agregar si encontramos las horas
                        if horas:
                            modulos.append({
                                'codigo': codigo,
                                'nombre': nombre,
                                'horas': int(horas)
                            })
            
            self.modulos = modulos
            print(f"✓ Módulos extraídos del cronograma: {len(modulos)}")
            for mod in modulos:
                print(f"  - {mod['codigo']}: {mod['nombre']} ({mod['horas']}h)")
            
        except Exception as e:
            print(f"⚠ No se pudieron extraer módulos del cronograma: {e}")
            self.modulos = []