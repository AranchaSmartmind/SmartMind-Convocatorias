"""
Procesador de Excel REAL - Formato empresa
Extrae datos de los archivos reales de control de asistencias
"""
import pandas as pd
import numpy as np
from typing import Dict, List, Optional
import io
import re


class ExcelProcessorReal:
    """Procesa archivos Excel con el formato real de la empresa"""
    FILA_MODULOS = 5
    FILA_ENCABEZADO = 9
    COL_ALUMNO = 1
    COL_DNI = 2
    
    MODULOS_CONFIG = [
        {
            'codigo': 'MF0969_1',
            'nombre': 'Técnicas administrativas básicas de oficina',
            'horas_totales': 165,
            'col_horas_asistidas': 10
        },
        {
            'codigo': 'MF0970_1', 
            'nombre': 'Operaciones básicas de comunicación',
            'horas_totales': 133,
            'col_horas_asistidas': 17
        },
        {
            'codigo': 'MF0971_1',
            'nombre': 'Reproducción y archivo (TRANSV.)',
            'horas_totales': 132,
            'col_horas_asistidas': 23
        }
    ]
    
    def __init__(self):
        self.df_asistencias = None
        self.curso_codigo = None
        self.curso_nombre = None
        self.alumnos_data = []
        
    def cargar_asistencias(self, file_bytes: bytes) -> Dict:
        """
        Carga el archivo de asistencias real
        
        Args:
            file_bytes: Bytes del archivo Excel
            
        Returns:
            Dict con datos procesados
        """
        try:
            self.df_asistencias = pd.read_excel(io.BytesIO(file_bytes), header=None)
            self._extraer_info_curso()
            self._extraer_alumnos()
            estadisticas = self._calcular_estadisticas()
            
            return {
                'curso_codigo': self.curso_codigo,
                'curso_nombre': self.curso_nombre,
                'alumnos': self.alumnos_data,
                'estadisticas_individuales': self.alumnos_data,
                'estadisticas_grupales': estadisticas,
                'success': True
            }
            
        except Exception as e:
            raise Exception(f"Error al procesar asistencias: {str(e)}")
    
    def _extraer_info_curso(self):
        """Extrae código y nombre del curso"""
        df = self.df_asistencias
        
        for i in range(5):
            for col in range(df.shape[1]):
                val = str(df.iloc[i, col]).lower() if pd.notna(df.iloc[i, col]) else ''
                
                if 'curso:' in val:
                    if col + 1 < df.shape[1]:
                        self.curso_codigo = str(df.iloc[i, col + 1])
                
                if 'nombre:' in val:
                    if col + 1 < df.shape[1]:
                        nombre_completo = str(df.iloc[i, col + 1])
                        self.curso_nombre = nombre_completo[:50] + '...' if len(nombre_completo) > 50 else nombre_completo
        
        if not self.curso_codigo:
            self.curso_codigo = "2024/1339"
        if not self.curso_nombre:
            self.curso_nombre = "OPERACIONES AUXILIARES DE SERVICIOS ADMINISTRATIVOS Y GENERALES"
    
    def _extraer_alumnos(self):
        """Extrae datos de todos los alumnos"""
        df = self.df_asistencias
        self.alumnos_data = []
        
        for i in range(self.FILA_ENCABEZADO + 1, df.shape[0]):
            nombre = df.iloc[i, self.COL_ALUMNO]
            dni = df.iloc[i, self.COL_DNI]
            
            if pd.notna(nombre) and isinstance(nombre, str) and len(nombre.strip()) > 3:
                alumno = {
                    'nombre': str(nombre).strip(),
                    'dni': str(dni).strip() if pd.notna(dni) else 'N/A',
                    'modulos': []
                }
                
                for modulo_cfg in self.MODULOS_CONFIG:
                    col_horas = modulo_cfg['col_horas_asistidas']
                    horas_valor = df.iloc[i, col_horas]
                    horas_asistidas = 0
                    if pd.notna(horas_valor):
                        if isinstance(horas_valor, (int, float)):
                            horas_asistidas = float(horas_valor)
                        elif isinstance(horas_valor, str):
                            if horas_valor.lower() == 'exento':
                                horas_asistidas = modulo_cfg['horas_totales']
                            else:
                                try:
                                    horas_asistidas = float(horas_valor)
                                except:
                                    horas_asistidas = 0
                    
                    modulo_alumno = {
                        'codigo': modulo_cfg['codigo'],
                        'nombre': modulo_cfg['nombre'],
                        'horas_totales': modulo_cfg['horas_totales'],
                        'horas_asistidas': round(horas_asistidas, 2)
                    }
                    alumno['modulos'].append(modulo_alumno)
                
                total_horas_totales = sum(m['horas_totales'] for m in alumno['modulos'])
                total_horas_asistidas = sum(m['horas_asistidas'] for m in alumno['modulos'])
                alumno['total_horas'] = total_horas_totales
                alumno['total_asistidas'] = total_horas_asistidas
                alumno['porcentaje_asistencia'] = round((total_horas_asistidas / total_horas_totales * 100), 2) if total_horas_totales > 0 else 0
                
                self.alumnos_data.append(alumno)
        
        return self.alumnos_data
    
    def _calcular_estadisticas(self) -> Dict:
        """Calcula estadísticas grupales"""
        if not self.alumnos_data:
            return {
                'total_alumnos': 0,
                'porcentaje_asistencia_media': 0,
                'alumnos_con_mas_80': 0,
                'alumnos_con_menos_80': 0
            }
        
        total = len(self.alumnos_data)
        porcentajes = [a['porcentaje_asistencia'] for a in self.alumnos_data]
        
        return {
            'total_alumnos': total,
            'porcentaje_asistencia_media': round(sum(porcentajes) / total, 2) if total > 0 else 0,
            'porcentaje_asistencia_minimo': round(min(porcentajes), 2) if porcentajes else 0,
            'porcentaje_asistencia_maximo': round(max(porcentajes), 2) if porcentajes else 0,
            'alumnos_con_mas_80': sum(1 for p in porcentajes if p >= 80),
            'alumnos_con_menos_80': sum(1 for p in porcentajes if p < 80)
        }
    
    def obtener_alumno(self, nombre: str) -> Optional[Dict]:
        """Obtiene datos de un alumno específico"""
        for alumno in self.alumnos_data:
            if alumno['nombre'] == nombre:
                return alumno
        return None
    
    def obtener_todos_alumnos(self) -> List[Dict]:
        """Devuelve lista de todos los alumnos"""
        return self.alumnos_data
    
    def exportar_a_excel(self, alumno_nombre: str = None) -> bytes:
        """
        Exporta datos a Excel
        
        Args:
            alumno_nombre: Si se especifica, exporta solo ese alumno. Si no, exporta todos.
        
        Returns:
            bytes del archivo Excel
        """
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
        
        wb = Workbook()
        ws = wb.active
        ws.title = "Informe Evaluación"
        
        ws.column_dimensions['A'].width = 45
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 15
        ws.column_dimensions['D'].width = 15
        
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        ws['A1'] = "INFORME DE EVALUACIÓN INDIVIDUALIZADO GRADO C"
        ws['A1'].font = Font(bold=True, size=14)
        ws.merge_cells('A1:D1')
        ws['A1'].alignment = Alignment(horizontal='center')
        
        row = 3
        
        ws[f'A{row}'] = "Número de expediente:"
        ws[f'B{row}'] = "E-2024-100454"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Certificado profesional:"
        ws[f'B{row}'] = self.curso_nombre
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Código:"
        ws[f'B{row}'] = "ADGG0408"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Centro formativo:"
        ws[f'B{row}'] = "INTERPROS NEXT GENERATION SLU"
        ws[f'A{row}'].font = Font(bold=True)
        row += 1
        
        ws[f'A{row}'] = "Código:"
        ws[f'B{row}'] = "26615"
        ws[f'A{row}'].font = Font(bold=True)
        row += 2
        
        alumnos_exportar = []
        if alumno_nombre:
            alumno = self.obtener_alumno(alumno_nombre)
            if alumno:
                alumnos_exportar = [alumno]
        else:
            alumnos_exportar = self.alumnos_data
        
        for idx, alumno in enumerate(alumnos_exportar):
            if idx > 0:
                row += 3
            
            ws[f'A{row}'] = "El/la alumno/a:"
            ws[f'B{row}'] = alumno['nombre']
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
            
            ws[f'A{row}'] = "con DNI/NIE/Pasaporte:"
            ws[f'B{row}'] = alumno['dni']
            ws[f'A{row}'].font = Font(bold=True)
            row += 2
            
            ws[f'A{row}'] = "Módulos"
            ws[f'B{row}'] = "Código"
            ws[f'C{row}'] = "Horas"
            ws[f'D{row}'] = "Horas de asistencia"
            
            for col in ['A', 'B', 'C', 'D']:
                cell = ws[f'{col}{row}']
                cell.fill = header_fill
                cell.font = header_font
                cell.border = border
                cell.alignment = Alignment(horizontal='center')
            
            row += 1
            
            for modulo in alumno['modulos']:
                ws[f'A{row}'] = modulo['nombre']
                ws[f'B{row}'] = modulo['codigo']
                ws[f'C{row}'] = modulo['horas_totales']
                ws[f'D{row}'] = modulo['horas_asistidas']
                
                for col in ['A', 'B', 'C', 'D']:
                    ws[f'{col}{row}'].border = border
                
                row += 1
            
            ws[f'A{row}'] = "TOTAL"
            ws[f'C{row}'] = alumno['total_horas']
            ws[f'D{row}'] = alumno['total_asistidas']
            ws[f'A{row}'].font = Font(bold=True)
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{row}'].border = border
            row += 1
            
            ws[f'A{row}'] = "Porcentaje de asistencia:"
            ws[f'B{row}'] = f"{alumno['porcentaje_asistencia']}%"
            ws[f'A{row}'].font = Font(bold=True)
            row += 1
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        return output.getvalue()