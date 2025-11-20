#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de Excel para extraer datos de alumnos
Compatible con word_generator_grupal.py
"""

import openpyxl
import re


class ExcelProcessor:
    """Procesa el Excel y extrae datos para el acta"""
    
    def __init__(self, excel_path):
        self.excel_path = excel_path
        
    def cargar_datos(self):
        """Carga datos desde el Excel"""
        print("   📊 Leyendo hojas del Excel...")
        
        wb = openpyxl.load_workbook(self.excel_path)
        ws_resumen = wb['RESUMEN']
        ws_asistencia = wb['ASISTENCIA']
        
        # Información del curso
        curso_codigo = ws_asistencia['C1'].value
        curso_nombre = ws_asistencia['C2'].value
        
        print(f"   ✓ Curso: {curso_codigo}")
        
        # Obtener encabezados
        headers = [cell.value for cell in ws_resumen[1]]
        
        # Identificar módulos MF
        modulos_info = []
        for i, header in enumerate(headers):
            if header and 'MF' in str(header) and '_' in str(header):
                modulos_info.append({
                    'columna': i,
                    'nombre': header.strip()
                })
        
        print(f"   ✓ Módulos encontrados: {len(modulos_info)}")
        for mod in modulos_info:
            print(f"      - {mod['nombre']}")
        
        # Cargar alumnos
        alumnos = self._cargar_alumnos(ws_resumen, modulos_info)
        
        wb.close()
        
        return {
            'curso_codigo': curso_codigo,
            'curso_nombre': curso_nombre,
            'fecha_inicio': '20/03/2025',  # Puedes cambiar esto
            'fecha_fin': '20/11/2025',      # Puedes cambiar esto
            'alumnos': alumnos
        }
    
    def _cargar_alumnos(self, ws_resumen, modulos_info):
        """Carga información de alumnos"""
        print("   📋 Procesando alumnos...")
        
        alumnos = []
        
        for row_idx in range(2, ws_resumen.max_row + 1):
            row = list(ws_resumen[row_idx])
            
            # Si no hay nombre, saltar
            if not row[1].value:
                continue
            
            alumno = {
                'dni': row[2].value or '',
                'nombre': row[1].value or '',
                'asistencia': row[5].value,  # Columna 6 (índice 5)
                'modulos': []
            }
            
            # Cargar calificaciones de módulos
            for mod_info in modulos_info:
                valor = row[mod_info['columna']].value
                nota, tipo = self._extraer_nota(valor)
                
                alumno['modulos'].append({
                    'nombre': mod_info['nombre'],
                    'nota': nota,
                    'tipo': tipo,
                    'calificacion': self._formatear_calificacion(nota, tipo)
                })
            
            alumnos.append(alumno)
        
        print(f"   ✓ {len(alumnos)} alumnos cargados")
        
        # Mostrar resumen de primeros 3 alumnos
        print("\n   📝 Resumen (primeros 3 alumnos):")
        for i, alumno in enumerate(alumnos[:3], 1):
            print(f"      {i}. {alumno['nombre']}")
            print(f"         DNI: {alumno['dni']}")
            asist_pct = f"{alumno['asistencia']*100:.1f}%" if alumno['asistencia'] else "N/A"
            print(f"         Asistencia: {asist_pct}")
            for mod in alumno['modulos']:
                print(f"         {mod['nombre']}: {mod['calificacion']}")
        
        if len(alumnos) > 3:
            print(f"      ... y {len(alumnos)-3} alumnos más")
        
        return alumnos
    
    def _extraer_nota(self, valor):
        """
        Extrae nota de formato '9 / Superado'
        Retorna: (nota_numerica, tipo)
        """
        if not valor:
            return None, None
        
        texto = str(valor).strip()
        
        # Patrón: "9 / Superado" o "5 / Convalida"
        match = re.match(r'^(\d+(?:\.\d+)?)\s*/\s*(.+)$', texto)
        if match:
            nota = float(match.group(1))
            tipo = match.group(2).strip().lower()
            return nota, tipo
        
        return None, None
    
    def _formatear_calificacion(self, nota, tipo):
        """
        Formatea calificación según especificaciones:
        - S / NS (para módulos)
        - CO-5 (para convalidaciones/exentos)
        """
        if nota is None:
            return ''
        
        # Convalidado/Exento = CO-5
        if tipo and 'convalida' in tipo:
            return 'CO-5'
        
        # Superado/No Superado (sin nota, solo letra)
        if nota >= 5:
            return 'S'
        else:
            return 'NS'
        