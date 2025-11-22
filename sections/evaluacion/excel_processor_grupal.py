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
        curso_nombre_completo = ws_asistencia['C2'].value
        
        # Limpiar paréntesis del nombre del curso
        import re
        curso_nombre = re.sub(r'\s*\([^)]*\)', '', curso_nombre_completo).strip()
        
        # Extraer código del certificado entre paréntesis (para información)
        match_certificado = re.search(r'\(([A-Z]{4}\d{4})\)', curso_nombre_completo)
        codigo_certificado = match_certificado.group(1) if match_certificado else ''
        
        print(f"   ✓ Curso: {curso_codigo}")
        print(f"   ✓ Certificado: {codigo_certificado}")
        print(f"   ✓ Nombre limpio: {curso_nombre}")
        
        # El nivel se extraerá después de identificar los módulos
        
        # Obtener encabezados
        headers = [cell.value for cell in ws_resumen[1]]
        
        # Inicializar nivel con valor por defecto
        nivel = '1'
        
        # Identificar módulos MF
        modulos_info = []
        for i, header in enumerate(headers):
            if header and 'MF' in str(header) and '_' in str(header):
                nombre_modulo = header.strip()
                
                # Obtener horas estándar del módulo (según catálogo oficial)
                # Diccionario de horas por módulo (certificado ADGG0408)
                horas_modulos = {
                    'MF0969_1': 90,   # Técnicas administrativas básicas de oficina
                    'MF0970_1': 120,  # Operaciones básicas de comunicación
                    'MF0971_1': 90,   # Reproducción y archivo
                }
                
                horas = horas_modulos.get(nombre_modulo, 0)
                
                modulos_info.append({
                    'columna': i,
                    'nombre': nombre_modulo,
                    'horas': horas,
                    'denominacion': self._get_denominacion_modulo(nombre_modulo)
                })
        
        print(f"   ✓ Módulos encontrados: {len(modulos_info)}")
        for mod in modulos_info:
            print(f"      - {mod['nombre']} ({mod['horas']}h)")
        
        # Extraer nivel del primer módulo (número después de _)
        # Ejemplo: MF0969_1 → nivel = 1, MF0969_2 → nivel = 2
        if modulos_info:
            primer_modulo = modulos_info[0]['nombre']
            match_nivel = re.search(r'_(\d+)', primer_modulo)
            if match_nivel:
                nivel = match_nivel.group(1)
                print(f"   ✓ Nivel del curso (extraído de {primer_modulo}): {nivel}")
        
        # Cargar alumnos
        alumnos = self._cargar_alumnos(ws_resumen, modulos_info)
        
        wb.close()
        
        return {
            'curso_codigo': curso_codigo,
            'curso_nombre': curso_nombre,
            'codigo_certificado': codigo_certificado,  # Código del certificado (ej: ADGG0408)
            'nivel': nivel,                     # Nivel del curso (1, 2 o 3)
            'fecha_inicio': '20/03/2025',       # Puedes cambiar esto
            'fecha_fin': '27/06/2025',          # Puedes cambiar esto
            'modulos_info': modulos_info,       # Información de módulos para encabezados
            'total_alumnos': len(alumnos),      # Total de alumnos en el Excel
            'alumnos': alumnos
        }
    
    def _get_denominacion_modulo(self, codigo_modulo: str) -> str:
        """Obtiene la denominación completa del módulo"""
        denominaciones = {
            'MF0969_1': 'TÉCNICAS ADMINISTRATIVAS BÁSICAS DE OFICINA',
            'MF0970_1': 'OPERACIONES BÁSICAS DE COMUNICACIÓN',
            'MF0971_1': 'REPRODUCCIÓN Y ARCHIVO',
        }
        return denominaciones.get(codigo_modulo, '')
    
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
        - S-10, S-9, S-8 (Superado con nota)
        - NS-0, NS-4 (No Superado con nota)
        - CO-5 (para convalidaciones/exentos)
        """
        if nota is None:
            return ''
        
        nota_int = int(nota)
        
        # Convalidado/Exento = CO-5
        if tipo and 'convalida' in tipo:
            return 'CO-5'
        
        # Superado/No Superado con nota
        if nota >= 5:
            return f'S-{nota_int}'
        else:
            return f'NS-{nota_int}'