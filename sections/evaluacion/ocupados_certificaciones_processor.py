"""
Procesador de datos para Certificaciones de Ocupados
Extrae información del PDF (justificante) y Excel (calificaciones)
"""

import re
import pandas as pd
import pdfplumber
from datetime import datetime
from typing import Dict, List, Optional


class CertificacionesOcupadosProcessor:
    """Procesa PDF y Excel para generar datos de certificaciones"""

    VALORES_FIJOS = {
        'director': 'PABLO LUIS LOBATO MURIENTE',
        'centro': 'INTERPROS NEXT GENERATION S.L.U.',
        'codigo_centro': '26615',
        'ciudad': 'AVILÉS',
        'direccion': 'C/Severo Ochoa 21 bajo Avilés'
    }
    
    def __init__(self, pdf_path: str, excel_path: str):
        """
        Inicializa el procesador
        
        Args:
            pdf_path: Ruta al PDF justificante
            excel_path: Ruta al Excel de calificaciones
        """
        self.pdf_path = pdf_path
        self.excel_path = excel_path
        self.datos_curso = {}
        self.alumnos = []
        
    def extraer_datos_pdf(self) -> Dict:
        """
        Extrae todos los datos necesarios del PDF
        
        Returns:
            Dict con datos del curso y alumnos
        """
        print(" Extrayendo datos del PDF...")
        
        with pdfplumber.open(self.pdf_path) as pdf:
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"

        expediente_match = re.search(r'Resumen comunicación\s+([^\n]+)', texto_completo)
        if expediente_match:
            expediente = expediente_match.group(1).strip().replace(' ', '')
        else:
            expediente = ""

        linea_grupo_match = re.search(r'(A\d+)\s+G\d+', texto_completo)
        codigo_curso = linea_grupo_match.group(1) if linea_grupo_match else ""
        modulo_match = re.search(r'\(([A-Z0-9_]+)\s+([^)]+)\)', texto_completo)
        
        if modulo_match:
            codigo_modulo = modulo_match.group(1)
            texto_modulo = modulo_match.group(2).strip()
            texto_modulo = texto_modulo.replace('\n', ' ')
            texto_modulo = texto_modulo.replace('?', '-')
            texto_modulo = texto_modulo.replace('Grupo ', '')
            texto_modulo = re.sub(r'\s+', ' ', texto_modulo)
            texto_modulo = texto_modulo.rstrip('-').strip()
            nombre_modulo = codigo_modulo + " " + texto_modulo
        else:
            codigo_modulo = ""
            nombre_modulo = ""
        
        nivel = codigo_modulo.split('_')[-1] if '_' in codigo_modulo else ""
        
        fecha_inicio_match = re.search(r'Fecha Inicio\s+(\d{2}/\d{2}/\d{4})', texto_completo)
        fecha_fin_match = re.search(r'Fecha Fin\s+(\d{2}/\d{2}/\d{4})', texto_completo)
        
        fecha_inicio = fecha_inicio_match.group(1) if fecha_inicio_match else ""
        fecha_fin = fecha_fin_match.group(1) if fecha_fin_match else ""

        horas_match = re.search(r'Horas\s+Modalidad.*?(\d+[,.]?\d*)\s+', texto_completo, re.DOTALL)
        horas_raw = horas_match.group(1) if horas_match else "0"
        horas = str(int(float(horas_raw.replace(',', '.'))))

        participantes = self._extraer_participantes(texto_completo)
        
        self.datos_curso = {
            'expediente': expediente,
            'codigo_curso': codigo_curso,
            'codigo_modulo': codigo_modulo,
            'nombre_modulo': nombre_modulo,
            'nivel': nivel,
            'horas': horas,
            'fecha_inicio': fecha_inicio,
            'fecha_fin': fecha_fin,
            **self.VALORES_FIJOS
        }
        
        self.alumnos = participantes
        
        print(f" Curso: {codigo_curso}")
        print(f" Módulo: {codigo_modulo}")
        print(f" Alumnos encontrados: {len(participantes)}")
        
        return self.datos_curso
    
    def _extraer_participantes(self, texto: str) -> List[Dict]:
        """
        Extrae lista de participantes del PDF
        
        Args:
            texto: Texto completo del PDF
            
        Returns:
            Lista de diccionarios con datos de alumnos
        """
        participantes = []
        
        patron = r'([A-ZÁÉÍÓÚÑ\s]+,\s+[A-ZÁÉÍÓÚÑ\s]+?)\s+\d{2}/\d{2}/\d{4}\s+([0-9]{8}[A-Z])\s+\d+\s+OCUPAD[AO]'
        
        matches = re.finditer(patron, texto)
        
        for match in matches:
            nombre_completo = match.group(1).strip()
            dni = match.group(2).strip()

            if ',' in nombre_completo:
                partes = nombre_completo.split(',')
                apellidos = partes[0].strip()
                nombre = partes[1].strip()
                nombre_formateado = f"{nombre} {apellidos}"
            else:
                nombre_formateado = nombre_completo
            
            participantes.append({
                'nombre': nombre_formateado,
                'dni': dni
            })
        
        return participantes
    
    def extraer_calificaciones_excel(self) -> Dict[str, str]:
        """
        Extrae calificaciones del Excel
        
        Returns:
            Dict con DNI -> calificación (formato S-9, NS, etc)
        """
        print("\n Extrayendo calificaciones del Excel...")

        df = pd.read_excel(self.excel_path, sheet_name=0, header=None)
        
        calificaciones = {}

        for alumno in self.alumnos:
            dni = alumno['dni']

            alumno_fila = None
            for idx, row in df.iterrows():
                if dni in str(row[0]):
                    alumno_fila = idx
                    break
            
            if alumno_fila is None:
                print(f" No se encontró calificación para {alumno['nombre']} ({dni})")
                calificaciones[dni] = "S-0"
                continue
            
            nota_final = None
            estado = None
            
            for offset in range(0, 20):
                fila_buscar = alumno_fila + offset
                if fila_buscar >= len(df):
                    break
                
                val_col21 = str(df.iloc[fila_buscar, 21])
                val_col24 = str(df.iloc[fila_buscar, 24])
                
                if 'APTO' in val_col21 or 'APTO' in val_col24:
                    if 'NO APTO' in val_col21 or 'NO APTO' in val_col24:
                        estado = 'NO APTO'
                    else:
                        estado = 'APTO'

                    nota_match = re.search(r'(\d+\.?\d*)', val_col21)
                    if nota_match:
                        nota_final = float(nota_match.group(1))
                    
                    break

            if estado == 'APTO' and nota_final is not None:
                nota_redondeada = round(nota_final)
                calificacion = f"S-{nota_redondeada}"
            elif estado == 'NO APTO':
                calificacion = "NS"
            else:
                calificacion = "S-0"
            
            calificaciones[dni] = calificacion
            print(f" {alumno['nombre']}: {calificacion}")
        
        return calificaciones
    
    def combinar_datos(self) -> List[Dict]:
        """
        Combina todos los datos para generar la lista completa
        
        Returns:
            Lista de diccionarios, uno por alumno con todos los datos
        """
        print("\n🔗 Combinando datos...")
        
        self.extraer_datos_pdf()

        calificaciones = self.extraer_calificaciones_excel()

        datos_completos = []
        
        for alumno in self.alumnos:
            dni = alumno['dni']
            
            datos_alumno = {
                'expediente': self.datos_curso['expediente'],
                'codigo_curso': self.datos_curso['codigo_curso'],
                'codigo_modulo': self.datos_curso['codigo_modulo'],
                'nombre_modulo': self.datos_curso['nombre_modulo'],
                'nivel': self.datos_curso['nivel'],
                'horas': self.datos_curso['horas'],
                'fecha_inicio': self.datos_curso['fecha_inicio'],
                'fecha_fin': self.datos_curso['fecha_fin'],

                'director': self.datos_curso['director'],
                'centro': self.datos_curso['centro'],
                'codigo_centro': self.datos_curso['codigo_centro'],
                'ciudad': self.datos_curso['ciudad'],
                'direccion': self.datos_curso['direccion'],

                'nombre_alumno': alumno['nombre'],
                'dni_alumno': dni,
                'calificacion': calificaciones.get(dni, 'S-0'),

                'firma': ''
            }
            
            datos_completos.append(datos_alumno)
        
        print(f"\n Datos completos para {len(datos_completos)} alumnos")
        
        return datos_completos


def procesar_certificaciones(pdf_path: str, excel_path: str) -> List[Dict]:
    """
    Función principal para procesar todos los datos
    
    Args:
        pdf_path: Ruta al PDF justificante
        excel_path: Ruta al Excel de calificaciones
        
    Returns:
        Lista de diccionarios con datos completos por alumno
    """
    processor = CertificacionesOcupadosProcessor(pdf_path, excel_path)
    return processor.combinar_datos()