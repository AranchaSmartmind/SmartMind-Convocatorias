"""
Procesador de datos para Certificaciones de Ocupados
Extrae información del PDF (justificante) y Excel (calificaciones)
VERSIÓN FINAL - Búsqueda dinámica sin posiciones fijas
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
        print("Extrayendo datos del PDF...")
        
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
        
        print(f"Curso: {codigo_curso}")
        print(f"Módulo: {codigo_modulo}")
        print(f"Alumnos encontrados: {len(participantes)}")
        
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
    
    def _buscar_columna_puntuacion(self, df: pd.DataFrame, fila_inicio: int, rango_filas: int = 10) -> Optional[int]:
        """
        Busca dinámicamente la columna que contiene 'PUNTUACIÓN FINAL' o 'DEL MÓDULO'
        
        Args:
            df: DataFrame del Excel
            fila_inicio: Fila desde donde empezar a buscar
            rango_filas: Cuántas filas buscar
            
        Returns:
            Índice de la columna o None
        """
        for offset in range(rango_filas):
            fila = fila_inicio + offset
            if fila >= len(df):
                break
            
            for col in range(len(df.columns)):
                celda = str(df.iloc[fila, col])
                if 'PUNTUACIÓN FINAL' in celda or 'DEL MÓDULO' in celda:
                    print(f"  Columna de puntuación encontrada: {col} (fila {fila})")
                    return col
        
        return None
    
    def extraer_calificaciones_excel(self) -> Dict[str, str]:
        """
        Extrae calificaciones del Excel
        Búsqueda dinámica de la PUNTUACIÓN FINAL en formato (X.X)
        
        Returns:
            Dict con DNI -> calificación (formato S-9, NS, etc)
        """
        print("\nExtrayendo calificaciones del Excel...")

        df = pd.read_excel(self.excel_path, sheet_name=0, header=None)
        
        calificaciones = {}

        for alumno in self.alumnos:
            dni = alumno['dni']

            # Buscar la fila donde aparece el DNI del alumno
            alumno_fila = None
            for idx, row in df.iterrows():
                # Buscar DNI en cualquier columna de la fila
                fila_texto = ' '.join([str(cell) for cell in row if pd.notna(cell)])
                if dni in fila_texto:
                    alumno_fila = idx
                    print(f"\n{alumno['nombre']} ({dni}) encontrado en fila {idx}")
                    break
            
            if alumno_fila is None:
                print(f"  ADVERTENCIA: No se encontró en Excel")
                calificaciones[dni] = "S-0"
                continue
            
            # Buscar dinámicamente la columna de PUNTUACIÓN FINAL
            columna_puntuacion = self._buscar_columna_puntuacion(df, alumno_fila, rango_filas=15)
            
            nota_final = None
            estado = None
            
            # ESTRATEGIA 1: Si encontramos la columna específica, buscar ahí
            if columna_puntuacion is not None:
                print(f"  Buscando en columna {columna_puntuacion}...")
                
                for offset in range(1, 25):
                    fila_buscar = alumno_fila + offset
                    if fila_buscar >= len(df):
                        break
                    
                    celda = df.iloc[fila_buscar, columna_puntuacion]
                    
                    if pd.notna(celda):
                        celda_str = str(celda).strip()
                        
                        # Buscar patrón (X.X) o (X)
                        match_puntuacion = re.search(r'\((\d+\.?\d*)\)', celda_str)
                        
                        if match_puntuacion:
                            nota_final = float(match_puntuacion.group(1))
                            print(f"    Fila {fila_buscar}: '{celda_str}' -> Nota: {nota_final}")
                            
                            # Verificar estado APTO/NO APTO en filas cercanas
                            for check_offset in range(-3, 4):
                                check_fila = fila_buscar + check_offset
                                if 0 <= check_fila < len(df):
                                    # Buscar en varias columnas cercanas
                                    for col_offset in range(-3, 4):
                                        col_check = columna_puntuacion + col_offset
                                        if 0 <= col_check < len(df.columns):
                                            celda_check = str(df.iloc[check_fila, col_check])
                                            if 'NO APTO' in celda_check:
                                                estado = 'NO APTO'
                                                break
                                            elif 'APTO' in celda_check and estado != 'NO APTO':
                                                estado = 'APTO'
                                if estado:
                                    break
                            
                            if not estado:
                                estado = 'APTO'
                            
                            break
            
            # ESTRATEGIA 2: Si no encontró en columna específica, búsqueda amplia
            if nota_final is None:
                print(f"  Búsqueda amplia en todas las columnas...")
                
                for offset in range(1, 30):
                    fila_buscar = alumno_fila + offset
                    if fila_buscar >= len(df):
                        break
                    
                    # Concatenar toda la fila
                    fila_completa = []
                    for col in range(len(df.columns)):
                        celda = df.iloc[fila_buscar, col]
                        if pd.notna(celda):
                            fila_completa.append(str(celda))
                    
                    fila_texto = ' '.join(fila_completa)
                    
                    # Buscar líneas que contengan indicadores clave
                    if any(keyword in fila_texto for keyword in ['PUNTUACIÓN FINAL', 'DEL MÓDULO', 'CALIFICACIÓN FINAL']):
                        # Extraer todos los números entre paréntesis
                        numeros = re.findall(r'\((\d+\.?\d*)\)', fila_texto)
                        
                        if numeros:
                            # Filtrar números que parecen notas (entre 0 y 10)
                            notas_validas = [float(n) for n in numeros if 0 <= float(n) <= 10]
                            
                            if notas_validas:
                                # Tomar el último (suele ser la puntuación final)
                                nota_final = notas_validas[-1]
                                print(f"    Fila {fila_buscar}: Números encontrados {numeros}")
                                print(f"    Nota seleccionada: {nota_final}")
                                
                                # Determinar estado
                                if 'NO APTO' in fila_texto:
                                    estado = 'NO APTO'
                                elif 'APTO' in fila_texto:
                                    estado = 'APTO'
                                
                                break
                    
                    # También buscar líneas con APTO y números
                    elif 'APTO' in fila_texto and '(' in fila_texto:
                        numeros = re.findall(r'\((\d+\.?\d*)\)', fila_texto)
                        notas_validas = [float(n) for n in numeros if 0 <= float(n) <= 10]
                        
                        if notas_validas:
                            nota_final = notas_validas[-1]
                            print(f"    Fila {fila_buscar}: APTO con nota {nota_final}")
                            
                            if 'NO APTO' in fila_texto:
                                estado = 'NO APTO'
                            else:
                                estado = 'APTO'
                            break

            # Generar calificación final
            if estado == 'NO APTO':
                calificacion = "NS"
                print(f"  RESULTADO: {calificacion} (NO APTO)")
            elif nota_final is not None:
                # Redondear sin decimales
                nota_redondeada = round(nota_final)
                calificacion = f"S-{nota_redondeada}"
                print(f"  RESULTADO: {calificacion} (nota: {nota_final} -> {nota_redondeada})")
            else:
                calificacion = "S-0"
                print(f"  ERROR: No se pudo extraer calificación, asignando S-0")
            
            calificaciones[dni] = calificacion
        
        return calificaciones
    
    def combinar_datos(self) -> List[Dict]:
        """
        Combina todos los datos para generar la lista completa
        
        Returns:
            Lista de diccionarios, uno por alumno con todos los datos
        """
        print("\nCombinando datos...")
        
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
        
        print(f"\nDatos completos para {len(datos_completos)} alumnos")
        
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