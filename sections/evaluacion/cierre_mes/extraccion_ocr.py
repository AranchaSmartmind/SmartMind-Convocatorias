"""
Módulo de extracción de datos con OCR
Procesa PDFs escaneados usando Tesseract
"""

import re
from datetime import datetime, timedelta
from collections import defaultdict
from pdf2image import convert_from_path
import pytesseract
import os
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# AGREGAR ESTAS DOS LÍNEAS:
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

def extraer_texto_con_ocr(pdf_path, dpi=300):
    """
    Convierte PDF a imágenes y extrae texto con OCR
    
    Args:
        pdf_path: Ruta al archivo PDF
        dpi: Resolución de la imagen (300 recomendado)
    
    Returns:
        Lista de textos extraídos por página
    """
    try:
        images = convert_from_path(pdf_path, dpi=dpi)
        textos_por_pagina = []
        
        for image in images:
            texto = pytesseract.image_to_string(image, lang='spa')
            textos_por_pagina.append(texto)
        
        return textos_por_pagina
    except Exception as e:
        print(f"Error en OCR: {e}")
        return []

def extraer_nombre_alumno_ocr(texto):
    """
    Extrae el nombre del alumno del texto OCR
    
    Args:
        texto: Texto extraído por OCR
    
    Returns:
        Nombre del alumno o None
    """
    patrones = [
        r'DATOS\s+DEL\s+ALUMNO.*?Nombre[:\s]+([A-ZÁÉÍÓÚÑ\s]+?)(?:\n|NIF|DNI)',
        r'Nombre[:\s]+([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+?)(?:\n|NIF|DNI)',
        r'ALUMNO.*?([A-Z][A-ZÁÉÍÓÚÑ]+(?:\s+[A-Z][A-ZÁÉÍÓÚÑ]+)+)',
    ]
    
    for patron in patrones:
        match = re.search(patron, texto, re.IGNORECASE | re.DOTALL)
        if match:
            nombre = match.group(1).strip()
            nombre = re.sub(r'\d+', '', nombre)
            nombre = re.sub(r'[^\w\s]', ' ', nombre)
            nombre = ' '.join(nombre.split())
            if len(nombre) > 5:
                return nombre
    
    return None

def extraer_fechas_de_pdf(pdf_path):
    """
    Extrae todas las fechas encontradas en un PDF usando OCR
    Prioriza "Fecha de inicio" y "Fecha de finalización"
    
    Args:
        pdf_path: Ruta al archivo PDF
    
    Returns:
        Lista de objetos datetime
    """
    fechas = []
    
    try:
        textos = extraer_texto_con_ocr(pdf_path, dpi=300)
        
        for texto in textos:
            # PRIORIDAD 1: Buscar "Fecha de inicio" y "Fecha de finalización"
            match_inicio = re.search(
                r'Fecha\s+de\s+inicio:\s+(\d{1,2})/(\d{1,2})/(\d{4})',
                texto,
                re.IGNORECASE
            )
            match_fin = re.search(
                r'Fecha\s+de\s+finalización:\s+(\d{1,2})/(\d{1,2})/(\d{4})',
                texto,
                re.IGNORECASE
            )
            
            if match_inicio:
                dia, mes, año = match_inicio.groups()
                try:
                    fecha = datetime(int(año), int(mes), int(dia))
                    if 2024 <= fecha.year <= 2026 and fecha.month <= 12:
                        fechas.append(fecha)
                except:
                    pass
            
            if match_fin:
                dia, mes, año = match_fin.groups()
                try:
                    fecha = datetime(int(año), int(mes), int(dia))
                    if 2024 <= fecha.year <= 2026 and fecha.month <= 12:
                        fechas.append(fecha)
                except:
                    pass
            
            # PRIORIDAD 2: Buscar patrón SEMANA DEL DD/MM AL DD/MM/YYYY
            patron_semana = r'SEMANA\s+DEL\s+(\d{1,2})/(\d{1,2})\s+AL\s+(\d{1,2})/(\d{1,2})/(\d{4})'
            matches = re.finditer(patron_semana, texto, re.IGNORECASE)
            
            for match in matches:
                dia1, mes1, dia2, mes2, año = match.groups()
                try:
                    fecha_inicio = datetime(int(año), int(mes1), int(dia1))
                    fecha_fin = datetime(int(año), int(mes2), int(dia2))
                    
                    # Validar rango razonable y meses válidos
                    if (fecha_fin - fecha_inicio).days <= 14 and fecha_inicio.month <= 12 and fecha_fin.month <= 12:
                        fechas.append(fecha_inicio)
                        fechas.append(fecha_fin)
                except:
                    pass
    except Exception as e:
        print(f"Error extrayendo fechas: {e}")
    
    return fechas

def contar_dias_con_firmas_por_alumno(pdf_path):
    """
    Cuenta los días con firma por alumno en un PDF
    
    Args:
        pdf_path: Ruta al archivo PDF
    
    Returns:
        Diccionario {nombre_alumno: numero_de_dias}
    """
    dias_por_alumno = {}
    
    try:
        textos = extraer_texto_con_ocr(pdf_path, dpi=300)
        
        for texto in textos:
            nombre = extraer_nombre_alumno_ocr(texto)
            if not nombre:
                continue
            
            patron_semana = r'SEMANA\s+DEL\s+(\d{1,2})/(\d{1,2})\s+AL\s+(\d{1,2})/(\d{1,2})/(\d{4})'
            semanas = re.finditer(patron_semana, texto, re.IGNORECASE)
            
            dias_con_firma = 0
            
            for semana_match in semanas:
                dia1, mes1, dia2, mes2, año = semana_match.groups()
                try:
                    fecha_inicio = datetime(int(año), int(mes1), int(dia1))
                    fecha_fin = datetime(int(año), int(mes2), int(dia2))
                    
                    # Extraer contexto alrededor de la semana
                    inicio_match = semana_match.start()
                    fin_contexto = min(inicio_match + 500, len(texto))
                    contexto = texto[inicio_match:fin_contexto]
                    
                    # Buscar horarios (HH:MM)
                    horarios = re.findall(r'\b(\d{1,2}):(\d{2})\b', contexto)
                    horarios_validos = [
                        (h, m) for h, m in horarios
                        if 0 <= int(h) <= 23 and 0 <= int(m) <= 59
                    ]
                    
                    # Si hay horarios, contar días laborables
                    if len(horarios_validos) >= 2:
                        current = fecha_inicio
                        while current <= fecha_fin:
                            if current.weekday() < 5:  # Lunes a Viernes
                                dias_con_firma += 1
                            current += timedelta(days=1)
                except:
                    pass
            
            if dias_con_firma > 0:
                dias_por_alumno[nombre] = dias_con_firma
    
    except Exception as e:
        print(f"Error procesando firmas: {e}")
    
    return dias_por_alumno

def calcular_dias_lectivos_y_asistencias(firmas_pdfs):
    """
    Calcula días lectivos totales y asistencias por alumno
    
    Lógica:
    1. Recolectar todas las fechas de todos los PDFs
    2. Días lectivos = días laborables desde MIN fecha hasta MAX fecha
    3. Por alumno: contar días con firmas (separando aula vs empresa)
    
    Args:
        firmas_pdfs: Lista de rutas a PDFs de firmas
    
    Returns:
        Tupla (dias_lectivos, asistencias_por_alumno, faltas_por_alumno)
    """
    print("\nCALCULANDO DIAS LECTIVOS Y ASISTENCIAS")
    print("-" * 80)
    
    todas_las_fechas = []
    asistencias_por_alumno = defaultdict(lambda: {'dias_empresa': 0, 'dias_aula': 0})
    
    # Procesar cada PDF
    for idx, pdf_path in enumerate(firmas_pdfs, 1):
        if not os.path.exists(pdf_path):
            print(f"Advertencia: No encontrado {pdf_path}")
            continue
        
        nombre_pdf = os.path.basename(pdf_path)
        print(f"\n[{idx}/{len(firmas_pdfs)}] Procesando: {nombre_pdf}")
        
        # Determinar si es PDF de aula o empresa
        es_aula = 'ParteFirma_30y31' in nombre_pdf or 'aula' in nombre_pdf.lower()
        
        # Extraer fechas
        fechas = extraer_fechas_de_pdf(pdf_path)
        todas_las_fechas.extend(fechas)
        print(f"    {len(fechas)} fechas detectadas")
        
        # Contar días por alumno
        dias_por_alumno = contar_dias_con_firmas_por_alumno(pdf_path)
        print(f"    {len(dias_por_alumno)} alumnos procesados")
        
        for nombre, dias in dias_por_alumno.items():
            if es_aula:
                asistencias_por_alumno[nombre]['dias_aula'] += dias
            else:
                asistencias_por_alumno[nombre]['dias_empresa'] += dias
            print(f"       - {nombre[:35]}: {dias} días ({'aula' if es_aula else 'empresa'})")
    
    # Calcular días lectivos totales
    if todas_las_fechas:
        fecha_min = min(todas_las_fechas)
        fecha_max = max(todas_las_fechas)
        
        print(f"\nPERIODO COMPLETO:")
        print(f"    Desde: {fecha_min.strftime('%d/%m/%Y')}")
        print(f"    Hasta: {fecha_max.strftime('%d/%m/%Y')}")
        
        # Contar días laborables
        dias_lectivos = 0
        current = fecha_min
        while current <= fecha_max:
            if current.weekday() < 5:
                dias_lectivos += 1
            current += timedelta(days=1)
        
        print(f"    Total días lectivos: {dias_lectivos}")
    else:
        dias_lectivos = 0
        print("Advertencia: No se encontraron fechas")
    
    # Calcular faltas
    faltas_por_alumno = {}
    for nombre, datos in asistencias_por_alumno.items():
        total_asistido = datos['dias_empresa'] + datos['dias_aula']
        faltas = max(0, dias_lectivos - total_asistido)
        faltas_por_alumno[nombre] = faltas
    
    print()
    return dias_lectivos, asistencias_por_alumno, faltas_por_alumno