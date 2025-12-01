"""
Módulo de procesamiento de datos - VERSIÓN FINAL CORREGIDA
Extrae datos de becas, justificantes y hojas de firmas
"""

import re
from datetime import datetime, timedelta
from collections import defaultdict
from pdf2image import convert_from_path
import pytesseract
import os

# Configuración
os.environ['TESSDATA_PREFIX'] = r'C:\Program Files\Tesseract-OCR\tessdata'
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
POPPLER_PATH = r'C:\Users\Arancha\Downloads\poppler-24.08.0\Library\bin'


def extraer_texto_con_ocr(pdf_path, dpi=300):
    """Convierte PDF a imágenes y extrae texto con OCR"""
    try:
        images = convert_from_path(pdf_path, dpi=dpi, poppler_path=POPPLER_PATH)
        textos_por_pagina = []
        
        for image in images:
            texto = pytesseract.image_to_string(image, lang='spa')
            textos_por_pagina.append(texto)
        
        return textos_por_pagina
    except Exception as e:
        print(f"Error en OCR: {e}")
        return []


def extraer_becas_ayudas_tabla(pdf_path):
    """
    Extrae ayudas del PDF de Otorgamiento (tabla estructurada)
    
    Columnas: Ayuda transporte | Manutención | Alojamiento | Beca | Ayuda conciliación | Otras
    Ignora columna "Beca"
    Retorna: {nombre_alumno: ['Transporte', 'Conciliación', ...]}
    """
    ayudas_por_alumno = defaultdict(list)
    
    try:
        print("\n📋 EXTRAYENDO BECAS Y AYUDAS")
        print("-" * 80)
        
        textos = extraer_texto_con_ocr(pdf_path, dpi=300)
        
        for texto in textos:
            lineas = texto.split('\n')
            
            for i, linea in enumerate(lineas):
                # Buscar líneas con NIE/DNI (formato: NOMBRE + DNI + datos)
                match_alumno = re.search(r'([A-ZÁÉÍÓÚÑ\s,]+)\s+([Z\d]\d{7}[A-Z])', linea)
                
                if match_alumno:
                    nombre_completo = match_alumno.group(1).strip()
                    
                    # Limpiar nombre
                    nombre_completo = re.sub(r'\s+', ' ', nombre_completo)
                    
                    ayudas = []
                    
                    # Transporte (buscar X después del importe)
                    if re.search(r'[0-9,]+€.*?X', linea):
                        ayudas.append('Transporte')
                    
                    # Conciliación
                    if 'conciliaci' in linea.lower():
                        # Buscar X cerca de conciliación
                        resto = linea.lower().split('conciliaci')[1] if 'conciliaci' in linea.lower() else ''
                        if 'x' in resto[:30]:
                            ayudas.append('Conciliación')
                    
                    # Discapacidad
                    if 'DISCAPACIDAD' in linea.upper():
                        ayudas.append('Discapacidad')
                    
                    if ayudas:
                        ayudas_por_alumno[nombre_completo] = ayudas
                        print(f"  ✅ {nombre_completo}: {', '.join(ayudas)}")
        
        print(f"\n✅ {len(ayudas_por_alumno)} alumnos con ayudas")
        
    except Exception as e:
        print(f"❌ Error extrayendo becas: {e}")
    
    return dict(ayudas_por_alumno)


def extraer_justificantes_mejorado(pdf_path):
    """
    Extrae justificantes del PDF
    
    Busca nombres de alumnos y cuenta cuántos justificantes tienen
    Retorna: {nombre_alumno: numero_justificantes}
    """
    justificantes_por_alumno = defaultdict(int)
    
    try:
        print("\n🏥 EXTRAYENDO JUSTIFICANTES")
        print("-" * 80)
        
        textos = extraer_texto_con_ocr(pdf_path, dpi=300)
        
        for texto in textos:
            # Buscar nombres en formato "APELLIDO APELLIDO, NOMBRE"
            # Patrón 1: Doña/Don NOMBRE APELLIDO
            matches = re.finditer(r'[Dd]o[ñn]a?\s+([A-ZÁÉÍÓÚÑ][a-záéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][a-záéíóúñ]+){1,4})', texto)
            for match in matches:
                nombre = match.group(1).strip()
                if len(nombre) > 8:  # Filtrar nombres muy cortos
                    justificantes_por_alumno[nombre] += 1
                    print(f"  📄 {nombre}: +1 justificante")
            
            # Patrón 2: Buscar nombres en mayúsculas después de JUSTIFICANTE
            if 'JUSTIFICANTE' in texto.upper():
                match = re.search(r'([A-ZÁÉÍÓÚÑ]{3,}(?:\s+[A-ZÁÉÍÓÚÑ]{3,}){2,4})', texto[texto.upper().find('JUSTIFICANTE'):])
                if match:
                    nombre = match.group(1).strip()
                    if len(nombre) > 10:
                        justificantes_por_alumno[nombre] += 1
                        print(f"  📄 {nombre}: +1 justificante")
        
        print(f"\n✅ {len(justificantes_por_alumno)} alumnos con justificantes")
        
    except Exception as e:
        print(f"❌ Error extrayendo justificantes: {e}")
    
    return dict(justificantes_por_alumno)


def calcular_dias_lectivos_y_faltas_corregido(firmas_pdfs, dias_lectivos_total):
    """
    MÉTODO MEJORADO PARA CONTAR DÍAS CON FIRMA
    
    Estrategias múltiples para detectar días firmados:
    1. Contar días de la semana individuales con horarios
    2. Contar semanas completas (SEMANA DEL)
    3. Contar filas de tabla con días
    
    Args:
        firmas_pdfs: Lista de rutas a PDFs de firmas
        dias_lectivos_total: Total de días lectivos del mes (ingresado por usuario)
    
    Returns:
        (dias_lectivos_total, dias_ausentes_por_alumno, dias_con_firma_por_alumno)
    """
    print("\n" + "="*80)
    print("CONTANDO DÍAS CON FIRMA - MÉTODO MEJORADO")
    print("="*80)
    print(f"📅 Días lectivos del mes: {dias_lectivos_total}")
    
    # Almacenar días con firma por alumno (acumulado de todos los PDFs)
    dias_firmados_por_alumno = defaultdict(int)
    
    for idx, pdf_path in enumerate(firmas_pdfs, 1):
        if not os.path.exists(pdf_path):
            print(f"⚠️  No encontrado: {pdf_path}")
            continue
        
        nombre_pdf = os.path.basename(pdf_path)
        print(f"\n[{idx}/{len(firmas_pdfs)}] {nombre_pdf}")
        
        try:
            textos = extraer_texto_con_ocr(pdf_path, dpi=300)
            
            for pagina_idx, texto in enumerate(textos, 1):
                print(f"\n  📄 Página {pagina_idx}")
                
                # Extraer nombre del alumno
                match_nombre = re.search(r'Nombre:\s*([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑa-záéíóúñ\s]+?)(?:\n|NIF|DNI)', texto, re.IGNORECASE)
                
                if not match_nombre:
                    print(f"    ⚠️  No se encontró nombre de alumno")
                    continue
                
                nombre = match_nombre.group(1).strip()
                nombre = ' '.join(nombre.split())
                
                print(f"    👤 Procesando: {nombre}")
                
                # ESTRATEGIA 1: Contar semanas completas
                semanas = re.findall(r'SEMANA\s+DEL', texto, re.IGNORECASE)
                dias_por_semanas = len(semanas) * 5
                
                if dias_por_semanas > 0:
                    print(f"    📅 Método SEMANAS: {len(semanas)} semanas = {dias_por_semanas} días")
                
                # ESTRATEGIA 2: Contar días individuales con horarios
                dias_semana_patterns = [
                    r'LUNES[^\n]*\d{1,2}:\d{2}',
                    r'MARTES[^\n]*\d{1,2}:\d{2}',
                    r'MI[ÉE]RCOLES[^\n]*\d{1,2}:\d{2}',
                    r'JUEVES[^\n]*\d{1,2}:\d{2}',
                    r'VIERNES[^\n]*\d{1,2}:\d{2}'
                ]
                
                dias_individuales = 0
                for pattern in dias_semana_patterns:
                    matches = re.findall(pattern, texto, re.IGNORECASE)
                    dias_individuales += len(matches)
                
                if dias_individuales > 0:
                    print(f"    📋 Método DÍAS INDIVIDUALES: {dias_individuales} días")
                
                # ESTRATEGIA 3: Contar bloques de tabla (cada bloque = 1 semana)
                # Buscar patrones de tabla con estructura repetitiva
                bloques_tabla = len(re.findall(r'LUNES.*?MARTES.*?MI[ÉE]RCOLES.*?JUEVES.*?VIERNES', texto, re.IGNORECASE | re.DOTALL))
                dias_por_bloques = bloques_tabla * 5
                
                if dias_por_bloques > 0:
                    print(f"    📊 Método BLOQUES TABLA: {bloques_tabla} bloques = {dias_por_bloques} días")
                
                # ESTRATEGIA 4: Buscar patrones específicos como "Del DD/MM al DD/MM"
                # y calcular días laborables
                rangos_fecha = re.findall(
                    r'del?\s+(\d{1,2})[/-](\d{1,2})[^\n]*al?\s+(\d{1,2})[/-](\d{1,2})',
                    texto,
                    re.IGNORECASE
                )
                
                dias_por_rangos = 0
                if rangos_fecha:
                    # Asumir que cada rango es una semana = 5 días
                    dias_por_rangos = len(rangos_fecha) * 5
                    print(f"    📆 Método RANGOS: {len(rangos_fecha)} rangos = {dias_por_rangos} días")
                
                # ELEGIR EL MEJOR MÉTODO (el que dé más días)
                dias_detectados = max(dias_por_semanas, dias_individuales, dias_por_bloques, dias_por_rangos)
                
                # Si no detectamos nada, intentar método de emergencia
                if dias_detectados == 0:
                    # Contar cuántas veces aparece cada día de la semana
                    cuenta_dias = 0
                    for dia in ['LUNES', 'MARTES', 'MIÉRCOLES', 'MIERCOLES', 'JUEVES', 'VIERNES']:
                        cuenta_dias += len(re.findall(dia, texto, re.IGNORECASE))
                    
                    # Si aparecen muchos días (probablemente tabla completa)
                    if cuenta_dias >= 15:  # Al menos 3 semanas
                        dias_detectados = (cuenta_dias // 5) * 5  # Redondear a semanas completas
                        print(f"    🔧 Método EMERGENCIA: {cuenta_dias} menciones → {dias_detectados} días")
                
                if dias_detectados > 0:
                    # Asegurarse de no exceder el total
                    dias_detectados = min(dias_detectados, dias_lectivos_total)
                    dias_firmados_por_alumno[nombre] += dias_detectados
                    print(f"    ✅ {nombre}: +{dias_detectados} días (total acumulado: {dias_firmados_por_alumno[nombre]})")
                else:
                    print(f"    ⚠️  {nombre}: No se detectaron días con firma")
        
        except Exception as e:
            print(f"  ❌ Error: {e}")
            import traceback
            traceback.print_exc()
    
    # CALCULAR AUSENCIAS
    dias_ausentes_por_alumno = {}
    
    print(f"\n📊 CÁLCULO FINAL DE AUSENCIAS:")
    print("-" * 80)
    
    for nombre, dias_firmados in dias_firmados_por_alumno.items():
        # Limitar días firmados al máximo de días lectivos
        dias_firmados = min(dias_firmados, dias_lectivos_total)
        
        # Ausencias = total - firmados
        ausencias = max(0, dias_lectivos_total - dias_firmados)
        dias_ausentes_por_alumno[nombre] = ausencias
        
        print(f"  👤 {nombre}")
        print(f"     - Días lectivos totales: {dias_lectivos_total}")
        print(f"     - Días con firma detectados: {dias_firmados}")
        print(f"     - Días AUSENTE: {ausencias}")
    
    print(f"\n✅ {len(dias_firmados_por_alumno)} alumnos procesados")
    
    return dias_lectivos_total, dias_ausentes_por_alumno, dict(dias_firmados_por_alumno)


def construir_observaciones_completas(nombre_alumno, ayudas_dict, dias_asistidos_dict, justificantes_dict):
    """
    Construye observaciones en formato de 2 líneas:
    Línea 1: [Ayudas] : [Días asistidos]
    Línea 2: [X] faltas justificadas
    """
    lineas = []
    
    # Línea 1: Ayudas y días
    ayudas = ayudas_dict.get(nombre_alumno, [])
    dias = dias_asistidos_dict.get(nombre_alumno, 0)
    
    ayudas_texto = ' + '.join(ayudas) if ayudas else ''
    linea1 = f"{ayudas_texto} : {dias}".strip()
    if linea1.startswith(':'):
        linea1 = linea1[1:].strip()
    
    if linea1:
        lineas.append(linea1)
    
    # Línea 2: Justificantes
    justif = justificantes_dict.get(nombre_alumno, 0)
    if justif > 0:
        lineas.append(f"{justif} falta{'s' if justif > 1 else ''} justificada{'s' if justif > 1 else ''}")
    
    return '\n'.join(lineas)


def obtener_mes_anterior():
    """Retorna el nombre del mes anterior en mayúsculas"""
    from datetime import datetime
    meses = ['ENERO', 'FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO',
             'JULIO', 'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE', 'DICIEMBRE']
    mes_actual = datetime.now().month
    mes_anterior = (mes_actual - 2) % 12
    return meses[mes_anterior]


def extraer_datos_curso_pdf(pdf_path):
    """
    Extrae número de curso y especialidad del PDF de becas
    Returns: (numero_curso, especialidad)
    """
    try:
        textos = extraer_texto_con_ocr(pdf_path, dpi=300)
        
        for texto in textos:
            # Buscar número de curso
            match_curso = re.search(r'(?:Nº de Curso|N° de Curso)[:\s]+(\d{4}/\d+)', texto, re.IGNORECASE)
            numero_curso = match_curso.group(1) if match_curso else ''
            
            # Buscar especialidad
            match_esp = re.search(r'OPERACIONES AUXILIARES[^\n]+', texto, re.IGNORECASE)
            especialidad = match_esp.group(0) if match_esp else ''
            
            if numero_curso or especialidad:
                return numero_curso, especialidad
        
        return '', ''
    except Exception as e:
        print(f"Error extrayendo datos curso: {e}")
        return '', ''


def extraer_alumnos_excel(excel_path):
    """
    Extrae lista de alumnos del Excel
    Returns: [{'nombre_completo': str, 'dni': str}, ...]
    """
    import openpyxl
    
    alumnos = []
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        # Buscar encabezados
        nombre_col = None
        dni_col = None
        
        for row in ws.iter_rows(min_row=1, max_row=10):
            for cell in row:
                if cell.value and 'nombre' in str(cell.value).lower():
                    nombre_col = cell.column
                if cell.value and ('dni' in str(cell.value).lower() or 'nif' in str(cell.value).lower()):
                    dni_col = cell.column
        
        if not nombre_col or not dni_col:
            print("⚠️ No se encontraron columnas de nombre/DNI en Excel")
            return []
        
        # Extraer datos
        for row in ws.iter_rows(min_row=2):
            nombre = row[nombre_col - 1].value
            dni = row[dni_col - 1].value
            
            if nombre and dni:
                nombre = str(nombre).strip()
                dni = str(dni).strip()
                
                alumnos.append({
                    'nombre_completo': nombre,
                    'dni': dni
                })
        
        print(f"✅ {len(alumnos)} alumnos extraídos del Excel")
        
    except Exception as e:
        print(f"❌ Error leyendo Excel: {e}")
    
    return alumnos


