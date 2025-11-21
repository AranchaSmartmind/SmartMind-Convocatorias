"""
GENERADOR DE ACTAS MULTIPÁGINA - CON MÓDULOS EN SEGUNDA PÁGINA
===============================================================
Versión mejorada que incluye código, nombre y horas de módulos en la segunda página

Autor: Sistema de generación de actas
Versión: 2.0 - Con módulos detallados
"""
import io
import re
import zipfile
from typing import Dict, List


class WordGeneratorActaGrupal:
    """Generador base para Acta de Evaluación Final (Grupal) - con módulos"""
    
    def __init__(self, plantilla_bytes: bytes):
        self.plantilla_bytes = plantilla_bytes
        self.plantilla_zip_parts = {}
        
        try:
            with zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r') as zf:
                for item in zf.namelist():
                    self.plantilla_zip_parts[item] = zf.read(item)
        except Exception as e:
            raise Exception(f"Error leyendo plantilla: {e}")
    
    def generar_acta_grupal(self, datos: Dict) -> bytes:
        """Genera acta grupal con todos los alumnos (máximo 15)"""
        
        if 'word/document.xml' not in self.plantilla_zip_parts:
            raise Exception("No se pudo leer word/document.xml")
        
        xml_string = self.plantilla_zip_parts['word/document.xml'].decode('utf-8')
        xml_modificado = self._rellenar_campos_simple(xml_string, datos)
        return self._crear_docx_seguro(xml_modificado)
    
    def _rellenar_campos_simple(self, xml: str, datos: Dict) -> str:
        """Rellena campos manteniendo estructura XML válida"""
        
        alumnos = datos.get('alumnos', [])[:15]  # Máximo 15 alumnos
        
        # Preparar valores (247 campos) - AHORA INCLUYE MÓDULOS EN SEGUNDA PÁGINA
        valores = self._preparar_valores_correctos(datos, alumnos)
        
        print(f"\n=== Procesando Acta Grupal ===")
        print(f"Alumnos a procesar: {len(alumnos)}")
        print(f"Módulos en segunda página: {len(datos.get('modulos_detalle', []))}")
        print(f"Valores preparados: {len(valores)}")
        
        contador = 0
        
        def rellenar_campo(match):
            nonlocal contador
            
            if contador >= len(valores):
                return match.group(0)
            
            valor = str(valores[contador]) if valores[contador] is not None else ''
            campo_actual = contador + 1
            contador += 1
            
            campo_completo = match.group(0)
            
            separate_match = re.search(
                r'<w:fldChar\s+w:fldCharType="separate"[^>]*/>',
                campo_completo
            )
            
            if not separate_match:
                return campo_completo
            
            pos_separate = separate_match.end()
            
            formato_match = re.search(
                r'<w:rPr>(.*?)</w:rPr>',
                campo_completo[:pos_separate],
                re.DOTALL
            )
            
            if formato_match:
                formato_original = formato_match.group(1)
            else:
                formato_original = '<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/>'
            
            # Ajustar tamaño para nombres largos
            campos_nombres = [23, 36, 49, 62, 75, 88, 101, 114, 127, 140, 153, 166, 179, 192, 205]
            
            if campo_actual in campos_nombres and len(valor) > 25:
                size_match = re.search(r'<w:sz w:val="(\d+)"/>', formato_original)
                size_actual = int(size_match.group(1)) if size_match else 20
                
                if len(valor) > 35:
                    nuevo_size = max(14, size_actual - 8)
                elif len(valor) > 30:
                    nuevo_size = max(16, size_actual - 6)
                else:
                    nuevo_size = max(18, size_actual - 4)
                
                formato_ajustado = re.sub(r'<w:sz w:val="\d+"/>', f'<w:sz w:val="{nuevo_size}"/>', formato_original)
                formato_ajustado = re.sub(r'<w:szCs w:val="\d+"/>', f'<w:szCs w:val="{nuevo_size}"/>', formato_ajustado)
                formato = f'<w:rPr>{formato_ajustado}</w:rPr>'
            else:
                formato = f'<w:rPr>{formato_original}</w:rPr>'
            
            contenido_match = re.search(
                r'(<w:fldChar\s+w:fldCharType="separate"[^>]*/>)(.*?)(<w:fldChar\s+w:fldCharType="end")',
                campo_completo,
                re.DOTALL
            )
            
            if not contenido_match:
                return campo_completo
            
            separate_tag = contenido_match.group(1)
            end_tag = contenido_match.group(3)
            
            nuevo_contenido = f'<w:r>{formato}<w:t xml:space="preserve">{valor}</w:t></w:r>'
            
            campo_nuevo = campo_completo.replace(
                separate_tag + contenido_match.group(2) + end_tag,
                separate_tag + nuevo_contenido + end_tag,
                1
            )
            
            return campo_nuevo
        
        patron = (
            r'<w:fldChar\s+w:fldCharType="begin"[^>]*>.*?'
            r'<w:fldChar\s+w:fldCharType="end"[^>]*/?>(?:</w:r>)?'
        )
        
        xml_modificado = re.sub(patron, rellenar_campo, xml, flags=re.DOTALL)
        
        print(f"✓ {contador} campos procesados")
        
        return xml_modificado
    
    def _preparar_valores_correctos(self, datos: Dict, alumnos: List[Dict]) -> List[str]:
        """Prepara lista de 247 valores en el orden correcto - INCLUYE MÓDULOS"""
        
        valores = []
        
        # CAMPOS 1-13: Encabezado del acta
        valores.extend([
            datos.get('fecha_inicio', '20/03/2025'),
            datos.get('fecha_fin', '27/06/2025'),
            '',  # Campo 3: Número de expediente (vacío)
            datos.get('curso_nombre', ''),
            datos.get('codigo_certificado', ''),
            datos.get('nivel', ''),
            'INTERPROS NEXT GENERATION SLU',
            '26615',
            'C/ DR. SEVERO OCHOA, 21, BJ - AVILÉS',
            '33401',
            'ASTURIAS',
            '985 525 111',
            'asturias@smartmind.net',
        ])
        
        # CAMPOS 14-21: Nombres de módulos (8 módulos máximo) - PRIMERA PÁGINA
        modulos_nombres = [mod['nombre'] for mod in datos.get('modulos_info', [])]
        for i in range(8):
            valores.append(modulos_nombres[i] if i < len(modulos_nombres) else '')
        
        # CAMPOS 22-216: Datos de 15 alumnos (13 campos por alumno)
        for i in range(15):
            if i < len(alumnos):
                alumno = alumnos[i]
                
                # Obtener calificaciones de los primeros 3 módulos
                mod1 = self._obtener_calificacion_modulo(alumno, 0) or 'NS-0'
                mod2 = self._obtener_calificacion_modulo(alumno, 1) or 'NS-0'
                mod3 = self._obtener_calificacion_modulo(alumno, 2) or 'NS-0'
                
                valores.extend([
                    alumno.get('dni', ''),
                    alumno.get('nombre', ''),
                    mod1,
                    mod2,
                    mod3,
                    '',
                    '',
                    '',
                    '',
                    '',
                    self._calc_calificacion_numerica(alumno) or 'NS-0.00',
                    self._calc_certificacion(alumno),
                    '',
                ])
            else:
                # Alumno vacío
                valores.extend([''] * 13)
        
        # CAMPO 217: Vacío
        valores.append('')
        
        # CAMPO 218: TOTAL de alumnos
        total_alumnos = datos.get('total_alumnos', len(datos.get('alumnos', [])))
        valores.append(str(total_alumnos))
        
        # ============================================================
        # CAMPOS 219-247: MÓDULOS DETALLADOS EN SEGUNDA PÁGINA (29 campos)
        # ============================================================
        modulos_detalle = datos.get('modulos_detalle', [])
        
        # Campos 219-227: Información de los 3 módulos (3 campos por módulo)
        for i in range(3):
            if i < len(modulos_detalle):
                modulo = modulos_detalle[i]
                valores.extend([
                    modulo.get('codigo', ''),      # Código: MF0969_1
                    modulo.get('nombre', ''),      # Nombre: Técnicas administrativas...
                    str(modulo.get('horas', '')),  # Horas: 150
                ])
            else:
                valores.extend(['', '', ''])  # Módulo vacío
        
        # Campos 228-247: Campos adicionales (20 campos) - Por ahora vacíos
        valores.extend([''] * 20)
        
        print(f"  → Módulos detalle agregados: {len(modulos_detalle)}")
        
        return valores
    
    def _obtener_calificacion_modulo(self, alumno: Dict, idx: int) -> str:
        """Obtiene calificación de un módulo específico"""
        modulos = alumno.get('modulos', [])
        if idx < len(modulos):
            return modulos[idx].get('calificacion', '')
        return ''
    
    def _calc_calificacion_numerica(self, alumno: Dict) -> str:
        """Calcula calificación numérica promedio"""
        modulos = alumno.get('modulos', [])
        if not modulos:
            return ''
        
        notas = [m.get('nota') for m in modulos if m.get('nota') is not None]
        if not notas:
            return ''
        
        promedio = sum(notas) / len(notas)
        return f"S-{promedio:.2f}" if promedio >= 5 else f"NS-{promedio:.2f}"
    
    def _calc_certificacion(self, alumno: Dict) -> str:
        """Determina si el alumno obtiene certificación"""
        modulos = alumno.get('modulos', [])
        if not modulos:
            return 'NO'
        
        notas = [m.get('nota') for m in modulos if m.get('nota') is not None]
        if not notas:
            return 'NO'
        
        promedio = sum(notas) / len(notas)
        return 'SÍ' if promedio >= 5.0 else 'NO'
    
    def _crear_docx_seguro(self, xml_modificado: str) -> bytes:
        """Crea documento DOCX con el XML modificado"""
        output = io.BytesIO()
        
        try:
            with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as docx:
                for nombre, contenido in self.plantilla_zip_parts.items():
                    if nombre == 'word/document.xml':
                        docx.writestr(nombre, xml_modificado.encode('utf-8'), compress_type=zipfile.ZIP_DEFLATED)
                    else:
                        docx.writestr(nombre, contenido, compress_type=zipfile.ZIP_DEFLATED)
            
            output.seek(0)
            return output.getvalue()
            
        except Exception as e:
            raise Exception(f"Error creando DOCX: {e}")


class WordGeneratorMultipaginaDuplicaTodo:
    """Generador que crea múltiples páginas completas (con módulos y leyenda)"""
    
    def __init__(self, plantilla_bytes: bytes):
        self.plantilla_bytes = plantilla_bytes
    
    def generar_acta_grupal(self, datos: Dict) -> bytes:
        """Genera actas múltiples cuando hay más de 15 alumnos"""
        
        alumnos = datos.get('alumnos', [])
        total_alumnos = len(alumnos)
        
        print(f"\n=== Generando Acta Multipágina (Duplica Todo) ===")
        print(f"Total de alumnos: {total_alumnos}")
        
        if total_alumnos <= 15:
            print("✓ Una sola acta suficiente")
            generador = WordGeneratorActaGrupal(self.plantilla_bytes)
            return generador.generar_acta_grupal(datos)
        
        # Generar actas separadas COMPLETAS
        print(f"✓ Generando {(total_alumnos + 14) // 15} actas completas...")
        
        actas_bytes = []
        alumno_idx = 0
        pagina = 1
        
        while alumno_idx < total_alumnos:
            fin_idx = min(alumno_idx + 15, total_alumnos)
            print(f"  Acta {pagina}: Alumnos {alumno_idx + 1}-{fin_idx} (COMPLETA)")
            
            datos_acta = datos.copy()
            datos_acta['alumnos'] = alumnos[alumno_idx:fin_idx]
            datos_acta['total_alumnos'] = total_alumnos
            
            generador = WordGeneratorActaGrupal(self.plantilla_bytes)
            acta_bytes = generador.generar_acta_grupal(datos_acta)
            actas_bytes.append(acta_bytes)
            
            alumno_idx = fin_idx
            pagina += 1
        
        print(f"\n✓ Combinando {len(actas_bytes)} actas...")
        acta_combinada = self._combinar_actas_completas(actas_bytes)
        
        print(f"✓ Acta combinada generada")
        return acta_combinada
    
    def _combinar_actas_completas(self, actas_bytes: List[bytes]) -> bytes:
        """Combina múltiples actas completas en un solo documento"""
        
        if len(actas_bytes) == 1:
            return actas_bytes[0]
        
        # Cargar primera acta
        primera_acta = {}
        with zipfile.ZipFile(io.BytesIO(actas_bytes[0]), 'r') as z:
            for item in z.namelist():
                primera_acta[item] = z.read(item)
        
        xml_base = primera_acta['word/document.xml'].decode('utf-8')
        
        # Extraer body
        match_body = re.search(r'<w:body>(.*?)</w:body>', xml_base, re.DOTALL)
        if not match_body:
            return actas_bytes[0]
        
        contenido_body = match_body.group(1)
        
        # Agregar contenido de otras actas
        for acta_bytes in actas_bytes[1:]:
            with zipfile.ZipFile(io.BytesIO(acta_bytes), 'r') as z:
                xml_acta = z.read('word/document.xml').decode('utf-8')
            
            match = re.search(r'<w:body>(.*?)</w:body>', xml_acta, re.DOTALL)
            if match:
                contenido_acta = match.group(1)
                salto = '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'
                contenido_body += salto + contenido_acta
        
        # Reconstruir XML
        xml_combinado = xml_base.replace(match_body.group(1), contenido_body, 1)
        
        # Crear DOCX
        output = io.BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:
            for nombre, contenido in primera_acta.items():
                if nombre == 'word/document.xml':
                    docx.writestr(nombre, xml_combinado.encode('utf-8'))
                else:
                    docx.writestr(nombre, contenido)
        
        output.seek(0)
        return output.getvalue()


# ============================================================
# FUNCIÓN AUXILIAR PARA EXTRAER MÓDULOS DEL CRONOGRAMA
# ============================================================

def extraer_modulos_de_cronograma(archivo) -> List[Dict]:
    """
    Extrae información de módulos del archivo Cronograma.
    Acepta rutas, bytes o archivos subidos vía Streamlit.
    """
    import pandas as pd
    import io

    try:
        # Si es un UploadedFile de Streamlit → usamos archivo.read()
        if hasattr(archivo, "read"):
            df = pd.read_excel(io.BytesIO(archivo.read()), sheet_name='Calculos_UF', header=None)
        else:
            # Ruta en disco tradicional
            df = pd.read_excel(archivo, sheet_name='Calculos_UF', header=None)

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
                                horas = float(siguiente[2])
                                break

                            if pd.notna(siguiente[0]) and str(siguiente[0]).startswith('MF'):
                                break

                    if horas:
                        modulos.append({
                            'codigo': codigo,
                            'nombre': nombre,
                            'horas': int(horas)
                        })
        
        return modulos

    except Exception as e:
        print(f"Error extrayendo módulos: {e}")
        return []
