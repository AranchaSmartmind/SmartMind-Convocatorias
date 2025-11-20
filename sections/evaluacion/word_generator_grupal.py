"""
Generador de Acta Grupal - VERSIÓN DEFINITIVA CORREGIDA
Mapeo exacto verificado: 247 campos
"""
import io
import re
import zipfile
from typing import Dict, List


class WordGeneratorActaGrupal:
    """Generador para Acta de Evaluación Final (Grupal)"""
    
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
        """Genera acta grupal con todos los alumnos"""
        
        if 'word/document.xml' not in self.plantilla_zip_parts:
            raise Exception("No se pudo leer word/document.xml")
        
        xml_string = self.plantilla_zip_parts['word/document.xml'].decode('utf-8')
        xml_modificado = self._rellenar_campos_simple(xml_string, datos)
        return self._crear_docx_seguro(xml_modificado)
    
    def _rellenar_campos_simple(self, xml: str, datos: Dict) -> str:
        """Rellena campos de forma simple y segura"""
        
        alumnos = datos.get('alumnos', [])[:15]  # Máximo 15 alumnos
        
        # Preparar valores (247 campos)
        valores = self._preparar_valores_correctos(datos, alumnos)
        
        print(f"\n=== Procesando Acta Grupal ===")
        print(f"Alumnos a procesar: {len(alumnos)}")
        print(f"Valores preparados: {len(valores)}")
        
        contador = 0
        
        def rellenar_campo(match):
            nonlocal contador
            
            if contador >= len(valores):
                return match.group(0)
            
            valor = str(valores[contador]) if valores[contador] is not None else ''
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
                formato = f'<w:rPr>{formato_match.group(1)}</w:rPr>'
            else:
                formato = '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr>'
            
            contenido_match = re.search(
                r'(<w:fldChar\s+w:fldCharType="separate"[^>]*/>)(.*?)(<w:fldChar\s+w:fldCharType="end")',
                campo_completo,
                re.DOTALL
            )
            
            if not contenido_match:
                return campo_completo
            
            separate_tag = contenido_match.group(1)
            contenido_viejo = contenido_match.group(2)
            end_tag = contenido_match.group(3)
            
            nuevo_contenido = f'<w:r>{formato}<w:t xml:space="preserve">{valor}</w:t></w:r>'
            
            campo_nuevo = campo_completo.replace(
                separate_tag + contenido_viejo + end_tag,
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
        """
        Prepara valores en el orden EXACTO verificado:
        
        Campos 1-3: Fila 3 (Acción Formativa, Fecha inicio, Fecha fin)
        Campos 4-6: Fila 5 (Certificado prof., Código cert., Nivel)
        Campos 7-9: Fila 6 (Centro formativo, Código centro, Localidad)
        Campos 10-13: Fila 7 (CP, Provincia, Teléfono, Email)
        Campos 14-21: Fila 10 (Encabezados de módulos - 8 campos vacíos)
        Campos 22-34: Fila 11 - Alumno 1 (13 campos)
        Campos 35-47: Fila 12 - Alumno 2 (13 campos)
        ...
        Campos 216: Fila 25 - Alumno 15 (último alumno)
        Campos 217-247: Resto (31 campos)
        """
        
        valores = []
        
        # PARTE 1: Fila 3 - Datos del curso
        # Según análisis: 3 campos pero visualmente aparecen 4 valores
        # Posiblemente: Acción, Fecha inicio, Fecha fin + otro campo extra
        valores.extend([
            datos.get('curso_nombre', ''),                        # 1. Acción Formativa
            datos.get('fecha_inicio', '20/03/2025'),             # 2. Fecha de inicio  
            datos.get('fecha_fin', '27/06/2025'),                # 3. Fecha de finalización
        ])
        
        # PARTE 2: Fila 5 - Certificado
        valores.extend([
            datos.get('curso_nombre', ''),                        # 4. Certificado profesional
            datos.get('curso_codigo', '').split('/')[-1] if datos.get('curso_codigo') else '1339',  # 5. Código
            'C',                                                  # 6. Nivel
        ])
        
        # PARTE 3: Fila 6 - Centro
        valores.extend([
            'INTERPROS NEXT GENERATION SLU',                     # 7. Centro formativo
            '26615',                                              # 8. Código centro
            'C/ DR. SEVERO OCHOA, 21, BJ - AVILÉS',             # 9. Localidad
        ])
        
        # PARTE 4: Fila 7 - Contacto  
        valores.extend([
            '33401',                                              # 10. CP
            'ASTURIAS',                                           # 11. Provincia
            '985 525 111',                                        # 12. Teléfono
            'asturias@smartmind.net',                            # 13. Email
        ])
        
        # PARTE 5: Fila 10 - Encabezados de módulos (8 campos VACÍOS)
        valores.extend([''] * 8)                                  # 14-21. Vacíos
        
        # PARTE 6: Filas 11-25 - Alumnos (15 alumnos × 13 campos = 195 campos)
        for i in range(15):
            if i < len(alumnos):
                alumno = alumnos[i]
                
                # 13 campos por alumno:
                valores.extend([
                    alumno.get('dni', ''),                       # 1. DNI
                    alumno.get('nombre', ''),                    # 2. Nombre
                    '',                                           # 3. Vacío
                    self._obtener_calificacion_modulo(alumno, 0), # 4. Módulo 1 (S/NS/CO-5)
                    self._obtener_calificacion_modulo(alumno, 1), # 5. Módulo 2 (S/NS/CO-5)
                    self._obtener_calificacion_modulo(alumno, 2), # 6. Módulo 3 (S/NS/CO-5)
                    '',                                           # 7. Vacío
                    '',                                           # 8. Vacío
                    '',                                           # 9. Vacío
                    '',                                           # 10. Vacío
                    self._calc_calificacion_numerica(alumno),    # 11. Calificación² (con decimales)
                    self._calc_certificacion(alumno),            # 12. Certificación (SÍ/NO según nota>=5)
                    '',                                           # 13. PFE (vacío)
                ])
            else:
                # Fila vacía (13 campos vacíos)
                valores.extend([''] * 13)
        
        # PARTE 7: Resto (31 campos finales hasta 247)
        campos_restantes = 247 - len(valores)
        valores.extend([''] * campos_restantes)
        
        # Debug
        print(f"\n  Verificación de valores:")
        print(f"    Total valores: {len(valores)}")
        print(f"    Campo 1 (Curso): {valores[0][:40]}...")
        print(f"    Campo 22 (DNI Al.1): {valores[21]}")
        print(f"    Campo 23 (Nombre Al.1): {valores[22][:30]}...")
        print(f"    Campo 25 (Mod1 Al.1): {valores[24]}")
        
        return valores
    
    def _obtener_calificacion_modulo(self, alumno: Dict, idx: int) -> str:
        """Obtiene la calificación formateada de un módulo (S/NS/CO-5)"""
        modulos = alumno.get('modulos', [])
        
        if idx < len(modulos):
            modulo = modulos[idx]
            return modulo.get('calificacion', '')
        
        return ''
    
    def _calc_calificacion_numerica(self, alumno: Dict) -> str:
        """
        Calcula calificación numérica (nota final) con 2 decimales
        Esta va en la columna "Calificación²"
        Formato: S-8.50, S-9.33, NS-0.00
        
        IMPORTANTE: Si no hay notas, retorna vacío (no NS-0.00)
        """
        modulos = alumno.get('modulos', [])
        
        if not modulos:
            return ''
        
        notas = []
        for modulo in modulos:
            nota = modulo.get('nota')
            if nota is not None:
                notas.append(nota)
        
        # Si no hay ninguna nota, retornar vacío
        if not notas:
            return ''
        
        promedio = sum(notas) / len(notas)
        
        # Formato: S-8.50 o NS-0.00
        if promedio >= 5:
            return f"S-{promedio:.2f}"
        else:
            return f"NS-{promedio:.2f}"
    
    def _calc_nota_final(self, alumno: Dict) -> str:
        """
        OBSOLETA - Usar _calc_calificacion_numerica
        Se mantiene por compatibilidad
        """
        return self._calc_calificacion_numerica(alumno)
    
    def _calc_certificacion(self, alumno: Dict) -> str:
        """
        Determina "Propuesta de certificación"
        SÍ si nota final >= 5.0
        NO si nota final < 5.0
        """
        modulos = alumno.get('modulos', [])
        
        if not modulos:
            return 'NO'
        
        # Calcular nota final
        notas = []
        for modulo in modulos:
            nota = modulo.get('nota')
            if nota is not None:
                notas.append(nota)
        
        if not notas:
            return 'NO'
        
        promedio = sum(notas) / len(notas)
        
        # SÍ si >= 5.0, NO si < 5.0
        return 'SÍ' if promedio >= 5.0 else 'NO'
    
    def _crear_docx_seguro(self, xml_modificado: str) -> bytes:
        """Crea DOCX copiando TODO del ZIP original"""
        
        if not self.plantilla_zip_parts:
            raise Exception("No hay plantilla cargada")
        
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