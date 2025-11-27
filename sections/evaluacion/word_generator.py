"""
Generador Word - BASADO EN CÓDIGO GRUPAL QUE FUNCIONA
Genera ZIP con múltiples documentos cuando hay más de 6 módulos
"""
import io
import re
import zipfile
from typing import Dict, List


class WordGeneratorSEPE:
    """
    Generador de informes individuales
    Basado en la estructura del WordGeneratorActaGrupal que SÍ funciona
    """
    
    def __init__(self, plantilla_bytes: bytes, es_xml: bool = False):
        self.plantilla_bytes = plantilla_bytes
        self.es_xml = es_xml
        self.plantilla_zip_parts = {}
        
        if plantilla_bytes and not es_xml:
            try:
                with zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r') as zf:
                    for item in zf.namelist():
                        self.plantilla_zip_parts[item] = zf.read(item)
            except Exception as e:
                print(f"Error leyendo DOCX: {e}")
    
    def generar_informe_individual(self, datos: Dict) -> bytes:
        """
        Genera informe individual
        Retorna .docx si ≤6 módulos, .zip si >6 módulos
        """
        
        alumno = datos.get('alumno', {})
        modulos = alumno.get('modulos', [])
        total_modulos = len(modulos)
        
        print(f"\n=== Generando Informe Individual ===")
        print(f"Alumno: {alumno.get('nombre', 'N/A')}")
        print(f"Total módulos: {total_modulos}")
        
        # Si hay ≤6 módulos, documento único
        if total_modulos <= 6:
            print("✓ Documento único")
            return self._generar_documento_unico(datos)
        
        # Si hay >6 módulos, ZIP con múltiples documentos
        print(f"✓ Generando {(total_modulos + 5) // 6} documentos en ZIP...")
        return self._generar_zip_multiples(datos)
    
    def es_zip(self, datos_bytes: bytes) -> bool:
        """Detecta si los bytes son un ZIP o un DOCX"""
        return datos_bytes[:2] == b'PK' and datos_bytes[2:4] != b'\x03\x04'
    
    def extraer_archivos_de_zip(self, zip_bytes: bytes) -> List[tuple]:
        """
        Extrae archivos de un ZIP
        Retorna: Lista de tuplas (nombre_archivo, contenido_bytes)
        """
        archivos = []
        try:
            with zipfile.ZipFile(io.BytesIO(zip_bytes), 'r') as zf:
                for nombre in zf.namelist():
                    contenido = zf.read(nombre)
                    archivos.append((nombre, contenido))
        except Exception as e:
            print(f"Error extrayendo ZIP: {e}")
        return archivos
    
    def _generar_documento_unico(self, datos: Dict) -> bytes:
        """Genera un solo documento .docx"""
        
        if 'word/document.xml' not in self.plantilla_zip_parts:
            raise Exception("No se pudo leer word/document.xml")
        
        xml_string = self.plantilla_zip_parts['word/document.xml'].decode('utf-8')
        xml_modificado = self._rellenar_campos(xml_string, datos)
        xml_modificado = self._rellenar_tabla_modulos(xml_modificado, datos)
        
        return self._crear_docx(xml_modificado)
    
    def _generar_zip_multiples(self, datos: Dict) -> bytes:
        """Genera ZIP con múltiples documentos (estilo grupal)"""
        
        alumno = datos.get('alumno', {})
        modulos = alumno.get('modulos', [])
        nombre_alumno = alumno.get('nombre', 'Alumno').replace(' ', '_').replace(',', '')[:50]
        
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            modulo_idx = 0
            pagina = 1
            
            while modulo_idx < len(modulos):
                fin_idx = min(modulo_idx + 6, len(modulos))
                print(f"  Documento {pagina}: Módulos {modulo_idx + 1}-{fin_idx}")
                
                # Crear datos para este documento
                datos_doc = {
                    'alumno': {
                        'nombre': alumno.get('nombre', ''),
                        'dni': alumno.get('dni', ''),
                        'modulos': modulos[modulo_idx:fin_idx]
                    },
                    'curso': datos.get('curso', {})
                }
                
                # Generar documento
                doc_bytes = self._generar_documento_unico(datos_doc)
                
                # Nombre del archivo
                nombre_archivo = f"{nombre_alumno}_parte{pagina}_modulos{modulo_idx+1}-{fin_idx}.docx"
                zf.writestr(nombre_archivo, doc_bytes)
                print(f"    ✓ {nombre_archivo}")
                
                modulo_idx = fin_idx
                pagina += 1
        
        zip_buffer.seek(0)
        print(f"✓ ZIP con {pagina-1} documentos generado")
        return zip_buffer.getvalue()
    
    def _rellenar_campos(self, xml: str, datos: Dict) -> str:
        """
        Rellena campos de formulario (IGUAL que código grupal)
        """
        
        alumno = datos.get('alumno', {})
        curso = datos.get('curso', {})
        modulos = alumno.get('modulos', [])
        
        print(f"📝 Rellenando campos de formulario...")
        
        valores = [
            '',  # 1. Expediente
            alumno.get('nombre', ''),  # 2. Nombre
            alumno.get('dni', ''),  # 3. DNI
            curso.get('nombre', ''),  # 4. Curso
            '',  # 5. Código certificado
            'INTERPROS NEXT GENERATION SLU',  # 6. Centro
            '26615',  # 7. Código centro
            'C/ DR. SEVERO OCHOA, 21, BJ',  # 8. Dirección
            'AVILÉS',  # 9. Localidad
            '33400',  # 10. CP
            'ASTURIAS',  # 11. Provincia
        ]
        
        # Agregar hasta 6 módulos
        for i in range(6):
            if i < len(modulos):
                mod = modulos[i]
                valores.extend([
                    mod.get('codigo', ''),
                    str(mod.get('horas_totales', 0)),
                    mod.get('nombre', ''),
                    str(mod.get('horas_asistidas', 0))
                ])
            else:
                valores.extend(['', '', '', ''])
        
        print(f"  → Valores preparados: {len(valores)}")
        
        campos_largos = [3]  # DNI
        for i in range(6):
            idx_nombre = 11 + (i * 4) + 2
            campos_largos.append(idx_nombre)
        
        contador = 0
        
        def rellenar_campo(match):
            nonlocal contador
            
            if contador >= len(valores):
                return match.group(0)
            
            valor = str(valores[contador]) if valores[contador] is not None else ''
            es_campo_largo = contador in campos_largos
            campo_actual = contador + 1
            contador += 1
            
            campo_completo = match.group(0)
            
            # Buscar el separate (IGUAL que código grupal)
            separate_match = re.search(
                r'<w:fldChar\s+w:fldCharType="separate"[^>]*/>',
                campo_completo
            )
            
            if not separate_match:
                return campo_completo
            
            pos_separate = separate_match.end()
            
            # Obtener formato original (IGUAL que código grupal)
            formato_match = re.search(
                r'<w:rPr>(.*?)</w:rPr>',
                campo_completo[:pos_separate],
                re.DOTALL
            )
            
            if formato_match:
                formato_original = formato_match.group(1)
                
                # Ajustar tamaño si es campo largo
                if es_campo_largo and len(valor) > 50:
                    formato_ajustado = re.sub(r'<w:sz w:val="\d+"/>', '<w:sz w:val="16"/>', formato_original)
                    formato_ajustado = re.sub(r'<w:szCs w:val="\d+"/>', '<w:szCs w:val="16"/>', formato_ajustado)
                    formato = f'<w:rPr>{formato_ajustado}</w:rPr>'
                else:
                    formato = f'<w:rPr>{formato_original}</w:rPr>'
            else:
                # Formato por defecto
                sz = '16' if (es_campo_largo and len(valor) > 50) else '20'
                formato = f'<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            
            # Reemplazar contenido (IGUAL que código grupal)
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
        
        # Patrón (IGUAL que código grupal)
        patron = (
            r'<w:fldChar\s+w:fldCharType="begin"[^>]*>.*?'
            r'<w:fldChar\s+w:fldCharType="end"[^>]*/?>(?:</w:r>)?'
        )
        
        xml_modificado = re.sub(patron, rellenar_campo, xml, flags=re.DOTALL)
        
        print(f"  ✓ {contador} campos procesados")
        
        return xml_modificado
    
    def _rellenar_tabla_modulos(self, xml: str, datos: Dict) -> str:
        """
        Rellena tabla de módulos (basado en código grupal)
        """
        
        alumno = datos.get('alumno', {})
        modulos = alumno.get('modulos', [])
        
        if not modulos:
            return xml
        
        print(f"📊 Rellenando tabla con {len(modulos)} módulo(s)...")
        
        # Buscar tabla
        idx_modulos = xml.find('Módulos')
        if idx_modulos < 0:
            print("  ⚠ No se encontró 'Módulos'")
            return xml
        
        idx_tabla = xml.find('<w:tbl>', idx_modulos)
        if idx_tabla < 0:
            print("  ⚠ No se encontró tabla")
            return xml
        
        idx_fin_tabla = xml.find('</w:tbl>', idx_tabla)
        if idx_fin_tabla < 0:
            return xml
        
        tabla_completa = xml[idx_tabla:idx_fin_tabla + 8]
        
        # Extraer filas
        filas = re.findall(r'<w:tr[^>]*>(.*?)</w:tr>', tabla_completa, re.DOTALL)
        
        if len(filas) < 2:
            print("  ⚠ No hay filas de datos")
            return xml
        
        print(f"  → Tabla tiene {len(filas)} filas")
        
        tabla_modificada = tabla_completa
        
        # Rellenar hasta 6 módulos
        for idx_mod, modulo in enumerate(modulos[:6]):
            if idx_mod + 1 >= len(filas):
                print(f"  ⚠ No hay fila para módulo {idx_mod + 1}")
                break
            
            fila_original = filas[idx_mod + 1]
            
            codigo = modulo.get('codigo', '')
            horas_totales = modulo.get('horas_totales', 0)
            nombre = modulo.get('nombre', '')
            horas_asistidas = modulo.get('horas_asistidas', 0)
            
            print(f"  • {codigo}: {nombre[:25]}...")
            
            fila_nueva = self._crear_fila_modulo(fila_original, codigo, horas_totales, nombre, horas_asistidas)
            
            tabla_modificada = tabla_modificada.replace(
                f'<w:tr{fila_original.split("</w:tr>")[0]}</w:tr>',
                fila_nueva,
                1
            )
        
        xml_modificado = xml.replace(tabla_completa, tabla_modificada, 1)
        
        print(f"  ✓ Tabla actualizada")
        
        return xml_modificado
    
    def _crear_fila_modulo(self, fila_original: str, codigo: str, horas: int, nombre: str, asistencia: float) -> str:
        """Crea fila de módulo (basado en código grupal)"""
        
        celdas = re.findall(r'(<w:tc>.*?</w:tc>)', fila_original, re.DOTALL)
        
        if len(celdas) < 4:
            return f'<w:tr>{fila_original}</w:tr>'
        
        valores = [str(codigo), str(horas), nombre, str(round(asistencia, 2))]
        celdas_modificadas = []
        
        for idx, celda in enumerate(celdas[:4]):
            if idx < len(valores):
                celda_nueva = self._insertar_texto_en_celda(celda, valores[idx], fuente_pequena=(idx == 2))
                celdas_modificadas.append(celda_nueva)
            else:
                celdas_modificadas.append(celda)
        
        # Resto de celdas
        celdas_modificadas.extend(celdas[4:])
        
        match_tr = re.search(r'<w:tr([^>]*)>', fila_original)
        atributos_tr = match_tr.group(1) if match_tr else ''
        
        fila_nueva = f'<w:tr{atributos_tr}>{"".join(celdas_modificadas)}</w:tr>'
        
        return fila_nueva
    
    def _insertar_texto_en_celda(self, celda: str, texto: str, fuente_pequena: bool = False) -> str:
        """Inserta texto en celda (basado en código grupal)"""
        
        match_p = re.search(r'(<w:p[^>]*>)(.*?)(</w:p>)', celda, re.DOTALL)
        
        if not match_p:
            return celda
        
        inicio_p = match_p.group(1)
        contenido_p = match_p.group(2)
        fin_p = match_p.group(3)
        
        if fuente_pequena:
            formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        else:
            formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        
        nuevo_run = f'<w:r>{formato}<w:t xml:space="preserve">{texto}</w:t></w:r>'
        
        if '<w:r>' in contenido_p:
            contenido_nuevo = re.sub(r'<w:r>.*?</w:r>', '', contenido_p, flags=re.DOTALL)
            contenido_nuevo = contenido_nuevo + nuevo_run
        else:
            contenido_nuevo = contenido_p + nuevo_run
        
        celda_nueva = celda.replace(
            f'{inicio_p}{contenido_p}{fin_p}',
            f'{inicio_p}{contenido_nuevo}{fin_p}',
            1
        )
        
        return celda_nueva
    
    def _crear_docx(self, xml_modificado: str) -> bytes:
        """Crea DOCX (IGUAL que código grupal)"""
        
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