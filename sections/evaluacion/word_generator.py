"""
Generador Word Inteligente - VERSIÓN FINAL
Busca y reemplaza directamente en el XML oficial sin necesidad de marcadores previos
"""
import io
import re
import zipfile
from typing import Dict


class WordGeneratorSEPE:
    """Generador que trabaja directamente con tu plantilla oficial"""
    
    def __init__(self, plantilla_bytes: bytes = None, es_xml: bool = False):
        """
        Args:
            plantilla_bytes: Bytes de tu plantilla oficial
            es_xml: True si es XML puro, False si es DOCX
        """
        self.plantilla_bytes = plantilla_bytes
        self.es_xml = es_xml
        self.plantilla_zip_parts = {}
        self.es_word_2003_xml = False
        
        if plantilla_bytes:
            # Detectar si es un XML de Word 2003 (como tu .doc)
            try:
                contenido_inicio = plantilla_bytes[:500].decode('utf-8', errors='ignore')
                if '<?xml' in contenido_inicio and 'pkg:package' in contenido_inicio:
                    self.es_word_2003_xml = True
                    self.es_xml = True
                    print("✓ Detectado: XML de Word 2003")
            except:
                pass
            
            # Si no es XML, intentar extraer como DOCX
            if not self.es_xml and not self.es_word_2003_xml:
                try:
                    with zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r') as zf:
                        for item in zf.namelist():
                            self.plantilla_zip_parts[item] = zf.read(item)
                    print("✓ Detectado: DOCX (ZIP)")
                except:
                    print("⚠ No se pudo leer como DOCX, intentando como XML...")
                    self.es_xml = True
    
    def generar_informe_individual(self, datos: Dict) -> bytes:
        """Genera informe desde tu plantilla oficial"""
        
        # Obtener el XML del documento
        if self.es_word_2003_xml and self.plantilla_bytes:
            # Es un XML de Word 2003 - extraer el document.xml interno
            xml_string = self._extraer_document_xml_de_word2003(self.plantilla_bytes)
        elif self.es_xml and self.plantilla_bytes:
            # XML puro
            xml_string = self.plantilla_bytes.decode('utf-8')
        elif 'word/document.xml' in self.plantilla_zip_parts:
            # DOCX - extraer document.xml
            xml_string = self.plantilla_zip_parts['word/document.xml'].decode('utf-8')
        else:
            raise Exception("No se pudo leer la plantilla")
        
        # Realizar reemplazos inteligentes
        xml_modificado = self._reemplazar_datos_en_xml(xml_string, datos)
        
        # Crear DOCX
        if self.es_word_2003_xml:
            # Para Word 2003 XML, extraer todas las partes y reconstruir
            return self._crear_docx_desde_word2003_xml(xml_modificado)
        elif self.es_xml:
            return self._crear_docx_desde_xml(xml_modificado)
        else:
            return self._crear_docx_con_partes_originales(xml_modificado)
    
    def _extraer_document_xml_de_word2003(self, word2003_bytes: bytes) -> str:
        """
        Extrae el document.xml de un archivo Word 2003 XML
        
        Estos archivos tienen estructura:
        <pkg:package>
            <pkg:part pkg:name="/word/document.xml">
                <pkg:xmlData>
                    [AQUÍ ESTÁ EL CONTENIDO]
                </pkg:xmlData>
            </pkg:part>
        </pkg:package>
        """
        import re
        
        xml_completo = word2003_bytes.decode('utf-8')
        
        # Buscar la parte que contiene word/document.xml
        patron = r'<pkg:part\s+pkg:name="/word/document\.xml"[^>]*>.*?<pkg:xmlData>(.*?)</pkg:xmlData>.*?</pkg:part>'
        match = re.search(patron, xml_completo, re.DOTALL)
        
        if match:
            document_xml = match.group(1)
            return document_xml.strip()
        else:
            # Si no se encuentra, devolver todo (mejor que nada)
            return xml_completo
    
    def _crear_docx_desde_word2003_xml(self, document_xml: str) -> bytes:
        """
        Crea un DOCX desde un XML de Word 2003
        
        Toma el document.xml extraído y crea un DOCX válido moderno
        """
        # Reconstruir el paquete Word 2003 XML completo con el document.xml modificado
        # O mejor aún, crear un DOCX moderno
        
        return self._crear_docx_desde_xml(document_xml)
    
    def _reemplazar_datos_en_xml(self, xml: str, datos: Dict) -> str:
        """Reemplaza datos en el XML de forma inteligente"""
        
        alumno = datos.get('alumno', {})
        curso = datos.get('curso', {})
        modulos = alumno.get('modulos', [])
        
        # ESTRATEGIA: Buscar patrones en el XML y reemplazarlos
        
        # 1. Expediente - buscar celda vacía después de "Número de expediente:"
        xml = re.sub(
            r'(<w:t>Número de expediente:</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>E-2024-100454</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 2. Nombre del alumno - después de "El/la alumno/a"
        nombre = alumno.get('nombre', '')
        xml = re.sub(
            r'(<w:t>El/la alumno/a</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{nombre}</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 3. DNI - después de "con DNI./NIE./Pasaporte:"
        dni = alumno.get('dni', '')
        xml = re.sub(
            r'(<w:t>/Pasaporte:</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{dni}</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 4. Certificado profesional
        nombre_curso = curso.get('nombre', '')[:80]  # Limitar longitud
        xml = re.sub(
            r'(<w:t>Certificado profesional:</w:t></w:r>.*?<w:tc>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{nombre_curso}</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 5. Código del curso
        xml = re.sub(
            r'(<w:t>Código</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>ADGG0408</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 6. Centro formativo
        xml = re.sub(
            r'(<w:t>Centro formativo:</w:t></w:r>.*?<w:tc>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>INTERPROS NEXT GENERATION SLU</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 7. Código centro
        xml = re.sub(
            r'(<w:t>Código</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>26615</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 8. Dirección
        xml = re.sub(
            r'(<w:t>Dirección:</w:t></w:r>.*?<w:tc>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>C/ DR. SEVERO OCHOA, 21, BJ</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 9. Localidad
        xml = re.sub(
            r'(<w:t>Localidad:</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>AVILÉS</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 10. Código Postal
        xml = re.sub(
            r'(<w:t>Código Postal:</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>33400</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 11. Provincia
        xml = re.sub(
            r'(<w:t>Provincia:</w:t></w:r>.*?<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>(?:<w:pPr[^>]*>.*?</w:pPr>)?)',
            rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>ASTURIAS</w:t></w:r>',
            xml,
            count=1,
            flags=re.DOTALL
        )
        
        # 12. Horas de módulos (en la tabla de módulos)
        # Buscar las celdas de "Horas de asistencia" y añadir valores
        if len(modulos) >= 1:
            horas_1 = modulos[0].get('horas_asistidas', 0)
            # Añadir primera hora de asistencia
            xml = re.sub(
                r'(<w:t>Horas de\s*asistencia</w:t></w:r>.*?</w:tc>\s*<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>)',
                rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{horas_1}</w:t></w:r>',
                xml,
                count=1,
                flags=re.DOTALL
            )
        
        if len(modulos) >= 2:
            horas_2 = modulos[1].get('horas_asistidas', 0)
            xml = re.sub(
                r'(<w:t>Horas de\s*asistencia</w:t></w:r>.*?</w:tc>\s*<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>)',
                rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{horas_2}</w:t></w:r>',
                xml,
                count=1,
                flags=re.DOTALL
            )
        
        if len(modulos) >= 3:
            horas_3 = modulos[2].get('horas_asistidas', 0)
            xml = re.sub(
                r'(<w:t>Horas de\s*asistencia</w:t></w:r>.*?</w:tc>\s*<w:tc>.*?<w:tcPr>.*?</w:tcPr>)(<w:p[^>]*>)',
                rf'\1\2<w:r><w:rPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:rPr><w:t>{horas_3}</w:t></w:r>',
                xml,
                count=1,
                flags=re.DOTALL
            )
        
        return xml
    
    def _crear_docx_con_partes_originales(self, xml_modificado: str) -> bytes:
        """Crea DOCX manteniendo todas las partes originales"""
        output = io.BytesIO()
        
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:
            for nombre, contenido in self.plantilla_zip_parts.items():
                if nombre == 'word/document.xml':
                    docx.writestr(nombre, xml_modificado.encode('utf-8'))
                else:
                    docx.writestr(nombre, contenido)
        
        output.seek(0)
        return output.getvalue()
    
    def _crear_docx_desde_xml(self, xml_content: str) -> bytes:
        """
        Crea DOCX válido desde XML
        
        IMPORTANTE: Un archivo .docx es un ZIP que contiene:
        - [Content_Types].xml
        - _rels/.rels
        - word/document.xml (el que modificamos)
        - word/_rels/document.xml.rels
        - word/styles.xml
        - word/numbering.xml (opcional)
        - word/fontTable.xml
        - word/settings.xml
        - word/webSettings.xml
        - word/theme/ (carpeta con temas)
        - docProps/ (propiedades del documento)
        
        Si falta alguna parte crítica, Word no puede abrir el archivo.
        """
        
        # Si tenemos las partes originales del DOCX, usarlas
        if self.plantilla_zip_parts:
            return self._crear_docx_con_partes_originales(xml_content)
        
        # Si no, crear un DOCX mínimo pero COMPLETO
        output = io.BytesIO()
        
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:
            # 1. Content Types - CRÍTICO
            docx.writestr('[Content_Types].xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
    <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
    <Default Extension="xml" ContentType="application/xml"/>
    <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
    <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
    <Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/>
    <Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>
    <Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/>
    <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
    <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>''')
            
            # 2. Relationships principales - CRÍTICO
            docx.writestr('_rels/.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>''')
            
            # 3. Document.xml - EL CONTENIDO MODIFICADO
            docx.writestr('word/document.xml', xml_content.encode('utf-8'))
            
            # 4. Document relationships - CRÍTICO
            docx.writestr('word/_rels/document.xml.rels', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
</Relationships>''')
            
            # 5. Styles - NECESARIO
            docx.writestr('word/styles.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:docDefaults>
        <w:rPrDefault>
            <w:rPr>
                <w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>
                <w:sz w:val="22"/>
            </w:rPr>
        </w:rPrDefault>
        <w:pPrDefault/>
    </w:docDefaults>
</w:styles>''')
            
            # 6. Font Table - NECESARIO
            docx.writestr('word/fontTable.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:font w:name="Arial">
        <w:panose1 w:val="020B0604020202020204"/>
        <w:charset w:val="00"/>
        <w:family w:val="swiss"/>
        <w:pitch w:val="variable"/>
    </w:font>
</w:fonts>''')
            
            # 7. Settings - NECESARIO
            docx.writestr('word/settings.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:zoom w:percent="100"/>
    <w:defaultTabStop w:val="708"/>
</w:settings>''')
            
            # 8. Web Settings - NECESARIO
            docx.writestr('word/webSettings.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:webSettings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
    <w:optimizeForBrowser/>
</w:webSettings>''')
            
            # 9. Core Properties - NECESARIO
            docx.writestr('docProps/core.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
    <dc:creator>Sistema SEPE</dc:creator>
    <cp:lastModifiedBy>Sistema SEPE</cp:lastModifiedBy>
    <cp:revision>1</cp:revision>
</cp:coreProperties>''')
            
            # 10. App Properties - NECESARIO
            docx.writestr('docProps/app.xml', '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">
    <Application>Microsoft Office Word</Application>
    <DocSecurity>0</DocSecurity>
    <HyperlinksChanged>false</HyperlinksChanged>
    <SharedDoc>false</SharedDoc>
</Properties>''')
        
        output.seek(0)
        return output.getvalue()
    
    def _get_content_types(self) -> str:
        return '''<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>'''
    
    def _get_rels(self) -> str:
        return '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'''
    
    def _get_doc_rels(self) -> str:
        return '''<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/></Relationships>'''