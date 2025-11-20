"""
Generador Word SEGURO - Rellena campos de formulario sin romper el XML
"""
import io
import re
import zipfile
from typing import Dict


class WordGeneratorSEPE:
    """
    Generador que rellena campos de formulario de forma SEGURA
    
    Estrategia:
    - NO elimina los campos de formulario
    - Solo añade el contenido entre begin y separate
    - Mantiene toda la estructura intacta
    """
    
    def __init__(self, plantilla_bytes: bytes = None, es_xml: bool = False):
        self.plantilla_bytes = plantilla_bytes
        self.es_xml = es_xml
        self.plantilla_zip_parts = {}
        
        if plantilla_bytes and not es_xml:
            try:
                with zipfile.ZipFile(io.BytesIO(plantilla_bytes), 'r') as zf:
                    for item in zf.namelist():
                        self.plantilla_zip_parts[item] = zf.read(item)
            except Exception as e:
                print(f" Error leyendo DOCX: {e}")
    
    def generar_informe_individual(self, datos: Dict) -> bytes:
        """Genera informe rellenando campos de formulario y tabla de módulos"""
        
        # Obtener document.xml
        if 'word/document.xml' in self.plantilla_zip_parts:
            xml_string = self.plantilla_zip_parts['word/document.xml'].decode('utf-8')
        else:
            raise Exception("No se pudo leer word/document.xml de la plantilla")
        
        # Rellenar campos de formulario
        xml_modificado = self._rellenar_campos(xml_string, datos)
        
        # Rellenar tabla de módulos
        xml_modificado = self._rellenar_tabla_modulos(xml_modificado, datos)
        
        # Crear DOCX
        return self._crear_docx(xml_modificado)
    
    def _rellenar_campos(self, xml: str, datos: Dict) -> str:
        """
        Rellena campos de formulario MANTENIENDO su estructura
        
        Los campos tienen esta estructura:
        <w:fldChar w:fldCharType="begin"/>
        <w:fldChar w:fldCharType="separate"/>
        [AQUÍ VA EL CONTENIDO] ← Añadimos aquí
        <w:fldChar w:fldCharType="end"/>
        
        NO eliminamos nada, solo añadimos contenido entre separate y end
        """
        
        alumno = datos.get('alumno', {})
        curso = datos.get('curso', {})
        modulos = alumno.get('modulos', [])
        
        # Valores en orden
        # IMPORTANTE: Solo hay 11 campos antes de la tabla de módulos
        valores = [
            '',                                         # 1. Expediente
            alumno.get('nombre', ''),                   # 2. Nombre
            alumno.get('dni', ''),                      # 3. DNI
            curso.get('nombre', ''),                    # 4. Curso (completo, sin recortar)
            '',                                         # 5. Código cert
            'INTERPROS NEXT GENERATION SLU',            # 6. Centro
            '26615',                                    # 7. Código centro
            'C/ DR. SEVERO OCHOA, 21, BJ',              # 8. Dirección
            'AVILÉS',                                   # 9. Localidad
            '33400',                                    # 10. CP
            'ASTURIAS',                                 # 11. Provincia
            # NO HAY campos 12-14, la tabla empieza directo en campo 12
        ]
        
        # Añadir TABLA DE MÓDULOS (5 filas x 4 campos = 20 campos)
        # Cada fila: Código | Horas Totales | Nombre Módulo | Horas Asistencia
        
        print(f"\n=== DEBUG: Módulos del alumno ===")
        print(f"Total módulos: {len(modulos)}")
        
        for i in range(5):  # 5 filas de módulos
            if i < len(modulos):
                # Módulo existe
                mod = modulos[i]
                
                codigo = mod.get('codigo', '')
                horas_tot = str(mod.get('horas_totales', 0))
                nombre = mod.get('nombre', '')
                horas_asist = str(mod.get('horas_asistidas', 0))
                
                print(f"\nMódulo {i+1}:")
                print(f"  Código: '{codigo}'")
                print(f"  Horas totales: '{horas_tot}'")
                print(f"  Nombre: '{nombre[:50]}...'")
                print(f"  Horas asistencia: '{horas_asist}'")
                
                valores.extend([
                    codigo,       # Código (ej: MF0969_1)
                    horas_tot,    # Horas totales (ej: 165)
                    nombre,       # Nombre completo del módulo
                    horas_asist   # Horas asistidas por alumno
                ])
            else:
                # No hay más módulos, dejar vacío
                valores.extend(['', '', '', ''])
        
        print(f"\n=== Total de valores preparados: {len(valores)} ===\n")
        
        # Debug: Mostrar TODOS los valores
        print("=== MAPEO DE CAMPOS ===")
        nombres_campos = [
            "1. Expediente", "2. Nombre alumno", "3. DNI", "4. Curso",
            "5. Código cert", "6. Centro", "7. Código centro", "8. Dirección",
            "9. Localidad", "10. CP", "11. Provincia",
            # Campos 12-31: Tabla de módulos (5 filas x 4 campos)
            "12. M1-Código", "13. M1-Horas", "14. M1-Nombre", "15. M1-Asist",
            "16. M2-Código", "17. M2-Horas", "18. M2-Nombre", "19. M2-Asist",
            "20. M3-Código", "21. M3-Horas", "22. M3-Nombre", "23. M3-Asist",
            "24. M4-Código", "25. M4-Horas", "26. M4-Nombre", "27. M4-Asist",
            "28. M5-Código", "29. M5-Horas", "30. M5-Nombre", "31. M5-Asist",
        ]
        
        for i, (nombre, valor) in enumerate(zip(nombres_campos, valores)):
            valor_corto = str(valor)[:50] if valor else '(vacío)'
            print(f"Campo {i+1:2d} ({nombre:20s}): {valor_corto}")
        print("="*60 + "\n")
        
        # Índices de campos que pueden ser largos (necesitan fuente pequeña)
        campos_largos = [3]  # Campo 4 (índice 3) = Nombre del curso
        # Añadir los nombres de módulos
        # Nombres están en posiciones: 14, 18, 22, 26, 30 (campos en el doc)
        # En el array son índices: 13, 17, 21, 25, 29
        for i in range(5):
            idx_nombre_modulo = 11 + (i * 4) + 2  # 13, 17, 21, 25, 29
            campos_largos.append(idx_nombre_modulo)
        
        contador = 0
        
        def rellenar_campo(match):
            """
            Rellena UN campo de formulario
            
            Match contiene:
            - begin
            - separate  
            - contenido existente
            - end
            """
            nonlocal contador
            
            if contador >= len(valores):
                return match.group(0)  # No modificar
            
            valor = valores[contador]
            es_campo_largo = contador in campos_largos
            contador += 1
            
            # Estructura completa del campo
            campo_completo = match.group(0)
            
            # Buscar dónde está el separate
            separate_match = re.search(r'<w:fldChar\s+w:fldCharType="separate"[^>]*/>', campo_completo)
            
            if not separate_match:
                return campo_completo  # No tiene separate, no modificar
            
            # Posición donde insertar el valor (después de separate)
            pos_separate = separate_match.end()
            
            # Extraer formato del campo si existe
            formato_match = re.search(r'<w:rPr>(.*?)</w:rPr>', campo_completo[:pos_separate], re.DOTALL)
            if formato_match:
                formato_base = formato_match.group(1)
                
                # Si es campo largo, ajustar tamaño de fuente
                if es_campo_largo and len(valor) > 50:
                    # Reemplazar el tamaño de fuente a algo más pequeño
                    formato_base = re.sub(r'<w:sz\s+w:val="[^"]*"/>', '<w:sz w:val="16"/>', formato_base)
                    formato_base = re.sub(r'<w:szCs\s+w:val="[^"]*"/>', '<w:szCs w:val="16"/>', formato_base)
                    
                    # Si no tiene sz, añadirlo
                    if '<w:sz' not in formato_base:
                        formato_base += '<w:sz w:val="16"/><w:szCs w:val="16"/>'
                        
                        
                
                formato = f'<w:rPr>{formato_base}</w:rPr>'
            else:
                # Formato por defecto
                if es_campo_largo and len(valor) > 50:
                    # Fuente más pequeña para textos largos
                    formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
                else:
                    formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
            
            # Buscar si ya hay contenido después de separate
            # Patrón: desde separate hasta antes de end
            contenido_existente_match = re.search(
                r'<w:fldChar\s+w:fldCharType="separate"[^>]*/>(.*?)<w:fldChar\s+w:fldCharType="end"',
                campo_completo,
                re.DOTALL
            )
            
            if contenido_existente_match:
                # Hay contenido, reemplazarlo
                contenido_antiguo = contenido_existente_match.group(1)
                nuevo_contenido = f'<w:r>{formato}<w:t xml:space="preserve">{valor}</w:t></w:r>'
                
                campo_nuevo = campo_completo.replace(contenido_antiguo, nuevo_contenido, 1)
                return campo_nuevo
            else:
                # No hay contenido, añadirlo
                nuevo_run = f'<w:r>{formato}<w:t xml:space="preserve">{valor}</w:t></w:r>'
                
                # Insertar después de separate
                campo_nuevo = (
                    campo_completo[:pos_separate] +
                    nuevo_run +
                    campo_completo[pos_separate:]
                )
                return campo_nuevo
        
        # Patrón: campo completo desde begin hasta end (inclusive)
        patron = (
            r'<w:fldChar\s+w:fldCharType="begin"[^>]*>.*?'
            r'<w:fldChar\s+w:fldCharType="end"[^>]*/?>(?:</w:r>)?'
        )
        
        xml_modificado = re.sub(patron, rellenar_campo, xml, flags=re.DOTALL)
        
        print(f" {contador} campos rellenados")
        
        return xml_modificado
    
    def _rellenar_tabla_modulos(self, xml: str, datos: Dict) -> str:
        """
        Rellena la tabla de módulos con:
        - Código del módulo (MF0969_1, etc.)
        - Horas totales del módulo
        - Nombre del módulo
        - Horas de asistencia del alumno
        """
        
        alumno = datos.get('alumno', {})
        modulos = alumno.get('modulos', [])
        
        if not modulos:
            return xml
        
        # Buscar la tabla de módulos (está después del header "Módulos")
        # La tabla tiene celdas con formato específico
        
        # Patrón para encontrar filas de la tabla después de "Módulos"
        # Buscamos la primera tabla después de "Módulos"
        
        idx_modulos = xml.find('Módulos')
        if idx_modulos < 0:
            print(" No se encontró 'Módulos' en el documento")
            return xml
        
        # Buscar la siguiente tabla después de "Módulos"
        # Las tablas están marcadas con <w:tbl>
        idx_tabla = xml.find('<w:tbl>', idx_modulos)
        if idx_tabla < 0:
            print(" No se encontró tabla después de 'Módulos'")
            return xml
        
        # Encontrar el final de la tabla
        idx_fin_tabla = xml.find('</w:tbl>', idx_tabla)
        if idx_fin_tabla < 0:
            return xml
        
        # Extraer la tabla completa
        tabla_completa = xml[idx_tabla:idx_fin_tabla + 8]
        
        # Buscar las filas de datos (después del header)
        # Las filas están marcadas con <w:tr>
        filas = re.findall(r'<w:tr[^>]*>(.*?)</w:tr>', tabla_completa, re.DOTALL)
        
        if len(filas) < 2:  # Necesitamos al menos header + 1 fila de datos
            print(" No se encontraron filas de datos en la tabla")
            return xml
        
        # La primera fila es el header, las siguientes son para datos
        # Necesitamos modificar las filas 2, 3, 4 (índices 1, 2, 3)
        
        tabla_modificada = tabla_completa
        
        for idx_mod, modulo in enumerate(modulos[:5]):  # Máximo 5 módulos
            if idx_mod + 1 >= len(filas):
                break
            
            fila_original = filas[idx_mod + 1]  # +1 porque saltamos el header
            
            # Datos del módulo
            codigo = modulo.get('codigo', '')
            horas_totales = modulo.get('horas_totales', 0)
            nombre = modulo.get('nombre', '')
            horas_asistidas = modulo.get('horas_asistidas', 0)
            
            # Crear nueva fila con los datos
            fila_nueva = self._crear_fila_modulo(fila_original, codigo, horas_totales, nombre, horas_asistidas)
            
            # Reemplazar fila en la tabla
            tabla_modificada = tabla_modificada.replace(
                f'<w:tr{fila_original.split("</w:tr>")[0]}</w:tr>',
                fila_nueva,
                1
            )
        
        # Reemplazar tabla completa en el XML
        xml_modificado = xml.replace(tabla_completa, tabla_modificada, 1)
        
        print(f"✓ Tabla de módulos rellenada con {len(modulos)} módulos")
        
        return xml_modificado
    
    def _crear_fila_modulo(self, fila_original: str, codigo: str, horas: int, nombre: str, asistencia: int) -> str:
        """
        Crea una fila de módulo con los datos
        
        Estructura esperada: 4 celdas
        1. Código (MF0969_1)
        2. Horas totales (165)
        3. Nombre del módulo (largo)
        4. Horas de asistencia (165)
        """
        
        # Extraer las celdas de la fila original
        celdas = re.findall(r'(<w:tc>.*?</w:tc>)', fila_original, re.DOTALL)
        
        if len(celdas) < 4:
            print(f"⚠ Fila tiene {len(celdas)} celdas, se esperaban 4")
            return f'<w:tr>{fila_original}</w:tr>'
        
        # Modificar cada celda
        valores = [str(codigo), str(horas), nombre, str(asistencia)]
        celdas_modificadas = []
        
        for idx, celda in enumerate(celdas[:4]):
            if idx < len(valores):
                celda_nueva = self._insertar_texto_en_celda(celda, valores[idx], fuente_pequena=(idx == 2))
                celdas_modificadas.append(celda_nueva)
            else:
                celdas_modificadas.append(celda)
        
        # Reconstruir fila
        # Extraer atributos de la fila original
        match_tr = re.search(r'<w:tr([^>]*)>', fila_original)
        atributos_tr = match_tr.group(1) if match_tr else ''
        
        fila_nueva = f'<w:tr{atributos_tr}>{"".join(celdas_modificadas)}</w:tr>'
        
        return fila_nueva
    
    def _insertar_texto_en_celda(self, celda: str, texto: str, fuente_pequena: bool = False) -> str:
        """Inserta texto en una celda de tabla"""
        
        # Buscar el párrafo dentro de la celda
        match_p = re.search(r'(<w:p[^>]*>)(.*?)(</w:p>)', celda, re.DOTALL)
        
        if not match_p:
            return celda
        
        inicio_p = match_p.group(1)
        contenido_p = match_p.group(2)
        fin_p = match_p.group(3)
        
        # Determinar tamaño de fuente
        if fuente_pequena:
            # Para nombres largos de módulo
            formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="16"/><w:szCs w:val="16"/></w:rPr>'
        else:
            # Para códigos y números
            formato = '<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>'
        
        # Crear nuevo run con el texto
        nuevo_run = f'<w:r>{formato}<w:t xml:space="preserve">{texto}</w:t></w:r>'
        
        # Si el párrafo tiene contenido, reemplazarlo
        # Si no, añadirlo
        if '<w:r>' in contenido_p:
            # Reemplazar todos los runs existentes
            contenido_nuevo = re.sub(r'<w:r>.*?</w:r>', '', contenido_p, flags=re.DOTALL)
            contenido_nuevo = contenido_nuevo + nuevo_run
        else:
            contenido_nuevo = contenido_p + nuevo_run
        
        # Reconstruir celda
        celda_nueva = celda.replace(
            f'{inicio_p}{contenido_p}{fin_p}',
            f'{inicio_p}{contenido_nuevo}{fin_p}',
            1
        )
        
        return celda_nueva
    
    def _crear_docx(self, xml_modificado: str) -> bytes:
        """Crea DOCX manteniendo TODAS las partes originales"""
        
        if not self.plantilla_zip_parts:
            raise Exception("No hay plantilla cargada")
        
        output = io.BytesIO()
        
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:
            for nombre, contenido in self.plantilla_zip_parts.items():
                if nombre == 'word/document.xml':
                    # Este es el único modificado
                    docx.writestr(nombre, xml_modificado.encode('utf-8'))
                else:
                    # TODO lo demás se copia tal cual
                    docx.writestr(nombre, contenido)
        
        output.seek(0)
        return output.getvalue()