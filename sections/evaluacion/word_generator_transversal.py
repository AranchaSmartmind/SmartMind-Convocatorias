"""
GENERADOR DE ACTAS TRANSVERSALES (FCOO03)
==========================================
Genera actas de evaluación final para competencias transversales
"""
import zipfile
import re
from io import BytesIO
from typing import Dict, List


class WordGeneratorTransversal:
    """Genera documentos Word para actas transversales"""
    
    def __init__(self, plantilla_bytes: bytes):
        """
        Inicializa el generador con la plantilla
        
        Args:
            plantilla_bytes: Bytes de la plantilla Word
        """
        self.plantilla_bytes = plantilla_bytes
        
        self.plantilla_parts = {}
        with zipfile.ZipFile(BytesIO(plantilla_bytes), 'r') as zf:
            for item in zf.namelist():
                self.plantilla_parts[item] = zf.read(item)
    
    def generar_acta(self, datos: Dict) -> bytes:
        """
        Genera el acta transversal completa
        
        Args:
            datos: Diccionario con todos los datos extraídos por TransversalesProcessor
        
        Returns:
            bytes: Documento Word generado
        """

        valores = self._preparar_valores(datos)
        
        print(f"\n Generando acta transversal con {len(valores)} campos")
        
        xml_content = self.plantilla_parts['word/document.xml'].decode('utf-8')
        xml_modificado = self._rellenar_campos(xml_content, valores)
        
        output = BytesIO()
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as docx:
            for nombre, contenido in self.plantilla_parts.items():
                if nombre == 'word/document.xml':
                    docx.writestr(nombre, xml_modificado.encode('utf-8'))
                else:
                    docx.writestr(nombre, contenido)
        
        output.seek(0)
        print(" Acta transversal generada correctamente")
        
        return output.getvalue()
    
    def _preparar_valores(self, datos: Dict) -> List[str]:
        """
        Prepara lista de 92 valores para rellenar los campos
        
        Estructura HORIZONTAL (4 campos por alumno):
        - Campos 11, 15, 19, 23... = DNI
        - Campos 12, 16, 20, 24... = Nombre
        - Campos 13, 17, 21, 25... = Horas
        - Campos 14, 18, 22, 26... = Calificación
        
        Args:
            datos: Datos extraídos por el procesador
        
        Returns:
            Lista de 92 valores
        """
        valores = []
        
        valores.append(datos.get('campo_1_convocatoria', ''))
        valores.append(datos.get('campo_2_accion', ''))
        valores.append(datos.get('campo_3_especialidad', ''))
        valores.append(datos.get('campo_4_codigo', ''))
        valores.append(datos.get('campo_5_centro', ''))
        valores.append(datos.get('campo_6_duracion', ''))
        valores.append(datos.get('campo_7_actividades', ''))
        valores.append(datos.get('campo_8_modalidad', ''))
        valores.append(datos.get('campo_9_fecha_inicio', ''))
        valores.append(datos.get('campo_10_fecha_fin', ''))

        alumnos = datos.get('alumnos', [])
        
        for i in range(20):
            if i < len(alumnos):
                alumno = alumnos[i]

                valores.append(str(alumno.get('dni', '')))
                valores.append(str(alumno.get('nombre', '')))
                valores.append(str(alumno.get('horas_actividades', '')))
                valores.append(str(alumno.get('calificacion_final', '')))
            else:

                valores.extend(['', '', '', ''])

        valores.append('')
        valores.append('')
        
        print(f"   Campos preparados: {len(valores)}")
        print(f"   Alumnos incluidos: {len(alumnos)}")
        print(f"   Estructura: 4 campos/alumno (DNI, Nombre, Horas, Calif)")
        
        return valores
    
    def _rellenar_campos(self, xml_content: str, valores: List[str]) -> str:
        """
        Rellena los campos del documento XML
        
        Args:
            xml_content: Contenido XML del documento
            valores: Lista de valores para cada campo
        
        Returns:
            XML modificado
        """
        contador = 0
        
        def rellenar_campo(match):
            nonlocal contador
            
            if contador >= len(valores):
                return match.group(0)
            
            valor = valores[contador]
            contador += 1
            
            campo_completo = match.group(0)
            
            separate_match = re.search(
                r'<w:fldChar\s+w:fldCharType="separate"[^>]*/>',
                campo_completo
            )
            
            if not separate_match:
                return campo_completo
            
            contenido_match = re.search(
                r'(<w:fldChar\s+w:fldCharType="separate"[^>]*/>)(.*?)(<w:fldChar\s+w:fldCharType="end")',
                campo_completo,
                re.DOTALL
            )
            
            if not contenido_match:
                return campo_completo
            
            separate_tag = contenido_match.group(1)
            end_tag = contenido_match.group(3)
            
            nuevo_contenido = f'<w:r><w:t xml:space="preserve">{self._escapar_xml(valor)}</w:t></w:r>'

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
        
        xml_modificado = re.sub(patron, rellenar_campo, xml_content, flags=re.DOTALL)
        
        print(f"   {contador} campos rellenados en el documento")
        
        return xml_modificado
    
    def _escapar_xml(self, texto: str) -> str:
        """Escapa caracteres especiales para XML"""
        if not texto:
            return ''
        
        texto = str(texto)
        texto = texto.replace('&', '&amp;')
        texto = texto.replace('<', '&lt;')
        texto = texto.replace('>', '&gt;')
        texto = texto.replace('"', '&quot;')
        texto = texto.replace("'", '&apos;')
        
        return texto

if __name__ == "__main__":
    import sys
    sys.path.insert(0, '/mnt/user-data/outputs')
    
    from transversales_processor import TransversalesProcessor
    
    print("\n PRUEBA DEL GENERADOR TRANSVERSAL")
    print("=" * 80)
    
    with open('/mnt/user-data/uploads/2024_1339_CTRL_Tareas_AREA.xlsx', 'rb') as f:
        control_bytes = f.read()
    
    with open('/mnt/user-data/uploads/1339_Cronograma_v01__1_.xlsx', 'rb') as f:
        cronograma_bytes = f.read()
    
    processor = TransversalesProcessor()
    datos = processor.cargar_datos(control_bytes, cronograma_bytes)
    
    print(f" Datos extraídos: {datos['total_alumnos']} alumnos")
    
    with open('/mnt/user-data/outputs/plantilla_transversal_oficial.docx', 'rb') as f:
        plantilla_bytes = f.read()
    
    print(f" Plantilla cargada: {len(plantilla_bytes):,} bytes")
    
    generador = WordGeneratorTransversal(plantilla_bytes)
    documento = generador.generar_acta(datos)
    
    with open('/mnt/user-data/outputs/ACTA_TRANSVERSAL_GENERADA.docx', 'wb') as f:
        f.write(documento)
    
    print(f"\n Documento generado: {len(documento):,} bytes")
    print(" Guardado en: ACTA_TRANSVERSAL_GENERADA.docx")