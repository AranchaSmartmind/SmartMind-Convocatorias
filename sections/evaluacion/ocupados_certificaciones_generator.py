"""
Generador de Certificados Word para Ocupados
Rellena la plantilla oficial con los datos de cada alumno
"""

from docx import Document
from docx.shared import Pt
from typing import Dict
import io


class CertificacionOcupadosGenerator:
    """Genera certificados Word individuales"""
    
    def __init__(self, plantilla_path: str):
        """
        Inicializa el generador
        
        Args:
            plantilla_path: Ruta a la plantilla .docx
        """
        self.plantilla_path = plantilla_path
        
    def generar_certificado(self, datos: Dict) -> bytes:
        """
        Genera un certificado para un alumno
        
        Args:
            datos: Diccionario con todos los datos del alumno
            
        Returns:
            Bytes del documento Word generado
        """
        # Cargar plantilla
        doc = Document(self.plantilla_path)
        
        # Obtener la tabla principal (primera tabla)
        tabla = doc.tables[0]
        
        # CAMPO #1: Código del curso (Fila 7, Celda 4) - VACÍO por defecto
        self._rellenar_celda(tabla, 6, 3, "")
        
        # CAMPO #2: Número de expediente (Fila 7, Celda 18)
        self._rellenar_celda(tabla, 6, 17, datos['expediente'])
        
        # CAMPO #3: Nombre del director (Fila 9, Celda 4)
        self._rellenar_celda(tabla, 8, 3, datos['director'])
        
        # CAMPO #4: Nombre del centro (Fila 10, Celda 8)
        self._rellenar_celda(tabla, 9, 7, datos['centro'])
        
        # CAMPO #5: Código de centro (Fila 11, Celda 7)
        self._rellenar_celda(tabla, 10, 6, datos['codigo_centro'])
        
        # CAMPO #6: Nombre del alumno (Fila 16, Celda 6)
        self._rellenar_celda(tabla, 15, 5, datos['nombre_alumno'])
        
        # CAMPO #7: DNI del alumno (Fila 16, Celda 20)
        self._rellenar_celda(tabla, 15, 19, datos['dni_alumno'])
        
        # CAMPO #7.5: Nombre del módulo después de la frase (Fila 17)
        # "según consta en el Acta de Evaluación Final, ha superado el siguiente Módulo Profesional:"
        # Seguido del nombre completo del módulo en la siguiente fila (Fila 18)
        
        # Agregar el nombre del módulo en la fila 18 (vacía)
        fila_18 = tabla.rows[17]  # Fila 18 (índice 17)
        celda_modulo = fila_18.cells[0]  # Primera celda de la fila 18
        
        # Limpiar si tiene algo
        for p in celda_modulo.paragraphs:
            p.clear()
        
        # Agregar el nombre del módulo
        if celda_modulo.paragraphs:
            p = celda_modulo.paragraphs[0]
        else:
            p = celda_modulo.add_paragraph()
        
        run = p.add_run(datos['nombre_modulo'])
        run.font.size = Pt(9)  # Fuente más pequeña (9pt) para que entre todo
        
        # CAMPO #8: Código del módulo (Fila 19, Celda 2)
        self._rellenar_celda(tabla, 18, 1, datos['codigo_modulo'])
        
        # CAMPO #9: Nivel (Fila 19, Celda 12)
        self._rellenar_celda(tabla, 18, 11, datos['nivel'])
        
        # CAMPO #10: Horas (Fila 19, Celda 17)
        self._rellenar_celda(tabla, 18, 16, datos['horas'])
        
        # CAMPO #11: Fecha inicio (Fila 20, Celda 3)
        self._rellenar_celda(tabla, 19, 2, datos['fecha_inicio'])
        
        # CAMPO #12: Fecha fin (Fila 20, Celda 11)
        self._rellenar_celda(tabla, 19, 10, datos['fecha_fin'])
        
        # CAMPO #13: Código módulo en tabla (Fila 23, Celda 1)
        self._rellenar_celda(tabla, 22, 0, datos['codigo_modulo'])
        
        # CAMPO #13.5: Denominación en tabla (Fila 23, Celdas 7-19)
        # La denominación va en las celdas centrales de la fila 23
        fila_23 = tabla.rows[22]
        for idx in range(6, 19):  # Celdas 7 a 19 (índices 6 a 18)
            try:
                celda = fila_23.cells[idx]
                # Limpiar contenido actual
                for paragraph in celda.paragraphs:
                    paragraph.clear()
                
                # Si no hay párrafos, crear uno
                if not celda.paragraphs:
                    celda.add_paragraph()
                
                # Agregar el nuevo texto con fuente más pequeña
                paragraph = celda.paragraphs[0]
                run = paragraph.add_run(datos['nombre_modulo'])
                run.font.size = Pt(9)  # Fuente 9pt para que entre completo
                break  # Solo rellenar la primera celda del rango
            except:
                pass
        
        # CAMPO #14: Horas en tabla (Fila 23, Celda 20)
        self._rellenar_celda(tabla, 22, 19, datos['horas'])
        
        # CAMPO #15: Calificación (Fila 23, Celda 22)
        self._rellenar_celda(tabla, 22, 21, datos['calificacion'])
        
        # CAMPO #16: Ciudad (Fila 25, después de "En")
        # La fila 25 tiene "En" seguido de espacio para la ciudad
        fila_ciudad = tabla.rows[24]
        for celda in fila_ciudad.cells:
            if 'En' in celda.text and len(celda.text.strip()) <= 3:
                # Reemplazar "En" por "En CIUDAD"
                for paragraph in celda.paragraphs:
                    for run in paragraph.runs:
                        if 'En' in run.text:
                            run.text = f"En {datos['ciudad']}"
                            break
                break
        
        # CAMPO #17: Firma (Fila 29, después de "Fdo:")
        # Agregar nombre del director después de "Fdo:"
        fila_firma = tabla.rows[28]  # Fila 29 (índice 28)
        for celda in fila_firma.cells:
            if 'Fdo.:' in celda.text or 'Fdo.' in celda.text:
                # Agregar nombre después de Fdo.:
                for paragraph in celda.paragraphs:
                    for run in paragraph.runs:
                        if 'Fdo.' in run.text:
                            run.text = f"Fdo.: {datos['director']}"
                            break
                break
        
        # Guardar en bytes
        output = io.BytesIO()
        doc.save(output)
        output.seek(0)
        
        return output.read()
    
    def _rellenar_celda(self, tabla, fila_idx: int, celda_idx: int, valor: str):
        """
        Rellena una celda específica de la tabla
        
        Args:
            tabla: Tabla del documento
            fila_idx: Índice de fila (0-based)
            celda_idx: Índice de celda (0-based)
            valor: Valor a insertar
        """
        try:
            celda = tabla.rows[fila_idx].cells[celda_idx]
            
            # Limpiar contenido actual
            for paragraph in celda.paragraphs:
                for run in paragraph.runs:
                    run.text = ""
            
            # Si no hay párrafos, crear uno
            if not celda.paragraphs:
                celda.add_paragraph()
            
            # Agregar el nuevo texto al primer párrafo
            celda.paragraphs[0].text = str(valor)
            
        except Exception as e:
            print(f"⚠ Error rellenando celda [{fila_idx}, {celda_idx}]: {e}")
    
    def generar_nombre_archivo(self, datos: Dict) -> str:
        """
        Genera nombre de archivo para el certificado
        Formato: NOMBRE_APELLIDO1_APELLIDO2_DNI
        
        Args:
            datos: Datos del alumno
            
        Returns:
            Nombre de archivo (sin extensión)
        """
        # Extraer nombre y apellidos
        nombre_completo = datos['nombre_alumno']
        partes = nombre_completo.split()
        
        if len(partes) >= 3:
            # Formato: NOMBRE APELLIDO1 APELLIDO2 (o más)
            nombre = partes[0]
            apellido1 = partes[1]
            apellido2 = partes[2]
            nombre_formateado = f"{nombre}_{apellido1}_{apellido2}"
        elif len(partes) == 2:
            # Solo NOMBRE APELLIDO1
            nombre = partes[0]
            apellido1 = partes[1]
            nombre_formateado = f"{nombre}_{apellido1}"
        else:
            # Un solo nombre
            nombre_formateado = nombre_completo.replace(' ', '_')
        
        dni = datos['dni_alumno']
        
        return f"{nombre_formateado}_{dni}".upper()


def generar_certificado_ocupado(plantilla_path: str, datos: Dict) -> tuple:
    """
    Función principal para generar un certificado
    
    Args:
        plantilla_path: Ruta a la plantilla
        datos: Datos del alumno
        
    Returns:
        Tuple (bytes del documento, nombre de archivo)
    """
    generator = CertificacionOcupadosGenerator(plantilla_path)
    certificado_bytes = generator.generar_certificado(datos)
    nombre_archivo = generator.generar_nombre_archivo(datos)
    
    return certificado_bytes, nombre_archivo