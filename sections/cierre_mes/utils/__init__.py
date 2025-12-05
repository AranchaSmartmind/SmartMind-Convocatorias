"""
Módulo de utilidades para Cierre de Mes
Generador de Parte Mensual
"""

from .extractor_ocr import calcular_dias_lectivos_y_asistencias
from .procesador_datos import (
    extraer_datos_curso_pdf,
    extraer_alumnos_excel,
    extraer_becas_ayudas_simple,
    extraer_justificantes,
    obtener_mes_anterior
)
from .generador_documento import rellenar_template, construir_observaciones
from .helpers import buscar_coincidencia, normalizar_nombre

__all__ = [
    'calcular_dias_lectivos_y_asistencias',
    'extraer_datos_curso_pdf',
    'extraer_alumnos_excel',
    'extraer_becas_ayudas_simple',
    'extraer_justificantes',
    'obtener_mes_anterior',
    'rellenar_template',
    'construir_observaciones',
    'buscar_coincidencia',
    'normalizar_nombre'
]