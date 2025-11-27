"""
Configuración general de la aplicación
"""

PAGE_CONFIG = {
    "page_title": "Interpros SmartMind",
    "page_icon": "logo.png",
    "layout": "wide",
    "initial_sidebar_state": "expanded"
}

SECCIONES = {
    "Captación": "Carga los documentos de captación de alumnos.",
    "Inicio": "Documentación de inicio de formación en empresa.",
    "Fin": "Documentación de finalización de formación en empresa.",
    "Evaluación": "Documentos de evaluación del curso.",
    "Cierre Mes": "Documentación de cierre mensual.",
}

SECCIONES = {
    "Captación": "Gestión de candidatos y procesos de selección",
    "Inicio": "Inicio de curso y documentación inicial",
    "Fin": "Finalización de curso y cierre de expedientes",
    "Evaluación": "Evaluación de alumnos y seguimiento académico",
    "Cierre Mes": "Generación automática de partes mensuales de asistencia"  # AÑADIR ESTA LÍNEA
}

TESSERACT_PATHS = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    r'C:\Tesseract-OCR\tesseract.exe'
]

ALLOWED_FILE_TYPES = {
    "images": ["png", "jpg", "jpeg", "bmp", "tiff", "gif"],
    "documents": ["pdf", "docx", "doc"],
    "spreadsheets": ["xlsx", "xls", "csv"]
}