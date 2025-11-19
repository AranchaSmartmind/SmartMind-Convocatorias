"""
Configuración general de la aplicación
"""

# Configuración de página
PAGE_CONFIG = {
    "page_title": "Interpros SmartMind",
    "page_icon": "logo.png",
    "layout": "wide",
    "initial_sidebar_state": "expanded"
}

# Secciones de la aplicación
SECCIONES = {
    "Captación": "Carga los documentos de captación de alumnos.",
    "Inicio": "Documentación de inicio de formación en empresa.",
    "Fin": "Documentación de finalización de formación en empresa.",
    "Evaluación": "Documentos de evaluación del curso.",
    "Cierre Mes": "Documentación de cierre mensual.",
}

# Configuración de Tesseract OCR
TESSERACT_PATHS = [
    r'C:\Program Files\Tesseract-OCR\tesseract.exe',
    r'C:\Program Files (x86)\Tesseract-OCR\tesseract.exe',
    r'C:\Tesseract-OCR\tesseract.exe'
]

# Tipos de archivo permitidos
ALLOWED_FILE_TYPES = {
    "images": ["png", "jpg", "jpeg", "bmp", "tiff", "gif"],
    "documents": ["pdf", "docx", "doc"],
    "spreadsheets": ["xlsx", "xls", "csv"]
}