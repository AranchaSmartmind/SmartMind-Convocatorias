"""
Aplicación Principal - SmartMind
Sistema de gestión de documentación para convocatorias
"""
import streamlit as st
import platform
import os
import sys
import pytesseract

# Añadir el directorio actual al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importar configuración
try:
    from config.settings import PAGE_CONFIG, TESSERACT_PATHS, SECCIONES
except ImportError as e:
    st.error(f"Error al importar configuración: {e}")
    st.stop()

# Importar estilos
try:
    from styles.custom_styles import get_custom_styles
except ImportError as e:
    st.error(f"Error al importar estilos: {e}")
    st.stop()

# Importar componentes
try:
    from components.header import render_header
    from components.navigation import render_navigation
except ImportError as e:
    st.error(f"Error al importar componentes: {e}")
    st.stop()

# Importar secciones
try:
    from sections.evaluacion import render_evaluacion
    from sections.formacion_fin import render_formacion_fin
    from sections.formacion_inicio import render_formacion_inicio
    from sections.captacion import render_captacion
    from sections.cierre_mes import render_cierre_mes
except ImportError as e:
    st.error(f"Error al importar secciones: {e}")
    st.stop()


# Configurar Tesseract OCR
if platform.system() == 'Windows':
    for ruta in TESSERACT_PATHS:
        if os.path.exists(ruta):
            pytesseract.pytesseract.tesseract_cmd = ruta
            break


# Configurar página
st.set_page_config(**PAGE_CONFIG)

# Aplicar estilos
st.markdown(get_custom_styles(), unsafe_allow_html=True)

# Renderizar header
render_header()

# Renderizar navegación
seccion_actual = render_navigation()

# Contenedor principal
st.markdown('<div style="padding: 2rem 3rem;">', unsafe_allow_html=True)

# Título de la sección
st.markdown(f"""
<div style="margin-bottom: 2rem;">
    <h1 style="font-size: 2rem; color: white; font-weight: 700; margin: 0;">
        {seccion_actual}
    </h1>
    <p style="color: #e2e8f0; font-size: 1rem; margin-top: 0.5rem;">
        {SECCIONES[seccion_actual]}
    </p>
</div>
""", unsafe_allow_html=True)

# Renderizar sección actual
if seccion_actual == "Captación":
    render_captacion()

elif seccion_actual == "Formación Empresa Inicio":
    render_formacion_inicio()

elif seccion_actual == "Formación Empresa Fin":
    render_formacion_fin()

elif seccion_actual == "Evaluación":
    render_evaluacion()

elif seccion_actual == "Cierre Mes":
    render_cierre_mes()

st.markdown("</div>", unsafe_allow_html=True)