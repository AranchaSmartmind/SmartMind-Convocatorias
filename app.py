"""
Aplicación Principal - SmartMind
Sistema de gestión de documentación para convocatorias
"""
import streamlit as st
import platform
import os
import sys
import pytesseract

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from config.settings import PAGE_CONFIG, TESSERACT_PATHS, SECCIONES

st.set_page_config(**PAGE_CONFIG)

try:
    from styles.custom_styles import get_custom_styles, get_robot_assistant, get_interpros_logo
    from styles.custom_styles import get_custom_styles, get_robot_assistant, get_interpros_logo
except ImportError as e:
    st.error(f"Error al importar estilos: {e}")
    st.stop()

try:
    from sections.evaluacion import render_evaluacion
    from sections.fin import render_fin
    from sections.inicio import render_inicio
    from sections.captacion import render_captacion
    from sections.cierre_mes import render_cierre_mes       

except ImportError as e:
    st.error(f"Error al importar secciones: {e}")
    st.stop()

if platform.system() == 'Windows':
    for ruta in TESSERACT_PATHS:
        if os.path.exists(ruta):
            pytesseract.pytesseract.tesseract_cmd = ruta
            break

st.markdown(get_custom_styles(), unsafe_allow_html=True)

st.markdown(get_robot_assistant('assets/robot_asistente.png'), unsafe_allow_html=True)

with st.sidebar:
    st.markdown(get_interpros_logo(), unsafe_allow_html=True)
    st.markdown("### SmartMind")
    st.markdown("---")

    if st.button("Captación", key="nav_captacion", use_container_width=True):
        st.session_state.seccion_actual = "Captación"
        st.rerun()

    if st.button("Inicio", key="nav_inicio", use_container_width=True):
        st.session_state.seccion_actual = "Inicio"
        st.rerun()

    if st.button("Fin", key="nav_fin", use_container_width=True):
        st.session_state.seccion_actual = "Fin"
        st.rerun()
    
    if st.button("Evaluación", key="nav_evaluacion", use_container_width=True):
        st.session_state.seccion_actual = "Evaluación"
        st.rerun()
    
    if st.button("Cierre Mes", key="nav_cierre", use_container_width=True):
        st.session_state.seccion_actual = "Cierre Mes"
        st.rerun()

if 'seccion_actual' not in st.session_state:
    st.session_state.seccion_actual = 'Inicio'

seccion_actual = st.session_state.seccion_actual

st.markdown(get_robot_assistant('assets/robot_asistente.png'), unsafe_allow_html=True)

with st.sidebar:
    st.markdown(get_interpros_logo(), unsafe_allow_html=True)
    st.markdown("### SmartMind")
    st.markdown("---")

    if st.button("Captación", key="nav_captacion", use_container_width=True):
        st.session_state.seccion_actual = "Captación"
        st.rerun()

    if st.button("Inicio", key="nav_inicio", use_container_width=True):
        st.session_state.seccion_actual = "Inicio"
        st.rerun()

    if st.button("Fin", key="nav_fin", use_container_width=True):
        st.session_state.seccion_actual = "Fin"
        st.rerun()
    
    if st.button("Evaluación", key="nav_evaluacion", use_container_width=True):
        st.session_state.seccion_actual = "Evaluación"
        st.rerun()
    
    if st.button("Cierre Mes", key="nav_cierre", use_container_width=True):
        st.session_state.seccion_actual = "Cierre Mes"
        st.rerun()

if 'seccion_actual' not in st.session_state:
    st.session_state.seccion_actual = 'Inicio'

seccion_actual = st.session_state.seccion_actual

st.markdown(f"""
<div style="margin-bottom: 2rem;">
    <h1 style="font-size: 2rem; color: white; font-weight: 700; margin: 0;">
        {seccion_actual}
    </h1>
    <p style="color: #e2e8f0; font-size: 1rem; margin-top: 0.5rem;">
        {SECCIONES.get(seccion_actual, '')}
        {SECCIONES.get(seccion_actual, '')}
    </p>
</div>
""", unsafe_allow_html=True)

if seccion_actual == "Captación":
    render_captacion()
elif seccion_actual == "Inicio":
    render_inicio()
elif seccion_actual == "Fin":
    render_fin()
elif seccion_actual == "Evaluación":
    render_evaluacion()
elif seccion_actual == "Cierre Mes":
    render_cierre_mes()