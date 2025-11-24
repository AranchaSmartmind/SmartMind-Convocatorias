"""
Sección de Evaluación - Módulo Principal
Coordina las diferentes subsecciones de evaluación
"""
import streamlit as st
from .desempleados import render_tab_desempleados
from .ocupados import render_tab_ocupados
from .gforma import render_tab_gforma


def render_evaluacion():
    """Renderiza la sección de Evaluación"""
    st.markdown("### Sistema de Gestión de Actas de Evaluación")
    st.markdown(
        '<div class="custom-card"><p>Selecciona el tipo de formación para generar las actas correspondientes.</p></div>',
        unsafe_allow_html=True
    )
    
    tab_ocupados, tab_desempleados, tab_gforma = st.tabs(
        ["Ocupados", "Desempleados", "GFORMA"]
    )
    
    with tab_ocupados:
        render_tab_ocupados()
    
    with tab_desempleados:
        render_tab_desempleados()
    
    with tab_gforma:
        render_tab_gforma()