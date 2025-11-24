"""
Sección de Evaluación
"""
import streamlit as st
import io
import zipfile
import openpyxl
import pandas as pd
from datetime import datetime
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from utils import (
    extraer_datos_multiples_documentos,
    rellenar_acta_desde_plantilla,
    visualizar_documento_word
)


def render_evaluacion():
    """Renderiza la sección de Evaluación"""
    st.markdown("### Sistema de Gestión de Actas de Evaluación")
    st.markdown('<div class="custom-card"><p>Selecciona el tipo de formación para generar las actas correspondientes.</p></div>', unsafe_allow_html=True)
    
    tab_ocupados, tab_desempleados, tab_gforma = st.tabs(["Ocupados", "Desempleados", "GFORMA"])
    
    with tab_ocupados:
        render_tab_ocupados()
    
    with tab_desempleados:
        render_tab_desempleados()
    
    with tab_gforma:
        render_tab_gforma()


def render_tab_ocupados():
    """Renderiza el tab de Ocupados"""
    st.markdown("#### Actas para Formación de Ocupados")
    st.markdown('<div class="custom-card"><p>Gestión de actas individuales, grupales y certificaciones para formación de trabajadores ocupados.</p></div>', unsafe_allow_html=True)

def render_tab_desempleados():
    """Renderiza el tab de Desempleados"""
    st.markdown("#### Actas para Formación de Desempleados")
    st.markdown('<div class="custom-card"><p>Gestión de actas individuales, grupales y transversales para formación de desempleados.</p></div>', unsafe_allow_html=True)

def render_tab_gforma():
    """Renderiza el tab de GFORMA"""
    st.markdown("#### Generación de Excel para Importación en GFORMA")
    st.markdown('<div class="custom-card"><p>Genera un archivo Excel con formato compatible para importar directamente en la plataforma GFORMA.</p></div>', unsafe_allow_html=True)
