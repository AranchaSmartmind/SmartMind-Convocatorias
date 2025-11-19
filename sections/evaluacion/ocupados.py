"""
Módulo de Evaluación - Ocupados
Gestión de actas para formación de trabajadores ocupados
"""
import streamlit as st


def render_tab_ocupados():
    """Renderiza el tab de Ocupados"""
    st.markdown("#### Actas para Formación de Ocupados")
    st.markdown(
        '<div class="custom-card"><p>Gestión de actas individuales, grupales y certificaciones para formación de trabajadores ocupados.</p></div>',
        unsafe_allow_html=True
    )
    
    st.info(" Esta sección está en desarrollo. Próximamente estará disponible.")
    
    # TODO: Implementar la lógica para actas de ocupados
    # - Actas individuales
    # - Actas grupales
    # - Certificaciones