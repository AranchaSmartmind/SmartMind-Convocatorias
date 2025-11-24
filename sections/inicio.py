"""
Sección de Inicio
"""
import streamlit as st


def render_inicio():
    """Renderiza la sección de Inicio"""
    st.markdown("### Documentación de Inicio de Formación en Empresa")
    st.markdown('<div class="custom-card"><p>Sección en desarrollo para la gestión de documentos de inicio de formación.</p></div>', unsafe_allow_html=True)
    
    st.info("Esta funcionalidad estará disponible próximamente.")

    st.markdown("""
    #### Documentos que se procesarán:
    - Acuerdos de formación en empresa
    - Planes de formación individual
    - Designación de tutores
    - Calendarios de formación
    - Seguros de responsabilidad civil
    """)