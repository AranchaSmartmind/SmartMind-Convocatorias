"""
Sección de Captación
"""
import streamlit as st


def render_captacion():
    """Renderiza la sección de Captación"""
    st.markdown("### Gestión de Documentación de Captación")
    st.markdown('<div class="custom-card"><p>Sección en desarrollo para la gestión de documentos de captación de alumnos.</p></div>', unsafe_allow_html=True)
    
    st.info("Esta funcionalidad estará disponible próximamente.")
    
    st.markdown("""
    #### Documentos que se procesarán:
    - Solicitudes de inscripción
    - Fichas de preinscripción
    - Documentación de identidad
    - Certificados académicos
    - Justificantes de situación laboral
    """)