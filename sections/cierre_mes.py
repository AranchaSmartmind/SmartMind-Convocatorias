"""
Sección de Cierre Mes
"""
import streamlit as st


def render_cierre_mes():
    """Renderiza la sección de Cierre Mes"""
    st.markdown("### Documentación de Cierre Mensual")
    st.markdown('<div class="custom-card"><p>Sección en desarrollo para la gestión de documentación de cierre mensual.</p></div>', unsafe_allow_html=True)
    
    st.info("Esta funcionalidad estará disponible próximamente.")

    st.markdown("""
    #### Documentos que se procesarán:
    - Partes de asistencia mensuales
    - Justificantes de ausencias
    - Informes de seguimiento
    - Facturas y justificantes de gastos
    - Resúmenes de actividad
    """)