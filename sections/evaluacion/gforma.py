"""
Módulo de Evaluación - GFORMA
Generación de archivos Excel para importación en GFORMA
"""
import streamlit as st


def render_tab_gforma():
    """Renderiza el tab de GFORMA"""
    st.markdown("#### Generación de Excel para Importación en GFORMA")
    st.markdown(
        '<div class="custom-card"><p>Genera un archivo Excel con formato compatible para importar directamente en la plataforma GFORMA.</p></div>',
        unsafe_allow_html=True
    )
    
    st.info(" Esta sección está en desarrollo. Próximamente estará disponible.")
    
    # TODO: Implementar la lógica para generación de Excel GFORMA
    # - Formato específico de GFORMA
    # - Validación de datos
    # - Exportación compatible