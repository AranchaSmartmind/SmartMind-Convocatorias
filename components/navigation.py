"""
Componente de navegación entre secciones
"""
import streamlit as st
from config import SECCIONES


def render_navigation():
    """Renderiza la barra de navegación"""
    if "seccion_actual" not in st.session_state:
        st.session_state.seccion_actual = list(SECCIONES.keys())[0]

    st.markdown('<div style="padding: 0 3rem; margin-top: 1.5rem;">', unsafe_allow_html=True)
    cols = st.columns(len(SECCIONES))

    for idx, (nombre, descripcion) in enumerate(SECCIONES.items()):
        with cols[idx]:
            if st.button(
                nombre,
                key=f"nav_{nombre}",
                use_container_width=True,
                type="primary" if st.session_state.seccion_actual == nombre else "secondary"
            ):
                st.session_state.seccion_actual = nombre
                st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)
    
    return st.session_state.seccion_actual