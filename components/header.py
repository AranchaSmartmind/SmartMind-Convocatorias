"""
Componente del header de la aplicación
"""
import streamlit as st
import base64


def render_header():
    """Renderiza el header con el logo"""
    try:
        with open("assets/logo.png", "rb") as f:
            logo_b64 = base64.b64encode(f.read()).decode()
            
        st.markdown(f'''
        <div style="background: transparent; 
                    padding: 1.5rem 3rem; 
                    border-bottom: 1px solid rgba(74, 144, 226, 0.2);
                    position: sticky;
                    top: 0;
                    z-index: 1000;">
            <div style="display: flex; align-items: center; gap: 1.5rem;">
                <div style="display: flex; align-items: center; gap: 1rem;">
                    <div style="
                        width: 70px; 
                        height: 70px; 
                        background: white; 
                        border-radius: 50%; 
                        display: flex; 
                        align-items: center; 
                        justify-content: center;
                        padding: 8px;
                        box-shadow: 0 4px 15px rgba(74, 144, 226, 0.4);
                        border: 2px solid rgba(74, 144, 226, 0.3);
                        flex-shrink: 0;
                    ">
                        <img src="data:image/png;base64,{logo_b64}" 
                             style="width: 100%; 
                                    height: 100%; 
                                    object-fit: contain; 
                                    border-radius: 50%;">
                    </div>
                    <div>
                        <h1 style="margin: 0; font-size: 1.6rem; color: #e8f1f8; font-weight: 800; letter-spacing: -0.02em;">
                            SmartMind
                        </h1>
                        <p style="margin: 0; font-size: 0.85rem; color: #4a90e2; font-weight: 500;">
                            Documentación Convocatorias
                        </p>
                    </div>
                </div>
            </div>
        </div>
        ''', unsafe_allow_html=True)
    except:
        st.warning("No se pudo cargar el logo")