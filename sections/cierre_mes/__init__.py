"""
SECCIÓN CIERRE MES - LAYOUT 2-2 CORREGIDO
Dos columnas arriba, dos columnas abajo
"""
import streamlit as st

def render_cierre_mes():
    """Render para Cierre Mes con layout 2-2"""
    
    st.markdown("## Cierre Mes")
    st.markdown("Generación automática de partes mensuales de asistencia procesando documentos con OCR")
    
    st.markdown("Sistema automatizado para generar partes mensuales de asistencia procesando documentos con OCR (Reconocimiento Óptico de Caracteres).")
    
    st.markdown("### Documentos Requeridos")
    st.markdown("Sube todos los archivos obligatorios en el panel izquierdo para comenzar.")
    
    # Expandable con información
    with st.expander("Información sobre los archivos"):
        st.markdown("**Descripción de archivos necesarios**")
    
    # ========================================
    # LAYOUT 2-2: PRIMERA FILA (2 COLUMNAS)
    # ========================================
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("### 1. Información del Curso")
        st.markdown("**PDF de Otorgamiento de Becas**")
        pdf_becas = st.file_uploader(
            "Drag and drop file here",
            key="cierre_mes_pdf_becas",
            type=['pdf'],
            help="PDF con información del curso"
        )
        if pdf_becas:
            st.success("Cargado")
        else:
            st.warning("Requerido")
        
        st.markdown("**Excel de Control de Alumnos**")
        excel_control = st.file_uploader(
            "Drag and drop file here",
            key="cierre_mes_excel_control",
            type=['xlsx', 'xls'],
            help="Excel con el listado de alumnos"
        )
        if excel_control:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col2:
        st.markdown("### 2. Hojas de Asistencia")
        st.markdown("**PDFs de Hojas de Firmas**")
        pdfs_firmas = st.file_uploader(
            "Drag and drop files here",
            key="cierre_mes_pdfs_firmas",
            type=['pdf'],
            accept_multiple_files=True,
            help="Múltiples PDFs con hojas de firmas"
        )
        if pdfs_firmas and len(pdfs_firmas) > 0:
            st.success(f"{len(pdfs_firmas)} archivo(s) cargado(s)")
        else:
            st.warning("Requerido")
        
        st.markdown("### 3. Justificantes (Opcional)")
        st.markdown("**PDF de Justificantes**")
        pdf_justificantes = st.file_uploader(
            "Drag and drop file here",
            key="cierre_mes_pdf_justificantes",
            type=['pdf'],
            help="PDF con justificantes de ausencias"
        )
        if pdf_justificantes:
            st.success("Cargado")
        else:
            st.info("Opcional")
    
    # Validación de archivos
    if not pdf_becas or not excel_control or not pdfs_firmas:
        st.info("Sube los archivos requeridos para continuar")
        return
    
    # Botón de generación
    st.markdown("---")
    if st.button("Generar Parte Mensual", type="primary", use_container_width=True):
        with st.spinner("Procesando documentos..."):
            try:
                # Aquí va tu lógica de procesamiento
                st.success("Parte mensual generado correctamente")
            except Exception as e:
                st.error(f"Error: {str(e)}")


# Versión alternativa si prefieres más separación visual:
def render_cierre_mes_alternativo():
    """Versión con más separación entre secciones"""
    
    st.markdown("## Cierre Mes")
    
    st.markdown("### Documentos Requeridos")
    
    # ========================================
    # PRIMERA FILA: 2 COLUMNAS
    # ========================================
    st.markdown("#### Información del Curso")
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**PDF de Otorgamiento de Becas**")
        pdf_becas = st.file_uploader(
            "Drag and drop file here",
            key="alt_pdf_becas",
            type=['pdf']
        )
    
    with col2:
        st.markdown("**Excel de Control de Alumnos**")
        excel_control = st.file_uploader(
            "Drag and drop file here",
            key="alt_excel_control",
            type=['xlsx', 'xls']
        )
    
    # ========================================
    # SEGUNDA FILA: 2 COLUMNAS
    # ========================================
    st.markdown("#### Documentación de Asistencia")
    col3, col4 = st.columns(2)
    
    with col3:
        st.markdown("**PDFs de Hojas de Firmas**")
        pdfs_firmas = st.file_uploader(
            "Drag and drop files here",
            key="alt_pdfs_firmas",
            type=['pdf'],
            accept_multiple_files=True
        )
    
    with col4:
        st.markdown("**PDF de Justificantes (Opcional)**")
        pdf_justificantes = st.file_uploader(
            "Drag and drop file here",
            key="alt_pdf_justificantes",
            type=['pdf']
        )


if __name__ == "__main__":
    render_cierre_mes()