"""
Sección de Cierre Mensual
"""
import streamlit as st
import os
import tempfile

def render_cierre_mes():
    """
    Renderiza la interfaz de Cierre Mensual
    """
    from evaluacion.cierre_mes.extraccion_ocr import calcular_dias_lectivos_y_asistencias
    from evaluacion.cierre_mes.procesamiento_datos import (
        obtener_mes_anterior,
        extraer_datos_curso_pdf,
        extraer_alumnos_excel,
        extraer_becas_ayudas_simple,
        extraer_justificantes
    )
    from evaluacion.cierre_mes.generacion_word import construir_observaciones, generar_parte_mensual
    from evaluacion.cierre_mes.utilidades import buscar_coincidencia
    
    tab1, tab2, tab3 = st.tabs(["📤 Subir Archivos", "⚙️ Procesar", "💾 Descargar"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Archivos básicos")
            excel_alumnos = st.file_uploader("Listado de Alumnos (Excel)", type=['xlsx'], key="cm_excel")
            pdf_becas = st.file_uploader("Becas y Ayudas (PDF)", type=['pdf'], key="cm_becas")
            pdf_justificantes = st.file_uploader("Justificantes (PDF)", type=['pdf'], key="cm_just")
            template_word = st.file_uploader("Template Word", type=['docx', 'doc'], key="cm_template")
        
        with col2:
            st.subheader("PDFs de Firmas")
            pdfs_firmas = st.file_uploader(
                "Hojas de Firmas (múltiples archivos)", 
                type=['pdf'], 
                accept_multiple_files=True, 
                key="cm_firmas"
            )
            if pdfs_firmas:
                st.success(f"✓ {len(pdfs_firmas)} archivos cargados")
                for pdf in pdfs_firmas:
                    st.text(f"  • {pdf.name}")
        
        archivos_completos = all([excel_alumnos, pdf_becas, pdf_justificantes, template_word, pdfs_firmas])
        
        if archivos_completos:
            st.success("✅ Todos los archivos cargados correctamente")
        else:
            st.info("📋 Por favor, sube todos los archivos necesarios")
    
    with tab2:
        if not archivos_completos:
            st.warning("⚠️ Primero debes subir todos los archivos en la pestaña anterior")
        else:
            if st.button("🚀 Iniciar Procesamiento", type="primary", use_container_width=True):
                with st.spinner("Procesando cierre mensual..."):
                    progress = st.progress(0)
                    status = st.empty()
                    
                    try:
                        temp_dir = tempfile.mkdtemp()
                        
                        # Guardar archivos
                        status.text("Guardando archivos...")
                        progress.progress(10)
                        
                        excel_path = os.path.join(temp_dir, excel_alumnos.name)
                        with open(excel_path, 'wb') as f:
                            f.write(excel_alumnos.getbuffer())
                        
                        becas_path = os.path.join(temp_dir, pdf_becas.name)
                        with open(becas_path, 'wb') as f:
                            f.write(pdf_becas.getbuffer())
                        
                        justif_path = os.path.join(temp_dir, pdf_justificantes.name)
                        with open(justif_path, 'wb') as f:
                            f.write(pdf_justificantes.getbuffer())
                        
                        template_path = os.path.join(temp_dir, template_word.name)
                        with open(template_path, 'wb') as f:
                            f.write(template_word.getbuffer())
                        
                        firmas_paths = []
                        for pdf_firma in pdfs_firmas:
                            firma_path = os.path.join(temp_dir, pdf_firma.name)
                            with open(firma_path, 'wb') as f:
                                f.write(pdf_firma.getbuffer())
                            firmas_paths.append(firma_path)
                        
                        output_path = os.path.join(temp_dir, 'parte_mensual.docx')
                        
                        # Procesar datos
                        status.text("Extrayendo datos del curso...")
                        progress.progress(20)
                        mes = obtener_mes_anterior()
                        numero_curso, especialidad = extraer_datos_curso_pdf(becas_path)
                        
                        status.text("Procesando alumnos...")
                        progress.progress(30)
                        alumnos_excel = extraer_alumnos_excel(excel_path)
                        
                        status.text("Extrayendo becas y ayudas...")
                        progress.progress(40)
                        ayudas_por_alumno = extraer_becas_ayudas_simple(becas_path)
                        
                        status.text("Procesando justificantes...")
                        progress.progress(50)
                        justificantes_por_alumno = extraer_justificantes(justif_path)
                        
                        status.text("Calculando días lectivos y asistencias (puede tardar)...")
                        progress.progress(60)
                        dias_lectivos, asistencias_por_alumno, faltas_por_alumno = calcular_dias_lectivos_y_asistencias(firmas_paths)
                        
                        status.text("Construyendo observaciones...")
                        progress.progress(80)
                        
                        alumnos_finales = []
                        for alumno_excel in alumnos_excel:
                            nombre = alumno_excel['nombre_completo']
                            dni = alumno_excel['dni']
                            
                            ayudas = buscar_coincidencia(nombre, ayudas_por_alumno) or []
                            justificantes = buscar_coincidencia(nombre, justificantes_por_alumno) or 0
                            asistencias = buscar_coincidencia(nombre, asistencias_por_alumno)
                            dias_empresa = asistencias['dias_empresa'] if asistencias else 0
                            dias_aula = asistencias['dias_aula'] if asistencias else 0
                            faltas = buscar_coincidencia(nombre, faltas_por_alumno) or 0
                            
                            observaciones = construir_observaciones(
                                ayudas=ayudas,
                                dias_aula=dias_aula,
                                dias_empresa=dias_empresa,
                                justificantes=justificantes,
                                dias_lectivos=dias_lectivos,
                                faltas=faltas
                            )
                            
                            alumnos_finales.append({
                                'nombre': nombre,
                                'dni': dni,
                                'faltas': faltas,
                                'observaciones': observaciones
                            })
                        
                        status.text("Generando documento Word...")
                        progress.progress(90)
                        
                        datos_documento = {
                            'numero_curso': numero_curso,
                            'especialidad': especialidad,
                            'centro': 'INTERPROS NEXT GENERATION SLU',
                            'mes': mes,
                            'dias_lectivos': dias_lectivos,
                            'horas_empresa': 21,
                            'horas_aula': 2,
                            'alumnos': alumnos_finales
                        }
                        
                        exito = generar_parte_mensual(template_path, output_path, datos_documento)
                        
                        progress.progress(100)
                        status.empty()
                        
                        if exito and os.path.exists(output_path):
                            st.success("✅ Procesamiento completado exitosamente")
                            
                            with open(output_path, 'rb') as f:
                                st.session_state['cm_documento'] = f.read()
                            st.session_state['cm_nombre'] = 'parte_mensual.docx'
                            
                            st.info("👉 Ve a la pestaña 'Descargar' para obtener tu documento")
                        else:
                            st.error("❌ Error durante la generación del documento")
                            
                    except Exception as e:
                        st.error(f"❌ Error: {str(e)}")
                        with st.expander("Ver detalles del error"):
                            import traceback
                            st.code(traceback.format_exc())
    
    with tab3:
        if 'cm_documento' in st.session_state:
            st.success("✅ Documento generado y listo para descargar")
            
            col1, col2 = st.columns([3, 1])
            
            with col1:
                st.download_button(
                    label="📥 Descargar Parte Mensual",
                    data=st.session_state['cm_documento'],
                    file_name=st.session_state['cm_nombre'],
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True
                )
            
            with col2:
                if st.button("🔄 Nuevo Proceso", use_container_width=True):
                    del st.session_state['cm_documento']
                    del st.session_state['cm_nombre']
                    st.rerun()
        else:
            st.info("📄 No hay documento generado todavía. Primero debes procesar los archivos en la pestaña anterior.")