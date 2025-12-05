"""
Sección de Cierre Mensual - CORREGIDO
"""
import streamlit as st
import os
import tempfile

def cargar_template_cierre_mes_por_defecto():
    """Carga la plantilla de cierre mensual"""
    try:
        plantilla_cierre = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 
            'evaluacion',
            'cierre_mes',
            'template_original.docx'
        )
        
        ubicaciones = [
            plantilla_cierre,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'cierre_mes', 'template_original.docx'),
        ]
        
        for ubicacion in ubicaciones:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'rb') as f:
                    contenido = f.read()
                    if len(contenido) > 1000:
                        return contenido
        
        return None
        
    except Exception as e:
        print(f"Error cargando plantilla: {e}")
        return None

def render_cierre_mes():
    """Renderiza la interfaz de Cierre Mensual"""
    from sections.evaluacion.cierre_mes.procesamiento_datos import (
        extraer_becas_ayudas_tabla,
        extraer_justificantes_mejorado,
        calcular_dias_lectivos_y_faltas_corregido,
        construir_observaciones_completas,
        extraer_alumnos_excel
    )
    from sections.evaluacion.cierre_mes.generacion_word import generar_parte_mensual
    from sections.evaluacion.cierre_mes.utilidades import buscar_coincidencia
    
    tab1, tab2, tab3 = st.tabs(["Subir Archivos", "Procesar", "Descargar"])
    
    with tab1:
        st.subheader("Archivos necesarios")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("**Listado de Alumnos (Excel)**")
            excel_alumnos = st.file_uploader(
                "Archivo Excel con datos de alumnos",
                type=['xlsx'],
                key="cm_excel"
            )
            if excel_alumnos:
                st.success("Cargado")
            else:
                st.warning("Requerido")
        
        with col2:
            st.markdown("**Hojas de Firmas (PDFs)**")
            pdfs_firmas = st.file_uploader(
                "Uno o más PDFs de firmas",
                type=['pdf'],
                accept_multiple_files=True,
                key="cm_firmas"
            )
            if pdfs_firmas:
                st.success(f"{len(pdfs_firmas)} archivos")
            else:
                st.warning("Requerido")
        
        col3, col4 = st.columns(2)
        
        with col3:
            st.markdown("**Becas y Ayudas (PDF)**")
            pdf_becas = st.file_uploader(
                "PDF de Otorgamiento",
                type=['pdf'],
                key="cm_becas"
            )
            if pdf_becas:
                st.success("Cargado")
            else:
                st.warning("Requerido")
        
        with col4:
            st.markdown("**Justificantes (PDF)**")
            pdf_justificantes = st.file_uploader(
                "PDF de justificaciones",
                type=['pdf'],
                key="cm_just"
            )
            if pdf_justificantes:
                st.success("Cargado")
            else:
                st.warning("Requerido")
        
        archivos_completos = all([excel_alumnos, pdf_becas, pdf_justificantes, pdfs_firmas])
        
        if archivos_completos:
            st.success("Todos los archivos listos")
        else:
            st.info("Sube todos los archivos para continuar")
    
    with tab2:
        if not archivos_completos:
            st.warning("Primero sube todos los archivos en la pestaña anterior")
        else:
            st.info("El sistema detectará automáticamente los días lectivos")
            
            if st.button("Iniciar Procesamiento", type="primary", use_container_width=True):
                with st.spinner("Procesando cierre mensual..."):
                    progress = st.progress(0)
                    status = st.empty()
                    
                    try:
                        temp_dir = tempfile.mkdtemp()
                        
                        status.text("Guardando archivos...")
                        progress.progress(5)
                        
                        # Guardar archivos
                        excel_path = os.path.join(temp_dir, excel_alumnos.name)
                        with open(excel_path, 'wb') as f:
                            f.write(excel_alumnos.getbuffer())
                        
                        becas_path = os.path.join(temp_dir, pdf_becas.name)
                        with open(becas_path, 'wb') as f:
                            f.write(pdf_becas.getbuffer())
                        
                        justif_path = os.path.join(temp_dir, pdf_justificantes.name)
                        with open(justif_path, 'wb') as f:
                            f.write(pdf_justificantes.getbuffer())
                        
                        template_path = os.path.join(temp_dir, 'template.docx')
                        template_bytes = cargar_template_cierre_mes_por_defecto()
                        if not template_bytes:
                            st.error("No se pudo cargar la plantilla")
                            return
                        
                        with open(template_path, 'wb') as f:
                            f.write(template_bytes)
                        
                        firmas_paths = []
                        for pdf_firma in pdfs_firmas:
                            firma_path = os.path.join(temp_dir, pdf_firma.name)
                            with open(firma_path, 'wb') as f:
                                f.write(pdf_firma.getbuffer())
                            firmas_paths.append(firma_path)
                        
                        output_path = os.path.join(temp_dir, 'parte_mensual.docx')
                        
                        # PROCESAMIENTO
                        status.text("Extrayendo alumnos del Excel...")
                        progress.progress(15)
                        alumnos_excel = extraer_alumnos_excel(excel_path)
                        
                        status.text("Extrayendo becas y ayudas...")
                        progress.progress(30)
                        # CORREGIDO: Ahora pasa alumnos_excel como segundo parámetro
                        ayudas_dict = extraer_becas_ayudas_tabla(becas_path, alumnos_excel)
                        
                        status.text("Extrayendo justificantes...")
                        progress.progress(45)
                        justificantes_dict = extraer_justificantes_mejorado(justif_path)
                        
                        status.text("Calculando días lectivos y faltas...")
                        progress.progress(60)
                        
                        dias_lectivos, ausencias_dict, dias_con_firma_dict = calcular_dias_lectivos_y_faltas_corregido(
                            firmas_paths,
                            alumnos_excel
                        )
                        
                        status.text("Construyendo observaciones...")
                        progress.progress(75)
                        
                        print("\n" + "="*80)
                        print("ASIGNACIÓN DE DATOS A ALUMNOS")
                        print("="*80)
                        
                        alumnos_finales = []
                        for alumno_excel in alumnos_excel:
                            nombre = alumno_excel['nombre_completo']
                            dni = alumno_excel['dni']
                            
                            ausencias = ausencias_dict.get(nombre, 0)
                            
                            print(f"\n{nombre}")
                            print(f"   DNI: {dni}")
                            print(f"   Faltas: {ausencias}")
                            
                            observaciones = construir_observaciones_completas(
                                nombre,
                                ayudas_dict,
                                dias_con_firma_dict,
                                justificantes_dict
                            )
                            
                            alumnos_finales.append({
                                'nombre': nombre,
                                'dni': dni,
                                'faltas': ausencias,
                                'observaciones': observaciones
                            })
                        
                        status.text("Generando documento Word...")
                        progress.progress(90)
                        
                        datos_documento = {
                            'alumnos': alumnos_finales
                        }
                        
                        exito = generar_parte_mensual(template_path, output_path, datos_documento)
                        
                        progress.progress(100)
                        status.empty()
                        
                        if exito and os.path.exists(output_path):
                            st.success("Procesamiento completado")
                            
                            with open(output_path, 'rb') as f:
                                st.session_state['cm_documento'] = f.read()
                            st.session_state['cm_nombre'] = 'parte_mensual.docx'
                            
                            st.info(f"""
                            Resumen:
                            - Días lectivos detectados: {dias_lectivos}
                            - Alumnos procesados: {len(alumnos_finales)}
                            - Con ayudas: {len(ayudas_dict)}
                            - Con justificantes: {len(justificantes_dict)}
                            """)
                            
                            st.info("Ve a la pestaña Descargar")
                        else:
                            st.error("Error generando el documento")
                            
                    except Exception as e:
                        st.error(f"Error: {str(e)}")
                        with st.expander("Ver detalles"):
                            import traceback
                            st.code(traceback.format_exc())
    
    with tab3:
        if 'cm_documento' in st.session_state:
            st.success("Documento listo para descargar")
            
            st.download_button(
                label="Descargar Parte Mensual",
                data=st.session_state['cm_documento'],
                file_name=st.session_state['cm_nombre'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            
            if st.button("Nuevo Proceso", use_container_width=True):
                del st.session_state['cm_documento']
                del st.session_state['cm_nombre']
                st.rerun()
        else:
            st.info("No hay documento generado todavía")