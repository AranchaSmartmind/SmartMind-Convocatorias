"""
FUNCIÓN RENDER_TRANSVERSALES COMPLETA
======================================
Esta función reemplaza el render_transversales actual en desempleados.py
"""

import streamlit as st # type: ignore
import pandas as pd # type: ignore

def render_transversales():
    """Render para actas transversales (FCOO03)"""
    
    st.markdown("### 📚 Actas Transversales (FCOO03)")
    st.markdown("Genera actas de evaluación final para competencias transversales")
    
    # Subida de archivos
    st.markdown("### Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="desempleados_transversal_cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("✅ Cargado")
        else:
            st.warning("⚠️ Requerido")
    
    with col2:
        st.markdown("**Control de Tareas**")
        control_file = st.file_uploader(
            "Excel CTRL_Tareas_AREA*",
            key="desempleados_transversal_control",
            type=['xlsx', 'xls']
        )
        if control_file:
            st.success("✅ Cargado")
        else:
            st.warning("⚠️ Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="desempleados_transversal_plantilla",
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial SEPE predeterminada"
        )
        if plantilla_file:
            st.success("✅ Personalizada")
        else:
            st.info("ℹ️ Por defecto")
    
    with st.expander("ℹ️ Información", expanded=False):
        st.markdown("""
        **Archivos necesarios:**
        
        1. **Cronograma** - Excel con fechas y datos del curso (Requerido)
        2. **Control de Tareas** - Excel CTRL_Tareas_AREA con asistencias y calificaciones (Requerido)
        3. **Plantilla** - Word con formato oficial (opcional)
        
        **Plantilla predeterminada:**
        
        Si NO subes una plantilla, se usará la plantilla oficial SEPE de actas transversales integrada.
        
        **El acta incluye:**
        - Datos del curso (Acción, Especialidad, Código)
        - Centro: INTERPROS NEXT GENERATION S.L.U
        - Fechas de inicio y finalización
        - Tabla con todos los alumnos (hasta 20)
        - DNI, Nombre, Horas de actividades
        - Calificación final (APTO/NO APTO)
        - Espacios para firmas
        """)
    
    if not cronograma_file or not control_file:
        st.info("📤 Sube el cronograma y control de tareas para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('⚙️ Procesando archivos...'):
            # Importar procesador
            from transversales_processor import TransversalesProcessor
            
            # Procesar datos
            processor = TransversalesProcessor()
            cronograma_file.seek(0)
            control_file.seek(0)
            datos = processor.cargar_datos(control_file.read(), cronograma_file.read())
        
        st.success("✅ Datos procesados correctamente")
        
        # Mostrar resumen
        st.markdown("### Resumen del Curso")
        
        with st.expander("Ver datos extraídos", expanded=True):
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Alumnos", datos['total_alumnos'])
            
            with col2:
                aptos = sum(1 for a in datos['alumnos'] if 'apto' in a['calificacion_final'].lower())
                st.metric("APTOS", aptos)
            
            with col3:
                no_aptos = sum(1 for a in datos['alumnos'] if 'no apto' in a['calificacion_final'].lower())
                st.metric("NO APTOS", no_aptos)
            
            with col4:
                st.metric("Horas Totales", datos['horas_totales'])
            
            st.markdown("---")
            
            # Información del curso
            st.markdown("**Información del Curso**")
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"Acción: {datos['campo_2_accion']}")
                st.text(f"Código: {datos['campo_4_codigo']}")
                st.text(f"Especialidad: {datos['campo_3_especialidad'][:50]}...")
            with col2:
                st.text(f"Fecha inicio: {datos['campo_9_fecha_inicio']}")
                st.text(f"Fecha fin: {datos['campo_10_fecha_fin']}")
                st.text(f"Modalidad: {datos['campo_8_modalidad']}")
            
            st.markdown("---")
            
            # Tabla de alumnos
            st.markdown("**Listado de Alumnos**")
            df = pd.DataFrame([{
                'Nº': a['numero'],
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Horas': a['horas_actividades'],
                'Calificación': a['calificacion_final']
            } for a in datos['alumnos']])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Generar acta
        st.markdown("---")
        st.markdown("### Generar Acta Transversal")
        
        if st.button("🚀 Generar Acta Transversal (FCOO03)", 
                    type="primary", 
                    use_container_width=True,
                    key="desempleados_transversal_generar"):
            try:
                # Obtener plantilla
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("📝 Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_transversal_por_defecto()
                    if plantilla_bytes:
                        st.info("📝 Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("❌ No se encontró la plantilla predeterminada")
                        st.warning("⚠️ Sube una plantilla manualmente")
                        return
                
                # Generar acta
                with st.spinner('⚙️ Generando acta transversal...'):
                    from word_generator_transversal import WordGeneratorTransversal
                    
                    gen = WordGeneratorTransversal(plantilla_bytes)
                    doc = gen.generar_acta(datos)
                    
                    # Guardar en session state
                    st.session_state['acta_transversal_desempleados'] = doc
                    st.session_state['nombre_acta_transversal_desempleados'] = f"Acta_Transversal_FCOO03_{datos['campo_2_accion'].replace('/', '_')}.docx"
                
                st.balloons()
                st.success("✅ ¡Acta transversal generada correctamente!")
                
            except Exception as e:
                st.error(f"❌ Error generando acta: {str(e)}")
                st.exception(e)
        
        # Descarga
        if 'acta_transversal_desempleados' in st.session_state:
            st.markdown("---")
            st.markdown("### Descargar")
            
            st.download_button(
                label="💾 Descargar Acta Transversal",
                data=st.session_state['acta_transversal_desempleados'],
                file_name=st.session_state['nombre_acta_transversal_desempleados'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key="desempleados_transversal_download"
            )
    
    except Exception as e:
        st.error(f"❌ Error procesando archivos: {str(e)}")
        st.exception(e)