"""
Módulo de Evaluación - Desempleados
Generación automática de actas SEPE
"""
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile

# Importaciones con manejo de errores
try:
    from .excel_processor import ExcelProcessorReal
    from .word_generator import WordGeneratorSEPE
except ImportError:
    try:
        from excel_processor import ExcelProcessorReal
        from word_generator import WordGeneratorSEPE
    except ImportError:
        import sys
        import os
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        from excel_processor import ExcelProcessorReal
        from word_generator import WordGeneratorSEPE


def render_tab_desempleados():
    """Función principal del módulo de desempleados"""
    st.markdown("#### 📋 Generación Automática de Actas SEPE")
    st.info("Sistema automatizado de generación de actas en formato Word")
    
    # Sección de carga de archivos
    st.markdown("### 📁 Cargar Archivos (Los 3 son obligatorios)")
    
    col1, col2, col3 = st.columns(3)
    
    # Cronograma
    with col1:
        st.markdown("**📅 Cronograma ⚠️**")
        cronograma_file = st.file_uploader(
            "Archivo Excel*",
            key="cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("✅ Cargado")
        else:
            st.warning("⚠️ Requerido")
    
    # Asistencias
    with col2:
        st.markdown("**📊 Asistencias ⚠️**")
        asistencias_file = st.file_uploader(
            "Archivo Excel*",
            key="asistencias",
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("✅ Cargado")
        else:
            st.warning("⚠️ Requerido")
    
    # Plantilla
    with col3:
        st.markdown("**📄 Plantilla ⚠️**")
        plantilla_file = st.file_uploader(
            "Archivo Word o XML*",
            key="plantilla",
            type=['docx', 'doc', 'xml']
        )
        if plantilla_file:
            tipo = "XML" if plantilla_file.name.endswith('.xml') else "Word"
            st.success(f"✅ {tipo}")
        else:
            st.warning("⚠️ Requerido")
    
    # Verificar si todos están cargados
    todos_cargados = cronograma_file and asistencias_file and plantilla_file
    
    if not todos_cargados:
        st.info("📌 **Sube los 3 archivos para continuar**")
        
        with st.expander("ℹ️ Información sobre los archivos"):
            st.markdown("""
            **Archivos necesarios:**
            
            1. **Cronograma** - Excel con fechas y módulos
            2. **Asistencias** - Excel de control (formato: `XXXX_CTRL_Tareas_AREA.xlsx`)
            3. **Plantilla** - Word (.docx) o XML (.xml) con marcadores
            
            **💡 RECOMENDADO: Usar XML para evitar bloqueos**
            
            Si tu documento Word está bloqueado por ser oficial:
            1. Extrae el XML del Word (ver guía)
            2. Sube el .xml en lugar del .docx
            3. ¡Sin problemas de bloqueos!
            
            **Datos del centro (se rellenan automáticamente):**
            - Centro: INTERPROS NEXT GENERATION SLU
            - Código: ADGG0408 / 26615
            - Dirección: C/ DR. SEVERO OCHOA, 21, BJ
            - Localidad: AVILÉS
            - C.P.: 33400
            - Provincia: ASTURIAS
            """)
        return
    
    # Procesar archivos
    st.markdown("---")
    
    try:
        with st.spinner('🔄 Procesando archivos...'):
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
        
        st.success("✅ Datos procesados correctamente")
        
        # Mostrar resumen
        st.markdown("### 📊 Resumen de Datos")
        
        with st.expander("Ver datos extraídos", expanded=True):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Curso", datos['curso_codigo'])
            with col2:
                st.metric("Alumnos", len(datos['alumnos']))
            with col3:
                st.metric("Asistencia Media", f"{datos['estadisticas_grupales']['porcentaje_asistencia_media']}%")
            
            st.markdown("---")
            
            # Tabla de alumnos
            df = pd.DataFrame([{
                'Nº': idx + 1,
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Total Horas': f"{a['total_asistidas']}/{a['total_horas']}",
                '%': f"{a['porcentaje_asistencia']}%"
            } for idx, a in enumerate(datos['alumnos'])])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Generación masiva
        st.markdown("---")
        st.markdown("### 🚀 Generar Actas")
        
        if st.button("📦 Generar TODAS las Actas (Word)", type="primary", use_container_width=True):
            try:
                alumnos = datos['alumnos']
                total = len(alumnos)
                
                with st.spinner(f'Generando {total} actas...'):
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        plantilla_file.seek(0)
                        plantilla_bytes = plantilla_file.read()
                        
                        progress = st.progress(0)
                        status = st.empty()
                        
                        for idx, alumno in enumerate(alumnos):
                            progress.progress((idx + 1) / total)
                            status.text(f"📄 {idx + 1}/{total}: {alumno['nombre'][:40]}")
                            
                            datos_alumno = {
                                'alumno': alumno,
                                'curso': {
                                    'nombre': datos['curso_nombre'],
                                    'codigo': datos['curso_codigo']
                                }
                            }
                            
                            # Detectar si es XML o DOCX
                            es_xml = plantilla_file.name.endswith('.xml')
                            
                            gen = WordGeneratorSEPE(plantilla_bytes, es_xml=es_xml)
                            doc = gen.generar_informe_individual(datos_alumno)
                            
                            nombre = alumno['nombre'].replace(' ', '_').replace(',', '')[:50]
                            zf.writestr(f"{nombre}.docx", doc)
                        
                        progress.progress(1.0)
                        status.text(f"✅ {total} actas generadas")
                    
                    zip_buffer.seek(0)
                    st.session_state['zip_actas'] = zip_buffer.getvalue()
                    st.session_state['nombre_zip'] = f"Actas_{datos['curso_codigo'].replace('/', '_')}.zip"
                
                st.balloons()
                st.success(f"🎉 {total} actas generadas correctamente")
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
        
        # Descarga
        if 'zip_actas' in st.session_state:
            st.markdown("---")
            st.download_button(
                label="📥 Descargar ZIP con todas las actas",
                data=st.session_state['zip_actas'],
                file_name=st.session_state['nombre_zip'],
                mime="application/zip",
                use_container_width=True
            )
        
        # Vista individual
        st.markdown("---")
        st.markdown("### 🔍 Vista Previa Individual")
        
        col_a, col_b = st.columns([3, 1])
        
        with col_a:
            nombres = [a['nombre'] for a in datos['alumnos']]
            alumno_sel = st.selectbox("Seleccionar alumno:", nombres)
        
        with col_b:
            if st.button("Generar"):
                try:
                    alumno = [a for a in datos['alumnos'] if a['nombre'] == alumno_sel][0]
                    
                    datos_ind = {
                        'alumno': alumno,
                        'curso': {
                            'nombre': datos['curso_nombre'],
                            'codigo': datos['curso_codigo']
                        }
                    }
                    
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    es_xml = plantilla_file.name.endswith('.xml')
                    
                    gen = WordGeneratorSEPE(plantilla_bytes, es_xml=es_xml)
                    doc = gen.generar_informe_individual(datos_ind)
                    
                    st.session_state['acta_ind'] = doc
                    st.session_state['nombre_ind'] = f"{alumno_sel.replace(' ', '_')[:40]}.docx"
                    st.success("✅ Generado")
                    
                except Exception as e:
                    st.error(f"❌ {str(e)}")
        
        if 'acta_ind' in st.session_state:
            st.download_button(
                label="⬇️ Descargar acta individual",
                data=st.session_state['acta_ind'],
                file_name=st.session_state['nombre_ind'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        
    except Exception as e:
        st.error(f"❌ Error al procesar: {str(e)}")
        st.exception(e)
        
        with st.expander("💡 Ayuda"):
            st.markdown("""
            **Verifica:**
            - Archivo de asistencias correcto
            - Plantilla con marcadores: `{{NOMBRE_ALUMNO}}`, `{{DNI}}`, etc.
            - Archivos no corruptos
            """)


# Verificar que la función existe
if __name__ == "__main__":
    print("✅ Módulo desempleados cargado correctamente")
    print("✅ Función render_tab_desempleados disponible")