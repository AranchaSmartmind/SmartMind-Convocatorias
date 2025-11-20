"""
Interfaz de Evaluación - Ocupados
Genera 3 tipos de actas: Individual, Grupal y Certificados
"""
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import os

# Importar los procesadores
import sys
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from excel_processor import ExcelProcessorReal
    from word_generator import WordGeneratorSEPE
    from cronograma_processor import CronogramaProcessor
    from word_generator_grupal import WordGeneratorActaGrupal
except:
    st.error("Error importando módulos")


# Rutas de plantillas por defecto (mismo archivo para las 3 por ahora)
PLANTILLA_POR_DEFECTO = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 
    'plantilla_oficial.docx'
)


def cargar_plantilla_por_defecto():
    """Carga la plantilla integrada en la aplicación"""
    try:
        ubicaciones = [
            PLANTILLA_POR_DEFECTO,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'plantilla_oficial.docx'),
            os.path.join(os.path.dirname(__file__), '..', '..', 'plantilla_oficial.docx'),
        ]
        
        for ubicacion in ubicaciones:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'rb') as f:
                    contenido = f.read()
                    if len(contenido) > 1000:
                        print(f"✓ Plantilla cargada desde: {ubicacion}")
                        return contenido
        
        print(" No se encontró plantilla en ninguna ubicación")
        return None
        
    except Exception as e:
        print(f" Error cargando plantilla: {e}")
        return None
    
def cargar_plantilla_grupal_por_defecto():
    """Carga la plantilla grupal integrada en la aplicación"""
    try:
        plantilla_grupal = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 
            'plantilla_grupal_oficial.docx'
        )
        
        ubicaciones = [
            plantilla_grupal,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'plantilla_grupal_oficial.docx'),
        ]
        
        for ubicacion in ubicaciones:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'rb') as f:
                    contenido = f.read()
                    if len(contenido) > 1000:
                        print(f"✓ Plantilla grupal cargada desde: {ubicacion}")
                        return contenido
        
        print(" No se encontró plantilla grupal")
        return None
        
    except Exception as e:
        print(f" Error cargando plantilla grupal: {e}")
        return None


def render_tab_ocupados():
    """Render tab para ocupados con selector de tipo de acta"""
    
    st.markdown("## Generador de Actas - Ocupados")
    
    # SELECTOR DE TIPO DE ACTA
    st.markdown("### Tipo de Acta")
    
    tipo_acta = st.radio(
        "Selecciona el tipo de acta a generar:",
        options=["individual", "grupal", "certificados"],
        format_func=lambda x: {
            "individual": " Acta Individual",
            "grupal": " Acta Grupal", 
            "certificados": " Certificados"
        }[x],
        key="ocupados_tipo_acta",
        horizontal=True
    )
    
    st.markdown("---")
    
    # Renderizar según el tipo seleccionado
    if tipo_acta == "individual":
        render_individual()
    elif tipo_acta == "grupal":
        render_grupal()
    else:  # certificados
        render_certificados()


def render_individual():
    """Render para actas individuales (igual que desempleados)"""
    
    st.markdown("### Acta Individual")
    st.markdown("Genera informes individualizados para cada alumno")
    
    # Subida de archivos
    st.markdown("### Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="ocupados_individual_cronograma",  # KEY ÚNICA
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col2:
        st.markdown("**Asistencias**")
        asistencias_file = st.file_uploader(
            "Excel control asistencias*",
            key="ocupados_individual_asistencias",  # KEY ÚNICA
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="ocupados_individual_plantilla",  # KEY ÚNICA
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial predeterminada"
        )
        if plantilla_file:
            st.success("Personalizada")
        else:
            st.info("Por defecto")
    
    with st.expander("Información", expanded=False):
        st.markdown("""
        **Archivos necesarios:**
        
        1. **Cronograma** - Excel con fechas y módulos (Requerido)
        2. **Asistencias** - Excel de control (Requerido)
        3. **Plantilla** - Word con formato oficial (opcional)
        
        **Plantilla predeterminada:**
        
        Si NO subes una plantilla, se usará la plantilla oficial SEPE integrada.
        
        **Datos del centro (se rellenan automáticamente):**
        - Centro: INTERPROS NEXT GENERATION SLU
        - Código: ADGG0408 / 26615
        - Dirección: C/ DR. SEVERO OCHOA, 21, BJ
        - Localidad: AVILÉS
        - C.P.: 33400
        - Provincia: ASTURIAS
        """)
    
    if not cronograma_file or not asistencias_file:
        st.info("Sube al menos el cronograma y asistencias para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('Procesando archivos...'):
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
        
        st.success("Datos procesados correctamente")
        
        # Mostrar resumen
        st.markdown("### Resumen de Datos")
        
        with st.expander("Ver datos extraídos", expanded=True):
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.metric("Curso", datos['curso_codigo'])
            with col2:
                st.metric("Alumnos", len(datos['alumnos']))
            with col3:
                st.metric("Asistencia Media", f"{datos['estadisticas_grupales']['porcentaje_asistencia_media']}%")
            
            st.markdown("---")
            
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
        st.markdown("### Generar Actas")
        
        if st.button("Generar TODAS las Actas (Word)", type="primary", use_container_width=True, key="ocupados_individual_generar_todas"):
            try:
                alumnos = datos['alumnos']
                total = len(alumnos)
                
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_por_defecto()
                    if plantilla_bytes:
                        st.info("Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("No se pudo cargar la plantilla predeterminada")
                        st.warning("Sube una plantilla manualmente")
                        return
                
                with st.spinner(f'Generando {total} actas...'):
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        progress = st.progress(0)
                        status = st.empty()
                        
                        for idx, alumno in enumerate(alumnos):
                            progress.progress((idx + 1) / total)
                            status.text(f"{idx + 1}/{total}: {alumno['nombre'][:40]}")
                            
                            datos_alumno = {
                                'alumno': alumno,
                                'curso': {
                                    'nombre': datos['curso_nombre'],
                                    'codigo': datos['curso_codigo']
                                }
                            }
                            
                            gen = WordGeneratorSEPE(plantilla_bytes, es_xml=False)
                            doc = gen.generar_informe_individual(datos_alumno)
                            
                            nombre = alumno['nombre'].replace(' ', '_').replace(',', '')[:50]
                            zf.writestr(f"{nombre}.docx", doc)
                        
                        progress.progress(1.0)
                        status.text(f"{total} actas generadas")
                    
                    zip_buffer.seek(0)
                    st.session_state['zip_actas_ocupados_individual'] = zip_buffer.getvalue()
                    st.session_state['nombre_zip_ocupados_individual'] = f"Actas_Individual_Ocupados_{datos['curso_codigo'].replace('/', '_')}.zip"
                
                st.balloons()
                st.success(f"{total} actas generadas correctamente")
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.exception(e)
        
        # Descarga
        if 'zip_actas_ocupados_individual' in st.session_state:
            st.markdown("---")
            st.markdown("### Descargar")
            
            st.download_button(
                label="Descargar ZIP con todas las actas",
                data=st.session_state['zip_actas_ocupados_individual'],
                file_name=st.session_state['nombre_zip_ocupados_individual'],
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="ocupados_individual_download"
            )
        
        # Vista individual
        st.markdown("---")
        st.markdown("### Vista Individual")
        
        alumno_seleccionado = st.selectbox(
            "Selecciona un alumno",
            options=range(len(datos['alumnos'])),
            format_func=lambda x: f"{x+1}. {datos['alumnos'][x]['nombre']} - {datos['alumnos'][x]['dni']}",
            key="ocupados_individual_selector"
        )
        
        if st.button("Generar vista previa", use_container_width=True, key="ocupados_individual_preview"):
            try:
                alumno = datos['alumnos'][alumno_seleccionado]
                
                datos_ind = {
                    'alumno': alumno,
                    'curso': {
                        'nombre': datos['curso_nombre'],
                        'codigo': datos['curso_codigo']
                    }
                }
                
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                else:
                    plantilla_bytes = cargar_plantilla_por_defecto()
                
                if plantilla_bytes:
                    gen = WordGeneratorSEPE(plantilla_bytes, es_xml=False)
                    doc = gen.generar_informe_individual(datos_ind)
                    
                    st.download_button(
                        label="Descargar informe individual",
                        data=doc,
                        file_name=f"{alumno['nombre'].replace(' ', '_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="ocupados_individual_download_one"
                    )
                else:
                    st.error("No hay plantilla disponible")
                    
            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.exception(e)
    
    except Exception as e:
        st.error(f"Error procesando archivos: {str(e)}")
        st.exception(e)


def render_grupal():
    """Render para acta grupal"""
    
    st.markdown("### Acta Grupal")
    st.markdown("Genera el acta de evaluación final con todos los alumnos del grupo")
    
    # Subida de archivos
    st.markdown("### Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="ocupados_grupal_cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col2:
        st.markdown("**Asistencias**")
        asistencias_file = st.file_uploader(
            "Excel control asistencias*",
            key="ocupados_grupal_asistencias",
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="ocupados_grupal_plantilla",
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial SEPE predeterminada"
        )
        if plantilla_file:
            st.success("Personalizada")
        else:
            st.info("Por defecto")
    
    with st.expander("Información", expanded=False):
        st.markdown("""
        **Archivos necesarios:**
        
        1. **Cronograma** - Excel con fechas del curso (Requerido)
        2. **Asistencias** - Excel de control (Requerido)
        3. **Plantilla** - Word con formato oficial (opcional)
        
        **Plantilla predeterminada:**
        
        Si NO subes una plantilla, se usará la plantilla oficial SEPE de acta grupal integrada.
        
        **El acta incluye:**
        - Datos del curso y centro
        - Fechas de inicio y finalización
        - Tabla con todos los alumnos
        - Calificaciones por módulo
        - Resultado final (APTO/NO APTO)
        - Firmas y observaciones
        """)
    
    if not cronograma_file or not asistencias_file:
        st.info("Sube el cronograma y asistencias para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('Procesando archivos...'):
            # Procesar asistencias
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
            
            # Procesar cronograma
            crono_processor = CronogramaProcessor()
            cronograma_file.seek(0)
            datos_cronograma = crono_processor.cargar_cronograma(cronograma_file.read())
        
        st.success("Datos procesados correctamente")
        
        # Mostrar resumen
        st.markdown("### Resumen del Grupo")
        
        with st.expander("Ver datos del grupo", expanded=True):
            # Métricas
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Alumnos", len(datos['alumnos']))
            
            with col2:
                aptos = sum(1 for a in datos['alumnos'] 
                           if a['porcentaje_asistencia'] >= 75)
                st.metric("APTOS", aptos)
            
            with col3:
                no_aptos = len(datos['alumnos']) - aptos
                st.metric("NO APTOS", no_aptos)
            
            with col4:
                st.metric("Asistencia Media", 
                         f"{datos['estadisticas_grupales']['porcentaje_asistencia_media']}%")
            
            st.markdown("---")
            
            # Información del curso
            st.markdown("**Información del Curso**")
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"Curso: {datos['curso_nombre']}")
                st.text(f"Código: {datos['curso_codigo']}")
            with col2:
                st.text(f"Fecha inicio: {datos_cronograma.get('fecha_inicio', 'N/A')}")
                st.text(f"Fecha fin: {datos_cronograma.get('fecha_fin', 'N/A')}")
            
            st.markdown("---")
            
            # Tabla de alumnos
            st.markdown("**Listado de Alumnos**")
            df = pd.DataFrame([{
                'Nº': idx + 1,
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Asistencia': f"{a['porcentaje_asistencia']}%",
                'Resultado': 'APTO' if a['porcentaje_asistencia'] >= 75 else 'NO APTO'
            } for idx, a in enumerate(datos['alumnos'])])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        # Generar acta
        st.markdown("---")
        st.markdown("### Generar Acta Grupal")
        
        if st.button("Generar Acta Grupal", 
                    type="primary", 
                    use_container_width=True,
                    key="ocupados_grupal_generar"):
            try:
                # Preparar datos para el acta
                datos_acta = {
                    'curso_codigo': datos['curso_codigo'],
                    'curso_nombre': datos['curso_nombre'],
                    'centro': 'INTERPROS NEXT GENERATION SLU',
                    'codigo_centro': '26615',
                    'fecha_inicio': datos_cronograma.get('fecha_inicio', ''),
                    'fecha_fin': datos_cronograma.get('fecha_fin', ''),
                    'alumnos': datos['alumnos']
                }
                
                # Obtener plantilla (personalizada o por defecto)
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_grupal_por_defecto()
                    if plantilla_bytes:
                        st.info("Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("No se encontró la plantilla predeterminada")
                        st.warning("Sube una plantilla manualmente")
                        return
                
                # Generar acta
                with st.spinner('Generando acta grupal...'):
                    gen = WordGeneratorActaGrupal(plantilla_bytes)
                    doc = gen.generar_acta_grupal(datos_acta)
                    
                    # Guardar en session state
                    st.session_state['acta_grupal_ocupados'] = doc
                    st.session_state['nombre_acta_grupal_ocupados'] = f"Acta_Grupal_Ocupados_{datos['curso_codigo'].replace('/', '_')}.docx"
                
                st.balloons()
                st.success("¡Acta grupal generada correctamente!")
                
            except Exception as e:
                st.error(f"Error generando acta: {str(e)}")
                st.exception(e)
        
        # Descarga
        if 'acta_grupal_ocupados' in st.session_state:
            st.markdown("---")
            st.markdown("### Descargar")
            
            st.download_button(
                label="Descargar Acta Grupal",
                data=st.session_state['acta_grupal_ocupados'],
                file_name=st.session_state['nombre_acta_grupal_ocupados'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
                key="ocupados_grupal_download"
            )
    
    except Exception as e:
        st.error(f"Error procesando archivos: {str(e)}")
        st.exception(e)


def render_certificados():
    """Render para certificados"""
    
    st.markdown("### Certificados")
    st.markdown("Genera certificados para los alumnos aprobados")
    
    st.info(" Funcionalidad en desarrollo")
    st.markdown("""
    **Próximamente:**
    - Certificados individuales
    - Validación de requisitos de aprobación
    - Generación masiva de certificados
    - Personalización de plantillas
    """)


if __name__ == "__main__":
    render_tab_ocupados()