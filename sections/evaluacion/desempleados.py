"""
Interfaz de Evaluación - Desempleados
CON PLANTILLA INTEGRADA POR DEFECTO + TRANSVERSALES
"""
import streamlit as st
import pandas as pd
from io import BytesIO
import zipfile
import os
import sys

from sections.evaluacion.word_generator_grupal import WordGeneratorMultipaginaDuplicaTodo
sys.path.append(os.path.dirname(os.path.abspath(__file__)))


def detectar_tipo_archivo(doc_bytes: bytes, curso_codigo: str, tipo: str = "Grupal") -> tuple:
    """
    Detecta si el documento es ZIP o DOCX y prepara los datos correctos para descarga
    
    Args:
        doc_bytes: Bytes del documento generado
        curso_codigo: Código del curso
        tipo: Tipo de acta
    
    Returns:
        tuple: (nombre_archivo, mime_type)
    """

    if doc_bytes[:4] == b'PK\x03\x04':
        try:
            with zipfile.ZipFile(BytesIO(doc_bytes), 'r') as zf:
                archivos = zf.namelist()
                if any(f.endswith('.docx') for f in archivos):
                    extension = '.zip'
                    mime_type = 'application/zip'
                    prefijo = 'Actas'
                else:
                    extension = '.docx'
                    mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    prefijo = 'Acta'
        except:
            extension = '.docx'
            mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            prefijo = 'Acta'
    else:
        extension = '.docx'
        mime_type = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        prefijo = 'Acta'
    
    codigo_limpio = curso_codigo.replace('/', '_').replace('\\', '_')
    nombre_archivo = f"{prefijo}_{tipo}_{codigo_limpio}{extension}"
    
    return nombre_archivo, mime_type


try:
    from .excel_processor import ExcelProcessorReal
    print("✅ excel_processor importado")
    from .word_generator import WordGeneratorSEPE
    print("✅ word_generator importado")
    from .word_generator_helper import generar_zip_todos_alumnos
    print("✅ word_generator_helper importado")
    from .cronograma_processor import CronogramaProcessor
    print("✅ cronograma_processor importado")
    from .word_generator_grupal import WordGeneratorActaGrupal
    print("✅ word_generator_grupal importado")
    from .transversales_processor import TransversalesProcessor
    print("✅ transversales_processor importado")
    from .word_generator_transversal import WordGeneratorTransversal
    print("✅ word_generator_transversal importado")
except Exception as e:
    st.error(f"❌ Error importando módulos: {e}")
    import traceback
    st.error(traceback.format_exc())
    raise

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
        
        print("⚠ No se encontró plantilla en ninguna ubicación")
        return None
        
    except Exception as e:
        print(f"⚠ Error cargando plantilla: {e}")
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
        
        print("⚠ No se encontró plantilla grupal")
        return None
        
    except Exception as e:
        print(f"⚠ Error cargando plantilla grupal: {e}")
        return None


def cargar_plantilla_transversal_por_defecto():
    """Carga la plantilla de acta transversal integrada en la aplicación"""
    try:
        plantilla_transversal = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 
            'plantilla_transversal_oficial.docx'
        )
        
        ubicaciones = [
            plantilla_transversal,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'plantilla_transversal_oficial.docx'),
        ]
        
        for ubicacion in ubicaciones:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'rb') as f:
                    contenido = f.read()
                    if len(contenido) > 1000:
                        print(f"✓ Plantilla transversal cargada desde: {ubicacion}")
                        return contenido
        
        print("⚠ No se encontró plantilla transversal")
        return None
        
    except Exception as e:
        print(f"⚠ Error cargando plantilla transversal: {e}")
        return None


def render_tab_desempleados():
    """Render tab para desempleados con selector de tipo de acta"""
    
    st.markdown("## Generador de Actas - Desempleados")
    st.markdown("### Tipo de Acta")
    st.markdown("Selecciona el tipo de acta a generar:")
    
    if 'desempleados_tipo_acta' not in st.session_state:
        st.session_state.desempleados_tipo_acta = "individual"
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("📄 Acta Individual", 
                     key="btn_individual_desempleados",
                     use_container_width=True):
            st.session_state.desempleados_tipo_acta = "individual"
            st.rerun()
    
    with col2:
        if st.button("📋 Acta Grupal", 
                     key="btn_grupal_desempleados",
                     use_container_width=True):
            st.session_state.desempleados_tipo_acta = "grupal"
            st.rerun()
    
    with col3:
        if st.button("🎓 Transversales", 
                     key="btn_transversales_desempleados",
                     use_container_width=True):
            st.session_state.desempleados_tipo_acta = "transversales"
            st.rerun()
    
    tipo_acta = st.session_state.desempleados_tipo_acta
    
    st.markdown("---")
    
    if tipo_acta == "individual":
        render_individual()
    elif tipo_acta == "grupal":
        render_grupal()
    else:
        render_transversales()


def render_individual():
    """Render para actas individuales"""
    
    st.markdown("### 📄 Acta Individual")
    st.markdown("Genera informes individualizados para cada alumno")
    st.markdown("### 📁 Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="desempleados_individual_cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col2:
        st.markdown("**Asistencias**")
        asistencias_file = st.file_uploader(
            "Excel control asistencias*",
            key="desempleados_individual_asistencias",
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="desempleados_individual_plantilla",
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial SEPE predeterminada"
        )
        if plantilla_file:
            st.success("✓ Personalizada")
        else:
            st.info("ℹ Por defecto")
    
    with st.expander("ℹ Información", expanded=False):
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
        st.info("ℹ Sube al menos el cronograma y asistencias para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('⏳ Procesando archivos...'):
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
        
        st.success("✅ Datos procesados correctamente")
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
            
            df = pd.DataFrame([{
                'Nº': idx + 1,
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Total Horas': f"{a['total_asistidas']}/{a['total_horas']}",
                '%': f"{a['porcentaje_asistencia']}%"
            } for idx, a in enumerate(datos['alumnos'])])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🚀 Generar Actas")
        
        if st.button("📦 Generar TODAS las Actas (ZIP)", type="primary", use_container_width=True, key="desempleados_individual_generar_todas"):
            try:
                alumnos = datos['alumnos']
                total = len(alumnos)
                
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("ℹ Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_por_defecto()
                    if plantilla_bytes:
                        st.info("ℹ Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("❌ No se pudo cargar la plantilla predeterminada")
                        st.warning("⚠ Sube una plantilla manualmente")
                        return
                
                with st.spinner(f'⏳ Generando {total} actas...'):
                    # CÓDIGO CORREGIDO - Usar helper
                    lista_datos = []
                    for alumno in alumnos:
                        datos_alumno = {
                            'alumno': alumno,
                            'curso': {
                                'nombre': datos['curso_nombre'],
                                'codigo': datos['curso_codigo']
                            }
                        }
                        lista_datos.append(datos_alumno)
                    
                    # Generar con manejo automático de ZIP
                    gen = WordGeneratorSEPE(plantilla_bytes, es_xml=False)
                    zip_bytes = generar_zip_todos_alumnos(gen, lista_datos)
                    
                    st.session_state['zip_actas_desempleados_individual'] = zip_bytes
                    st.session_state['nombre_zip_desempleados_individual'] = f"Actas_Individual_Desempleados_{datos['curso_codigo'].replace('/', '_')}.zip"
                
                st.balloons()
                st.success(f"✅ {total} actas generadas correctamente")
                
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
        
        if 'zip_actas_desempleados_individual' in st.session_state:
            st.markdown("---")
            st.markdown("### 📥 Descargar")
            
            st.download_button(
                label="📥 Descargar ZIP con todas las actas",
                data=st.session_state['zip_actas_desempleados_individual'],
                file_name=st.session_state['nombre_zip_desempleados_individual'],
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="desempleados_individual_download"
            )
        
        st.markdown("---")
        st.markdown("### 👤 Vista Individual")
        
        alumno_seleccionado = st.selectbox(
            "Selecciona un alumno",
            options=range(len(datos['alumnos'])),
            format_func=lambda x: f"{x+1}. {datos['alumnos'][x]['nombre']} - {datos['alumnos'][x]['dni']}",
            key="desempleados_individual_selector"
        )
        
        if st.button("🔍 Generar vista previa", use_container_width=True, key="desempleados_individual_preview"):
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
                    
                    # Detectar si es ZIP o DOCX
                    if doc[:2] == b'PK':
                        extension = '.zip'
                        mime = 'application/zip'
                    else:
                        extension = '.docx'
                        mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    
                    st.download_button(
                        label=f"📥 Descargar informe individual ({extension})",
                        data=doc,
                        file_name=f"{alumno['nombre'].replace(' ', '_')}{extension}",
                        mime=mime,
                        use_container_width=True,
                        key="desempleados_individual_download_one"
                    )
                else:
                    st.error("❌ No hay plantilla disponible")
                    
            except Exception as e:
                st.error(f"❌ Error: {str(e)}")
                st.exception(e)
    
    except Exception as e:
        st.error(f"❌ Error procesando archivos: {str(e)}")
        st.exception(e)


def render_grupal():
    """Render para acta grupal"""
    
    st.markdown("### 📋 Acta Grupal")
    st.markdown("Genera el acta de evaluación final con todos los alumnos del grupo")
    st.markdown("### 📁 Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="desempleados_grupal_cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col2:
        st.markdown("**Asistencias**")
        asistencias_file = st.file_uploader(
            "Excel control asistencias*",
            key="desempleados_grupal_asistencias",
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="desempleados_grupal_plantilla",
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial SEPE predeterminada"
        )
        if plantilla_file:
            st.success("✓ Personalizada")
        else:
            st.info("ℹ Por defecto")
    
    with st.expander("ℹ Información", expanded=False):
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
        st.info("ℹ Sube el cronograma y asistencias para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('⏳ Procesando archivos...'):
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
            crono_processor = CronogramaProcessor()
            cronograma_file.seek(0)
            datos_cronograma = crono_processor.cargar_cronograma(cronograma_file.read())
        
        st.success("✅ Datos procesados correctamente")
        st.markdown("### 📊 Resumen del Grupo")
        
        with st.expander("Ver datos del grupo", expanded=True):

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
            st.markdown("**Información del Curso**")
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"Curso: {datos['curso_nombre']}")
                st.text(f"Código: {datos['curso_codigo']}")
            with col2:
                st.text(f"Fecha inicio: {datos_cronograma.get('fecha_inicio', 'N/A')}")
                st.text(f"Fecha fin: {datos_cronograma.get('fecha_fin', 'N/A')}")
            
            st.markdown("---")
            st.markdown("**Listado de Alumnos**")
            df = pd.DataFrame([{
                'Nº': idx + 1,
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Asistencia': f"{a['porcentaje_asistencia']}%",
                'Resultado': 'APTO' if a['porcentaje_asistencia'] >= 75 else 'NO APTO'
            } for idx, a in enumerate(datos['alumnos'])])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🚀 Generar Acta Grupal")
        
        if st.button("📋 Generar Acta Grupal", 
                    type="primary", 
                    use_container_width=True,
                    key="desempleados_grupal_generar"):
            try:
                datos_acta = {
                    'curso_codigo': datos['curso_codigo'],
                    'curso_nombre': datos['curso_nombre'],
                    'centro': 'INTERPROS NEXT GENERATION SLU',
                    'codigo_centro': '26615',
                    'fecha_inicio': datos_cronograma.get('fecha_inicio', ''),
                    'fecha_fin': datos_cronograma.get('fecha_fin', ''),
                    'alumnos': datos['alumnos'],
                    'total_alumnos': len(datos['alumnos']),
                    'modulos_detalle': datos_cronograma.get('modulos', []),
                    'modulos_info': datos_cronograma.get('modulos', [])
                }
                
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("ℹ Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_grupal_por_defecto()
                    if plantilla_bytes:
                        st.info("ℹ Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("❌ No se encontró la plantilla predeterminada")
                        st.warning("⚠ Sube una plantilla manualmente")
                        return
                
                with st.spinner('⏳ Generando acta grupal...'):
                    gen = WordGeneratorMultipaginaDuplicaTodo(plantilla_bytes)
                    doc = gen.generar_acta_grupal(datos_acta)
                    nombre, mime = detectar_tipo_archivo(doc, datos['curso_codigo'], "Grupal_Desempleados")
                    
                    st.session_state['acta_grupal_desempleados'] = doc
                    st.session_state['nombre_acta_grupal_desempleados'] = nombre
                    st.session_state['mime_acta_grupal_desempleados'] = mime
                
                st.balloons()
                st.success("✅ ¡Acta grupal generada correctamente!")
                
            except Exception as e:
                st.error(f"❌ Error generando acta: {str(e)}")
                st.exception(e)
        
        if 'acta_grupal_desempleados' in st.session_state:
            st.markdown("---")
            st.markdown("### 📥 Descargar")
            
            st.download_button(
                label="📥 Descargar Acta Grupal",
                data=st.session_state['acta_grupal_desempleados'],
                file_name=st.session_state['nombre_acta_grupal_desempleados'],
                mime=st.session_state.get('mime_acta_grupal_desempleados', 'application/zip'),
                type="primary",
                use_container_width=True,
                key="desempleados_grupal_download"
            )
    
    except Exception as e:
        st.error(f"❌ Error procesando archivos: {str(e)}")
        st.exception(e)


def render_transversales():
    """Render para actas transversales (FCOO03)"""
    
    st.markdown("### 🎓 Actas Transversales (FCOO03)")
    st.markdown("Genera actas de evaluación final para competencias transversales")
    st.markdown("### 📁 Archivos")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="desempleados_transversal_cronograma",
            type=['xlsx', 'xls']
        )
        if cronograma_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col2:
        st.markdown("**Control de Tareas**")
        control_file = st.file_uploader(
            "Excel CTRL_Tareas_AREA*",
            key="desempleados_transversal_control",
            type=['xlsx', 'xls']
        )
        if control_file:
            st.success("✓ Cargado")
        else:
            st.warning("⚠ Requerido")
    
    with col3:
        st.markdown("**Plantilla (Opcional)**")
        plantilla_file = st.file_uploader(
            "Archivo Word (opcional)",
            key="desempleados_transversal_plantilla",
            type=['docx', 'doc'],
            help="Si no subes ninguna, se usará la plantilla oficial SEPE predeterminada"
        )
        if plantilla_file:
            st.success("✓ Personalizada")
        else:
            st.info("ℹ Por defecto")
    
    with st.expander("ℹ Información", expanded=False):
        st.markdown("""
        **Archivos necesarios:**
        
        1. **Cronograma** - Excel con fechas y datos del curso (Requerido)
        2. **Control de Tareas** - Excel CTRL_Tareas_AREA con asistencias y calificaciones (Requerido)
        3. **Plantilla** - Word con formato oficial (opcional)
        
        **Plantilla predeterminada:**
        
        Si NO subes una plantilla, se usará la plantilla oficial SEPE de actas transversales integrada.
        
        **El acta incluye:**
        - Datos del curso (Acción, Especialidad, Código FCOO03)
        - Centro: INTERPROS NEXT GENERATION S.L.U
        - Fechas de inicio y finalización
        - Tabla con todos los alumnos (hasta 20)
        - DNI, Nombre completo, Horas de actividades
        - Calificación final (APTO/NO APTO/EXENTO/A)
        - Espacios para firmas del responsable y formador/a
        """)
    
    if not cronograma_file or not control_file:
        st.info("ℹ Sube el cronograma y control de tareas para continuar")
        return
    
    st.markdown("---")
    
    try:
        with st.spinner('⏳ Procesando archivos...'):
            from .transversales_processor import TransversalesProcessor
            
            processor = TransversalesProcessor()
            cronograma_file.seek(0)
            control_file.seek(0)
            datos = processor.cargar_datos(control_file.read(), cronograma_file.read())
        
        st.success("✅ Datos procesados correctamente")
        st.markdown("### 📊 Resumen del Curso")
        
        with st.expander("Ver datos extraídos", expanded=True):

            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Alumnos", datos['total_alumnos'])
            
            with col2:
                aptos = sum(1 for a in datos['alumnos'] if 'APTO' in a['calificacion_final'].upper() and 'NO' not in a['calificacion_final'].upper())
                st.metric("APTOS", aptos)
            
            with col3:
                no_aptos = sum(1 for a in datos['alumnos'] if 'NO APTO' in a['calificacion_final'].upper())
                st.metric("NO APTOS", no_aptos)
            
            with col4:
                st.metric("Duración", f"{datos['campo_6_duracion']} horas")
            
            st.markdown("---")
        
            st.markdown("**Información del Curso**")
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"Acción: {datos['campo_2_accion']}")
                st.text(f"Código: {datos['campo_4_codigo']}")
                especialidad_corta = datos['campo_3_especialidad'][:60] + "..." if len(datos['campo_3_especialidad']) > 60 else datos['campo_3_especialidad']
                st.text(f"Especialidad: {especialidad_corta}")
            with col2:
                st.text(f"Fecha inicio: {datos['campo_9_fecha_inicio']}")
                st.text(f"Fecha fin: {datos['campo_10_fecha_fin']}")
                st.text(f"Modalidad: {datos['campo_8_modalidad']}")
            
            st.markdown("---")
            
            st.markdown("**Listado de Alumnos**")
            df = pd.DataFrame([{
                'Nº': a['numero'],
                'Alumno': a['nombre'],
                'DNI': a['dni'],
                'Horas': a['horas_actividades'],
                'Calificación': a['calificacion_final']
            } for a in datos['alumnos']])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### 🚀 Generar Acta Transversal")
        
        if st.button("🎓 Generar Acta Transversal (FCOO03)", 
                    type="primary", 
                    use_container_width=True,
                    key="desempleados_transversal_generar"):
            try:
                if plantilla_file:
                    plantilla_file.seek(0)
                    plantilla_bytes = plantilla_file.read()
                    st.info("ℹ Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_transversal_por_defecto()
                    if plantilla_bytes:
                        st.info("ℹ Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("❌ No se encontró la plantilla predeterminada")
                        st.warning("⚠ Sube una plantilla manualmente")
                        return
                
                with st.spinner('⏳ Generando acta transversal...'):
                    from .word_generator_transversal import WordGeneratorTransversal
                    
                    gen = WordGeneratorTransversal(plantilla_bytes)
                    doc = gen.generar_acta(datos)
                    
                    st.session_state['acta_transversal_desempleados'] = doc
                    st.session_state['nombre_acta_transversal_desempleados'] = f"Acta_Transversal_FCOO03_{datos['campo_2_accion'].replace('/', '_')}.docx"
                
                st.balloons()
                st.success("✅ ¡Acta transversal generada correctamente!")
                
            except Exception as e:
                st.error(f"❌ Error generando acta: {str(e)}")
                st.exception(e)
        
        if 'acta_transversal_desempleados' in st.session_state:
            st.markdown("---")
            st.markdown("### 📥 Descargar")
            
            st.download_button(
                label="📥 Descargar Acta Transversal",
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


if __name__ == "__main__":
    render_tab_desempleados()