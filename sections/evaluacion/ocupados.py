"""
Interfaz de Evaluación - Ocupados
Genera 3 tipos de actas: Individual, Grupal y Certificados
LAYOUT MEJORADO: 2 columnas arriba, 1 abajo
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
    print("excel_processor importado")
    from .word_generator import WordGeneratorSEPE
    print("word_generator importado")
    from .word_generator_helper import generar_zip_todos_alumnos
    print("word_generator_helper importado")
    from .cronograma_processor import CronogramaProcessor
    print("cronograma_processor importado")
    from .word_generator_grupal import WordGeneratorActaGrupal
    print("word_generator_grupal importado")
    from .transversales_processor import TransversalesProcessor
    print("transversales_processor importado")
    from .word_generator_transversal import WordGeneratorTransversal
    print("word_generator_transversal importado")
except Exception as e:
    st.error(f"Error importando módulos: {e}")
    import traceback
    st.error(traceback.format_exc())
    raise

try:
    from sections.evaluacion.ocupados_certificaciones_processor import procesar_certificaciones
    from sections.evaluacion.ocupados_certificaciones_generator import generar_certificado_ocupado
    CERTIFICACIONES_DISPONIBLE = True
except ImportError as e:
    st.warning(f"Módulo de certificaciones no disponible: {e}")
    CERTIFICACIONES_DISPONIBLE = False
except Exception as e:
    st.error(f"Error importando módulo de certificaciones: {e}")
    CERTIFICACIONES_DISPONIBLE = False


PLANTILLA_POR_DEFECTO = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 
    'plantilla_oficial.docx'
)

PLANTILLA_CERTIFICACION_POR_DEFECTO = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), 
    'plantilla_certificacion_ocupados.docx'
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
                        print(f"Plantilla cargada desde: {ubicacion}")
                        return contenido
        
        print("No se encontró plantilla en ninguna ubicación")
        return None
        
    except Exception as e:
        print(f"Error cargando plantilla: {e}")
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
                        print(f"Plantilla grupal cargada desde: {ubicacion}")
                        return contenido
        
        print("No se encontró plantilla grupal")
        return None
        
    except Exception as e:
        print(f"Error cargando plantilla grupal: {e}")
        return None


def cargar_plantilla_certificacion_por_defecto():
    """Carga la plantilla de certificación integrada"""
    try:
        ubicaciones = [
            PLANTILLA_CERTIFICACION_POR_DEFECTO,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'plantilla_certificacion_ocupados.docx'),
        ]
        
        for ubicacion in ubicaciones:
            if os.path.exists(ubicacion):
                with open(ubicacion, 'rb') as f:
                    contenido = f.read()
                    if len(contenido) > 1000:
                        print(f"Plantilla certificación cargada desde: {ubicacion}")
                        return contenido
        
        print("No se encontró plantilla de certificación")
        return None
        
    except Exception as e:
        print(f"Error cargando plantilla certificación: {e}")
        return None


def render_tab_ocupados():
    """Render tab para ocupados con selector de tipo de acta"""
    
    st.markdown("## Generador de Actas - Ocupados")
    st.markdown("### Tipo de Acta")
    st.markdown("Selecciona el tipo de acta a generar:")
    
    if 'ocupados_tipo_acta' not in st.session_state:
        st.session_state.ocupados_tipo_acta = "individual"
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if st.button("Acta Individual", 
                     key="btn_individual_ocupados",
                     use_container_width=True):
            st.session_state.ocupados_tipo_acta = "individual"
            st.rerun()
    
    with col2:
        if st.button("Acta Grupal", 
                     key="btn_grupal_ocupados",
                     use_container_width=True):
            st.session_state.ocupados_tipo_acta = "grupal"
            st.rerun()
    
    with col3:
        if st.button("Certificados", 
                     key="btn_certificados_ocupados",
                     use_container_width=True):
            st.session_state.ocupados_tipo_acta = "certificados"
            st.rerun()
    
    tipo_acta = st.session_state.ocupados_tipo_acta
    
    st.markdown("---")

    if tipo_acta == "individual":
        render_individual()
    elif tipo_acta == "grupal":
        render_grupal()
    else:
        render_certificados()


def render_individual():
    """Render para actas individuales - LAYOUT 2-1"""
    
    st.markdown("### Acta Individual")
    st.markdown("Genera informes individualizados para cada alumno")
    st.markdown("### Archivos")
    
    # Fila superior: 2 columnas
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Cronograma**")
        cronograma_file = st.file_uploader(
            "Excel cronograma*",
            key="ocupados_individual_cronograma",
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
            key="ocupados_individual_asistencias",
            type=['xlsx', 'xls']
        )
        if asistencias_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    # Fila inferior: Plantilla
    st.markdown("**Plantilla (Opcional)**")
    plantilla_file = st.file_uploader(
        "Archivo Word (opcional)",
        key="ocupados_individual_plantilla",
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
                    
                    if doc[:2] == b'PK':
                        extension = '.zip'
                        mime = 'application/zip'
                    else:
                        extension = '.docx'
                        mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
                    
                    st.download_button(
                        label="Descargar informe individual",
                        data=doc,
                        file_name=f"{alumno['nombre'].replace(' ', '_')}{extension}",
                        mime=mime,
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
    """Render para acta grupal - LAYOUT 2-1"""
    
    st.markdown("### Acta Grupal")
    st.markdown("Genera el acta de evaluación final con todos los alumnos del grupo")
    st.markdown("### Archivos")
    
    # Fila superior: 2 columnas
    col1, col2 = st.columns(2)
    
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
    
    # Fila inferior: Plantilla
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
            processor = ExcelProcessorReal()
            datos = processor.cargar_asistencias(asistencias_file.read())
            
            crono_processor = CronogramaProcessor()
            cronograma_file.seek(0)
            datos_cronograma = crono_processor.cargar_cronograma(cronograma_file.read())
        
        st.success("Datos procesados correctamente")
        st.markdown("### Resumen del Grupo")
        
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
        st.markdown("### Generar Acta Grupal")
        
        if st.button("Generar Acta Grupal", 
                    type="primary", 
                    use_container_width=True,
                    key="ocupados_grupal_generar"):
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
                    st.info("Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_grupal_por_defecto()
                    if plantilla_bytes:
                        st.info("Usando plantilla oficial SEPE predeterminada")
                    else:
                        st.error("No se encontró la plantilla predeterminada")
                        st.warning("Sube una plantilla manualmente")
                        return
                
                with st.spinner('Generando acta grupal...'):
                    gen = WordGeneratorMultipaginaDuplicaTodo(plantilla_bytes)
                    doc = gen.generar_acta_grupal(datos_acta)
                    
                    nombre, mime = detectar_tipo_archivo(doc, datos['curso_codigo'], "Grupal_Ocupados")
                    
                    st.session_state['acta_grupal_ocupados'] = doc
                    st.session_state['nombre_acta_grupal_ocupados'] = nombre
                    st.session_state['mime_acta_grupal_ocupados'] = mime
                
                st.balloons()
                st.success("¡Acta grupal generada correctamente!")
                
            except Exception as e:
                st.error(f"Error generando acta: {str(e)}")
                st.exception(e)
        
        if 'acta_grupal_ocupados' in st.session_state:
            st.markdown("---")
            st.markdown("### Descargar")
            
            st.download_button(
                label="Descargar Acta Grupal",
                data=st.session_state['acta_grupal_ocupados'],
                file_name=st.session_state['nombre_acta_grupal_ocupados'],
                mime=st.session_state.get('mime_acta_grupal_ocupados', 'application/zip'),
                type="primary",
                use_container_width=True,
                key="ocupados_grupal_download"
            )
    
    except Exception as e:
        st.error(f"Error procesando archivos: {str(e)}")
        st.exception(e)


def render_certificados():
    """Render para certificados - LAYOUT 2-1"""
    
    st.markdown("### Certificados de Ocupados")
    st.markdown("Genera certificados individuales automáticamente desde PDF y Excel")
    
    if not CERTIFICACIONES_DISPONIBLE:
        st.error("El módulo de certificaciones no está disponible")
        st.info("""
        **Para habilitar esta funcionalidad:**
        
        1. Copia estos archivos a `sections/evaluacion/`:
           - ocupados_certificaciones_processor.py
           - ocupados_certificaciones_generator.py
           - plantilla_certificacion_ocupados.docx
        
        2. Reinicia la aplicación
        """)
        return
    
    st.markdown("### Archivos Necesarios")
    
    # Fila superior: 2 columnas
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**PDF Justificante**")
        pdf_file = st.file_uploader(
            "PDF Justificante Comunicación*",
            key="ocupados_cert_pdf",
            type=['pdf'],
            help="PDF con datos del curso y alumnos"
        )
        if pdf_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    with col2:
        st.markdown("**Excel Calificaciones**")
        excel_file = st.file_uploader(
            "Excel Informe Estadístico*",
            key="ocupados_cert_excel",
            type=['xlsx', 'xls'],
            help="Excel con calificaciones de los alumnos"
        )
        if excel_file:
            st.success("Cargado")
        else:
            st.warning("Requerido")
    
    # Fila inferior: Plantilla
    st.markdown("**Plantilla (Opcional)**")
    plantilla_cert_file = st.file_uploader(
        "Plantilla Word (opcional)",
        key="ocupados_cert_plantilla",
        type=['docx'],
        help="Si no subes ninguna, se usará la plantilla oficial"
    )
    if plantilla_cert_file:
        st.success("Personalizada")
    else:
        st.info("Por defecto")
    
    with st.expander("Información", expanded=False):
        st.markdown("""
        **Archivos necesarios:**
        
        1. **PDF Justificante** - Comunicación con datos del curso y listado de alumnos
        2. **Excel Calificaciones** - Informe estadístico con notas finales
        3. **Plantilla** - Certificado oficial (opcional, se usa plantilla por defecto)
        
        **Datos extraídos automáticamente:**
        
        Del PDF:
        - Expediente completo
        - Código del curso
        - Código y nombre del módulo
        - Nivel y horas
        - Fechas inicio/fin
        - DNI y nombre de alumnos
        
        Del Excel:
        - Calificaciones (S-9, S-10, NS, etc.)
        
        **Nomenclatura de archivos:**
        - Formato: NOMBRE_APELLIDO1_APELLIDO2_DNI.docx
        - Ejemplo: ANTONIO_AZUAGA_LOBATO_02648439N.docx
        
        **Datos fijos incluidos:**
        - Director: PABLO LUIS LOBATO MURIENTE
        - Centro: INTERPROS NEXT GENERATION S.L.U.
        - Código centro: 26615
        - Ciudad: AVILÉS
        """)
    
    if not pdf_file or not excel_file:
        st.info("Sube el PDF y Excel para continuar")
        return
    
    st.markdown("---")
    
    try:
        import tempfile
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_pdf:
            tmp_pdf.write(pdf_file.read())
            pdf_path = tmp_pdf.name
        
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_excel:
            tmp_excel.write(excel_file.read())
            excel_path = tmp_excel.name
        
        with st.spinner('Extrayendo datos del PDF y Excel...'):
            datos_completos = procesar_certificaciones(pdf_path, excel_path)
        
        os.unlink(pdf_path)
        os.unlink(excel_path)
        
        st.success(f"Datos extraídos correctamente - {len(datos_completos)} alumnos encontrados")
        st.markdown("### Resumen de Datos")
        
        with st.expander("Ver datos extraídos", expanded=True):

            col1, col2, col3, col4 = st.columns(4)
            
            primer = datos_completos[0]
            
            with col1:
                st.metric("Total Alumnos", len(datos_completos))
            
            with col2:
                aptos = sum(1 for d in datos_completos if d['calificacion'] != 'NS')
                st.metric("APTOS", aptos)
            
            with col3:
                no_aptos = len(datos_completos) - aptos
                st.metric("NO APTOS", no_aptos)
            
            with col4:
                st.metric("Horas", primer['horas'])
            
            st.markdown("---")
            
            st.markdown("**Información del Curso**")
            col1, col2 = st.columns(2)
            with col1:
                st.text(f"Expediente: {primer['expediente']}")
                st.text(f"Código curso: {primer['codigo_curso']}")
                st.text(f"Código módulo: {primer['codigo_modulo']}")
            with col2:
                st.text(f"Nivel: {primer['nivel']}")
                st.text(f"Fecha inicio: {primer['fecha_inicio']}")
                st.text(f"Fecha fin: {primer['fecha_fin']}")
            
            st.markdown("---")

            st.markdown("**Listado de Alumnos**")
            df = pd.DataFrame([{
                'Nº': idx + 1,
                'Alumno': d['nombre_alumno'],
                'DNI': d['dni_alumno'],
                'Calificación': d['calificacion'],
                'Resultado': 'APTO' if d['calificacion'] != 'NS' else 'NO APTO'
            } for idx, d in enumerate(datos_completos)])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("### Generar Certificados")
        
        if st.button("Generar TODOS los Certificados", 
                    type="primary", 
                    use_container_width=True,
                    key="ocupados_cert_generar_todos"):
            try:
                total = len(datos_completos)

                if plantilla_cert_file:
                    plantilla_cert_file.seek(0)
                    plantilla_bytes = plantilla_cert_file.read()
                    st.info("Usando plantilla personalizada")
                else:
                    plantilla_bytes = cargar_plantilla_certificacion_por_defecto()
                    if plantilla_bytes:
                        st.info("Usando plantilla oficial predeterminada")
                    else:
                        st.error("No se encontró la plantilla de certificación")
                        st.warning("Sube una plantilla manualmente")
                        return
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.docx') as tmp_plantilla:
                    tmp_plantilla.write(plantilla_bytes)
                    plantilla_path = tmp_plantilla.name
                
                with st.spinner(f'Generando {total} certificados...'):
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                        progress = st.progress(0)
                        status = st.empty()
                        
                        for idx, datos_alumno in enumerate(datos_completos):
                            progress.progress((idx + 1) / total)
                            status.text(f"{idx + 1}/{total}: {datos_alumno['nombre_alumno'][:40]}")
                            
                            certificado_bytes, nombre_archivo = generar_certificado_ocupado(
                                plantilla_path, 
                                datos_alumno
                            )
                            
                            zipf.writestr(f"{nombre_archivo}.docx", certificado_bytes)
                        
                        progress.progress(1.0)
                        status.text(f"{total} certificados generados")
                    
                    zip_buffer.seek(0)
                    st.session_state['zip_certificados_ocupados'] = zip_buffer.getvalue()
                    st.session_state['nombre_zip_certificados'] = f"Certificados_Ocupados_{primer['expediente'].replace('/', '_')}.zip"
                
                os.unlink(plantilla_path)
                
                st.balloons()
                st.success(f"{total} certificados generados correctamente")
                
            except Exception as e:
                st.error(f"Error generando certificados: {str(e)}")
                st.exception(e)
        

        if 'zip_certificados_ocupados' in st.session_state:
            st.markdown("---")
            st.markdown("### Descargar")
            
            st.download_button(
                label="Descargar ZIP con todos los certificados",
                data=st.session_state['zip_certificados_ocupados'],
                file_name=st.session_state['nombre_zip_certificados'],
                mime="application/zip",
                type="primary",
                use_container_width=True,
                key="ocupados_cert_download_all"
            )
    
    except Exception as e:
        st.error(f"Error procesando archivos: {str(e)}")
        import traceback
        st.code(traceback.format_exc())


if __name__ == "__main__":
    render_tab_ocupados()