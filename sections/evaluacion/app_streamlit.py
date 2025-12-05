import streamlit as st
import os
import sys
import tempfile
from io import StringIO

from cierre_mes.extraccion_ocr import calcular_dias_lectivos_y_asistencias
from cierre_mes.procesamiento_datos import (
    obtener_mes_anterior,
    extraer_datos_curso_pdf,
    extraer_alumnos_excel,
    extraer_becas_ayudas_simple,
    extraer_justificantes
)
from cierre_mes.generacion_word import construir_observaciones, generar_parte_mensual
from cierre_mes.utilidades import normalizar_nombre, buscar_coincidencia

def procesar_cierre_mensual(archivos_config):
    print("INICIANDO CIERRE MENSUAL")
    
    try:
        mes = obtener_mes_anterior()
        print(f"Mes a procesar: {mes}")
        
        numero_curso, especialidad = extraer_datos_curso_pdf(archivos_config['pdf_becas'])
        print(f"Curso: {numero_curso}")
        
        alumnos_excel = extraer_alumnos_excel(archivos_config['excel_alumnos'])
        print(f"{len(alumnos_excel)} alumnos encontrados")
        
        ayudas_por_alumno = extraer_becas_ayudas_simple(archivos_config['pdf_becas'])
        print(f"{len(ayudas_por_alumno)} alumnos con ayudas")
        
        justificantes_por_alumno = extraer_justificantes(archivos_config['pdf_justificantes'])
        print(f"{len(justificantes_por_alumno)} alumnos con justificantes")
        
        dias_lectivos, asistencias_por_alumno, faltas_por_alumno = calcular_dias_lectivos_y_asistencias(
            archivos_config['pdfs_firmas']
        )
        print(f"Dias lectivos totales: {dias_lectivos}")
        
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
                'observaciones': observaciones,
                'ayudas': ayudas,
                'justificantes': justificantes,
                'dias_aula': dias_aula,
                'dias_empresa': dias_empresa
            })
        
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
        
        exito = generar_parte_mensual(
            template_path=archivos_config['template_word'],
            output_path=archivos_config['output_word'],
            datos=datos_documento
        )
        
        return exito
            
    except Exception as e:
        print(f"ERROR: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    st.set_page_config(page_title="Cierre Mensual", layout="wide")
    st.title("Sistema de Cierre Mensual")
    
    tab1, tab2, tab3 = st.tabs(["Subir Archivos", "Procesar", "Descargar"])
    
    with tab1:
        st.header("Subir Archivos")
        
        col1, col2 = st.columns(2)
        
        with col1:
            excel_alumnos = st.file_uploader("Listado de Alumnos (Excel)", type=['xlsx', 'xls'], key="excel")
            pdf_becas = st.file_uploader("Becas y Ayudas (PDF)", type=['pdf'], key="becas")
            pdf_justificantes = st.file_uploader("Justificantes (PDF)", type=['pdf'], key="justificantes")
            template_word = st.file_uploader("Template Word", type=['docx'], key="template")
        
        with col2:
            pdfs_firmas = st.file_uploader(
                "PDFs de Firmas (multiples)",
                type=['pdf'],
                accept_multiple_files=True,
                key="firmas"
            )
            
            if pdfs_firmas:
                st.success(f"{len(pdfs_firmas)} archivos de firmas subidos")
        
        archivos_completos = all([excel_alumnos, pdf_becas, pdf_justificantes, template_word, pdfs_firmas])
        
        if archivos_completos:
            st.success("Todos los archivos cargados")
        else:
            st.warning("Sube todos los archivos necesarios")
    
    with tab2:
        st.header("Procesar Cierre Mensual")
        
        if not archivos_completos:
            st.error("Sube todos los archivos primero")
        else:
            if st.button("Iniciar Procesamiento", type="primary", use_container_width=True):
                with st.spinner("Procesando..."):
                    try:
                        temp_dir = tempfile.mkdtemp()
                        
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
                        
                        archivos_config = {
                            'excel_alumnos': excel_path,
                            'pdf_becas': becas_path,
                            'pdfs_firmas': firmas_paths,
                            'pdf_justificantes': justif_path,
                            'template_word': template_path,
                            'output_word': output_path
                        }
                        
                        old_stdout = sys.stdout
                        sys.stdout = log_buffer = StringIO()
                        
                        exito = procesar_cierre_mensual(archivos_config)
                        
                        sys.stdout = old_stdout
                        logs = log_buffer.getvalue()
                        
                        with st.expander("Ver logs", expanded=False):
                            st.code(logs)
                        
                        if exito and os.path.exists(output_path):
                            st.success("Procesamiento completado")
                            
                            with open(output_path, 'rb') as f:
                                st.session_state['documento_generado'] = f.read()
                            st.session_state['nombre_documento'] = 'parte_mensual.docx'
                            
                            st.info("Ve a la pestana Descargar")
                        else:
                            st.error("Error durante el procesamiento")
                            
                    except Exception as e:
                        st.error(f"Error: {e}")
                        import traceback
                        st.code(traceback.format_exc())
    
    with tab3:
        st.header("Descargar Documento")
        
        if 'documento_generado' in st.session_state:
            st.success("Documento listo")
            
            st.download_button(
                label="Descargar Parte Mensual",
                data=st.session_state['documento_generado'],
                file_name=st.session_state['nombre_documento'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            
            if st.button("Procesar Nuevo Mes", use_container_width=True):
                del st.session_state['documento_generado']
                del st.session_state['nombre_documento']
                st.rerun()
        else:
            st.info("No hay documento generado todavia")

if __name__ == "__main__":
    main()