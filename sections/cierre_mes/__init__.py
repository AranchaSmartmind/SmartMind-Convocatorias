"""
Módulo Cierre de Mes
Generador de Parte Mensual de Asistencia
"""

import streamlit as st
import pandas as pd
import os
import tempfile
from datetime import datetime

# Importar módulos propios
from sections.cierre_mes.utils.extractor_ocr import calcular_dias_lectivos_y_asistencias
from sections.cierre_mes.utils.procesador_datos import (
    extraer_datos_curso_pdf,
    extraer_alumnos_excel,
    extraer_becas_ayudas_simple,
    extraer_justificantes,
    obtener_mes_anterior
)
from sections.cierre_mes.utils.generador_documento import rellenar_template, construir_observaciones
from sections.cierre_mes.utils.helpers import buscar_coincidencia

def render_cierre_mes():
    """Renderiza la sección de Cierre de Mes"""
    
    # Inicializar session state
    if 'cierre_processed' not in st.session_state:
        st.session_state.cierre_processed = False
    if 'cierre_resultado' not in st.session_state:
        st.session_state.cierre_resultado = None
    
    # Descripción
    st.markdown("""
    Sistema automatizado para generar partes mensuales de asistencia procesando 
    documentos con OCR (Reconocimiento Óptico de Caracteres).
    """)
    
    st.divider()
    
    # Layout con columnas
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Documentos Requeridos")
        
        st.markdown("##### 1. Información del Curso")
        becas_file = st.file_uploader(
            "PDF de Otorgamiento de Becas",
            type=['pdf'],
            help="PDF con información del curso y becas otorgadas",
            key="cierre_becas"
        )
        
        excel_file = st.file_uploader(
            "Excel de Control de Alumnos",
            type=['xlsx', 'xls'],
            help="Excel con el listado de alumnos",
            key="cierre_excel"
        )
        
        st.markdown("##### 2. Hojas de Asistencia")
        firmas_files = st.file_uploader(
            "PDFs de Hojas de Firmas",
            type=['pdf'],
            accept_multiple_files=True,
            help="PDFs de firmas (aula y empresa)",
            key="cierre_firmas"
        )
        
        st.markdown("##### 3. Justificantes (Opcional)")
        justificantes_file = st.file_uploader(
            "PDF de Justificantes",
            type=['pdf'],
            help="PDF con justificantes",
            key="cierre_just"
        )
        
        st.divider()
        
        # Botón de procesamiento
        procesar_btn = st.button(
            "Generar Parte Mensual",
            type="primary",
            use_container_width=True,
            disabled=not (becas_file and excel_file and firmas_files),
            key="cierre_procesar"
        )
    
    with col2:
        if not (becas_file and excel_file and firmas_files):
            st.info("Sube todos los archivos obligatorios en el panel izquierdo para comenzar.")
            
            with st.expander("Información sobre los archivos"):
                st.markdown("""
                **Archivos Obligatorios:**
                
                1. **PDF de Otorgamiento de Becas**
                   - Número de curso, especialidad, centro
                   - Becas y ayudas por alumno
                
                2. **Excel de Control de Alumnos**
                   - Listado completo con DNI
                   - Debe tener columna "APELLIDOS, NOMBRE"
                
                3. **PDFs de Hojas de Firmas**
                   - Registros de asistencia (aula y empresa)
                   - Se procesan con OCR automático
                   - Puedes subir múltiples archivos
                
                **Archivos Opcionales:**
                
                4. **PDF de Justificantes**
                   - Justificantes médicos o de empresa
                """)
        
        # Procesamiento
        if procesar_btn:
            with st.spinner("Procesando documentos..."):
                try:
                    # Crear directorio temporal
                    with tempfile.TemporaryDirectory() as tmpdir:
                        # Guardar archivos
                        becas_path = os.path.join(tmpdir, becas_file.name)
                        with open(becas_path, 'wb') as f:
                            f.write(becas_file.read())
                        
                        excel_path = os.path.join(tmpdir, excel_file.name)
                        with open(excel_path, 'wb') as f:
                            f.write(excel_file.read())
                        
                        firmas_paths = []
                        for firma_file in firmas_files:
                            firma_path = os.path.join(tmpdir, firma_file.name)
                            with open(firma_path, 'wb') as f:
                                f.write(firma_file.read())
                            firmas_paths.append(firma_path)
                        
                        just_path = None
                        if justificantes_file:
                            just_path = os.path.join(tmpdir, justificantes_file.name)
                            with open(just_path, 'wb') as f:
                                f.write(justificantes_file.read())
                        
                        # Procesar datos
                        progreso = st.progress(0, text="Extrayendo datos del curso...")
                        
                        # 1. Datos del curso
                        numero_curso, especialidad = extraer_datos_curso_pdf(becas_path)
                        centro = "INTERPROS NEXT GENERATION SLU"
                        mes = obtener_mes_anterior()
                        progreso.progress(15, text="Datos del curso extraídos")
                        
                        # 2. Procesar firmas con OCR
                        progreso.progress(20, text="Procesando hojas de asistencia...")
                        dias_lectivos, asistencias, faltas = calcular_dias_lectivos_y_asistencias(firmas_paths)
                        progreso.progress(50, text="Asistencias procesadas")
                        
                        # 3. Becas y ayudas
                        progreso.progress(55, text="Extrayendo becas...")
                        becas = extraer_becas_ayudas_simple(becas_path)
                        progreso.progress(65, text="Becas extraídas")
                        
                        # 4. Justificantes
                        progreso.progress(70, text="Extrayendo justificantes...")
                        justificantes = {}
                        if just_path:
                            justificantes = extraer_justificantes(just_path)
                        progreso.progress(75, text="Justificantes extraídos")
                        
                        # 5. Consolidar datos de alumnos
                        progreso.progress(80, text="Consolidando datos...")
                        alumnos_excel = extraer_alumnos_excel(excel_path)
                        
                        alumnos_final = []
                        for alumno in alumnos_excel:
                            nombre = alumno['nombre_completo']
                            
                            asist = buscar_coincidencia(nombre, asistencias) or {'dias_aula': 0, 'dias_empresa': 0}
                            ayudas = buscar_coincidencia(nombre, becas) or []
                            falta = buscar_coincidencia(nombre, faltas) or 0
                            just = buscar_coincidencia(nombre, justificantes) or 0
                            
                            obs = construir_observaciones(
                                ayudas,
                                asist['dias_aula'],
                                asist['dias_empresa'],
                                just,
                                dias_lectivos
                            )
                            
                            alumnos_final.append({
                                'nombre': nombre,
                                'dni': alumno['dni'],
                                'faltas': falta,
                                'observaciones': obs,
                                'dias_aula': asist['dias_aula'],
                                'dias_empresa': asist['dias_empresa']
                            })
                        
                        progreso.progress(90, text="Generando documento...")
                        
                        # 6. Generar documento
                        template_path = 'sections/cierre_mes/templates/TEMPLATE_Original.docx'
                        
                        print(f"DEBUG: Buscando template en: {template_path}")
                        
                        if not os.path.exists(template_path):
                            raise FileNotFoundError(f"No se encontró el template en: {template_path}")
                        
                        datos = {
                            'numero_curso': numero_curso,
                            'especialidad': especialidad,
                            'centro': centro,
                            'mes': mes,
                            'dias_lectivos': dias_lectivos,
                            'alumnos': alumnos_final
                        }
                        
                        print(f"DEBUG: Datos preparados. Alumnos: {len(alumnos_final)}")
                        progreso.progress(91, text="Abriendo template...")
                        
                        doc = rellenar_template(template_path, datos)
                        
                        if not doc:
                            raise Exception("No se pudo generar el documento desde el template")
                        
                        print("DEBUG: Documento generado, guardando...")
                        progreso.progress(95, text="Guardando documento...")
                        
                        output_path = os.path.join(tmpdir, f'Parte_Mensual_{mes}_2025.docx')
                        doc.save(output_path)
                        
                        print(f"DEBUG: Documento guardado en: {output_path}")
                        progreso.progress(98, text="Leyendo archivo...")
                        
                        # Esperar un momento para asegurar que el archivo se guardó
                        import time
                        time.sleep(0.2)
                        
                        with open(output_path, 'rb') as f:
                            doc_bytes = f.read()
                        
                        print(f"DEBUG: Archivo leído, tamaño: {len(doc_bytes)} bytes")
                        progreso.progress(100, text="Completado")
                        
                        st.session_state.cierre_resultado = {
                            'doc_bytes': doc_bytes,
                            'filename': f'Parte_Mensual_{mes}_2025.docx',
                            'datos': datos,
                            'alumnos': alumnos_final,
                            'dias_lectivos': dias_lectivos
                        }
                        st.session_state.cierre_processed = True
                        st.rerun()
                    
                except Exception as e:
                    st.error(f"Error durante el procesamiento: {str(e)}")
                    with st.expander("Ver detalles del error"):
                        import traceback
                        st.code(traceback.format_exc())
        
        # Mostrar resultados
        if st.session_state.cierre_processed and st.session_state.cierre_resultado:
            resultado = st.session_state.cierre_resultado
            
            st.success("Documento generado exitosamente")
            
            # Botón de descarga
            st.download_button(
                label="Descargar Documento",
                data=resultado['doc_bytes'],
                file_name=resultado['filename'],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary",
                key="cierre_download"
            )
            
            st.divider()
            
            # Estadísticas
            st.subheader("Resumen General")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                st.metric("Total Alumnos", len(resultado['alumnos']))
            
            with col2:
                st.metric("Días Lectivos", resultado['dias_lectivos'])
            
            with col3:
                total_faltas = sum(a['faltas'] for a in resultado['alumnos'])
                st.metric("Total Faltas", total_faltas)
            
            with col4:
                if resultado['dias_lectivos'] > 0:
                    promedio_asist = 100 * (1 - total_faltas / (len(resultado['alumnos']) * resultado['dias_lectivos']))
                    st.metric("Asistencia Promedio", f"{promedio_asist:.1f}%")
            
            st.divider()
            
            # Tabla de alumnos
            st.subheader("Detalle por Alumno")
            
            df = pd.DataFrame([
                {
                    'Alumno': a['nombre'],
                    'DNI': a['dni'],
                    'Días Aula': a['dias_aula'],
                    'Días Empresa': a['dias_empresa'],
                    'Faltas': a['faltas'],
                    'Observaciones': a['observaciones']
                }
                for a in resultado['alumnos']
            ])
            
            st.dataframe(df, use_container_width=True, hide_index=True)
            
            # Botón para nuevo parte
            st.divider()
            if st.button("Generar Nuevo Parte", use_container_width=True, key="cierre_nuevo"):
                st.session_state.cierre_processed = False
                st.session_state.cierre_resultado = None
                st.rerun()

__all__ = ['render_cierre_mes']