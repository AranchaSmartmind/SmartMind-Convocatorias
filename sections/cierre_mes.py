"""
SECCIÓN CIERRE MES - Generador de Parte Mensual
Mantiene el mismo diseño que el resto de la aplicación
"""

import streamlit as st
import tempfile
import os
from pathlib import Path
import sys

def render_cierre_mes():
    """
    Renderiza la sección de Cierre Mes con el diseño de INTERPROS
    """
    
    # Contenedor principal con diseño holográfico
    with st.container():
        st.markdown("""
        <div style='
            background: rgba(20,25,35,0.8);
            border: 2px solid #4fc3f7;
            border-radius: 15px;
            padding: 2rem;
            margin-bottom: 2rem;
            box-shadow: 0 8px 32px rgba(79,195,247,0.3);
        '>
        """, unsafe_allow_html=True)
        
        # Título y descripción
        st.markdown("""
        <div style='text-align: center; margin-bottom: 2rem;'>
            <h2 style='font-family: "Orbitron", sans-serif; color: #4fc3f7; margin-bottom: 0.5rem;'>
                 GENERADOR DE PARTE MENSUAL
            </h2>
            <p style='font-size: 1.1rem; color: #e6f7ff;'>
                Sistema automático para generar partes mensuales de faltas y asistencia
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        # Upload de archivos
        st.markdown("###  SUBIR ARCHIVOS REQUERIDOS")
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            archivo_becas = st.file_uploader(
                " PDF de Otorgamiento de Becas",
                type=['pdf'],
                help="Archivo PDF con la información de becas y ayudas",
                key="becas_upload"
            )
            
            archivo_excel = st.file_uploader(
                " Excel CTRL de Alumnos", 
                type=['xlsx'],
                help="Archivo Excel con el control de alumnos",
                key="excel_upload"
            )
        
        with col2:
            archivo_parte_centro = st.file_uploader(
                " Parte de Firma del Centro",
                type=['pdf'],
                help="PDF con el parte de firma de asistencia del centro",
                key="parte_centro_upload"
            )
            
            archivo_justificantes = st.file_uploader(
                " Justificantes de Asistencia",
                type=['pdf'],
                help="PDF con justificantes médicos/laborales (opcional)",
                key="justificantes_upload"
            )
        
        st.markdown("---")
        
        # Información adicional
        with st.expander(" INFORMACIÓN SOBRE LOS ARCHIVOS"):
            st.markdown("""
            **Archivos necesarios:**
            -  **PDF de Otorgamiento de Becas**: Contiene información de becas y lista de alumnos
            -  **Excel CTRL de Alumnos**: Control de asistencia y datos de alumnos
            -  **Parte de Firma del Centro**: Registro de asistencia en el centro de formación
            
            **Archivo opcional:**
            -  **Justificantes**: Documentación médica/laboral para justificar faltas
            """)
        
        # Botón de procesamiento
        col_btn1, col_btn2, col_btn3 = st.columns([1,2,1])
        
        with col_btn2:
            if st.button(" GENERAR PARTE MENSUAL", use_container_width=True, key="generar_parte"):
                procesar_cierre_mensual(archivo_becas, archivo_excel, archivo_parte_centro, archivo_justificantes)
        
        st.markdown("</div>", unsafe_allow_html=True)

def procesar_cierre_mensual(archivo_becas, archivo_excel, archivo_parte_centro, archivo_justificantes):
    """
    Procesa los archivos y genera el parte mensual
    """
    
    # Validar archivos requeridos
    if not archivo_becas or not archivo_excel or not archivo_parte_centro:
        st.error(" Se requieren los archivos: PDF Becas, Excel Alumnos y Parte Centro")
        return
    
    with st.spinner(" Procesando archivos y generando parte mensual..."):
        try:
            # Importar el procesador desde el mismo directorio
            from sections.cierre_mensual import ProcesadorCierreMensual  # ← DIRECTAMENTE desde sections
            
            # Crear procesador
            procesador = ProcesadorCierreMensual()
            
            # Preparar archivos
            archivos = {
                'becas': archivo_becas,
                'excel': archivo_excel,
                'parte_centro': archivo_parte_centro,
                'justificantes': archivo_justificantes
            }
            
            # Ruta de plantilla (ajusta según tu estructura)
            plantilla_path = "templates/PARTEMENSUAL_template.docx"
            
            # Verificar que existe la plantilla
            if not os.path.exists(plantilla_path):
                st.error(f" No se encuentra la plantilla en: {plantilla_path}")
                st.info(" Crea una carpeta 'templates' en la raíz de tu proyecto y pon ahí el archivo PARTEMENSUAL_template.docx")
                return
            
            # Procesar
            exito, archivo_generado, estadisticas = procesador.procesar_completo(archivos, plantilla_path)
            
            if exito and archivo_generado:
                # Mostrar estadísticas
                st.success(" Parte mensual generado exitosamente!")
                
                # Mostrar métricas
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total Alumnos", estadisticas.get('total_alumnos', 0))
                with col2:
                    st.metric("Alumnos con Faltas", estadisticas.get('alumnos_con_faltas', 0))
                with col3:
                    st.metric("Mes", estadisticas.get('mes', 'No detectado'))
                
                # Información adicional
                with st.expander(" DETALLES DEL PROCESAMIENTO", expanded=True):
                    st.write(f"**Curso:** {estadisticas.get('curso', 'No detectado')}")
                    st.write(f"**Código:** {estadisticas.get('codigo', 'No detectado')}")
                    st.write(f"**Centro:** {estadisticas.get('centro', 'INTERPROS NEXT GENERATION SLU')}")
                    st.write(f"**Días lectivos:** {estadisticas.get('dias_lectivos', '2')}")
                
                # Botón de descarga
                with open(archivo_generado, "rb") as file:
                    nombre_archivo = f"PARTE_MENSUAL_{estadisticas.get('curso', 'CURSO').replace(' ', '_')}_{estadisticas.get('mes', 'MES')}.docx"
                    
                    st.download_button(
                        label=" DESCARGAR PARTE MENSUAL",
                        data=file,
                        file_name=nombre_archivo,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True,
                        key="descargar_parte"
                    )
                
                # Limpiar archivo temporal
                try:
                    os.unlink(archivo_generado)
                except:
                    pass
                
            else:
                error_msg = estadisticas.get('error', 'Error desconocido al generar el documento')
                st.error(f" Error al generar el parte: {error_msg}")
                
        except ImportError as e:
            st.error(f" Error: No se pudo importar el procesador de cierre mensual")
            st.error(f"Detalle: {e}")
            st.info(" Asegúrate de que el archivo 'cierre_mensual.py' está en la carpeta 'sections/'")
            
        except Exception as e:
            st.error(f" Error inesperado en el procesamiento: {str(e)}")
            st.info(" **Solución:** Verifica que todos los archivos subidos sean los correctos y no estén corruptos.")