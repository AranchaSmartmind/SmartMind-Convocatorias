"""
Helper para generar ZIP de múltiples alumnos
Cada alumno puede tener 1 o más documentos dependiendo de sus módulos
"""
import io
import zipfile
from typing import List, Dict


def generar_zip_todos_alumnos(generador, lista_datos_alumnos: List[Dict]) -> bytes:
    """
    Genera un ZIP con informes de todos los alumnos
    
    Args:
        generador: Instancia de WordGeneratorSEPE
        lista_datos_alumnos: Lista de datos de alumnos, cada uno con formato:
            {
                'alumno': {'nombre': '...', 'dni': '...', 'modulos': [...]},
                'curso': {'nombre': '...', 'codigo': '...'}
            }
    
    Returns:
        bytes: ZIP con todos los informes
    """
    
    print(f"\n{'='*80}")
    print(f"📦 GENERANDO ZIP CON {len(lista_datos_alumnos)} ALUMNOS")
    print(f"{'='*80}\n")
    
    zip_buffer = io.BytesIO()
    
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_principal:
        
        for idx, datos_alumno in enumerate(lista_datos_alumnos):
            alumno = datos_alumno.get('alumno', {})
            nombre = alumno.get('nombre', f'Alumno_{idx+1}')
            num_modulos = len(alumno.get('modulos', []))
            
            print(f"\n[{idx+1}/{len(lista_datos_alumnos)}] {nombre} ({num_modulos} módulos)")
            
            # Generar informe (puede ser .docx o .zip)
            resultado = generador.generar_informe_individual(datos_alumno)
            
            # Detectar si es ZIP o DOCX
            es_zip_interno = resultado[:2] == b'PK' and len(resultado) > 100
            
            if es_zip_interno:
                # Es un ZIP (alumno con >6 módulos)
                # Extraer sus archivos y agregarlos al ZIP principal
                print(f"  → Extrayendo múltiples documentos del alumno...")
                
                try:
                    with zipfile.ZipFile(io.BytesIO(resultado), 'r') as zip_alumno:
                        for nombre_archivo in zip_alumno.namelist():
                            contenido = zip_alumno.read(nombre_archivo)
                            zip_principal.writestr(nombre_archivo, contenido)
                            print(f"    ✓ {nombre_archivo}")
                except Exception as e:
                    print(f"    ✗ Error extrayendo ZIP del alumno: {e}")
                    # Agregar como está
                    nombre_archivo = f"{nombre.replace(' ', '_')[:50]}.zip"
                    zip_principal.writestr(nombre_archivo, resultado)
            else:
                # Es un DOCX simple (alumno con ≤6 módulos)
                nombre_archivo = f"{nombre.replace(' ', '_').replace(',', '')[:50]}.docx"
                zip_principal.writestr(nombre_archivo, resultado)
                print(f"  ✓ {nombre_archivo}")
    
    zip_buffer.seek(0)
    print(f"\n✅ ZIP PRINCIPAL GENERADO\n")
    return zip_buffer.getvalue()


# EJEMPLO DE USO EN TU CÓDIGO:
"""
# En desempleados.py o ocupados.py:

from .word_generator import WordGeneratorSEPE
from .word_generator_helper import generar_zip_todos_alumnos

# ... tu código para procesar Excel ...

# Preparar lista de datos
lista_datos = []
for alumno in alumnos:
    datos_alumno = {
        'alumno': {
            'nombre': alumno['nombre'],
            'dni': alumno['dni'],
            'modulos': alumno['modulos']
        },
        'curso': {
            'nombre': curso_nombre,
            'codigo': curso_codigo
        }
    }
    lista_datos.append(datos_alumno)

# Cargar plantilla
with open('plantilla_oficial.docx', 'rb') as f:
    plantilla_bytes = f.read()

# Generar ZIP con todos los alumnos
generador = WordGeneratorSEPE(plantilla_bytes)
zip_bytes = generar_zip_todos_alumnos(generador, lista_datos)

# Descargar
st.download_button(
    label="📥 Descargar todos los informes (ZIP)",
    data=zip_bytes,
    file_name=f"Informes_{curso_codigo}.zip",
    mime="application/zip"
)
"""