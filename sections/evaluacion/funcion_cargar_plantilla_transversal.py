"""
FUNCIÓN PARA CARGAR PLANTILLA TRANSVERSAL POR DEFECTO
======================================================
Esta función se agregará a desempleados.py para cargar la plantilla automáticamente
"""
import os

def cargar_plantilla_transversal_por_defecto():
    """
    Carga la plantilla de acta transversal integrada en la aplicación
    
    Returns:
        bytes: Contenido de la plantilla o None si no se encuentra
    """
    try:
        # Ubicaciones posibles de la plantilla
        plantilla_transversal = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), 
            'plantilla_transversal_oficial.docx'
        )
        
        ubicaciones = [
            plantilla_transversal,
            os.path.join(os.getcwd(), 'sections', 'evaluacion', 'plantilla_transversal_oficial.docx'),
            os.path.join(os.path.dirname(__file__), '..', 'plantilla_transversal_oficial.docx'),
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
        print(f"❌ Error cargando plantilla transversal: {e}")
        return None


"""
INSTRUCCIONES PARA INSTALAR LA PLANTILLA:
==========================================

1. Copiar la plantilla al directorio de evaluación:
   
   cp plantilla_transversal_oficial.docx sections/evaluacion/

2. Agregar la función cargar_plantilla_transversal_por_defecto() 
   al archivo desempleados.py (después de las otras funciones de carga)

3. En la sección render_transversales(), usar:
   
   plantilla_bytes = cargar_plantilla_transversal_por_defecto()
   if plantilla_bytes:
       st.info("📝 Usando plantilla oficial transversal predeterminada")
   else:
       st.error("❌ No se encontró la plantilla")
       return
"""

# Verificar que la plantilla existe
if __name__ == "__main__":
    import sys
    sys.path.insert(0, '/mnt/user-data/outputs')
    
    # Simular carga
    plantilla_path = '/mnt/user-data/outputs/plantilla_transversal_oficial.docx'
    if os.path.exists(plantilla_path):
        with open(plantilla_path, 'rb') as f:
            contenido = f.read()
            print(f"✓ Plantilla encontrada: {len(contenido):,} bytes")
    else:
        print("❌ Plantilla NO encontrada")