"""
Script para crear la estructura completa del proyecto
"""
import os

# Definir estructura
estructura = {
    "config": ["__init__.py", "settings.py"],
    "styles": ["__init__.py", "custom_styles.py"],
    "utils": ["__init__.py", "document_extractors.py", "excel_processors.py", "file_handlers.py"],
    "components": ["__init__.py", "header.py", "navigation.py"],
    "sections": ["__init__.py", "captacion.py", "formacion_inicio.py", "formacion_fin.py", "evaluacion.py", "cierre_mes.py"],
    "assets": []
}

def crear_estructura():
    """Crea toda la estructura de carpetas y archivos"""
    
    # Crear carpetas
    for carpeta in estructura.keys():
        os.makedirs(carpeta, exist_ok=True)
        print(f"✓ Carpeta '{carpeta}' creada/verificada")
    
    # Crear archivos
    for carpeta, archivos in estructura.items():
        for archivo in archivos:
            ruta = os.path.join(carpeta, archivo)
            if not os.path.exists(ruta):
                with open(ruta, 'w', encoding='utf-8') as f:
                    f.write(f'"""\n{archivo}\n"""\n')
                print(f"✓ Archivo '{ruta}' creado")
            else:
                print(f"  Archivo '{ruta}' ya existe")
    
    # Crear archivos raíz si no existen
    archivos_raiz = ["app.py", "requirements.txt", "README.md", ".gitignore"]
    for archivo in archivos_raiz:
        if not os.path.exists(archivo):
            with open(archivo, 'w', encoding='utf-8') as f:
                f.write(f'# {archivo}\n')
            print(f"✓ Archivo '{archivo}' creado")
    
    print("\n¡Estructura creada exitosamente!")

if __name__ == "__main__":
    crear_estructura()
    