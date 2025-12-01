"""
TEST DE IMPORT - Diagnostica problemas con render_cierre_mes
Guarda este archivo como: test_import.py
Ejecuta desde: C:\\Users\\Arancha\\Desktop\\Arancha\\Repos
"""

import sys
import os

def diagnosticar_import():
    print("="*80)
    print(" 🔍 DIAGNÓSTICO DE IMPORT")
    print("="*80)
    print()
    
    # 1. Verificar directorio
    directorio_actual = os.getcwd()
    print(f"📁 Directorio actual:")
    print(f"   {directorio_actual}\n")
    
    # 2. Verificar estructura
    print("📂 Verificando estructura de carpetas:")
    
    checks = [
        ("sections/", "Carpeta principal"),
        ("sections/__init__.py", "Archivo de inicialización"),
        ("sections/cierre_mes.py", "Módulo cierre_mes"),
        ("sections/evaluacion/", "Subcarpeta evaluacion (opcional)"),
        ("sections/evaluacion/__init__.py", "Init evaluacion (opcional)"),
    ]
    
    estructura_ok = True
    for ruta, descripcion in checks:
        existe = os.path.exists(ruta)
        simbolo = "✅" if existe else "❌"
        print(f"   {simbolo} {ruta:<35} {descripcion}")
        
        if "sections/__init__.py" in ruta and not existe:
            estructura_ok = False
            print(f"      💡 CREAR: type nul > {ruta}")
        
        if "sections/cierre_mes.py" in ruta and not existe:
            estructura_ok = False
            print(f"      ⚠️  Este archivo es NECESARIO")
    
    print()
    
    if not estructura_ok:
        print("❌ La estructura tiene problemas. Sigue las sugerencias arriba.\n")
        return False
    
    # 3. Verificar contenido del archivo
    print("📄 Verificando contenido de cierre_mes.py:")
    
    if os.path.exists("sections/cierre_mes.py"):
        try:
            with open("sections/cierre_mes.py", 'r', encoding='utf-8') as f:
                contenido = f.read()
                
            # Verificar función
            if 'def render_cierre_mes' in contenido:
                print("   ✅ Función render_cierre_mes() encontrada")
                
                # Verificar sintaxis
                try:
                    compile(contenido, 'sections/cierre_mes.py', 'exec')
                    print("   ✅ Sintaxis correcta")
                except SyntaxError as e:
                    print(f"   ❌ ERROR DE SINTAXIS:")
                    print(f"      Línea {e.lineno}: {e.msg}")
                    print(f"      {e.text}")
                    return False
                
            else:
                print("   ❌ Función render_cierre_mes() NO encontrada")
                print("      Verifica que la función esté definida correctamente")
                
                # Mostrar primeras líneas
                lineas = contenido.split('\n')[:30]
                print("\n   📝 Primeras 30 líneas del archivo:")
                for i, linea in enumerate(lineas, 1):
                    print(f"      {i:3}: {linea[:70]}")
                
                return False
                
        except Exception as e:
            print(f"   ❌ Error leyendo archivo: {e}")
            return False
    
    print()
    
    # 4. Intentar import
    print("🔧 Intentando importar:")
    
    try:
        from sections.cierre_mes import render_cierre_mes
        print("   ✅ Import exitoso!")
        print(f"   📍 Función: {render_cierre_mes}")
        print(f"   📍 Módulo: {render_cierre_mes.__module__}")
        print(f"   📍 Archivo: {render_cierre_mes.__code__.co_filename}")
        print()
        print("="*80)
        print(" ✅ TODO CORRECTO - El import funciona!")
        print("="*80)
        return True
        
    except ImportError as e:
        print(f"   ❌ Error de import: {e}")
        print()
        print("="*80)
        print(" 💡 POSIBLES SOLUCIONES:")
        print("="*80)
        print()
        print("1. Crear __init__.py:")
        print("   type nul > sections\\__init__.py")
        print()
        print("2. Verificar que render_cierre_mes esté definida correctamente")
        print()
        print("3. Limpiar caché:")
        print("   rmdir /s /q sections\\__pycache__")
        print()
        print("4. Verificar sintaxis:")
        print("   python -m py_compile sections\\cierre_mes.py")
        print()
        return False
        
    except Exception as e:
        print(f"   ❌ Error inesperado: {e}")
        import traceback
        traceback.print_exc()
        return False

def verificar_imports_del_modulo():
    """
    Verifica que los imports dentro de cierre_mes.py funcionen
    """
    print("\n" + "="*80)
    print(" 🔍 VERIFICANDO IMPORTS DENTRO DE cierre_mes.py")
    print("="*80)
    print()
    
    if not os.path.exists("sections/cierre_mes.py"):
        print("❌ Archivo no encontrado")
        return
    
    with open("sections/cierre_mes.py", 'r', encoding='utf-8') as f:
        contenido = f.read()
    
    # Buscar imports
    import re
    imports = re.findall(r'^(?:from|import)\s+(.+?)(?:\s+import|\n)', contenido, re.MULTILINE)
    
    if imports:
        print("📦 Imports encontrados:")
        for imp in imports[:10]:  # Primeros 10
            print(f"   • {imp.strip()}")
        
        print("\n💡 Si alguno falla, instálalo con:")
        print("   pip install <nombre_paquete>")
    else:
        print("ℹ️  No se encontraron imports (o el patrón no los capturó)")

def crear_archivos_faltantes():
    """
    Ofrece crear los archivos __init__.py faltantes
    """
    print("\n" + "="*80)
    print(" 🛠️  CREAR ARCHIVOS FALTANTES")
    print("="*80)
    print()
    
    archivos_a_crear = []
    
    if not os.path.exists("sections/__init__.py"):
        archivos_a_crear.append("sections/__init__.py")
    
    if os.path.exists("sections/evaluacion") and not os.path.exists("sections/evaluacion/__init__.py"):
        archivos_a_crear.append("sections/evaluacion/__init__.py")
    
    if archivos_a_crear:
        print("Los siguientes archivos __init__.py faltan:")
        for archivo in archivos_a_crear:
            print(f"   • {archivo}")
        
        respuesta = input("\n¿Quieres crearlos automáticamente? (s/n): ").lower()
        
        if respuesta == 's':
            for archivo in archivos_a_crear:
                try:
                    # Crear directorio si no existe
                    directorio = os.path.dirname(archivo)
                    if directorio and not os.path.exists(directorio):
                        os.makedirs(directorio)
                    
                    # Crear archivo vacío
                    with open(archivo, 'w') as f:
                        f.write("# Archivo de inicialización automático\n")
                    print(f"   ✅ Creado: {archivo}")
                except Exception as e:
                    print(f"   ❌ Error creando {archivo}: {e}")
            
            print("\n✨ Archivos creados. Intenta ejecutar tu app de nuevo.")
        else:
            print("\n💡 Créalos manualmente con:")
            for archivo in archivos_a_crear:
                print(f"   type nul > {archivo}")
    else:
        print("✅ Todos los archivos __init__.py necesarios existen")

if __name__ == "__main__":
    print("""
╔══════════════════════════════════════════════════════════════════════════════╗
║                                                                              ║
║                    TEST DE IMPORT - render_cierre_mes                        ║
║                                                                              ║
║  Este script diagnostica problemas con el import de render_cierre_mes        ║
║                                                                              ║
╚══════════════════════════════════════════════════════════════════════════════╝
    """)
    
    # Diagnóstico principal
    resultado = diagnosticar_import()
    
    if not resultado:
        # Si falló, ofrecer más ayuda
        verificar_imports_del_modulo()
        crear_archivos_faltantes()
        
        print("\n" + "="*80)
        print(" 📞 NECESITAS MÁS AYUDA?")
        print("="*80)
        print()
        print("Comparte esta información:")
        print("1. La salida completa de este script")
        print("2. El contenido de sections/cierre_mes.py (primeras 50 líneas)")
        print("3. Tu archivo principal (app.py o main.py)")
        print()
    
    input("\nPresiona ENTER para salir...")