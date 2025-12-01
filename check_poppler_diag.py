# Script para verificar si pdf2image puede encontrar poppler

import sys
import os

print("="*80)
print("DIAGNÓSTICO DE POPPLER")
print("="*80)

print(f"\nPython executable: {sys.executable}")

print("\n1. Verificando PATH...")
path_dirs = os.environ.get("PATH", "").split(os.pathsep)
poppler_found_in_path = False
for p in path_dirs:
    if "poppler" in p.lower():
        print(f"  ✅ Poppler en PATH: {p}")
        poppler_found_in_path = True
        # Verificar si pdftoppm.exe existe
        pdftoppm_path = os.path.join(p, "pdftoppm.exe")
        if os.path.exists(pdftoppm_path):
            print(f"     ✅ pdftoppm.exe existe")
        else:
            print(f"     ❌ pdftoppm.exe NO encontrado")

if not poppler_found_in_path:
    print("  ❌ Poppler NO está en el PATH")

print("\n2. Intentando importar pdf2image...")
try:
    from pdf2image import convert_from_path
    print("  ✅ pdf2image importado correctamente")
    
    print("\n3. Verificando poppler_path...")
    # Intentar con poppler_path explícito
    poppler_paths_to_try = [
        r"C:\Users\Arancha\Desktop\Arancha\poppler-24.02.0\Library\bin",
        r"C:\poppler\Library\bin",
    ]
    
    for pp in poppler_paths_to_try:
        if os.path.exists(pp):
            print(f"  ✅ Ruta existe: {pp}")
            pdftoppm = os.path.join(pp, "pdftoppm.exe")
            if os.path.exists(pdftoppm):
                print(f"     ✅ pdftoppm.exe encontrado")
            else:
                print(f"     ❌ pdftoppm.exe NO encontrado")
        else:
            print(f"  ❌ Ruta NO existe: {pp}")
    
except ImportError as e:
    print(f"  ❌ No se pudo importar pdf2image: {e}")
    print("     Instala con: pip install pdf2image")

print("\n" + "="*80)