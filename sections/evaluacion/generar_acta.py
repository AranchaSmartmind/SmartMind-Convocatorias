#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script principal para generar Acta Grupal MULTIPÁGINA
"""

import sys
import os

# Importar el generador MULTIPÁGINA
from word_generator_grupal import WordGeneratorMultipaginaDuplicaTodo

# Importar el procesador de Excel
from excel_processor_grupal import ExcelProcessor


def main():
    """Función principal"""
    
    if len(sys.argv) != 4:
        print("❌ Uso incorrecto\n")
        print("Uso:")
        print("  python generar_acta.py <excel> <plantilla.docx> <salida.docx>\n")
        print("Ejemplo:")
        print("  python generar_acta.py PTEREvisar_2024_1339_CTRL_Tareas_AREA.xlsx plantilla_grupal_oficial.docx acta_final.docx")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    plantilla_path = sys.argv[2]
    output_path = sys.argv[3]
    
    # Verificar que existen los archivos
    if not os.path.exists(excel_path):
        print(f"❌ Error: No se encuentra el archivo Excel: {excel_path}")
        sys.exit(1)
    
    if not os.path.exists(plantilla_path):
        print(f"❌ Error: No se encuentra la plantilla Word: {plantilla_path}")
        sys.exit(1)
    
    try:
        print("="*60)
        print("🎓 GENERADOR DE ACTA GRUPAL MULTIPÁGINA")
        print("="*60)
        
        # 1. Procesar Excel
        print(f"\n📂 Procesando Excel: {excel_path}")
        processor = ExcelProcessor(excel_path)
        datos = processor.cargar_datos()
        
        total_alumnos = len(datos['alumnos'])
        num_paginas = (total_alumnos + 14) // 15
        
        print(f"\n📊 Resumen:")
        print(f"   - Curso: {datos['curso_codigo']}")
        print(f"   - Alumnos: {total_alumnos}")
        print(f"   - Páginas a generar: {num_paginas}")
        
        # 2. Cargar plantilla
        print(f"\n📄 Cargando plantilla: {plantilla_path}")
        with open(plantilla_path, 'rb') as f:
            plantilla_bytes = f.read()
        
        # 3. Usar el generador MULTIPÁGINA
        print("\n🔄 Generando acta multipágina...")
        generador = WordGeneratorMultipaginaDuplicaTodo(plantilla_bytes)
        acta_bytes = generador.generar_acta_grupal(datos)
        
        # 4. Guardar
        print(f"\n💾 Guardando en: {output_path}")
        with open(output_path, 'wb') as f:
            f.write(acta_bytes)
        
        # 5. Verificar que se generó correctamente
        import zipfile
        import re
        import io
        
        with zipfile.ZipFile(io.BytesIO(acta_bytes), 'r') as z:
            xml = z.read('word/document.xml').decode('utf-8')
        
        dnis = re.findall(r'<w:t[^>]*>(\d{8}[A-Z]|[XYZ]\d{7}[A-Z])</w:t>', xml)
        
        print("\n" + "="*60)
        print("✅ ¡ACTA GENERADA EXITOSAMENTE!")
        print("="*60)
        print(f"\n📄 Archivo generado: {output_path}")
        print(f"📏 Tamaño: {len(acta_bytes) / 1024:.1f} KB")
        print(f"👥 Alumnos en el documento: {len(dnis)}")
        
        if len(dnis) == total_alumnos:
            print(f"\n✅ Todos los {total_alumnos} alumnos están incluidos")
            print(f"✅ El documento tiene {num_paginas} páginas de alumnos")
        else:
            print(f"\n⚠️ ADVERTENCIA: Se esperaban {total_alumnos} alumnos pero se encontraron {len(dnis)}")
        
        print(f"\n💡 Si al abrir en Word no ves todos los alumnos:")
        print(f"   1. Presiona Ctrl+A (seleccionar todo)")
        print(f"   2. Presiona F9 (actualizar campos)")
        
    except FileNotFoundError as e:
        print(f"\n❌ Error: Archivo no encontrado")
        print(f"   {str(e)}")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error durante la generación:")
        print(f"   {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()