#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script principal para generar Acta Grupal
Usa tu archivo word_generator_grupal.py existente
"""

import sys
import os

# Importar tu generador existente
from word_generator_grupal import WordGeneratorActaGrupal

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
        print("🎓 GENERADOR DE ACTA GRUPAL")
        print("="*60)
        
        # 1. Procesar Excel
        print(f"\n📂 Procesando Excel: {excel_path}")
        processor = ExcelProcessor(excel_path)
        datos = processor.cargar_datos()
        
        print(f"\n📊 Resumen:")
        print(f"   - Curso: {datos['curso_codigo']}")
        print(f"   - Alumnos: {len(datos['alumnos'])}")
        
        # 2. Cargar plantilla
        print(f"\n📄 Cargando plantilla: {plantilla_path}")
        with open(plantilla_path, 'rb') as f:
            plantilla_bytes = f.read()
        
        # 3. Generar acta usando TU clase
        print("\n🔄 Generando acta con word_generator_grupal.py...")
        generador = WordGeneratorActaGrupal(plantilla_bytes)
        acta_bytes = generador.generar_acta_grupal(datos)
        
        # 4. Guardar
        print(f"\n💾 Guardando en: {output_path}")
        with open(output_path, 'wb') as f:
            f.write(acta_bytes)
        
        print("\n" + "="*60)
        print("✅ ¡ACTA GENERADA EXITOSAMENTE!")
        print("="*60)
        print(f"\n📄 Archivo generado: {output_path}")
        print(f"📏 Tamaño: {len(acta_bytes) / 1024:.1f} KB")
        
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