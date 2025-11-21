print("="*60)
print("TEST SIMPLE")
print("="*60)

# Ver archivos en la carpeta
import os
print("\nArchivos en esta carpeta:")
archivos = [f for f in os.listdir('.') if f.endswith('.py')]
for archivo in archivos:
    print(f"  - {archivo}")

# Intentar importar
print("\nIntentando importar módulos...")

try:
    from excel_processor_grupal import ExcelProcessor
    print("✅ excel_processor_grupal importado")
except Exception as e:
    print(f"❌ Error: {e}")

try:
    from word_generator_grupal import WordGeneratorMultipaginaDuplicaTodo
    print("✅ word_generator_grupal importado")
except Exception as e:
    print(f"❌ Error: {e}")

# Si funciona, probar cargar datos
try:
    processor = ExcelProcessor('PTEREvisar_2024_1339_CTRL_Tareas_AREA.xlsx')
    datos = processor.cargar_datos()
    print(f"\n✅ Datos cargados: {datos['total_alumnos']} alumnos")
except Exception as e:
    print(f"\n❌ Error cargando datos: {e}")