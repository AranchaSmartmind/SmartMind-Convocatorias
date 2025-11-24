"""
SCRIPT DE DIAGNÓSTICO
=====================
Ejecuta este script en tu plataforma para ver qué está pasando
"""
import sys
import zipfile
import re
import io

print("="*60)
print(" DIAGNÓSTICO DEL GENERADOR")
print("="*60)

try:
    from excel_processor_grupal import ExcelProcessor
    print(" excel_processor importado correctamente")
except Exception as e:
    print(f" Error importando excel_processor: {e}")
    sys.exit(1)

try:
    from word_generator_grupal import WordGeneratorMultipaginaDuplicaTodo
    print(" word_generator_duplica_todo_AUTONOMO importado correctamente")
except Exception as e:
    print(f" Error importando generador: {e}")
    sys.exit(1)

print("\n" + "="*60)
print(" PASO 1: Cargar datos del Excel")
print("="*60)

ARCHIVO_EXCEL = 'PTEREvisar_2024_1339_CTRL_Tareas_AREA.xlsx'

try:
    processor = ExcelProcessor(ARCHIVO_EXCEL)
    datos = processor.cargar_datos()
    print(f" Datos cargados: {datos['total_alumnos']} alumnos")
except Exception as e:
    print(f" Error cargando Excel: {e}")
    sys.exit(1)

print("\n" + "="*60)
print(" PASO 2: Cargar plantilla")
print("="*60)

PLANTILLA = 'plantilla_grupal_oficial.docx'

try:
    with open(PLANTILLA, 'rb') as f:
        plantilla_bytes = f.read()
    print(f" Plantilla cargada: {len(plantilla_bytes)} bytes")
except Exception as e:
    print(f" Error cargando plantilla: {e}")
    sys.exit(1)

print("\n" + "="*60)
print(" PASO 3: Generar documento")
print("="*60)

try:
    generador = WordGeneratorMultipaginaDuplicaTodo(plantilla_bytes)
    print(" Generador creado")
    
    alumnos = datos['alumnos']
    total_alumnos = len(alumnos)
    num_actas = (total_alumnos + 14) // 15
    
    print(f"\n Configuración:")
    print(f"   Total alumnos: {total_alumnos}")
    print(f"   Actas a generar: {num_actas}")
    print(f"   Acta 1: Alumnos 1-15")
    if num_actas > 1:
        print(f"   Acta 2: Alumnos 16-{total_alumnos}")
    
    print("\n Generando...")
    acta_bytes = generador.generar_acta_grupal(datos)
    print(f" Documento generado: {len(acta_bytes)} bytes")
    
except Exception as e:
    print(f" Error generando documento: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "="*60)
print("🔍 PASO 4: Verificar contenido del documento")
print("="*60)

try:
    with zipfile.ZipFile(io.BytesIO(acta_bytes), 'r') as z:
        xml = z.read('word/document.xml').decode('utf-8')
    
    dnis = re.findall(r'<w:t[^>]*>(\d{8}[A-Z]|[XYZ]\d{7}[A-Z])</w:t>', xml)
    
    print(f"\n DNIs encontrados en el XML: {len(dnis)}")
    print(f"   Primeros 3: {dnis[:3]}")
    if len(dnis) > 15:
        print(f"   Últimos 3: {dnis[-3:]}")
    
    nombres = re.findall(r'<w:t[^>]*>([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+,\s+[A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ\s]+)</w:t>', xml)
    print(f"\n Nombres encontrados: {len(nombres)}")
    
    saltos = len(re.findall(r'<w:br\s+w:type="page"', xml))
    print(f"\n Saltos de página: {saltos}")
    
    campos = len(re.findall(r'<w:fldChar\s+w:fldCharType="begin"', xml))
    print(f" Campos de formulario: {campos}")
    
    print("\n Verificación de alumnos clave:")
    
    alumno_1 = '58459759S'
    alumno_16 = 'Z2374947H'
    
    if alumno_1 in xml:
        print(f"   Alumno 1 ({alumno_1}) presente")
    else:
        print(f"   Alumno 1 ({alumno_1}) NO encontrado")
    
    if alumno_16 in xml:
        print(f"   Alumno 16 ({alumno_16}) presente")
    else:
        print(f"   Alumno 16 ({alumno_16}) NO encontrado")
    
except Exception as e:
    print(f" Error verificando contenido: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)

print("\n" + "="*60)
print(" PASO 5: Guardar documento")
print("="*60)

ARCHIVO_SALIDA = 'ACTA_DIAGNOSTICO.docx'

try:
    with open(ARCHIVO_SALIDA, 'wb') as f:
        f.write(acta_bytes)
    print(f" Guardado: {ARCHIVO_SALIDA}")
except Exception as e:
    print(f" Error guardando: {e}")
    sys.exit(1)

print("\n" + "="*60)
print(" RESUMEN")
print("="*60)

if len(dnis) == total_alumnos:
    print(f"\n ¡TODO CORRECTO!")
    print(f"   El documento tiene los {total_alumnos} alumnos")
    print(f"   Si no los ves en Word, actualiza los campos:")
    print(f"   1. Abre {ARCHIVO_SALIDA}")
    print(f"   2. Ctrl+A (seleccionar todo)")
    print(f"   3. F9 (actualizar campos)")
elif len(dnis) == 15:
    print(f"\n PROBLEMA DETECTADO")
    print(f"   Solo se generó 1 acta (15 alumnos)")
    print(f"   Deberían ser {num_actas} actas ({total_alumnos} alumnos)")
    print(f"\n Posibles causas:")
    print(f"   - La función _combinar_actas_completas no está funcionando")
    print(f"   - Solo se está generando la primera acta")
else:
    print(f"\n Resultado inesperado: {len(dnis)} DNIs")

print("\n" + "="*60)
print("FIN DEL DIAGNÓSTICO")
print("="*60)