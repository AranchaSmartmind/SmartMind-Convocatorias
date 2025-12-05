"""
Script de prueba para verificar que la generación funciona con datos completos
"""
from sections.evaluacion.cierre_mes.generacion_word import generar_parte_mensual
import os

# Datos de prueba COMPLETOS
datos_prueba = {
    'numero_curso': '2024/1339',
    'especialidad': 'OPERACIONES AUXILIARES DE SERVICIOS ADMINISTRATIVOS Y GENERALES',
    'centro': 'INTERPROS NEXT GENERATION SLU',
    'mes': 'OCTUBRE',
    'dias_lectivos': 22,  # ← Valor real en lugar de 0
    'horas_empresa': 21,
    'horas_aula': 2,
    'alumnos': [
        {
            'nombre': 'BRAÑA MANCHADO, NURIA',
            'dni': '71879712C',
            'faltas': 2,  # ← Valor real en lugar de 0
            'observaciones': 'Beca de transporte'  # ← Observación real
        },
        {
            'nombre': 'BRIZUELA TILL, JOSE ALEJANDRO',
            'dni': 'Z2751442A',
            'faltas': 0,
            'observaciones': 'Empresa: 4 días'
        },
        {
            'nombre': 'CASADO VIÑAS, LAURA MARIA',
            'dni': '11422042N',
            'faltas': 1,
            'observaciones': '1 faltas justificadas'
        },
        {
            'nombre': 'DE SOUZA ROS, CATHERINE',
            'dni': '71906405X',
            'faltas': 0,
            'observaciones': 'Beca de comedor / Ayuda de transporte'
        },
        {
            'nombre': 'DELGADO ORTIZ, MARYELY YORSELY',
            'dni': 'Z2762262J',
            'faltas': 3,
            'observaciones': 'Empresa: 5 días / 2 faltas justificadas'
        }
    ]
}

# Rutas
template_path = r"C:\Users\Arancha\Desktop\Arancha\Repos\sections\evaluacion\cierre_mes\template_original.docx"
output_path = r"C:\Users\Arancha\Desktop\Arancha\Repos\PRUEBA_CON_DATOS_COMPLETOS.docx"

print("="*80)
print("PRUEBA CON DATOS COMPLETOS")
print("="*80)

if not os.path.exists(template_path):
    print(f"❌ Template no encontrado: {template_path}")
else:
    print(f"✅ Template encontrado")
    
    exito = generar_parte_mensual(template_path, output_path, datos_prueba)
    
    if exito and os.path.exists(output_path):
        print(f"\n✅ ¡ÉXITO! Documento generado en:")
        print(f"   {output_path}")
        print(f"\n📂 Abre el archivo para verificar:")
        print(f"   - Mes: OCTUBRE")
        print(f"   - Días lectivos: 22")
        print(f"   - Alumna 1 con 2 faltas y observación 'Beca de transporte'")
        print(f"   - Alumno 5 con 3 faltas y observación 'Empresa: 5 días / 2 faltas justificadas'")
    else:
        print("\n❌ Error al generar el documento")