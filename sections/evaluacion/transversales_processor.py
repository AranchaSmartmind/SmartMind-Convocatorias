"""
PROCESADOR DE DATOS TRANSVERSALES (FCOO03)
===========================================
Extrae datos de archivos Excel para generar actas transversales
"""
import pandas as pd
from io import BytesIO
from typing import Dict, List
import re


class TransversalesProcessor:
    """
    Procesa archivos Excel de transversales (FCOO03)
    Extrae datos del curso y alumnos
    """
    
    def __init__(self):
        pass
    
    def cargar_datos(self, control_bytes: bytes, cronograma_bytes: bytes) -> Dict:
        """
        Carga y procesa datos de transversales desde los archivos Excel
        
        Args:
            control_bytes: Bytes del archivo CTRL_Tareas_AREA
            cronograma_bytes: Bytes del archivo Cronograma
        
        Returns:
            Dict con todos los datos procesados
        """
        
        try:
            print("\n📊 Procesando archivos de transversales...")
            
            # Leer pestaña ASISTENCIA para obtener datos del curso
            df_asist = pd.read_excel(BytesIO(control_bytes), sheet_name='ASISTENCIA', header=None)
            
            # Extraer datos del curso de las primeras filas
            curso_codigo = self._extraer_valor(df_asist, 0, 'curso:')
            curso_nombre = self._extraer_valor(df_asist, 1, 'nombre:')
            convocatoria = self._extraer_valor(df_asist, 2, 'convocatoria:')
            
            print(f"  📋 Curso: {curso_codigo}")
            print(f"  📝 Nombre: {curso_nombre}")
            
            # Buscar fila con "FCOO03" para obtener fechas
            fila_fcoo03 = None
            for idx, row in df_asist.iterrows():
                if any('FCOO03' in str(val) for val in row if pd.notna(val)):
                    fila_fcoo03 = idx
                    break
            
            fechas = ''
            horas_fcoo03 = ''
            if fila_fcoo03 is not None:
                celda_fcoo03 = df_asist.iloc[fila_fcoo03, 5]  # Columna donde está FCOO03
                if pd.notna(celda_fcoo03):
                    texto = str(celda_fcoo03)
                    # Extraer fechas del formato "FCOO03 \nFechas: 20/02/2025 - 27/06/2025"
                    match_fechas = re.search(r'Fechas:\s*(\d{2}/\d{2}/\d{4})\s*-\s*(\d{2}/\d{2}/\d{4})', texto)
                    if match_fechas:
                        fecha_inicio = match_fechas.group(1)
                        fecha_fin = match_fechas.group(2)
                        fechas = f"{fecha_inicio} - {fecha_fin}"
                
                # Horas en la siguiente columna
                horas_val = df_asist.iloc[fila_fcoo03, 6]
                if pd.notna(horas_val):
                    horas_fcoo03 = str(horas_val)
            
            print(f"  📅 Fechas FCOO03: {fechas}")
            print(f"  ⏱️ Horas: {horas_fcoo03}")
            
            # Buscar fila de encabezados (donde está "ALUMNO", "DNI", etc.)
            fila_encabezado = None
            for idx, row in df_asist.iterrows():
                if any('ALUMNO' in str(val).upper() for val in row if pd.notna(val)):
                    fila_encabezado = idx
                    break
            
            if fila_encabezado is None:
                raise Exception("No se encontró la fila de encabezados con 'ALUMNO'")
            
            print(f"  📌 Encabezado en fila: {fila_encabezado}")
            
            # Leer datos de alumnos desde la fila después del encabezado
            df_alumnos = pd.read_excel(
                BytesIO(control_bytes), 
                sheet_name='ASISTENCIA', 
                header=fila_encabezado,
                skiprows=0
            )
            
            # Procesar alumnos
            alumnos = []
            for idx, row in df_alumnos.iterrows():
                # Saltar filas vacías o sin DNI
                if pd.isna(row.get('DNI')) or row.get('DNI') == '':
                    continue
                
                nombre = str(row.get('ALUMNO', '')).strip()
                dni = str(row.get('DNI', '')).strip()
                
                # Buscar columna de horas FCOO03 (la columna 5 o 6 según estructura)
                horas_actividades = horas_fcoo03
                
                # Determinar calificación (si hay datos de baja, es NO APTO)
                baja = row.get('BAJA MOTIVOS', '')
                if pd.notna(baja) and str(baja).strip() != '':
                    calificacion = 'NO APTO'
                else:
                    # Buscar porcentaje de asistencia
                    porcentaje_col = None
                    for col in df_alumnos.columns:
                        if '%' in str(col):
                            porcentaje_col = col
                            break
                    
                    if porcentaje_col and pd.notna(row.get(porcentaje_col)):
                        porcentaje = float(row.get(porcentaje_col, 0))
                        if porcentaje >= 0.75:  # 75% o más
                            calificacion = 'APTO'
                        else:
                            calificacion = 'NO APTO'
                    else:
                        calificacion = 'APTO'  # Por defecto
                
                if nombre and dni:
                    alumnos.append({
                        'numero': len(alumnos) + 1,
                        'dni': dni,
                        'nombre': nombre,
                        'horas_actividades': horas_actividades,
                        'calificacion_final': calificacion
                    })
            
            print(f"  👥 Alumnos procesados: {len(alumnos)}")
            
            # Extraer fechas del cronograma si es necesario
            fecha_inicio = ''
            fecha_fin = ''
            if fechas:
                partes = fechas.split(' - ')
                if len(partes) == 2:
                    fecha_inicio = partes[0]
                    fecha_fin = partes[1]
            
            # Construir diccionario de salida
            datos = {
                'campo_1_convocatoria': convocatoria or '',
                'campo_2_accion': curso_codigo or '',
                'campo_3_especialidad': curso_nombre or '',
                'campo_4_codigo': 'FCOO03',
                'campo_5_centro': 'INTERPROS NEXT GENERATION S.L.U.',
                'campo_6_duracion': horas_fcoo03 or '10',
                'campo_7_actividades': horas_fcoo03 or '10',
                'campo_8_modalidad': 'PRESENCIAL',
                'campo_9_fecha_inicio': fecha_inicio,
                'campo_10_fecha_fin': fecha_fin,
                'alumnos': alumnos,
                'total_alumnos': len(alumnos)
            }
            
            print(f"✅ Datos procesados correctamente\n")
            
            return datos
            
        except Exception as e:
            print(f"❌ Error procesando Excel transversales: {str(e)}")
            raise Exception(f"Error procesando Excel transversales: {str(e)}")
    
    def _extraer_valor(self, df: pd.DataFrame, fila: int, clave: str) -> str:
        """
        Extrae un valor de una fila que tiene formato 'clave: valor'
        
        Args:
            df: DataFrame
            fila: Número de fila
            clave: Texto clave a buscar
        
        Returns:
            Valor extraído como string
        """
        try:
            row = df.iloc[fila]
            for val in row:
                if pd.notna(val) and clave in str(val).lower():
                    # El valor está en la siguiente columna
                    idx = row.tolist().index(val)
                    if idx + 1 < len(row):
                        valor = row.iloc[idx + 1]
                        if pd.notna(valor):
                            return str(valor).strip()
            return ''
        except Exception as e:
            print(f"  ⚠ Error extrayendo {clave}: {e}")
            return ''


if __name__ == "__main__":
    # Prueba del procesador
    import sys
    sys.path.insert(0, '/mnt/user-data/uploads')
    
    print("\n🧪 PRUEBA DEL PROCESADOR TRANSVERSALES")
    print("=" * 80)
    
    with open('PTEREvisar_2024_1334_CTRL_Tareas_AREA_v001__1_.xlsx', 'rb') as f:
        control_bytes = f.read()
    
    # Usar el mismo archivo como cronograma por ahora
    cronograma_bytes = control_bytes
    
    processor = TransversalesProcessor()
    datos = processor.cargar_datos(control_bytes, cronograma_bytes)
    
    print(f"\n📊 RESULTADO:")
    print(f"  Curso: {datos['campo_2_accion']}")
    print(f"  Especialidad: {datos['campo_3_especialidad']}")
    print(f"  Total alumnos: {datos['total_alumnos']}")
    print(f"\n  Primeros 3 alumnos:")
    for alumno in datos['alumnos'][:3]:
        print(f"    {alumno['numero']}. {alumno['nombre']} - {alumno['calificacion_final']}")