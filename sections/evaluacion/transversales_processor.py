"""
PROCESADOR DE DATOS TRANSVERSALES
==================================
Extrae datos del Excel de control de tareas para generar actas transversales
"""
import pandas as pd
import io
from typing import Dict, List


class TransversalesProcessor:
    """Procesa datos de competencias transversales desde Excel"""
    
    def __init__(self):
        self.df_resumen = None
        self.df_calificaciones = None
    
    def cargar_datos(self, control_bytes: bytes, cronograma_bytes: bytes) -> Dict:
        """
        Carga datos del Excel de control y cronograma
        
        Args:
            control_bytes: Bytes del archivo Excel CTRL_Tareas_AREA.xlsx
            cronograma_bytes: Bytes del archivo Cronograma.xlsx
            
        Returns:
            Dict con datos procesados para el acta
        """
        try:
            self.df_resumen = pd.read_excel(
                io.BytesIO(control_bytes), 
                sheet_name='RESUMEN',
                header=0
            )
            
            self.df_calificaciones = pd.read_excel(
                io.BytesIO(control_bytes),
                sheet_name='CALIFICACIONES',
                header=7
            )
            
            campo_1_convocatoria = ''
            
            df_asistencia = pd.read_excel(
                io.BytesIO(control_bytes),
                sheet_name='ASISTENCIA',
                header=None
            )
            campo_2_accion = str(df_asistencia.iloc[0, 2]) if pd.notna(df_asistencia.iloc[0, 2]) else ''

            df_cronograma = pd.read_excel(
                io.BytesIO(cronograma_bytes),
                sheet_name='Cronograma',
                header=None
            )

            valor_fila20 = str(df_cronograma.iloc[19, 0])

            partes = valor_fila20.split(None, 1)
            if len(partes) >= 2:
                campo_4_codigo = partes[0]
                campo_3_especialidad = partes[1]
            else:
                campo_4_codigo = ''
                campo_3_especialidad = valor_fila20
            
            campo_5_centro = 'INTERPROS NEXT GENERATION S.L.U'

            campo_6_duracion = str(df_cronograma.iloc[19, 2]) if pd.notna(df_cronograma.iloc[19, 2]) else ''

            campo_7_actividades = '1'

            campo_8_modalidad = 'Presencial'
            for col in range(15):
                valor = df_cronograma.iloc[9, col]
                if pd.notna(valor) and 'Modalidad' in str(valor):
                    if 'Presencial' in str(valor):
                        campo_8_modalidad = 'Presencial'
                        break
            
            valor_fechas = str(df_cronograma.iloc[19, 8])
            
            if '-' in valor_fechas:
                partes_fechas = valor_fechas.split('-')
                if len(partes_fechas) == 2:
                    campo_9_fecha_inicio = partes_fechas[0].strip()
                    campo_10_fecha_fin = partes_fechas[1].strip()
                else:
                    campo_9_fecha_inicio = ''
                    campo_10_fecha_fin = ''
            else:
                campo_9_fecha_inicio = ''
                campo_10_fecha_fin = ''
            
            alumnos = self._procesar_alumnos(control_bytes)
            
            horas_totales = self._calcular_horas_totales()
            
            return {
                'campo_1_convocatoria': campo_1_convocatoria,
                'campo_2_accion': campo_2_accion,
                'campo_3_especialidad': campo_3_especialidad,
                'campo_4_codigo': campo_4_codigo,
                'campo_5_centro': campo_5_centro,
                'campo_6_duracion': campo_6_duracion,
                'campo_7_actividades': campo_7_actividades,
                'campo_8_modalidad': campo_8_modalidad,
                'campo_9_fecha_inicio': campo_9_fecha_inicio,
                'campo_10_fecha_fin': campo_10_fecha_fin,
                'alumnos': alumnos,
                'total_alumnos': len(alumnos),
                'horas_totales': horas_totales,
                'success': True
            }
            
        except Exception as e:
            raise Exception(f"Error procesando Excel transversales: {str(e)}")
    
    def _procesar_alumnos(self, control_bytes: bytes) -> List[Dict]:
        """
        Procesa lista de alumnos desde ASISTENCIA y RESUMEN
        
        Extrae:
        - Columna B (ASISTENCIA): Nombre
        - Columna C (ASISTENCIA): DNI
        - Columna Z (ASISTENCIA): Horas (% total convertido a horas)
        - Columna K (RESUMEN): Calificación PRL
        """

        df_asistencia = pd.read_excel(
            io.BytesIO(control_bytes),
            sheet_name='ASISTENCIA',
            header=None
        )
        
        df_resumen = pd.read_excel(
            io.BytesIO(control_bytes),
            sheet_name='RESUMEN',
            header=0
        )
        
        alumnos = []

        fila_inicio = 10
        
        for i in range(20):
            fila_idx = fila_inicio + i

            if fila_idx >= len(df_asistencia):
                break

            nombre = df_asistencia.iloc[fila_idx, 1]

            if pd.isna(nombre) or str(nombre).strip() == '':
                break

            dni = df_asistencia.iloc[fila_idx, 2]
            dni = str(dni) if pd.notna(dni) else ''

            horas_valor = df_asistencia.iloc[fila_idx, 25]

            if pd.notna(horas_valor):
                try:
                    horas_float = float(horas_valor)

                    if horas_float == int(horas_float):
                        horas_actividades = str(int(horas_float))
                    else:
                        horas_actividades = str(round(horas_float, 1))
                except:
                    horas_actividades = '0'
            else:
                horas_actividades = '0'

            if i < len(df_resumen):
                fcoo03 = df_resumen.iloc[i, 10]
                fcoo03_str = str(fcoo03).strip().upper()

                if 'APTO' in fcoo03_str and 'NO' not in fcoo03_str:
                    calificacion = 'APTO'
                elif 'NO APTO' in fcoo03_str:
                    calificacion = 'NO APTO'
                elif 'CONV' in fcoo03_str:
                    calificacion = 'PENDIENTE'
                elif 'EXENTA' in fcoo03_str or 'EXENTO' in fcoo03_str:
                    calificacion = 'EXENTO/A'
                elif fcoo03_str == 'NO' or fcoo03_str == 'NAN':
                    calificacion = 'NO APTO'
                else:
                    calificacion = ''
            else:
                calificacion = ''
            
            alumno = {
                'numero': i + 1,
                'dni': dni,
                'nombre': str(nombre),
                'horas_actividades': horas_actividades,
                'calificacion_final': calificacion
            }
            
            alumnos.append(alumno)
        
        return alumnos
    
    def _calcular_horas_totales(self) -> int:
        """Calcula total de horas del módulo transversal FCOO03"""
        return 10


def extraer_info_curso_transversales(file_bytes: bytes) -> Dict:
    """
    Función auxiliar para extraer información básica del curso
    
    Args:
        file_bytes: Bytes del archivo Excel
        
    Returns:
        Dict con información del curso
    """
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name='RESUMEN', header=None)
        
        return {
            'codigo': str(df.iloc[0, 2]) if pd.notna(df.iloc[0, 2]) else '',
            'nombre': str(df.iloc[1, 2]) if pd.notna(df.iloc[1, 2]) else '',
            'convocatoria': str(df.iloc[2, 2]) if pd.notna(df.iloc[2, 2]) else ''
        }
    except Exception as e:
        print(f"Error extrayendo info curso: {e}")
        return {}

if __name__ == "__main__":
    with open('/mnt/user-data/uploads/2024_1339_CTRL_Tareas_AREA.xlsx', 'rb') as f:
        file_bytes = f.read()
    
    processor = TransversalesProcessor()
    datos = processor.cargar_datos(file_bytes)
    
    print("\n DATOS EXTRAÍDOS:")
    print(f"Curso: {datos['curso_codigo']} - {datos['curso_nombre']}")
    print(f"Convocatoria: {datos['convocatoria']}")
    print(f"Total alumnos: {datos['total_alumnos']}")
    print(f"Horas totales: {datos['horas_totales']}")
    print(f"\n ALUMNOS:")
    for alumno in datos['alumnos'][:5]:
        print(f"  {alumno['numero']}. {alumno['nombre']} - {alumno['dni']}")
        print(f"     Horas: {alumno['horas_actividades']} | Calificación: {alumno['calificacion_final']}")