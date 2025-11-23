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
            # Leer hojas necesarias del control
            self.df_resumen = pd.read_excel(
                io.BytesIO(control_bytes), 
                sheet_name='RESUMEN',
                header=0
            )
            
            self.df_calificaciones = pd.read_excel(
                io.BytesIO(control_bytes),
                sheet_name='CALIFICACIONES',
                header=7  # La fila 7 tiene los encabezados
            )
            
            # ============================================================
            # EXTRAER CAMPOS 1-10 DEL ENCABEZADO
            # ============================================================
            
            # CAMPO 1: Convocatoria (vacío por defecto)
            campo_1_convocatoria = ''
            
            # CAMPO 2: Acción - De ASISTENCIA, fila 1, columna C
            df_asistencia = pd.read_excel(
                io.BytesIO(control_bytes),
                sheet_name='ASISTENCIA',
                header=None
            )
            campo_2_accion = str(df_asistencia.iloc[0, 2]) if pd.notna(df_asistencia.iloc[0, 2]) else ''
            
            # CAMPOS 3, 4, 6, 8, 9, 10: Del Cronograma
            df_cronograma = pd.read_excel(
                io.BytesIO(cronograma_bytes),
                sheet_name='Cronograma',
                header=None
            )
            
            # CAMPO 3 y 4: Fila 20, columna A
            valor_fila20 = str(df_cronograma.iloc[19, 0])  # Fila 20 = índice 19
            
            # Separar código (FCOO03) y nombre (resto)
            partes = valor_fila20.split(None, 1)
            if len(partes) >= 2:
                campo_4_codigo = partes[0]  # FCOO03
                campo_3_especialidad = partes[1]  # Inserc. Lab., Sensib...
            else:
                campo_4_codigo = ''
                campo_3_especialidad = valor_fila20
            
            # CAMPO 5: Centro (fijo)
            campo_5_centro = 'INTERPROS NEXT GENERATION S.L.U'
            
            # CAMPO 6: Duración - Fila 20, columna C
            campo_6_duracion = str(df_cronograma.iloc[19, 2]) if pd.notna(df_cronograma.iloc[19, 2]) else ''
            
            # CAMPO 7: Actividades totales (fijo = 1)
            campo_7_actividades = '1'
            
            # CAMPO 8: Modalidad - Fila 10, buscar "Presencial"
            campo_8_modalidad = 'Presencial'  # Por defecto
            for col in range(15):
                valor = df_cronograma.iloc[9, col]  # Fila 10 = índice 9
                if pd.notna(valor) and 'Modalidad' in str(valor):
                    if 'Presencial' in str(valor):
                        campo_8_modalidad = 'Presencial'
                        break
            
            # CAMPOS 9 y 10: Fechas - Fila 20, columna I
            valor_fechas = str(df_cronograma.iloc[19, 8])  # Columna I = índice 8
            
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
            
            # ============================================================
            # FIN CAMPOS 1-10
            # ============================================================
            
            # Procesar alumnos
            alumnos = self._procesar_alumnos(control_bytes)
            
            # Calcular horas totales
            horas_totales = self._calcular_horas_totales()
            
            return {
                # Campos del encabezado (1-10)
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
                # Datos de alumnos
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
        
        # Leer hojas
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
        
        # Los datos empiezan en fila 11 del Excel (índice 10)
        # La fila 10 es el encabezado "ALUMNO"
        fila_inicio = 10
        
        for i in range(20):  # Máximo 20 alumnos
            fila_idx = fila_inicio + i
            
            # Verificar que no pasamos el final del Excel
            if fila_idx >= len(df_asistencia):
                break
            
            # Columna B = índice 1 (Nombre)
            nombre = df_asistencia.iloc[fila_idx, 1]
            
            # Si no hay nombre, terminamos
            if pd.isna(nombre) or str(nombre).strip() == '':
                break
            
            # Columna C = índice 2 (DNI)
            dni = df_asistencia.iloc[fila_idx, 2]
            dni = str(dni) if pd.notna(dni) else ''
            
            # Columna Z = índice 25 (Horas - YA son horas directas, no porcentaje)
            horas_valor = df_asistencia.iloc[fila_idx, 25]
            
            # Las horas ya vienen calculadas en la columna Z
            if pd.notna(horas_valor):
                try:
                    horas_float = float(horas_valor)
                    # Redondear a 1 decimal si es necesario
                    if horas_float == int(horas_float):
                        horas_actividades = str(int(horas_float))
                    else:
                        horas_actividades = str(round(horas_float, 1))
                except:
                    horas_actividades = '0'
            else:
                horas_actividades = '0'
            
            # Columna FCOO03 del RESUMEN (índice 10) - Calificación
            if i < len(df_resumen):
                fcoo03 = df_resumen.iloc[i, 10]  # Columna FCOO03
                fcoo03_str = str(fcoo03).strip().upper()
                
                # Convertir a formato del acta
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
        # El módulo FCOO03 tiene 10 horas
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


# EJEMPLO DE USO:
if __name__ == "__main__":
    # Probar con el archivo subido
    with open('/mnt/user-data/uploads/2024_1339_CTRL_Tareas_AREA.xlsx', 'rb') as f:
        file_bytes = f.read()
    
    processor = TransversalesProcessor()
    datos = processor.cargar_datos(file_bytes)
    
    print("\n📊 DATOS EXTRAÍDOS:")
    print(f"Curso: {datos['curso_codigo']} - {datos['curso_nombre']}")
    print(f"Convocatoria: {datos['convocatoria']}")
    print(f"Total alumnos: {datos['total_alumnos']}")
    print(f"Horas totales: {datos['horas_totales']}")
    print(f"\n👥 ALUMNOS:")
    for alumno in datos['alumnos'][:5]:  # Mostrar solo primeros 5
        print(f"  {alumno['numero']}. {alumno['nombre']} - {alumno['dni']}")
        print(f"     Horas: {alumno['horas_actividades']} | Calificación: {alumno['calificacion_final']}")