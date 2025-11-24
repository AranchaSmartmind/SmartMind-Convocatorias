"""
Funciones para procesar archivos Excel
"""
import io
import pandas as pd
import openpyxl
import streamlit as st
import re
from datetime import datetime, date


def leer_datos_ctrl(excel_file):
    """Lee la pestaña CTRL del Excel CTRL de Alumnos"""
    datos_ctrl = {}

    try:
        excel_file.seek(0)

        if "CTRL" not in pd.ExcelFile(excel_file).sheet_names:
            st.warning("No se encontró la pestaña 'CTRL' en el Excel CTRL")
            return datos_ctrl

        df_ctrl = pd.read_excel(excel_file, sheet_name="CTRL")

        col_nombre = None
        col_dni = None
        col_corporacion = None
        col_baja = None
        col_motivo = None
        col_baja_ocupacion = None
        col_fecha_incorporacion = None

        for col in df_ctrl.columns:
            col_lower = str(col).lower().strip()

            if 'nombre' in col_lower or 'alumno' in col_lower:
                col_nombre = col
            elif 'dni' in col_lower or 'nif' in col_lower:
                col_dni = col
            elif ('incorporacion' in col_lower or 'incorporación' in col_lower) and 'fecha' in col_lower:
                col_fecha_incorporacion = col
            elif 'corporacion' in col_lower or 'corporación' in col_lower:
                col_corporacion = col
            elif 'motivo' in col_lower:
                col_motivo = col
            elif ('baja' in col_lower and ('ocupacion' in col_lower or 'ocupación' in col_lower)) or \
                 ('ocupacion' in col_lower or 'ocupación' in col_lower) and '%' in col_lower:
                col_baja_ocupacion = col
            elif 'baja' in col_lower and 'fecha' in col_lower:
                col_baja = col

        for _, row in df_ctrl.iterrows():
            if col_nombre and pd.notna(row.get(col_nombre)):
                nombre = str(row[col_nombre]).strip().upper()

                if nombre:
                    fecha_incorporacion_valor = None

                    if col_fecha_incorporacion:
                        fecha_val = row.get(col_fecha_incorporacion)

                        if pd.notna(fecha_val):
                            if isinstance(fecha_val, pd.Timestamp):
                                fecha_incorporacion_valor = date(fecha_val.year, fecha_val.month, fecha_val.day)
                            elif isinstance(fecha_val, datetime):
                                fecha_incorporacion_valor = date(fecha_val.year, fecha_val.month, fecha_val.day)
                            else:
                                fecha_incorporacion_valor = str(fecha_val)

                    baja_valor = None
                    if col_baja and pd.notna(row.get(col_baja)):
                        baja_raw = row[col_baja]
                        if isinstance(baja_raw, pd.Timestamp):
                            baja_valor = date(baja_raw.year, baja_raw.month, baja_raw.day)
                        elif isinstance(baja_raw, datetime):
                            baja_valor = date(baja_raw.year, baja_raw.month, baja_raw.day)
                        else:
                            baja_valor = str(baja_raw)

                    datos_ctrl[nombre] = {
                        "dni": str(row[col_dni]) if col_dni and pd.notna(row.get(col_dni)) else "",
                        "corporacion_a_clase": str(row[col_corporacion]) if col_corporacion and pd.notna(row.get(col_corporacion)) else "",
                        "baja": baja_valor if baja_valor else "",
                        "motivo": str(row[col_motivo]) if col_motivo and pd.notna(row.get(col_motivo)) else "",
                        "motivo_sin_parentesis": re.sub(r'\s*\([^)]*\)', '', str(row[col_motivo])).strip() if col_motivo and pd.notna(row.get(col_motivo)) else "",
                        "baja_ocupacion": str(row[col_baja_ocupacion]) if col_baja_ocupacion and pd.notna(row.get(col_baja_ocupacion)) else "",
                        "fecha_incorporacion": fecha_incorporacion_valor
                    }

        st.success(f"Datos del CTRL leídos: {len(datos_ctrl)} alumnos encontrados")
        return datos_ctrl

    except Exception as e:
        st.error(f"Error al leer Excel CTRL: {str(e)}")
        import traceback
        st.code(traceback.format_exc())
        return datos_ctrl


def leer_datos_excel(excel_file, datos_evaluacion=None):
    """Lee las pestañas CALIFICACIONES y ASISTENCIA del Excel"""
    datos = {
        "alumnos": {},
    }

    
    return datos