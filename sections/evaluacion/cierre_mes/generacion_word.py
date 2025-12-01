import os
from docx import Document

def construir_observaciones(ayudas, dias_aula, dias_empresa, justificantes, dias_lectivos, faltas=0):
    partes = []
    
    if ayudas:
        if 'Discapacidad' in ayudas:
            if len(ayudas) == 1:
                partes.append(f"Discapacidad+{dias_lectivos}")
            else:
                otras = [a for a in ayudas if a != 'Discapacidad']
                texto_otras = '+ '.join(otras)
                partes.append(f"{texto_otras}+ Discapacidad: {dias_lectivos}")
        else:
            texto_ayudas = '+ '.join(ayudas)
            partes.append(f"{texto_ayudas}: {dias_lectivos}")
    
    if dias_aula > 0:
        partes.append(f"{dias_aula} dias aula")
    
    if justificantes > 0:
        if justificantes == 1:
            partes.append("1 falta justificada")
        else:
            partes.append(f"{justificantes} faltas justificadas")
    
    return ' '.join(partes) if partes else ''

def formatear_nombre_oficial(nombre_completo):
    """
    Convierte 'APELLIDOS, NOMBRE' a 'NOMBRE APELLIDOS'
    Ejemplo: 'BRAÑA MANCHADO, NURIA' -> 'NURIA BRAÑA MANCHADO'
    """
    if ',' in nombre_completo:
        partes = nombre_completo.split(',')
        apellidos = partes[0].strip()
        nombre = partes[1].strip()
        return f"{nombre} {apellidos}"
    return nombre_completo

def rellenar_template(template_path, datos):
    try:
        doc = Document(template_path)
        
        if not doc.tables:
            print("Error: No se encontro tabla")
            return None
        
        table = doc.tables[0]
        
        # FILA 0: Año Nº de
        table.rows[0].cells[10].text = datos['numero_curso']
        
        # FILA 1: Especialidad (ocupa múltiples celdas combinadas)
        especialidad = datos['especialidad']
        for i in range(1, 11):
            table.rows[1].cells[i].text = especialidad
        
        # FILA 2: Centro (línea completa)
        centro = datos.get('centro', 'INTERPROS NEXT GENERATION SLU')
        for i in range(1, 11):
            table.rows[2].cells[i].text = centro
        
        # FILA 3: Mes y información de horas
        mes = datos['mes']
        dias_lectivos = datos['dias_lectivos']
        horas_empresa = datos.get('horas_empresa', 21)
        horas_aula = datos.get('horas_aula', 2)
        
        table.rows[3].cells[0].text = f"Mes de {mes}"
        # La celda combinada con información
        table.rows[3].cells[1].text = f"23 ({horas_empresa} horas en empresa +{horas_aula}"
        # "Número de días lectivos" en celdas 5-10
        texto_dias = f"Número de días lectivos {dias_lectivos} aula)"
        for i in range(5, 10):
            table.rows[3].cells[i].text = texto_dias
        
        # FILA 4: "LISTADO DEL ALUMNADO" (ya está en template)
        # FILA 5: Headers (ya están en template)
        
        # FILAS 6+: Datos de alumnos
        for idx, alumno in enumerate(datos['alumnos']):
            if idx >= 20:
                break
            
            fila_idx = 6 + idx
            if fila_idx >= len(table.rows):
                break
                
            row = table.rows[fila_idx]
            
            # Columna 1: Número de orden
            row.cells[1].text = str(idx + 1)
            
            # Columnas 2-3: Nombre y Apellidos (FORMATO OFICIAL: NOMBRE APELLIDOS)
            nombre_oficial = formatear_nombre_oficial(alumno['nombre'])
            row.cells[2].text = nombre_oficial
            row.cells[3].text = nombre_oficial
            
            # Columnas 4-7: NIF
            dni = alumno.get('dni', '')
            for i in range(4, 8):
                row.cells[i].text = dni
            
            # Columna 8: Nº de faltas
            row.cells[8].text = str(alumno['faltas'])
            
            # Columnas 9-10: Observaciones
            obs = alumno['observaciones']
            row.cells[9].text = obs
            row.cells[10].text = obs
        
        return doc
    
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return None

def generar_parte_mensual(template_path, output_path, datos):
    try:
        doc = rellenar_template(template_path, datos)
        
        if doc is None:
            return False
        
        doc.save(output_path)
        print(f"Documento generado: {output_path}")
        return True
        
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        return False