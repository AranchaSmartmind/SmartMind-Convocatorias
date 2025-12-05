"""
Módulo de funciones auxiliares
Funciones de utilidad general
"""

def normalizar_nombre(nombre):
    """
    Normaliza un nombre para facilitar comparaciones
    
    Args:
        nombre: Nombre a normalizar
    
    Returns:
        str: Nombre normalizado en mayúsculas sin espacios extras
    """
    return ' '.join(nombre.upper().split())

def buscar_coincidencia(nombre_buscar, diccionario):
    """
    Busca un nombre en un diccionario con coincidencia flexible
    Útil para relacionar datos entre diferentes fuentes
    
    Args:
        nombre_buscar: Nombre a buscar
        diccionario: Diccionario donde buscar
    
    Returns:
        El valor encontrado o None
    """
    if not diccionario:
        return None
    
    nombre_norm = normalizar_nombre(nombre_buscar)
    
    # Extraer apellidos
    if ',' in nombre_norm:
        apellidos = nombre_norm.split(',')[0].strip()
    else:
        apellidos = nombre_norm.split()[0] if nombre_norm.split() else nombre_norm
    
    # Búsqueda exacta
    for key in diccionario.keys():
        if normalizar_nombre(key) == nombre_norm:
            return diccionario[key]
    
    # Búsqueda por apellidos
    for key in diccionario.keys():
        key_norm = normalizar_nombre(key)
        
        if ',' in key_norm:
            key_apellidos = key_norm.split(',')[0].strip()
        else:
            key_apellidos = key_norm.split()[0] if key_norm.split() else key_norm
        
        # Comparar apellidos
        if apellidos in key_norm or key_apellidos in nombre_norm:
            return diccionario[key]
    
    return None

def validar_dni(dni):
    """
    Valida formato básico de DNI/NIE
    
    Args:
        dni: Cadena con el DNI/NIE
    
    Returns:
        bool: True si tiene formato válido
    """
    if not dni:
        return False
    
    dni = dni.upper().strip()
    
    # Formato DNI: 8 dígitos + letra
    if len(dni) == 9:
        return dni[:8].isdigit() and dni[8].isalpha()
    
    # Formato NIE: X/Y/Z + 7 dígitos + letra
    if len(dni) == 9 and dni[0] in 'XYZ':
        return dni[1:8].isdigit() and dni[8].isalpha()
    
    return False

def formatear_fecha(fecha, formato='%d/%m/%Y'):
    """
    Formatea un objeto datetime a string
    
    Args:
        fecha: Objeto datetime
        formato: Formato de salida
    
    Returns:
        str: Fecha formateada
    """
    try:
        return fecha.strftime(formato)
    except:
        return str(fecha)