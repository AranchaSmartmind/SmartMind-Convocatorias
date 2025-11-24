"""
═══════════════════════════════════════════════════════════════════════════════
SMARTMIND - CONFIGURACIÓN PARA PARTE MENSUAL
═══════════════════════════════════════════════════════════════════════════════
Archivo de configuración para personalizar la generación de partes mensuales
═══════════════════════════════════════════════════════════════════════════════
"""

# ═══════════════════════════════════════════════════════════════════════════════
# DÍAS LECTIVOS (MANUAL)
# ═══════════════════════════════════════════════════════════════════════════════

# Si el cálculo automático no es correcto, especifica el número exacto:
DIAS_LECTIVOS_MANUAL = None  # Cambiar a número si se necesita (ej: 23)

# ═══════════════════════════════════════════════════════════════════════════════
# FALTAS POR ALUMNO (MANUAL)
# ═══════════════════════════════════════════════════════════════════════════════

# Formato: {'NIF': número_de_faltas}
# Si un alumno no está aquí, tendrá 0 faltas

FALTAS_MANUALES = {}

# Ejemplo:
# FALTAS_MANUALES = {
#     '71879712C': 2,   # NURIA BRAÑA MANCHADO - 2 faltas
#     '43759848W': 0,   # MARIA DELIA GUERRA OJEDA - 0 faltas
#     '71779864S': 4,   # EDGAR PORRON VAZQUEZ - 4 faltas
# }

# ═══════════════════════════════════════════════════════════════════════════════
# FALTAS JUSTIFICADAS (MANUAL)
# ═══════════════════════════════════════════════════════════════════════════════

# Formato: {'NIF': número_de_faltas_justificadas}
# Se añadirá "X falta/s justificada/s" al final de las observaciones

FALTAS_JUSTIFICADAS = {}

# Ejemplo:
# FALTAS_JUSTIFICADAS = {
#     'Z2762262J': 1,   # MARYELY YORSELY - 1 falta justificada
#     '71779864S': 1,   # EDGAR PORRON - 1 falta justificada
# }

# ═══════════════════════════════════════════════════════════════════════════════
# OBSERVACIONES POR ALUMNO (MANUAL)
# ═══════════════════════════════════════════════════════════════════════════════

# Si quieres especificar observaciones manualmente (sobrescribe cálculo automático)
# Formato: {'NIF': 'Observación exacta'}

OBSERVACIONES_MANUALES = {}

# Ejemplo:
# OBSERVACIONES_MANUALES = {
#     'Z2751442A': 'Transporte + Conciliación: 23',  # JOSE ALEJANDRO
#     '11422042N': 'Discapacidad: 23',                # LAURA MARIA
#     'Z2762262J': 'Transporte + Conciliación: 19',   # MARYELY YORSELY
# }

# ═══════════════════════════════════════════════════════════════════════════════
# NOTAS:
# ═══════════════════════════════════════════════════════════════════════════════
# 
# - Si OBSERVACIONES_MANUALES está vacío {}, se calculan automáticamente
#   desde el PDF de otorgamiento
# 
# - Las observaciones automáticas tienen el formato:
#   "Transporte: XX" o "Transporte + Conciliación: XX" 
#   donde XX = días asistidos (días lectivos - faltas)
# 
# - Las faltas justificadas se añaden al final:
#   "Transporte: 23, 1 falta justificada"
# 
# - Para usar NIF correcto, consulta el Excel CTRL o el PDF de otorgamiento
# 
# ═══════════════════════════════════════════════════════════════════════════════