def get_custom_styles():
    return """
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@600;800&family=Poppins:wght@300;400;500&display=swap');

:root {
    --holo-blue: #4fc3f7;
    --holo-glow: rgba(79,195,247,0.7);
    --panel-dark: rgba(20,25,35,0.94);
    --text-light: #e6f7ff;
    --text-bright: #ffffff;
}

.stApp {
    background: radial-gradient(circle at 20% 20%, #1c2331, #0b0e14 70%);
    color: var(--text-bright);
    font-family: 'Poppins', sans-serif;
}

/* ========================================
   SIDEBAR FIJO - SIEMPRE VISIBLE
   ======================================== */
section[data-testid="stSidebar"],
[data-testid="stSidebar"] {
    display: block !important;
    visibility: visible !important;
    opacity: 1 !important;
    position: relative !important;
    left: 0 !important;
    transform: translateX(0) !important;
    width: 21rem !important;
    min-width: 21rem !important;
    background: var(--panel-dark) !important;
    border-right: 3px solid var(--holo-blue) !important;
    padding-top: 0 !important;
    z-index: 999 !important;
}

/* OCULTAR BOTÓN DE COLAPSAR - SIDEBAR FIJO */
[data-testid="collapsedControl"],
button[kind="header"] {
    display: none !important;
    visibility: hidden !important;
}

/* ELIMINAR SOLO EL HEADER DEL SIDEBAR (no todo el contenido) */
[data-testid="stSidebarHeader"] {
    display: none !important;
    visibility: hidden !important;
    height: 0 !important;
    padding: 0 !important;
    margin: 0 !important;
}

/* Ocultar solo el logo de Streamlit */
div[data-testid="stSidebar"] img[alt*="streamlit"] {
    display: none !important;
}

/* Ocultar el contenedor morado del header */
div.st-emotion-cache-6f82ta4 {
    display: none !important;
}

header, footer, #MainMenu { 
    visibility: hidden !important;
}

.main .block-container {
    padding-top: 1rem !important;
    max-width: 100% !important;
}

/* ========================================
   ESTILOS INTERNOS DEL SIDEBAR
   ======================================== */
[data-testid="stSidebar"] > div:first-child {
    background: var(--panel-dark) !important;
    padding-top: 0 !important;
    padding-bottom: 0 !important;
}

[data-testid="stSidebar"] .block-container,
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
    padding-top: 0 !important;
    padding-bottom: 0 !important;
    gap: 0 !important;
}

[data-testid="stSidebar"] > div > div {
    padding-top: 0 !important;
    margin-top: 0 !important;
}

[data-testid="stSidebar"] h3 {
    font-family: 'Orbitron', sans-serif !important;
    color: var(--holo-blue) !important;
    font-weight: 800 !important;
    font-size: 1.2rem !important;
    text-align: center !important;
    margin-top: 0rem !important;
    margin-bottom: 0.4rem !important;
    text-shadow: 0 0 10px var(--holo-glow);
}

[data-testid="stSidebar"] hr {
    border-color: var(--holo-blue) !important;
    opacity: 0.3 !important;
    margin-top: 1rem !important;
    margin-bottom: 0.4rem !important;
}

[data-testid="stSidebar"] button {
    background: transparent !important;
    color: var(--text-light) !important;
    font-family: 'Orbitron', sans-serif !important;
    margin-top: 0.5rem !important;
    font-weight: 600 !important;
    font-size: 1.05rem !important;
    border: 2px solid transparent !important;
    border-radius: 10px !important;
    padding: 1rem 1.2rem !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    width: 100% !important;
    text-align: left !important;
    margin-bottom: 0.5rem !important;
    letter-spacing: 0.5px !important;
}

[data-testid="stSidebar"] button:hover {
    background: rgba(79,195,247,0.15) !important;
    border-color: var(--holo-blue) !important;
    color: var(--holo-blue) !important;
    transform: translateX(5px) !important;
    box-shadow: 0 4px 15px rgba(79,195,247,0.3) !important;
}

[data-testid="stSidebar"] button:active,
[data-testid="stSidebar"] button:focus {
    background: rgba(79,195,247,0.25) !important;
    border-color: var(--holo-blue) !important;
    color: var(--holo-blue) !important;
    box-shadow: 0 4px 20px rgba(79,195,247,0.4) !important;
}

[data-testid="stSidebar"]::-webkit-scrollbar {
    width: 10px;
}

[data-testid="stSidebar"]::-webkit-scrollbar-thumb {
    background: var(--holo-blue);
    border-radius: 5px;
}

[data-testid="stSidebar"]::-webkit-scrollbar-track {
    background: rgba(0,0,0,0.2);
}

/* ========================================
   BOTONES PRINCIPALES
   ======================================== */
.stButton > button,
.main .stButton > button,
button[kind="primary"],
button[kind="secondary"],
.stDownloadButton > button,
button[data-baseweb="button"] {
    background: transparent !important;
    color: var(--text-light) !important;
    font-family: 'Orbitron', sans-serif !important;
    font-weight: 600 !important;
    font-size: 1.05rem !important;
    border: 2px solid var(--holo-blue) !important;
    border-radius: 10px !important;
    padding: 0.75rem 1.5rem !important;
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    letter-spacing: 0.5px !important;
}

.stButton > button:hover,
.main .stButton > button:hover,
button[kind="primary"]:hover,
button[kind="secondary"]:hover,
.stDownloadButton > button:hover,
button[data-baseweb="button"]:hover {
    background: rgba(79,195,247,0.15) !important;
    border-color: var(--holo-blue) !important;
    color: var(--holo-blue) !important;
    transform: translateY(-2px) !important;
    box-shadow: 0 4px 15px rgba(79,195,247,0.3) !important;
}

.stButton > button:active,
.stButton > button:focus,
.main .stButton > button:active,
button[kind="primary"]:active,
button[kind="secondary"]:active,
.stDownloadButton > button:active,
button[data-baseweb="button"]:active {
    background: rgba(79,195,247,0.25) !important;
    transform: translateY(0px) !important;
    box-shadow: 0 4px 20px rgba(79,195,247,0.4) !important;
}

/* ========================================
   TEXTOS Y TIPOGRAFÍA
   ======================================== */
p, span, div, label {
    color: var(--text-bright) !important;
}

h1, h2, h3, h4, h5, h6 {
    font-family: 'Orbitron', sans-serif !important;
    font-weight: 800 !important;
    color: var(--text-bright) !important;
}

h1 {
    font-size: 2rem !important;
    margin-bottom: 0.5rem !important;
}

label {
    color: var(--text-bright) !important;
    font-weight: 500 !important;
}

/* ========================================
   FILE UPLOADER
   ======================================== */
[data-testid="stFileUploader"] > label {
    color: var(--text-bright) !important;
    font-weight: 600 !important;
}

[data-testid="stFileUploader"] section {
    background-color: white !important;
}

[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section div {
    color: #1c2331 !important;
    font-weight: 500 !important;
}

[data-testid="stFileUploader"] section > div > div > div > span {
    color: #0b0e14 !important;
    font-weight: 600 !important;
}

[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] small {
    color: #4a5568 !important;
    font-weight: 400 !important;
}

[data-testid="stFileUploader"] button {
    color: #1c2331 !important;
    border-color: #1c2331 !important;
    background-color: transparent !important;
    font-weight: 600 !important;
}

[data-testid="stFileUploader"] button:hover {
    background-color: rgba(28, 35, 49, 0.1) !important;
    color: #0b0e14 !important;
}

/* ========================================
   TABS
   ======================================== */
button[data-baseweb="tab"] {
    color: rgba(255, 255, 255, 0.6) !important;
    font-weight: 500 !important;
    border-bottom: 3px solid transparent !important;
    transition: all 0.3s ease !important;
}

button[data-baseweb="tab"]:hover {
    color: var(--text-bright) !important;
    border-bottom-color: rgba(79, 195, 247, 0.5) !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    color: var(--holo-blue) !important;
    font-weight: 700 !important;
    border-bottom: 3px solid var(--holo-blue) !important;
    background: rgba(79, 195, 247, 0.1) !important;
}

.stTabs [data-baseweb="tab-highlight"] {
    background-color: var(--holo-blue) !important;
    height: 3px !important;
}

.stTabs [data-baseweb="tab-border"] {
    background-color: rgba(79, 195, 247, 0.3) !important;
}

[data-baseweb="tab-list"] {
    gap: 0 !important;
}

button[data-baseweb="tab"] {
    padding: 0.75rem 1.5rem !important;
}
</style>
"""


def get_interpros_logo(image_path='assets/logo.png'):
    """Genera el HTML del logo de INTERPROS - VERSIÓN MINI"""
    import base64
    import os

    ubicaciones = [
        image_path,
        os.path.join(os.getcwd(), 'assets', 'logo.png'),
        os.path.join(os.path.dirname(__file__), '..', 'assets', 'logo.png'),
    ]
    
    image_base64 = None
    for ubicacion in ubicaciones:
        if os.path.exists(ubicacion):
            try:
                with open(ubicacion, 'rb') as f:
                    image_base64 = base64.b64encode(f.read()).decode()
                print(f"✅ Logo cargado desde: {ubicacion}")
                break
            except Exception as e:
                print(f"❌ Error cargando {ubicacion}: {e}")
                continue
    
    if not image_base64:
        print("⚠️ No se encontró el logo en ninguna ubicación")
        return ""
    
    return f"""
<style>
.interpros-logo {{
    text-align: center;
    margin: 0;
    margin-top: 2rem;
    margin-bottom: 2rem;
    padding: 0;
    line-height: 0;
}}

.interpros-logo img {{
    width: 70px;
    height: 70px;
    border-radius: 50%;
    background: white;
    padding: 10px;
    object-fit: contain;
    display: block;
    margin: 0 auto;
}}
</style>
<div class="interpros-logo">
    <img src="data:image/png;base64,{image_base64}" alt="Logo">
</div>
"""


def get_robot_assistant(image_path='assets/robot_asistente.png'):
    """Genera el HTML del robot asistente - SIEMPRE en esquina inferior izquierda"""
    import base64
    import os
    
    image_base64 = None
    if os.path.exists(image_path):
        try:
            with open(image_path, 'rb') as f:
                image_base64 = base64.b64encode(f.read()).decode()
        except:
            pass
    
    if not image_base64:
        image_base64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=="
    
    return f"""
<style>
.robot-assistant {{
    position: fixed !important;
    bottom: 60px !important;
    left: 80px !important;
    width: 150px !important;
    height: 150px !important;
    z-index: 999999 !important;
    cursor: pointer !important;
    animation: float-robot 4s ease-in-out infinite !important;
    transition: all 0.4s ease !important;
}}

.robot-assistant img {{
    width: 100%;
    height: 100%;
    object-fit: contain;
    filter: drop-shadow(0 10px 25px rgba(79, 195, 247, 0.4));
    transition: all 0.4s ease;
}}

@keyframes float-robot {{
    0%, 100% {{ transform: translateY(0px); }}
    50% {{ transform: translateY(-15px); }}
}}

.robot-assistant:hover {{
    transform: scale(1.08) !important;
}}

.robot-assistant:hover img {{
    filter: drop-shadow(0 15px 40px rgba(79, 195, 247, 0.6))
           drop-shadow(0 0 30px rgba(79, 195, 247, 0.4))
           brightness(1.1);
    animation: wiggle 2s ease-in-out infinite !important;
}}

@keyframes wiggle {{
    0%, 100% {{ transform: rotate(0deg); }}
    10% {{ transform: rotate(-8deg); }}
    20% {{ transform: rotate(8deg); }}
    30% {{ transform: rotate(-8deg); }}
    40% {{ transform: rotate(8deg); }}
    50% {{ transform: rotate(-5deg); }}
    60% {{ transform: rotate(5deg); }}
    70% {{ transform: rotate(-3deg); }}
    80% {{ transform: rotate(3deg); }}
    90% {{ transform: rotate(0deg); }}
}}

.robot-assistant:active {{
    transform: scale(1.02) translateY(-5px) !important;
}}

.robot-assistant::before {{
    content: '';
    position: absolute;
    width: 180px;
    height: 180px;
    background: radial-gradient(circle, rgba(79, 195, 247, 0.2), transparent 70%);
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    border-radius: 50%;
    z-index: -1;
    animation: glow-blue 3s ease-in-out infinite;
}}

@keyframes glow-blue {{
    0%, 100% {{ 
        opacity: 0.3; 
        transform: translate(-50%, -50%) scale(0.9); 
    }}
    50% {{ 
        opacity: 0.6; 
        transform: translate(-50%, -50%) scale(1.15); 
    }}
}}

.robot-assistant::after {{
    content: '¡Hola! Soy tu Asistente SmartMind';
    position: absolute;
    top: 50%;
    left: 165px;
    background: linear-gradient(135deg, rgba(79, 195, 247, 0.95), rgba(79, 195, 247, 0.85));
    color: white;
    padding: 12px 20px;
    border-radius: 20px;
    font-family: 'Poppins', sans-serif;
    font-size: 14px;
    font-weight: 500;
    white-space: nowrap;
    box-shadow: 0 8px 20px rgba(79, 195, 247, 0.4);
    z-index: 1000000;
    pointer-events: none;
    transform: translateY(-50%) scale(0);
    opacity: 0;
    transition: all 0.3s ease;
}}

.robot-assistant:hover::after {{
    transform: translateY(-50%) scale(1);
    opacity: 1;
}}
</style>

<div class="robot-assistant">
    <img src="data:image/png;base64,{image_base64}" alt="Robot Asistente">
</div>
"""
