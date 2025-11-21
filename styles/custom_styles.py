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

[data-testid="stSidebar"] { display: block !important; }
header, footer, #MainMenu { visibility: hidden; }

.main .block-container {
    padding-top: 1rem !important;
    max-width: 100% !important;
}

/* =============================
   SIDEBAR
============================= */
[data-testid="stSidebar"] {
    background: var(--panel-dark) !important;
    border-right: 3px solid var(--holo-blue) !important;
}

[data-testid="stSidebar"] > div:first-child {
    background: var(--panel-dark) !important;
    padding-top: 2rem !important;
}

[data-testid="stSidebar"] h3 {
    font-family: 'Orbitron', sans-serif !important;
    color: var(--holo-blue) !important;
    font-weight: 800 !important;
    font-size: 1.5rem !important;
    text-align: center !important;
    margin-bottom: 1rem !important;
    text-shadow: 0 0 10px var(--holo-glow);
}

[data-testid="stSidebar"] hr {
    border-color: var(--holo-blue) !important;
    opacity: 0.3 !important;
}

[data-testid="stSidebar"] button {
    background: transparent !important;
    color: var(--text-light) !important;
    font-family: 'Orbitron', sans-serif !important;
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

/* =============================
   TEXTOS VISIBLES
============================= */
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

/* File Uploader - texto oscuro sobre fondo blanco */
[data-testid="stFileUploader"] {
    background-color: transparent !important;
}

/* Label principal - blanco (fuera del cuadro) */
[data-testid="stFileUploader"] > label {
    color: var(--text-bright) !important;
    font-weight: 600 !important;
}

/* Texto DENTRO del cuadro blanco - azul oscuro */
[data-testid="stFileUploader"] section {
    background-color: white !important;
}

[data-testid="stFileUploader"] section span,
[data-testid="stFileUploader"] section p,
[data-testid="stFileUploader"] section div {
    color: #1c2331 !important;
    font-weight: 500 !important;
}

/* "Drag and drop file here" - azul más oscuro */
[data-testid="stFileUploader"] section > div > div > div > span {
    color: #0b0e14 !important;
    font-weight: 600 !important;
}

/* "Limit 200MB..." - gris oscuro */
[data-testid="stFileUploader"] section small,
[data-testid="stFileUploader"] small {
    color: #4a5568 !important;
    font-weight: 400 !important;
}

/* Botón Browse files */
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

/* =============================
   FORZAR RADIO BUTTONS AZULES
============================= */

/* ELIMINAR TODO EL ROJO */
[data-testid="stRadio"] [style*="rgb(255, 75, 75)"],
[data-testid="stRadio"] [style*="#ff4b4b"],
[data-testid="stRadio"] [style*="red"] {
    background-color: var(--holo-blue) !important;
    border-color: var(--holo-blue) !important;
}

/* Contenedor */
[data-testid="stRadio"] {
    background-color: transparent !important;
}

[data-testid="stRadio"] > div {
    background-color: transparent !important;
}

/* Labels - SIEMPRE VISIBLES */
[data-testid="stRadio"] label {
    color: var(--text-bright) !important;
    background-color: transparent !important;
    display: flex !important;
    align-items: center !important;
    gap: 0.5rem !important;
}

/* Texto del label - blanco por defecto */
[data-testid="stRadio"] label span {
    color: var(--text-bright) !important;
    font-weight: 500 !important;
}

/* Texto del label cuando está checked - azul */
[data-testid="stRadio"] label:has(input:checked) span {
    color: var(--holo-blue) !important;
}

/* SVG circles (Streamlit usa esto) */
[data-testid="stRadio"] svg {
    display: block !important;
}

[data-testid="stRadio"] svg circle {
    stroke: white !important;
    fill: transparent !important;
}

/* SVG cuando está checked */
[data-testid="stRadio"] input:checked ~ * svg circle {
    stroke: var(--holo-blue) !important;
    fill: var(--holo-blue) !important;
}

/* Forzar estilos inline rojos a azul */
[data-testid="stRadio"] [style*="background-color: rgb(255, 75, 75)"],
[data-testid="stRadio"] [style*="background-color: rgb(255,75,75)"],
[data-testid="stRadio"] [style*="background-color:#ff4b4b"],
[data-testid="stRadio"] [style*="background: rgb(255, 75, 75)"] {
    background-color: var(--holo-blue) !important;
    background: var(--holo-blue) !important;
}

/* Input radio nativo */
[data-testid="stRadio"] input[type="radio"] {
    accent-color: var(--holo-blue) !important;
    opacity: 1 !important;
    width: 20px !important;
    height: 20px !important;
}

/* Divs alrededor del input */
[data-testid="stRadio"] input[type="radio"] + div {
    display: block !important;
    opacity: 1 !important;
}

/* Role radio */
[data-testid="stRadio"] [role="radio"] {
    border-color: white !important;
    display: block !important;
    opacity: 1 !important;
}

[data-testid="stRadio"] [role="radio"][aria-checked="true"] {
    background-color: var(--holo-blue) !important;
    border-color: var(--holo-blue) !important;
}

/* Baseweb */
[data-testid="stRadio"] [data-baseweb="radio"] {
    display: block !important;
    opacity: 1 !important;
}

[data-testid="stRadio"] [data-baseweb="radio"] > div:first-child {
    border-color: white !important;
    background-color: transparent !important;
}

[data-testid="stRadio"] [data-baseweb="radio"][aria-checked="true"] > div:first-child {
    background-color: var(--holo-blue) !important;
    border-color: var(--holo-blue) !important;
}

/* =============================
   TABS AZUL
============================= */
button[data-baseweb="tab"] {
    color: var(--text-bright) !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    border-bottom-color: var(--holo-blue) !important;
    color: var(--holo-blue) !important;
}

.stTabs [data-baseweb="tab-highlight"],
.stTabs [data-baseweb="tab-border"] {
    background-color: var(--holo-blue) !important;
}
</style>
"""


def get_robot_assistant(image_path='assets/robot_asistente.png'):
    """Genera el HTML del robot asistente"""
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
    bottom: 20px !important;
    left: 20px !important;
    width: 150px !important;
    height: 150px !important;
    z-index: 999999 !important;
    cursor: pointer !important;
    animation: float-robot 4s ease-in-out infinite !important;
    transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1) !important;
}}

.robot-assistant img {{
    width: 100%;
    height: 100%;
    object-fit: contain;
    filter: drop-shadow(0 10px 25px rgba(79, 195, 247, 0.4));
    transition: all 0.4s ease;
}}

@keyframes float-robot {{
    0%, 100% {{ transform: translateY(0px) rotate(0deg); }}
    50% {{ transform: translateY(-20px) rotate(2deg); }}
}}

.robot-assistant:hover {{
    transform: scale(1.15) !important;
}}

.robot-assistant:hover img {{
    filter: drop-shadow(0 15px 40px rgba(79, 195, 247, 0.6))
           drop-shadow(0 0 30px rgba(79, 195, 247, 0.4))
           brightness(1.1);
    animation: wiggle 0.5s ease !important;
}}

@keyframes wiggle {{
    0%, 100% {{ transform: rotate(0deg); }}
    25% {{ transform: rotate(-5deg); }}
    75% {{ transform: rotate(5deg); }}
}}

.robot-assistant:active {{
    transform: scale(1.05) translateY(-10px) !important;
}}

.robot-assistant::before {{
    content: '';
    position: absolute;
    width: 180px;
    height: 180px;
    background: radial-gradient(circle, rgba(79, 195, 247, 0.2), transparent 70%);
    top: -15px;
    left: -15px;
    border-radius: 50%;
    z-index: -1;
    animation: glow-blue 3s ease-in-out infinite;
}}

@keyframes glow-blue {{
    0%, 100% {{ opacity: 0.3; transform: scale(0.9); }}
    50% {{ opacity: 0.6; transform: scale(1.15); }}
}}
</style>

<div class="robot-assistant" title="¡Hola! Soy tu asistente SmartMind 🤖✨">
    <img src="data:image/png;base64,{image_base64}" alt="Robot Asistente">
</div>
"""