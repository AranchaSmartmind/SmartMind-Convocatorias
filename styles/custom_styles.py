"""
Estilos CSS y estructura HTML para SmartMind
Réplica exacta del diseño de la imagen de referencia
"""

def get_custom_styles():
    """Retorna los estilos CSS personalizados para SmartMind"""
    return """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800;900&display=swap');
    
    :root {
        /* Colores de los botones circulares */
        --color-design: #1abc9c;
        --color-web: #e67e22;
        --color-market: #e91e63;
        --color-graphic: #9b59b6;
        --color-search: #e74c3c;
        --color-develop: #3498db;
        
        /* Colores de fondo */
        --bg-brown: #7d6b5d;
        --bg-warm: #a89080;
        
        /* Textos */
        --text-dark: #2c2c2c;
        --text-light: #ffffff;
        --text-medium: #555555;
    }
    
    * {
        margin: 0;
        padding: 0;
        box-sizing: border-box;
        font-family: 'Poppins', sans-serif !important;
    }
    
    /* Ocultar elementos de Streamlit */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    /* Fondo principal - gradiente cálido como la imagen */
    .stApp {
        background: linear-gradient(135deg, #7d6b5d 0%, #9d8b7a 50%, #b8a89a 100%);
        background-attachment: fixed;
    }
    
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    /* Hero Banner - Réplica exacta */
    .hero-banner {
        background: linear-gradient(135deg, rgba(50, 50, 50, 0.9) 0%, rgba(70, 70, 70, 0.85) 100%);
        border-radius: 0;
        padding: 0;
        margin: -2rem -2rem 2rem -2rem;
        min-height: 400px;
        position: relative;
        overflow: hidden;
        display: flex;
        align-items: center;
    }
    
    .hero-content {
        padding: 3rem 4rem;
        flex: 1;
        z-index: 2;
    }
    
    .hero-badge {
        display: inline-block;
        color: var(--text-light);
        font-size: 0.75rem;
        font-weight: 500;
        letter-spacing: 2px;
        margin-bottom: 1rem;
        text-transform: uppercase;
    }
    
    .hero-title {
        color: var(--text-light) !important;
        font-size: 3.5rem !important;
        font-weight: 900 !important;
        line-height: 1.1 !important;
        margin-bottom: 1rem !important;
        text-transform: uppercase;
        letter-spacing: -2px;
    }
    
    .hero-subtitle {
        color: var(--text-light) !important;
        font-size: 1.5rem !important;
        font-weight: 400 !important;
        margin-bottom: 2rem !important;
    }
    
    .hero-image {
        position: absolute;
        right: 5%;
        top: 50%;
        transform: translateY(-50%);
        max-width: 400px;
        z-index: 1;
    }
    
    .hero-image img {
        width: 100%;
        height: auto;
    }
    
    /* Botones circulares de navegación - EXACTOS a la imagen */
    .navigation-circles {
        display: flex;
        gap: 1rem;
        flex-wrap: wrap;
        margin-top: 2rem;
        padding: 0;
    }
    
    .circle-button {
        width: 90px;
        height: 90px;
        border-radius: 50%;
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        color: white;
        font-weight: 700;
        font-size: 0.75rem;
        text-transform: uppercase;
        letter-spacing: 0.5px;
        cursor: pointer;
        transition: all 0.3s ease;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
        border: none;
        text-align: center;
        line-height: 1.2;
    }
    
    .circle-button:hover {
        transform: translateY(-5px) scale(1.05);
        box-shadow: 0 8px 25px rgba(0, 0, 0, 0.3);
    }
    
    .circle-button.active {
        transform: scale(1.1);
        box-shadow: 0 8px 30px rgba(0, 0, 0, 0.4);
    }
    
    .btn-design { background: var(--color-design); }
    .btn-web { background: var(--color-web); }
    .btn-market { background: var(--color-market); }
    .btn-graphic { background: var(--color-graphic); }
    .btn-search { background: var(--color-search); }
    .btn-develop { background: var(--color-develop); }
    
    /* Sección de clientes/logos - Como en la imagen */
    .clients-bar {
        background: rgba(255, 255, 255, 0.3);
        padding: 1.5rem 2rem;
        margin: 2rem -2rem -2rem -2rem;
        text-align: center;
    }
    
    .clients-title {
        color: var(--text-medium);
        font-size: 0.75rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 2px;
        margin-bottom: 1.5rem;
    }
    
    .clients-logos {
        display: flex;
        justify-content: center;
        align-items: center;
        gap: 3rem;
        flex-wrap: wrap;
    }
    
    .client-logo {
        height: 30px;
        opacity: 0.7;
        filter: grayscale(100%);
        transition: all 0.3s ease;
    }
    
    .client-logo:hover {
        opacity: 1;
        filter: grayscale(0%);
    }
    
    /* Contenedor de contenido principal */
    .content-section {
        background: rgba(255, 255, 255, 0.95);
        border-radius: 20px;
        padding: 3rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 40px rgba(0, 0, 0, 0.15);
    }
    
    /* Títulos */
    h1, h2, h3 {
        color: var(--text-dark) !important;
        font-weight: 700 !important;
    }
    
    h1 {
        font-size: 2.5rem !important;
        margin-bottom: 1.5rem !important;
    }
    
    h2 {
        font-size: 2rem !important;
        margin-bottom: 1rem !important;
    }
    
    h3 {
        font-size: 1.5rem !important;
        margin-bottom: 0.875rem !important;
    }
    
    /* Textos */
    p, label, span {
        color: var(--text-dark) !important;
        line-height: 1.6;
    }
    
    /* Botones principales de Streamlit */
    .stButton > button {
        background: var(--color-web) !important;
        color: white !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        border: none !important;
        border-radius: 50px !important;
        padding: 1rem 2.5rem !important;
        text-transform: uppercase;
        letter-spacing: 1px;
        box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2) !important;
        transition: all 0.3s ease !important;
    }
    
    .stButton > button:hover {
        transform: translateY(-3px) !important;
        box-shadow: 0 6px 20px rgba(0, 0, 0, 0.3) !important;
    }
    
    /* Inputs */
    .stTextInput > div > div > input,
    .stSelectbox > div > div > select,
    .stTextArea > div > div > textarea {
        background: white !important;
        border: 2px solid #ddd !important;
        border-radius: 10px !important;
        padding: 0.875rem 1rem !important;
        font-size: 1rem !important;
    }
    
    .stTextInput > div > div > input:focus,
    .stSelectbox > div > div > select:focus,
    .stTextArea > div > div > textarea:focus {
        border-color: var(--color-web) !important;
        box-shadow: 0 0 0 3px rgba(230, 126, 34, 0.1) !important;
    }
    
    /* Labels */
    .stTextInput > label,
    .stSelectbox > label,
    .stTextArea > label {
        color: var(--text-dark) !important;
        font-weight: 600 !important;
        margin-bottom: 0.5rem !important;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0.5rem;
        background: transparent;
        border-bottom: none;
    }
    
    .stTabs [data-baseweb="tab"] {
        background: rgba(255, 255, 255, 0.7);
        border-radius: 10px;
        padding: 0.75rem 1.5rem;
        font-weight: 600;
        border: none;
    }
    
    .stTabs [aria-selected="true"] {
        background: var(--color-web) !important;
        color: white !important;
    }
    
    /* File uploader */
    [data-testid="stFileUploader"] {
        background: white;
        border: 2px dashed #ddd;
        border-radius: 15px;
        padding: 2rem;
    }
    
    [data-testid="stFileUploader"]:hover {
        border-color: var(--color-web);
    }
    
    /* Métricas */
    [data-testid="stMetricValue"] {
        color: var(--color-web) !important;
        font-size: 2rem !important;
        font-weight: 800 !important;
    }
    
    /* Scrollbar */
    ::-webkit-scrollbar {
        width: 10px;
    }
    
    ::-webkit-scrollbar-track {
        background: rgba(0, 0, 0, 0.1);
    }
    
    ::-webkit-scrollbar-thumb {
        background: var(--color-web);
        border-radius: 10px;
    }
    
    /* Responsive */
    @media (max-width: 768px) {
        .hero-banner {
            margin: -1rem -1rem 2rem -1rem;
            min-height: 300px;
        }
        
        .hero-content {
            padding: 2rem;
        }
        
        .hero-title {
            font-size: 2.5rem !important;
        }
        
        .hero-image {
            display: none;
        }
        
        .circle-button {
            width: 70px;
            height: 70px;
            font-size: 0.65rem;
        }
        
        .content-section {
            padding: 2rem;
        }
    }
</style>
"""

def get_hero_html():
    """Retorna el HTML del hero banner con botones circulares"""
    return """
<div class="hero-banner">
    <div class="hero-content">
        <div class="hero-badge">SIT BACK & RELAX, WE WILL</div>
        <h1 class="hero-title">APRENDE<br>MEJORA &<br>CRECE</h1>
        <p class="hero-subtitle">CON NUESTRAS FORMACIONES</p>
        
        <div class="navigation-circles">
            <button class="circle-button btn-design" onclick="navigateToSection('design')">
                DISEÑO
            </button>
            <button class="circle-button btn-web" onclick="navigateToSection('web')">
                WEB
            </button>
            <button class="circle-button btn-market" onclick="navigateToSection('marketing')">
                MARKETING
            </button>
            <button class="circle-button btn-graphic" onclick="navigateToSection('grafico')">
                GRÁFICO
            </button>
            <button class="circle-button btn-search" onclick="navigateToSection('seo')">
                SEO
            </button>
            <button class="circle-button btn-develop" onclick="navigateToSection('desarrollo')">
                DESARROLLO
            </button>
        </div>
    </div>
    
    <div class="hero-image">
        <img src="https://images.unsplash.com/photo-1586023492125-27b2c045efd7?w=500" alt="Formación" />
    </div>
</div>

<script>
function navigateToSection(section) {
    // Aquí puedes añadir la lógica para cambiar de pestaña
    console.log('Navegando a:', section);
    
    // Ejemplo: si usas Streamlit con query params
    const url = new URL(window.location);
    url.searchParams.set('section', section);
    window.history.pushState({}, '', url);
    
    // O si prefieres scroll a una sección
    const element = document.getElementById(section);
    if (element) {
        element.scrollIntoView({ behavior: 'smooth' });
    }
}

// Marcar botón activo según la sección actual
document.addEventListener('DOMContentLoaded', function() {
    const urlParams = new URLSearchParams(window.location.search);
    const currentSection = urlParams.get('section');
    
    if (currentSection) {
        document.querySelectorAll('.circle-button').forEach(btn => {
            btn.classList.remove('active');
        });
        
        const activeBtn = document.querySelector(`.circle-button[onclick*="${currentSection}"]`);
        if (activeBtn) {
            activeBtn.classList.add('active');
        }
    }
});
</script>

<div class="clients-bar">
    <div class="clients-title">ALGUNOS DE NUESTROS CLIENTES</div>
    <div class="clients-logos">
        <!-- Aquí puedes añadir los logos de tus clientes -->
        <span style="color: #666; font-weight: 600;">CLIENTE 1</span>
        <span style="color: #666; font-weight: 600;">CLIENTE 2</span>
        <span style="color: #666; font-weight: 600;">CLIENTE 3</span>
        <span style="color: #666; font-weight: 600;">CLIENTE 4</span>
        <span style="color: #666; font-weight: 600;">CLIENTE 5</span>
    </div>
</div>
"""