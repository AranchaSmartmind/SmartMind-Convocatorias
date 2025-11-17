"""
Estilos CSS personalizados para la aplicación
"""

def get_custom_styles():
    """Retorna los estilos CSS personalizados"""
    return """
<style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700;800&display=swap');
    
    :root {
        --primary-color: #2c3e50;
        --secondary-color: #34495e;
        --accent-color: #1abc9c;
        --accent-alt-color: #e67e22;
        --bg-primary: #f5f7fa;
        --bg-secondary: #ffffff;
        --text-primary: #2c3e50;
        --text-secondary: #7f8c8d;
        --glass-bg: rgba(255, 255, 255, 0.15);
        --glass-border: rgba(0, 0, 0, 0.1);
        --glass-shadow: 0 4px 16px rgba(0, 0, 0, 0.1);
        --neon-cyan: #00f5ff;
        --neon-purple: #d946ef;
        --neon-pink: #ff0080;
        --neon-blue: #0080ff;
        --neon-green: #00ff88;
    }
    
    * {
        font-family: 'Poppins', sans-serif;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    #MainMenu, footer, header {visibility: hidden;}
    
    .stApp {
        background: linear-gradient(135deg, #0a0e27 0%, #1a1f4d 50%, #0a0e27 100%);
        background-size: 200% 200%;
        animation: gradientShift 15s ease infinite;
    }
    
    @keyframes gradientShift {
        0% { background-position: 0% 50%; }
        50% { background-position: 100% 50%; }
        100% { background-position: 0% 50%; }
    }
    
    .custom-card {
        background: rgba(255, 255, 255, 0.12);
        backdrop-filter: blur(16px) saturate(180%);
        border: 1px solid rgba(255, 255, 255, 0.3);
        border-radius: 20px;
        padding: 1.75rem;
        margin-bottom: 1.25rem;
        box-shadow: var(--glass-shadow);
        transition: all 0.4s cubic-bezier(0.4, 0, 0.2, 1);
    }
    
    .custom-card:hover {
        transform: translateY(-5px) scale(1.02);
        border-color: var(--neon-cyan);
        box-shadow: 0 12px 40px rgba(0, 245, 255, 0.3);
        background: rgba(255, 255, 255, 0.15);
    }
    
    .result-container {
        background: rgba(0, 255, 136, 0.15);
        backdrop-filter: blur(20px);
        border: 2px solid var(--neon-green);
        border-radius: 20px;
        padding: 2.5rem;
        margin: 2rem 0;
        box-shadow: 0 0 40px rgba(0, 255, 136, 0.3);
        animation: pulse 3s ease-in-out infinite;
    }
    
    .result-title {
        color: #ffffff !important;
        font-size: 1.75rem;
        font-weight: 700;
        margin-bottom: 1rem;
        text-shadow: 0 0 20px rgba(0, 255, 136, 0.5);
    }
    
    h1, h2, h3 {
        color: #ffffff !important;
    }
    
    p, label, span, div {
        color: #e2e8f0 !important;
    }
    
    [data-testid="stSidebar"] {
        display: none !important;
    }
    
    .stButton button {
        background: linear-gradient(135deg, var(--neon-cyan), var(--neon-blue));
        color: #0a0e27;
        font-weight: 700;
        border: none;
        border-radius: 16px;
        padding: 0.875rem 2rem;
        box-shadow: 0 0 20px rgba(0, 245, 255, 0.4);
    }
    
    .stButton button:hover {
        transform: translateY(-3px) scale(1.05);
        box-shadow: 0 0 40px rgba(0, 245, 255, 0.6);
    }
    
    [data-testid="column"] button[kind="secondary"] {
        background: rgba(255, 255, 255, 0.06) !important;
        color: #e2e8f0 !important;
        border: 1.5px solid rgba(255, 255, 255, 0.15) !important;
        font-weight: 600 !important;
        border-radius: 12px !important;
    }
    
    [data-testid="column"] button[kind="secondary"]:hover {
        background: rgba(0, 245, 255, 0.12) !important;
        border-color: var(--neon-cyan) !important;
        color: var(--neon-cyan) !important;
        box-shadow: 0 0 25px rgba(0, 245, 255, 0.4) !important;
    }
</style>
"""