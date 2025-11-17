# SmartMind - Sistema de Gestión de Documentación

Sistema automatizado para la gestión de documentación de convocatorias de formación.

## Estructura del Proyecto
```
Doc_Convocatoria/
├── app.py                 # Aplicación principal
├── requirements.txt       # Dependencias
├── README.md             # Documentación
├── assets/               # Recursos (logos, imágenes)
├── config/               # Configuración general
├── styles/               # Estilos CSS personalizados
├── utils/                # Utilidades y funciones auxiliares
├── components/           # Componentes reutilizables
└── sections/             # Secciones de la aplicación
```

## Instalación

1. Clonar el repositorio
2. Crear entorno virtual:
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

3. Instalar dependencias:
```bash
pip install -r requirements.txt
```

4. Ejecutar la aplicación:
```bash
streamlit run app.py
```

## Características

- **Captación**: Gestión de documentos de inscripción
- **Formación Empresa Inicio**: Control de inicio de formación
- **Formación Empresa Fin**: Procesamiento de documentación final
- **Evaluación**: Generación automática de actas y certificaciones
- **Cierre Mes**: Documentación de cierre mensual

## Tecnologías

- Python 3.8+
- Streamlit
- Pandas
- OpenPyXL
- python-docx
- Pillow
- pytesseract
- PyPDF2

## Autor

Arancha Fernández Fernández
