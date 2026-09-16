# Incoming Engine Report Generator — V2.1

Generador de reportes fotográficos **ENG INCOMING** en PowerPoint.

La V2.1 corrige el problema de importación observado en Streamlit Cloud y mantiene todos los archivos en la raíz del repositorio.

## Estructura obligatoria

```text
generador-reporte-incoming-eng/
├── app.py
├── report_generator.py
├── requirements.txt
├── README.md
└── .gitignore
```

**Importante:** `app.py` y `report_generator.py` deben estar al mismo nivel.

## Corrección V2.1

`app.py` carga `report_generator.py` usando la ruta física del propio archivo:

```python
BASE_DIR = Path(__file__).resolve().parent
GENERATOR_PATH = BASE_DIR / "report_generator.py"
```

Esto evita depender del directorio de ejecución de Streamlit Cloud y proporciona un mensaje claro si el archivo no fue subido al repositorio.

## Funciones

- PN, SN, QC y fecha.
- Fotografías por sección.
- Hasta 4 fotografías por página.
- Páginas adicionales automáticas.
- Conservación del aspect ratio.
- Encabezado técnico.
- Diseño magenta inspirado en el reporte de referencia.
- Sección de discrepancias.
- Hasta 5 discrepancias.
- Fotografías de discrepancias.
- Finding / Observation.
- Diapositiva final de reporte.
- Descarga directa del PowerPoint.

## Instalación

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud

1. Sube los cinco archivos a la raíz de GitHub.
2. En Streamlit Cloud selecciona el repositorio.
3. Selecciona `app.py` como **Main file path**.
4. Haz redeploy/reboot.
5. La aplicación generará `INCOMING_<SN>.pptx`.

## Nota visual

El diseño es una aproximación basada en el reporte proporcionado. No contiene logotipos, imágenes ni activos oficiales de Volaris.
