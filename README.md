# Incoming Engine Report Generator

Aplicación Python/Streamlit para generar reportes **ENG INCOMING** en PowerPoint a partir de fotografías organizadas por sección.

El diseño de las secciones iniciales está basado en el reporte de referencia suministrado para el motor **PW1133GA-JM / S/N P800341**.

## Características de esta V1

- Captura PN, SN, QC y fecha.
- Carga múltiples fotografías para cada sección.
- Mantiene la proporción original de cada fotografía.
- Distribuye hasta 4 fotografías por diapositiva.
- Crea diapositivas adicionales automáticamente cuando una sección contiene más fotografías.
- Genera `INCOMING_<SN>.pptx`.
- Permite descargar el PowerPoint desde la interfaz.

## Instalación

```bash
python -m venv .venv
```

Windows:

```bash
.venv\Scripts\activate
```

macOS/Linux:

```bash
source .venv/bin/activate
```

Instalar dependencias:

```bash
pip install -r requirements.txt
```

Ejecutar:

```bash
streamlit run app.py
```

## Estructura

```text
incoming-engine-report/
├── app.py
├── requirements.txt
├── README.md
├── src/
│   ├── __init__.py
│   └── report_generator.py
├── templates/
├── assets/
└── output/
```

## Próximas mejoras

1. Reproducir exactamente el diseño visual del PowerPoint corporativo.
2. Incorporar logo, encabezado y pie de página.
3. Agregar formulario de discrepancias con fotografías.
4. Permitir cargar carpetas completas por sección.
5. Generar PDF automáticamente.
6. Validar orientación, resolución y nombre de fotografías.
7. Crear una plantilla `.pptx` editable.
8. Incorporar numeración de páginas.
9. Agregar configuración para cambiar cantidad de fotos por página.
10. Preparar una versión ejecutable para Windows.
