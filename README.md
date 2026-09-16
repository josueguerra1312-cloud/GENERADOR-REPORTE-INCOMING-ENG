# Incoming Engine Report Generator — V2

Generador de reportes fotográficos **ENG INCOMING** en PowerPoint, construido para conservar una apariencia técnica y ordenada inspirada en el reporte de referencia proporcionado.

> **Nota de diseño:** la V2 reproduce características visuales observadas en el ejemplo (estructura técnica, encabezados, uso de magenta, espacios blancos, composición fotográfica y sección de discrepancias). No incorpora logotipos, imágenes o activos oficiales de Volaris.

## V2 incluye

- Interfaz Streamlit.
- PN, SN, QC y fecha.
- Fotografías independientes por sección.
- Hasta 4 fotografías por diapositiva en una composición 2×2.
- Creación automática de páginas adicionales.
- Aspect ratio preservado.
- Encabezado `ENG INCOMING REPORT S/N`.
- Títulos de sección.
- Estilo magenta inspirado en el documento de referencia.
- Panel especial para `DISCREPANCIES`.
- Hasta 5 discrepancias por reporte.
- Fotografías de discrepancias.
- Diapositiva final de reporte completado.
- Salida directa en PowerPoint.

## Estructura plana

```text
incoming-engine-report/
├── app.py
├── report_generator.py
├── requirements.txt
├── README.md
└── .gitignore
```

## Instalación local

Windows:

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

macOS/Linux:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

## Flujo de uso

1. Ejecuta `streamlit run app.py`.
2. Introduce PN, SN, QC y fecha.
3. Abre cada sección.
4. Carga las fotografías correspondientes.
5. Si existen discrepancias, activa el apartado correspondiente y agrega descripción, finding, referencia y fotografías.
6. Pulsa **GENERATE POWERPOINT**.
7. Descarga `INCOMING_<SN>.pptx`.

## Siguiente evolución recomendada

- Crear una plantilla `.pptx` corporativa editable.
- Incorporar el logo/activos oficiales proporcionados por el usuario.
- Ajustar posiciones y tipografías a una copia autorizada del formato corporativo.
- Añadir conversión PPTX → PDF.
- Añadir carga por carpetas.
- Agregar reordenamiento drag-and-drop de fotografías.
- Añadir numeración y control de páginas.
