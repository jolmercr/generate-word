# GenerarWord

Proyecto para generar documentos de Word usando Python.

## Configuración del Entorno de Desarrollo

### 1. Activar el entorno virtual

#### En macOS/Linux:
```bash
source venv/bin/activate
```

#### En Windows:
```bash
venv\Scripts\activate
```

### 2. Instalar dependencias

Una vez activado el entorno virtual, instala las dependencias:

```bash
pip install -r requirements.txt
```

### 3. Desactivar el entorno virtual

Cuando termines de trabajar:

```bash
deactivate
```

## Estructura del Proyecto

```
GenerarWord/
├── venv/                 # Entorno virtual (no se versiona)
├── requirements.txt      # Dependencias del proyecto
├── .gitignore           # Archivos ignorados por git
├── README.md            # Este archivo
└── src/                 # Código fuente (crear según necesites)
```

## Uso

```python
from docx import Document

# Crear un nuevo documento
doc = Document()
doc.add_heading('Mi Documento', 0)
doc.add_paragraph('Contenido del documento...')
doc.save('mi_documento.docx')
```

## Dependencias Principales

- **python-docx**: Librería para crear y modificar documentos de Word (.docx)
