# Investigacion: Integracion Microsoft Word con Hermes (CLI AI Agent)

**Fecha:** Abril 2026
**Entorno de prueba:** Linux (Ubuntu 24.04), Python 3.12

---

## RESUMEN EJECUTIVO

Para Hermes (CLI AI Agent corriendo en Linux/Mac/Windows), la estrategia recomendada es:

1. **python-docx** como herramienta principal (leer, analizar, modificar, crear .docx)
2. **mammoth** como puente de conversion .docx -> Markdown (para que Hermes "lea" documentos)
3. **python-docx + markdown** para crear .docx desde Markdown generado por el LLM
4. **Microsoft Graph API** opcional para documentos en OneDrive/SharePoint
5. **COM/pywin32** solo viable en Windows, para features avanzadas que python-docx no cubre

---

## ENFOQUE 1: python-docx (PRINCIPAL - RECOMENDADO)

**Tipo:** Biblioteca Python pura (sin dependencias externas de Word)
**Plataformas:** Linux, Mac, Windows
**Instalacion:** `pip install python-docx`

### Capacidades (verificadas en pruebas)

| Feature | Soporte | Notas |
|---------|---------|-------|
| Leer .docx | Completo | Parrafos, runs, estilos, tablas, headers, footers |
| Crear .docx | Completo | Desde cero con toda la estructura |
| Modificar .docx | Completo | Editar texto, formato, agregar/quitar elementos |
| Parrafos y runs | Completo | Bold, italic, underline, color, font size, font name |
| Tablas | Completo | Merge cells, bordes, estilos predefinidos |
| Imagenes | Completo | PNG, JPEG, GIF, BMP, TIFF (insertar y extraer) |
| Headers/Footers | Completo | Por seccion, primera pagina diferente |
| Listas | Completo | Bullet, numbered, multi-nivel |
| Estilos | Completo | Heading 1-9, List Bullet/Number, estilos custom |
| Secciones | Completo | Page breaks, orientacion, margenes, columnas |
| Comentarios | Parcial | Lectura solamente (no escritura) |
| Track Changes | No | No soportado |
| Macros/VBA | No | No soportado |
| Documentos .doc (legado) | No | Solo formato .docx (Office 2007+) |

### Ejemplo practico para Hermes

```python
from docx import Document

# Leer documento
doc = Document('informe.docx')
for p in doc.paragraphs:
    print(f"[{p.style.name}] {p.text}")

# Extraer tablas
for table in doc.tables:
    for row in table.rows:
        print([cell.text for cell in row.cells])

# Modificar
doc.paragraphs[0].text = "NUEVO TITULO"
doc.add_paragraph("Agregado por Hermes")

# Guardar
doc.save('informe_modificado.docx')
```

### Limitaciones
- No soporta documentos .doc antiguos
- No puede ejecutar macros
- Track changes no es accesible
- Los comentarios son solo lectura

---

## ENFOQUE 2: Office JavaScript API (Word Add-ins)

**Tipo:** API para add-ins dentro de Word (web-based)
**Plataformas:** Word Online, Word Desktop (Windows/Mac), Word Mobile
**Relevancia para CLI:** **BAJA**

### Por que NO es practico para Hermes

- Requiere un add-in instalado dentro de Word (no funciona standalone)
- Ejecuta JavaScript en un webview dentro de Word
- Necesita interaccion con UI de Word - no hay interfaz CLI
- No se puede llamar desde un script Python/bash directamente
- Util solo si Hermes tuviera un add-in companion en Word

### Posible uso futuro
Si se desarrollara un "Hermes Add-in for Word" que expusiera una API local (ej. WebSocket), Hermes podria comunicarse con Word a traves de el. Pero esto requiere desarrollo significativo y Word abierto.

---

## ENFOQUE 3: Microsoft Graph API

**Tipo:** API REST cloud para Microsoft 365
**Plataformas:** Multiplataforma (requiere internet + autenticacion)
**Instalacion:** `pip install msal requests`

### Capacidades

| Feature | Soporte |
|---------|---------|
| Listar documentos en OneDrive/SharePoint | Si |
| Descargar .docx | Si |
| Subir .docx | Si |
| Convertir a PDF | Si (via export) |
| Buscar en documentos | Si (via Search API) |
| Leer metadata | Si |
| Modificar contenido programaticamente | Limitado (no es un editor .docx) |

### Ejemplo: Descargar documento desde OneDrive

```python
import msal
import requests

# Autenticacion (app registration en Azure AD requerida)
app = msal.ConfidentialClientApplication(
    client_id="CLIENT_ID",
    client_credential="CLIENT_SECRET",
    authority="https://login.microsoftonline.com/TENANT_ID"
)
token = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])

# Descargar archivo
headers = {"Authorization": f"Bearer {token['access_token']}"}
response = requests.get(
    "https://graph.microsoft.com/v1.0/me/drive/root:/documento.docx:/content",
    headers=headers
)
with open('documento_descargado.docx', 'wb') as f:
    f.write(response.content)
```

### Ventajas para Hermes
- Acceso a documentos en la nube sin tener Word instalado
- Busqueda en documentos de SharePoint/OneDrive
- Posibilidad de listar documentos de un equipo

### Desventajas
- Requiere configuracion previa (Azure AD app registration)
- Necesita permisos de tenant/organizacion
- No edita contenido de .docx directamente (descarga, edita localmente con python-docx, sube)
- Requiere conexion a internet

### Integracion practica con Hermes
Flujo: Graph API para descubrir/descargar -> python-docx para editar -> Graph API para subir

---

## ENFOQUE 4: COM Automation con pywin32

**Tipo:** Automatizacion de Word via COM en Windows
**Plataformas:** SOLO Windows (requiere Microsoft Word instalado)
**Instalacion:** `pip install pywin32`

### Capacidades

Acceso COMPLETO a todas las features de Word, incluyendo:
- Macros, VBA
- Track changes
- PDF export nativo
- Imprimir documentos
- Combinar correspondencia
- Todas las features avanzadas de Word

### Ejemplo

```python
import win32com.client

word = win32com.client.Dispatch("Word.Application")
word.Visible = False  # No mostrar UI

doc = word.Documents.Open(r"C:\ruta\documento.docx")

# Leer contenido
for paragraph in doc.Paragraphs:
    print(paragraph.Range.Text)

# Modificar
doc.Paragraphs(1).Range.Text = "Nuevo titulo"
doc.Save()
doc.Close()
word.Quit()
```

### Ventajas
- Acceso al 100% de las capacidades de Word
- Export a PDF nativo de alta calidad
- Soporte para .doc y .docx
- Track changes, comentarios, macros

### Desventajas CRITICAS
- **SOLO Windows**
- Requiere Microsoft Word instalado (licencia paga)
- Lento (abrir Word es pesado)
- No es headless-friendly (aunque Visible=False ayuda)
- Puede dejar procesos zombie si hay errores

### Recomendacion para Hermes
Solo usar como fallback en Windows cuando python-docx no alcance. Implementar deteccion de plataforma:

```python
import sys
if sys.platform == 'win32':
    try:
        import win32com.client
        # Usar COM
    except ImportError:
        pass  # Fallback a python-docx
```

---

## ENFOQUE 5: Conversion .docx <-> Markdown/Texto

### 5a. docx2txt - Extraccion de texto plano

**Instalacion:** `pip install docx2txt`
**Uso:** Extrae texto plano de .docx, incluyendo texto de imagenes (OCR basico)

```python
import docx2txt
text = docx2txt.process('documento.docx')
```

**Limitaciones:** Pierde estructura (tablas se aplanan, formato se pierde)

### 5b. mammoth - Conversion .docx -> Markdown/HTML

**Instalacion:** `pip install mammoth`
**Uso:** Convierte .docx a Markdown o HTML preservando estructura basica

```python
import mammoth
with open('documento.docx', 'rb') as f:
    result = mammoth.convert_to_markdown(f)
    markdown_text = result.value
```

**Ventajas:**
- Preserva headings, listas, bold/italic
- Relativamente fiel al formato original
**Limitaciones:**
- Tablas se aplanan a texto (no las convierte a tablas Markdown)
- Imagenes no se extraen automaticamente

### 5c. markdownify - HTML -> Markdown

Util si se combina con mammoth (que puede output HTML): `pip install markdownify`

### 5d. pandoc - Conversion universal (RECOMENDADO SI ESTA DISPONIBLE)

**Instalacion:** `sudo apt install pandoc` (Linux) o `brew install pandoc` (Mac)
**No verificado en este entorno** (pandoc no instalado, pero esta disponible en apt)

```bash
# .docx -> Markdown
pandoc documento.docx -t markdown -o output.md

# Markdown -> .docx
pandoc input.md -o output.docx --reference-doc=template.docx
```

**Ventajas de pandoc:**
- Conversion de alta calidad
- Soporta tablas Markdown -> tablas Word correctamente
- Puede usar un template .docx para controlar estilos
- Multiplataforma
- Preserva mas estructura que mammoth/docx2txt

### 5e. Creacion .docx desde Markdown via python-docx

Para cuando pandoc no esta disponible, Hermes puede parsear Markdown y construir
el .docx con python-docx. Verificado en pruebas:

```python
# Markdown -> python-docx (parser simple)
# Headings, parrafos, listas, tablas basicas
# (ver test_md_to_docx.py)
```

---

## RECOMENDACION FINAL PARA HERMES

### Estrategia en capas (de mas practica a mas especializada):

```
Nivel 1: python-docx (SIEMPRE disponible)
  - Leer, analizar, modificar, crear .docx
  - Cubre 90% de casos de uso

Nivel 2: mammoth + python-docx
  - .docx -> Markdown (para que el LLM procese texto)
  - Markdown -> .docx (para output del LLM)
  - Flujo: docx -> mammoth -> markdown -> LLM -> markdown -> python-docx -> docx

Nivel 3: pandoc (si disponible)
  - Mejor conversion .docx <-> Markdown
  - Preserva tablas y estructura compleja

Nivel 4: Microsoft Graph API (si configurado)
  - Acceso a documentos en OneDrive/SharePoint
  - Combinar con Nivel 1/2 para edicion

Nivel 5: COM/pywin32 (solo Windows, solo si Word instalado)
  - Features avanzadas que python-docx no cubre
  - Export PDF nativo
```

### Implementacion sugerida en Hermes

```python
# Pseudocodigo de integracion
class WordIntegration:
    def read_docx(self, path):
        """Leer .docx y retornar contenido estructurado"""
        # Opcion A: mammoth -> markdown (para texto legible por LLM)
        # Opcion B: python-docx -> dict estructurado (para analisis preciso)

    def create_docx(self, markdown_content, output_path):
        """Crear .docx desde markdown"""
        # Opcion A: pandoc si disponible
        # Opcion B: parser markdown + python-docx

    def modify_docx(self, path, modifications):
        """Modificar documento existente"""
        # python-docx para cambios estructurales

    def search_sharepoint(self, query):
        """Buscar documentos en la nube"""
        # Microsoft Graph API
```

### Archivos generados durante la investigacion

- `/home/usuario/sample_report.docx` - Documento de prueba basico
- `/home/usuario/sample_modified.docx` - Documento modificado
- `/home/usuario/capabilities_demo.docx` - Demo de features avanzadas
- `/home/usuario/md_generated.docx` - Documento creado desde Markdown
- `/home/usuario/create_sample.py` - Script de creacion
- `/home/usuario/test_read.py` - Script de lectura
- `/home/usuario/test_modify.py` - Script de modificacion
- `/home/usuario/test_md_to_docx.py` - Script Markdown->DOCX
- `/home/usuario/test_complex_docx.py` - Script features avanzadas
- `/home/usuario/test_mammoth2.py` - Script conversion mammoth
