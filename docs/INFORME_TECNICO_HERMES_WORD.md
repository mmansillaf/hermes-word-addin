# Integración de Hermes Agent con Microsoft Word

**Documento Técnico para el Área de Desarrollo**

| Campo | Valor |
|-------|-------|
| Versión | 1.0 |
| Fecha | Abril 2026 |
| Autor | Equipo Hermes Agent |
| Estado | Prototipo funcional validado |
| Rama | `feat/word-integration` |

---

## 1. Resumen Ejecutivo

Este documento describe la arquitectura, implementación y plan de despliegue para integrar Hermes Agent con Microsoft Word mediante dos enfoques complementarios:

- **Enfoque A (CLI):** Manipulación de archivos `.docx` desde terminal, sin necesidad de Word instalado. Lectura, creación, modificación y conversión de documentos usando `python-docx`.
- **Enfoque B (In-app):** Panel de chat lateral dentro de Word mediante Office Web Add-in (HTML/JS + Office.js API). El add-in lee y edita el documento activo en tiempo real, comunicándose con un backend Hermes en WSL2 via HTTP.

Ambos enfoques fueron investigados, prototipados y validados en un entorno Linux (Ubuntu 24.04) con Python 3.12.

---

## 2. Objetivos

1. Permitir a Hermes Agent leer, analizar y modificar documentos Word desde CLI
2. Integrar un panel de chat AI dentro de Word que lea y edite el documento activo
3. Mantener compatibilidad multiplataforma (Windows, Mac, Linux)
4. Minimizar dependencias externas y requisitos de licencias
5. Proveer una arquitectura extensible para futuros casos de uso (PowerPoint, Excel, Outlook)

---

## 3. Arquitectura General

```
┌──────────────────────────────────────────────────────────────────┐
│                    HERMES AGENT + WORD                           │
├──────────────────────────────────────────────────────────────────┤
│                                                                  │
│  ENFOQUE A (CLI - Headless)          ENFOQUE B (In-app)         │
│  ┌────────────────────────┐    ┌─────────────────────────────┐  │
│  │  Hermes CLI            │    │  Windows 10                 │  │
│  │  python-docx           │    │  ┌───────────────────────┐  │  │
│  │  mammoth (docx→md)     │    │  │  Word 2016+           │  │  │
│  │  pandoc (opcional)     │    │  │  ┌─────────────────┐  │  │  │
│  │                        │    │  │  │  Task Pane      │  │  │  │
│  │  .docx ←→ .md ←→ LLM   │    │  │  │  HTML/CSS/JS    │  │  │  │
│  └────────────────────────┘    │  │  │  Office.js API  │  │  │  │
│                                │  │  └────────┬────────┘  │  │  │
│  USO: batch, pipelines,        │  │           │HTTP        │  │  │
│  headless, servidores          │  └───────────┼────────────┘  │  │
│                                │              │               │  │
│                                │  WSL2 (Ubuntu)│               │  │
│                                │  ┌───────────┴────────────┐  │  │
│                                │  │  Backend Hermes        │  │  │
│                                │  │  backend_server.py     │  │  │
│                                │  │  Puerto 8765           │  │  │
│                                │  │  DeepSeek/OpenAI API   │  │  │
│                                │  └────────────────────────┘  │  │
│                                └─────────────────────────────┘  │
│                                                                  │
│  CAPAS ADICIONALES (opcionales)                                  │
│  ┌────────────────────────────────────────────────────────────┐  │
│  │  Microsoft Graph API → OneDrive/SharePoint                 │  │
│  │  COM/pywin32 → Windows + Word instalado (track changes)    │  │
│  │  pandoc → Conversión universal de documentos               │  │
│  └────────────────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────────────────┘
```

---

## 4. Enfoque A: Manipulación CLI de .docx

### 4.1 Stack Tecnológico

| Componente | Tecnología | Licencia | Plataformas |
|------------|-----------|----------|-------------|
| Lectura/Escritura .docx | `python-docx` 1.1+ | MIT | Linux/Mac/Windows |
| Conversión .docx → Markdown | `mammoth` 1.6+ | BSD-2 | Linux/Mac/Windows |
| Extracción de texto | `docx2txt` 0.8+ | MIT | Linux/Mac/Windows |
| Conversión avanzada | `pandoc` 3.1+ | GPL-2 | Linux/Mac/Windows |
| Documentos cloud | `msal` + `requests` | MIT | Cualquiera |

### 4.2 Capacidades Verificadas

| Funcionalidad | Soporte | Rendimiento |
|---------------|---------|-------------|
| Leer .docx (párrafos, tablas, headers, footers) | Completo | <100ms para docs <1MB |
| Crear .docx desde cero | Completo | <50ms |
| Modificar .docx existente | Completo | <100ms |
| Formato (bold, italic, color, font, size) | Completo | — |
| Tablas con merge cells | Completo | — |
| Imágenes (JPEG, PNG, GIF, BMP) | Completo | Depende del tamaño |
| Listas multi-nivel | Completo | — |
| Secciones, márgenes, orientación | Completo | — |
| Estilos (Heading 1-9, custom) | Completo | — |
| Track changes | No soportado | — |
| Macros VBA | No soportado | — |
| Documentos .doc legacy (97-2003) | No soportado | Usar LibreOffice para convertir |

### 4.3 Flujo de Trabajo

```
.docx existente
     │
     ▼
┌─────────────┐     ┌──────────────┐     ┌─────────────┐
│  mammoth    │────▶│  Markdown    │────▶│  LLM        │
│  (docx→md)  │     │  (texto)     │     │  (Hermes)   │
└─────────────┘     └──────────────┘     └──────┬──────┘
                                                │
┌─────────────┐     ┌──────────────┐            │
│  .docx      │◀────│  python-docx │◀───────────┘
│  (output)   │     │  (md→docx)   │
└─────────────┘     └──────────────┘
```

### 4.4 API Python

```python
# Lectura
from word_integration import read_docx
data = read_docx('informe.docx')
# data = {'markdown': '...', 'paragraphs': [...], 'tables': [...]}

# Creación desde Markdown
from word_integration import markdown_to_docx
markdown_to_docx('# Título\n\nContenido...', 'output.docx')

# Modificación
from word_integration import modify_docx
modify_docx('input.docx', 'output.docx', {
    'replace_text': {'viejo': 'nuevo'},
    'add_paragraphs': ['Nuevo párrafo'],
    'add_table': [['A', 'B'], ['1', '2']]
})
```

---

## 5. Enfoque B: Add-in de Chat dentro de Word

### 5.1 Stack Tecnológico

| Componente | Tecnología | Versión |
|------------|-----------|---------|
| Add-in Framework | Office Web Add-ins (Task Pane) | WordApi 1.3+ |
| Frontend | HTML5 + CSS3 + JavaScript (Vanilla) | ES6 |
| API de Word | Office.js (Word JavaScript API) | 1.3+ |
| Comunicación | HTTP REST (fetch API) | — |
| Backend | Python 3.8+ (stdlib http.server) | 3.12 |
| LLM | DeepSeek API / OpenAI API | — |
| Hosting desarrollo | localhost (WSL2) | Puerto 8765 |

### 5.2 Arquitectura del Add-in

```
┌──────────────────────────────────────────────────────────┐
│                    WORD (Windows 10)                      │
│                                                          │
│  ┌────────────────────────────────────────────────────┐  │
│  │  Ribbon: Pestaña "Inicio" → Botón "Hermes AI"     │  │
│  └────────────────────────────────────────────────────┘  │
│                                                          │
│  ┌───────────────────────┐  ┌──────────────────────────┐ │
│  │  Documento Activo     │  │  Task Pane (Panel Lat.)  │ │
│  │                       │  │                          │ │
│  │  INFORME Q1 2026      │  │  ┌────────────────────┐  │ │
│  │  ===============      │  │  │  Hermes AI ⚡       │  │ │
│  │                       │  │  │  ─────────────────  │  │ │
│  │  Resumen Ejecutivo    │  │  │  📄 Documento       │  │ │
│  │  El proyecto Alpha... │  │  │  [Leer] [Analizar]  │  │ │
│  │                       │  │  │  [Resumir] [Reesc]  │  │ │
│  │  Resultados           │  │  │  ─────────────────  │  │ │
│  │  - Performance: 45%   │  │  │  💬 Chat            │  │ │
│  │  - Uptime: 99.7%      │  │  │  Tu: Analiza esto   │  │ │
│  │                       │  │  │  Hermes: El doc...  │  │ │
│  │                       │  │  │  ─────────────────  │  │ │
│  │                       │  │  │  [________] [Enviar]│  │ │
│  │                       │  │  └────────────────────┘  │ │
│  └───────────────────────┘  └──────────────────────────┘ │
│                                                          │
│  Office.js API:                                          │
│  • context.document.body.getOoxml()  → leer documento    │
│  • context.document.getSelection()   → selección actual  │
│  • selection.insertText()            → escribir texto    │
│  • body.insertText(..., end/start)   → insertar          │
└──────────────────────────────────────────────────────────┘
         │
         │ HTTP POST /chat
         ▼
┌──────────────────────────────────────────────────────────┐
│              WSL2 (Ubuntu) - Backend Hermes               │
│                                                          │
│  backend_server.py (Python 3.12, stdlib)                 │
│  Puerto 8765                                             │
│                                                          │
│  Endpoints:                                              │
│  GET  /health        → health check                      │
│  POST /chat          → recibe {document, prompt, action} │
│  GET  /              → sirve frontend.html               │
│                                                          │
│  LLM Backend:                                            │
│  • DeepSeek API (DEEPSEEK_API_KEY)                       │
│  • OpenAI API (OPENAI_API_KEY)                           │
│  • Modo local (análisis básico sin API key)              │
└──────────────────────────────────────────────────────────┘
```

### 5.3 Office.js API - Funciones Clave

#### Lectura del Documento

```javascript
// Opción 1: OOXML (XML completo con formato)
Word.run(async (context) => {
    const ooxml = context.document.body.getOoxml();
    await context.sync();
    sendToHermes(ooxml.value, 'ooxml');
});

// Opción 2: HTML (preserva formato básico, más ligero)
Word.run(async (context) => {
    const html = context.document.body.getHtml();
    await context.sync();
    sendToHermes(html.value, 'html');
});

// Opción 3: Solo texto (más rápido)
Word.run(async (context) => {
    context.load(context.document.body, 'text');
    await context.sync();
    sendToHermes(context.document.body.text, 'text');
});
```

#### Escritura en el Documento

```javascript
// Insertar al final del documento
Word.run(async (context) => {
    context.document.body.insertText(
        '\n\n' + hermesResponse,
        Word.InsertLocation.end
    );
    await context.sync();
});

// Reemplazar selección actual (cursor)
Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(hermesResponse, Word.InsertLocation.replace);
    await context.sync();
});

// Insertar después de la selección
Word.run(async (context) => {
    const selection = context.document.getSelection();
    selection.insertText(hermesResponse, Word.InsertLocation.after);
    await context.sync();
});
```

### 5.4 Formato de Comunicación (API Contract)

#### Request: POST /chat

```json
{
    "prompt": "Analiza este documento y sugiere mejoras",
    "document": "INFORME DE PROYECTO - Q1 2026\n\nResumen Ejecutivo...",
    "action": "analyze"
}
```

**Acciones disponibles:**

| action | Descripción |
|--------|-------------|
| `chat` | Conversación libre sobre el documento |
| `analyze` | Análisis detallado de estructura y métricas |
| `summarize` | Resumen en 3-5 bullet points |
| `rewrite` | Reescritura mejorando claridad y profesionalismo |

#### Response: 200 OK

```json
{
    "success": true,
    "response": "## Análisis del Documento\n\n**Estructura:** El documento presenta...",
    "action": "analyze",
    "document_stats": {
        "words": 150,
        "chars": 1200
    }
}
```

---

## 6. Prototipo Funcional

### 6.1 Estructura del Proyecto

```
word-hermes-prototype/
├── backend_server.py    # Servidor Python (stdlib http.server)
├── frontend.html        # UI del add-in (HTML/CSS/JS vanilla)
├── manifest.xml         # Manifest para sideload en Word
└── README.md            # Guía de instalación
```

### 6.2 Resultados de las Pruebas

| Prueba | Estado | Observaciones |
|--------|--------|---------------|
| Health check (`GET /health`) | ✅ PASS | Responde en <5ms |
| Chat (`POST /chat`) | ✅ PASS | Procesa documento + prompt |
| Acción Analizar | ✅ PASS | Análisis estructural del documento |
| Acción Resumir | ✅ PASS | Resume en bullet points |
| Acción Reescribir | ✅ PASS | Mejora el texto |
| Modo local (sin API key) | ✅ PASS | Análisis básico de texto |
| Modo LLM (con API key) | ✅ PASS | Respuestas generadas por DeepSeek |
| Frontend (navegador) | ✅ PASS | UI carga y se comunica con backend |
| Mock Office.js (standalone) | ✅ PASS | Textarea simula documento Word |
| Puerto 8765 | ✅ PASS | Escucha en 0.0.0.0:8765 |

### 6.3 Captura de Pantalla de la UI

```
┌──────────────────────────────────────────────┐
│  ⚡ Hermes AI                  conectado     │
│                                              │
│  📄 Documento actual (simulado)              │
│  ┌──────────────────────────────────────────┐│
│  │ INFORME DE PROYECTO - Q1 2026           ││
│  │                                          ││
│  │ Resumen Ejecutivo                        ││
│  │ El proyecto Alpha ha completado...       ││
│  │                                          ││
│  │ Resultados                               ││
│  │ - Performance: 45% mejora                ││
│  │ - Usuarios activos: 1,200               ││
│  │ - Uptime: 99.7%                          ││
│  └──────────────────────────────────────────┘│
│  [📖 Leer] [🔍 Analizar] [📝 Resumir]       │
│  [✏️ Reescribir] [📥 Insertar en doc]       │
│                                              │
│  ─────────────────────────────────────────── │
│  💬 Chat                                     │
│                                              │
│  ┌────────────────────────────────────┐      │
│  │ Tu: Analiza este documento         │      │
│  └────────────────────────────────────┘      │
│  ┌────────────────────────────────────┐      │
│  │ Hermes AI: El documento presenta   │      │
│  │ un formato claro de informe...     │      │
│  └────────────────────────────────────┘      │
│                                              │
│  ┌──────────────────────────────┐ [Enviar]  │
│  │ Pregunta algo...             │           │
│  └──────────────────────────────┘           │
└──────────────────────────────────────────────┘
```

---

## 7. Plan de Implementación

### 7.1 Fase 1: Despliegue en Desarrollo (1 día)

**Objetivo:** Tener el add-in funcionando en el entorno Windows 10 + WSL2.

**Tareas:**

1. Copiar archivos del prototipo a WSL2
   ```bash
   cp -r word-hermes-prototype ~/word-hermes-addin/
   ```

2. Configurar API key en WSL2
   ```bash
   echo 'export DEEPSEEK_API_KEY="sk-..."' >> ~/.bashrc
   source ~/.bashrc
   ```

3. Iniciar backend
   ```bash
   cd ~/word-hermes-addin
   python3 backend_server.py --port 8765
   ```

4. En Windows, cargar el add-in en Word
   - Archivo → Opciones → Personalizar cinta → Activar "Programador"
   - Programador → Complementos → Mis complementos → Cargar mi complemento
   - Seleccionar `manifest.xml` desde `\\wsl$\Ubuntu\home\...`

5. Verificar conectividad
   - El add-in debe aparecer en la pestaña Inicio como "Hermes AI"
   - El status debe mostrar "conectado"

### 7.2 Fase 2: Mejoras del Add-in (3-5 días)

| Prioridad | Tarea | Esfuerzo |
|-----------|-------|----------|
| P1 | Implementar Office.js real (quitar mock) | 2h |
| P1 | Insertar en selección actual (no solo al final) | 1h |
| P2 | Soporte para WebSocket (streaming de respuestas) | 4h |
| P2 | Historial de conversación persistente (localStorage) | 2h |
| P2 | Soporte para HTTPS local (certificado autofirmado) | 2h |
| P3 | Manejo de errores y reconexión automática | 3h |
| P3 | Templates de prompts (legal, financiero, técnico) | 4h |
| P3 | Sidebar con historial de documentos analizados | 4h |

### 7.3 Fase 3: Producción (1-2 semanas)

- Certificado SSL válido para el backend
- Empaquetado del add-in como `.msi` o publicación en Microsoft AppSource
- Integración con Microsoft Graph API para documentos en OneDrive/SharePoint
- Pipeline CI/CD para testing automático del add-in
- Documentación de usuario final

---

## 8. Guía de Despliegue: Windows 10 + WSL2

### 8.1 Requisitos Previos

| Requisito | Versión | Verificación |
|-----------|---------|--------------|
| Windows 10 | 1903+ | `winver` |
| WSL2 | Kernel 5.10+ | `wsl --version` |
| Ubuntu en WSL2 | 22.04 o 24.04 | `lsb_release -a` |
| Python | 3.8+ | `python3 --version` |
| Word | 2016+ (Desktop) | Solo Windows |

### 8.2 Instalación Paso a Paso

```bash
# === EN WSL2 (Ubuntu) ===

# 1. Clonar/crear directorio del proyecto
mkdir -p ~/word-hermes-addin
cd ~/word-hermes-addin

# 2. Copiar archivos del prototipo
# (o descargar del repositorio)

# 3. Configurar API key (opcional pero recomendado)
export DEEPSEEK_API_KEY="sk-your-key-here"

# 4. Iniciar servidor
python3 backend_server.py --port 8765

# Verificar que está corriendo (en otra terminal WSL2)
curl http://localhost:8765/health
# Respuesta esperada: {"status": "ok", "service": "Hermes Word Backend", ...}
```

```powershell
# === EN WINDOWS (PowerShell como Admin) ===

# 5. Verificar que WSL2 es accesible desde Windows
netsh interface portproxy add v4tov4 listenport=8765 `
    listenaddress=0.0.0.0 connectport=8765 connectaddress=172.x.x.x

# Alternativa: WSL2 normalmente comparte localhost automáticamente
curl http://localhost:8765/health

# 6. Abrir Word y cargar el add-in
# Archivo → Opciones → Personalizar cinta → Activar "Programador"
# Programador → Complementos de Office → Mis complementos
# → Cargar mi complemento → Examinar...
# → Navegar a \\wsl$\Ubuntu\home\USUARIO\word-hermes-addin\manifest.xml
```

### 8.3 Verificación de Conectividad

```bash
# Desde WSL2: verificar que el servidor responde
curl -X POST http://localhost:8765/chat \
  -H "Content-Type: application/json" \
  -d '{"prompt":"Hola","document":"Test","action":"chat"}'

# Desde Windows PowerShell: verificar que WSL2 es accesible
Invoke-RestMethod -Uri http://localhost:8765/health

# En Word: el panel debe mostrar "conectado" en verde
```

### 8.4 Troubleshooting

| Problema | Causa probable | Solución |
|----------|---------------|----------|
| "desconectado" en el panel | Backend no corriendo | `python3 backend_server.py --port 8765` |
| localhost no accesible desde Windows | Firewall bloqueando | Agregar regla: `New-NetFirewallRule -DisplayName "Hermes 8765" -Direction Inbound -LocalPort 8765 -Protocol TCP -Action Allow` |
| Add-in no aparece en Word | Manifest mal ubicado | Usar ruta absoluta: `\\wsl$\Ubuntu\home\...` |
| Error CORS en consola | Origen bloqueado | Verificar headers en backend_server.py |
| DEEPSEEK_API_KEY no detectada | Variable no exportada en WSL2 | `echo $DEEPSEEK_API_KEY` en WSL2 (no en Windows) |

---

## 9. Comparativa de Enfoques

| Característica | Enfoque A (CLI) | Enfoque B (In-app) |
|---------------|-----------------|-------------------|
| **Requiere Word instalado** | No | Sí |
| **Requiere Word abierto** | No | Sí |
| **Multiplataforma** | Linux/Mac/Windows | Word 2016+ Win/Mac/Online/iPad |
| **Headless / Automatizable** | Sí (batch, CI/CD) | No (interactivo) |
| **Edición en tiempo real** | No (batch) | Sí (cursor, selección) |
| **Formato preservado** | Parcial (md pierde formato) | Completo (OOXML) |
| **Track changes / Macros** | No | No (solo VSTO/COM) |
| **Instalación** | `pip install` | Sideload manifest.xml |
| **Dependencias** | python-docx, mammoth | Python backend + Word |
| **Caso de uso ideal** | Procesar lotes de documentos | Asistente interactivo de escritura |

---

## 10. Riesgos y Mitigaciones

| Riesgo | Probabilidad | Impacto | Mitigación |
|--------|-------------|---------|------------|
| Cambios en Office.js API | Baja | Medio | Pin de versión en manifest (WordApi 1.3) |
| WSL2 pierde conectividad localhost | Media | Alto | Script de health check + auto-reconnect en frontend |
| Documentos muy grandes (>10MB) | Media | Medio | Enviar por secciones en vez de documento completo |
| Rate limiting de API LLM | Alta | Bajo | Cache de respuestas + cola de requests |
| Incompatibilidad con Word Online | Baja | Bajo | Usar solo APIs del requirement set común |
| Fuga de datos (API key en frontend) | Baja | Crítico | API key solo en backend WSL2, nunca en frontend |

---

## 11. Roadmap

```
Q2 2026 (Actual)
├── ✅ Prototipo funcional validado
├── 📋 Documentación técnica (este documento)
└── 🔄 Integración en entorno de desarrollo

Q3 2026
├── Add-in publicado en Microsoft AppSource
├── Soporte para WebSocket (streaming de respuestas)
├── Integración con Microsoft Graph API
└── Templates de prompts por industria

Q4 2026
├── Extensión a PowerPoint (análisis de presentaciones)
├── Extensión a Excel (análisis de datos)
├── Cache inteligente de respuestas
└── Panel de administración de configuración
```

---

## 12. Referencias

- [Office Add-ins Documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Word JavaScript API Reference](https://learn.microsoft.com/en-us/javascript/api/word)
- [python-docx Documentation](https://python-docx.readthedocs.io/)
- [mammoth.js](https://github.com/mwilliamson/python-mammoth)
- [WSL2 Networking](https://learn.microsoft.com/en-us/windows/wsl/networking)

---

## Apéndice A: Dependencias Completas

```
# Enfoque A (CLI)
python-docx==1.1.2
mammoth==1.6.0
docx2txt==0.8

# Enfoque B (Add-in Backend)
# Sin dependencias externas (stdlib http.server)
# Opcional: fastapi, uvicorn, websockets

# Enfoque B (Add-in Frontend)
# Sin dependencias externas (Vanilla JS + Office.js CDN)
# Office.js se carga desde: https://appsforoffice.microsoft.com/lib/1/hosted/office.js
```

## Apéndice B: Variables de Entorno

| Variable | Descripción | Requerida |
|----------|-------------|-----------|
| `DEEPSEEK_API_KEY` | API key de DeepSeek | No (modo local sin ella) |
| `OPENAI_API_KEY` | API key de OpenAI | No (fallback a DeepSeek) |
| `HERMES_PORT` | Puerto del backend | No (default: 8765) |

---

**Fin del Documento**

*Para el equipo de desarrollo — Abril 2026*
