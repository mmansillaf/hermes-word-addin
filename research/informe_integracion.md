# Hermes Agent ↔ Microsoft Word: Investigacion de Integracion

**Informe de Arquitectura y Estrategia**
**Fecha:** 29 de abril de 2026
**Autor:** Hermes Agent (DeepSeek v4-pro)
**Version:** 1.0

---

## Indice

1. [Resumen Ejecutivo](#1-resumen-ejecutivo)
2. [Estado del Arte: Ecosistema Word + AI](#2-estado-del-arte-ecosistema-word--ai)
3. [Enfoques de Integracion](#3-enfoques-de-integracion)
4. [MCP Servers para Word: Analisis Comparativo](#4-mcp-servers-para-word-analisis-comparativo)
5. [Arquitectura Propuesta](#5-arquitectura-propuesta)
6. [Plan de Implementacion](#6-plan-de-implementacion)
7. [Competitive Landscape](#7-competitive-landscape)
8. [Riesgos y Desafios](#8-riesgos-y-desafios)
9. [Conclusiones y Recomendaciones](#9-conclusiones-y-recomendaciones)
10. [Referencias](#10-referencias)

---

## 1. Resumen Ejecutivo

### Objetivo
Integrar Hermes Agent como un panel de chat lateral dentro de Microsoft Word, permitiendo interaccion bidireccional: leer el documento activo, procesarlo con LLMs, e insertar respuestas directamente en el documento.

### Hallazgo Principal
**Existen 3 enfoques viables y 6 MCP servers open source** que pueden aprovecharse. La opcion recomendada combina un **Office Web Add-in (HTML/JS + Office.js)** como frontend dentro de Word, con un **backend Python (FastAPI + WebSocket)** que expone el LLM via Hermes Agent, complementado por un **MCP Server** para operaciones avanzadas sobre .docx.

### Metricas Clave
- **Tiempo estimado de desarrollo:** 3-5 dias (MVP), 2-3 semanas (produccion)
- **Dificultad:** Media
- **Cobertura multiplataforma:** Windows, Mac, Web, iPad (Office.js)
- **Proyectos open source existentes:** 6 MCP servers, 0 add-ins completos con chat+lectura/escritura

---

## 2. Estado del Arte: Ecosistema Word + AI

### 2.1 Panorama General

El ecosistema de integracion Word+AI esta fragmentado en 3 capas:

```
┌────────────────────────────────────────────────────────────┐
│  CAPA 1: IN-APP (Word Add-ins)                             │
│  Office.js API · VSTO/COM · Task Pane · Dialog             │
│  ─────────────────────────────────────────────              │
│  CAPA 2: HEADLESS (sin Word)                               │
│  python-docx · pandoc · mammoth · OOXML puro               │
│  ─────────────────────────────────────────────              │
│  CAPA 3: PROTOCOLO (MCP)                                   │
│  MCP Servers · Stdio/SSE · FastMCP                         │
└────────────────────────────────────────────────────────────┘
```

### 2.2 Microsoft Copilot (Competitive Baseline)

Microsoft Copilot en Word (2024+) es el referente de mercado:
- Panel lateral con chat contextual al documento
- Capacidades: draft, rewrite, summarize, analyze
- Usa modelos GPT-4/GPT-4o via Azure OpenAI
- Integracion profunda con Microsoft Graph (OneDrive/SharePoint)
- **Limitacion:** Solo disponible con Microsoft 365 Copilot ($30/usuario/mes), no es open source, no permite custom models

### 2.3 Proyectos Open Source Relevantes

| Proyecto | Estrellas | Enfoque | Estado |
|----------|-----------|---------|--------|
| GongRzhe/Office-Word-MCP-Server | ★1,902 | MCP + python-docx | Archivado Mar 2026 |
| OfficeMCP/OfficeMCP | ★78 | COM/Windows + Python | Activo |
| PsychQuant/che-word-mcp | ★3 | Swift puro + OOXML | Activo (v3.13.5) |
| ForLegalAI/mcp-ms-office-documents | ★25 | MCP + python-docx/pptx | Activo |
| vAirpower/macos-office365-mcp-server | ★13 | AppleScript + python-docx | Activo |
| ecator/cs-office-mcp-server | ★? | C# MCP server | Activo |

**Hallazgo critico:** Ninguno de estos proyectos implementa un panel de chat DENTRO de Word. Son servidores MCP que manipulan archivos .docx, pero sin interfaz de usuario en Word.

---

## 3. Enfoques de Integracion

### 3.1 Enfoque A: Office Web Add-in (HTML/JS + Office.js)
**Recomendado para el MVP**

```
┌──────────────────────────────────────────────────────┐
│  WORD (Windows / Mac / Online / iPad)                 │
│  ┌────────────────────────────────────────────────┐  │
│  │  TASK PANE (panel lateral HTML/CSS/JS)         │  │
│  │  ┌──────────────────────────────────────────┐  │  │
│  │  │  Chat UI                                  │  │  │
│  │  │  - Historial de conversacion              │  │  │
│  │  │  - Input de texto                        │  │  │
│  │  │  - Botones: Leer/Analizar/Resumir/       │  │  │
│  │  │    Reescribir/Insertar                   │  │  │
│  │  └──────────────────────────────────────────┘  │  │
│  │           ↕ Office.js API                       │  │
│  │  ┌──────────────────────────────────────────┐  │  │
│  │  │  Documento activo                         │  │  │
│  │  │  - body.getOoxml() / getHtml() / text     │  │  │
│  │  │  - getSelection().insertText()            │  │  │
│  │  │  - paragraphs, tables, content controls   │  │  │
│  │  └──────────────────────────────────────────┘  │  │
│  └────────────────────────────────────────────────┘  │
│              ↕ WebSocket / SSE (streaming)            │
│  ┌────────────────────────────────────────────────┐  │
│  │  BACKEND HERMES (local/WSL/Docker)             │  │
│  │  - FastAPI + WebSocket                         │  │
│  │  - Recibe OOXML → convierte a texto limpio     │  │
│  │  - LLM (DeepSeek v4, GPT, Claude)              │  │
│  │  - Devuelve respuesta + instrucciones de       │  │
│  │    insercion                                   │  │
│  └────────────────────────────────────────────────┘  │
└──────────────────────────────────────────────────────┘
```

**Ventajas:**
- Multiplataforma (Word 2016+ en Windows, Mac, Web, iPad)
- Desarrollo web estandar (HTML/CSS/JS)
- Office.js API oficial de Microsoft
- No requiere instalacion de software adicional en la maquina

**Desventajas:**
- Requiere HTTPS para desarrollo (cert autofirmado o ngrok)
- Office.js es asincrono (modelo request/context.sync)
- getOoxml() tiene limites de tamano (~10MB)
- Solo funciona con Word abierto

### 3.2 Enfoque B: VSTO/COM Add-in (.NET, solo Windows)

```
Word.exe → VSTO Add-in (.NET/C#) → Named Pipe / HTTP → Hermes Backend
```

**Ventajas:**
- Acceso TOTAL al modelo de objetos de Word
- Soporte para Track Changes, revisiones, macros
- Mejor rendimiento que Office.js
- Integracion nativa con ribbon/cintas

**Desventajas:**
- SOLO Windows (no Mac, no Web, no iPad)
- Requiere Visual Studio + .NET Framework
- Curva de aprendizaje alta
- Distribucion compleja (MSI/ClickOnce)

### 3.3 Enfoque C: MCP Server + Hermes Agent

```
Claude/Cursor/Hermes → MCP Protocol → Word MCP Server → .docx files
```

**Ventajas:**
- No requiere Word abierto (headless)
- Integracion con cualquier MCP client
- Reutiliza servidores existentes (GongRzhe, OfficeMCP, che-word-mcp)

**Desventajas:**
- Sin interfaz de usuario dentro de Word
- Manipula archivos, no el documento activo
- El usuario debe saber usar MCP/CLI

### 3.4 Enfoque D: COM/pywin32 (Python, solo Windows)

```python
import win32com.client
word = win32com.client.Dispatch("Word.Application")
doc = word.ActiveDocument
text = doc.Content.Text
```

**Ventajas:**
- Python nativo, facil de integrar con Hermes
- Acceso completo a la API COM de Word

**Desventajas:**
- SOLO Windows + Word instalado
- Fragil (puede dejar procesos zombie)
- Sin UI integrada (requiere GUI separada)

### 3.5 Matriz Comparativa

| Criterio | Add-in Office.js | VSTO/.NET | MCP Server | COM/pywin32 |
|----------|:---:|:---:|:---:|:---:|
| Multiplataforma | ★★★★★ | ★☆☆☆☆ | ★★★★★ | ★☆☆☆☆ |
| UI en Word | ★★★★★ | ★★★★★ | ☆☆☆☆☆ | ★☆☆☆☆ |
| Facilidad desarrollo | ★★★★☆ | ★★☆☆☆ | ★★★★★ | ★★★☆☆ |
| Acceso API Word | ★★★☆☆ | ★★★★★ | ☆☆☆☆☆ | ★★★★★ |
| Headless (sin Word) | ☆☆☆☆☆ | ☆☆☆☆☆ | ★★★★★ | ☆☆☆☆☆ |
| Streaming (SSE/WS) | ★★★★☆ | ★★★★☆ | ★★★☆☆ | ★★★★☆ |
| Distribucion | ★★★☆☆ | ★★☆☆☆ | ★★★★★ | ★★★☆☆ |

---

## 4. MCP Servers para Word: Analisis Comparativo

### 4.1 GongRzhe/Office-Word-MCP-Server ★1,902

- **Stack:** Python + python-docx + FastMCP
- **Herramientas:** 50+ (crear, leer, formatear, tablas, comentarios, proteccion)
- **Estado:** ARCHIVADO (marzo 2026)
- **Fortalezas:** Mas popular, documentacion extensa, disponible en PyPI y Smithery
- **Debilidades:** Archivado, sin mantenimiento activo, basado en python-docx (no preserva 100% del formato)

```json
// Configuracion tipica para Claude Desktop
{
  "mcpServers": {
    "word-document-server": {
      "command": "uvx",
      "args": ["--from", "office-word-mcp-server", "word_mcp_server"]
    }
  }
}
```

### 4.2 OfficeMCP/OfficeMCP ★78

- **Stack:** Python + COM (Windows) + FastMCP
- **Herramientas:** Full Office suite (Word, Excel, PowerPoint, Access, Outlook, Visio, Project)
- **Estado:** Activo (v1.0.5)
- **Fortalezas:** Suite completa de Office, COM da acceso total, soporta WPS Office
- **Debilidades:** SOLO Windows (usa COM), requiere Office instalado, seguridad (ejecuta codigo Python arbitrario via RunPython)

```bash
# Modo SSE (recomendado)
uvx officemcp sse --port 8888
# URL: http://127.0.0.1:8888/sse
```

### 4.3 PsychQuant/che-word-mcp ★3

- **Stack:** Swift puro + OOXML nativo (sin dependencias)
- **Herramientas:** 233 herramientas MCP
- **Estado:** MUY activo (v3.13.5, abril 2026)
- **Fortalezas:** 
  - Preservacion de bytes (byte-perfect round-trip)
  - Track Changes programaticos (ins/del/move)
  - 100% cobertura Office.js OOXML Roadmap P0
  - Binario unico, sin runtime externo
- **Debilidades:** Solo macOS (universal binary x86_64+arm64), comunidad pequena, documentacion densa

```bash
# Instalacion via MCPB (one-click)
# Descargar .mcpb de Releases y doble-click
```

### 4.4 ForLegalAI/mcp-ms-office-documents ★25

- **Stack:** Python + python-docx + python-pptx
- **Herramientas:** Crear documentos Office (.docx, .pptx, .xlsx, .eml)
- **Estado:** Activo, enfoque legal
- **Fortalezas:** Especializado en documentos legales, multi-formato
- **Debilidades:** Solo creacion, no lectura/edicion del documento activo

### 4.5 vAirpower/macos-office365-mcp-server ★13

- **Stack:** Python + AppleScript + python-docx
- **Herramientas:** PowerPoint, Word, Excel via AppleScript
- **Estado:** Activo, PoC personal
- **Fortalezas:** Controla la app de Office en macOS via AppleScript
- **Debilidades:** Solo macOS, requiere Office instalado, PoC no produccion

---

## 5. Arquitectura Propuesta

### 5.1 Arquitectura Hibrida (Recomendada)

Combinamos los 3 mejores enfoques en capas:

```
┌─────────────────────────────────────────────────────────────────────┐
│                    HERMES WORD INTEGRATION                           │
│                                                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ CAPA 1: IN-APP (Word Add-in)                           ⭐MVP │   │
│  │                                                                │   │
│  │  taskpane.html  ←→  taskpane.js  ←→  backend_server.py       │   │
│  │  (Chat UI)           (Office.js)       (FastAPI :8765)        │   │
│  │                                                                │   │
│  │  Flujos:                                                       │   │
│  │  1. Leer doc → getOoxml() → POST /chat → LLM → respuesta      │   │
│  │  2. Insertar → insertText() en seleccion/final/inicio         │   │
│  │  3. Analizar → extraer metricas, estructura, sugerencias      │   │
│  │  4. Resumir → bullet points, abstract ejecutivo               │   │
│  │  5. Reescribir → mejorar claridad, tono, gramatica            │   │
│  │  6. Traducir → EN↔ES manteniendo formato                      │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                                                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ CAPA 2: HEADLESS (CLI + python-docx)                          │   │
│  │                                                                │   │
│  │  hermes word read contrato.docx                                │   │
│  │  hermes word create --from-md informe.md --output informe.docx │   │
│  │  hermes word analyze --metrics --suggestions propuesta.docx    │   │
│  │  hermes word convert contrato.docx --to markdown               │   │
│  └──────────────────────────────────────────────────────────────┘   │
│                                                                      │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │ CAPA 3: MCP SERVER (protocolo estandar)                       │   │
│  │                                                                │   │
│  │  Integrar che-word-mcp o GongRzhe como MCP server             │   │
│  │  → Hermes Agent lo descubre automaticamente                   │   │
│  │  → 233 herramientas disponibles sin codigo adicional          │   │
│  └──────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
```

### 5.2 Stack Tecnologico

```
Frontend (Word Task Pane):
├── HTML5 + CSS3 + Vanilla JS (o Preact para lightweight)
├── Office.js (API oficial de Microsoft)
├── Markdown rendering (marked.js)
└── Code highlighting (highlight.js)

Backend (Hermes Server):
├── FastAPI + uvicorn (Python 3.8+)
├── WebSocket (streaming de respuestas)
├── SSE (Server-Sent Events, alternativa a WS)
├── DeepSeek v4 / GPT-4o / Claude (LLM)
├── python-docx (lectura/escritura de .docx)
├── mammoth (docx → markdown)
└── markitdown (conversion avanzada Microsoft)

MCP Integration (opcional):
├── che-word-mcp (Swift, macOS) o
├── office-word-mcp-server (Python, multiplataforma) o
└── OfficeMCP (Python, Windows COM)
```

### 5.3 API Endpoints

```yaml
POST /chat:
  description: Enviar mensaje + documento para procesamiento
  request:
    document: string (OOXML o HTML del doc activo)
    message: string (consulta del usuario)
    action: "chat" | "analyze" | "summarize" | "rewrite" | "translate"
    history: array (conversacion previa)
    options:
      model: string (deepseek-v4, gpt-4o, etc.)
      temperature: float
      target_language: string (para traduccion)
  response:
    response: string (respuesta del LLM)
    insert_mode: "cursor" | "end" | "start" | "replace_selection"
    stats:
      words: int
      tokens_used: int
      cost: float

GET /health:
  description: Health check del servidor

WS /ws:
  description: Canal WebSocket para streaming bidireccional
  protocol: JSON messages
  events: "message", "token", "done", "error"
```

### 5.4 Seguridad

```
┌─────────────────────────────────────────────────────────┐
│ Capa de Seguridad                                        │
│                                                          │
│ 1. HTTPS (TLS 1.3) - cert autofirmado local              │
│ 2. CORS restringido a localhost:3000                     │
│ 3. API keys en variables de entorno (nunca en frontend) │
│ 4. Rate limiting (100 req/min)                           │
│ 5. Sanitizacion de input OOXML (defensa XXE)            │
│ 6. Logs de auditoria (sin datos de documentos)          │
│ 7. Timeout de sesion (30 min inactividad)               │
└─────────────────────────────────────────────────────────┘
```

---

## 6. Plan de Implementacion

### Fase 1: MVP (3-5 dias)

```
Dia 1: Backend Hermes
├── FastAPI server con endpoint /chat
├── Integracion con DeepSeek v4
├── Procesamiento basico de OOXML → texto
└── Tests de integracion

Dia 2: Frontend Add-in
├── taskpane.html con UI de chat
├── Office.js: leer documento (getOoxml)
├── Office.js: insertar texto (insertText)
├── Comunicacion fetch() con backend
└── Sideload en Word para pruebas

Dia 3: Funcionalidades Core
├── Boton "Leer documento"
├── Boton "Analizar" (estructura, metricas)
├── Boton "Resumir" (3-5 bullet points)
├── Boton "Reescribir" (mejora de texto)
├── Boton "Insertar en documento"
└── Historial de chat en panel

Dia 4: Pulido y UX
├── Streaming de respuestas (SSE o WS)
├── Markdown rendering en chat
├── Indicador de carga/typing
├── Manejo de errores elegante
└── Estilos visuales profesionales

Dia 5: Testing y Documentacion
├── Pruebas con documentos reales
├── README + instrucciones de instalacion
├── Manifest.xml para sideload
└── Script de setup automatizado
```

### Fase 2: Produccion (2-3 semanas)

```
Semana 1:
├── MCP server integration (che-word-mcp o GongRzhe)
├── Soporte para multiples LLMs (DeepSeek, OpenAI, Claude)
├── Templates de prompts por tipo de documento
├── Configuracion persistente (JSON config file)
└── Gestion de documentos grandes (>10MB, chunking)

Semana 2:
├── WebSocket para sesiones persistentes
├── Historial de conversaciones (SQLite local)
├── Insertar en seleccion actual (no solo al final)
├── Formato rico (negritas, cursivas, tablas)
├── Comandos slash (/analyze, /summarize, /rewrite)
└── Atajos de teclado (Ctrl+Enter para enviar)

Semana 3:
├── Microsoft Graph API (OneDrive/SharePoint docs)
├── Autenticacion (si se requiere multi-usuario)
├── Packaging para distribucion
├── Publicacion en GitHub Releases
├── Tests E2E con Playwright
└── CI/CD pipeline (GitHub Actions)
```

### Fase 3: Premium (1-2 meses)

```
├── Track Changes revisiones (via che-word-mcp)
├── Comparacion de documentos (diff visual)
├── Traduccion multi-idioma preservando formato
├── Voice dictation (Web Speech API)
├── Integracion con bases de conocimiento RAG
├── Modo colaborativo (multiples usuarios)
├── Custom skills/tools por tipo de documento
└── Publicacion en Microsoft AppSource
```

---

## 7. Competitive Landscape

### 7.1 Comparativa con Soluciones Existentes

| Solucion | Chat en Word | Open Source | Custom LLM | Track Changes | Precio |
|----------|:---:|:---:|:---:|:---:|---|
| **Hermes Word Add-in** | ✅ | ✅ | ✅ | 🔜 | Gratis |
| Microsoft Copilot | ✅ | ❌ | ❌ | ✅ | $30/usr/mes |
| Grammarly | ✅ | ❌ | ❌ | ❌ | $12/usr/mes |
| WordTune | ✅ | ❌ | ❌ | ❌ | $10/usr/mes |
| ChatGPT (manual) | ❌ | ❌ | ❌ | ❌ | $20/usr/mes |
| Claude (manual) | ❌ | ❌ | ❌ | ❌ | $20/usr/mes |

### 7.2 Ventajas Competitivas de Hermes

1. **Open source total** - Sin vendor lock-in, auditable, customizable
2. **Multi-LLM** - DeepSeek, OpenAI, Claude, modelos locales (llama.cpp)
3. **Privacidad** - Datos nunca salen de la maquina local
4. **Extensible** - Skills/plugins por tipo de documento
5. **Integracion MCP** - Compatible con el ecosistema de herramientas AI
6. **Costo cero** - Sin suscripciones (solo API keys de LLM)

---

## 8. Riesgos y Desafios

### 8.1 Riesgos Tecnicos

| Riesgo | Impacto | Probabilidad | Mitigacion |
|--------|---------|-------------|------------|
| Office.js limita tamano de OOXML | Medio | Alta | Chunking por secciones, usar getHtml() |
| HTTPS requerido para desarrollo | Bajo | Alta | Cert autofirmado o ngrok |
| Cambios en API de Office.js | Alto | Baja | Tests E2E, monitoreo de changelog |
| Latencia LLM > 5s | Medio | Media | Streaming + indicador de progreso |
| Corrupcion de OOXML en round-trip | Alto | Media | Usar che-word-mcp (byte-perfect) o python-docx con tests |
| WSL2 networking issues | Bajo | Media | Script de diagnostico automatico |

### 8.2 Riesgos de Producto

| Riesgo | Impacto | Probabilidad | Mitigacion |
|--------|---------|-------------|------------|
| Microsoft mejora Copilot y absorbe el mercado | Alto | Alta | Diferenciacion: open source, custom LLMs, privacidad |
| Baja adopcion (instalacion compleja) | Medio | Media | Script de setup automatizado, MSI installer |
| Office Web Add-ins deprecados | Alto | Baja | Microsoft esta invirtiendo en la plataforma |

---

## 9. Conclusiones y Recomendaciones

### 9.1 Recomendacion Principal

**Implementar el Enfoque A (Office Web Add-in) como MVP**, complementado con:

1. **Backend Python (FastAPI + WebSocket)** corriendo localmente
2. **Integracion MCP opcional** via che-word-mcp (macOS) o office-word-mcp-server (cross-platform)
3. **CLI complementaria** para operaciones headless via python-docx

### 9.2 Por que NO VSTO/COM como primera opcion

- Solo funciona en Windows (perdemos Mac, Web, iPad)
- Requiere tooling especifico (Visual Studio, .NET)
- Distribucion mas compleja
- El ecosistema AI se mueve hacia web/JS

### 9.3 Quick Wins (bajo esfuerzo, alto impacto)

1. **CLI `hermes word`** con python-docx (1 dia) → utilidad inmediata
2. **Integracion MCP server** (1 dia) → 50-233 herramientas gratis
3. **Add-in basico con chat** (3 dias) → producto diferenciador
4. **Skill en ~/.hermes/skills/** para reutilizar el conocimiento

### 9.4 Estado Actual del Proyecto

El proyecto ya cuenta con:
- ✅ Skill `word-hermes-addin` con arquitectura definida
- ✅ Skill `word-office-integration` con enfoques A y B
- ✅ Backend server funcional (backend_server.py)
- ✅ Frontend HTML/JS basico (frontend.html)
- ✅ Manifest XML para sideload
- ✅ Modo simulacion sin Word (pruebas en navegador)

**Proximo paso inmediato:** Completar el frontend con streaming SSE/WebSocket y pulir la UX.

---

## 10. Referencias

### Proyectos GitHub

| Proyecto | URL |
|----------|-----|
| GongRzhe/Office-Word-MCP-Server | https://github.com/GongRzhe/Office-Word-MCP-Server |
| OfficeMCP/OfficeMCP | https://github.com/OfficeMCP/OfficeMCP |
| PsychQuant/che-word-mcp | https://github.com/PsychQuant/che-word-mcp |
| ForLegalAI/mcp-ms-office-documents | https://github.com/ForLegalAI/mcp-ms-office-documents |
| vAirpower/macos-office365-mcp-server | https://github.com/vAirpower/macos-office365-mcp-server |
| ecator/cs-office-mcp-server | https://github.com/ecator/cs-office-mcp-server |

### Documentacion Oficial

- Office Add-ins docs: https://learn.microsoft.com/en-us/office/dev/add-ins/
- Word JavaScript API: https://learn.microsoft.com/en-us/javascript/api/word
- Model Context Protocol: https://modelcontextprotocol.io/
- python-docx: https://python-docx.readthedocs.io/
- FastMCP: https://github.com/modelcontextprotocol/python-sdk

### Skills Internas de Hermes

- `word-hermes-addin` (~/.hermes/skills/productivity/word-hermes-addin/SKILL.md)
- `word-office-integration` (~/.hermes/skills/word-office-integration/SKILL.md)

---

**Fin del informe.**
