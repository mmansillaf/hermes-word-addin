<p align="center">
  <img src="https://img.shields.io/badge/python-3.8+-blue?logo=python&logoColor=white" alt="Python">
  <img src="https://img.shields.io/badge/word-2016+-2B579A?logo=microsoft-word&logoColor=white" alt="Word">
  <img src="https://img.shields.io/badge/license-MIT-green" alt="License">
  <img src="https://img.shields.io/badge/status-prototype--funcional-yellow" alt="Status">
  <img src="https://img.shields.io/badge/plataforma-Windows%20%7C%20macOS%20%7C%20Web-lightgrey" alt="Platform">
</p>

<h1 align="center">Hermes Word Add-in</h1>

<p align="center"><strong>Panel de chat AI dentro de Microsoft Word.</strong><br>
Leé, analizá, resumí, reescribí y editá documentos<br>usando DeepSeek, OpenAI o Claude sin salir de Word.</p>

<p align="center">
  🌐 <a href="https://mmansillaf.github.io/hermes-word-addin/"><strong>Documentación online</strong></a>
</p>

<p align="center">
  <a href="#-quick-start">Quick Start</a> ·
  <a href="#-instalacion-completa">Instalación</a> ·
  <a href="#-arquitectura">Arquitectura</a> ·
  <a href="#-features">Features</a> ·
  <a href="#-documentacion">Docs</a>
</p>

---

## ¿Qué es?

Un **Office Web Add-in** que agrega un panel lateral de chat AI a Microsoft Word. El add-in:

- **Lee** el documento activo via Office.js
- **Procesa** el texto con LLMs (DeepSeek, OpenAI, Claude)
- **Escribe** las respuestas directamente en el documento

Todo corre **localmente** en tu máquina. Tus documentos nunca salen de tu PC — solo las consultas viajan a la API del LLM que elijas.

```
┌──────────────────────────────────────────────────┐
│  Microsoft Word                                   │
│  ┌──────────────────────┐    ┌─────────────────┐ │
│  │  📄 Documento activo │    │  💬 Chat Panel  │ │
│  │                      │◄──►│                 │ │
│  │  "Informe Q4..."     │    │  > Resumir doc  │ │
│  │                      │    │  > Reescribir   │ │
│  └──────────┬───────────┘    │  > Analizar     │ │
│             │ Office.js      └────────┬────────┘ │
│             └──────────► HTTP :8765 ◄─┘          │
└──────────────────────────────────────────────────┘
                           │
              ┌────────────▼────────────┐
              │   Backend Python        │
              │   FastAPI + DeepSeek    │
              └─────────────────────────┘
```

---

## Quick Start

```bash
# 1. Clonar
git clone https://github.com/mmansillaf/hermes-word-addin.git
cd hermes-word-addin

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Configurar API key
# Windows PowerShell:
$env:DEEPSEEK_API_KEY="sk-tu-key"

# 4. Lanzar backend
cd src
python backend_server.py --port 8765

# 5. En Word: Programador > Complementos > Cargar manifest.xml
```

Abrí `http://localhost:8765` en tu navegador para probar sin Word.

---

## Instalación Completa

Guía detallada paso a paso para Windows:

- 📘 [docs/guia_instalacion_windows.md](docs/guia_instalacion_windows.md)
- 🌐 [docs/guia_instalacion_windows.html](docs/guia_instalacion_windows.html) (diseño profesional)

---

## Funcionalidades

| Botón | Qué hace |
|-------|----------|
| 🔍 **Leer doc** | Carga el texto del documento activo en el panel |
| 📊 **Analizar** | Hermes analiza estructura, métricas, sugiere mejoras |
| 📝 **Resumir** | Genera resumen en bullet points del documento |
| ✨ **Reescribir** | Mejora claridad, tono y profesionalismo |
| 📥 **Insertar** | Inserta la respuesta al final del documento |
| 💬 **Chat libre** | Consultas abiertas sobre el documento |

### Próximamente

- [ ] Streaming de respuestas (SSE/WebSocket)
- [ ] Insertar en selección actual
- [ ] Historial de conversaciones (SQLite)
- [ ] Formato rico (negritas, tablas, listas)
- [ ] Integración con Microsoft Graph (OneDrive)
- [ ] Modo oscuro (sigue el tema de Word)
- [ ] Track Changes y comentarios

---

## Arquitectura

```
┌─────────────────────────────────────────────────────────────┐
│                    HERMES WORD ADD-IN                        │
│                                                              │
│  CAPA 1 — IN-APP (Word)                                      │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  taskpane.html  ←→  taskpane.js  ←→  Office.js API  │   │
│  │  (Chat UI)           (Lógica)          (Word doc)     │   │
│  │                                                       │   │
│  │  Office.js APIs usadas:                               │   │
│  │  · body.getOoxml()     → leer documento               │   │
│  │  · body.getHtml()      → leer como HTML               │   │
│  │  · body.insertText()   → insertar al final            │   │
│  │  · getSelection()      → leer selección               │   │
│  └────────────────────┬─────────────────────────────────┘   │
│                       │ HTTP (localhost:8765)                │
│  CAPA 2 — BACKEND                                           │
│  ┌────────────────────┴─────────────────────────────────┐   │
│  │  backend_server.py (Python)                          │   │
│  │  · Sirve frontend.html                               │   │
│  │  · POST /chat  →  procesa consultas + documento      │   │
│  │  · Conecta con DeepSeek/OpenAI API                   │   │
│  └────────────────────┬─────────────────────────────────┘   │
│                       │ HTTPS                                │
│  CAPA 3 — LLM (NUBE)                                        │
│  ┌────────────────────┴─────────────────────────────────┐   │
│  │  DeepSeek v4 / GPT-4o / Claude                       │   │
│  └──────────────────────────────────────────────────────┘   │
│                                                              │
│  CAPA HEADLESS (CLI, sin Word) — opcional                    │
│  ┌──────────────────────────────────────────────────────┐   │
│  │  python-docx + mammoth + convert_to_formats.py       │   │
│  │  · Leer/crear/modificar .docx desde terminal         │   │
│  │  · Convertir docx ↔ markdown                         │   │
│  └──────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
```

---

## ¿Por qué Hermes Word Add-in?

| | Hermes Word | Copilot | Grammarly | ChatGPT |
|---|:---:|:---:|:---:|:---:|
| **Open source** | ✅ | ❌ | ❌ | ❌ |
| **Custom LLM** | ✅ | ❌ | ❌ | ❌ |
| **Local (privacidad)** | ✅ | ❌ | ❌ | ❌ |
| **Chat en Word** | ✅ | ✅ | ✅ | ❌ |
| **Lee documento** | ✅ | ✅ | ✅ | ❌ |
| **Escribe en doc** | ✅ | ✅ | ✅ | ❌ |
| **Sin suscripción** | ✅ | ❌ | ❌ | ❌ |
| **Extensible** | ✅ | ❌ | ❌ | ❌ |

---

## Documentación

| Documento | Descripción |
|-----------|-------------|
| [Informe de integración](research/informe_integracion.md) | Investigación completa de opciones y estrategia |
| [Informe de viabilidad](research/feasibility_report.md) | Análisis técnico de factibilidad |
| [Especificación de protocolo](research/protocol_spec.md) | Protocolo WebSocket JSON v1.0 |
| [Diagramas de arquitectura](research/architecture_diagrams.md) | Diagramas ASCII + matriz de plataformas |
| [Guía de instalación Windows](docs/guia_instalacion_windows.md) | Paso a paso detallado |
| [Informe técnico](docs/INFORME_TECNICO_HERMES_WORD.md) | Documento para equipo de desarrollo |

---

## Stack Tecnológico

| Capa | Tecnología |
|------|-----------|
| **Frontend** | HTML5, CSS3, Vanilla JS, Office.js |
| **Backend** | Python 3.8+, FastAPI, uvicorn |
| **LLM** | DeepSeek v4 (default), OpenAI, Claude |
| **Word API** | Office JavaScript API (WordApi 1.1+) |
| **Headless** | python-docx, mammoth, markitdown |
| **Protocolo** | HTTP REST + WebSocket (planificado) |

---

## Contribuir

El proyecto está en fase de prototipo funcional. Issues y PRs son bienvenidos.

Áreas donde se necesita ayuda:
- Testing en Mac y Word Online
- Implementación de streaming SSE/WebSocket
- Mejoras de UX/UI en el panel de chat
- Empaquetado para distribución (MSI, AppSource)

---

## Licencia

MIT — [LICENSE](LICENSE)

---

<p align="center">
  <sub>Hecho con ❤️ por <a href="https://github.com/mmansillaf">mmansillaf</a> y <a href="https://hermes-agent.nousresearch.com">Hermes Agent</a></sub>
</p>
