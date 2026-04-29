# Hermes Word Add-in

Panel de chat AI dentro de Microsoft Word. Conecta el documento activo con Hermes Agent (DeepSeek, OpenAI, Claude) para leer, analizar, resumir, reescribir y modificar documentos.

## Arquitectura

```
Word (Win/Mac)                    Backend Hermes (local)
┌──────────────────────┐         ┌──────────────────────┐
│  Task Pane Add-in    │  WSS    │  FastAPI + WebSocket │
│  (HTML/JS/CSS)       │◄───────►│  localhost:8765      │
│                      │  JSON   │                      │
│  Office.js API:      │         │  DeepSeek v4 / GPT   │
│  - getOoxml() leer   │         │  (o LLM local)       │
│  - insertText()      │         │                      │
└──────────────────────┘         └──────────────────────┘
```

## Estructura

```
hermes-word-addin/
├── src/
│   ├── backend_server.py      # Servidor Python (FastAPI wrapper)
│   ├── frontend.html          # UI del add-in (HTML/CSS/JS + Office.js)
│   ├── manifest.xml           # Manifest para sideload en Word
│   └── convert_to_formats.py  # Conversión docx ↔ md
├── research/                  # Investigación y análisis
│   ├── informe_integracion.md/.html  # Informe completo (Abr 2026)
│   ├── feasibility_report.md         # Viabilidad técnica
│   ├── protocol_spec.md              # Especificación WS JSON
│   ├── architecture_diagrams.md      # Diagramas de arquitectura
│   └── word_integration_research.md  # Investigación inicial
├── scripts/                   # Scripts de utilidad
├── docs/                      # Documentación adicional
└── README.md
```

## Instalación

### Requisitos

- **Word 2016+** (Windows o Mac)
- **Python 3.8+**
- API key de LLM (DeepSeek, OpenAI, o Claude)

### Setup Rápido

```bash
# 1. Clonar
git clone <repo-url>
cd hermes-word-addin

# 2. Dependencias Python
pip install fastapi uvicorn websockets python-docx mammoth

# 3. Configurar API key
export DEEPSEEK_API_KEY="sk-..."

# 4. Lanzar backend
cd src
python backend_server.py --port 8765

# 5. En Word: Archivo > Opciones > Programador > Complementos > Cargar manifest.xml
```

### Sin Word (pruebas)

```bash
python backend_server.py --port 8765
# Abrir http://localhost:8765 en navegador
```

## Uso

1. Abrir un documento en Word
2. Click en "Hermes AI" en la pestaña Inicio
3. Usar botones: **Leer doc**, **Analizar**, **Resumir**, **Reescribir**, **Insertar en doc**
4. O escribir consultas libres en el chat

## Referencias

- [Office Add-ins docs](https://learn.microsoft.com/en-us/office/dev/add-ins/)
- [Word JavaScript API](https://learn.microsoft.com/en-us/javascript/api/word)
- [Model Context Protocol](https://modelcontextprotocol.io/)
- [python-docx](https://python-docx.readthedocs.io/)
