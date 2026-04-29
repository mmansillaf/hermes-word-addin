# Guia de Instalacion: Hermes Word Add-in en Windows

**Paso a paso para tener el panel de chat AI dentro de Microsoft Word.**

---

## ¿Necesito instalar Hermes Agent?

**NO.** El backend (`backend_server.py`) es un servidor Python autónomo que llama directamente a la API de DeepSeek (o OpenAI). No requiere Hermes Agent instalado.

Si en el futuro queres usar features avanzadas de Hermes (skills, memoria persistente, RAG), podes instalar Hermes Agent en la misma maquina y hacer que el backend le delegue tareas. Pero para el MVP **no es necesario**.

---

## Requisitos Previos

| Requisito | Detalle |
|-----------|---------|
| Windows 10/11 | 64-bit |
| Microsoft Word | 2016 o superior (Desktop, no Online) |
| Python | 3.8 o superior |
| API Key | DeepSeek, OpenAI, Anthropic, o endpoint compatible |
| Git | Para clonar el repo |

---

## Paso 1: Instalar Python en Windows

1. Descargar Python desde https://www.python.org/downloads/
2. Ejecutar el instalador
3. **IMPORTANTE:** Marcar la casilla "Add Python to PATH"
4. Click en "Install Now"
5. Verificar abriendo PowerShell:
   ```
   python --version
   ```
   Debe mostrar `Python 3.x.x`

---

## Paso 2: Clonar el repositorio

Abrir PowerShell y ejecutar:

```powershell
cd C:\Users\TU_USUARIO\
git clone https://github.com/mmansillaf/hermes-word-addin.git
cd hermes-word-addin
```

Si no tenes Git, descargalo de https://git-scm.com/download/win

---

## Paso 3: Instalar dependencias Python

En PowerShell, dentro de la carpeta `hermes-word-addin`:

```powershell
pip install fastapi uvicorn websockets
```

Si pip da error, probar:
```powershell
python -m pip install fastapi uvicorn websockets
```

---

## Paso 4: Configurar proveedor LLM y API Key

Elegi tu proveedor configurando estas variables de entorno:

**DeepSeek (recomendado, mejor relacion costo/calidad):**
```powershell
$env:LLM_PROVIDER="deepseek"
$env:DEEPSEEK_API_KEY="sk-tu-key"
```

**OpenAI:**
```powershell
$env:LLM_PROVIDER="openai"
$env:OPENAI_API_KEY="sk-tu-key"
$env:LLM_MODEL="gpt-4o"    # opcional, default: gpt-4o-mini
```

**Anthropic Claude:**
```powershell
$env:LLM_PROVIDER="anthropic"
$env:ANTHROPIC_API_KEY="sk-ant-tu-key"
$env:LLM_MODEL="claude-sonnet-4-20250514"  # opcional
```

**Endpoint OpenAI-compatible (Groq, Together, llama.cpp, vLLM, Ollama, etc.):**
```powershell
$env:LLM_PROVIDER="openai-compatible"
$env:LLM_BASE_URL="https://api.groq.com/openai/v1/chat/completions"
$env:LLM_API_KEY="gsk_tu-key"
$env:LLM_MODEL="llama-3.1-8b-instant"
```

Para que persista entre reinicios, crea variables de entorno del sistema:
1. Win + R, escribir `sysdm.cpl`
2. Pestana "Opciones avanzadas" > "Variables de entorno"
3. Nueva variable de sistema para cada una de las que necesites

---

## Paso 5: Lanzar el backend

Desde la carpeta `hermes-word-addin`:

```powershell
cd src
python backend_server.py --port 8765
```

Debes ver algo como:
```
Hermes Word Backend Server
Listening on http://localhost:8765
Open http://localhost:8765 in browser for standalone mode
```

**NO cierres esta ventana.** El servidor debe estar corriendo mientras usas Word.

### Probar sin Word (opcional)

Abri tu navegador en http://localhost:8765
Vas a ver el panel de chat. Escribi un mensaje y proba que responde.

---

## Paso 6: Instalar el Add-in en Word

### 6.1 Habilitar la pestana Programador

1. Abri Microsoft Word
2. Click en **Archivo** > **Opciones**
3. Selecciona **Personalizar cinta de opciones**
4. En la columna derecha, marca la casilla **Programador**
5. Click **Aceptar**

### 6.2 Cargar el manifest

1. En Word, ve a la pestana **Programador**
2. Click en **Complementos** (icono de pieza de puzzle)
3. En la ventana que se abre, selecciona **MIS COMPLEMENTOS**
4. Click en **Cargar mi complemento** (al final de la lista desplegable)
5. Navega hasta la carpeta del proyecto: `C:\Users\TU_USUARIO\hermes-word-addin\src\`
6. Selecciona el archivo `manifest.xml`
7. Click **Abrir**

### 6.3 Abrir el panel de chat

El add-in aparece como un boton "Hermes AI" en la pestana **Inicio** de Word. Click para abrir el panel lateral.

---

## Paso 7: Usar el Add-in

El panel tiene estas secciones:

| Boton | Que hace |
|-------|----------|
| **Leer doc** | Carga el texto del documento activo en el panel |
| **Analizar** | Hermes analiza estructura, metricas, sugiere mejoras |
| **Resumir** | Genera 3-5 bullet points del documento |
| **Reescribir** | Mejora claridad y profesionalismo del texto |
| **Insertar en doc** | Inserta la respuesta de Hermes al final del documento |

Tambien podes escribir consultas libres en el campo de texto y presionar Enter.

---

## Solucion de Problemas

### "No se pudo cargar el complemento"

- Verifica que el backend esta corriendo (Paso 5)
- Abri http://localhost:8765 en el navegador para confirmar
- Cierra y reabre Word

### "Failed to fetch" en el panel

El backend no esta accesible. Verifica:
1. La ventana de PowerShell sigue abierta con el servidor corriendo
2. No hay firewall bloqueando el puerto 8765
3. Ejecuta `netstat -an | findstr 8765` para ver si el puerto esta en uso

### "El texto no se inserta en el documento"

Office.js requiere que el documento este en modo edicion (no solo lectura). Asegurate de que el documento no esta protegido.

### Error de API key

Si ves "No API key configured", la variable de entorno no esta configurada. Revisa el Paso 4.

---

## Arquitectura

```
Windows 10/11
┌────────────────────────────────────────────┐
│  Microsoft Word                             │
│  ┌──────────────────────────────────────┐  │
│  │  Task Pane (panel lateral)           │  │
│  │  frontend.html + Office.js           │  │
│  │  Leer doc: getOoxml()                │  │
│  │  Escribir: insertText()              │  │
│  └──────────────┬───────────────────────┘  │
│                 │ HTTP (localhost:8765)     │
│  ┌──────────────▼───────────────────────┐  │
│  │  backend_server.py                   │  │
│  │  - Sirve frontend.html               │  │
│  │  - Endpoint /chat                    │  │
│  │  - Llama a DeepSeek API              │  │
│  └──────────────┬───────────────────────┘  │
│                 │ HTTPS (api.deepseek.com)  │
└─────────────────┼───────────────────────────┘
                  ▼
        DeepSeek API (nube)
```

---

## Proximos Pasos (Mejoras)

1. **Streaming:** Respuestas en tiempo real (SSE/WebSocket)
2. **HTTPS local:** Certificado autofirmado para desarrollo
3. **Historial persistente:** Guardar conversaciones en SQLite
4. **Insertar en seleccion:** Reemplazar el texto seleccionado en vez de insertar al final
5. **Formato rico:** Negritas, cursivas, tablas en las respuestas
