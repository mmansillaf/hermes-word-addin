#!/usr/bin/env python3
"""
Hermes Word Backend Server — Multi-Proveedor
Recibe texto/documentos desde el add-in de Word y responde usando LLM.

Uso: python backend_server.py [--port 8765]

Proveedores soportados via variable LLM_PROVIDER:
  deepseek          → DeepSeek API (default)
  openai            → OpenAI API
  anthropic         → Anthropic Claude API
  openai-compatible → Cualquier endpoint compatible (Groq, Together, llama.cpp, vLLM, etc.)

Configuracion:
  LLM_PROVIDER   = deepseek | openai | anthropic | openai-compatible
  LLM_MODEL      = modelo a usar (opcional, tiene default por provider)
  LLM_BASE_URL   = URL base para openai-compatible (obligatorio para ese provider)

API Keys (segun provider):
  deepseek           → DEEPSEEK_API_KEY
  openai             → OPENAI_API_KEY
  anthropic          → ANTHROPIC_API_KEY
  openai-compatible  → LLM_API_KEY
"""

import http.server
import json
import os
import sys
import urllib.parse
import urllib.request
from pathlib import Path

PORT = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[1] == '--port' else 8765

SYSTEM_PROMPT = """Eres Hermes, un asistente AI integrado en Microsoft Word.
Analizas documentos y respondes consultas del usuario.
Se conciso y practico. Si el usuario pide modificar el documento,
indica claramente QUE cambiar y propone el texto exacto."""

# ---------------------------------------------------------------------------
# Configuracion de proveedor
# ---------------------------------------------------------------------------

PROVIDER = os.environ.get('LLM_PROVIDER', 'deepseek')

PROVIDER_CONFIG = {
    'deepseek': {
        'url': 'https://api.deepseek.com/v1/chat/completions',
        'model': os.environ.get('LLM_MODEL', 'deepseek-chat'),
        'api_key': os.environ.get('DEEPSEEK_API_KEY', ''),
        'auth_header': 'Bearer {key}',
        'body': lambda model, messages: {
            'model': model,
            'messages': messages,
            'max_tokens': 2000,
            'temperature': 0.7,
        },
        'parse_response': lambda data: data['choices'][0]['message']['content'],
    },
    'openai': {
        'url': 'https://api.openai.com/v1/chat/completions',
        'model': os.environ.get('LLM_MODEL', 'gpt-4o-mini'),
        'api_key': os.environ.get('OPENAI_API_KEY', ''),
        'auth_header': 'Bearer {key}',
        'body': lambda model, messages: {
            'model': model,
            'messages': messages,
            'max_tokens': 2000,
        },
        'parse_response': lambda data: data['choices'][0]['message']['content'],
    },
    'anthropic': {
        'url': 'https://api.anthropic.com/v1/messages',
        'model': os.environ.get('LLM_MODEL', 'claude-sonnet-4-20250514'),
        'api_key': os.environ.get('ANTHROPIC_API_KEY', ''),
        'auth_header': '{key}',  # Anthropic usa x-api-key
        'extra_headers': lambda: {
            'x-api-key': os.environ.get('ANTHROPIC_API_KEY', ''),
            'anthropic-version': '2023-06-01',
        },
        'body': lambda model, messages: {
            'model': model,
            'system': messages[0]['content'] if messages[0]['role'] == 'system' else SYSTEM_PROMPT,
            'messages': [m for m in messages if m['role'] != 'system'],
            'max_tokens': 2000,
        },
        'parse_response': lambda data: data['content'][0]['text'],
    },
    'openai-compatible': {
        'url': os.environ.get('LLM_BASE_URL', 'http://localhost:8080/v1/chat/completions'),
        'model': os.environ.get('LLM_MODEL', 'llama-3.1-8b-instruct'),
        'api_key': os.environ.get('LLM_API_KEY', ''),
        'auth_header': 'Bearer {key}',
        'body': lambda model, messages: {
            'model': model,
            'messages': messages,
            'max_tokens': 2000,
            'temperature': 0.7,
        },
        'parse_response': lambda data: data['choices'][0]['message']['content'],
    },
}

# ---------------------------------------------------------------------------
# LLM Backend unificado
# ---------------------------------------------------------------------------

def call_llm(prompt, document_text=''):
    """Llama al LLM segun LLM_PROVIDER. Sin API key → analisis local."""
    cfg = PROVIDER_CONFIG.get(PROVIDER)

    if not cfg or not cfg['api_key']:
        return local_analysis(prompt, document_text)

    messages = [
        {'role': 'system', 'content': SYSTEM_PROMPT},
        {'role': 'user', 'content': f'{prompt}\n\n--- DOCUMENTO ACTUAL ---\n{document_text[:8000]}'},
    ]

    body = cfg['body'](cfg['model'], messages)
    url = cfg['url']
    headers = {'Content-Type': 'application/json'}

    if PROVIDER == 'anthropic':
        headers.update(cfg['extra_headers']())
    else:
        headers['Authorization'] = cfg['auth_header'].format(key=cfg['api_key'])

    try:
        req = urllib.request.Request(url, data=json.dumps(body).encode(), headers=headers)
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())
            return cfg['parse_response'](result)
    except Exception as e:
        return f'[Error {PROVIDER}: {e}]\n\n{local_analysis(prompt, document_text)}'


def local_analysis(prompt, document_text):
    """Analisis local del documento (sin LLM). Util para testing."""
    words = len(document_text.split()) if document_text else 0
    lines = document_text.count('\n') + 1 if document_text else 0
    chars = len(document_text)

    first_lines = '\n'.join(document_text.split('\n')[:3]) if document_text else '(vacio)'

    return f"""[MODO LOCAL - Sin API key configurada para {PROVIDER}]

📊 ANALISIS DEL DOCUMENTO:
- Palabras: {words}
- Lineas: {lines}
- Caracteres: {chars}

📝 PRIMERAS LINEAS:
{first_lines[:300]}

💬 Tu consulta: "{prompt}"

⚙️  Configuracion actual:
    LLM_PROVIDER = {PROVIDER}
    Para usar IA real, configura la API key correspondiente.
    Ver README o docs/guia_instalacion_windows.md
"""

# ---------------------------------------------------------------------------
# HTTP Server
# ---------------------------------------------------------------------------

class WordHermesHandler(http.server.BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self.send_response(200)
        self._cors_headers()
        self.end_headers()

    def do_GET(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/health':
            self.send_response(200)
            self._cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'status': 'ok',
                'service': 'Hermes Word Backend',
                'version': '0.2.0',
                'provider': PROVIDER,
                'model': PROVIDER_CONFIG.get(PROVIDER, {}).get('model', 'none'),
                'has_api_key': bool(PROVIDER_CONFIG.get(PROVIDER, {}).get('api_key')),
            }).encode())
            return

        if parsed.path == '/':
            frontend_path = Path(__file__).parent / 'frontend.html'
            if frontend_path.exists():
                self.send_response(200)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.end_headers()
                self.wfile.write(frontend_path.read_bytes())
                return

        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        parsed = urllib.parse.urlparse(self.path)

        if parsed.path == '/chat':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)

            try:
                data = json.loads(body)
                prompt = data.get('prompt', '')
                document_text = data.get('document', '')
                action = data.get('action', 'chat')

                print(f'\n[Word→Hermes] Provider: {PROVIDER} | Action: {action}')
                print(f'[Document] {len(document_text)} chars, {len(document_text.split())} words')

                full_prompt = self._build_prompt(prompt, document_text, action)
                response = call_llm(full_prompt, document_text)

                print(f'[Hermes→Word] Response: {len(response)} chars')

                self.send_response(200)
                self._cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': True,
                    'response': response,
                    'action': action,
                    'provider': PROVIDER,
                    'document_stats': {
                        'words': len(document_text.split()),
                        'chars': len(document_text),
                    },
                }).encode())

            except json.JSONDecodeError:
                self.send_response(400)
                self.end_headers()
                self.wfile.write(b'{"error": "Invalid JSON"}')
            except Exception as e:
                self.send_response(500)
                self._cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': False,
                    'error': str(e),
                }).encode())
        else:
            self.send_response(404)
            self.end_headers()

    def _build_prompt(self, user_prompt, document_text, action):
        prompts = {
            'chat': f'Usuario pregunta: {user_prompt}',
            'analyze': f'Analiza este documento y responde: {user_prompt}\n\n---\n{document_text[:6000]}',
            'rewrite': f'Reescribe el siguiente texto segun esta instruccion: {user_prompt}\n\n---\n{document_text[:6000]}',
            'summarize': f'Resume este documento. Instruccion adicional: {user_prompt}\n\n---\n{document_text[:6000]}',
        }
        return prompts.get(action, user_prompt)

    def _cors_headers(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')

    def log_message(self, format, *args):
        print(f'[{self.log_date_time_string()}] {args[0]}')


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == '__main__':
    cfg = PROVIDER_CONFIG.get(PROVIDER, {})
    has_key = bool(cfg.get('api_key'))
    model = cfg.get('model', 'none')

    mode_str = f'LLM: {PROVIDER} ({model})' if has_key else f'LOCAL (sin key para {PROVIDER})'

    print(f"""
╔══════════════════════════════════════════════╗
║   Hermes Word Backend Server v0.2.0         ║
╠══════════════════════════════════════════════╣
║  Puerto:  {PORT}                              ║
║  Provider: {PROVIDER:<31} ║
║  Model:   {model[:31]:<31} ║
║  Modo:    {mode_str[:31]} ║
╠══════════════════════════════════════════════╣
║  Health:  GET  http://localhost:{PORT}/health ║
║  Chat:    POST http://localhost:{PORT}/chat   ║
║  Frontend:     http://localhost:{PORT}        ║
╚══════════════════════════════════════════════╝
""")

    server = http.server.HTTPServer(('0.0.0.0', PORT), WordHermesHandler)
    print(f'Servidor corriendo en http://localhost:{PORT}')
    print('Presiona Ctrl+C para detener\n')

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServidor detenido.')
        server.shutdown()
