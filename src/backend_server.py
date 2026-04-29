#!/usr/bin/env python3
"""
Hermes Word Backend Server
Recibe texto/documentos desde el add-in de Word y responde usando LLM.

Uso: python backend_server.py [--port 8765]
Para produccion, conecta a DeepSeek/OpenAI via API key en env vars.
"""

import http.server
import json
import os
import sys
import urllib.parse
from pathlib import Path

PORT = int(sys.argv[2]) if len(sys.argv) > 2 and sys.argv[1] == '--port' else 8765

# --- LLM Backend (usa DeepSeek si hay API key, sino responde con analisis local) ---
def call_llm(prompt, document_text=""):
    """Llama al LLM. Si no hay API key, hace analisis local del texto."""
    
    api_key = os.environ.get('DEEPSEEK_API_KEY') or os.environ.get('OPENAI_API_KEY')
    
    if api_key and os.environ.get('DEEPSEEK_API_KEY'):
        return call_deepseek(api_key, prompt, document_text)
    elif api_key:
        return call_openai(api_key, prompt, document_text)
    else:
        return local_analysis(prompt, document_text)


def call_deepseek(api_key, prompt, document_text):
    """Llama a DeepSeek API."""
    import urllib.request
    
    system_msg = """Eres Hermes, un asistente AI integrado en Microsoft Word. 
Analizas documentos y respondes consultas del usuario. 
Se conciso y practico. Si el usuario pide modificar el documento, 
indica claramente QUE cambiar y propone el texto exacto."""
    
    full_prompt = f"{prompt}\n\n--- DOCUMENTO ACTUAL ---\n{document_text[:8000]}"
    
    data = json.dumps({
        "model": "deepseek-chat",
        "messages": [
            {"role": "system", "content": system_msg},
            {"role": "user", "content": full_prompt}
        ],
        "max_tokens": 2000,
        "temperature": 0.7
    }).encode()
    
    req = urllib.request.Request(
        "https://api.deepseek.com/v1/chat/completions",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
    )
    
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())
            return result['choices'][0]['message']['content']
    except Exception as e:
        return f"[Error DeepSeek: {e}]\n\n{local_analysis(prompt, document_text)}"


def call_openai(api_key, prompt, document_text):
    """Llama a OpenAI API."""
    import urllib.request
    
    data = json.dumps({
        "model": "gpt-4o-mini",
        "messages": [
            {"role": "system", "content": "Eres Hermes, asistente AI en Word. Responde conciso."},
            {"role": "user", "content": f"{prompt}\n\nDocumento:\n{document_text[:8000]}"}
        ],
        "max_tokens": 2000
    }).encode()
    
    req = urllib.request.Request(
        "https://api.openai.com/v1/chat/completions",
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
    )
    
    try:
        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read())
            return result['choices'][0]['message']['content']
    except Exception as e:
        return f"[Error OpenAI: {e}]\n\n{local_analysis(prompt, document_text)}"


def local_analysis(prompt, document_text):
    """Analisis local del documento (sin LLM). Util para testing."""
    words = len(document_text.split())
    lines = document_text.count('\n') + 1
    chars = len(document_text)
    
    # Extraer primeras lineas (posible titulo)
    first_lines = '\n'.join(document_text.split('\n')[:3])
    
    return f"""[MODO LOCAL - Sin API key configurada]

📊 ANALISIS DEL DOCUMENTO:
- Palabras: {words}
- Lineas: {lines}
- Caracteres: {chars}

📝 PRIMERAS LINEAS:
{first_lines[:300]}

💬 Tu consulta: "{prompt}"

⚠️  Para respuestas con IA real, configura DEEPSEEK_API_KEY o OPENAI_API_KEY.
    El servidor detectara automaticamente la API disponible.
    Ejemplo: export DEEPSEEK_API_KEY="sk-..."
"""


# --- HTTP Server ---
class WordHermesHandler(http.server.BaseHTTPRequestHandler):
    
    def do_OPTIONS(self):
        """CORS preflight"""
        self.send_response(200)
        self._cors_headers()
        self.end_headers()
    
    def do_GET(self):
        """Sirve archivos estaticos y health check"""
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path == '/health':
            self.send_response(200)
            self._cors_headers()
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({
                'status': 'ok',
                'service': 'Hermes Word Backend',
                'version': '0.1.0'
            }).encode())
            return
        
        if parsed.path == '/':
            # Servir el frontend
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
        """Endpoint principal: recibe documento + prompt, devuelve respuesta"""
        parsed = urllib.parse.urlparse(self.path)
        
        if parsed.path == '/chat':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            
            try:
                data = json.loads(body)
                prompt = data.get('prompt', '')
                document_text = data.get('document', '')
                action = data.get('action', 'chat')  # chat | analyze | rewrite | summarize
                
                print(f"\n[Word→Hermes] Action: {action} | Prompt: {prompt[:80]}...")
                print(f"[Document] {len(document_text)} chars, {len(document_text.split())} words")
                
                # Construir prompt completo segun accion
                full_prompt = self._build_prompt(prompt, document_text, action)
                
                # Llamar al LLM
                response = call_llm(full_prompt, document_text)
                
                print(f"[Hermes→Word] Response: {len(response)} chars")
                
                self.send_response(200)
                self._cors_headers()
                self.send_header('Content-Type', 'application/json')
                self.end_headers()
                self.wfile.write(json.dumps({
                    'success': True,
                    'response': response,
                    'action': action,
                    'document_stats': {
                        'words': len(document_text.split()),
                        'chars': len(document_text)
                    }
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
                    'error': str(e)
                }).encode())
        else:
            self.send_response(404)
            self.end_headers()
    
    def _build_prompt(self, user_prompt, document_text, action):
        """Construye el prompt para el LLM segun la accion."""
        prompts = {
            'chat': f"Usuario pregunta: {user_prompt}",
            'analyze': f"Analiza este documento y responde: {user_prompt}\n\n---\n{document_text[:6000]}",
            'rewrite': f"Reescribe el siguiente texto segun esta instruccion: {user_prompt}\n\n---\n{document_text[:6000]}",
            'summarize': f"Resume este documento. Instruccion adicional: {user_prompt}\n\n---\n{document_text[:6000]}"
        }
        return prompts.get(action, user_prompt)
    
    def _cors_headers(self):
        """Headers CORS para permitir conexion desde el add-in (origen file:// o https)"""
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type, Authorization')
    
    def log_message(self, format, *args):
        """Override para logs mas limpios"""
        print(f"[{self.log_date_time_string()}] {args[0]}")


if __name__ == '__main__':
    print(f"""
╔══════════════════════════════════════════╗
║   Hermes Word Backend Server v0.1.0     ║
╠══════════════════════════════════════════╣
║  Puerto: {PORT}                         ║
║  Health: http://localhost:{PORT}/health ║
║  Chat:   POST http://localhost:{PORT}/chat ║
║  Frontend: http://localhost:{PORT}      ║
╠══════════════════════════════════════════╣
║  Modo: {'LLM (DeepSeek/OpenAI)' if os.environ.get('DEEPSEEK_API_KEY') or os.environ.get('OPENAI_API_KEY') else 'LOCAL (analisis basico)'} ║
╚══════════════════════════════════════════╝
""")
    
    server = http.server.HTTPServer(('0.0.0.0', PORT), WordHermesHandler)
    print(f"Servidor corriendo en http://localhost:{PORT}")
    print("Presiona Ctrl+C para detener\n")
    
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
        server.shutdown()
