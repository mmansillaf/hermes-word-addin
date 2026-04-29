HERMES-WORD WEBSOCKET PROTOCOL SPECIFICATION v1.0
============================================================

TRANSPORT: WSS (WebSocket Secure)
DEFAULT URL: wss://localhost:8443/ws
CONTENT TYPE: JSON text frames
ENCODING: UTF-8

MESSAGE FORMAT
==============
All messages are JSON objects with a required "id" (UUID v4 string)
and "action" (client->server) or "type" (server->client).

CLIENT -> SERVER MESSAGES
==========================

1. CHAT REQUEST
{
  "id": "550e8400-e29b-41d4-a716-446655440000",
  "action": "chat",
  "conversation": [
    {"role": "user", "content": "Hello"},
    {"role": "assistant", "content": "Hi! How can I help?"}
  ],
  "doc_context": {
    "full_text": "optional full document text...",
    "selection": "currently selected text",
    "surrounding_before": "text before selection",
    "surrounding_after": "text after selection",
    "metadata": {
      "paragraph_count": 42,
      "word_count": 1500,
      "title": "My Document"
    }
  },
  "options": {
    "max_tokens": 2000,
    "temperature": 0.7
  }
}

2. CANCEL REQUEST
{
  "id": "...",
  "action": "cancel",
  "request_id": "550e8400-..."
}

3. PING
{
  "id": "...",
  "action": "ping"
}

4. GET DOCUMENT ANALYSIS
{
  "id": "...",
  "action": "analyze",
  "doc_context": { ... },
  "analysis_type": "summarize" | "outline" | "key_points" | "grammar_check"
}

5. EDIT REQUEST (request structured edit)
{
  "id": "...",
  "action": "request_edit",
  "edit_type": "rewrite" | "format" | "expand" | "shorten" | "translate",
  "target_text": "text to edit",
  "instructions": "make it more formal",
  "doc_context": { ... }
}

SERVER -> CLIENT MESSAGES
==========================

1. RESPONSE CHUNK (streaming)
{
  "id": "...",
  "type": "chunk",
  "request_id": "550e8400-...",
  "content": "partial response text"
}

2. RESPONSE COMPLETE
{
  "id": "...",
  "type": "done",
  "request_id": "550e8400-...",
  "content": "full response text",
  "metadata": {
    "model": "deepseek-v4-pro",
    "tokens_used": 150,
    "processing_time_ms": 2300
  },
  "edit": {                          // optional structured edit
    "action": "replace",
    "target": "selection",
    "text": "edited text to apply"
  }
}

3. ERROR
{
  "id": "...",
  "type": "error",
  "request_id": "550e8400-...",
  "code": "timeout" | "auth" | "rate_limit" | "invalid_request" | "server_error",
  "message": "Human-readable error description"
}

4. PONG
{
  "id": "...",
  "type": "pong"
}

5. SYSTEM NOTIFICATION
{
  "id": "...",
  "type": "notification",
  "level": "info" | "warning" | "error",
  "message": "Connection to AI service established"
}

RECONNECTION BEHAVIOR
=====================
- Client initiates connection to wss://localhost:8443/ws
- On disconnect: exponential backoff (1s, 2s, 4s, 8s, 16s, 30s cap)
- Heartbeat: client sends ping every 30s; server must respond pong within 10s
- Server may close idle connections after 5 minutes
- Message IDs allow deduplication on reconnect (server should cache last N responses)

FALLBACK: REST API
==================
When WebSocket is unavailable, the add-in falls back to HTTPS REST:

POST https://localhost:8443/api/chat
Content-Type: application/json
{
  "conversation": [...],
  "doc_context": {...},
  "options": {...}
}
Response: { "id": "...", "content": "response text", "metadata": {...} }

POST https://localhost:8443/api/analyze
Content-Type: application/json
{ "doc_context": {...}, "analysis_type": "summarize" }
Response: { "id": "...", "content": "analysis result" }

GET https://localhost:8443/api/health
Response: { "status": "ok", "version": "1.0.0", "uptime_seconds": 3600 }
