======================================================================
HERMES-WORD INTEGRATION: TECHNICAL FEASIBILITY REPORT
======================================================================

1. OFFICE WEB ADD-INS (Task Pane) - The Official Route
----------------------------------------------------------------------
An Office Web Add-in is:
  - HTML5/JS/CSS web app inside embedded browser (Edge WebView2 on Win,
    Safari WebKit on Mac, browser iframe on Word Online)
  - Packaged as .zip with manifest XML; deployed to web server
    (localhost with HTTPS during dev)
  - Manifest declares: task pane dimensions, URLs per platform, permissions
    (ReadWriteDocument minimum)
  - Loads Office.js from CDN: appsforoffice.microsoft.com/lib/1/hosted/office.js
  - Entry point: Office.onReady() then Word.run(async (ctx) => {...})

CRITICAL INSIGHT: The add-in UI runs in a BROWSER context (NOT Node.js).
  - Standard WebSocket (wss://) WORKS - it is a regular browser!
  - fetch() to REST APIs WORKS
  - EventSource (SSE) WORKS
  - Web Workers available
  - NO child_process, NO Node.js modules, NO filesystem access
  - HTTPS/WSS required for production (self-signed OK for localhost dev)

This is EXCELLENT: the add-in can open a WebSocket to a local
Hermes backend and exchange JSON messages natively.

2. PROPOSED ARCHITECTURE (Hybrid: Add-in + Hermes Backend)

  +-----------------------+      wss://        +------------------+
  |   Microsoft Word      |<------------------>|  Hermes Backend  |
  |   +-----------------+ |  JSON protocol     |  (Python)        |
  |   | Task Pane Addin | |                    |  localhost:8443  |
  |   | (HTML/JS/CSS)   | |  REST fallback    |                  |
  |   |                 | |<------------------>|                  |
  |   | - Chat UI       | |                    |  - aiohttp WS    |
  |   | - Office.js     | |                    |  - AI agent      |
  |   | - WS client     | |                    |  - doc context   |
  |   +-----------------+ |                    |                  |
  |          |             |                    +------------------+
  |   Office.js API        |
  |   (read/write doc)     |
  +------------------------+

FLOW:
  1. User types in panel: "Rewrite this paragraph formally"
  2. Add-in JS calls Word.run() to get document text/selection
  3. Add-in sends JSON via WebSocket: {action:"chat", text:"...", msg:"..."}
  4. Hermes processes with LLM, streams response back in chunks
  5. Add-in displays response chunks in chat UI
  6. User clicks "Apply" to insert AI response into document
     (via body.insertText() or selection.insertText())

3. OFFICE.JS WORD API CAPABILITIES (Key for Hermes)
----------------------------------------------------------------------
Min requirement: WordApi 1.1 (available on ALL platforms: Win/Mac/Online/iPad)

READING DOCUMENT:
  context.document.body.getText()       -> plain text
  context.document.body.getHtml()       -> HTML (preserves formatting)
  context.document.body.getOoxml()      -> Office Open XML (full fidelity)
  context.document.getSelection()       -> current selection Range
  body.paragraphs                       -> iterate paragraphs with items
  body.search("term", {matchCase})      -> find text, returns ranges

WRITING/MODIFYING DOCUMENT:
  body.insertText("text", "Start"/"End"/"Replace")
  body.insertParagraph("text", "Start"/"End")
  body.insertHtml("<b>text</b>", "End")
  body.insertOoxml(oxml, "End")
  selection.insertText("text", "Replace")    -> modify current selection
  selection.insertParagraph(...)
  body.clear()                               -> WordApi 1.3+

HIGHER VERSIONS:
  WordApi 1.3: Sections, headers/footers, body.clear()
  WordApi 1.4: Document change events, comment read
  WordApi 1.6: Comment insert (write), comment replies
  WordApi 1.7: Checkbox controls

KEY INSIGHT: WordApi 1.1 provides EVERYTHING needed for basic
chat-assisted document editing. Read text, send to AI, insert
response. Comment insertion (1.6) is optional for AI suggestions.

4. WEBSOCKET CONSTRAINTS AND SOLUTIONS
----------------------------------------------------------------------
CONSTRAINT: Office Add-ins require HTTPS for the add-in source URL
and WSS for WebSocket connections (even to localhost).

SOLUTION:
  a) Use mkcert (github.com/FiloSottile/mkcert) to generate a local CA
     and certificate for localhost. mkcert automatically trusts the CA
     in the system trust store.
  b) Certificates needed:
     - localhost:3000 for add-in dev server (webpack-dev-server)
     - localhost:8443 for Hermes WS backend
  c) One-time setup per user machine; can be automated by installer.

5. PRODUCTION DEPLOYMENT OPTIONS
----------------------------------------------------------------------
OPTION A: Local-only (Recommended for v1)
  - Hermes backend as local process (systemd/launchd/Windows service)
  - Add-in hosted on bundled local static server
  - WebSocket to localhost:8443
  - Best for privacy, latency, offline operation

OPTION B: Cloud backend
  - Hermes backend on cloud server
  - Add-in connects via WSS to remote endpoint with auth tokens
  - Works on Word Online without local install
  - Simpler for users but requires cloud costs

OPTION C: Hybrid (Local with cloud fallback)
  - Try localhost first; if Hermes not running, fall back to cloud

RECOMMENDATION: Option A for v1 (aligns with Hermes as local CLI tool).

6. SIMILAR OPEN-SOURCE PROJECTS
----------------------------------------------------------------------
- GitHub Copilot for Office: Closed-source; uses similar task pane
  architecture with cloud backend
- Microsoft Yo Office templates: Yeoman generator scaffolding for
  task pane add-ins (yo office --projectType taskpane --host word)
- OfficeDev/Office-Add-in-samples: Official Microsoft samples on GitHub
- No open-source AI chat panel specifically for Word found (this is
  a relatively unexplored niche)
- VS Code extension model is conceptually identical: sidebar panel
  communicating with a language server backend

7. DETAILED IMPLEMENTATION PLAN
----------------------------------------------------------------------

PHASE 1: SCAFFOLDING (Week 1)
  - Set up dev environment:
    * Node.js LTS + npm
    * Yeoman + generator-office:
      npm install -g yo generator-office
    * yo office --projectType taskpane --host word --name hermes-word --js
  - Install and configure mkcert for local HTTPS:
    * brew install mkcert (Mac) / choco install mkcert (Win) / apt install mkcert
    * mkcert -install (trusts CA in system store)
    * mkcert localhost 127.0.0.1 ::1 (generates cert+key)
  - Configure webpack-dev-server with HTTPS using mkcert certs
  - Update manifest.xml to point to https://localhost:3000
  - Verify add-in loads in Word Desktop (sideload manifest)

PHASE 2: HERMES BACKEND WRAPPER (Week 1-2)
  - Create hermes_word_server.py:
    * Async WebSocket server using aiohttp or FastAPI
    * Listen on wss://localhost:8443 with mkcert cert
    * Accept JSON messages:
      {
        "id": "uuid",
        "action": "chat",
        "conversation": [{"role": "user"/"assistant", "content": "..."}],
        "doc_context": {
          "full_text": "optional full document text",
          "selection": "currently selected text",
          "surrounding": "paragraphs around cursor"
        },
        "system_prompt": "You are an AI assistant embedded in Word..."
      }
    * Response messages (streamed):
      {"id": "uuid", "type": "chunk", "content": "partial response..."}
      {"id": "uuid", "type": "done", "content": "full text"}
      {"id": "uuid", "type": "error", "content": "error message"}
    * Invoke Hermes agent (subprocess or library call) with doc context
    * Stream LLM output chunks back over WebSocket
  - REST endpoints:
    * GET /health -> {"status": "ok", "version": "1.0.0"}
    * POST /chat -> non-streaming chat (fallback if WS fails)
    * POST /analyze -> full document analysis
  - Configuration via environment variables:
    * HERMES_WORD_PORT=8443
    * HERMES_WORD_HOST=localhost
    * HERMES_WORD_CERT_PATH=/path/to/cert.pem
    * HERMES_WORD_KEY_PATH=/path/to/key.pem

PHASE 3: ADD-IN CORE DEVELOPMENT (Week 2-3)
  - Chat UI components:
    * Message list with user and assistant bubbles
    * Input area with text field and send button
    * "Apply to Document" button on each AI response
    * "Insert at Cursor" / "Replace Selection" / "Append to End" actions
    * Connection status indicator (green/yellow/red dot)
    * Settings panel (Hermes backend URL, API key)
  - WordBridge module (word-bridge.js):
    * getDocumentText(): context.document.body.getText()
    * getSelection(): context.document.getSelection().getText()
    * getContextAroundSelection(): surrounding paragraphs
    * insertAtCursor(text): selection.insertText(text, "Replace")
    * insertAtEnd(text): body.insertText(text, "End")
    * replaceSelection(text): selection.insertText(text, "Replace")
    * clearDocument(): body.clear() [WordApi 1.3+]
  - WebSocket client (ws-client.js):
    * Auto-reconnect with exponential backoff (1s, 2s, 4s, 8s, max 30s)
    * Heartbeat/ping every 30 seconds
    * Message queue for offline resilience (buffer messages when disconnected)
    * Connection timeout detection
  - Document change detection:
    * Listen to Office.EventType.DocumentSelectionChanged
    * Debounce and auto-update doc context sent to Hermes

PHASE 4: COMMUNICATION PROTOCOL (Week 3)
  - Full JSON protocol specification:
    {
      "version": "1.0",
      "messages": {
        "chat_request": {
          "id": "uuid",
          "action": "chat",
          "conversation": [...],
          "doc_context": {...},
          "options": {"max_tokens": 2000, "temperature": 0.7}
        },
        "chat_response_chunk": {
          "id": "uuid",
          "type": "chunk",
          "content": "partial text"
        },
        "chat_response_done": {
          "id": "uuid",
          "type": "done",
          "content": "full response",
          "metadata": {"tokens_used": 150, "model": "gpt-4"}
        },
        "edit_command": {
          "id": "uuid",
          "type": "edit",
          "action": "insert"/"replace"/"comment",
          "target": "selection"/"cursor"/"end",
          "text": "text to apply",
          "description": "human-readable summary"
        },
        "error": {
          "id": "uuid",
          "type": "error",
          "code": "timeout"/"auth"/"rate_limit",
          "message": "human-readable error"
        }
      }
    }
  - Structured edit commands: Hermes can propose edits as {type:"edit"}
    messages. The add-in renders them as "Suggested edit" cards with
    Apply/Dismiss buttons, or auto-applies if user enables auto-apply.

PHASE 5: POLISH AND UX (Week 3-4)
  - Context-aware system prompt that includes document metadata
  - Quick commands in chat:
    * "/rewrite" - rewrite selected text
    * "/summarize" - summarize document
    * "/continue" - continue writing from cursor
    * "/format" - improve formatting
    * "/translate" - translate selection
  - Markdown rendering in chat bubbles (AI uses markdown)
  - Code block syntax highlighting
  - Document outline awareness (read headings for context)
  - Dark mode support (follow Word's theme)
  - Error handling: show user-friendly errors
  - Loading states: typing indicator, spinner for long ops

PHASE 6: PACKAGING AND DISTRIBUTION (Week 4+)
  - Production build optimization (webpack minification)
  - Add-in packaging:
    * npm run build produces dist/ folder
    * Zip dist/ + manifest.xml -> hermes-word-addin.zip
  - Installer script for each platform:
    * Linux: shell script + systemd service
    * macOS: .pkg installer + launchd plist
    * Windows: NSIS/MSI installer + Windows service
  - Installer steps:
    1. Install Python dependency: pip install hermes-word-server
    2. Run mkcert -install (trust local CA)
    3. Generate localhost certs
    4. Copy add-in to local static server directory
    5. Copy manifest to Word's shared folder / registry
    6. Register Hermes as system service
    7. Launch service
  - Alternative: Side-loading via network share manifest (enterprise)
  - Alternative: Office Store submission (public distribution)

8. LIMITATIONS AND MITIGATIONS
----------------------------------------------------------------------

LIMITATION 1: Document size
  - Full document can be 100+ pages, too large for LLM context windows
  - Mitigation: Default to sending only selection + surrounding 5
    paragraphs. Offer "Analyze full document" with chunked processing
    (split doc, process segments, aggregate).

LIMITATION 2: HTTPS/WSS cert management
  - Self-signed certs are painful for non-technical users
  - Mitigation: One-click setup via mkcert in installer; clear docs;
    consider bundled cert approach.

LIMITATION 3: Office.js batch API
  - Word.run() executes async batches; no streaming edits
  - Mitigation: Buffer edits; apply in single Word.run() on "Apply".
    Show streaming response in chat UI only.

LIMITATION 4: No undo grouping
  - Office.js edits create separate undo steps; AI batch edit
    cannot be undone as one operation
  - Mitigation: Use tracked changes (WordApi 1.3+) or comments
    (WordApi 1.6+) so user can review before accepting.

LIMITATION 5: Word Online rate limiting
  - Stricter API call limits on web version
  - Mitigation: Batch calls aggressively; debounce; show progress.

LIMITATION 6: Platform CSS differences
  - Mac = Safari WebKit, Win = Edge WebView2, Online = host browser iframe
  - Mitigation: Test on all platforms; use conservative CSS; prefer
    Office UI Fabric (Fluent UI) for native look.

LIMITATION 7: Authentication
  - If cloud backend: need auth tokens
  - If local only: no auth needed (local trust boundary)
  - Mitigation: v1 = local only (no auth). v2 = add token-based auth.

9. VIABILITY VERDICT
----------------------------------------------------------------------

FEASIBLE: YES - with high confidence (9/10)

The hybrid architecture (Office Web Add-in task pane + local Hermes
WebSocket backend) is technically sound and well-supported:

  + Task pane add-in: Official, documented, stable Microsoft API
  + Office.js Word read/write: Sufficient (WordApi 1.1 minimum)
  + WebSocket from browser: Native, no special permissions
  + Python async WS server: Mature ecosystem (aiohttp, FastAPI, etc.)
  + Hermes agent: Already exists; needs thin WS wrapper layer
  + Cross-platform: Win/Mac/Online via same add-in code

No insurmountable technical blockers identified. Main engineering
challenges are:
  1. Local HTTPS cert setup (well-solved by mkcert)
  2. Cross-platform installer (standard desktop app problem)
  3. Large document handling (chunking + context window mgmt)

The overall effort is estimated at 3-4 weeks for a working MVP
with core chat capabilities (Phase 1-4), with another 2-4 weeks
for polish, packaging, and distribution.

======================================================================
