HERMES-WORD: ARCHITECTURE DIAGRAMS
=====================================

ARCHITECTURE OVERVIEW
---------------------

                        USER
                         |
             +-----------v-----------+
             |   Microsoft Word      |
             |  (Win / Mac / Online) |
             |                       |
             |  +-----------------+  |
             |  |  Task Pane      |  |
             |  |  (HTML/JS/CSS)  |  |
             |  |                 |  |
             |  | +-------------+ |  |
             |  | | ChatApp     | |  |   WebSocket (wss://)
             |  | | - MsgList   | |  |   JSON protocol
             |  | | - Input     | |  |        |
             |  | | - Status    | |<-+--------+
             |  | +-------------+ |  |        |
             |  |                 |  |   +----v---------------+
             |  | +-------------+ |  |   |  Hermes Backend    |
             |  | | WordBridge  |-+-+--->  (Python)           |
             |  | | (Office.js) | |  |   |                    |
             |  | +-------------+ |  |   |  localhost:8443    |
             |  +--------|--------+  |   |                    |
             |           |           |   |  +---------------+ |
             |    Office.js API      |   |  | aiohttp Server| |
             |    (Read/Write doc)   |   |  | - /ws         | |
             +-----------------------+   |  | - /api/health | |
                                          |  | - /api/chat   | |
                                          |  +---------------+ |
                                          |         |          |
                                          |  +-----v--------+  |
                                          |  | Hermes Agent |  |
                                          |  | (LLM calls)  |  |
                                          |  +--------------+  |
                                          +--------------------+

PLATFORM SUPPORT MATRIX
------------------------
Feature               Win Desktop  Mac Desktop  Word Online  iPad
Task Pane                Y            Y            Y          Y*
Office.js Word APIs    1.1-1.7      1.1-1.7      1.1-1.7    1.1-1.3
WebSocket (wss://)       Y            Y            Y          Y
Local backend (locahost) Y            Y            N**        N**
Read document            Y            Y            Y          Y
Write document           Y            Y            Y          Y
Tracked changes          Y            Y            Y          Y
Comments (insert)        Y            Y            Y          N

* iPad: Task pane limited; width constraints
** Word Online/iPad: Cannot reach localhost; need cloud backend
   or browser extension bridge. Target Desktop for v1.
