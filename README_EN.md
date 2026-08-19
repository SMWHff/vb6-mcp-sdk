# vb6-mcp-sdk — VB6.0 MCP Server Development Framework

[中文](README.md) | **English**

![License](https://img.shields.io/badge/license-MIT-blue)
![Language](https://img.shields.io/badge/language-VB6%20(32--bit)-brightgreen)
![Platform](https://img.shields.io/badge/platform-Windows-0078d6)
![Protocol](https://img.shields.io/badge/MCP-2024--11--05-purple)
![Transport](https://img.shields.io/badge/transport-stdio%20%2F%20Streamable%20HTTP-orange)

An **MCP (Model Context Protocol) server development framework written in VB6 (32-bit)**. It encapsulates protocol handling (JSON-RPC 2.0), dual transports (stdio / Streamable HTTP), JSON parsing, UTF-8 encoding/decoding, and logging, covering **MCP's three capabilities (Tools / Prompts / Resources)** — **implement the interfaces, register them, and you have a working MCP server**.

```text
Framework = protocol engine (McpServer) + dual transports (stdio/HTTP) + three capability interfaces
You write = tool classes (Implements ITool) + prompt classes (Implements IPrompt)
          + resource classes (Implements IResource) + entry assembly (mcp_main.bas)
```

## Features

- 🧩 **Protocol engine**: Full JSON-RPC 2.0 + MCP lifecycle (initialize / notifications / ping), compatible with the 2024-11-05 revision and 2025-series clients
- 🔌 **Dual transports**: stdio (spawned by Claude Desktop / Cursor, etc.) and Streamable HTTP (self-implemented with Winsock, zero third-party dependencies)
- 🧰 **Three capabilities**: Tools / Prompts / Resources fully supported — implement an interface, register it, done
- 🔤 **Chinese-friendly**: UTF-8 end-to-end encoding; ASCII tool names with Chinese descriptions and results
- 🛡️ **Error semantics**: tool errors automatically become `isError` results; missing required arguments are validated automatically (-32602) so the AI client sees readable error text
- 📦 **Zero dependencies**: VB6 + Win32 API + MSScriptControl only, no third-party VB6 controls
- 🧪 **Testable**: 33-case stdio test suite + 13 HTTP-specific checks, verified against the official Python SDK over both transports

---

## Directory Structure

```
vb6-mcp-sdk\
├── sdk\                        ← SDK core (reuse, usually no changes needed)
│   ├── ITool.cls               ← Tool interface
│   ├── IPrompt.cls             ← Prompt template interface
│   ├── IResource.cls           ← Resource interface
│   ├── McpServer.cls           ← Server class: register/dispatch the three capabilities, start transports
│   ├── mcp_transport_stdio.bas ← stdio transport (kernel32 framing + UTF-8)
│   ├── mcp_transport_http.bas  ← Streamable HTTP transport (Winsock API)
│   ├── mcp_json.bas            ← JSON utilities (JsonInit/JsonGet/JsonQuote)
│   └── mcp_log.bas             ← Logging (logs\mcp.log)
├── tools\                      ← Put your capabilities here (with examples)
│   ├── ToolAdd.cls             ← Example tool: arithmetic
│   ├── ToolEcho.cls            ← Example tool: strings
│   ├── ToolGetTime.cls         ← Example tool: time
│   ├── ToolSysInfo.cls         ← Example tool: system info (Windows API)
│   ├── ToolReadFile.cls        ← Example tool: whitelist-secured file reading
│   ├── ToolWordCount.cls       ← Example tool: text statistics (demonstrates isError)
│   ├── SamplePrompt.cls        ← Example prompt (code review assistant)
│   └── SampleResource.cls      ← Example resource (server info)
├── mcp_main.bas                   ← Entry point: create server, register capabilities, start
├── vb6-mcp-sdk.vbp              ← Project file (double-click to open)
├── json-polyfill.js            ← Runtime dependency (must sit next to the exe)
├── README.md                   ← 中文文档
├── README_EN.md                ← English docs (this file)
├── run_test.bat                ← one-click launcher (auto-compile VB6 / self-install deps / full tests / logs in logs\)
├── scripts\
│   └── fix-console.ps1         ← Must run after every compile (GUI→Console subsystem)
└── tests\
    ├── test.ps1                ← stdio smoke test
    ├── test-suite.py           ← ★ full stdio test suite (33 cases)
    ├── http-test.py            ← ★ HTTP transport tests (13 checks)
    ├── client-sdk.py           ← official Python SDK handshake (stdio)
    └── client-sdk-http.py      ← official Python SDK over HTTP
```

---

## Installation

```bash
# Option 1: git clone
git clone git@github.com:SMWHff/vb6-mcp-sdk.git

# Option 2: download the ZIP (GitHub page → Code → Download ZIP)
```

Double-click `vb6-mcp-sdk.vbp` to open the project in the VB6 IDE; use **File → Make vb6-mcp-sdk.exe** to build the executable (details in Quick Start below).

> 💡 This framework is designed for **secondary development**: write your own capability classes under `tools/`, register them in `mcp_main.bas`, and compile your own MCP server. The bundled example tools (arithmetic / time / file reading, etc.) work out of the box for a quick try.

---

## Quick Start (Build an MCP server in 3 steps)

### Step 1: Implement the ITool interface (create a new class module)

```vb
' Class name: ToolGreeting (Project → Add Class Module, name it ToolGreeting)
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "greeting"          ' Tool name (unique, ASCII)
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "Generate a greeting"   ' Description (shown in tools/list)
End Property

Private Property Get ITool_InputSchema() As String
    ' Argument schema (JSON string): tells the client what arguments this tool needs
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""name"":{""type"":""string""}}," _
        & """required"":[""name""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    ' requestJson is the full request; use JsonGet to read arguments
    Dim name As String
    name = CStr(JsonGet(requestJson, ".params.arguments.name"))
    ITool_Execute = "Hello, " & name & "!"
End Function
```

### Step 2: Register the tool (edit mcp_main.bas)

```vb
' Add one line inside Main in mcp_main.bas:
server.RegisterTool New ToolGreeting
```

### Step 3: Compile & start

1. Double-click `vb6-mcp-sdk.vbp` → **File → Make vb6-mcp-sdk.exe** into the project root
2. `pwsh .\scripts\fix-console.ps1` (**required after every recompile**)
3. Start:
   - stdio: spawned by an MCP client (Claude Desktop, etc.)
   - HTTP: `.\vb6-mcp-sdk.exe /http` (port 8080) or `/http:9000`

Done. Your tool is now a standard MCP server.

---

## ITool Interface

| Member | Description |
|---|---|
| `ToolName` | Tool name, used by clients to call it; must be unique (ASCII) |
| `Description` | Tool description, helps the AI decide when to call it (Chinese OK) |
| `InputSchema` | JSON Schema for arguments; the client generates call arguments from it |
| `Execute(requestJson)` | Executes the tool; returns result text; raise via `Err.Raise` on invalid arguments |

**Reading arguments in Execute**: use the framework's `JsonGet(requestJson, ".params.arguments.xxx")` — it returns scalars (number / string / boolean). Tools never deal with transports or encoding — the framework handles UTF-8.

**Errors**: just `Err.Raise vbObjectError + 1, , "argument a must be a number"`; the framework converts it into a JSON-RPC error (-32602) for the client.

**Automatic argument validation**: before dispatching `tools/call`, the framework parses the `required` array of the tool's `InputSchema` and checks that all required arguments are present in `arguments` — **missing required arguments return 「缺少必需参数: xxx」 directly without invoking the tool**. Tools only need to validate types/business rules; no duplicate existence checks.

**isError semantics (automatic)**: when a tool raises (`Err.Raise`), the framework returns an MCP-conformant result `{"content":[...],"isError":true}` instead of a JSON-RPC error — **the AI client sees the error text** (e.g. 「执行失败: 非法路径」), which is friendlier than a protocol-level error. Missing arguments / unknown tools still go through JSON-RPC error (-32602).

---

## The Three Capability Interfaces

MCP defines three capabilities; the SDK provides one interface for each, all registered via `server.RegisterXxx`:

| Capability | Interface | Registration | Client calls |
|---|---|---|---|
| Tools | `ITool` | `RegisterTool New YourClass` | `tools/list`, `tools/call` |
| Prompts | `IPrompt` | `RegisterPrompt New YourClass` | `prompts/list`, `prompts/get` |
| Resources | `IResource` | `RegisterResource New YourClass` | `resources/list`, `resources/read` |

### IPrompt (prompt template)

```vb
Option Explicit
Implements IPrompt

Private Property Get IPrompt_PromptName() As String
    IPrompt_PromptName = "code_review"              ' Prompt name (unique, ASCII)
End Property

Private Property Get IPrompt_PromptDescription() As String
    IPrompt_PromptDescription = "Generate code review comments"   ' Description
End Property

Private Property Get IPrompt_PromptArguments() As String
    ' Argument declaration (JSON array string)
    IPrompt_PromptArguments = "[{""name"":""language"",""required"":true}]"
End Property

Private Function IPrompt_GetPromptText(ByVal requestJson As String) As String
    Dim lang As String
    lang = CStr(JsonGet(requestJson, ".params.arguments.language"))
    IPrompt_GetPromptText = "Please review the following " & lang & " code:"
End Function
```

### IResource (resource)

```vb
Option Explicit
Implements IResource

Private Property Get IResource_ResourceUri() As String
    IResource_ResourceUri = "demo://server/info"    ' Unique URI (standard format)
End Property

Private Property Get IResource_ResourceName() As String
    IResource_ResourceName = "Server info"
End Property

Private Property Get IResource_ResourceMimeType() As String
    IResource_ResourceMimeType = "text/plain"
End Property

Private Function IResource_ReadResourceText(ByVal uri As String) As String
    IResource_ReadResourceText = "dynamically generated content..."
End Function
```

Reference implementations: `tools\SamplePrompt.cls` (code review assistant), `tools\SampleResource.cls` (server info).

---

## Entry Point: mcp_main.bas

```vb
Public Sub Main()
    Dim server As McpServer
    Set server = New McpServer

    server.SetServerInfo "my-mcp-server", "1.0.0"   ' optional: server identification

    ' Register your tools (add as many as you like)
    server.RegisterTool New ToolGreeting
    server.RegisterTool New ToolAdd

    ' Register prompts and resources (optional)
    server.RegisterPrompt New SamplePrompt
    server.RegisterResource New SampleResource

    ' Pick the transport based on the command line
    If LCase$(Command$) Like "*http*" Then
        server.RunHttp ParsePort(Command$)
    Else
        server.RunStdio
    End If
End Sub
```

---

## Integrating with MCP Clients (stdio)

Once compiled, register the exe in any MCP client.

**Claude Desktop** (Windows; config file `%APPDATA%\Claude\claude_desktop_config.json`):

```json
{
  "mcpServers": {
    "vb6-mcp-sdk": {
      "command": "C:\\path\\to\\vb6-mcp-sdk.exe"
    }
  }
}
```

**Cursor**: `Settings → MCP → Add new MCP server`, Type `command`, Command = the full path to the exe (e.g. `C:\path\to\vb6-mcp-sdk.exe`).

> 📌 stdio mode starts by default and needs no arguments; the exe automatically speaks the stdin/stdout framed protocol when spawned by a client.

---

## HTTP Transport Mode (Streamable HTTP)

```powershell
# Default port 8080
.\vb6-mcp-sdk.exe /http

# Custom port
.\vb6-mcp-sdk.exe /http:9000
```

- Client endpoint: `http://localhost:9000/mcp`
- **CORS** supported — browser / web MCP clients can call it cross-origin
- Transport is self-implemented with raw Winsock API; no third-party HTTP component
- Use `tests/client-sdk-http.py` (official Python SDK) for quick integration testing

---

## Verification & Testing

**One-click**: `.\run_test.bat` — auto-checks/installs uv, auto-compiles with VB6 when the exe is missing, runs fix-console, then executes all 5 test groups; the full output goes to the console and a `logs\run_test_*.log` file at the same time. Exit codes: 0=all passed / 1=failures / 2=environment not ready.

Or run each test individually:

```powershell
# 1) Smoke test (stdio, 5 messages)
pwsh .\tests\test.ps1

# 2) ★ Full test suite (stdio, 33 cases: handshake/tools/prompts/resources/security/raw protocol errors)
uv run --with mcp python .\tests\test-suite.py

# 3) ★ HTTP transport tests (13 checks: CORS/202/404/OPTIONS/Chinese/missing args)
#    Start first: .\vb6-mcp-sdk.exe /http:9002
uv run python .\tests\http-test.py http://localhost:9002/mcp

# 4) Official Python SDK full handshake (stdio)
uv run --with mcp python .\tests\client-sdk.py

# 5) Official Python SDK over HTTP
uv run --with mcp python .\tests\client-sdk-http.py http://localhost:9000/mcp

# 6) Official Inspector (run in the project root)
npx @modelcontextprotocol/inspector .\vb6-mcp-sdk.exe
```

**Test suite coverage** (`test-suite.py`, 33 cases): handshake 4 · tools 14 (negative numbers / decimals / special characters / 10KB long text / unknown tool / missing args / isError) · prompts 3 · resources 3 · security 3 (path traversal / absolute path / illegal extension) · raw protocol 6 (invalid JSON / unknown method / notification without response / string id / numeric id / CRLF). **The raw-protocol cases catch bugs the official SDK client cannot** (e.g. string ids without quotes, `\uXXXX` escape parsing) — they have already surfaced and fixed framework bugs twice.

---

## Architecture

```
┌──────────────────────────────────────────────────┐
│ Your code (SDK untouched)                         │
│  mcp_main.bas → register capabilities → RunStdio / RunHttp
│  tool classes (Implements ITool)                  │
│  prompt classes (Implements IPrompt)              │
│  resource classes (Implements IResource)          │
├──────────────────────────────────────────────────┤
│ McpServer.cls (protocol engine, MCP capabilities) │
│  tools:    /list /call                            │
│  prompts:  /list /get                             │
│  resources:/list /read                            │
│  capability dispatch via interface polymorphism   │
├──────────────────┬───────────────────────────────┤
│ stdio transport  │ Streamable HTTP transport      │
│ mcp_transport_   │ mcp_transport_http.bas         │
│ stdio.bas        │ (Winsock API)                  │
├──────────────────┴───────────────────────────────┤
│ mcp_json.bas (JsonGet/JsonQuote)                  │
│ mcp_transport_stdio.bas (Utf8Encode/Decode)       │
│ mcp_log.bas                                       │
└──────────────────────────────────────────────────┘
```

---

## Requirements & Dependencies

- VB6 IDE (32-bit compilation)
- 64-bit systems: `regsvr32 C:\Windows\SysWOW64\msscript.ocx` (as administrator, MSScriptControl)
- `json-polyfill.js` must sit next to the exe (read at runtime; keep pure ASCII, do not change encoding)

---

## Known Pitfalls (lessons learned)

| Symptom | Root cause | Fix |
|---|---|---|
| Compile error 「JsonInit 未定义」 | Module not in project / .bas has LF line endings | Double-click .vbp to load; convert all files to CRLF |
| Compile error 「缺少标识符」 | API declaration parameter collides with a keyword (`len`) | Rename to `buflen`; avoid len/name/type/error |
| Compile error 「XX 未定义」 (cross-module) | `Private Declare` not visible across modules | Use `Public Declare` when called from other modules |
| Runtime error 62 (input past end of file) | Text-mode `Input$(LOF(f))` reading CRLF | Read the polyfill in binary mode |
| All tool calls return 438 | late binding into JS nested objects | Use `JsonGet` to fetch scalars on the JS side |
| No output in pwsh pipeline | exe is a GUI subsystem | Run `fix-console.ps1` |
| All HTTP requests return 404 | Header extraction took only the first line / 0-byte placeholder pollution | Slice `sepPos+4` for the full header; track length with `accLen` |

---

## Relationship with vb6mcp (Monolithic Version)

`vb6mcp` is the predecessor of this framework — protocol, transports, and tools were all crammed into a single `mcp.bas`. `vb6-mcp-sdk` refactors it into a reusable layered framework: **the protocol engine is decoupled from concrete tools; adding a tool only requires implementing ITool and registering it**.

---

## License & Contributing

Repository: [github.com/SMWHff/vb6-mcp-sdk](https://github.com/SMWHff/vb6-mcp-sdk) — if you find it useful, ⭐ Star it, or open an [Issue](https://github.com/SMWHff/vb6-mcp-sdk/issues) / [PR](https://github.com/SMWHff/vb6-mcp-sdk/pulls).

- **License**: MIT (see `LICENSE`)
- **Version**: 1.0.0 (protocol 2024-11-05; compatible with 2025-03-26 / 2025-06-18 clients)
- **Tech stack**: VB6 (32-bit) + Win32 API + MSScriptControl (JSON parsing) — no third-party VB6 controls
- **Contributing**:
  - Add example tools: implement `ITool` / `IPrompt` / `IResource`, follow the templates under `tools/`
  - Fix bugs: make sure `tests/test-suite.py` (33 cases) + `tests/http-test.py` (13 checks) pass before submitting
  - Commit style: Conventional Commits (`feat(scope): description`)
- **Testing**: full verification against the official Python SDK over both transports (stdio + Streamable HTTP) plus raw-protocol error cases (which catch what official clients miss)
- **Related resources**: [MCP specification](https://modelcontextprotocol.io/) · [Python SDK](https://github.com/modelcontextprotocol/python-sdk) · [Inspector](https://github.com/modelcontextprotocol/inspector)
