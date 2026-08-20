# vb6-mcp-sdk —— VB6.0 MCP Server 开发框架

**中文** | [English](README_EN.md)

![License](https://img.shields.io/badge/license-MIT-blue)
![Language](https://img.shields.io/badge/language-VB6%20(32--bit)-brightgreen)
![Platform](https://img.shields.io/badge/platform-Windows-0078d6)
![Protocol](https://img.shields.io/badge/MCP-2024--11--05-purple)
![Transport](https://img.shields.io/badge/transport-stdio%20%2F%20Streamable%20HTTP-orange)

一个用 **VB6（32 位）** 编写的 MCP（Model Context Protocol）服务器开发框架。封装了协议处理（JSON-RPC 2.0）、双传输（stdio / Streamable HTTP）、JSON 解析、UTF-8 编解码、日志，覆盖 **MCP 三大能力（Tools / Prompts / Resources）**——**你只需要实现接口、注册进去，就是一个可用的 MCP server**。

```text
本框架 = 协议引擎（McpServer）+ 双传输（stdio/HTTP）+ 三大能力接口
你写   = 工具类（Implements ITool）+ 提示词类（Implements IPrompt）
        + 资源类（Implements IResource）+ 入口组装（mcp_main.bas）
```

## 特性

- 🧩 **协议引擎**：完整 JSON-RPC 2.0 + MCP 生命周期（initialize / notifications / ping），**协议版本协商**（支持 2024-11-05 / 2025-03-26 / 2025-06-18 / 2025-11-25），含 **completion/complete 参数补全、logging/setLevel、resources 订阅、structuredContent 结构化结果**
- 🔌 **双传输**：stdio（被 Claude Desktop / Cursor 等拉起）与 Streamable HTTP（Winsock 自实现，零第三方依赖）
- 🧰 **四大能力**：Tools / Prompts / Resources / **Resource Templates** 全覆盖，实现接口 + 注册即用
- 🔤 **中文友好**：UTF-8 全链路编解码，工具名用 ASCII、描述与结果可中文
- 🛡️ **错误语义**：工具抛错自动转 `isError` 结果、缺参自动校验（-32602），AI 端可见错误文本
- 📦 **零依赖**：仅 VB6 + Win32 API + 纯 VB6 JSON 库（VBJSON），无任何第三方 VB6 控件
- 🧪 **可测试**：85 用例 stdio 测试套件 + 33 项 HTTP 专项，官方 Python SDK 双传输验证

---

## 目录结构

```
vb6-mcp-sdk\
├── sdk\                        ← SDK 核心（复用，一般不需要改）
│   ├── ITool.cls               ← 工具接口
│   ├── IPrompt.cls             ← 提示词模板接口
│   ├── ITemplate.cls            ← 资源模板接口（动态 URI，如 demo://greet/{name}）
│   ├── IResource.cls            ← 资源接口
│   ├── McpServer.cls            ← 服务器类：注册/分派四大能力、双传输启动
│   ├── mcp_transport_stdio.bas ← stdio 传输层（kernel32 分帧 + UTF-8）
│   ├── mcp_transport_http.bas  ← Streamable HTTP 传输层（Winsock API）
│   ├── mcp_json.bas            ← JSON 工具（JsonGet/JsonQuote/JsonBuild，纯 VB6 VBJSON）
│   ├── json\                   ← VBJSON 库（JSON.bas + cStringBuilder.cls）
│   └── mcp_log.bas             ← 日志（logs\mcp.log）
├── tools\                      ← 你的能力放这里（含示例）
│   ├── ToolAdd.cls             ← 示例工具：算术
│   ├── ToolEcho.cls            ← 示例工具：字符串
│   ├── ToolGetTime.cls         ← 示例工具：时间
│   ├── ToolSysInfo.cls         ← 示例工具：系统信息（Windows API）
│   ├── ToolReadFile.cls        ← 示例工具：白名单安全读文件
│   ├── ToolWordCount.cls       ← 示例工具：文本统计（演示 isError）
│   ├── ToolRand.cls            ← 示例工具：随机整数（演示数字参数/边界校验）
│   ├── ToolText.cls            ← 示例工具：字符串处理（枚举参数 upper/lower/trim/reverse）
│   ├── ToolEnv.cls             ← 示例工具：环境变量查询（可选默认值）
│   ├── ToolMes.cls             ← MES 开放 API 工具（18 个接口：SN/MAC/料号/工单/治具/烧录文件，见教程文档）
│   ├── SamplePrompt.cls        ← 示例提示词（代码审查助手）
│   ├── SampleTranslate.cls     ← 示例提示词（翻译助手）
│   ├── SampleResource.cls      ← 示例资源（服务器信息）
│   └── SampleTemplate.cls      ← 示例资源模板（demo://greet/{name} 动态问候）
├── mcp_main.bas                   ← 入口：创建 server、注册能力、启动
├── vb6-mcp-sdk.vbp              ← 工程文件（双击打开）
├── README.md                   ← 中文文档
├── README_EN.md                ← English docs
├── run_test.bat                ← 一键测试启动器（自动编译 VB6 / 自装依赖 / 全量测试 / 日志输出 logs\）
├── scripts\                    ← 构建 / 部署脚本
│   ├── fix-console.ps1         ← 编译后必跑（GUI→Console 子系统）
│   ├── fix-vb6ext.ps1          ← 修复 VB6「不能 VB6EXT.OLB 注册」弹窗
│   └── utf8-to-gbk.ps1         ← VB6 源码 UTF-8（含 BOM）转 GBK，防中文乱码
└── tests\                      ← 测试用例
    ├── test.ps1                ← stdio 冒烟测试（10 条消息 / 13 项结构化断言）
    ├── smoke-sdk.py            ← ★ 官方 Python SDK 冒烟（参考 SDK 测试断言风格，19 项，双传输）
    ├── test-suite.py           ← ★ stdio 全面测试套件（85 用例）
    ├── http-test.py            ← ★ HTTP 传输层专项（33 项：会话/边界/三能力全链路/SSE）
    ├── coverage.py             ← ★ 测试用例覆盖率自动统计（工具/提示词/资源/模板/协议方法）
    ├── coverage-whitebox.py    ← ★ 白盒覆盖率自动统计（VB6 源码过程级调用图，仅自研代码，VBJSON 第三方库排除）
    ├── claude-mcp-test.py      ← ★ Claude Code（Anthropic CLI）MCP 兼容性：注册/连接/headless 调工具
    ├── dsh-mcp-test.py         ← ★ DeepSeek Harness（DSH）MCP 兼容性：生成配置块 + 按 dsh-mcp-client 方式握手
    ├── client-sdk.py           ← 官方 Python SDK 握手（stdio）
    └── client-sdk-http.py      ← 官方 Python SDK 走 HTTP
```

---

## 安装与获取

```bash
# 方式一：git clone
git clone git@github.com:SMWHff/vb6-mcp-sdk.git

# 方式二：直接下载 ZIP（GitHub 页面 → Code → Download ZIP）
```

拿到源码后双击 `vb6-mcp-sdk.vbp` 即可在 VB6 IDE 中打开工程；「文件 → 生成 vb6-mcp-sdk.exe」编译出可执行文件（详见下方快速开始）。

> 💡 本框架面向**二次开发**：你在 `tools/` 下写自己的能力类并注册进 `mcp_main.bas`，再编译成你自己的 MCP server。仓库里的示例工具（算术/时间/读文件等）即开即用，可直接用来体验。

---

## 快速开始（三步开发一个 MCP server）

### 第 1 步：实现 ITool 接口（新建一个类模块）

```vb
' 类名：ToolGreeting（新建 → 类模块，命名为 ToolGreeting）
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "greeting"          ' 工具名（唯一，ASCII）
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "生成问候语"      ' 描述（可中文，显示在 tools/list）
End Property

Private Property Get ITool_InputSchema() As String
    ' 参数 schema（JSON 字符串），告诉客户端这个工具要什么参数
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""name"":{""type"":""string""}}," _
        & """required"":[""name""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    ' requestJson 是完整请求，用 JsonGet 取参数
    Dim name As String
    name = CStr(JsonGet(requestJson, ".params.arguments.name"))
    ITool_Execute = "你好，" & name & "！"
End Function
```

### 第 2 步：注册工具（改 mcp_main.bas）

```vb
' mcp_main.bas 的 Main 里加一行：
server.RegisterTool New ToolGreeting
```

### 第 3 步：编译 + 启动

1. 双击 `vb6-mcp-sdk.vbp` → 「文件 → 生成 vb6-mcp-sdk.exe」到项目根目录
2. `pwsh .\scripts\fix-console.ps1`（**每次重编译后必跑**）
3. 启动：
   - stdio：被 MCP 客户端（Claude Desktop 等）拉起
   - HTTP：`.\vb6-mcp-sdk.exe /http`（端口 8080）或 `/http:9000`

完成。你的工具已经是一个标准的 MCP server。

---

## ITool 接口详解

| 成员 | 说明 |
|---|---|
| `ToolName` | 工具名，客户端用这个名字调用，必须唯一（ASCII） |
| `Description` | 工具描述，帮助 AI 判断何时调用（可中文） |
| `InputSchema` | 参数 JSON Schema，客户端按它生成调用参数 |
| `Execute(requestJson)` | 执行工具；返回结果文本；参数非法时 `Err.Raise` 抛错 |

**在 Execute 里取参数**：用框架提供的 `JsonGet(requestJson, ".params.arguments.xxx")`，返回标量（数字/字符串/布尔）。工具内部无需关心传输层和编码——框架已经处理了 UTF-8。

**工具出错**：直接 `Err.Raise vbObjectError + 1, , "参数 a 必须为数字"`，框架会转成 JSON-RPC error（-32602）返回给客户端。

**框架参数校验（自动）**：`tools/call` 分派前，框架会解析工具的 `InputSchema` 里的 `required` 数组，检查请求的 `arguments` 是否齐全——**缺少必需参数时直接返回「缺少必需参数: xxx」**，不会调用工具。工具内部只需校验类型/业务规则，无需重复写存在性检查。

**isError 语义（自动）**：工具执行时抛错（`Err.Raise`），框架按 MCP 规范返回 `{"content":[...],"isError":true}` 的结果而非 JSON-RPC error——**AI 客户端能看到错误文本**（如「执行失败: 非法路径」），比协议层错误更友好。缺参/未知工具仍走 JSON-RPC error（-32602）。

---

## 四大能力接口

MCP 定义了核心能力，SDK 各提供一个接口，全部通过 `server.RegisterXxx` 注册：

| 能力 | 接口 | 注册 | 客户端调用 |
|---|---|---|---|
| Tools 工具 | `ITool` | `RegisterTool New 你的类` | `tools/list`、`tools/call` |
| Prompts 提示词 | `IPrompt` | `RegisterPrompt New 你的类` | `prompts/list`、`prompts/get` |
| Resources 资源 | `IResource` | `RegisterResource New 你的类` | `resources/list`、`resources/read` |
| Resource Templates 模板 | `ITemplate` | `RegisterTemplate New 你的类` | `resources/templates/list`、`resources/read`（动态 URI） |

### IPrompt（提示词模板）

```vb
Option Explicit
Implements IPrompt

Private Property Get IPrompt_PromptName() As String
    IPrompt_PromptName = "code_review"              ' 提示词名（唯一，ASCII）
End Property

Private Property Get IPrompt_PromptDescription() As String
    IPrompt_PromptDescription = "生成代码审查意见"   ' 描述
End Property

Private Property Get IPrompt_PromptArguments() As String
    ' 参数声明（JSON 数组字符串）
    IPrompt_PromptArguments = "[{""name"":""language"",""required"":true}]"
End Property

Private Function IPrompt_GetPromptText(ByVal requestJson As String) As String
    Dim lang As String
    lang = CStr(JsonGet(requestJson, ".params.arguments.language"))
    IPrompt_GetPromptText = "请审查以下 " & lang & " 代码："
End Function
```

### IResource（资源）

```vb
Option Explicit
Implements IResource

Private Property Get IResource_ResourceUri() As String
    IResource_ResourceUri = "demo://server/info"    ' 唯一 URI（标准格式）
End Property

Private Property Get IResource_ResourceName() As String
    IResource_ResourceName = "服务器信息"
End Property

Private Property Get IResource_ResourceMimeType() As String
    IResource_ResourceMimeType = "text/plain"
End Property

Private Function IResource_ReadResourceText(ByVal uri As String) As String
    IResource_ReadResourceText = "动态生成的内容..."
End Function
```

参考实现：`tools\SamplePrompt.cls`（代码审查助手）、`tools\SampleResource.cls`（服务器信息）。

---

## 入口 mcp_main.bas

```vb
Public Sub Main()
    Dim server As McpServer
    Set server = New McpServer

    server.SetServerInfo "my-mcp-server", "1.0.0"   ' 可选：服务器标识

    ' 注册你的工具（想加多少加多少）
    server.RegisterTool New ToolGreeting
    server.RegisterTool New ToolAdd

    ' 注册提示词与资源（可选）
    server.RegisterPrompt New SamplePrompt
    server.RegisterResource New SampleResource

    ' 按命令行参数选择传输
    If LCase$(Command$) Like "*http*" Then
        server.RunHttp ParsePort(Command$)
    Else
        server.RunStdio
    End If
End Sub
```

---

## 与 MCP 客户端集成（stdio）

编译出 exe 后，把它注册到任意 MCP 客户端即可。

**Claude Desktop**（Windows，配置文件 `%APPDATA%\Claude\claude_desktop_config.json`）：

```json
{
  "mcpServers": {
    "vb6-mcp-sdk": {
      "command": "C:\\path\\to\\vb6-mcp-sdk.exe"
    }
  }
}
```

**Cursor**：`Settings → MCP → Add new MCP server`，Type 选 `command`，Command 填 exe 的完整路径（如 `C:\path\to\vb6-mcp-sdk.exe`）。

> 📌 stdio 模式默认启动，无需任何参数；exe 被客户端拉起时自动走 stdin/stdout 分帧协议。

---

## HTTP 传输模式（Streamable HTTP）

```powershell
# 默认端口 8080
.\vb6-mcp-sdk.exe /http

# 指定端口
.\vb6-mcp-sdk.exe /http:9000
```

- 客户端接入端点：`http://localhost:9000/mcp`
- **会话管理**：首次请求自动创建会话并返回 `Mcp-Session-Id` 响应头，后续请求携带该头复用；`DELETE /mcp` 关闭会话（无效会话返回 404）
- 支持 **CORS**，浏览器/Web 端 MCP 客户端可直接跨域调用
- 传输层为纯 Winsock API 自实现，无第三方 HTTP 组件
- 结合 `tests/client-sdk-http.py`（官方 Python SDK）可快速联调

---

## 验证

**一键运行**：`.\run_test.bat` —— 自动检查/安装 uv、exe 缺失时自动调用 VB6 编译、跑 fix-console、执行 6 组测试（冒烟 13 项断言 + SDK 冒烟 19 项断言 + 85 用例全面套件 + 33 项 HTTP 专项 + 官方 SDK 双传输），**并在最后自动输出两份覆盖率报告**：黑盒报告（`tests/coverage.py` 统计工具/提示词/资源/模板/协议方法）+ 白盒报告（`tests/coverage-whitebox.py` 对全部 VB6 源码做过程级调用图可达性分析，**只统计自研代码，VBJSON 第三方库（sdk/json/）排除**，宿主公开计数属性列为豁免项，当前 100%）；完整输出同时显示在控制台并记录到 `logs\run_test_*.log`。退出码：0=全部通过 / 1=有失败 / 2=环境未就绪。

也可以单独执行各测试：

```powershell
# 1) 冒烟测试（stdio，10 条消息 / 13 项断言：握手/ping/tools/prompts/resources/错误码/中文）
pwsh .\tests\test.ps1

# 1b) ★ 官方 Python SDK 冒烟（参考 SDK 测试断言风格，19 项；带 URL 参数则跑 HTTP 传输）
uv run --with mcp python .\tests\smoke-sdk.py
uv run --with mcp python .\tests\smoke-sdk.py http://localhost:9000/mcp

# 2) ★ 全面测试套件（stdio，85 用例：握手/工具/提示词/资源/模板/安全/裸协议/边界/数据完整性/MES/协议扩展/健壮性）
uv run --with mcp python .\tests\test-suite.py

# 3) ★ HTTP 传输层专项（33 项：CORS/202/404/OPTIONS/中文/缺参/会话/边界/三能力全链路/GET-SSE）
#    先启动：.\vb6-mcp-sdk.exe /http:9002
uv run python .\tests\http-test.py http://localhost:9002/mcp

# 4) 官方 Python SDK 完整握手（stdio）
uv run --with mcp python .\tests\client-sdk.py

# 5) 官方 Python SDK 走 HTTP
uv run --with mcp python .\tests\client-sdk-http.py http://localhost:9000/mcp

# 6) 官方 Inspector（在项目根目录执行）
npx @modelcontextprotocol/inspector .\vb6-mcp-sdk.exe

# 7) ★ Claude Code（Anthropic CLI）MCP 兼容性（需已安装并登录 claude CLI）
uv run python .\tests\claude-mcp-test.py

# 8) ★ DeepSeek Harness（DSH）MCP 兼容性（生成 cordis 配置块 + 按 dsh-mcp-client 方式握手）
uv run python .\tests\dsh-mcp-test.py
```

**测试套件覆盖**（`test-suite.py` 85 用例）：握手 5（含 capabilities 订阅/日志/补全声明）· 工具 37（含负数/小数/特殊字符/10KB 长文本/未知工具/缺参/isError/数据完整性/MES 4 例/structuredContent 2 例）· 提示词 5 · 资源 6 · 安全 3（路径穿越/绝对路径/非法扩展名）· 裸协议 29（非法 JSON/未知方法/通知无响应/id 类型/CRLF/补全/日志/订阅/错误后恢复/版本协商边界/cursor/initialize 容错）。**裸协议用例能抓到官方 SDK 客户端测不出的问题**（如字符串 id 不带引号、`\uXXXX` 转义解析）——已两次真实发现并修复框架 bug。MES 用例依赖内网服务器 192.168.20.151（接口见教程文档），不可达时对应用例会报失败。

---

## 架构

```
┌──────────────────────────────────────────────────┐
│ 你的代码（不改 SDK）                               │
│  mcp_main.bas → 注册能力 → RunStdio / RunHttp        │
│  工具类（Implements ITool）                       │
│  提示词类（Implements IPrompt）                   │
│  资源类（Implements IResource）                   │
├──────────────────────────────────────────────────┤
│ McpServer.cls（协议引擎，MCP 三大能力）            │
│  tools:    /list /call                            │
│  prompts:  /list /get                             │
│  resources:/list /read                            │
│  能力遍历分派（接口多态）                          │
├──────────────────┬───────────────────────────────┤
│ stdio 传输层      │ Streamable HTTP 传输层         │
│ mcp_transport_   │ mcp_transport_http.bas        │
│ stdio.bas        │ （Winsock API）                │
├──────────────────┴───────────────────────────────┤
│ mcp_json.bas（VBJSON 纯 VB6 解析）                  │
│ mcp_transport_stdio.bas（Utf8Encode/Decode）      │
│ mcp_log.bas                                       │
└──────────────────────────────────────────────────┘
```

---

## 环境要求与依赖

- VB6 IDE（32 位编译）
- JSON 解析采用纯 VB6 库（VBJSON，`sdk/json/`），**无需 MSScriptControl / msscript.ocx 注册，无运行期附加文件**

---

## 已知坑位备忘（踩过的雷）

| 症状 | 根因 | 处理 |
|---|---|---|
| 编译报「JsonInit 未定义」 | 模块未进工程 / .bas 是 LF 换行 | 双击 .vbp 加载；全部文件转 CRLF |
| 编译报「缺少标识符」 | API 声明参数名撞关键字（`len`） | 改名 `buflen`，勿用 len/name/type/error |
| 编译报「XX 未定义」（跨模块） | Private Declare 跨模块不可见 | 被其他模块调用时用 `Public Declare` |
| 首条消息 Parse error | 客户端（.NET StreamWriter 等）在消息前带 UTF-8 BOM | `JsonGet` 解析前自动剥离 U+FEFF |
| 运行时错误 450（对象→Variant） | VB6 学习版（VB6Expr）不支持对象赋值进 Variant | 遍历解析树只用 `Set` 对象指针，叶子直接取 `Item()` |
| 编译错「无效限定符」 | 函数内用 `step`/`cur` 等 VB6 敏感标识 | 改名（如 `j`/`curVal`），勿用关键字或内置名 |
| pwsh 管道无输出 | exe 是 GUI 子系统 | 跑 `fix-console.ps1` |
| HTTP 请求全部 404 | 请求头截取只取第一行 / 占位 0 字节污染 | 截 `sepPos+4` 完整头；用 `accLen` 跟踪长度 |

---

## 与 vb6mcp（单体版）的关系

`vb6mcp` 是这套框架的前身——把协议、传输、工具全写在一个 `mcp.bas` 里。`vb6-mcp-sdk` 把它拆成了可复用的分层框架：**协议引擎与具体工具解耦，新增工具只需实现 ITool 并注册**。

---

## 开源与贡献

仓库：[github.com/SMWHff/vb6-mcp-sdk](https://github.com/SMWHff/vb6-mcp-sdk) —— 觉得有用的话欢迎 ⭐ Star、提 [Issue](https://github.com/SMWHff/vb6-mcp-sdk/issues) 或发 [PR](https://github.com/SMWHff/vb6-mcp-sdk/pulls)。

- **许可证**：MIT（见 `LICENSE`）
- **版本**：1.0.0（协议 2024-11-05，兼容 2025-03-26 / 2025-06-18 / 2025-11-25 客户端）
- **技术栈**：VB6（32 位）+ Win32 API + VBJSON（纯 VB6 JSON 库）——无任何第三方 VB6 控件依赖
- **贡献方式**：
  - 新增示例工具：实现 `ITool` / `IPrompt` / `IResource` / `ITemplate` 接口，参考 `tools/` 下的模板
  - 修复 bug：跑通 `tests/test-suite.py`（41 用例）+ `tests/http-test.py`（28 项）后再提交
  - 提交规范：Conventional Commits（`feat(scope): 描述`，描述用中文）
- **测试**：官方 Python SDK（stdio + Streamable HTTP 双传输）全链路验证 + 裸协议错误用例（能抓到官方客户端测不出的问题）
- **相关资源**：[MCP 官方规范](https://modelcontextprotocol.io/) · [Python SDK](https://github.com/modelcontextprotocol/python-sdk) · [Inspector 调试器](https://github.com/modelcontextprotocol/inspector)

---

## Star History

<a href="https://www.star-history.com/?repos=SMWHff%2Fvb6-mcp-sdk&type=timeline&logscale=&legend=top-left">
 <picture>
   <source media="(prefers-color-scheme: dark)" srcset="https://api.star-history.com/chart?repos=SMWHff/vb6-mcp-sdk&type=timeline&theme=dark&logscale&legend=top-left&sealed_token=8lEBCwZdiL8GlGtz3Rj-Akv9rPBsQlHSKgQEbQsDlnPlK7Q6CSQocCQ4mIB2ke2rVfNaya-AsEkHsHNg2VkRmWqXkBLJzkmnAwUAnKLZ3yXwFoB1tvO0JuTz8CzIkHT5L8u-Tr9NesCK5Zhooyz3ZWh0FWeAstxSrkPULKdA7kySDbv0UI6MXvnEo2WI" />
   <source media="(prefers-color-scheme: light)" srcset="https://api.star-history.com/chart?repos=SMWHff/vb6-mcp-sdk&type=timeline&logscale&legend=top-left&sealed_token=8lEBCwZdiL8GlGtz3Rj-Akv9rPBsQlHSKgQEbQsDlnPlK7Q6CSQocCQ4mIB2ke2rVfNaya-AsEkHsHNg2VkRmWqXkBLJzkmnAwUAnKLZ3yXwFoB1tvO0JuTz8CzIkHT5L8u-Tr9NesCK5Zhooyz3ZWh0FWeAstxSrkPULKdA7kySDbv0UI6MXvnEo2WI" />
   <img alt="Star History Chart" src="https://api.star-history.com/chart?repos=SMWHff/vb6-mcp-sdk&type=timeline&logscale&legend=top-left&sealed_token=8lEBCwZdiL8GlGtz3Rj-Akv9rPBsQlHSKgQEbQsDlnPlK7Q6CSQocCQ4mIB2ke2rVfNaya-AsEkHsHNg2VkRmWqXkBLJzkmnAwUAnKLZ3yXwFoB1tvO0JuTz8CzIkHT5L8u-Tr9NesCK5Zhooyz3ZWh0FWeAstxSrkPULKdA7kySDbv0UI6MXvnEo2WI" />
 </picture>
</a>
