VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ITool —— MCP 工具接口（SDK 使用者实现此接口即可注册工具）
' 实现方式：类模块里写 Implements ITool，然后实现
'   ITool_ToolName / ITool_Description / ITool_InputSchema / ITool_Execute
' 注册方式：server.RegisterTool New 你的工具类
' ============================================================
Option Explicit

' 工具名（唯一，ASCII）
Public Property Get ToolName() As String
End Property

' 工具描述（会出现在 tools/list 里，可含中文）
Public Property Get Description() As String
End Property

' 输入参数 schema（JSON 字符串，如 {"type":"object","properties":{...},"required":[...]}）
Public Property Get InputSchema() As String
End Property

' 执行工具。
' requestJson = 完整的 tools/call 请求 JSON，
' 工具内用 JsonGet(requestJson, ".params.arguments.xxx") 取参数。
' 返回：结果文本（纯文本，框架自动包成 MCP content）。
' 参数不合法时抛 Err（框架会转成 JSON-RPC error -32602）。
Public Function Execute(ByVal requestJson As String) As String
End Function
