VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' IPrompt —— MCP 提示词模板接口（SDK 使用者实现此接口注册 prompt）
' 实现：类模块写 Implements IPrompt，实现
'   IPrompt_PromptName / IPrompt_PromptDescription /
'   IPrompt_PromptArguments / IPrompt_GetPromptText
' 注册：server.RegisterPrompt New 你的类
' ============================================================
Option Explicit

' 提示词名（唯一，ASCII）
Public Property Get PromptName() As String
End Property

' 提示词描述（显示在 prompts/list）
Public Property Get PromptDescription() As String
End Property

' 参数声明（JSON 数组字符串，如 [{"name":"lang","required":true}]）
Public Property Get PromptArguments() As String
End Property

' 生成提示词文本。
' requestJson = 完整的 prompts/get 请求 JSON，
' 参数用 JsonGet(requestJson, ".params.arguments.xxx") 取。
' 返回：提示词文本（框架包成 messages[0].content.text）。
Public Function GetPromptText(ByVal requestJson As String) As String
End Function
