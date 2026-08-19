VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SamplePrompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' SamplePrompt —— 示例提示词：代码审查助手
' 演示：如何实现 IPrompt（带两个参数）
' ============================================================
Option Explicit
Implements IPrompt

Private Property Get IPrompt_PromptName() As String
    IPrompt_PromptName = "code_review"
End Property

Private Property Get IPrompt_PromptDescription() As String
    IPrompt_PromptDescription = "生成代码审查意见"
End Property

Private Property Get IPrompt_PromptArguments() As String
    IPrompt_PromptArguments = "[" _
        & "{""name"":""language"",""description"":""代码语言"",""required"":true}," _
        & "{""name"":""code"",""description"":""待审查的代码片段"",""required"":true}" _
        & "]"
End Property

Private Function IPrompt_GetPromptText(ByVal requestJson As String) As String
    Dim lang As String, code As String
    lang = CStr(JsonGet(requestJson, ".params.arguments.language"))
    code = CStr(JsonGet(requestJson, ".params.arguments.code"))
    IPrompt_GetPromptText = "请以资深工程师视角审查以下 " & lang & " 代码，" _
        & "指出潜在 bug、安全问题和可读性改进建议：" & vbCrLf & vbCrLf & code
End Function
