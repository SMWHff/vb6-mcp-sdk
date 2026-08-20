VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IPrompt

Private Property Get IPrompt_PromptName() As String
    IPrompt_PromptName = "translate"
End Property

Private Property Get IPrompt_PromptDescription() As String
    IPrompt_PromptDescription = "翻译助手：按目标语言翻译文本"
End Property

Private Property Get IPrompt_PromptArguments() As String
    IPrompt_PromptArguments = "[{""name"":""language"",""required"":true},{""name"":""text"",""required"":true}]"
End Property

Private Function IPrompt_GetPromptText(ByVal requestJson As String) As String
    Dim lang As String
    Dim text As String
    lang = CStr(JsonGet(requestJson, ".params.arguments.language"))
    text = CStr(JsonGet(requestJson, ".params.arguments.text"))
    IPrompt_GetPromptText = "请将以下文本翻译为" & lang & "，仅返回翻译结果：" & vbCrLf & text
End Function
