VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITemplate

Private Property Get ITemplate_UriTemplate() As String
    ITemplate_UriTemplate = "demo://greet/{name}"
End Property

Private Property Get ITemplate_TemplateName() As String
    ITemplate_TemplateName = "greet"
End Property

Private Property Get ITemplate_TemplateDescription() As String
    ITemplate_TemplateDescription = "生成问候语（模板示例，动态解析 {name}）"
End Property

Private Property Get ITemplate_TemplateMimeType() As String
    ITemplate_TemplateMimeType = "text/plain"
End Property

Private Function ITemplate_ResolveText(ByVal uri As String) As String
    Dim prefix As String
    prefix = "demo://greet/"
    Dim name As String
    If Left$(uri, Len(prefix)) = prefix Then
        name = Mid$(uri, Len(prefix) + 1)
    Else
        name = uri
    End If
    ITemplate_ResolveText = "你好，" & name & "！"
End Function
