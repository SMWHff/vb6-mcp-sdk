VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ITemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' ITemplate —— Resource Template 接口（MCP resources/templates）
'   实现类提供动态资源模板：uriTemplate 形如 "demo://greet/{name}"
'   （RFC 6570 简化版，仅支持 {name} 形式变量）
'   框架匹配模板后调用 ResolveText(具体URI) 动态生成内容
'   注册：server.RegisterTemplate New 你的模板类
' ============================================================

' URI 模板（含 {var} 占位符），如 "demo://greet/{name}"
Public Property Get UriTemplate() As String
End Property

' 模板名称（客户端展示）
Public Property Get TemplateName() As String
End Property

' 模板描述（可中文）
Public Property Get TemplateDescription() As String
End Property

' 内容 MIME 类型，如 "text/plain"
Public Property Get TemplateMimeType() As String
End Property

' 按具体 URI 解析模板变量并生成文本内容
Public Function ResolveText(ByVal uri As String) As String
End Function
