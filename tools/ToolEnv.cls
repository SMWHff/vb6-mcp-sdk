VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "getenv"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "查询环境变量（可选默认值）"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{""name"":{""type"":""string""},""default"":{""type"":""string""}},""required"":[""name""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim name As String
    name = CStr(JsonGet(requestJson, ".params.arguments.name"))
    Dim val As String
    val = Environ$(name)
    If Len(val) = 0 Then
        Dim dflt As Variant
        dflt = JsonGet(requestJson, ".params.arguments.default")
        If IsEmpty(dflt) Or IsNull(dflt) Then
            ITool_Execute = name & " 未定义"
        Else
            ITool_Execute = CStr(dflt)
        End If
    Else
        ITool_Execute = val
    End If
End Function
