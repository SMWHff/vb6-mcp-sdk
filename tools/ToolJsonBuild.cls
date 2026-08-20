VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolJsonBuild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "json_build"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "JsonBuild 单元测试：返回构造的 JSON 对象（演示字符串/数字/布尔/null/嵌套）"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{""mode"":{""type"":""string"",""enum"":[""full"",""empty""]}},""required"":[]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim mode As String
    mode = CStr(JsonGet(requestJson, ".params.arguments.mode"))
    If Len(mode) = 0 Or mode = "full" Then
        ITool_Execute = JsonBuild( _
            "name", "vb6", _
            "count", 3, _
            "ratio", 3.5, _
            "ok", True, _
            "no", False, _
            "nil", Null, _
            "nested", JsonBuild( _
                "a", 1 _
            ) _
        )
    ElseIf mode = "empty" Then
        ITool_Execute = JsonBuild()
    Else
        ITool_Execute = JsonBuild("error", mode)
    End If
End Function
