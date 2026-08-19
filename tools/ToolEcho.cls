VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolEcho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolEcho —— 示例工具：原样返回传入文本
' 演示：字符串参数 + 中文传输
' ============================================================
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "echo"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "原样返回传入文本"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""text"":{""type"":""string""}}," _
        & """required"":[""text""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim txt As Variant
    txt = JsonGet(requestJson, ".params.arguments.text")
    ITool_Execute = CStr(txt)
End Function
