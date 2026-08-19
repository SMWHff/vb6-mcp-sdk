VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolAdd —— 示例工具：两个数相加
' 演示：如何实现 ITool 接口（Implements + ITool_xxx 命名）
' ============================================================
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "add"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "两个数相加"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""a"":{""type"":""number""},""b"":{""type"":""number""}}," _
        & """required"":[""a"",""b""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim a As Double, b As Double
    a = CDbl(JsonGet(requestJson, ".params.arguments.a"))
    b = CDbl(JsonGet(requestJson, ".params.arguments.b"))
    ITool_Execute = CStr(a + b)
End Function
