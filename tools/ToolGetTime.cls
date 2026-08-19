VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolGetTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolGetTime —— 示例工具：返回服务器当前时间
' 演示：无参数工具
' ============================================================
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "get_time"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "返回服务器当前时间"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{}}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    ITool_Execute = Format$(Now, "yyyy-mm-dd hh:nn:ss")
End Function
