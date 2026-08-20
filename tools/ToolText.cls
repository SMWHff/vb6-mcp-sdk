VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "text_case"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "字符串处理：按 mode 转换大小写/去首尾空白/反转"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{""text"":{""type"":""string""},""mode"":{""type"":""string"",""enum"":[""upper"",""lower"",""trim"",""reverse""]}},""required"":[""text"",""mode""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim text As String
    Dim mode As String
    text = CStr(JsonGet(requestJson, ".params.arguments.text"))
    mode = CStr(JsonGet(requestJson, ".params.arguments.mode"))
    Select Case mode
        Case "upper": ITool_Execute = UCase$(text)
        Case "lower": ITool_Execute = LCase$(text)
        Case "trim": ITool_Execute = Trim$(text)
        Case "reverse": ITool_Execute = StrReverse(text)
        Case Else: Err.Raise vbObjectError + 1, , "mode 必须为 upper/lower/trim/reverse"
    End Select
End Function
