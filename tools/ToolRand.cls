VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ToolRand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "rand"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "生成 [min, max] 区间内的随机整数"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{""min"":{""type"":""integer""},""max"":{""type"":""integer""}},""required"":[""min"",""max""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim minV As Long
    Dim maxV As Long
    minV = CLng(JsonGet(requestJson, ".params.arguments.min"))
    maxV = CLng(JsonGet(requestJson, ".params.arguments.max"))
    If maxV < minV Then Err.Raise vbObjectError + 1, , "max 必须大于等于 min"
    Randomize
    ITool_Execute = CStr(Int((maxV - minV + 1) * Rnd + minV))
End Function
