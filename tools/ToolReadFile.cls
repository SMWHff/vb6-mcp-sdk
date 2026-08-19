VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolReadFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolReadFile —— 示例工具：安全读取项目内文本文件
' 安全约束（白名单）：
'   - 只允许相对路径（拒绝绝对路径）
'   - 拒绝路径穿越（..\ 或 ../）
'   - 扩展名白名单：txt/md/bas/cls/json/ini/log
'   - 仅限 App.Path（exe 所在目录）下的文件
' 演示：MCP 工具的安全边界实践
' ============================================================
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "read_file"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "读取项目内的文本文件（白名单限制，防路径穿越）"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""path"":{""type"":""string"",""description"":""相对路径，如 README.md""}}," _
        & """required"":[""path""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim relPath As String
    relPath = CStr(JsonGet(requestJson, ".params.arguments.path"))

    If Not IsSafePath(relPath) Then
        Err.Raise vbObjectError + 1, , "非法路径: " & relPath & "（只允许项目内相对路径）"
    End If

    Dim fullPath As String
    fullPath = App.Path & "\" & relPath

    Dim f As Integer, n As Long
    f = FreeFile
    Open fullPath For Binary As #f
    n = LOF(f)
    Dim buf() As Byte
    If n > 0 Then
        ReDim buf(0 To n - 1)
        Get #f, , buf
    End If
    Close #f
    ITool_Execute = Utf8Decode(buf, n)
End Function

' 路径安全检查
Private Function IsSafePath(ByVal p As String) As Boolean
    On Error GoTo NotSafe
    p = Trim$(p)
    If Len(p) = 0 Then Exit Function
    ' 拒绝绝对路径（盘符开头 或 反斜杠/斜杠开头）
    If Len(p) >= 2 Then
        If Mid$(p, 2, 1) = ":" Then Exit Function
    End If
    If Left$(p, 1) = "\" Or Left$(p, 1) = "/" Then Exit Function
    ' 拒绝路径穿越
    If InStr(1, p, "..") > 0 Then Exit Function
    ' 扩展名白名单
    Dim ext As String
    ext = LCase$(Mid$(p, InStrRev(p, ".") + 1))
    Select Case ext
    Case "txt", "md", "bas", "cls", "json", "ini", "log"
        IsSafePath = True
    End Select
    Exit Function
NotSafe:
    IsSafePath = False
End Function
