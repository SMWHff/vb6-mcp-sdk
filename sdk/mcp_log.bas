' ============================================================
' mcp_log.bas —— SDK 日志工具
' 所有日志写 logs\mcp.log（App.Path 下），绝不弹窗、绝不写 stdout
' ============================================================
Option Explicit

' 写一行日志（自动带时间戳）
Public Sub LogFile(ByVal msg As String)
    On Error Resume Next
    MkDir App.Path & "\logs"
    Dim f As Integer
    f = FreeFile
    Open App.Path & "\logs\mcp.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    Close #f
End Sub
