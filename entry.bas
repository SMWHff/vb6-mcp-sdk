' ============================================================
' entry.bas —— SDK 使用示例：组装入口
' 你的 MCP server 只需要改这里：
'   1) 注册你的工具（实现 ITool 接口的类）
'   2) （可选）SetServerInfo 设置服务器标识
' 启动：
'   vb6-mcp-sdk.exe          -> stdio 传输（默认）
'   vb6-mcp-sdk.exe /http    -> Streamable HTTP（端口 8080）
'   vb6-mcp-sdk.exe /http:9000 -> Streamable HTTP（端口 9000）
' ============================================================
Option Explicit

Private Const DEFAULT_PORT As Long = 8080

Public Sub Main()
    Dim server As McpServer
    Set server = New McpServer

    ' ===== 服务器标识（可选）=====
    server.SetServerInfo "vb6-mcp-sdk-demo", "1.0.0"

    ' ===== 注册你的工具（核心步骤）=====
    server.RegisterTool New ToolAdd
    server.RegisterTool New ToolEcho
    server.RegisterTool New ToolGetTime
    server.RegisterPrompt New SamplePrompt
    server.RegisterResource New SampleResource
    server.RegisterTool New ToolSysInfo
    server.RegisterTool New ToolReadFile
    server.RegisterTool New ToolWordCount

    ' ===== 启动（按命令行参数选传输）=====
    If LCase$(Command$) Like "*http*" Then
        server.RunHttp ParsePort(Command$)
    Else
        server.RunStdio
    End If
End Sub

' 从命令行解析端口：/http 或 /http:8080
Private Function ParsePort(ByVal cmd As String) As Long
    ParsePort = DEFAULT_PORT
    Dim p As Long
    p = InStr(1, LCase$(cmd), "http:")
    If p > 0 Then
        Dim rest As String
        rest = Mid$(cmd, p + 5)
        Dim num As String
        Dim i As Long
        For i = 1 To Len(rest)
            If Mid$(rest, i, 1) >= "0" And Mid$(rest, i, 1) <= "9" Then
                num = num & Mid$(rest, i, 1)
            ElseIf Len(num) > 0 Then
                Exit For
            End If
        Next i
        If Len(num) > 0 Then ParsePort = Val(num)
    End If
End Function
