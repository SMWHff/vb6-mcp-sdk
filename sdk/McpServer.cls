VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "McpServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' McpServer —— MCP 服务器类（SDK 核心，覆盖 MCP 三大能力）
'   Tools    ：server.RegisterTool    New 工具类     (Implements ITool)
'   Prompts  ：server.RegisterPrompt  New 提示词类   (Implements IPrompt)
'   Resources：server.RegisterResource New 资源类    (Implements IResource)
' 启动：server.RunStdio 或 server.RunHttp 8080
' ============================================================
Option Explicit

Private Const PROTOCOL_VERSION As String = "2024-11-05"

Private Declare Function GetConsoleWindow Lib "kernel32" () As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_HIDE As Long = 0

Private m_tools As Collection
Private m_prompts As Collection
Private m_resources As Collection
Private m_serverName As String
Private m_serverVersion As String

Private Sub Class_Initialize()
    Set m_tools = New Collection
    Set m_prompts = New Collection
    Set m_resources = New Collection
    m_serverName = "vb6-mcp-sdk"
    m_serverVersion = "1.0.0"
End Sub

' 设置服务器标识（可选）
Public Sub SetServerInfo(ByVal name As String, ByVal version As String)
    m_serverName = name
    m_serverVersion = version
End Sub

' ============ 注册（三大能力） ============

Public Sub RegisterTool(ByVal tool As ITool)
    m_tools.Add tool
End Sub

Public Sub RegisterPrompt(ByVal prompt As IPrompt)
    m_prompts.Add prompt
End Sub

Public Sub RegisterResource(ByVal resource As IResource)
    m_resources.Add resource
End Sub

Public Property Get ToolCount() As Long
    ToolCount = m_tools.Count
End Property

Public Property Get PromptCount() As Long
    PromptCount = m_prompts.Count
End Property

Public Property Get ResourceCount() As Long
    ResourceCount = m_resources.Count
End Property

' ===================== 传输模式 =====================

Public Sub RunStdio()
    On Error GoTo Fail
    HideConsole
    StdioInit
    JsonInit
    LogFile "=== MCP stdio server 启动 ==="
    Dim line As String
    line = StdioReadLine()
    Do While Len(line) > 0
        Dim resp As String
        resp = HandleMessage(line)
        If Len(resp) > 0 Then StdioWriteLine resp
        line = StdioReadLine()
    Loop
    LogFile "=== MCP stdio server 退出 ==="
    Exit Sub
Fail:
    LogFile "致命错误: " & Err.Number & " " & Err.Description
End Sub

Public Sub RunHttp(ByVal port As Long)
    On Error GoTo Fail
    JsonInit
    LogFile "=== HTTP MCP server 启动 (端口 " & port & ") ==="
    Dim listenSock As Long
    listenSock = HttpStart(port)
    If listenSock <= 0 Then
        LogFile "监听失败，错误码: " & WSAGetLastError() & "（端口可能被占用）"
        Exit Sub
    End If
    LogFile "监听成功，等待连接..."
    Dim conn As Long
    Do
        conn = HttpAccept(listenSock)
        If conn <= 0 Then Exit Do
        Call HandleHttpConnection(conn)
        Call HttpClose(conn)
    Loop
    Call HttpClose(listenSock)
    Call HttpCleanup
    LogFile "=== HTTP MCP server 退出 ==="
    Exit Sub
Fail:
    LogFile "HTTP 致命错误: " & Err.Number & " " & Err.Description
End Sub

' ===================== 协议处理 =====================

Public Function HandleMessage(ByVal json As String) As String
    On Error GoTo BadJson
    Dim method As String
    method = CStr(JsonGet(json, ".method"))

    If Left$(method, 14) = "notifications/" Then
        HandleMessage = ""
        Exit Function
    End If

    Dim id As Variant
    id = JsonGet(json, ".id")
    Dim idText As String
    If IsNull(id) Or IsEmpty(id) Then
        idText = "null"
    ElseIf VarType(id) = vbString Then
        idText = JsonQuote(CStr(id))
    Else
        idText = CStr(id)
    End If

    Select Case method
    Case "initialize"
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{" _
            & """protocolVersion"":""" & PROTOCOL_VERSION & """," _
            & """capabilities"":{""tools"":{""listChanged"":false},""prompts"":{},""resources"":{}}," _
            & """serverInfo"":{""name"":""" & m_serverName & """,""version"":""" & m_serverVersion & """}}}"

    Case "ping"
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{}}"

    Case "tools/list"
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""tools"":[" & BuildToolsList() & "]}}"

    Case "tools/call"
        HandleMessage = CallTool(json, idText)

    Case "prompts/list"
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""prompts"":[" & BuildPromptsList() & "]}}"

    Case "prompts/get"
        HandleMessage = CallPrompt(json, idText)

    Case "resources/list"
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""resources"":[" & BuildResourcesList() & "]}}"

    Case "resources/read"
        HandleMessage = ReadResourceValue(json, idText)

    Case Else
        HandleMessage = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32601," _
            & """message"":""方法未实现""}}"
    End Select
    Exit Function

BadJson:
    LogFile "处理失败: " & Err.Number & " " & Err.Description & " | 原文: " & json
    HandleMessage = "{""jsonrpc"":""2.0"",""id"":null,""error"":{""code"":-32700,""message"":""Parse error""}}"
End Function

' ===================== Tools 内部 =====================

Private Function BuildToolsList() As String
    Dim list As String
    list = ""
    Dim t As ITool
    For Each t In m_tools
        If Len(list) > 0 Then list = list & ","
        list = list & "{""name"":""" & t.ToolName & """," _
            & """description"":" & JsonQuote(t.Description) & "," _
            & """inputSchema"":" & t.InputSchema & "}"
    Next t
    BuildToolsList = list
End Function

Private Function CallTool(ByVal json As String, ByVal idText As String) As String
    Dim toolName As String
    toolName = CStr(JsonGet(json, ".params.name"))
    Dim t As ITool
    For Each t In m_tools
        If t.ToolName = toolName Then
            On Error GoTo ToolErr
            Dim missing As String
            missing = JsonCheckRequired(t.InputSchema, json)
            If Len(missing) > 0 Then
                CallTool = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
                    & """message"":" & JsonQuote("缺少必需参数: " & missing) & "}}"
                Exit Function
            End If
            Dim resultText As String
            resultText = t.Execute(json)
            CallTool = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""content"":[" _
                & "{""type"":""text"",""text"":" & JsonQuote(resultText) & "}]}}"
            Exit Function
        End If
    Next t
    CallTool = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
        & """message"":" & JsonQuote("未知工具: " & toolName) & "}}"
    Exit Function
ToolErr:
    ' 工具执行失败：按规范返回 isError:true 的结果（客户端能看到错误文本）
    CallTool = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""content"":[" _
        & "{""type"":""text"",""text"":" & JsonQuote("执行失败: " & Err.Description) & "}],""isError"":true}}"
End Function

' ===================== Prompts 内部 =====================

Private Function BuildPromptsList() As String
    Dim list As String
    list = ""
    Dim p As IPrompt
    For Each p In m_prompts
        If Len(list) > 0 Then list = list & ","
        list = list & "{""name"":""" & p.PromptName & """," _
            & """description"":" & JsonQuote(p.PromptDescription) & "," _
            & """arguments"":" & p.PromptArguments & "}"
    Next p
    BuildPromptsList = list
End Function

Private Function CallPrompt(ByVal json As String, ByVal idText As String) As String
    Dim pName As String
    pName = CStr(JsonGet(json, ".params.name"))
    Dim p As IPrompt
    For Each p In m_prompts
        If p.PromptName = pName Then
            On Error GoTo PromptErr
            Dim text As String
            text = p.GetPromptText(json)
            CallPrompt = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""description"":" _
                & JsonQuote(p.PromptDescription) & ",""messages"":[{""role"":""user""," _
                & """content"":{""type"":""text"",""text"":" & JsonQuote(text) & "}}]}}"
            Exit Function
        End If
    Next p
    CallPrompt = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
        & """message"":" & JsonQuote("未知提示词: " & pName) & "}}"
    Exit Function
PromptErr:
    CallPrompt = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
        & """message"":" & JsonQuote("提示词生成失败: " & Err.Description) & "}}"
End Function

' ===================== Resources 内部 =====================

Private Function BuildResourcesList() As String
    Dim list As String
    list = ""
    Dim r As IResource
    For Each r In m_resources
        If Len(list) > 0 Then list = list & ","
        list = list & "{""uri"":""" & r.ResourceUri & """," _
            & """name"":" & JsonQuote(r.ResourceName) & "," _
            & """mimeType"":""" & r.ResourceMimeType & """," _
            & """description"":" & JsonQuote(r.ResourceDescription) & "}"
    Next r
    BuildResourcesList = list
End Function

Private Function ReadResourceValue(ByVal json As String, ByVal idText As String) As String
    Dim uri As String
    uri = CStr(JsonGet(json, ".params.uri"))
    Dim r As IResource
    For Each r In m_resources
        If r.ResourceUri = uri Then
            On Error GoTo ResErr
            Dim text As String
            text = r.ReadResourceText(uri)
            ReadResourceValue = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""result"":{""contents"":[" _
                & "{""uri"":""" & uri & """,""mimeType"":""" & r.ResourceMimeType & """," _
                & """text"":" & JsonQuote(text) & "}]}}"
            Exit Function
        End If
    Next r
    ReadResourceValue = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
        & """message"":" & JsonQuote("未知资源: " & uri) & "}}"
    Exit Function
ResErr:
    ReadResourceValue = "{""jsonrpc"":""2.0"",""id"":" & idText & ",""error"":{""code"":-32602," _
        & """message"":" & JsonQuote("资源读取失败: " & Err.Description) & "}}"
End Function

' ===================== HTTP 连接 =====================

Private Sub HandleHttpConnection(ByVal conn As Long)
    On Error GoTo ConnFail
    Dim method As String, path As String, reqBody As String
    Dim rc As Long
    rc = HttpReadRequest(conn, method, path, reqBody)
    If rc <> 0 Then Exit Sub

    If method = "OPTIONS" Then
        Call HttpSendPreflight(conn)
    ElseIf method = "POST" And path = "/mcp" Then
        Dim resp As String
        resp = HandleMessage(reqBody)
        If Len(resp) > 0 Then
            Call HttpSendJson(conn, resp)
        Else
            Call HttpSendAccepted(conn)
        End If
    ElseIf method = "GET" Then
        Call HttpSendJson(conn, "{""name"":""" & m_serverName & """,""mode"":""streamable-http"",""endpoint"":""/mcp""}")
    Else
        Call HttpSendNotFound(conn)
    End If
    Exit Sub
ConnFail:
    LogFile "连接处理失败: " & Err.Number & " " & Err.Description
End Sub

Private Sub HideConsole()
    On Error Resume Next
    Dim h As Long
    h = GetConsoleWindow()
    If h <> 0 Then Call ShowWindow(h, SW_HIDE)
End Sub
