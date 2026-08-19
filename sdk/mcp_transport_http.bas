' ============================================================
' mcp_transport_http.bas —— Streamable HTTP 传输层（Winsock API）
' 职责：监听 TCP、accept、读取 HTTP 请求（body 按 UTF-8 解码）、
'       发送 JSON 响应（含 CORS 头）、关闭连接。纯 IO，不含业务逻辑。
' 依赖：mcp_transport_stdio.bas 的 Utf8Encode / Utf8Decode
' 模型：阻塞、一次一连接、Connection: close
' 踩坑备忘：Declare 参数名勿用关键字（len→buflen）；跨模块调用的
'        Declare 必须 Public；字节缓冲用独立长度变量防 0 字节污染
' ============================================================
Option Explicit

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As Any) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal sType As Long, ByVal protocol As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As Any, ByVal namelen As Long) As Long
Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As Any, ByRef addrlen As Long) As Long
Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal flags As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer

Private Const AF_INET As Long = 2
Private Const SOCK_STREAM As Long = 1
Private Const IPPROTO_TCP As Long = 6
Private Const SOMAXCONN As Long = 5
Private Const INADDR_ANY As Long = 0
Private Const RECV_CHUNK As Long = 4096

Private Type sockaddr_in
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero As String * 8
End Type

Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

' 初始化并监听端口，返回监听 socket（失败返回 0）
Public Function HttpStart(ByVal port As Long) As Long
    Dim wsa As WSADATA
    If WSAStartup(&H202, wsa) <> 0 Then Exit Function
    Dim s As Long
    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s <= 0 Then Exit Function
    Dim addr As sockaddr_in
    addr.sin_family = AF_INET
    addr.sin_port = htons(port)
    addr.sin_addr = INADDR_ANY
    If bind(s, addr, LenB(addr)) <> 0 Then
        Call closesocket(s): Exit Function
    End If
    If listen(s, SOMAXCONN) <> 0 Then
        Call closesocket(s): Exit Function
    End If
    HttpStart = s
End Function

' 阻塞等待连接，返回连接 socket（失败返回 0）
Public Function HttpAccept(ByVal listenSock As Long) As Long
    Dim addr As sockaddr_in
    Dim addrLen As Long
    addrLen = LenB(addr)
    HttpAccept = accept(listenSock, addr, addrLen)
End Function

' 读取一个完整 HTTP 请求；返回 0=成功，其他=失败
Public Function HttpReadRequest(ByVal conn As Long, ByRef method As String, ByRef path As String, ByRef body As String) As Long
    method = "": path = "": body = ""
    Dim acc() As Byte
    ReDim acc(0 To 0)          ' 占位数组；真实长度由 accLen 跟踪
    Dim accLen As Long
    accLen = 0
    Dim b(0 To RECV_CHUNK - 1) As Byte
    Dim n As Long, sepPos As Long

    ' 收数据直到出现 \r\n\r\n（头结束）
    Do
        n = recv(conn, b(0), RECV_CHUNK, 0)
        If n <= 0 Then Exit Function
        Call ByteAppend(acc, b, n, accLen)
        sepPos = ByteFindCrlfCrlf(acc, accLen)
        If sepPos >= 0 Then Exit Do
        If accLen > 65536 Then Exit Function
    Loop

    ' 完整头 = acc(0..sepPos+3)，含 \r\n\r\n
    Dim head As String
    head = BytesAscii(acc, sepPos + 4)
    Dim line As String
    line = Left$(head, InStr(1, head, vbCrLf) - 1)
    Dim sp1 As Long, sp2 As Long
    sp1 = InStr(1, line, " ")
    sp2 = InStr(sp1 + 1, line, " ")
    If sp1 > 0 And sp2 > 0 Then
        method = Left$(line, sp1 - 1)
        path = Mid$(line, sp1 + 1, sp2 - sp1 - 1)
    Else
        Exit Function
    End If

    ' body：按 Content-Length 从 acc 提取，不足则继续 recv
    Dim cl As Long
    cl = GetHeaderLong(head, "Content-Length")
    If cl > 0 Then
        Dim bodyStart As Long
        bodyStart = sepPos + 4
        Do While accLen < bodyStart + cl
            n = recv(conn, b(0), RECV_CHUNK, 0)
            If n <= 0 Then Exit Function
            Call ByteAppend(acc, b, n, accLen)
        Loop
        Dim bodyBytes() As Byte
        ReDim bodyBytes(0 To cl - 1)
        Dim i As Long
        For i = 0 To cl - 1
            bodyBytes(i) = acc(bodyStart + i)
        Next i
        body = Utf8Decode(bodyBytes, cl)   ' body 是 UTF-8 JSON
    End If
    HttpReadRequest = 0
End Function

' 发送 200 JSON 响应
Public Function HttpSendJson(ByVal conn As Long, ByVal jsonBody As String) As Long
    HttpSendJson = HttpSendRaw(conn, 200, "OK", "application/json", jsonBody)
End Function

' 发送 202（通知类请求）
Public Function HttpSendAccepted(ByVal conn As Long) As Long
    HttpSendAccepted = HttpSendRaw(conn, 202, "Accepted", "application/json", "")
End Function

' 发送 404
Public Function HttpSendNotFound(ByVal conn As Long) As Long
    HttpSendNotFound = HttpSendRaw(conn, 404, "Not Found", "application/json", "{""error"":""not found""}")
End Function

' 发送 204（OPTIONS 预检）
Public Function HttpSendPreflight(ByVal conn As Long) As Long
    HttpSendPreflight = HttpSendRaw(conn, 204, "No Content", "", "")
End Function

' 核心发送：组头 + UTF-8 body，一次发送
Public Function HttpSendRaw(ByVal conn As Long, ByVal status As Long, ByVal statusText As String, _
    ByVal ctype As String, ByVal body As String) As Long
    On Error GoTo FailSend
    Dim b() As Byte
    b = Utf8Encode(body)
    Dim blen As Long
    blen = UBound(b) + 1
    If Len(ctype) = 0 Then ctype = "text/plain"

    Dim head As String
    head = "HTTP/1.1 " & status & " " & statusText & vbCrLf _
        & "Content-Type: " & ctype & "; charset=utf-8" & vbCrLf _
        & "Content-Length: " & blen & vbCrLf _
        & "Connection: close" & vbCrLf _
        & "Access-Control-Allow-Origin: *" & vbCrLf _
        & "Access-Control-Allow-Headers: Content-Type, MCP-Protocol-Version, Mcp-Session-Id" & vbCrLf _
        & "Access-Control-Allow-Methods: POST, GET, OPTIONS" & vbCrLf _
        & vbCrLf
    Dim hb() As Byte
    hb = Utf8Encode(head)
    Call send(conn, hb(0), UBound(hb) + 1, 0)
    If blen > 0 Then Call send(conn, b(0), blen, 0)
    HttpSendRaw = 0
    Exit Function
FailSend:
    HttpSendRaw = -1
End Function

' 关闭连接
Public Sub HttpClose(ByVal conn As Long)
    On Error Resume Next
    If conn > 0 Then Call closesocket(conn)
End Sub

' 程序退出前清理
Public Sub HttpCleanup()
    On Error Resume Next
    Call WSACleanup
End Sub

' ---- 内部辅助 ----
' 在 acc 前 accLen 字节中找 \r\n\r\n，返回 \r 的 0 基索引；找不到返回 -1
Private Function ByteFindCrlfCrlf(ByRef acc() As Byte, ByVal accLen As Long) As Long
    Dim i As Long
    For i = 0 To accLen - 4
        If acc(i) = 13 And acc(i + 1) = 10 And acc(i + 2) = 13 And acc(i + 3) = 10 Then
            ByteFindCrlfCrlf = i
            Exit Function
        End If
    Next i
    ByteFindCrlfCrlf = -1
End Function

' 把 b 的前 n 字节追加到 acc（真实长度由 accLen 跟踪，避免占位 0 字节污染）
Private Sub ByteAppend(ByRef acc() As Byte, ByRef b() As Byte, ByVal n As Long, ByRef accLen As Long)
    ReDim Preserve acc(0 To accLen + n - 1)
    Dim i As Long
    For i = 0 To n - 1
        acc(accLen + i) = b(i)
    Next i
    accLen = accLen + n
End Sub

' 前 count 字节转 ASCII 字符串（HTTP 头）
Private Function BytesAscii(ByRef acc() As Byte, ByVal count As Long) As String
    If count <= 0 Then Exit Function
    Dim s As String
    s = Space$(count)
    Dim i As Long
    For i = 0 To count - 1
        Mid$(s, i + 1, 1) = Chr$(acc(i))
    Next i
    BytesAscii = s
End Function

' 从头里取整数头字段（如 Content-Length）
Private Function GetHeaderLong(ByVal head As String, ByVal name As String) As Long
    Dim pos As Long
    pos = InStr(1, LCase$(head), LCase$(name) & ":")
    If pos = 0 Then Exit Function
    Dim line As String
    line = Mid$(head, pos + Len(name) + 1)
    line = Left$(line, InStr(1, line, vbCrLf) - 1)
    GetHeaderLong = Val(Trim$(line))
End Function
