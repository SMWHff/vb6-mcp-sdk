' ============================================================
' mcp_transport_stdio.bas —— stdio 传输层（kernel32，零依赖）
' 职责：按行分帧（LF，兼容 CRLF）+ UTF-8/UTF-16 转换
' 注意：stdout 是协议通道，只允许写 JSON-RPC 消息，禁止用于日志
' 关键：读行按【字节】累积、整行一次性 UTF-8 解码——逐块解码会把
'       多字节字符在块边界截断（UTF-8 汉字 3 字节，8192 边界会切断）
' ============================================================
Option Explicit

Private Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As Long, _
    ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
    ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
    ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const STD_INPUT_HANDLE As Long = -10
Private Const STD_OUTPUT_HANDLE As Long = -11
Private Const CP_UTF8 As Long = 65001
Private Const CHUNK As Long = 8192

Private m_hIn As Long
Private m_hOut As Long
Private m_acc() As Byte            ' 输入字节缓冲
Private m_accLen As Long           ' 缓冲真实长度（避免占位 0 字节污染）

Public Sub StdioInit()
    m_hIn = GetStdHandle(STD_INPUT_HANDLE)
    m_hOut = GetStdHandle(STD_OUTPUT_HANDLE)
    ReDim m_acc(0 To 0)
    m_accLen = 0
End Sub

' 读一行：以 LF 分帧，去掉行尾 CR；EOF 返回 ""（缓冲有残余则作为最后一行）
Public Function StdioReadLine() As String
    Dim bytes(0 To CHUNK - 1) As Byte
    Dim got As Long, i As Long, lfPos As Long

    Do
        ' 在缓冲中找 LF
        lfPos = -1
        For i = 0 To m_accLen - 1
            If m_acc(i) = 10 Then lfPos = i: Exit For
        Next i
        If lfPos >= 0 Then
            StdioReadLine = TakeLine(m_acc, lfPos)
            Call ConsumeLine(m_acc, m_accLen, lfPos)
            Exit Function
        End If
        got = 0
        If ReadFile(m_hIn, VarPtr(bytes(0)), CHUNK, got, 0) = 0 Then Exit Do
        If got = 0 Then Exit Do
        Call ByteAppend(m_acc, m_accLen, bytes, got)
    Loop

    ' EOF：剩余字节作为最后一行
    If m_accLen > 0 Then
        Dim last As String
        last = TakeLine(m_acc, m_accLen)
        m_accLen = 0
        StdioReadLine = last
    Else
        StdioReadLine = ""
    End If
End Function

' 写一行：UTF-8 + LF。这是唯一的 stdout 出口
Public Sub StdioWriteLine(ByVal s As String)
    Dim b() As Byte
    b = Utf8Encode(s & vbLf)
    Dim n As Long
    If UBound(b) >= 0 Then Call WriteFile(m_hOut, VarPtr(b(0)), UBound(b) + 1, n, 0)
End Sub

' ---- 内部辅助 ----

' 提取一行（0..lfPos 前的内容，去 \n 和行尾 \r），整体 UTF-8 解码
Private Function TakeLine(ByRef acc() As Byte, ByVal lfPos As Long) As String
    Dim lineLen As Long
    lineLen = lfPos
    If lineLen > 0 And acc(lineLen - 1) = 13 Then lineLen = lineLen - 1
    If lineLen <= 0 Then Exit Function
    Dim lineBytes() As Byte
    ReDim lineBytes(0 To lineLen - 1)
    Dim i As Long
    For i = 0 To lineLen - 1
        lineBytes(i) = acc(i)
    Next i
    TakeLine = Utf8Decode(lineBytes, lineLen)
End Function

' 消费已读行（含 \n），剩余字节前移
Private Sub ConsumeLine(ByRef acc() As Byte, ByRef accLen As Long, ByVal lfPos As Long)
    Dim remain As Long
    remain = accLen - lfPos - 1
    Dim i As Long
    For i = 0 To remain - 1
        acc(i) = acc(lfPos + 1 + i)
    Next i
    accLen = remain
End Sub

' 追加字节到缓冲
Private Sub ByteAppend(ByRef acc() As Byte, ByRef accLen As Long, ByRef bytes() As Byte, ByVal n As Long)
    Dim oldLen As Long
    oldLen = accLen
    ReDim Preserve acc(0 To oldLen + n - 1)
    Dim i As Long
    For i = 0 To n - 1
        acc(oldLen + i) = bytes(i)
    Next i
    accLen = oldLen + n
End Sub

' ---- UTF-8 <-> UTF-16（HTTP 传输层也复用）----
Public Function Utf8Encode(ByVal s As String) As Byte()
    Dim n As Long
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If n <= 0 Then n = 1
    Dim b() As Byte
    ReDim b(0 To n - 1)
    If n > 1 Then Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0)
    Utf8Encode = b
End Function

Public Function Utf8Decode(ByRef bytes() As Byte, ByVal count As Long) As String
    If count <= 0 Then Exit Function
    Dim n As Long
    n = MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(0)), count, 0, 0)
    If n <= 0 Then Exit Function
    Dim s As String
    s = String$(n, vbNullChar)
    Call MultiByteToWideChar(CP_UTF8, 0, VarPtr(bytes(0)), count, StrPtr(s), n)
    Utf8Decode = s
End Function
