VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolWordCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolWordCount —— 示例工具：文本统计
' 演示：字符串处理 + isError 语义（传空文本时抛错，框架返回 isError:true）
' ============================================================
Option Explicit
Implements ITool

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "word_count"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "统计文本的字数、字符数和行数"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{""text"":{""type"":""string"",""description"":""要统计的文本""}}," _
        & """required"":[""text""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim txt As String
    txt = CStr(JsonGet(requestJson, ".params.arguments.text"))
    If Len(txt) = 0 Then
        Err.Raise vbObjectError + 1, , "文本不能为空"   ' 演示 isError 语义
    End If

    Dim wordCount As Long, lineCount As Long, charCount As Long
    charCount = Len(txt)

    ' 行数：按换行拆分（兼容 \r\n 和 \n）
    Dim norm As String
    norm = Replace(txt, vbCrLf, vbLf)
    lineCount = 1
    Dim i As Long
    For i = 1 To Len(norm)
        If Mid$(norm, i, 1) = vbLf Then lineCount = lineCount + 1
    Next i
    If Len(norm) = 0 Then lineCount = 0

    ' 字数：按空白拆分为词
    Dim inWord As Boolean
    inWord = False
    wordCount = 0
    For i = 1 To Len(txt)
        Dim ch As String
        ch = Mid$(txt, i, 1)
        If ch = " " Or ch = vbTab Or ch = vbCr Or ch = vbLf Then
            If inWord Then inWord = False
        Else
            If Not inWord Then
                inWord = True
                wordCount = wordCount + 1
            End If
        End If
    Next i

    ITool_Execute = "字数: " & wordCount & vbCrLf _
        & "字符数: " & charCount & vbCrLf _
        & "行数: " & lineCount
End Function
