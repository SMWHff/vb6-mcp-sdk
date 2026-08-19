' ============================================================
' mcp_json.bas —— SDK JSON 工具（MSScriptControl + polyfill）
' 注意1：polyfill 用二进制模式读取——文本模式 Input$ + LOF 遇 CRLF 会抛
'        「错误 62 输入超出文件尾」；二进制模式对 CRLF/LF 通吃。
' 注意2：所有取值用 JsonGet 在 JS 表达式里取【标量】——VB6 late binding
'        访问 ScriptControl 返回对象的嵌套属性会抛「错误 438」。
' 要求：json-polyfill.js 与 exe 同目录（App.Path 定位）
' ============================================================
Option Explicit

Private m_sc As Object

' 初始化 JScript 引擎并注入 JSON.parse polyfill（幂等，可重复调用）
Public Sub JsonInit()
    Set m_sc = CreateObject("MSScriptControl.ScriptControl")
    m_sc.Language = "JScript"
    m_sc.AllowUI = False
    m_sc.UseSafeSubset = True

    Dim f As Integer, n As Long
    f = FreeFile
    Open App.Path & "\json-polyfill.js" For Binary As #f
    n = LOF(f)
    Dim buf() As Byte
    If n > 0 Then
        ReDim buf(0 To n - 1)
        Get #f, , buf
    End If
    Close #f
    If n <= 0 Then Err.Raise 53

    m_sc.AddCode Utf8Decode(buf, n)
End Sub

' 在 JS 表达式里解析 JSON 并取 path 指向的值（返回标量；取不到返回 Empty）
' 例：JsonGet(json, ".params.arguments.a")
Public Function JsonGet(ByVal json As String, ByVal path As String) As Variant
    JsonGet = m_sc.Eval("JSON.parse(" & JsonQuote(json) & ")" & path)
End Function

' VB6 字符串 -> JS 字符串字面量（双引号包裹 + 完整转义；先 \ 后 "）
Public Function JsonQuote(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    s = Replace(s, ChrW(&H2028), "\u2028")
    s = Replace(s, ChrW(&H2029), "\u2029")
    JsonQuote = """" & s & """"
End Function

' 校验请求的 arguments 是否包含 schema 声明的全部必需参数
' 返回：缺失的参数名（| 分隔）；空字符串 = 校验通过
Public Function JsonCheckRequired(ByVal schemaJson As String, ByVal requestJson As String) As String
    Dim expr As String
    expr = "(function(){var req=JSON.parse(" & JsonQuote(requestJson) & ").params.arguments||{};" _
        & "var s=JSON.parse(" & JsonQuote(schemaJson) & ");var r=s.required||[];var m=[];" _
        & "for(var i=0;i<r.length;i++){if(typeof req[r[i]]==='undefined')m.push(r[i]);}" _
        & "return m.join('|');})()"
    JsonCheckRequired = CStr(m_sc.Eval(expr))
End Function