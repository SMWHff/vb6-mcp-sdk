Attribute VB_Name = "mcp_json"
' ============================================================
' mcp_json.bas -- SDK JSON utilities (pure VB6 via VBJSON)
' Replaces the old MSScriptControl + json-polyfill.js approach.
' JSON tree: object -> Scripting.Dictionary, array -> Collection
' (1-based Collection index; add 1 when mapping from JSON 0-based)
' NOTE: the VB6 Learning Edition runtime raises error 450 when an
' object is assigned into a Variant, so traversal uses only Set
' (object pointers) and reads leaves straight from Item().
' ============================================================
Option Explicit

Private m_cachedJson As String
Private m_cachedObj As Object

' Kept for compatibility with McpServer startup (VBJSON needs none)
Public Sub JsonInit()
End Sub

' Strip UTF-8 BOM (U+FEFF) at message start (some clients like .NET prepend it)
Private Function StripBom(ByVal s As String) As String
    If LenB(s) > 0 And Left$(s, 1) = ChrW(&HFEFF) Then s = Mid$(s, 2)
    StripBom = s
End Function

' Parse JSON and fetch the value at a dot path, e.g. ".params.arguments.a"
' Returns a scalar (String/Long/Double/Boolean) or Empty when the path is missing.
' Raises an error when the JSON itself is invalid (caller maps it to -32700).
Public Function JsonGet(ByVal json As String, ByVal path As String) As Variant
    On Error GoTo ErrHandler
    Dim obj As Object
    Set obj = ParseCached(json)
    If obj Is Nothing Then Err.Raise vbObjectError + 1, , "Invalid JSON"
    Dim parts() As String, j As Long
    parts = Split(path, ".")
    Dim node As Object
    Set node = obj
    For j = 1 To UBound(parts)
        If Len(parts(j)) = 0 Then Exit For
        Dim d As Scripting.Dictionary
        Dim c As Collection
        If TypeName(node) = "Dictionary" Then
            Set d = node
            If Not d.Exists(parts(j)) Then
                JsonGet = Empty
                Exit Function
            End If
            If TypeName(d.Item(parts(j))) = "Dictionary" Or TypeName(d.Item(parts(j))) = "Collection" Then
                Set node = d.Item(parts(j))
            Else
                JsonGet = d.Item(parts(j))
                Exit Function
            End If
        ElseIf TypeName(node) = "Collection" Then
            Set c = node
            On Error GoTo NotFound
            If TypeName(c.Item(CLng(parts(j)) + 1)) = "Dictionary" Or TypeName(c.Item(CLng(parts(j)) + 1)) = "Collection" Then
                Set node = c.Item(CLng(parts(j)) + 1)
            Else
                JsonGet = c.Item(CLng(parts(j)) + 1)
                Exit Function
            End If
            On Error GoTo 0
        Else
            JsonGet = Empty
            Exit Function
        End If
    Next j
    JsonGet = Empty
    Exit Function
NotFound:
    JsonGet = Empty
    Exit Function
ErrHandler:
    Err.Raise Err.Number
End Function

' VB6 string -> JSON string literal (quoted + escaped)
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

' Build a JSON object string from key/value pairs:
'   JsonBuild("name", "vb6", "count", 3, "ok", True)
'   -> {"name":"vb6","count":3,"ok":true}
' Value handling:
'   String        -> quoted + escaped (unless it starts with { or [ -> embedded as nested JSON)
'   Long/Integer/Byte/Single/Double/Currency -> raw number (comma decimal -> dot)
'   Boolean       -> true / false
'   Null / Empty  -> null
'   Date          -> quoted "yyyy-mm-dd hh:nn:ss"
'   other         -> quoted CStr
Public Function JsonBuild(ParamArray kv() As Variant) As String
    Dim parts As String
    parts = ""
    Dim i As Long
    For i = 0 To UBound(kv) Step 2
        If i + 1 <= UBound(kv) Then
            If Len(parts) > 0 Then parts = parts & ","
            parts = parts & JsonQuote(CStr(kv(i))) & ":" & JsonBuildValue(kv(i + 1))
        End If
    Next i
    JsonBuild = "{" & parts & "}"
End Function

Private Function JsonBuildValue(ByVal v As Variant) As String
    Dim t As Integer
    t = VarType(v)
    Select Case t
        Case vbString
            Dim sv As String
            sv = CStr(v)
            Dim h As String
            h = LTrim$(sv)
            If Left$(h, 1) = "{" Or Left$(h, 1) = "[" Then
                JsonBuildValue = sv
            Else
                JsonBuildValue = JsonQuote(sv)
            End If
        Case vbLong, vbInteger, vbByte, vbSingle, vbDouble, vbCurrency
            JsonBuildValue = Replace$(CStr(v), ",", ".")
        Case vbBoolean
            If v Then JsonBuildValue = "true" Else JsonBuildValue = "false"
        Case vbNull, vbEmpty
            JsonBuildValue = "null"
        Case vbDate
            JsonBuildValue = JsonQuote(Format$(v, "yyyy-mm-dd hh:nn:ss"))
        Case Else
            JsonBuildValue = JsonQuote(CStr(v))
    End Select
End Function

' Check request arguments against the schema "required" list.
' Returns missing names joined by "|", or "" when all are present.
Public Function JsonCheckRequired(ByVal schemaJson As String, ByVal requestJson As String) As String
    On Error GoTo Done
    Dim req As Scripting.Dictionary, s As Scripting.Dictionary
    Dim argsDict As Scripting.Dictionary, reqColl As Collection
    Dim rItem As Variant, missing As String
    Set req = ParseCached(StripBom(requestJson))
    Set s = ParseCached(StripBom(schemaJson))
    If req Is Nothing Or s Is Nothing Then GoTo Done
    Set argsDict = Nothing
    If req.Exists("params") Then
        If req.Item("params").Exists("arguments") Then
            Set argsDict = req.Item("params").Item("arguments")
        End If
    End If
    If Not s.Exists("required") Then GoTo Done
    Set reqColl = s.Item("required")
    For Each rItem In reqColl
        If argsDict Is Nothing Or Not argsDict.Exists(CStr(rItem)) Then
            If Len(missing) > 0 Then missing = missing & "|"
            missing = missing & CStr(rItem)
        End If
    Next rItem
    JsonCheckRequired = missing
    Exit Function
Done:
    JsonCheckRequired = ""
End Function

' Parse with per-message cache (HandleMessage + tool Execute share the same JSON text)
Private Function ParseCached(ByVal s As String) As Object
    s = StripBom(s)
    If s = m_cachedJson Then
        Set ParseCached = m_cachedObj
    Else
        Set m_cachedObj = Nothing
        Set m_cachedObj = json.parse(s)
        If Len(json.GetParserErrors()) > 0 Then Set m_cachedObj = Nothing
        m_cachedJson = s
        Set ParseCached = m_cachedObj
    End If
End Function
