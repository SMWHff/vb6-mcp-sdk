VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolMes —— MES 开放 API 工具（见教程文档，18 个接口）
' 统一 GET /mrs/getField；写接口（wifimac）走 /mrs/updateField
' 参数：
'   field   接口标识（友好名或原始 fieldName，白名单校验）
'   pcb_seq 主键（SN / MAC / 料号 / 工单号 / 组合）
'   value   写接口专用（set_mac 上传 MAC 地址）
'   server  MES 服务器地址（默认 http://192.168.20.151）
' ============================================================
Option Explicit
Implements ITool

Private Const DEFAULT_SERVER As String = "http://192.168.20.151"

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "mes_query"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "查询/写入 MES 开放 API（SN/MAC/料号/工单/治具/烧录文件）"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object""," _
        & """properties"":{" _
        & """field"":{""type"":""string"",""description"":""接口标识：mac/customer_sn/set_mac/next_route/repair_status/item_no/route_pass/item_name/bom_version/mac_bound/jig_control/jig_control_on/tool_register/tool_bind/mac_route_sn/firmware/work_order_sns/item_work_orders""}," _
        & """pcb_seq"":{""type"":""string"",""description"":""主键：SN/MAC/料号/工单号（如 BH08E901600001 或 BH08E901600001-A03）""}," _
        & """value"":{""type"":""string"",""description"":""写接口专用（set_mac 上传 MAC 地址）""}," _
        & """server"":{""type"":""string"",""description"":""MES 服务器地址，默认 http://192.168.20.151""}}," _
        & """required"":[""field"",""pcb_seq""]}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim field As String
    field = CStr(JsonGet(requestJson, ".params.arguments.field"))
    Dim pcbSeq As String
    pcbSeq = CStr(JsonGet(requestJson, ".params.arguments.pcb_seq"))
    Dim value As String
    value = CStr(JsonGet(requestJson, ".params.arguments.value"))
    Dim server As String
    server = CStr(JsonGet(requestJson, ".params.arguments.server"))
    If Len(server) = 0 Then server = DEFAULT_SERVER

    Dim fieldName As String
    fieldName = MapField(field)
    If Len(fieldName) = 0 Then
        Err.Raise vbObjectError + 1, , "未知 field: " & field & "（支持友好名或原始 fieldName）"
    End If
    If Len(pcbSeq) = 0 Then
        Err.Raise vbObjectError + 1, , "缺少 pcb_seq"
    End If

    Dim url As String
    If LCase$(fieldName) = "wifimac" Then
        If Len(value) = 0 Then
            Err.Raise vbObjectError + 1, , "set_mac 需要 value（MAC 地址）"
        End If
        url = server & "/mrs/updateField?pcbSeq=" & UrlEncode(pcbSeq) _
            & "&fieldName=wifimac&fieldValue=" & UrlEncode(value)
    Else
        url = server & "/mrs/getField?pcbSeq=" & UrlEncode(pcbSeq) _
            & "&fieldName=" & fieldName
    End If

    ITool_Execute = HttpGet(url)
End Function

' 接口标识 -> fieldName（友好名 + 原始名白名单）
Private Function MapField(ByVal f As String) As String
    Dim up As String
    up = LCase$(Trim$(f))
    Select Case up
    Case "mac":                  MapField = "mac"
    Case "customer_sn":          MapField = "shPcb"
    Case "set_mac":              MapField = "wifimac"
    Case "next_route":           MapField = "nextRoute"
    Case "repair_status":        MapField = "isRepairing"
    Case "item_no":              MapField = "ITEMNO"
    Case "route_pass":           MapField = "isPass"
    Case "item_name":            MapField = "ITEMNAME"
    Case "bom_version":          MapField = "BOMVER"
    Case "mac_bound":            MapField = "checkMAC"
    Case "jig_control":          MapField = "isJIGCODE"
    Case "jig_control_on":       MapField = "JIGCODE_ON"
    Case "tool_register":        MapField = "AddTool"
    Case "tool_bind":            MapField = "AddToolBindItem"
    Case "mac_route_sn":         MapField = "macSN"
    Case "firmware":             MapField = "FW"
    Case "work_order_sns":       MapField = "GetSNList"
    Case "item_work_orders":     MapField = "GetProdNoList"
    Case "shpcb", "nextroute", "isrepairing", "itemno", "ispass", "itemname", _
         "bomver", "checkmac", "isjigcode", "jigcode_on", "addtool", _
         "addtoolbinditem", "macsn", "fw", "getsnlist", "getprodnolist", "wifimac"
        MapField = f
    End Select
End Function

' HTTP GET（MSXML ServerXMLHTTP，超时 5s/15s）
Private Function HttpGet(ByVal url As String) As String
    Dim req As Object
    Set req = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    req.setTimeouts 5000, 5000, 5000, 15000
    req.Open "GET", url, False
    req.Send
    If req.Status <> 200 Then
        Err.Raise vbObjectError + 2, , "MES HTTP " & req.Status & ": " & Left$(req.responseText, 200)
    End If
    HttpGet = req.responseText
End Function

' URL 编码：字母数字 -_.~ 保留；ASCII 其它 %XX；非 ASCII 按 UTF-8
Private Function UrlEncode(ByVal s As String) As String
    Dim i As Long, c As Long
    Dim res As String
    res = ""
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or (c >= 48 And c <= 57) _
           Or c = 45 Or c = 95 Or c = 46 Or c = 126 Then
            res = res & ChrW$(c)
        ElseIf c < 128 Then
            res = res & "%" & HexByte(c)
        ElseIf c < &H800 Then
            res = res & "%" & HexByte(&HC0 Or (c \ 64)) & "%" & HexByte(&H80 Or (c And &H3F))
        ElseIf c < &H10000 Then
            res = res & "%" & HexByte(&HE0 Or (c \ 4096)) _
                & "%" & HexByte(&H80 Or ((c \ 64) And &H3F)) _
                & "%" & HexByte(&H80 Or (c And &H3F))
        Else
            res = res & "%" & HexByte(&HF0 Or (c \ 262144)) _
                & "%" & HexByte(&H80 Or ((c \ 4096) And &H3F)) _
                & "%" & HexByte(&H80 Or ((c \ 64) And &H3F)) _
                & "%" & HexByte(&H80 Or (c And &H3F))
        End If
    Next i
    UrlEncode = res
End Function

Private Function HexByte(ByVal b As Long) As String
    Dim h As String
    h = Hex$(b And &HFF)
    If Len(h) = 1 Then h = "0" & h
    HexByte = h
End Function
