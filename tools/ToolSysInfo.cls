VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ToolSysInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' ToolSysInfo —— 示例工具：系统信息（只读，无副作用）
' 演示：调用 Windows API 的真实业务工具
' ============================================================
Option Explicit
Implements ITool

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Sub GlobalMemoryStatusEx Lib "kernel32" (ByRef lpBuffer As MEMORYSTATUSEX)
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToCaller As Currency, ByRef lpTotalNumberOfBytes As Currency, ByRef lpTotalNumberOfFreeBytes As Currency) As Long

Private Type MEMORYSTATUSEX
    dwLength As Long
    dwMemoryLoad As Long
    ullTotalPhys As Currency
    ullAvailPhys As Currency
    ullTotalPageFile As Currency
    ullAvailPageFile As Currency
    ullTotalVirtual As Currency
    ullAvailVirtual As Currency
    ullAvailExtendedVirtual As Currency
End Type

Private Property Get ITool_ToolName() As String
    ITool_ToolName = "sys_info"
End Property

Private Property Get ITool_Description() As String
    ITool_Description = "返回本机系统信息（机器名/用户/内存/磁盘）"
End Property

Private Property Get ITool_InputSchema() As String
    ITool_InputSchema = "{""type"":""object"",""properties"":{}}"
End Property

Private Function ITool_Execute(ByVal requestJson As String) As String
    Dim s As String
    s = "机器名: " & GetComputerNameStr() & vbCrLf _
      & "用户名: " & GetUserNameStr() & vbCrLf _
      & "内存使用率: " & GetMemoryLoad() & "%" & vbCrLf _
      & "C盘可用: " & GetDriveFree("C:\") & " MB"
    ITool_Execute = s
End Function

Private Function GetComputerNameStr() As String
    Dim buf As String
    buf = Space$(64)
    Dim n As Long
    n = 64
    Call GetComputerName(buf, n)
    GetComputerNameStr = TrimNull(buf)
End Function

Private Function GetUserNameStr() As String
    Dim buf As String
    buf = Space$(64)
    Dim n As Long
    n = 64
    Call GetUserName(buf, n)
    GetUserNameStr = TrimNull(buf)
End Function

Private Function GetMemoryLoad() As String
    Dim m As MEMORYSTATUSEX
    m.dwLength = LenB(m)
    Call GlobalMemoryStatusEx(m)
    GetMemoryLoad = CStr(m.dwMemoryLoad)
End Function

Private Function GetDriveFree(ByVal drive As String) As String
    Dim freeToCaller As Currency, total As Currency, freeTotal As Currency
    Call GetDiskFreeSpaceEx(drive, freeToCaller, total, freeTotal)
    ' Currency 是 1/10000 定点：值 / 10000 = 字节数
    GetDriveFree = CStr(CLng(CDbl(freeToCaller) * 10000# / 1048576#))
End Function

' API 返回的 ANSI 字符串以空字符结尾，截取到空字符之前
Private Function TrimNull(ByVal s As String) As String
    Dim pos As Long
    pos = InStr(1, s, vbNullChar)
    If pos > 0 Then TrimNull = Left$(s, pos - 1) Else TrimNull = Trim$(s)
End Function