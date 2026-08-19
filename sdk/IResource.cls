VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' IResource —— MCP 资源接口（SDK 使用者实现此接口注册 resource）
' 实现：类模块写 Implements IResource，实现
'   IResource_ResourceUri / IResource_ResourceName /
'   IResource_ResourceMimeType / IResource_ResourceDescription /
'   IResource_ReadResourceText
' 注册：server.RegisterResource New 你的类
' ============================================================
Option Explicit

' 资源 URI（唯一，标准 URI 格式，如 demo://server/info）
Public Property Get ResourceUri() As String
End Property

' 资源名（显示在 resources/list）
Public Property Get ResourceName() As String
End Property

' MIME 类型（如 text/plain、application/json）
Public Property Get ResourceMimeType() As String
End Property

' 资源描述
Public Property Get ResourceDescription() As String
End Property

' 读取资源内容。
' uri = 客户端请求的 URI（框架已校验匹配）。
' 返回：资源文本内容；读取失败可抛 Err（框架转 -32602）。
Public Function ReadResourceText(ByVal uri As String) As String
End Function
