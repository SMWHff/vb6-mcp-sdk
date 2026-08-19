VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SampleResource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ============================================================
' SampleResource —— 示例资源：服务器信息
' 演示：如何实现 IResource（动态内容资源）
' ============================================================
Option Explicit
Implements IResource

Private Property Get IResource_ResourceUri() As String
    IResource_ResourceUri = "demo://server/info"
End Property

Private Property Get IResource_ResourceName() As String
    IResource_ResourceName = "服务器信息"
End Property

Private Property Get IResource_ResourceMimeType() As String
    IResource_ResourceMimeType = "text/plain"
End Property

Private Property Get IResource_ResourceDescription() As String
    IResource_ResourceDescription = "SDK 演示服务器运行信息（动态生成）"
End Property

Private Function IResource_ReadResourceText(ByVal uri As String) As String
    IResource_ReadResourceText = "服务器: vb6-mcp-sdk-demo" & vbCrLf _
        & "框架: VB6.0 MCP SDK 1.0.0" & vbCrLf _
        & "时间: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf _
        & "说明: 这是一个通过 IResource 接口动态生成的文本资源"
End Function
