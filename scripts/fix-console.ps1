# fix-console.ps1 —— 把 vb6mcp-sdk.exe 从 GUI 子系统改为控制台子系统
# 为什么：PowerShell 管道无法与 GUI 子系统（subsystem=2）程序交换 stdio；
#         改成控制台子系统（subsystem=3）后，pwsh 管道与 MCP 客户端都能正常通信。
# 用法：每次重新编译后运行一次：pwsh .\scripts\fix-console.ps1
$ErrorActionPreference = "Stop"

$exe = Join-Path $PSScriptRoot "..\vb6mcp-sdk.exe"
if (-not (Test-Path $exe)) {
    Write-Host "未找到 vb6mcp-sdk.exe，请先编译。" -ForegroundColor Red
    exit 1
}

$b = [IO.File]::ReadAllBytes($exe)
if ($b.Length -lt 0x40) { throw "无效的 PE 文件" }

$peOff = [BitConverter]::ToInt32($b, 0x3C)
$subOff = $peOff + 24 + 68     # Optional header 中 Subsystem 字段偏移
$old = [BitConverter]::ToUInt16($b, $subOff)

if ($old -eq 3) {
    Write-Host "已是控制台子系统（Console=3），无需修改。" -ForegroundColor Yellow
    exit 0
}
if ($old -ne 2) {
    Write-Host "意外的子系统值: $old，拒绝修改（期望 GUI=2 或 Console=3）。" -ForegroundColor Red
    exit 1
}

$b[$subOff] = 3
$b[$subOff + 1] = 0
[IO.File]::WriteAllBytes($exe, $b)

$b2 = [IO.File]::ReadAllBytes($exe)
$v = [BitConverter]::ToUInt16($b2, $subOff)
if ($v -eq 3) {
    Write-Host "已将 vb6mcp-sdk.exe 从 GUI(2) 改为 Console(3)。" -ForegroundColor Green
} else {
    Write-Host "修改失败：子系统 = $v" -ForegroundColor Red
    exit 1
}
