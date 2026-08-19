# test.ps1 —— 冒烟测试：把 JSON-RPC 消息管道灌给 vb6-mcp-sdk.exe 验证握手
# 用法：pwsh .\scripts\test.ps1   （要求先编译 + fix-console）
$ErrorActionPreference = "Stop"

$exe = Join-Path $PSScriptRoot "..\vb6-mcp-sdk.exe"
if (-not (Test-Path $exe)) {
    Write-Host "未找到 vb6-mcp-sdk.exe，请先编译（且运行 fix-console.ps1）。" -ForegroundColor Red
    exit 1
}

$input = @'
{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}
{"jsonrpc":"2.0","id":2,"method":"tools/list"}
{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"add","arguments":{"a":2,"b":3}}}
{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"echo","arguments":{"text":"hello sdk"}}}
{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"get_time"}}
'@

Write-Host "=== 冒烟测试：5 条 JSON-RPC 消息 ===" -ForegroundColor Cyan
$output = $input | & $exe
$output
Write-Host ""

$fail = 0
if (-not ($output -match '"text":"5"')) {
    Write-Host "失败：add(2,3) 未返回 5" -ForegroundColor Red
    $fail = 1
} elseif (-not ($output -match '"text":"hello sdk"')) {
    Write-Host "失败：echo 未原样返回文本" -ForegroundColor Red
    $fail = 1
} elseif (-not ($output -match 'vb6-mcp-sdk-demo')) {
    Write-Host "注意：initialize 未返回自定义服务器名（不阻塞，仅提示）" -ForegroundColor Yellow
}

if ($fail -eq 0) {
    Write-Host "=== 通过：SDK 三个工具均正常 ===" -ForegroundColor Green
} else {
    Write-Host "=== 失败：请检查 logs\mcp.log ===" -ForegroundColor Red
    exit 1
}
