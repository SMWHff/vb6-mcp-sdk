# test.ps1 - 冒烟测试：JSON-RPC 消息管道灌给 vb6-mcp-sdk.exe，逐条结构化断言
# 覆盖：握手（版本协商）/ping/tools/prompts/resources/错误码/中文 UTF-8
# 用法：pwsh .\tests\test.ps1   （要求先编译 + fix-console）
$ErrorActionPreference = "Stop"

$exe = Join-Path $PSScriptRoot "..\vb6-mcp-sdk.exe"
if (-not (Test-Path $exe)) {
    Write-Host "未找到 vb6-mcp-sdk.exe，请先编译（且运行 fix-console.ps1）。" -ForegroundColor Red
    exit 1
}

$msgs = @(
    '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"smoke","version":"1.0"}}}'
    '{"jsonrpc":"2.0","id":2,"method":"ping"}'
    '{"jsonrpc":"2.0","id":3,"method":"tools/list"}'
    '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"add","arguments":{"a":2,"b":3}}}'
    '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"echo","arguments":{"text":"\u4f60\u597d SDK"}}}'
    '{"jsonrpc":"2.0","id":6,"method":"tools/call","params":{"name":"get_time"}}'
    '{"jsonrpc":"2.0","id":7,"method":"prompts/list"}'
    '{"jsonrpc":"2.0","id":8,"method":"resources/list"}'
    '{"jsonrpc":"2.0","id":9,"method":"no_such_method"}'
    'this is not json'
)

Write-Host "=== 冒烟测试：$($msgs.Count) 条 JSON-RPC 消息 ===" -ForegroundColor Cyan
$output = $msgs | & $exe

# 按 id 收集响应（非法 JSON 的 -32700 响应 id 为 null，单独识别）
$resp = @{}
$parseErr = $false
foreach ($line in $output) {
    try {
        $r = $line | ConvertFrom-Json
        if ($null -ne $r.id) { $resp["$($r.id)"] = $r }
        elseif ($r.error.code -eq -32700) { $parseErr = $true }
    } catch {
        if ($line -match '-32700') { $parseErr = $true }
    }
}
$output | ForEach-Object { Write-Host $_ }
Write-Host ""

$pass = 0; $fail = 0
function Check([string]$name, [bool]$cond, [string]$detail = "") {
    if ($cond) { $script:pass++; Write-Host "  PASS  $name" -ForegroundColor Green }
    else { $script:fail++; Write-Host "  FAIL  $name  -> $detail" -ForegroundColor Red }
}

# 1) 握手：版本协商 + 服务器信息 + capabilities
$r1 = $resp["1"]
Check "initialize 返回协议版本" ($r1.result.protocolVersion -in @("2024-11-05","2025-03-26","2025-06-18","2025-11-25")) "$($r1.result.protocolVersion)"
Check "initialize 服务器名" ($r1.result.serverInfo.name -eq "vb6-mcp-sdk-demo") "$($r1.result.serverInfo.name)"
Check "initialize capabilities 三能力" ($null -ne $r1.result.capabilities.tools -and $null -ne $r1.result.capabilities.prompts -and $null -ne $r1.result.capabilities.resources) ""

# 2) ping
$r2 = $resp["2"]
Check "ping 返回空 result" ($null -ne $r2.result) ""

# 3) tools/list：11 个工具且含 add
$r3 = $resp["3"]
$names = @($r3.result.tools | ForEach-Object { $_.name })
Check "tools/list 返回 11 个工具" ($names.Count -eq 11) "count=$($names.Count)"
Check "tools/list 含 add" ($names -contains "add") ($names -join ",")

# 4) 工具调用
$r4 = $resp["4"]
Check "add(2,3)=5" ($r4.result.content[0].text -eq "5") "$($r4.result.content[0].text)"
$r5 = $resp["5"]
Check "echo 中文原样" ($r5.result.content[0].text -eq "你好 SDK") "$($r5.result.content[0].text)"
$r6 = $resp["6"]
Check "get_time 格式" ($r6.result.content[0].text -match "^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$") "$($r6.result.content[0].text)"

# 5) prompts / resources
$r7 = $resp["7"]
$pnames = @($r7.result.prompts | ForEach-Object { $_.name })
Check "prompts/list 含 code_review" ($pnames -contains "code_review") ($pnames -join ",")
$r8 = $resp["8"]
$uris = @($r8.result.resources | ForEach-Object { $_.uri })
Check "resources/list 含 demo://server/info" ($uris -contains "demo://server/info") ($uris -join ",")

# 6) 错误处理
$r9 = $resp["9"]
Check "未知方法 -> -32601" ($r9.error.code -eq -32601) "code=$($r9.error.code)"
Check "非法 JSON -> -32700" $parseErr ""

Write-Host ""
if ($fail -eq 0) {
    Write-Host "=== 通过：$pass/$($pass + $fail) 冒烟全部正常 ===" -ForegroundColor Green
    exit 0
} else {
    Write-Host "=== 失败：$fail 项，请查看 logs\mcp.log ===" -ForegroundColor Red
    exit 1
}
