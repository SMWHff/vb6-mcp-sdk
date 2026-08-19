# ============================================================
# fix-vb6ext.ps1 —— 修复 VB6 启动报「不能 'VB6EXT.OLB' 注册」
#
# 问题：VB6 每次启动都会把扩展库 VB6EXT.OLB 注册到
#       HKLM\SOFTWARE\WOW6432Node\Classes\TypeLib，
#       普通用户对该键无写权限 → 注册失败 → 每次启动弹窗。
#       （命令行编译 /make 不加载 IDE 扩展检查，所以不弹；
#        双击打开 .vbp 触发，才会弹。）
#
# 修复：① 用 32 位进程 LoadTypeLibEx 注册两个位置的 VB6EXT.OLB
#       ② 给 TypeLib 键授予 Users 完全控制（VB6 启动会写该键）
#       ③ 验证注册结果
#
# 用法：powershell -ExecutionPolicy Bypass -File scripts\fix-vb6ext.ps1
#       脚本会自动请求管理员权限（UAC 弹窗点「是」）
# ============================================================

# ---- 自动检测管理员，非管理员则提权（用 32 位 PowerShell 保证注册到 WOW6432Node）----
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "需要管理员权限，正在请求提升（请在弹出的 UAC 中点击「是」）..."
    $psPath = "$env:WINDIR\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
    if (-not (Test-Path $psPath)) { $psPath = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe" }
    Start-Process $psPath -Verb RunAs -ArgumentList '-NoProfile','-ExecutionPolicy','Bypass','-File',"`"$PSCommandPath`""
    exit
}

$ErrorActionPreference = "Stop"
Write-Host "=== 修复 VB6EXT.OLB 注册 ===" -ForegroundColor Cyan

# ---- ① 注册 OLB 类型库 ----
Add-Type -TypeDefinition @"
using System.Runtime.InteropServices;
public class VB6TLB {
    [DllImport("oleaut32.dll", CharSet=CharSet.Unicode)]
    public static extern int LoadTypeLibEx(string f, uint k, out System.IntPtr p);
}
"@

$olbFiles = @(
    "$env:USERPROFILE\VB6Expr\VB6EXT.OLB",
    "C:\Program Files (x86)\VB6Expr\VB6EXT.OLB"
)
$ok = $true
foreach ($olb in $olbFiles) {
    if (-not (Test-Path $olb)) { Write-Host "  跳过（文件不存在）: $olb" -ForegroundColor DarkGray; continue }
    $p = [IntPtr]::Zero
    $hr = [VB6TLB]::LoadTypeLibEx($olb, 1, [ref]$p)   # 1 = REGKIND_REGISTER
    if ($hr -eq 0) { Write-Host "  ✓ 注册成功: $olb" -ForegroundColor Green }
    else { Write-Host "  ✗ 注册失败 (0x$('{0:X8}' -f $hr)): $olb" -ForegroundColor Red; $ok = $false }
}

# ---- ② 授权 TypeLib 键（VB6 启动会写该键）----
$key = "HKLM:\SOFTWARE\WOW6432Node\Classes\TypeLib"
try {
    $acl = Get-Acl $key
    $rule = New-Object System.Security.AccessControl.RegistryAccessRule("Users","FullControl","ContainerInherit,ObjectInherit","None","Allow")
    $acl.SetAccessRule($rule)
    Set-Acl $key $acl
    Write-Host "  ✓ 已授予 Users 对 TypeLib 键的写权限" -ForegroundColor Green
} catch {
    Write-Host "  ✗ 授权失败: $($_.Exception.Message)" -ForegroundColor Red
    $ok = $false
}

# ---- ③ 验证 ----
$chk = reg query "HKLM\SOFTWARE\WOW6432Node\Classes\TypeLib" /s 2>$null | Select-String "VB6EXT"
if ($chk) { Write-Host "  ✓ 验证通过：VB6EXT.OLB 已注册（$($chk.Count) 处条目）" -ForegroundColor Green }
else { Write-Host "  ⚠ 验证未找到 VB6EXT 条目，请检查 VB6Expr 安装位置" -ForegroundColor Yellow }

Write-Host ""
if ($ok) { Write-Host "修复完成，请重新双击 vb6-mcp-sdk.vbp 测试（应不再弹窗）。" -ForegroundColor Green }
else { Write-Host "部分步骤失败，请检查上方输出。" -ForegroundColor Red }
