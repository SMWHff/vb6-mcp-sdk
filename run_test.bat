@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion

rem ============================================================
rem  run_test.bat —— vb6-mcp-sdk 一键测试启动器（放项目根目录）
rem
rem  流程：
rem    1) 环境自检与自安装：uv（缺失时自动安装）、PowerShell（无 7 用 5.1 兜底）
rem    2) 自动部署：exe 缺失时尝试 VB6 命令行编译；编译/已存在后自动跑 fix-console
rem    3) 自动运行全部 5 组测试：冒烟 / 85 用例套件 / HTTP 33 项 / 官方 SDK stdio / HTTP
rem    4) 汇总报告
rem
rem  退出码：0 = 全部通过    1 = 有测试失败    2 = 环境未就绪（exe 无法准备）
rem ============================================================

set "ROOT=%~dp0"
cd /d "%ROOT%"

rem ---------- 0) 定位 PowerShell 执行器（优先 pwsh 7，否则 powershell 5.1 兜底） ----------
set "PS=powershell -NoProfile -ExecutionPolicy Bypass"
where pwsh >nul 2>nul && set "PS=pwsh -NoProfile -ExecutionPolicy Bypass"

rem ---------- 0.5) 日志：自我 tee（输出同时进控制台与 logs 日志文件） ----------
for /f "delims=" %%i in ('%PS% -NoProfile -Command Get-Date -Format yyyyMMdd_HHmmss') do set "TS=%%i"
set "LOG=%ROOT%logs\run_test_%TS%.log"
if not exist "%ROOT%logs" mkdir "%ROOT%logs"
if "%TEE_DONE%"=="1" goto :tee_done
set "TEE_DONE=1"
%PS% -NoProfile -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; cmd /c '%~f0' 2>&1 | ForEach-Object { Write-Output $_; $_ | Out-File -Append -Encoding utf8 '%LOG%' }; exit $LASTEXITCODE"
exit /b %ERRORLEVEL%
:tee_done
echo [日志] 本次完整输出已记录到：%LOG%

echo ============================================================
echo  vb6-mcp-sdk 一键测试  ^| 工作目录：%ROOT%
echo ============================================================

rem ---------- 1) 自检 / 自安装：uv ----------
echo.
echo [1/3] 检查依赖 uv ...
set "UV=%USERPROFILE%\.local\bin\uv.exe"
if exist "%UV%" goto :uv_ok
where uv >nul 2>nul
if not errorlevel 1 (for /f "delims=" %%i in ('where uv') do set "UV=%%i" & goto :uv_ok)
echo [安装] 未找到 uv，正在自动安装（官方脚本）...
%PS% -Command "irm https://astral.sh/uv/install.ps1 | iex"
if exist "%UV%" goto :uv_ok
echo [错误] uv 自动安装失败，请手动安装：https://docs.astral.sh/uv/
exit 2
:uv_ok
echo [环境] uv 就绪
"%UV%" --version >nul 2>nul || (echo [错误] uv 不可执行 & exit 2)

rem ---------- 2) 部署：exe 检查 + 自动编译 + fix-console ----------
set "EXE=%ROOT%vb6-mcp-sdk.exe"
echo.
echo [2/3] 检查可执行文件 vb6-mcp-sdk.exe ...
rem ---- 查找 VB6 编译器（优先用户目录可写副本；(x86) 目录用 8.3 短路径避免 cmd 括号块解析问题）----
set "VB6="
if defined VB6_PATH set "VB6=%VB6_PATH%\vb6.exe"
if not defined VB6 if exist "%USERPROFILE%\VB6Expr\VB6.EXE" set "VB6=%USERPROFILE%\VB6Expr\VB6.EXE"
if not defined VB6 for /f "skip=2 tokens=2,*" %%a in ('reg query "HKLM\SOFTWARE\WOW6432Node\Microsoft\VisualStudio\6.0\Setup" /v VB6Path 2^>nul') do set "VB6=%%b\vb6.exe"
if not defined VB6 if exist "%ProgramFiles%\Microsoft Visual Studio\VB98\vb6.exe" set "VB6=%ProgramFiles%\Microsoft Visual Studio\VB98\vb6.exe"
for %%d in ("%ProgramFiles(x86)%") do set "PF86=%%~sd"
if not defined VB6 if exist "%PF86%\Microsoft Visual Studio\VB98\vb6.exe" set "VB6=%PF86%\Microsoft Visual Studio\VB98\vb6.exe"
if not defined VB6 if exist "%PF86%\VB6Expr\VB6.EXE" set "VB6=%PF86%\VB6Expr\VB6.EXE"

if not exist "%EXE%" (
    echo [部署] exe 不存在，尝试 VB6 命令行编译 ...
    if not defined VB6 (
        echo [错误] 未找到 VB6 编译器（vb6.exe），无法自动编译。
        echo         请手动编译：双击 vb6-mcp-sdk.vbp - 「文件 - 生成 vb6-mcp-sdk.exe」
        echo         编译完成后重新运行本脚本。
        exit 2
    )
    echo [编译] 使用 %VB6%
    rem 注意：VB6Expr 忽略 /out 参数，exe 按 vbp 的 ExeName32 输出到项目目录；start /wait 等待 GUI 程序
    start /wait "" "%VB6%" /make "%ROOT%vb6-mcp-sdk.vbp"
    if not exist "%EXE%" (
        echo [错误] VB6 编译失败，请检查 vbp 工程后手动编译。
        exit 2
    )
    findstr /b /m "MZ" "%EXE%" >nul 2>nul
    if errorlevel 1 (
        echo [错误] 编译产物异常（非有效 PE）。当前 VB6 可能不支持命令行编译，请手动编译后重跑。
        exit 2
    )
    echo [编译] 编译成功
)
echo [部署] 执行 fix-console.ps1（确保 Console 子系统，幂等）...
%PS% -File "%ROOT%scripts\fix-console.ps1"
if errorlevel 1 (echo [错误] fix-console 失败 & exit 2)

rem ---------- 3) 运行全部测试 ----------
set /a TPASS=0
set /a TFAIL=0

echo.
echo ============================================================
echo [3/3] 测试开始（5 组）
echo ============================================================

rem ---- 测试 1：stdio 冒烟 ----
echo.
echo ----- [1/5] 冒烟测试（stdio，10 条消息 / 13 项断言）-----
%PS% -File "%ROOT%tests\test.ps1"
if errorlevel 1 (echo [结果] 冒烟测试：失败 & set /a TFAIL+=1) else (echo [结果] 冒烟测试：通过 & set /a TPASS+=1)

rem ---- 测试 2：全面测试套件（85 用例，需 mcp 包，首次自动下载）----
echo.
echo ----- [2/5] 全面测试套件（85 用例：握手/工具/提示词/资源/安全/裸协议）-----
echo        （首次运行会通过 uv 自动下载 mcp 包，可能需要几分钟）
"%UV%" run --with mcp python "%ROOT%tests\test-suite.py"
if errorlevel 1 (echo [结果] 全面测试套件：失败 & set /a TFAIL+=1) else (echo [结果] 全面测试套件：通过 & set /a TPASS+=1)

rem ---- 测试 3：HTTP 传输层专项（33 项）----
echo.
echo ----- [3/5] HTTP 传输层专项（33 项）-----
echo        启动 vb6-mcp-sdk.exe /http:9002 ...
%PS% -Command "Start-Process -FilePath '%EXE%' -ArgumentList '/http:9002' -PassThru | ForEach-Object { $_.Id } | Out-File -Encoding ascii '%TEMP%\vb6mcp-sdk-9002.pid'"
ping 127.0.0.1 -n 3 >nul
"%UV%" run python "%ROOT%tests\http-test.py" http://localhost:9002/mcp
set "RC=!errorlevel!"
if exist "%TEMP%\vb6mcp-sdk-9002.pid" (
    for /f %%p in (%TEMP%\vb6mcp-sdk-9002.pid) do taskkill /pid %%p /f >nul 2>nul
    del "%TEMP%\vb6mcp-sdk-9002.pid" >nul 2>nul
)
if "!RC!"=="0" (echo [结果] HTTP 专项：通过 & set /a TPASS+=1) else (echo [结果] HTTP 专项：失败 & set /a TFAIL+=1)

rem ---- 测试 4：官方 Python SDK（stdio）----
echo.
echo ----- [4/5] 官方 Python SDK 完整握手（stdio）-----
"%UV%" run --with mcp python "%ROOT%tests\client-sdk.py"
if errorlevel 1 (echo [结果] 官方 SDK stdio：失败 & set /a TFAIL+=1) else (echo [结果] 官方 SDK stdio：通过 & set /a TPASS+=1)

rem ---- 测试 5：官方 Python SDK（HTTP）----
echo.
echo ----- [5/5] 官方 Python SDK（Streamable HTTP）-----
echo        启动 vb6-mcp-sdk.exe /http:9000 ...
%PS% -Command "Start-Process -FilePath '%EXE%' -ArgumentList '/http:9000' -PassThru | ForEach-Object { $_.Id } | Out-File -Encoding ascii '%TEMP%\vb6mcp-sdk-9000.pid'"
ping 127.0.0.1 -n 3 >nul
"%UV%" run --with mcp python "%ROOT%tests\client-sdk-http.py" http://localhost:9000/mcp
set "RC=!errorlevel!"
if exist "%TEMP%\vb6mcp-sdk-9000.pid" (
    for /f %%p in (%TEMP%\vb6mcp-sdk-9000.pid) do taskkill /pid %%p /f >nul 2>nul
    del "%TEMP%\vb6mcp-sdk-9000.pid" >nul 2>nul
)
if "!RC!"=="0" (echo [结果] 官方 SDK HTTP：通过 & set /a TPASS+=1) else (echo [结果] 官方 SDK HTTP：失败 & set /a TFAIL+=1)

echo.
echo ============================================================
echo  [4/3] 测试用例覆盖率（自动统计：工具/提示词/资源/模板/协议方法）
echo ============================================================
"%UV%" run --with mcp python "%ROOT%tests\coverage.py"

echo.
echo ============================================================
echo  [4/4] 白盒覆盖率（VB6 源码过程级：调用图可达性分析）
echo ============================================================
"%UV%" run python "%ROOT%tests\coverage-whitebox.py"

rem ---------- 4) 汇总 ----------
echo.
echo ============================================================
if %TFAIL% EQU 0 (
    echo 汇总：全部 %TPASS% 组测试通过 ✔
    echo ============================================================
    exit 0
)
echo 汇总：通过 %TPASS% 组，失败 %TFAIL% 组 ✘
echo ============================================================
echo 失败原因请查看上方输出；框架日志见 logs\mcp.log
exit 1
