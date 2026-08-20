# test-suite.py —— vb6-mcp-sdk 全面测试套件（官方 SDK + 裸 stdio 协议错误）
# 用法：uv run --with mcp python scripts/test-suite.py
# 前置：已编译 vb6-mcp-sdk.exe 且跑过 fix-console.ps1（GUI 子系统下裸管道读不到输出）
import asyncio
import json
import os
import re
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.shared.exceptions import MCPError

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXE = os.path.join(BASE_DIR, "vb6-mcp-sdk.exe")
EXE_DIR = os.path.dirname(EXE)

RESULTS = []


def record(category, name, passed, detail=""):
    RESULTS.append((category, name, passed, detail))
    mark = "PASS" if passed else "FAIL"
    print(f"  {mark}  [{category}] {name}" + ("" if passed else f"  -> {detail}"))


# ==================== 官方 SDK 会话用例 ====================
async def sdk_session_cases():
    server = StdioServerParameters(command=EXE, args=[], env=None)
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            record("握手", "initialize 协议版本 2024-11-05",
                   init.protocol_version == "2024-11-05", init.protocol_version)
            record("握手", "serverInfo 名称与版本",
                   init.server_info.name == "vb6-mcp-sdk-demo" and init.server_info.version == "1.0.0",
                   f"{init.server_info.name} {init.server_info.version}")
            caps = init.capabilities
            record("握手", "capabilities 含三大能力",
                   caps.tools is not None and caps.prompts is not None and caps.resources is not None,
                   str(caps))
            await session.send_ping()
            record("握手", "ping 成功", True)

            # ---- Tools ----
            tools = await session.list_tools()
            names = [t.name for t in tools.tools]
            record("工具", "tools/list 返回 7 个工具", len(tools.tools) == 7, str(names))
            record("工具", "工具字段完整",
                   all(t.name and t.description and t.input_schema.get("type") == "object" for t in tools.tools))

            r = await session.call_tool("add", {"a": 2, "b": 3})
            record("工具", "add(2,3)=5", r.content[0].text == "5", r.content[0].text)
            r = await session.call_tool("add", {"a": -1.5, "b": 0.25})
            record("工具", "add 负数与小数", r.content[0].text == "-1.25", r.content[0].text)

            r = await session.call_tool("echo", {"text": "你好，MCP SDK"})
            record("工具", "echo 中文", r.content[0].text == "你好，MCP SDK", r.content[0].text)
            special = '说"引号" 用\\斜杠 换\n行'
            r = await session.call_tool("echo", {"text": special})
            record("工具", "echo 特殊字符", r.content[0].text == special, repr(r.content[0].text))
            big = "长" * 5000
            r = await session.call_tool("echo", {"text": big})
            record("工具", "echo 长文本 10KB", len(r.content[0].text) == 5000, len(r.content[0].text))

            r = await session.call_tool("get_time", {})
            record("工具", "get_time 格式", bool(re.match(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$", r.content[0].text)),
                   r.content[0].text)
            r = await session.call_tool("sys_info", {})
            record("工具", "sys_info 含机器名", "机器名" in r.content[0].text, r.content[0].text.splitlines()[0])
            r = await session.call_tool("read_file", {"path": "README.md"})
            record("工具", "read_file 读 README.md", "# vb6-mcp-sdk" in r.content[0].text,
                   r.content[0].text.splitlines()[0][:30])
            r = await session.call_tool("word_count", {"text": "你好 MCP SDK\n第二行"})
            txt = r.content[0].text
            record("工具", "word_count 统计", "字数: 4" in txt and "行数: 2" in txt, txt.replace("\n", "|"))

            r = await session.call_tool("json_build", {})
            jb = r.content[0].text
            try:
                pj = json.loads(jb)
                record("工具", "json_build 返回合法 JSON", isinstance(pj, dict), jb[:60])
                record("工具", "json_build 类型正确",
                       pj.get("name") == "vb6" and pj.get("count") == 3 and pj.get("ratio") == 3.5
                       and pj.get("ok") is True and pj.get("no") is False and pj.get("nil") is None
                       and isinstance(pj.get("nested"), dict) and pj["nested"].get("a") == 1, jb[:80])
            except Exception as _e:
                record("工具", "json_build 返回合法 JSON", False, str(_e))

            try:
                await session.call_tool("no_such_tool", {})
                record("工具", "未知工具 -> error", False, "未报错")
            except MCPError as e:
                record("工具", "未知工具 -> error", e.code == -32602, f"code={e.code}")
            try:
                await session.call_tool("add", {"a": 1})
                record("工具", "缺必需参数 -> 拦截", False, "未拦截")
            except MCPError as e:
                record("工具", "缺必需参数 -> 拦截", "缺少必需参数" in e.message, e.message)
            r = await session.call_tool("read_file", {"path": "..\\..\\secret.txt"})
            record("工具", "工具抛错 -> isError", r.is_error is True and "执行失败" in r.content[0].text,
                   f"is_error={r.is_error}")

            # ---- Prompts ----
            prompts = await session.list_prompts()
            record("提示词", "prompts/list 含 code_review",
                   any(p.name == "code_review" for p in prompts.prompts), [p.name for p in prompts.prompts])
            pr = await session.get_prompt("code_review", {"language": "VB6", "code": "Dim x"})
            ptxt = pr.messages[0].content.text
            record("提示词", "prompts/get 包含参数", "VB6" in ptxt and "Dim x" in ptxt, ptxt[:40])
            try:
                await session.get_prompt("no_such_prompt", {})
                record("提示词", "未知提示词 -> 错误", False, "未报错")
            except MCPError as e:
                record("提示词", "未知提示词 -> 错误", e.code == -32602, f"code={e.code}")

            # ---- Resources ----
            resources = await session.list_resources()
            record("资源", "resources/list 含 demo://server/info",
                   any(rr.uri == "demo://server/info" for rr in resources.resources),
                   [rr.uri for rr in resources.resources])
            rr = await session.read_resource("demo://server/info")
            record("资源", "resources/read 返回内容", "vb6-mcp-sdk-demo" in rr.contents[0].text,
                   rr.contents[0].text.splitlines()[0])
            try:
                await session.read_resource("demo://no/such")
                record("资源", "未知资源 -> 错误", False, "未报错")
            except MCPError as e:
                record("资源", "未知资源 -> 错误", e.code == -32602, f"code={e.code}")

            # ---- 安全 ----
            r = await session.call_tool("read_file", {"path": "..\\..\\Windows\\win.ini"})
            record("安全", "路径穿越拒绝", r.is_error and "非法路径" in r.content[0].text, r.content[0].text[:40])
            r = await session.call_tool("read_file", {"path": "C:\\Windows\\win.ini"})
            record("安全", "绝对路径拒绝", r.is_error, r.content[0].text[:40])
            r = await session.call_tool("read_file", {"path": "secret.exe"})
            record("安全", "非法扩展名拒绝", r.is_error, r.content[0].text[:40])


# ==================== 裸 stdio 协议用例 ====================
def raw_run(messages):
    p = subprocess.Popen([EXE], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                         stderr=subprocess.DEVNULL, cwd=EXE_DIR)
    for m in messages:
        p.stdin.write(m.encode("utf-8") + b"\n")
    p.stdin.flush()
    p.stdin.close()
    out = []
    while True:
        line = p.stdout.readline()
        if not line:
            break
        out.append(line.decode("utf-8", errors="replace").strip())
    p.wait(timeout=10)
    return out


def raw_cases():
    out = raw_run(["this is not json"])
    record("裸协议", "非法 JSON -> -32700", any('"code":-32700' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":9,"method":"no_such_method"}'])
    record("裸协议", "未知方法 -> -32601", any('"code":-32601' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","method":"notifications/initialized"}'])
    record("裸协议", "通知 -> 无响应", len(out) == 0, str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":"abc","method":"ping"}'])
    record("裸协议", "字符串 id 正确回显", any('"id":"abc"' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"ping"}'])
    record("裸协议", "数字 id 回显", any('"id":1' in o for o in out), str(out))

    p = subprocess.Popen([EXE], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                         stderr=subprocess.DEVNULL, cwd=EXE_DIR)
    p.stdin.write(b'{"jsonrpc":"2.0","id":1,"method":"ping"}\r\n')
    p.stdin.flush()
    p.stdin.close()
    line = p.stdout.readline().decode("utf-8", errors="replace").strip()
    p.wait(timeout=10)
    record("裸协议", "CRLF 行尾兼容", '"id":1' in line and '"result"' in line, line)


# ==================== 汇总 ====================
def main():
    asyncio.run(sdk_session_cases())
    raw_cases()

    passed = sum(1 for _, _, ok, _ in RESULTS if ok)
    print("\n" + "=" * 56)
    print(f"结果: {passed}/{len(RESULTS)} 通过")
    cats = {}
    for cat, _, ok, _ in RESULTS:
        cats.setdefault(cat, [0, 0])
        cats[cat][0] += 1
        if ok:
            cats[cat][1] += 1
    for cat, (total, ok) in sorted(cats.items()):
        print(f"  {cat:<6} {ok}/{total}")
    failed = [(c, n, d) for c, n, ok, d in RESULTS if not ok]
    if failed:
        print("\n失败明细:")
        for c, n, d in failed:
            print(f"  [{c}] {n}: {d}")
    sys.exit(0 if passed == len(RESULTS) else 1)


main()
