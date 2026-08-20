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
            record("握手", "initialize 协议版本协商",
                   init.protocol_version in ("2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25"),
                   init.protocol_version)
            record("握手", "serverInfo 名称与版本",
                   init.server_info.name == "vb6-mcp-sdk-demo" and init.server_info.version == "1.0.0",
                   f"{init.server_info.name} {init.server_info.version}")
            caps = init.capabilities
            record("握手", "capabilities 含三大能力",
                   caps.tools is not None and caps.prompts is not None and caps.resources is not None,
                   str(caps))
            record("握手", "capabilities 含订阅/日志/补全",
                   getattr(caps.resources, "subscribe", False) is True
                   and getattr(caps, "logging", None) is not None
                   and getattr(caps, "completions", None) is not None, str(caps))
            await session.send_ping()
            record("握手", "ping 成功", True)

            # ---- Tools ----
            tools = await session.list_tools()
            names = [t.name for t in tools.tools]
            record("工具", "tools/list 返回 11 个工具", len(tools.tools) == 11, str(names))
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

            r = await session.call_tool("echo", {"text": "你好👋😀🚀"})
            record("工具", "echo emoji 4字节Unicode", r.content[0].text == "你好👋😀🚀", r.content[0].text[:20])
            r = await session.call_tool("echo", {"text": ""})
            record("工具", "echo 空字符串", r.content[0].text == "", repr(r.content[0].text))
            r = await session.call_tool("echo", {"text": "a" * 100000})
            record("工具", "echo 100KB 大文本", len(r.content[0].text) == 100000, len(r.content[0].text))

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

            r = await session.call_tool("json_build", {"mode": "empty"})
            record("工具", "json_build 空参数返回 {}", r.content[0].text == "{}", r.content[0].text)

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

            r = await session.call_tool("rand", {"min": 1, "max": 10})
            record("工具", "rand 区间内", 1 <= int(r.content[0].text) <= 10, r.content[0].text)
            vals = []
            for _ in range(20):
                r = await session.call_tool("rand", {"min": -5, "max": 5})
                vals.append(int(r.content[0].text))
            record("工具", "rand 负区间边界", all(-5 <= x <= 5 for x in vals), str(vals[:6]))
            r = await session.call_tool("rand", {"min": 10, "max": 1})
            record("工具", "rand 参数错误 isError", r.is_error is True and "max" in r.content[0].text,
                   r.content[0].text[:40])

            r = await session.call_tool("add", {"a": "x", "b": 1})
            record("工具", "add 类型错误 -> isError", r.is_error is True, r.content[0].text[:40])
            r = await session.call_tool("add", {"a": 1, "b": 2, "c": 99})
            record("工具", "额外参数忽略", r.content[0].text == "3", r.content[0].text)

            r = await session.call_tool("text_case", {"text": "Hello VB6", "mode": "upper"})
            record("工具", "text_case upper", r.content[0].text == "HELLO VB6", r.content[0].text)
            r = await session.call_tool("text_case", {"text": "  hello  ", "mode": "trim"})
            record("工具", "text_case trim", r.content[0].text == "hello", repr(r.content[0].text))
            r = await session.call_tool("text_case", {"text": "abc", "mode": "reverse"})
            record("工具", "text_case reverse", r.content[0].text == "cba", r.content[0].text)
            r = await session.call_tool("text_case", {"text": "x", "mode": "bogus"})
            record("工具", "text_case 非法mode isError", r.is_error is True, r.content[0].text[:30])
            r = await session.call_tool("getenv", {"name": "PATH"})
            record("工具", "getenv 存在变量", len(r.content[0].text) > 0, r.content[0].text[:20])
            r = await session.call_tool("getenv", {"name": "MCP_NO_SUCH_VAR_XYZ", "default": "fb"})
            record("工具", "getenv 默认值兜底", r.content[0].text == "fb", r.content[0].text)

            # ---- 结构化结果（structuredContent，2025-03-26+）----
            r = await session.call_tool("json_build", {"mode": "full"})
            sc = r.structured_content
            record("工具", "structuredContent 结构化结果",
                   isinstance(sc, dict) and sc.get("name") == "vb6" and sc.get("count") == 3, str(sc)[:80])
            r = await session.call_tool("echo", {"text": "plain"})
            record("工具", "非 JSON 文本无 structuredContent", r.structured_content is None,
                   str(r.structured_content)[:40])

            # ---- MES 工具（依赖内网 MES 服务器 192.168.20.151，接口见教程文档）----
            r = await session.call_tool("mes_query", {"field": "no_such_field", "pcb_seq": "BH08E901600001"})
            record("工具", "mes_query 非法 field 拒绝", r.is_error is True, r.content[0].text[:40])
            r = await session.call_tool("mes_query", {"field": "item_no", "pcb_seq": "BH08E901600001"})
            record("工具", "mes_query 查询料号", (not r.is_error) and r.content[0].text.strip() == "9001001114",
                   r.content[0].text[:40])
            r = await session.call_tool("mes_query", {"field": "ITEMNO", "pcb_seq": "BH08E901600001"})
            record("工具", "mes_query 原始 fieldName 透传",
                   (not r.is_error) and r.content[0].text.strip() == "9001001114", r.content[0].text[:40])
            r = await session.call_tool("mes_query", {"field": "set_mac", "pcb_seq": "BH08E901600001"})
            record("工具", "mes_query set_mac 缺 value 拒绝", r.is_error is True, r.content[0].text[:40])

            # ---- Prompts ----
            prompts = await session.list_prompts()
            record("提示词", "prompts/list 含 code_review",
                   any(p.name == "code_review" for p in prompts.prompts), [p.name for p in prompts.prompts])
            record("提示词", "prompts/list 含 translate",
                   any(p.name == "translate" for p in prompts.prompts), [p.name for p in prompts.prompts])
            pr = await session.get_prompt("code_review", {"language": "VB6", "code": "Dim x"})
            ptxt = pr.messages[0].content.text
            record("提示词", "prompts/get 包含参数", "VB6" in ptxt and "Dim x" in ptxt, ptxt[:40])
            pr = await session.get_prompt("translate", {"language": "英语", "text": "你好"})
            ptxt = pr.messages[0].content.text
            record("提示词", "translate 生成指令", "英语" in ptxt and "你好" in ptxt, ptxt[:40])
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

            # ---- Resource Templates ----
            tpls = await session.list_resource_templates()
            tpl_uris = [t.uri_template for t in tpls.resource_templates]
            record("资源", "templates/list 含 greet 模板", "demo://greet/{name}" in tpl_uris, str(tpl_uris))
            rt = await session.read_resource("demo://greet/张三")
            record("资源", "模板 read 动态解析", "你好，张三！" in rt.contents[0].text, rt.contents[0].text)

            try:
                await session.read_resource("demo://other/1")
                record("资源", "模板不匹配 -> 未知资源", False, "未报错")
            except MCPError as e:
                record("资源", "模板不匹配 -> 未知资源", e.code == -32602, f"code={e.code}")

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

    # ---- 边界与健壮性（参考官方 SDK 测试的输入覆盖）----
    out = raw_run([""])
    record("裸协议", "空消息 -> 无响应", len(out) == 0, str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"ping"'])
    record("裸协议", "截断 JSON -> -32700", any('"code":-32700' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"ping"} garbage'])
    record("裸协议", "尾随垃圾宽容处理", any('"id":1' in o and '"result"' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1.5,"method":"ping"}'])
    record("裸协议", "浮点 id 回显", any('"id":1.5' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":null,"method":"ping"}'])
    record("裸协议", "null id 回显", any('"id":null' in o and '"result"' in o for o in out), str(out))

    out = raw_run(['[{"jsonrpc":"2.0","id":1,"method":"ping"}]'])
    record("裸协议", "batch 数组 -> -32601", any('"code":-32601' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"PING"}'])
    record("裸协议", "方法名区分大小写", any('"code":-32601' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"ping","extra":{"x":1}}'])
    record("裸协议", "未知字段忽略", any('"id":1' in o and '"result"' in o for o in out), str(out))

    out = raw_run(['\ufeff{"jsonrpc":"2.0","id":1,"method":"ping"}'])
    record("裸协议", "BOM 前缀正常", any('"id":1' in o and '"result"' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":999999999999,"method":"ping"}'])
    record("裸协议", "超大 id 回显", any('"id":999999999999' in o for o in out), str(out))

    # ---- 协议扩展端点（2025-06-18+：补全/日志/订阅）----
    out = raw_run(['{"jsonrpc":"2.0","id":40,"method":"completion/complete",'
                   '"params":{"ref":{"type":"ref/prompt","name":"code_review"},'
                   '"argumentName":"language","arguments":{}}}'])
    record("裸协议", "completion/complete 补全语言", any('"VB6"' in o and "hasMore" in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":41,"method":"logging/setLevel","params":{"level":"debug"}}'])
    record("裸协议", "logging/setLevel 成功", any('"id":41' in o and '"result":{}' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":42,"method":"resources/subscribe","params":{"uri":"demo://server/info"}}',
                   '{"jsonrpc":"2.0","id":43,"method":"resources/unsubscribe","params":{"uri":"demo://server/info"}}'])
    record("裸协议", "resources/subscribe+unsubscribe",
           any('"id":42' in o and '"result":{}' in o for o in out)
           and any('"id":43' in o and '"result":{}' in o for o in out), str(out))

    # ---- 协议健壮性（参考官方 Python SDK 客户端测试思路）----
    out = raw_run(["bad json", '{"jsonrpc":"2.0","id":1,"method":"ping"}'])
    record("裸协议", "错误后恢复（坏消息后 ping 正常）",
           any('"code":-32700' in o for o in out)
           and any('"id":1' in o and '"result"' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2026-01-01"}}'])
    record("裸协议", "未来版本协商到最高支持",
           any('"protocolVersion":"2025-11-25"' in o for o in out), str(out)[:80])

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2023-01-01"}}'])
    record("裸协议", "极旧版本回退基线",
           any('"protocolVersion":"2024-11-05"' in o for o in out), str(out)[:80])

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"initialize"}'])
    record("裸协议", "initialize 无 params 容错",
           any('"protocolVersion":"2024-11-05"' in o for o in out), str(out)[:80])

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"initialize","params":{}}'])
    record("裸协议", "initialize 无 protocolVersion 容错",
           any('"protocolVersion":"2024-11-05"' in o for o in out), str(out)[:80])

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18"}}'])
    record("裸协议", "initialize 响应含三字段",
           any('"protocolVersion"' in o and '"capabilities"' in o and '"serverInfo"' in o for o in out),
           str(out)[:80])

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"tools/list","params":{"cursor":"abc"}}'])
    record("裸协议", "list 方法 cursor 参数兼容",
           any('"id":1' in o and '"tools"' in o for o in out), str(out)[:60])

    out = raw_run(['{"jsonrpc":"2.0","method":"ping"}'])
    record("裸协议", "无 id 请求宽松处理（按 null id 响应）",
           len(out) == 1 and '"id":null' in out[0] and '"result"' in out[0], str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":123}'])
    record("裸协议", "method 非字符串容错", any('"code":-32601' in o for o in out), str(out))

    out = raw_run(['{"jsonrpc":"2.0","id":1,"method":"tools/call","params":{"name":"get_time"}}'])
    record("裸协议", "tools/call 无 arguments 正常",
           any('"id":1' in o and '"result"' in o and '"text"' in o for o in out), str(out)[:80])


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
