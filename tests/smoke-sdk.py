#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""smoke-sdk.py - official Python MCP SDK smoke test (assert-style).

Follows the SDK's own example tests (tests/test_examples.py): drive the
server through the official `ClientSession` high-level API and assert on
structured fields (TextContent type, .text, structured_content, MCPError
codes) instead of raw-string matching. Complements tests/test.ps1 which
exercises the raw stdio pipe.

Run:
  uv run --with mcp python tests/smoke-sdk.py            # stdio smoke
  uv run --with mcp python tests/smoke-sdk.py http://localhost:9000/mcp  # + HTTP smoke
"""
import asyncio
import os
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client
from mcp.shared.exceptions import MCPError
from mcp.types import TextContent, TextResourceContents

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXE = os.path.join(BASE_DIR, "vb6-mcp-sdk.exe")

PASS = 0
FAIL = []


def check(name, cond, detail=""):
    global PASS
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL.append((name, detail))
        print(f"  FAIL  {name}  -> {detail}")


async def smoke_session(session, label):
    print(f"--- {label} ---")

    init = await session.initialize()
    check("initialize 协议版本", init.protocol_version in (
        "2024-11-05", "2025-03-26", "2025-06-18", "2025-11-25"), init.protocol_version)
    check("initialize serverInfo", init.server_info.name == "vb6-mcp-sdk-demo"
          and init.server_info.version == "1.0.0", f"{init.server_info.name} {init.server_info.version}")
    check("initialize capabilities",
          init.capabilities.tools is not None and init.capabilities.prompts is not None
          and init.capabilities.resources is not None, str(init.capabilities)[:80])

    tools = await session.list_tools()
    names = [t.name for t in tools.tools]
    check("tools/list 返回 11 个工具", len(tools.tools) == 11, str(len(tools.tools)))
    check("tools/list 含 add/getenv/mes_query",
          all(n in names for n in ("add", "getenv", "mes_query")), str(names)[:100])

    # 工具调用：结构化字段断言（参考 SDK test_examples.py）
    r = await session.call_tool("add", {"a": 2, "b": 3})
    check("call add(2,3) TextContent", len(r.content) == 1 and isinstance(r.content[0], TextContent),
          str(type(r.content[0]).__name__ if r.content else "empty"))
    check("call add(2,3)=5", r.content[0].text == "5", r.content[0].text)
    check("call add 无 structured_content", r.structured_content is None, str(r.structured_content)[:40])

    r = await session.call_tool("echo", {"text": "你好 SDK"})
    check("call echo 中文", r.content[0].text == "你好 SDK", r.content[0].text)
    r = await session.call_tool("echo", {"text": "🚀"})
    check("call echo emoji", r.content[0].text == "🚀", r.content[0].text[:20])

    r = await session.call_tool("json_build", {"mode": "full"})
    check("call json_build 带 structured_content",
          isinstance(r.structured_content, dict) and r.structured_content.get("name") == "vb6",
          str(r.structured_content)[:60])

    # 错误路径（MCPError 断言）
    try:
        await session.call_tool("no_such_tool", {})
        check("未知工具 -> -32602", False, "未报错")
    except MCPError as e:
        check("未知工具 -> -32602", e.code == -32602, f"code={e.code}")
    try:
        await session.call_tool("add", {"a": 1})
        check("缺必需参数 -> -32602", False, "未报错")
    except MCPError as e:
        check("缺必需参数 -> -32602", e.code == -32602, f"code={e.code}")

    # 资源 + 提示词 + 模板
    res = await session.read_resource("demo://server/info")
    check("read_resource 类型", len(res.contents) == 1
          and isinstance(res.contents[0], TextResourceContents), str(type(res.contents[0]).__name__))
    check("read_resource 内容", "vb6-mcp-sdk-demo" in res.contents[0].text, res.contents[0].text[:40])

    tpls = await session.list_resource_templates()
    check("templates/list 含 greet", "demo://greet/{name}" in [t.uri_template for t in tpls.resource_templates],
          str([t.uri_template for t in tpls.resource_templates]))
    rt = await session.read_resource("demo://greet/李雷")
    check("模板 read 动态解析", "你好，李雷！" in rt.contents[0].text, rt.contents[0].text[:30])

    pr = await session.get_prompt("code_review", {"language": "VB6", "code": "Dim x As Long"})
    check("get_prompt 返回内容", "VB6" in pr.messages[0].content.text
          and "Dim x As Long" in pr.messages[0].content.text, pr.messages[0].content.text[:40])

    await session.send_ping()
    check("ping 成功", True)

    print()


async def smoke_stdio():
    print("=== 官方 Python SDK 冒烟（stdio）===\n")
    server = StdioServerParameters(command=EXE, args=[], env=None)
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await smoke_session(session, "stdio")


async def smoke_http(url):
    print("=== 官方 Python SDK 冒烟（Streamable HTTP）===\n")
    from mcp.client.streamable_http import streamable_http_client

    async with streamable_http_client(url) as (read, write):
        async with ClientSession(read, write) as session:
            await smoke_session(session, f"http {url}")


def main():
    args = sys.argv[1:]
    if args:
        asyncio.run(smoke_http(args[0]))
    else:
        asyncio.run(smoke_stdio())

    print("=" * 52)
    print(f"结果: {PASS}/{PASS + len(FAIL)} 通过")
    for name, detail in FAIL:
        print(f"  FAIL  {name}: {detail}")
    print("=" * 52)
    sys.exit(0 if not FAIL else 1)


if __name__ == "__main__":
    main()
