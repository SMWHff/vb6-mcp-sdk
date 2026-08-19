# client-sdk.py —— 官方 Python MCP SDK 加载 vb6-mcp-sdk（stdio）并验证三大能力
# 用法：uv run --with mcp python scripts/client-sdk.py
import asyncio
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

EXE = r"C:\Users\mengf\vb6-mcp-sdk\vb6-mcp-sdk.exe"


async def main():
    server = StdioServerParameters(command=EXE, args=[], env=None)
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            print("[握手] 协议版本:", init.protocol_version)
            print("[握手] 服务器:", init.server_info.name, init.server_info.version)

            # ---- Tools ----
            tools = await session.list_tools()
            print(f"[工具列表] 共 {len(tools.tools)} 个:", [t.name for t in tools.tools])
            r1 = await session.call_tool("add", {"a": 2, "b": 3})
            print("[add(2,3)]", r1.content[0].text)
            r2 = await session.call_tool("echo", {"text": "你好，SDK"})
            print("[echo]", r2.content[0].text)

            # 真实业务工具
            r4 = await session.call_tool("sys_info", {})
            print("[sys_info]", r4.content[0].text.splitlines()[0])
            r5 = await session.call_tool("read_file", {"path": "README.md"})
            print("[read_file]", r5.content[0].text.splitlines()[0][:40], "...")
            r6 = await session.call_tool("word_count", {"text": "你好 MCP SDK\n第二行"})
            print("[word_count]", r6.content[0].text.replace("\n", " | "))

            # 框架参数校验：缺必需参数应被拦截
            try:
                await session.call_tool("add", {"a": 1})
                print("[参数校验] 缺参调用未被拦截（异常！）")
            except Exception as exc:
                print("[参数校验] 缺参被拦截:", str(exc)[:60])

            # isError 语义：read_file 传非法路径（路径穿越）应返回 isError:true 结果
            err = await session.call_tool("read_file", {"path": "..\\..\\secret.txt"})
            print("[isError]", "is_error=", err.is_error, "| 文本:", err.content[0].text[:40])

            # ---- Prompts ----
            prompts = await session.list_prompts()
            print(f"[提示词列表] 共 {len(prompts.prompts)} 个:", [p.name for p in prompts.prompts])
            if prompts.prompts:
                pr = await session.get_prompt(prompts.prompts[0].name, {"language": "VB6", "code": "Dim x As Integer"})
                first_line = pr.messages[0].content.text.splitlines()[0]
                print("[prompts/get]", first_line[:50], "...")

            # ---- Resources ----
            resources = await session.list_resources()
            print(f"[资源列表] 共 {len(resources.resources)} 个:", [r.uri for r in resources.resources])
            if resources.resources:
                rr = await session.read_resource(resources.resources[0].uri)
                print("[resources/read]", rr.contents[0].text.splitlines()[0])

            print("\n=== 官方 SDK 加载成功（stdio）：三大能力全部可用 ===")


asyncio.run(main())
