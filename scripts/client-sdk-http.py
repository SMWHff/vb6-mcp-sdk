# client-sdk-http.py —— 官方 Python SDK 通过 Streamable HTTP 连接 vb6-mcp-sdk
# 前置：先启动 .\vb6-mcp-sdk.exe /http:9000
# 用法：uv run --with mcp python scripts/client-sdk-http.py [URL]
import asyncio
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from mcp import ClientSession
from mcp.client.streamable_http import streamable_http_client

URL = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:8080/mcp"


async def main():
    async with streamable_http_client(URL) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            print("[握手] 协议版本:", init.protocol_version)
            print("[握手] 服务器:", init.server_info.name, init.server_info.version)

            tools = await session.list_tools()
            print(f"[工具列表] 共 {len(tools.tools)} 个:", [t.name for t in tools.tools])

            r1 = await session.call_tool("add", {"a": 2, "b": 3})
            print("[add(2,3)]", r1.content[0].text)

            r2 = await session.call_tool("echo", {"text": "你好，HTTP-SDK"})
            print("[echo]", r2.content[0].text)

            r3 = await session.call_tool("get_time", {})
            print("[get_time]", r3.content[0].text)

            print("\n=== 官方 SDK 加载成功（Streamable HTTP）：全部工具可用 ===")


asyncio.run(main())
