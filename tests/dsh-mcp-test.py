#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""dsh-mcp-test.py - DeepSeek Harness (DSH) MCP compatibility test.

DSH loads MCP servers through the `@deepseek-ai/dsh-mcp-client` cordis
component in `profiles/<profile>/cordis.patch.yml` (same serverName scheme
as mcp__mes__ etc.). This test:

  1. Generates the cordis patch block that registers vb6-mcp-sdk.
  2. Validates the block (required fields, command/args, transport).
  3. Boots the server exactly as DSH would (command + args over stdio) and
     performs the full handshake DSH's client does: initialize -> tools/list
     -> tools/call, asserting the server answers correctly.
  4. Prints how to enable it in the current DSH profile.

Run: uv run python tests/dsh-mcp-test.py
"""
import json
import os
import re
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXE = os.path.join(BASE_DIR, "vb6-mcp-sdk.exe")
SERVER_NAME = "vb6-mcp-sdk"

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


def gen_config_block():
    """The cordis.patch.yml insert block DSH needs to register the server."""
    exe_fwd = EXE.replace("\\", "/")
    return (
        "# vb6-mcp-sdk: VB6 MCP Server（本仓库编译产物，stdio 传输）。\n"
        "# 提供 mcp__vb6-mcp-sdk__* 工具：add/echo/get_time/sys_info/read_file/\n"
        "# word_count/json_build/rand/text_case/getenv/mes_query。\n"
        "- insert:\n"
        "    - id: mcp-vb6-mcp-sdk\n"
        "      name: '@deepseek-ai/dsh-mcp-client'\n"
        "      config:\n"
        "        serverName: " + SERVER_NAME + "\n"
        "        transport: stdio\n"
        "        command: " + exe_fwd + "\n"
        "        args: []\n"
        "        toolCallTimeoutMs: 30000\n"
        "        failOnStartupError: false\n"
    )


def validate_block(block):
    """Field-level validation of the generated cordis block."""
    checks = []
    checks.append(("serverName", SERVER_NAME in block))
    checks.append(("component name", "@deepseek-ai/dsh-mcp-client" in block))
    checks.append(("transport stdio", "transport: stdio" in block))
    checks.append(("command 指向 exe", "command: " + EXE.replace("\\", "/") in block))
    checks.append(("args 数组", "args: []" in block))
    checks.append(("failOnStartupError 兜底", "failOnStartupError: false" in block))
    ok = all(c for _, c in checks)
    for name, c in checks:
        check(f"配置校验: {name}", c, block[:200] if not c else "")
    return ok


def run_stdio_exchange():
    """Boot the server the way DSH does (stdio, command+args) and handshake."""
    msgs = [
        {"jsonrpc": "2.0", "id": 1, "method": "initialize",
         "params": {"protocolVersion": "2025-11-25", "capabilities": {},
                    "clientInfo": {"name": "dsh-mcp-client", "version": "0.0.0"}}},
        {"jsonrpc": "2.0", "id": 2, "method": "tools/list"},
        {"jsonrpc": "2.0", "id": 3, "method": "tools/call",
         "params": {"name": "add", "arguments": {"a": 2, "b": 3}}},
        {"jsonrpc": "2.0", "id": 4, "method": "ping"},
    ]
    p = subprocess.Popen([EXE], stdin=subprocess.PIPE, stdout=subprocess.PIPE,
                         stderr=subprocess.DEVNULL, cwd=BASE_DIR)
    for m in msgs:
        p.stdin.write((json.dumps(m) + "\n").encode("utf-8"))
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


def main():
    print("=== DeepSeek Harness (DSH) MCP 兼容性测试 ===\n")
    print(f"  服务器: {EXE}\n")

    # 1) 生成 + 校验配置块
    block = gen_config_block()
    print("  [1/3] 生成的 DSH 配置块（cordis.patch.yml insert）：")
    print("  " + block.replace("\n", "\n  "))
    validate_block(block)

    # 2) 以 DSH 客户端方式真实握手
    print("\n  [2/3] 按 DSH dsh-mcp-client 方式启动服务器并握手…")
    out = run_stdio_exchange()

    def has(pat, mid=None):
        for o in out:
            if mid is not None and f'"id":{mid}' not in o:
                continue
            if re.search(pat, o):
                return True
        return False

    check("initialize 返回协议版本", has(r'"protocolVersion":"2025-11-25"', 1), str(out)[:200])
    check("initialize 返回 capabilities", has(r'"capabilities"', 1), str(out)[:200])
    try:
        tools_n = len(json.loads(out[1])["result"]["tools"])
    except Exception as e:
        tools_n = -1
        check("tools/list 返回 11 个工具", False, str(e))
    check("tools/list 返回 11 个工具", tools_n == 11, f"tools={tools_n}")
    check("tools/call add(2,3) = 5", has(r'"text":"5"', 3), str(out[2])[:150])
    check("ping 正常响应", has(r'"result"', 4), str(out[3])[:120])

    # 3) 接入提示
    print("\n  [3/3] 接入 DSH（将上述配置块追加到当前 profile 的 cordis.patch.yml）")
    dsh_home = os.environ.get("DSH_HOME", os.path.expanduser("~/.dsh"))
    prof_dir = os.path.join(dsh_home, "profiles", "web")
    active = os.path.join(prof_dir, "cordis.patch.yml")
    check("找到 DSH web profile 配置", os.path.exists(active), active)
    if os.path.exists(active):
        with open(active, "r", encoding="utf-8", errors="replace") as f:
            cur = f.read()
        check("配置块尚未重复添加", "mcp-vb6-mcp-sdk" not in cur, "已存在则跳过")
    else:
        check("配置块尚未重复添加", True, "profile 不存在则跳过")

    print("")
    print("=" * 52)
    print(f"结果: {PASS}/{PASS + len(FAIL)} 通过")
    for name, detail in FAIL:
        print(f"  FAIL  {name}: {detail}")
    print("=" * 52)
    print("接入方法: 把 [1/3] 输出的配置块粘贴到")
    print(f"  {active}")
    print("然后重启 DSH（或刷新 MCP 工具列表）即可使用 mcp__vb6-mcp-sdk__* 工具。")
    sys.exit(0 if not FAIL else 1)


if __name__ == "__main__":
    main()
