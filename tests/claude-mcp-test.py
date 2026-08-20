#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""claude-mcp-test.py - Claude Code (Anthropic CLI) MCP compatibility test.

Verifies that vb6-mcp-sdk can be registered with Claude Code, connects
(initialize + tools/list succeed), and is actually callable end-to-end
(Claude Code -> stdio MCP -> VB6 server -> tool -> result).

Prereqs: `claude` CLI installed (npm i -g @anthropic-ai/claude-code)
         and authenticated (an API key or `claude` login), because the
         headless tool-call step needs model access.

Run: uv run python tests/claude-mcp-test.py
"""
import os
import re
import shutil
import subprocess
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXE = os.path.join(BASE_DIR, "vb6-mcp-sdk.exe")
SERVER_NAME = "vb6-mcp-sdk"
HAIKU = "claude-3-5-haiku-latest"

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


def find_claude():
    for name in ("claude.cmd", "claude", "claude.exe"):
        p = shutil.which(name)
        if p:
            return p
    return None


def run(args, timeout=90):
    """Run a claude CLI command, return (returncode, stdout)."""
    p = subprocess.run(args, capture_output=True, text=True, timeout=timeout,
                       encoding="utf-8", errors="replace")
    return p.returncode, p.stdout


def main():
    claude = find_claude()
    if not claude:
        print("claude CLI 未安装，跳过（npm i -g @anthropic-ai/claude-code 后重试）")
        sys.exit(2)

    print("=== Claude Code (MCP 客户端) 兼容性测试 ===\n")
    print(f"  claude: {claude}")
    rc, ver = run([claude, "--version"], timeout=30)
    print(f"  版本: {ver.strip().splitlines()[0] if ver.strip() else rc}\n")

    registered = False
    try:
        # 0) 清理可能的历史注册（幂等，避免同名重复注册失败）
        run([claude, "mcp", "remove", SERVER_NAME], timeout=30)

        # 1) 注册服务器（local scope：写入 ~/.claude.json，仅本机）
        rc, out = run([claude, "mcp", "add", SERVER_NAME, "--scope", "local", "--", EXE], timeout=60)
        registered = rc == 0
        check("claude mcp add 注册成功", registered, f"rc={rc} {out[:120]}")

        # 2) 连接检查（Claude Code 实际握手 initialize + tools/list）
        rc, out = run([claude, "mcp", "list"], timeout=90)
        check("claude mcp list 显示服务器", SERVER_NAME in out, out[:200])
        check("服务器连接成功 (Connected)",
              re.search(r"vb6-mcp-sdk.*Connected", out, re.I) is not None, out[:200])

        # 3) headless 端到端工具调用（Claude Code -> MCP -> add 工具）
        prompt = "调用 add 工具计算 2+3，直接输出结果数字"
        try:
            rc, out = run([claude, "-p", prompt,
                           "--allowedTools", f"mcp__{SERVER_NAME}__add",
                           "--model", HAIKU], timeout=180)
            result = out.strip()
            check("headless 调用 add(2,3) = 5",
                  rc == 0 and "5" in result and "错误" not in result[:20], f"rc={rc} out={result[:120]}")
        except subprocess.TimeoutExpired:
            check("headless 调用 add(2,3) = 5", False, "超时（模型推理过慢）")
        except Exception as e:
            check("headless 调用 add(2,3) = 5", False, str(e)[:120])
    finally:
        # 4) 清理：移除注册的服务器
        if registered:
            try:
                run([claude, "mcp", "remove", SERVER_NAME], timeout=30)
            except Exception:
                pass

    print("")
    print("=" * 52)
    ok = PASS == len(FAIL) + PASS and len(FAIL) == 0
    print(f"结果: {PASS}/{PASS + len(FAIL)} 通过")
    for name, detail in FAIL:
        print(f"  FAIL  {name}: {detail}")
    print("=" * 52)
    print("说明: 注册/连接/清理由脚本自动完成；headless 调用需要模型额度。")
    sys.exit(0 if not FAIL else 1)


if __name__ == "__main__":
    main()
