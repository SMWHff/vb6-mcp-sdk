#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""coverage.py - compute test-case coverage (tools/prompts/resources/templates/methods).

Approach (black-box): query the running server for everything it exposes, then
scan the test sources for which of those are exercised, and report coverage.
Run: uv run --with mcp python tests/coverage.py
"""
import asyncio
import os
import re
import subprocess
import sys

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
EXE = os.path.join(ROOT, "vb6-mcp-sdk.exe")
TESTS_DIR = os.path.join(ROOT, "tests")
TEST_FILES = ["test-suite.py", "http-test.py", "client-sdk.py", "client-sdk-http.py"]

# \u540d\u79f0\u6807\u7b7e
LBL = {
    "tools": "\u5de5\u5177", "prompts": "\u63d0\u793a\u8bcd", "resources": "\u8d44\u6e90",
    "templates": "\u6a21\u677f", "methods": "\u534f\u8bae\u65b9\u6cd5", "total": "\u603b\u4f53",
}


async def fetch_exposed():
    """Query the live server for everything it exposes."""
    server = StdioServerParameters(command=EXE, args=[], env=None)
    async with stdio_client(server) as (read, write):
        async with ClientSession(read, write) as session:
            await session.initialize()
            tools = [t.name for t in (await session.list_tools()).tools]
            prompts = [p.name for p in (await session.list_prompts()).prompts]
            resources = [r.uri for r in (await session.list_resources()).resources]
            templates = [t.uri_template for t in (await session.list_resource_templates()).resource_templates]
            return tools, prompts, resources, templates


def scan_sources():
    """Scan test sources for exercised tools/prompts/resources/templates/methods."""
    text = ""
    for fn in TEST_FILES:
        p = os.path.join(TESTS_DIR, fn)
        if os.path.exists(p):
            with open(p, "r", encoding="utf-8", errors="replace") as f:
                text += f.read() + "\n"

    tools = set(re.findall(r'call_tool\("([^"]+)"', text))
    prompts = set(re.findall(r'get_prompt\("([^"]+)"', text))
    resource_uris = set(re.findall(r'read_resource\("([^"]+)"', text))

    methods = set()
    for pat, name in [
        (r'session\.initialize\(', "initialize"),
        (r'"method":"ping"|send_ping\(', "ping"),
        (r'"method":"tools/list"|list_tools\(', "tools/list"),
        (r'call_tool\(', "tools/call"),
        (r'"method":"prompts/list"|list_prompts\(', "prompts/list"),
        (r'get_prompt\(', "prompts/get"),
        (r'"method":"resources/list"|list_resources\(', "resources/list"),
        (r'read_resource\(', "resources/read"),
        (r'list_resource_templates\(', "resources/templates/list"),
        (r'notifications/', "notifications"),
    ]:
        if re.search(pat, text):
            methods.add(name)
    return tools, prompts, resource_uris, methods


def uri_matches_template(uri, tpl):
    """demo://greet/{name} matches demo://greet/xxx (static prefix before '{')."""
    brace = tpl.find("{")
    static = tpl[:brace] if brace >= 0 else tpl
    return uri.startswith(static) if static else True


def main():
    try:
        tools, prompts, resources, templates = asyncio.run(fetch_exposed())
    except Exception as e:
        print("ERROR: \u65e0\u6cd5\u542f\u52a8 server \u83b7\u53d6\u80fd\u529b\u6e05\u5355:", e)
        print("      \u8bf7\u5148\u7f16\u8bd1 vb6-mcp-sdk.exe")
        sys.exit(1)

    c_tools, c_prompts, c_uris, c_methods = scan_sources()

    exposed_methods = [
        "initialize", "ping", "tools/list", "tools/call",
        "prompts/list", "prompts/get", "resources/list", "resources/read",
        "resources/templates/list", "notifications",
    ]

    t_miss = [t for t in tools if t not in c_tools]
    p_miss = [p for p in prompts if p not in c_prompts]
    r_miss = [r for r in resources if r not in c_uris]
    tpl_miss = [t for t in templates if not any(uri_matches_template(u, t) for u in c_uris)]
    m_miss = [m for m in exposed_methods if m not in c_methods]

    rows = [
        ("tools", tools, t_miss),
        ("prompts", prompts, p_miss),
        ("resources", resources, r_miss),
        ("templates", templates, tpl_miss),
        ("methods", exposed_methods, m_miss),
    ]

    print("")
    print("=" * 52)
    print("  \u6d4b\u8bd5\u7528\u4f8b\u8986\u76d6\u7387\u62a5\u544a")
    print("=" * 52)
    total_cov = 0
    total_all = 0
    for key, all_items, miss in rows:
        covered = len(all_items) - len(miss)
        total_all += len(all_items)
        total_cov += covered
        pct = (covered / len(all_items) * 100) if all_items else 100.0
        mark = "✓" if not miss else "✗"
        print(f"  {LBL[key]:<4} {covered:>3}/{len(all_items):<3} ({pct:5.1f}%)  {mark}")
        if miss:
            print(f"         \u672a\u8986\u76d6: {', '.join(miss)}")
    total_pct = (total_cov / total_all * 100) if total_all else 100.0
    print("-" * 52)
    print(f"  {LBL['total']:<4} {total_cov:>3}/{total_all:<3} ({total_pct:5.1f}%)")
    print("=" * 52)


if __name__ == "__main__":
    main()
