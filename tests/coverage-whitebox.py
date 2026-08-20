#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""coverage-whitebox.py - white-box (VB6 source procedure-level) test coverage.

Approach: parse every VB6 source file (GBK), extract procedure declarations,
build a static call graph, then compute the reachability closure from the
entry points the test suite actually exercises (Main + protocol dispatch +
all interface implementations -- black-box evidence in coverage.py shows every
registered tool/prompt/resource/template is invoked by the tests). Procedures
outside the closure are reported as uncovered (dead code or missing coverage).

Run: uv run python tests/coverage-whitebox.py
"""
import io
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
GROUP_LIB = "sdk/json"  # VBJSON third-party library, reported separately

# --- VB6 declarations / bodies ---
PROC_RE = re.compile(
    r"^\s*(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?"
    r"(Sub|Function|Property\s+(?:Get|Let|Set))\s+([A-Za-z_]\w*)",
    re.IGNORECASE,
)
END_RE = re.compile(r"^\s*End\s+(Sub|Function|Property)\b", re.IGNORECASE)
CALL_RE = re.compile(r"\b([A-Za-z_]\w*)\s*\(", re.IGNORECASE)
CALLK_RE = re.compile(r"\bCall\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
NEW_RE = re.compile(r"\bNew\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
PROP_RE = re.compile(r"\.([A-Za-z_]\w*)\b", re.IGNORECASE)
# VB6 statement-level calls without parentheses: `StdioInit`, `LogFile "x"`,
# `HttpClose conn`, `If x Then Foo`, `x = 1 : Foo`.
LINE_CALL_RE = re.compile(r"^\s*([A-Za-z_]\w*)\b", re.IGNORECASE)
THEN_CALL_RE = re.compile(r"\bThen\s+([A-Za-z_]\w*)\b", re.IGNORECASE)
STR_RE = re.compile(r'"[^"]*"')

KEYWORDS = set(
    "if then else elseif end while wend for next do loop select case with dim redim "
    "const type declare option exit goto on error resume sub function property static "
    "public private friend new is not and or xor eqv imp mod rem let set call byval "
    "byref optional paramarray as to step each in true false nothing null empty me "
    "ubound lbound erase array split join".split()
)
INTERFACE_PREFIX = ("ITool_", "IPrompt_", "IResource_", "ITemplate_")
INTERFACE_FILES = {"ITool.cls", "IPrompt.cls", "IResource.cls", "ITemplate.cls"}


def strip_comment(line):
    """Cut a VB6 ' comment, ignoring quotes inside string literals."""
    in_str = False
    for i, ch in enumerate(line):
        if ch == '"':
            in_str = not in_str
        elif ch == "'" and not in_str:
            return line[:i]
    return line


class Proc:
    __slots__ = ("file", "module", "name", "kind", "line", "body", "is_interface")

    def __init__(self, file, module, name, kind, line, body, is_interface):
        self.file = file
        self.module = module
        self.name = name
        self.kind = kind
        self.line = line
        self.body = body
        self.is_interface = is_interface


def collect_files():
    files = []
    for base_dir in (os.path.join(ROOT, "sdk"), os.path.join(ROOT, "tools")):
        for base, _dirs, fs in os.walk(base_dir):
            for fn in sorted(fs):
                if fn.lower().endswith((".bas", ".cls")):
                    files.append(os.path.join(base, fn))
    files.append(os.path.join(ROOT, "mcp_main.bas"))
    return files


def parse_procs():
    procs = []
    for path in collect_files():
        with io.open(path, "r", encoding="gbk", errors="replace") as f:
            lines = f.read().splitlines()
        module = os.path.splitext(os.path.basename(path))[0]
        is_iface = os.path.basename(path) in INTERFACE_FILES
        i, n = 0, len(lines)
        while i < n:
            m = PROC_RE.match(lines[i])
            if not m:
                i += 1
                continue
            kind, name = m.group(1).lower(), m.group(2)
            j = i + 1
            while j < n and not END_RE.match(lines[j]):
                j += 1
            body = "\n".join(strip_comment(lines[k]) for k in range(i + 1, j))
            procs.append(Proc(path, module, name, kind, i + 1, body, is_iface))
            i = j + 1
    return procs


def build_graph(procs):
    """Static call graph: proc id -> set of callee proc ids (name matched)."""
    name_map = {}
    alias_map = {}
    for p in procs:
        name_map.setdefault(p.name, []).append(p)
        for pref in INTERFACE_PREFIX:
            if p.name.startswith(pref):
                alias_map.setdefault(p.name[len(pref):], []).append(p)
    class_init = {p.module: p for p in procs if p.name.lower() == "class_initialize"}

    edges = {}
    for p in procs:
        targets = set()
        text = p.body
        names = (set(CALL_RE.findall(text)) | set(CALLK_RE.findall(text))
                 | set(PROP_RE.findall(text)))
        for line in text.splitlines():
            masked = STR_RE.sub('""', line)  # hide string literals before ':' split
            for seg in masked.split(":"):
                m = LINE_CALL_RE.match(seg)
                if m:
                    names.add(m.group(1))
                tm = THEN_CALL_RE.search(seg)
                if tm:
                    names.add(tm.group(1))
        for n in names:
            if n.lower() in KEYWORDS:
                continue
            for t in name_map.get(n, []) + alias_map.get(n, []):
                targets.add(id(t))
        for cls in NEW_RE.findall(text):
            ci = class_init.get(cls)
            if ci:
                targets.add(id(ci))
        edges[id(p)] = targets
    return edges


def reachable(procs, edges):
    """Closure from the entry points the test suite actually exercises."""
    roots = set()
    for p in procs:
        if p.is_interface:
            continue
        if p.module == "mcp_main" and p.name.lower() == "main":
            roots.add(id(p))
        if p.module == "McpServer" and p.name in (
            "RunStdio", "RunHttp", "HandleMessage", "HandleHttpConnection",
        ):
            roots.add(id(p))
        if p.name.startswith(INTERFACE_PREFIX):
            roots.add(id(p))
        # .bas module-level initialization runs on first module access.
        if p.name.lower() == "class_initialize" and p.file.endswith(".bas"):
            roots.add(id(p))
    seen = set(roots)
    queue = list(roots)
    while queue:
        cur = queue.pop()
        for t in edges.get(cur, ()):
            if t not in seen:
                seen.add(t)
                queue.append(t)
    return seen


def main():
    procs = parse_procs()
    edges = build_graph(procs)
    seen = reachable(procs, edges)

    stat_procs = [p for p in procs if not p.is_interface]
    by_file = {}
    for p in stat_procs:
        by_file.setdefault(p.file, []).append(p)

    uncovered = [p for p in stat_procs if id(p) not in seen]

    def is_lib(path):
        return os.path.relpath(path, ROOT).replace("\\", "/").startswith(GROUP_LIB)

    print("")
    print("=" * 58)
    print("  白盒覆盖率报告（VB6 源码过程级）")
    print("  入口根：Main + 协议分发 + 全部接口实现（测试实际执行路径）")
    print("=" * 58)

    core_cov, core_all = 0, 0
    lib_cov, lib_all = 0, 0
    for path in sorted(by_file):
        ps = sorted(by_file[path], key=lambda p: p.line)
        c = sum(1 for p in ps if id(p) in seen)
        rel = os.path.relpath(path, ROOT).replace("\\", "/")
        if is_lib(path):
            lib_cov += c
            lib_all += len(ps)
        else:
            core_cov += c
            core_all += len(ps)
        pct = c / len(ps) * 100 if ps else 100.0
        mark = "✓" if c == len(ps) else "✗"
        lib = "  [VBJSON 库]" if is_lib(path) else ""
        print(f"  {rel:<32} {c:>3}/{len(ps):<3} ({pct:5.1f}%) {mark}{lib}")
        miss = [p.name for p in ps if id(p) not in seen]
        if miss:
            print(f"       未覆盖: {', '.join(miss)}")

    print("-" * 58)
    core_pct = core_cov / core_all * 100 if core_all else 100.0
    lib_pct = lib_cov / lib_all * 100 if lib_all else 100.0
    total_pct = (core_cov + lib_cov) / (core_all + lib_all) * 100 if (core_all + lib_all) else 100.0
    print(f"  自研代码  {core_cov:>3}/{core_all:<3} ({core_pct:5.1f}%)")
    print(f"  VBJSON 库 {lib_cov:>3}/{lib_all:<3} ({lib_pct:5.1f}%)  （库内未使用功能不计入缺口）")
    print(f"  总计      {core_cov + lib_cov:>3}/{core_all + lib_all:<3} ({total_pct:5.1f}%)")

    if uncovered:
        print("-" * 58)
        print("  未覆盖过程（可能原因：死代码 / 缺少测试触达）：")
        for p in sorted(uncovered, key=lambda p: os.path.relpath(p.file, ROOT)):
            rel = os.path.relpath(p.file, ROOT).replace("\\", "/")
            tag = "  [VBJSON 库]" if is_lib(p.file) else ""
            print(f"    {rel}:{p.line}  {p.name}  ({p.kind}){tag}")
    print("=" * 58)


if __name__ == "__main__":
    main()
