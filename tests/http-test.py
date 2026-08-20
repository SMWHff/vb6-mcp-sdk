# http-test.py —— vb6-mcp-sdk Streamable HTTP 传输层专项测试
# 用法：先启动 .\vb6-mcp-sdk.exe /http:9002
#       uv run python scripts/http-test.py [URL]
# 覆盖：POST /mcp、OPTIONS 预检+CORS、GET 健康检查、404、202 通知、
#       非法 JSON、Content-Type/Length 头、中文 UTF-8
import json
import sys
import urllib.error
import urllib.request

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

URL = sys.argv[1] if len(sys.argv) > 1 else "http://localhost:9002/mcp"

PASS = 0
FAIL = []


def post(path, body, headers=None):
    req = urllib.request.Request(
        URL.rsplit("/mcp", 1)[0] + path if path != "/mcp" else URL,
        data=body.encode("utf-8") if body is not None else None,
        headers={"Content-Type": "application/json", **(headers or {})},
        method="POST" if body is not None else "GET",
    )
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        return resp.status, dict(resp.headers), resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        return e.code, dict(e.headers), e.read().decode("utf-8", errors="replace")


def get(path):
    return post(path, None)


def check(name, cond, detail=""):
    global PASS
    if cond:
        PASS += 1
        print(f"  PASS  {name}")
    else:
        FAIL.append((name, detail))
        print(f"  FAIL  {name}  -> {detail}")


def main():
    print("=== Streamable HTTP 传输层测试 ===\n")

    # 1. POST /mcp initialize
    status, headers, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {}}))
    check("POST /mcp initialize -> 200", status == 200, status)
    check("initialize 返回协议版本", '"protocolVersion":"2024-11-05"' in body, body[:80])
    check("Content-Type: application/json", "application/json" in headers.get("Content-Type", ""), headers.get("Content-Type"))

    # 2. 工具调用（中文回显）
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 2, "method": "tools/call",
                                               "params": {"name": "echo", "arguments": {"text": "你好 HTTP"}}}))
    check("tools/call echo 中文", '"你好 HTTP"' in body, body[:80])

    # 3. 通知 -> 202
    status, headers, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "method": "notifications/initialized"}))
    check("通知 -> 202 Accepted", status == 202, status)

    # 4. 非法 JSON -> 200 + -32700
    status, _, body = post("/mcp", "not json at all")
    check("非法 JSON -> 200 + -32700", status == 200 and "-32700" in body, f"{status} {body[:60]}")

    # 5. OPTIONS 预检 -> 204 + CORS
    req = urllib.request.Request(URL, method="OPTIONS")
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        h = dict(resp.headers)
        check("OPTIONS -> 204", resp.status == 204, resp.status)
        check("CORS Allow-Origin: *", h.get("Access-Control-Allow-Origin") == "*", h.get("Access-Control-Allow-Origin"))
        check("CORS Allow-Methods", "POST" in h.get("Access-Control-Allow-Methods", ""), h.get("Access-Control-Allow-Methods"))
    except Exception as e:
        check("OPTIONS -> 204", False, str(e))

    # 6. GET / 健康检查
    status, _, body = get("/")
    check("GET / 健康检查 -> 200", status == 200 and "streamable-http" in body, f"{status} {body[:60]}")

    # 6b. GET /mcp -> SSE 流（2025-06-18 Streamable HTTP 要求）
    status, headers, body = get("/mcp")
    check("GET /mcp -> 200", status == 200, status)
    check("GET /mcp Content-Type: text/event-stream",
          "text/event-stream" in headers.get("Content-Type", ""), headers.get("Content-Type"))
    check("GET /mcp 含 endpoint 事件", "event: endpoint" in body and "data: /mcp" in body, body[:60])

    # 7. 未知路径 -> 404
    status, _, _ = post("/wrong", json.dumps({"jsonrpc": "2.0", "id": 1, "method": "ping"}))
    check("POST 未知路径 -> 404", status == 404, status)

    # 8. 未知工具 -> JSON-RPC error
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 3, "method": "tools/call",
                                               "params": {"name": "nope"}}))
    check("未知工具 -> -32602", status == 200 and "-32602" in body, body[:60])

    # 9. 缺参 -> 框架拦截
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 4, "method": "tools/call",
                                               "params": {"name": "add", "arguments": {"a": 1}}}))
    check("缺参拦截 -> 缺少必需参数", "缺少必需参数" in body, body[:60])

    # 10. 会话管理（Mcp-Session-Id）
    status, headers, _ = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 10, "method": "ping"}))
    sess = headers.get("Mcp-Session-Id", "")
    check("首次请求返回 Mcp-Session-Id", len(sess) > 0, sess[:30])
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 11, "method": "ping"}),
                           {"Mcp-Session-Id": sess})
    check("带会话请求正常", status == 200 and '"result"' in body, f"{status} {body[:40]}")
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 14, "method": "tools/list"}),
                           {"Mcp-Session-Id": sess})
    check("同会话 tools/list 复用", status == 200 and '"tools"' in body, f"{status} {body[:40]}")
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 15, "method": "tools/call",
                                               "params": {"name": "get_time"}}),
                           {"Mcp-Session-Id": sess})
    check("同会话 tools/call 复用", status == 200 and '"text"' in body, f"{status} {body[:40]}")
    status, _, _ = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 12, "method": "ping"}),
                        {"Mcp-Session-Id": "vb6-invalid-session"})
    check("无效会话 -> 404", status == 404, status)
    req = urllib.request.Request(URL, method="DELETE", headers={"Mcp-Session-Id": sess})
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        check("DELETE 关会话 -> 204", resp.status == 204, resp.status)
    except urllib.error.HTTPError as e:
        check("DELETE 关会话 -> 204", e.code == 204, e.code)

    # 11. HTTP 边界与健壮性
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 13, "method": "tools/call",
                                               "params": {"name": "echo", "arguments": {"text": "长" * 5000}}}))
    check("HTTP 10KB 大 body", status == 200 and ("长" * 5000) in body, status)
    status, _, body = post("/mcp", "")
    check("无 body -> -32700", status == 200 and "-32700" in body, f"{status} {body[:40]}")
    status, _, body = post("/mcp?x=1", json.dumps({"jsonrpc": "2.0", "id": 14, "method": "ping"}))
    check("带 query 正常", status == 200 and '"result"' in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 15, "method": "ping"}),
                           {"Content-Type": "text/plain"})
    check("非 JSON Content-Type 宽容", status == 200 and '"result"' in body, status)
    req = urllib.request.Request(URL.rsplit("/mcp", 1)[0] + "/", method="HEAD")
    try:
        resp = urllib.request.urlopen(req, timeout=10)
        check("HEAD -> 404", resp.status == 404, resp.status)
    except urllib.error.HTTPError as e:
        check("HEAD -> 404", e.code == 404, e.code)

    # 12. 三大能力全链路（HTTP 层）
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 20, "method": "prompts/list"}))
    check("HTTP prompts/list", status == 200 and "code_review" in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 21, "method": "prompts/get",
                                               "params": {"name": "code_review", "arguments": {"language": "VB6", "code": "x"}}}))
    check("HTTP prompts/get", status == 200 and "VB6" in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 22, "method": "resources/list"}))
    check("HTTP resources/list", status == 200 and "demo://server/info" in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 23, "method": "resources/read",
                                               "params": {"uri": "demo://server/info"}}))
    check("HTTP resources/read", status == 200 and "vb6-mcp-sdk-demo" in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 24, "method": "resources/templates/list"}))
    check("HTTP templates/list", status == 200 and "demo://greet/{name}" in body, status)
    status, _, body = post("/mcp", json.dumps({"jsonrpc": "2.0", "id": 25, "method": "resources/read",
                                               "params": {"uri": "demo://greet/测试"}}))
    check("HTTP 模板 read", status == 200 and "你好" in body, status)

    print("\n" + "=" * 56)
    print(f"结果: {PASS}/{PASS + len(FAIL)} 通过")
    for name, detail in FAIL:
        print(f"  FAIL  {name}: {detail}")
    sys.exit(0 if not FAIL else 1)


main()
