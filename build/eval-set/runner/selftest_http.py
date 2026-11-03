# -*- coding: utf-8 -*-
"""
自检：起一个本地假接口，验证 run_eval.py 的 HTTP 通路（请求体、响应解析、5xx 处理）。

用途只有一个——在没有被测服务时确认「脚本这一侧」没坏。
假接口对一半请求返 200、一半返 500，因此跑到「不通过（技术失败）」是正确的。

  python selftest_http.py
"""
import json
import subprocess
import sys
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

PORT = 8899
CALLS = {"n": 0}


class Handler(BaseHTTPRequestHandler):
    def do_POST(self):
        n = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(n).decode("utf-8"))
        CALLS["n"] += 1
        if CALLS["n"] % 2 == 0:                       # 故意制造 5xx，验证硬失败判定
            self.send_response(500)
            self.end_headers()
            self.wfile.write(b'{"error":"boom"}')
            return
        q = body.get("question", "")
        ans = f"（假接口）已收到：{q[:20]}；ACCTREC.cpy ACCT-BAL S9(11)V99"
        payload = {"answer": ans, "sources": [{"loc": "ACCTREC.cpy:214-216"}],
                   "latency_ms": 900, "echo_filters": body.get("filters")}
        raw = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(raw)))
        self.end_headers()
        self.wfile.write(raw)

    def log_message(self, *a):                          # 静音
        pass


def main():
    here = Path(__file__).parent
    srv = HTTPServer(("127.0.0.1", PORT), Handler)
    threading.Thread(target=srv.serve_forever, daemon=True).start()
    try:
        cmd = [sys.executable, str(here / "run_eval.py"),
               "--input", str(here / "eval_set.jsonl"),
               "--endpoint", f"http://127.0.0.1:{PORT}/qa",
               "--out", str(here / "_selftest_results.csv"),
               "--include-sample", "--limit", "4"]
        print("[selftest] " + " ".join(cmd[1:]))
        r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8")
        print(r.stdout.strip())
        if r.returncode != 0:
            print(r.stderr)
            return 1
        out = here / "_selftest_results.csv"
        lines = out.read_text(encoding="utf-8-sig").splitlines()
        assert len(lines) == 5, f"应有表头+4行，实际 {len(lines)} 行"
        assert "不通过" in "\n".join(lines), "5xx 未被判为不通过，硬失败逻辑没生效"
        print(f"[selftest] 通过：4 次调用成功落库，500 响应被判为「不通过」")
        out.unlink()
        return 0
    finally:
        srv.shutdown()


if __name__ == "__main__":
    sys.exit(main())
