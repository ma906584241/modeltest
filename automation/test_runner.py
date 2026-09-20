"""Local fake-server tests. No internal model endpoint is called."""
import http.server
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest

import run_tests as runner


class Handler(http.server.BaseHTTPRequestHandler):
    protocol_version = "HTTP/1.1"
    received = []

    def log_message(self, *args):
        pass

    def do_POST(self):
        payload = json.loads(self.rfile.read(int(self.headers["Content-Length"])))
        self.received.append(payload)
        if not payload["stream"]:
            value = {"model":"mock", "choices":[{"index":0,"message":{"role":"assistant","content":"0 测试，1 测试"},"finish_reason":"length" if self.path=="/json_length" else "stop"}],"usage":{"completion_tokens":512},"prefill_time":220,"decode_time_arr":[60,61]}
            body = json.dumps(value,ensure_ascii=False).encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type","application/json")
            self.send_header("Content-Length",str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if self.path == "/error":
            self.send_response(404)
            self.send_header("Content-Length", "13")
            self.end_headers()
            self.wfile.write(b"unknown model")
            return
        self.send_response(200)
        self.send_header("Content-Type", "text/event-stream")
        self.send_header("Connection", "close")
        self.end_headers()
        if self.path == "/timeout":
            time.sleep(.3)
            return
        def emit(value):
            self.wfile.write(("data: " + json.dumps(value,ensure_ascii=False)+"\r\n\r\n").encode())
            self.wfile.flush()
        emit({"model":"mock", "choices":[{"delta":{"role":"assistant"}}]})
        emit({"choices":[{"delta":{"reasoning_content":"思考"}}]})
        if self.path == "/empty":
            chunks = []
        elif self.path == "/think":
            chunks = ["<thi", "nk>hidden", "</think>", "0 测试，1 测试"]
        else:
            chunks = ["0 测试，", "1 测试"]
        for chunk in chunks:
            emit({"choices":[{"delta":{"content":chunk}}]})
        if self.path != "/broken":
            emit({"choices":[{"delta":{}, "finish_reason":"length" if self.path=="/length" else "stop"}]})
            emit({"usage":{"completion_tokens":10}, "choices":[]})
            self.wfile.write(b"data: [DONE]\n\n")
            self.wfile.flush()
        self.close_connection = True


class RunnerTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.server = http.server.ThreadingHTTPServer(("127.0.0.1",0), Handler)
        cls.thread = threading.Thread(target=cls.server.serve_forever,daemon=True)
        cls.thread.start()

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def call(self, path):
        cfg = json.loads((runner.HERE / "config.json").read_text(encoding="utf-8"))
        cfg["stream"] = not path.startswith("json")
        cfg["endpoint"] = f"http://127.0.0.1:{self.server.server_port}/{path}"
        case = dict(id="test",group="test",source="mock",prompt="test",rule="count",end=1)
        with tempfile.TemporaryDirectory() as folder:
            result = runner.request_model(cfg,"mock",case,Path(folder),"test",timeout=.1 if path=="timeout" else 3)
            runner.report(Path(folder),[result],"mock",[])
            self.assertTrue((Path(folder)/"report.html").exists())
        return result

    def test_stream_and_usage(self):
        r = self.call("ok")
        self.assertEqual((r["status"],r["quality"]),("OK","PASS"))
        self.assertEqual(r["completion_tokens"],10)
        self.assertEqual(r["reasoning_chars"],2)
        self.assertGreaterEqual(r["first_answer_s"],r["first_output_s"])
        self.assertEqual(len(Handler.received[-1]["messages"]),1)

    def test_nonstream(self):
        r = self.call("json")
        self.assertEqual(r["status"],"OK")
        self.assertIsNone(r["first_output_s"])
        self.assertIsNone(r["first_answer_s"])
        self.assertEqual(r["server_metrics"]["prefill_time"],220)
        self.assertEqual(Handler.received[-1]["temperature"],0.6)
        self.assertEqual(Handler.received[-1]["max_tokens"],4096)

    def test_nonstream_length(self):
        r = self.call("json_length")
        self.assertEqual(r["status"],"TRUNCATED")
        self.assertEqual(r["completion_tokens"],512)
        self.assertEqual(r["requested_max_tokens"],4096)

    def test_split_think_tag(self):
        r = self.call("think")
        self.assertEqual(r["quality"],"PASS")
        self.assertEqual(r["answer_chars"],len("0 测试，1 测试"))

    def test_truncation(self):
        self.assertEqual(self.call("length")["status"],"TRUNCATED")

    def test_broken_stream(self):
        self.assertEqual(self.call("broken")["status"],"ERROR")

    def test_http_error(self):
        r = self.call("error")
        self.assertEqual(r["http_status"],404)
        self.assertEqual(r["status"],"ERROR")

    def test_deadline(self):
        self.assertEqual(self.call("timeout")["status"],"TIMEOUT")

    def test_empty_answer(self):
        self.assertEqual(self.call("empty")["status"],"ERROR")

    def test_count_missing_duplicate_out_of_order(self):
        case = dict(rule="count",end=2)
        for text in ("0 测试，2 测试", "0 测试，1 测试，1 测试，2 测试", "2 测试，1 测试，0 测试"):
            self.assertEqual(runner.evaluate(case,text,"stop")[0],"FAIL")


if __name__ == "__main__":
    unittest.main(verbosity=2)
