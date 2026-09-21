"""Local fake-server tests. No internal model endpoint is called."""
import http.server
import json
import contextlib
import io
import sys
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch
import quality_policy as policy

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
        if self.path in ("/error", "/server_error"):
            self.send_response(503 if self.path == "/server_error" else 404)
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

    def test_repeat_whitespace_and_prefix(self):
        case = dict(rule="repeat")
        for tail in ("", "测", "测试", "测试输"):
            self.assertEqual(runner.evaluate(case, "测试输出\n\t 测试输出" + tail, "length")[0], "PASS")
        for text in ("测试输出其他", "测试输出测输", "测试输出测试"):
            self.assertEqual(runner.evaluate(case, text, "stop")[0], "FAIL")

    def test_strict_full_sequences(self):
        for end in (1000, 5000):
            case = dict(rule="count", end=end)
            answer = "\n".join(f"{i} 测试" for i in range(end + 1))
            self.assertEqual(runner.evaluate(case, answer, "stop")[0], "PASS")
            for wrong in (answer + "\n解释", answer.replace("5 测试\n", "", 1),
                          answer.replace("5 测试\n", "4 测试\n", 1), answer + "...", answer[:-8]):
                self.assertEqual(runner.evaluate(case, wrong, "stop")[0], "FAIL")
            self.assertEqual(runner.evaluate(case, answer, "length")[0], "FAIL")

    def test_translation_structure_language_and_ending(self):
        case = dict(rule="translate", source_text="第一章 总则\n1.1 内容\n第二章 实施\n2.1 内容")
        good = "Chapter 1 General\n1.1 A complete sentence.\nChapter 2 Implementation\n2.1 Another complete sentence."
        self.assertEqual(runner.evaluate(case, good, "stop")[0], "PASS")
        for text in (good.replace("Chapter 2 Implementation\n", ""), good + "中文。", good[:-1], good.replace("2.1", "2.2")):
            self.assertEqual(runner.evaluate(case, text, "stop")[0], "FAIL")
        self.assertEqual(runner.evaluate(dict(rule="translate", source_text="没有编号的原文。"), "A sentence.", "stop")[0], "REVIEW")

    def test_input_limit_requires_exact_acknowledgement(self):
        case = dict(rule="limit_input")
        self.assertEqual(runner.evaluate(case, "输入已接收。", "stop")[0], "PASS")
        self.assertEqual(runner.evaluate(case, "输入已接收", "stop")[0], "PASS")
        self.assertEqual(runner.evaluate(case, "输入已接收。补充内容", "stop")[0], "FAIL")
        self.assertEqual(runner.evaluate(case, "输入已接收。", "length")[0], "FAIL")

    def config(self):
        cfg = json.loads((runner.HERE / "config.json").read_text(encoding="utf-8"))
        cfg.update(network_retries=0, retry_backoff_seconds=0)
        cfg["endpoint"] = f"http://127.0.0.1:{self.server.server_port}/json_length"
        return cfg

    def test_preflight_blocks_before_network(self):
        cfg = self.config()
        case = dict(id="large", group="test", source="mock", prompt="x", max_tokens=65536)
        before = len(Handler.received)
        with tempfile.TemporaryDirectory() as directory:
            result = runner.request_model(cfg, "mock", case, Path(directory), "large")
        self.assertEqual(len(Handler.received), before)
        self.assertEqual(result["failure_type"], "CONFIG_LIMIT")
        self.assertEqual(result["attempt_count"], 0)
        self.assertEqual(result["final_quality"], "FAIL")
        payload = dict(max_tokens=32768, messages=[dict(role="user", content="测" * 11000)])
        cfg["strict_estimated_preflight"] = True
        with self.assertRaisesRegex(policy.PreflightError, "max_seq_len"):
            policy.preflight(cfg, "mock", dict(id="context"), payload)

    def test_count_truncation_classified_and_bounded(self):
        case = dict(id="count_1000", group="test", source="mock", prompt="x", rule="count", end=1000, max_tokens=8192)
        with tempfile.TemporaryDirectory() as directory:
            result = runner.request_model(self.config(), "mock", case, Path(directory), "count")
            self.assertTrue((Path(directory) / "count_a1.request.json").exists())
            self.assertTrue((Path(directory) / "count_a2.request.json").exists())
        self.assertEqual(result["attempt_count"], 2)
        self.assertEqual(result["failure_type"], "CONFIG_LIMIT")
        self.assertEqual(result["raw_quality"], "FAIL")
        self.assertEqual(result["final_quality"], "FAIL")
        self.assertEqual([a["requested_max_tokens"] for a in result["attempts"]], [8192, 16384])

    def test_real_tokenizer_budget_contract(self):
        class Tokenizer:
            def encode(self, text, **kwargs):
                self.answer = text
                return [0] * 36000
            def apply_chat_template(self, messages, **kwargs):
                return [0] * 100
        tokenizer = Tokenizer()
        cfg = self.config()
        cfg["tokenizer_paths"] = {"mock": "local-test-tokenizer"}
        cfg["server_limits"]["max_iter_times"] = 49152
        cfg["supports_min_new_tokens"] = True
        payload = dict(max_tokens=49152, messages=[dict(role="user", content="x")])
        with patch.object(policy, "load_tokenizer", return_value=tokenizer):
            budget = policy.preflight(cfg, "mock", dict(id="count_5000"), payload)
        self.assertEqual(len(tokenizer.answer.splitlines()), 5001)
        self.assertEqual(payload["max_tokens"], 39600)
        self.assertEqual(payload["min_new_tokens"], 36000)
        self.assertEqual(budget["prompt_tokens_budget"], 100)

    def test_count5000_missing_tokenizer_blocks(self):
        cfg = self.config()
        cfg["require_count_tokenizer"] = True
        with self.assertRaises(policy.PreflightError):
            policy.preflight(cfg, "mock", dict(id="count_5000"), dict(max_tokens=49152, messages=[]))

    def test_estimate_warns_and_request_reaches_server(self):
        cfg = self.config()
        case = dict(id="input_8192_r1", group="test", source="mock", prompt="测" * 16400, max_tokens=512)
        before = len(Handler.received)
        with tempfile.TemporaryDirectory() as directory:
            result = runner.request_model(cfg, "mock", case, Path(directory), "estimated")
        self.assertEqual(len(Handler.received), before + 1)
        self.assertTrue(result["preflight_warnings"])
        self.assertNotEqual(result["request_phase"], "preflight")

    def test_exact_input_limit_still_blocks(self):
        class Tokenizer:
            def apply_chat_template(self, *args, **kwargs):
                return [0] * 40000
        cfg = self.config()
        cfg["tokenizer_paths"] = {"mock": "fixture"}
        with patch.object(policy, "load_tokenizer", return_value=Tokenizer()):
            with self.assertRaisesRegex(policy.PreflightError, "max_input_token_len"):
                policy.preflight(cfg, "mock", {}, dict(max_tokens=512, messages=[]))

    def test_count5000_fallback_budget(self):
        payload = dict(max_tokens=49152, messages=[])
        budget = policy.preflight(self.config(), "mock", dict(id="count_5000"), payload)
        self.assertEqual(payload["max_tokens"], 32768)
        self.assertIsNone(budget["standard_answer_tokens"])
        self.assertTrue(budget["preflight_warnings"])
        self.assertNotIn("min_new_tokens", payload)

    def test_upload_document_discovery_and_explicit_path(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            upload = root / "upload"
            upload.mkdir()
            (upload / "制度.doc").write_bytes(b"legacy")
            (upload / "制度.txt").write_text("测试制度正文", encoding="utf-8")
            with patch.object(runner, "ROOT", root):
                self.assertEqual(runner.resolve_data_dir({}), upload.resolve())
                cases, sources = runner.build_cases(self.config())
                self.assertEqual(len(sources), 1)
                self.assertEqual(len([c for c in cases if c["id"].startswith("doc_")]), 4)
                self.assertEqual(runner.resolve_data_dir({"data_dir": "upload"}), upload.resolve())
                with self.assertRaises(ValueError):
                    runner.resolve_data_dir({"data_dir": "missing"})

    def test_input_summary_uses_measured_tokens_and_quality(self):
        rows = [dict(case_id=f"input_8192_r{i}", status="OK", quality="PASS", usage=dict(prompt_tokens=n))
                for i, n in enumerate((7300, 7310, 7320), 1)]
        self.assertEqual(runner.limit_summary(rows, [], {})["max_stable_input_tokens"], 7300)
        rows[0]["quality"] = "FAIL"
        self.assertIsNone(runner.limit_summary(rows, [], {})["max_stable_input_tokens"])

    def test_historical_full_matrix_contract(self):
        cfg = self.config()
        cfg["data_dir"] = str(runner.ROOT)
        cases, sources = runner.build_cases(cfg)
        capacity = runner.build_capacity_cases(cfg)
        limits = runner.build_limit_cases(cfg)
        self.assertEqual(len(sources), 7)
        self.assertEqual(len(cases), 45)
        self.assertEqual(len(capacity), 9)
        self.assertEqual(len(limits), 24)
        self.assertEqual(len({case["id"] for case in cases + capacity + limits}), 78)
        stress = 4 * sum(cfg["concurrency_levels"]) * cfg["stress_rounds"]
        concurrency = sum(cfg["limit_concurrency_levels"]) * cfg["limit_repeats"]
        self.assertEqual(1 + len(cases + capacity + limits) + stress + concurrency, 91)

    def test_missing_documents_stops_full_run(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(runner, "ROOT", Path(directory)):
            with patch.object(sys, "argv", ["run_tests.py", "--mode", "full", "--dry-run"]):
                with contextlib.redirect_stderr(io.StringIO()), contextlib.redirect_stdout(io.StringIO()):
                    with self.assertRaises(SystemExit) as raised:
                        runner.main()
            self.assertEqual(raised.exception.code, 2)
            self.assertFalse((Path(directory) / "模型测评结果").exists())

    def test_full_dry_run_loads_explicit_documents(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(runner, "ROOT", Path(directory)):
            args = ["run_tests.py", "--mode", "full", "--dry-run", "--data-dir", str(runner.HERE.parent)]
            with patch.object(sys, "argv", args), contextlib.redirect_stdout(io.StringIO()):
                self.assertEqual(runner.main(), 0)
            output = Path(directory) / "模型测评结果"
            run = next(p for p in output.iterdir() if p.is_dir())
            sources = json.loads((run / "sources.json").read_text(encoding="utf-8"))
            self.assertTrue(sources)
            cases = json.loads((run / "cases.json").read_text(encoding="utf-8"))
            self.assertTrue(any(c["id"].startswith("input_") for c in cases))
            self.assertTrue(any(c["id"].startswith("doc_") for c in cases))
            page = (run / "report.html").read_text(encoding="utf-8")
            self.assertNotIn("????", page)
            self.assertNotIn("<!--BOUNDARY_NOTE-->", page)
            self.assertNotIn("<!--RUN_CONFIG-->", page)
            self.assertIn("<h2>失败原因与恢复记录</h2>", page)

    def test_count5000_whole_request_retry_only(self):
        cfg = self.config()
        cfg["endpoint"] = cfg["endpoint"].replace("json_length", "ok")
        cfg["server_limits"]["max_iter_times"] = 49152
        case = dict(id="count_5000", group="test", source="mock", prompt="all 5001 lines", rule="count", end=5000, stream=True, max_tokens=49152)
        before = len(Handler.received)
        budget = dict(prompt_tokens_budget=100, token_budget_method="test_fixture", standard_answer_tokens=36000, safety_margin=1024)
        with tempfile.TemporaryDirectory() as directory, patch.object(runner, "preflight", return_value=budget):
            result = runner.request_model(cfg, "mock", case, Path(directory), "count")
        self.assertEqual(len(Handler.received) - before, 2)
        self.assertEqual(result["final_quality"], "FAIL")
        for payload in Handler.received[before:]:
            self.assertTrue(payload["stream"])
            self.assertEqual(payload["messages"], [dict(role="user", content=case["prompt"])])

    def test_speech_validation_and_recovery(self):
        # Unique sentences avoid fabricating a pass through repeated padding.
        def section(index):
            sentences = ["这是一段正式发言正文" + chr(0x4e00 + index * 100 + j) * 2 + "我们应当认真落实工作要求并积极参与具体行动。" for j in range(16)]
            return f"第{index}章 标题\n" + "".join(sentences)
        complete = "\n".join(section(i) for i in range(1, 9))
        case = dict(id="speech", group="test", source="mock", prompt="speech", rule="speech", minimum=4000)
        self.assertEqual(runner.evaluate(case, complete, "stop")[0], "PASS")
        bold = "\n".join("**" + line + "**" if line.startswith("第") else line for line in complete.splitlines())
        self.assertEqual(runner.evaluate(case, bold, "stop")[0], "PASS")
        self.assertEqual(runner.evaluate(case, complete.replace("第8章", "第7章"), "stop")[0], "FAIL")
        self.assertEqual(runner.evaluate(case, complete + "\n无法直接完成发言稿", "stop")[0], "FAIL")
        pieces = [section(1), "\n".join(section(i) for i in range(2, 9))]
        def fake(cfg, model, task, folder, rid, timeout):
            text = pieces.pop(0)
            (folder / (rid + ".answer.txt")).write_text(text, encoding="utf-8")
            quality, note = runner.evaluate(task, text, "stop")
            return dict(request_id=rid, status="OK", quality=quality, quality_note=note, finish_reason="stop", request_phase="completed", failure_type="MODEL_BEHAVIOR", failure_reason=note, attempt_count=1, total_s=1, requested_max_tokens=8192)
        with tempfile.TemporaryDirectory() as directory, patch.object(runner, "request_once", side_effect=fake):
            result = runner.request_model(self.config(), "mock", case, Path(directory), "speech")
        self.assertEqual(result["raw_quality"], "FAIL")
        self.assertEqual(result["final_quality"], "PASS")
        self.assertEqual(result["attempt_count"], 2)

    def test_bad_document_blocked(self):
        case = dict(id="doc", source_text="损坏\ufffd")
        with self.assertRaises(policy.PreflightError) as raised:
            policy.preflight(self.config(), "mock", case, dict(max_tokens=1, messages=[]))
        self.assertEqual(raised.exception.category, "DATA_QUALITY")

    def test_report_has_no_historical_diagnosis(self):
        with tempfile.TemporaryDirectory() as directory:
            runner.report(Path(directory), [], "empty", [])
            page = (Path(directory) / "report.html").read_text(encoding="utf-8")
        self.assertNotIn("仅完整输出 0-622", page)
        self.assertNotIn("优先拆成每段 500 项", page)

    def test_recovery_limit_speech(self):
        case = dict(id="speech", group="test", source="mock", prompt="x", rule="speech", minimum=4000)
        with tempfile.TemporaryDirectory() as directory:
            result = runner.request_model(self.config(), "mock", case, Path(directory), "speech")
        self.assertEqual(result["attempt_count"], 4)
        self.assertEqual(result["final_quality"], "FAIL")

    def test_case_parameters(self):
        cases, _ = runner.build_cases(self.config())
        by_id = {case["id"]: case for case in cases}
        self.assertEqual(by_id["count_1000"]["max_tokens"], 8192)
        self.assertFalse(by_id["count_5000"]["do_sample"])
        self.assertTrue(by_id["count_5000"]["stream"])
        self.assertEqual(by_id["count_5000"]["request_timeout_seconds"], 3000)
        self.assertEqual(by_id["speech"]["max_tokens"], 8192)

    def test_network_5xx_retry_and_4xx_no_retry(self):
        case = dict(id="network", group="test", source="mock", prompt="x", stream=True)
        cfg = {**self.config(), "network_retries": 1}
        for endpoint, expected in (("server_error", 2), ("error", 1)):
            cfg["endpoint"] = f"http://127.0.0.1:{self.server.server_port}/{endpoint}"
            with tempfile.TemporaryDirectory() as directory:
                result = runner.request_model(cfg, "mock", case, Path(directory), endpoint)
            self.assertEqual(result["attempt_count"], expected)
            self.assertEqual(result["failure_type"], "NETWORK_ERROR")
            self.assertEqual(result["final_quality"], "FAIL")

    def test_translation_chunk_recovery_retains_raw_failure(self):
        source = "第一章 内容\n1. 内容。\n"
        translation = "Chapter 1 Contents\n1. Complete translation."
        case = dict(id="doc_translate", group="test", source="mock", prompt="translate", source_text=source, rule="translate", max_tokens=8192)
        answers = ["Incomplete", "Still incomplete", translation]
        def fake(cfg, model, task, folder, rid, timeout):
            answer = answers.pop(0)
            (folder / (rid + ".answer.txt")).write_text(answer, encoding="utf-8")
            quality, note = runner.evaluate(task, answer, "stop")
            return dict(request_id=rid, status="OK", quality=quality, quality_note=note, finish_reason="stop", request_phase="completed", failure_type="MODEL_BEHAVIOR" if quality == "FAIL" else "", failure_reason=note, attempt_count=1, total_s=1, requested_max_tokens=task["max_tokens"])
        cfg = {**self.config(), "translation_chunk_recovery": True}
        with tempfile.TemporaryDirectory() as directory, patch.object(runner, "request_once", side_effect=fake):
            result = runner.request_model(cfg, "mock", case, Path(directory), "translation")
        self.assertEqual(result["raw_quality"], "FAIL")
        self.assertEqual(result["final_quality"], "PASS")
        self.assertEqual(result["attempt_count"], 3)
        text = source * 700
        chunks = policy.translation_chunks(text)
        self.assertEqual("".join(chunks), text)
        self.assertTrue(all(len(c) <= 4000 for c in chunks))


if __name__ == "__main__":
    unittest.main(verbosity=2)
