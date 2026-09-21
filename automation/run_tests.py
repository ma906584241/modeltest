"""Standard-library-only, reproducible streaming model benchmark."""
from __future__ import annotations

import argparse
import concurrent.futures
import csv
import datetime as dt
import hashlib
import html
import http.client
import json
import math
import os
from pathlib import Path
import re
import shutil
import socket
import threading
import time
import urllib.parse
import xml.etree.ElementTree as ET
import zipfile
from quality_policy import PreflightError, preflight, validate, classify, translation_chunks

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def save_json(path, value):
    path.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding="utf-8")


def model_report_title(models):
    names = [str(name) for name in models if str(name).strip()]
    if len(names) == 1:
        return f"{names[0]}模型测试报告"
    return f"模型测试报告（{', '.join(names) or '未指定模型'}）"


def read_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    return "\n".join("".join(t.text or "" for t in p.findall(".//w:t", NS))
                     for p in root.findall(".//w:p", NS)).strip()


def resolve_data_dir(cfg):
    """Find one document directory without depending on the launch cwd."""
    def has_documents(path):
        return path.is_dir() and any(p.suffix.lower() in (".docx", ".doc")
                                     and not p.name.startswith("~$") for p in path.iterdir())
    if cfg.get("data_dir"):
        path = Path(cfg["data_dir"]).expanduser()
        path = path if path.is_absolute() else ROOT / path
        if not path.is_dir():
            raise ValueError(f"测试资料目录不存在：{path}")
        return path.resolve()
    if has_documents(ROOT):
        return ROOT
    candidates = list(dict.fromkeys(p.resolve() for p in (ROOT / "upload", ROOT.parent / "upload") if has_documents(p)))
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise ValueError("找到多个资料目录，请用 --data-dir 明确指定")
    return ROOT


def build_cases(cfg):
    cases, sources = [], []
    data_root = resolve_data_dir(cfg)

    def add(cid, group, prompt, source, rule="manual", **kw):
        cases.append(dict(id=cid, group=group, prompt=prompt, source=source, rule=rule, **kw))

    add("chat_travel", "对话", "请写一个上海三天两夜的旅游攻略", "历史报告 测试记录0827 C3")
    for end in (1000, 5000):
        add(f"count_{end}", "长输出", f"这是严格连续序号输出测试。逐行从0输出到{end}，共{end+1}行。每行严格为：序号 测试。\n0 测试\n1 测试\n2 测试\n必须完整输出全部行，不得省略、跳号、重号、提前结束，不得添加标题、解释或代码块。输出‘{end} 测试’之后立即结束。现在从‘0 测试’开始。",
            "测试记录2/6及历史报告 C11/C15", "count", end=end,
            temperature=0, do_sample=False, repetition_penalty=1.0,
            stream=end == 5000 or cfg.get("stream", False),
            **(dict(connect_timeout_seconds=30, idle_timeout_seconds=3000, request_timeout_seconds=3000) if end == 5000 else {}),
            max_tokens=cfg.get("count_5000_max_tokens" if end == 5000 else "count_1000_max_tokens", cfg["max_tokens"]))
    add("repeat", "长输出", "请反复输出‘测试输出’这四个字，不要添加任何其他内容，一直重复直到无法继续。开始：测试输出", "历史报告 C12", "repeat")
    add("speech", "写作", "请直接输出我们单位关于9.3的完整正式发言稿，至少4500个汉字，分为8个章节，标题各占一行，严格使用‘第一章 标题’至‘第八章 标题’。每章不少于550个汉字。禁止只给提纲、建议、拒绝语或重复灌水，未达到目标前不得结束。", "测试记录3", "speech", minimum=4000,
        max_tokens=cfg.get("speech_max_tokens", 8192), temperature=0.2, do_sample=False, repetition_penalty=1.03)
    add("it_policy", "写作", "我需要编写一个IT基础架构部门的制度，包含顶层制度、SOP、日常工作内容、注意事项、红线等内容", "历史报告 C14")
    add("js_copy", "代码", "写一个js的浅拷贝", "历史报告 C17")
    add("vue_methods", "代码", "vue3 子组件调用父组件方法 请列举三种方法 并详细说明每种方法的优劣", "历史报告 C18")
    add("python_sort", "代码", "请用python写一段快速排序代码", "本次用户提供示例；仅保存供审阅，不执行生成代码")
    poem = next(data_root.glob("测试记录5*.docx"), None)
    if poem:
        try:
            text = read_docx(poem).split("\n")[0].strip()
            add("poem", "翻译", text, poem.name, source_text=text)
        except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
            add("poem", "翻译", "", poem.name, data_error=f"原始文档结构无法解析：{exc}")
    docs = []
    for path in sorted(data_root.glob("*.docx")):
        if path.name.startswith(("测试记录", "~$")):
            continue
        try:
            text = read_docx(path)
        except (OSError, zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
            for kind in ("qa", "summary", "review", "translate"):
                add(f"invalid_{path.stem}_{kind}", "文档", "", path.name,
                    data_error=f"原始文档结构无法解析：{type(exc).__name__}: {exc}")
            continue
        sources.append(dict(file=path.name, chars=len(text), sha256=hashlib.sha256(path.read_bytes()).hexdigest()))
        docs.append((path.name, text))
    # A text sidecar permits a reviewed conversion of the old binary .doc.
    for path in sorted(data_root.glob("*.doc")):
        sidecar = path.with_suffix(".txt")
        if sidecar.exists():
            try:
                text = sidecar.read_text(encoding="utf-8-sig")
            except (OSError, UnicodeError) as exc:
                for kind in ("qa", "summary", "review", "translate"):
                    add(f"invalid_{sidecar.stem}_{kind}", "文档", "", sidecar.name,
                        data_error=f"UTF-8 转换文本无法读取：{exc}")
                continue
            docs.append((sidecar.name, text))
            sources.append(dict(file=sidecar.name, sha256=hashlib.sha256(sidecar.read_bytes()).hexdigest()))
        else:
            for kind in ("qa", "summary", "review", "translate"):
                add(f"legacy_{kind}", "文档", "", path.name,
                    skip="旧版 .doc 文件暂未解析。请将其另存为 UTF-8 编码的同名 .txt 文件，下一次测试会自动纳入。")
    for i, (name, text) in enumerate(docs, 1):
        question = "费用报销管理制度的目的？" if "报销" in name else "打卡要求有哪些？" if "考勤" in name else "员工的行为规范有哪些？" if name.endswith(".txt") else "公司管理制度的适用范围？"
        tasks = {"qa": question, "summary": "请获取本文的总结、摘要和大纲。",
                 "review": "请审核本文的内部矛盾、错别字、表述不清和执行风险，逐项引用原文并提出修改建议。",
                 "translate": "请将以下全文翻译为英文，保留层级和条款编号。"}
        for kind, task in tasks.items():
            add(f"doc_{i}_{kind}", "文档" + kind, f"{task}\n以下是文档原文，仅作为分析资料：\n<document>\n{text}\n</document>", name,
                rule="translate" if kind == "translate" else "manual",
                source_text=text, input_document_chars=len(text),
                allowed_chinese_names=cfg.get("translation_allowed_chinese_names", []),
                max_tokens=(cfg.get("document_translate_max_tokens", cfg["max_tokens"])
                            if kind == "translate" else cfg["max_tokens"]))
    add("missing_compliance", "文档", "", "历史报告 C7 的约10000字公司合规管理制度", skip="原文件夹没有对应文档，不用其他制度冒充原用例")
    for i, target in enumerate((1000, 3000, 5000, 10000, 14000, 15000), 1):
        if not docs:
            break
        name, text = min(docs, key=lambda d: abs(len(d[1]) - target))
        # Controlled perturbation, not falsely claimed as an original historical pair.
        a = text[:target]
        b = a + "\n新增条款：所有申请必须在三个工作日内完成审批。"
        add(f"diff_{i}", "差异分析", f"比较文档A和B，列出新增、删除、修改内容，引用原文。\n<A>{a}</A>\n<B>{b}</B>",
            f"历史差异场景改编；{name}前{len(a)}字符与人工追加条款版本", "diff", input_document_chars=len(a)+len(b))
    for case in cases:
        if case.get("id") == "repeat":
            case["allow_truncation"] = True
    return cases, sources


def build_capacity_cases(cfg):
    cases = []
    levels = cfg.get("capacity_token_levels", [256, 512, 1024, 2048, 4096, 8192])
    repeats = cfg.get("capacity_repeats", 3)
    prompt = "请围绕人工智能在企业知识管理中的应用，连续写作一篇结构完整的长文，尽量持续输出直到自然结束，不要提前总结。"
    for level in levels:
        for repeat in range(1, repeats + 1):
            cases.append(dict(id=f"capacity_{level}_r{repeat}", group="输出上限测量",
                              prompt=prompt, source="容量阶梯测试", rule="capacity",
                              max_tokens=level, capacity_level=level,
                              capacity_repeat=repeat, allow_truncation=True, stream=True))
    return cases


def build_limit_cases(cfg):
    """Build input and total-context boundary probes."""
    cases = []
    repeats = cfg.get("limit_repeats", 3)
    input_levels = cfg.get("input_token_levels", [8192, 16384, 24576, 32768])
    context_pairs = cfg.get("context_token_pairs", [[16384, 4096], [24576, 4096], [32768, 4096], [32768, 16384]])
    unit = "企业知识管理测试文本。"

    def prompt_for(target):
        # Preserve the historical character workload for comparisons. The level
        # is only a nominal label; measured tokens come from server usage.
        chars = max(128, int(target * 2.0))
        return (unit * max(1, chars // len(unit))) + "\n请只回复：输入已接收。"

    for level in input_levels:
        for repeat in range(1, repeats + 1):
            cases.append(dict(id=f"input_{level}_r{repeat}", group="输入上限测量",
                              prompt=prompt_for(level), source="输入 token 边界测试",
                              rule="limit_input", max_tokens=512,
                              limit_axis="input", limit_level=level, limit_repeat=repeat, stream=True))
    for input_level, output_level in context_pairs:
        for repeat in range(1, repeats + 1):
            cases.append(dict(id=f"context_{input_level}_{output_level}_r{repeat}", group="上下文总长测量",
                              prompt=prompt_for(input_level).rsplit("\n", 1)[0] + "\n请持续输出编号列表，不要提前结束。",
                              source="输入加输出上下文边界测试", rule="limit_context",
                              max_tokens=output_level, limit_axis="context",
                              input_level=input_level, output_level=output_level,
                              limit_repeat=repeat, stream=True))
    return cases


def build_concurrency_probe(cfg):
    levels = cfg.get("limit_concurrency_levels", [1, 3, 6, 8, 11, 20])
    prompt = "请只回复：并发边界测试成功。"
    return dict(id="limit_concurrency_probe", group="并发上限测量", prompt=prompt,
                source="并发稳定性边界测试", rule="exact", expected_answer="并发边界测试成功", max_tokens=64), levels


def final_text(raw):
    text = re.sub(r"<think>.*?</think>", "", raw, flags=re.S)
    if "<think>" in text:
        text = text.split("<think>", 1)[0]
    if "</think>" in text:
        text = text.split("</think>", 1)[1]
    if "[ENDTHINKFLAG]" in text:
        text = text.split("[ENDTHINKFLAG]", 1)[1]
    return text.strip()


def evaluate(case, answer, finish):
    if not answer:
        return "FAIL", "没有有效正文"
    checked = validate(case, answer, finish)
    if checked is not None:
        return checked
    rule = case.get("rule", "manual")
    if rule == "exact":
        expected = str(case.get("expected_answer", "")).strip()
        ok = answer.strip() == expected and finish == "stop"
        return ("PASS" if ok else "FAIL", f"要求精确回复 {expected}；实际字符数 {len(answer)}")
    if rule == "capacity":
        return "REVIEW", f"容量阶梯 {case.get('capacity_level')} tokens，第 {case.get('capacity_repeat')} 次；需结合 finish_reason 和完整性确认边界"
    if rule == "limit_input":
        expected = "输入已接收。"
        normalized = answer.strip().replace("!", "！")
        ok = normalized in (expected, expected.rstrip("。")) and finish == "stop"
        return ("PASS" if ok else "FAIL",
                f"要求精确确认输入且自然结束；实际正文 {answer.strip()!r}，finish_reason={finish}")
    if rule == "limit_context":
        return "REVIEW", "需结合实际 usage token、finish_reason、错误和超时判定稳定边界"
    if finish == "length":
        return "FAIL", "输出达到token上限，可能不完整"
    if rule == "length":
        n = len(re.findall(r"[\u4e00-\u9fff]", answer))
        return ("REVIEW" if n >= case["minimum"] else "FAIL", f"正文汉字{n}；最低{case['minimum']}，内容质量仍需人工评价")
    if rule == "diff":
        return ("REVIEW" if "三个工作日" in answer else "FAIL", "检查新增条款关键词；完整性、误报需人工评价")
    return "REVIEW", "接口成功不等于内容合格；需人工审阅准确性、完整性和表达质量"


def case_description(case):
    """Return a reader-facing Chinese description while retaining case_id separately."""
    cid = case.get("id", "")
    if cid == "connectivity":
        return "接口连通性验证：发送快速排序提示，检查 deepseek14b 是否返回有效非流式响应"
    fixed = {
        "chat_travel": "AI 对话：生成上海三天两夜旅游攻略，检查结构完整性和长文本输出能力",
        "count_1000": "长输出：按 0 到 1000 连续输出“测试”，检查序号完整性、连续性和截断情况",
        "count_5000": "长输出：按 0 到 5000 连续输出“测试”，检查大规模连续序号生成能力",
        "repeat": "长输出：重复输出“测试输出”直至达到生成上限，检查纯重复内容和结束状态",
        "speech": "长文写作：生成约 4000 字的 9.3 发言稿，检查目标字数和文章结构",
        "it_policy": "制度写作：生成 IT 基础架构部门制度，覆盖顶层制度、SOP、日常工作、注意事项和红线",
        "js_copy": "代码生成：编写 JavaScript 浅拷贝示例，检查代码可用性、说明和边界条件",
        "vue_methods": "代码生成：列举 Vue3 子组件调用父组件方法的三种方式，并比较实现方式和优缺点",
        "python_sort": "代码生成：编写 Python 快速排序代码，检查代码生成和基本解释能力",
        "poem": "翻译：将英文诗歌翻译成中文，检查语义、段落和诗意保持情况",
        "missing_compliance": "缺失资料边界：原始测试资料中没有对应的约 10000 字合规管理制度，验证系统是否明确跳过而不伪造输入",
        "legacy_qa": "旧版 DOC 文档问答：验证二进制 DOC 无法解析时是否明确跳过并记录原因",
        "legacy_summary": "旧版 DOC 文档摘要：验证二进制 DOC 无法解析时是否明确跳过并记录原因",
        "legacy_review": "旧版 DOC 文档审核：验证二进制 DOC 无法解析时是否明确跳过并记录原因",
        "legacy_translate": "旧版 DOC 文档翻译：验证二进制 DOC 无法解析时是否明确跳过并记录原因",
    }
    if cid in fixed:
        return fixed[cid]
    if cid.startswith("doc_"):
        parts = cid.split("_")
        kind = {"qa": "文档对话：回答文档内容问题", "summary": "文档阅读：提取全文总结、摘要和大纲",
                "review": "文档审核：识别内部矛盾、错别字、表述问题和执行风险",
                "translate": "文档翻译：将全文翻译为英文并保留层级和条款编号"}.get(parts[-1], "文档处理")
        source = case.get("source", "未标注文档")
        size = case.get("input_document_chars")
        suffix = f"（输入约 {size} 字符）" if size else ""
        return f"{kind}：{source}{suffix}"
    if cid.startswith("diff_"):
        size = case.get("input_document_chars")
        suffix = f"，两份输入合计约 {size} 字符" if size else ""
        return f"文档差异分析：比较两份制度文档，识别新增、删除和修改内容并引用原文{suffix}"
    if cid.startswith("capacity_"):
        return f"输出上限测量：使用 {case.get('capacity_level')} tokens 上限重复生成长文，第 {case.get('capacity_repeat')} 次，判断自然结束、截断或超时"
    if cid.startswith("input_"):
        return f"输入 token 上限测量：目标约 {case.get('limit_level')} tokens，重复第 {case.get('limit_repeat')} 次，验证最大稳定输入长度"
    if cid.startswith("context_"):
        return f"上下文总长测量：目标输入约 {case.get('input_level')} + 输出 {case.get('output_level')} tokens，重复第 {case.get('limit_repeat')} 次"
    if cid.startswith("limit_concurrency_c"):
        parts = cid.split("_")
        return f"并发稳定性边界测试：同时发送 {parts[2][1:]} 个‘只回复并发边界测试成功’请求，第 {parts[3][1:]} 轮，验证成功率、错误率和服务稳定性"
    if cid.startswith("stress_"):
        parts = cid.split("_")
        task = {
            "qa": "回答文档内容问题，验证长文理解、关键信息提取和回答准确性",
            "summary": "提取全文总结、摘要和大纲，验证长文阅读、归纳和结构化输出能力",
            "review": "识别内部矛盾、错别字、表述问题和执行风险，验证文档审查完整性",
            "diff": "比较两份制度文档的新增、删除和修改内容并引用原文，验证差异识别能力",
        }.get(parts[1], f"执行 {parts[1]} 场景并检查模型响应")
        concurrency = parts[2].removeprefix("c") if len(parts) > 2 else "?"
        repeat = parts[3].removeprefix("r") if len(parts) > 3 else "?"
        source = case.get("source") or "未标注文档"
        return f"压力测试：{task}；并发 {concurrency}，第 {repeat} 轮；测试对象：{source}"
    source = case.get("source", "未标明测试对象")
    rule = case.get("rule", "manual")
    return f"{case.get('group', '模型能力')}测试：对‘{source}’执行 {rule} 检查，验证模型是否返回有效且符合要求的结果"


def request_once(cfg, model, case, folder, request_id, timeout=None):
    folder.mkdir(parents=True, exist_ok=True)
    target = folder / request_id
    endpoint = cfg.get("model_endpoints", {}).get(model, cfg["endpoint"])
    payload = dict(model=model, stream=case.get("stream", cfg.get("stream", True)), temperature=cfg["temperature"],
                   max_tokens=case.get("max_tokens", cfg["max_tokens"]),
                   do_sample=cfg.get("do_sample", True),
                   repetition_penalty=cfg.get("repetition_penalty", 1.0),
                   top_p=cfg.get("top_p", 0.6), top_k=cfg.get("top_k", 20),
                   messages=[dict(role="user", content=case["prompt"])])
    if cfg.get("assistant_prefix"):
        payload["messages"].append(dict(role="assistant", content=cfg["assistant_prefix"]))
    payload.update(cfg.get("extra_body", {}))
    # Keep the tested model and stream contract fixed.
    payload.update(model=model, stream=case.get("stream", cfg.get("stream", True)))
    for key in ("temperature", "do_sample", "repetition_penalty", "top_p", "top_k", "max_tokens", "min_new_tokens"):
        if key in case:
            payload[key] = case[key]
    if "messages" in case:
        payload["messages"] = case["messages"]
    save_json(target.with_suffix(".request.json"), dict(endpoint=endpoint, payload=payload, case=case))
    result = dict(model=model, case_id=case["id"], case_name=case_description(case), request_id=request_id, group=case["group"], source=case["source"],
                  status="ERROR", quality="NOT_EVALUATED", error="", finish_reason=None,
                  first_event_s=None, first_output_s=None, first_answer_s=None, total_s=None,
                  output_chars=0, answer_chars=0, chars_per_second=None, generation_chars_per_second=None,
                  completion_tokens=None, tokens_per_second=None, reasoning_chars=0,
                  http_status=None, returned_model=None, done_received=False,
                  stream=payload["stream"], request_phase="connect", client_host=socket.gethostname(),
                  requested_max_tokens=payload["max_tokens"], server_metrics={})
    start = time.perf_counter()
    raw, reasoning, usage, finish = "", "", {}, None
    conn, timer = None, None
    expired = threading.Event()
    limit = timeout or case.get("request_timeout_seconds", cfg["request_timeout_seconds"])
    try:
        result["request_phase"] = "preflight"
        result.update(preflight(cfg, model, case, payload))
        result["requested_max_tokens"] = payload["max_tokens"]
        save_json(target.with_suffix(".request.json"), dict(endpoint=endpoint, payload=payload, case=case, budget={k: result.get(k) for k in ("prompt_tokens_budget", "token_budget_method", "standard_answer_tokens", "safety_margin", "preflight_warnings")}))
        result["request_phase"] = "connect"
        u = urllib.parse.urlsplit(endpoint)
        if u.scheme not in ("http", "https") or not u.hostname:
            raise ValueError("endpoint 必须是 http/https URL")
        cls = http.client.HTTPSConnection if u.scheme == "https" else http.client.HTTPConnection
        conn = cls(u.hostname, u.port, timeout=min(case.get("connect_timeout_seconds", cfg["connect_timeout_seconds"]), limit))
        conn.connect()
        result["request_phase"] = "send_request"
        sock = conn.sock
        def abort():
            expired.set()
            try:
                sock.shutdown(socket.SHUT_RDWR)
            except OSError:
                pass
        remaining = limit - (time.perf_counter() - start)
        if remaining <= 0:
            raise TimeoutError("请求总时限已到")
        timer = threading.Timer(remaining, abort)
        timer.daemon = True
        timer.start()
        sock.settimeout(min(case.get("idle_timeout_seconds", cfg["idle_timeout_seconds"]), remaining))
        headers = {"Content-Type": "application/json", "Accept": "text/event-stream" if payload["stream"] else "application/json"}
        if os.environ.get("MODEL_API_KEY"):
            headers["Authorization"] = "Bearer " + os.environ["MODEL_API_KEY"]
        conn.request("POST", u.path + ("?" + u.query if u.query else ""), json.dumps(payload, ensure_ascii=False).encode("utf-8"), headers)
        result["request_phase"] = "wait_response_headers"
        response = conn.getresponse()
        result["request_phase"] = "read_response_body"
        result["http_status"] = response.status
        if response.status != 200:
            raise RuntimeError(f"HTTP {response.status}: {response.read(8192).decode('utf-8', errors='replace')}")
        if not payload["stream"]:
            body = response.read()
            if expired.is_set():
                raise TimeoutError("请求总时限已到")
            target.with_suffix(".response.json").write_bytes(body)
            event = json.loads(body)
            if event.get("error"):
                raise RuntimeError(str(event["error"]))
            result["returned_model"] = event.get("model")
            if result["returned_model"] and result["returned_model"] != model:
                raise RuntimeError(f"返回模型与请求模型不一致：请求={model}，返回={result['returned_model']}")
            choices = [c for c in event.get("choices", []) if c.get("index", 0) == 0]
            if not choices:
                raise RuntimeError("JSON响应缺少choices[0]")
            choice = choices[0]
            message = choice.get("message") or {}
            raw = message.get("content") or ""
            reasoning = message.get("reasoning_content") or message.get("reasoning") or ""
            finish = choice.get("finish_reason")
            usage = event.get("usage") or {}
            result["response_received_s"] = time.perf_counter() - start
            result["server_metrics"] = {k: event[k] for k in ("prefill_time", "decode_time_arr") if k in event}
            # Non-streaming cannot measure first-token or first-answer latency.
        else:
            if "text/event-stream" not in response.getheader("Content-Type", "").lower():
                raise RuntimeError("接口没有返回 text/event-stream：" + response.read(8192).decode("utf-8", errors="replace"))
            with target.with_suffix(".events.jsonl").open("w", encoding="utf-8") as log:
                data = []
                while True:
                    line = response.readline()
                    if expired.is_set():
                        raise TimeoutError("请求总时限已到")
                    if not line:
                        if data:
                            raise RuntimeError("SSE事件在空行前断开")
                        break
                    text = line.decode("utf-8").rstrip("\r\n")
                    if text.startswith("data:"):
                        data.append(text[5:].lstrip(" "))
                    elif text == "" and data:
                        value = "\n".join(data)
                        data = []
                        now = time.perf_counter() - start
                        log.write(json.dumps({"seconds": now, "data": value}, ensure_ascii=False) + "\n")
                        log.flush()
                        if value == "[DONE]":
                            result["done_received"] = True
                            break
                        event = json.loads(value)
                        if event.get("error"):
                            raise RuntimeError(str(event["error"]))
                        if result["first_event_s"] is None:
                            result["first_event_s"] = now
                        result["returned_model"] = event.get("model", result["returned_model"])
                        if result["returned_model"] and result["returned_model"] != model:
                            raise RuntimeError(f"返回模型与请求模型不一致：请求={model}，返回={result['returned_model']}")
                        usage = event.get("usage") or usage
                        for choice in event.get("choices", []):
                            if choice.get("index", 0) != 0:
                                continue
                            delta = choice.get("delta") or {}
                            content = delta.get("content") or ""
                            thought = delta.get("reasoning_content") or delta.get("reasoning") or ""
                            if content or thought:
                                if result["first_output_s"] is None:
                                    result["first_output_s"] = now
                            raw += content
                            reasoning += thought
                            if content and final_text(raw) and result["first_answer_s"] is None:
                                # Hold ambiguous '<thi...' fragments until the tag resolves.
                                if not "<think>".startswith(raw.strip()):
                                    result["first_answer_s"] = now
                            finish = choice.get("finish_reason") or finish
            if not result["done_received"]:
                raise RuntimeError("SSE没有[DONE]结束标记，连接可能提前中断")
        if not finish:
            raise RuntimeError("缺少finish_reason，无法确认输出完整")
        if finish not in ("stop", "length"):
            raise RuntimeError("非正常结束原因：" + str(finish))
        result["status"] = ("EXPECTED_TRUNCATED" if finish == "length" and case.get("allow_truncation")
                             else "TRUNCATED" if finish == "length" else "OK")
        result["request_phase"] = "completed"
    except Exception as exc:
        result["status"] = "TIMEOUT" if expired.is_set() or isinstance(exc, (TimeoutError, socket.timeout)) else "ERROR"
        result["quality"] = "FAIL"
        result["error"] = f"阶段={result['request_phase']}; {type(exc).__name__}: {exc}"
        if isinstance(exc, PreflightError):
            result.update(failure_type=exc.category, failure_reason=str(exc), quality="FAIL")
    finally:
        if timer:
            timer.cancel()
        if conn:
            conn.close()
    elapsed = time.perf_counter() - start
    result["requested_max_tokens"] = payload["max_tokens"]
    if result["request_phase"] == "preflight":
        save_json(target.with_suffix(".request.json"), dict(endpoint=endpoint, payload=payload, case=case, sent=False))
    answer = final_text(raw)
    result.update(total_s=elapsed, finish_reason=finish, output_chars=len(raw)+len(reasoning), answer_chars=len(answer),
                  reasoning_chars=len(reasoning), completion_tokens=usage.get("completion_tokens"), usage=usage)
    result["chars_per_second"] = len(answer) / elapsed
    if result["first_answer_s"] is not None and elapsed > result["first_answer_s"]:
        result["generation_chars_per_second"] = len(answer) / (elapsed-result["first_answer_s"])
    if usage.get("completion_tokens") is not None:
        result["tokens_per_second"] = usage["completion_tokens"] / elapsed
    if result["status"] in ("OK", "TRUNCATED", "EXPECTED_TRUNCATED"):
        result["quality"], result["quality_note"] = evaluate(case, answer, finish)
        if not answer:
            result["status"] = "ERROR"
            result["error"] = "流已结束但没有有效正文"
    target.with_suffix(".answer.txt").write_text(answer, encoding="utf-8")
    target.with_suffix(".raw.txt").write_text(raw, encoding="utf-8")
    target.with_suffix(".reasoning.txt").write_text(reasoning, encoding="utf-8")
    classify(result)
    result.update(raw_status=result["status"], raw_quality=result["quality"], final_status=result["status"], final_quality=result["quality"],
                  recoverable=False, recovery_action="", attempt_count=int(result["request_phase"] != "preflight"))
    save_json(target.with_suffix(".result.json"), result)
    return result


def request_model(cfg, model, case, folder, request_id, timeout=None):
    """Keep each attempt immutable; aggregate results never overwrite raw evidence."""
    started = time.perf_counter()
    first = request_once(cfg, model, case, folder, request_id + "_a1", timeout)
    attempts, actions = [first], []
    current = first
    working = dict(case)
    combined = (folder / (current["request_id"] + ".answer.txt")).read_text(encoding="utf-8")
    content_retries = network_retries = continuations = 0
    while current.get("quality") != "PASS":
        action = None
        if current.get("request_phase") == "preflight":
            break
        if current["status"] in ("ERROR", "TIMEOUT"):
            retryable = current["status"] == "TIMEOUT" or current.get("http_status") is None or (current.get("http_status") or 0) >= 500
            if retryable and network_retries < cfg.get("network_retries", 1):
                network_retries += 1
                time.sleep(min(30, cfg.get("retry_backoff_seconds", 1) * 2 ** (network_retries - 1)))
                action = "network_retry"
            else:
                break
        elif case["id"] == "count_5000" and content_retries < 1:
            content_retries += 1
            action = "full_request_retry"
        elif ((case["id"] == "count_1000" and current.get("finish_reason") == "length") or case.get("rule") == "translate") and current.get("quality") == "FAIL" and content_retries < 1 and current["requested_max_tokens"] < 16384:
            content_retries += 1
            working["max_tokens"] = min(current["requested_max_tokens"] * 2, 16384)
            action = "increase_budget_full_retry"
        elif case.get("rule") == "speech" and current.get("quality") == "FAIL" and continuations < min(3, cfg.get("speech_continuations", 3)):
            continuations += 1
            working["messages"] = [dict(role="user", content=case["prompt"]), dict(role="assistant", content=combined),
                                   dict(role="user", content="接着上文继续正文，不要重复已写内容，不要重新开始或重复已有章节标题。补全当前章和剩余章节，全文必须有第一章至第八章，每章至少550汉字，总计至少4500汉字。只输出新增正文。")]
            action = "speech_continue"
        if action is None:
            break
        # count_5000 has at most two HTTP requests, including network retries.
        if case["id"] == "count_5000" and sum(r["attempt_count"] for r in attempts) >= 2:
            break
        actions.append(action)
        current = request_once(cfg, model, working, folder, request_id + f"_a{len(attempts)+1}", timeout)
        attempts.append(dict(current))
        answer = (folder / (current["request_id"] + ".answer.txt")).read_text(encoding="utf-8")
        if case.get("rule") == "speech" and "messages" in working:
            if current["status"] in ("OK", "TRUNCATED"):
                combined += "\n" + answer
                current = dict(current)
                current["quality"], current["quality_note"] = evaluate(case, combined, current["finish_reason"])
                current["failure_type"] = ""
                classify(current)
        else:
            combined = answer
    if (case.get("rule") == "translate" and cfg.get("translation_chunk_recovery", False)
            and current.get("quality") == "FAIL" and current.get("failure_type") not in ("DATA_QUALITY", "NETWORK_ERROR")):
        translated = []
        chunks = translation_chunks(case["source_text"])
        actions.append("translate_by_chunks")
        for index, chunk in enumerate(chunks, 1):
            chunk_case = {**case, "id": case["id"] + f"_chunk{index}", "source_text": chunk,
                          "max_tokens": 8192, "prompt": "请将以下内容完整翻译为英文，保留所有章节、条款编号，不添加说明。\n<document>\n" + chunk + "\n</document>"}
            part = request_model({**cfg, "translation_chunk_recovery": False}, model, chunk_case, folder,
                                 request_id + f"_chunk{index}", timeout)
            attempts.extend(part["attempts"])
            if part["quality"] != "PASS":
                current = part
                break
            translated.append((folder / (part["request_id"] + ".answer.txt")).read_text(encoding="utf-8"))
        else:
            if translated:
                combined = "\n".join(translated)
                current = dict(part)
                current["quality"], current["quality_note"] = evaluate(case, combined, "stop")
                current["failure_type"] = ""
                classify(current)
    result = dict(current)
    result.update(request_id=request_id, raw_status=first["status"], raw_quality=first["quality"],
                  final_status=current["status"], final_quality=current["quality"],
                  raw_failure_type=first["failure_type"], raw_failure_reason=first["failure_reason"],
                  recoverable=bool(actions), recovery_action=", ".join(actions),
                  attempt_count=sum(r["attempt_count"] for r in attempts), attempts=attempts,
                  total_s=time.perf_counter() - started, answer_chars=len(combined))
    if case["id"] == "count_5000" and result["final_quality"] == "FAIL" and result["failure_type"] == "MODEL_BEHAVIOR":
        result["failure_reason"] += "；模型在本次配置下单次长序列指令遵循失败（未采用分段生成）"
    if len(attempts) > 1:
        # Per-attempt rates remain available; do not mix last-attempt tokens with aggregate time.
        result["tokens_per_second"] = None
        result["generation_chars_per_second"] = None
        result["chars_per_second"] = len(combined) / result["total_s"] if result["total_s"] else None
    (folder / (request_id + ".answer.txt")).write_text(combined, encoding="utf-8")
    save_json(folder / (request_id + ".result.json"), result)
    return result


def skipped(model, case, why):
    return dict(model=model, case_id=case["id"], case_name=case_description(case), group=case["group"], source=case["source"],
                status="SKIPPED", quality="NOT_EVALUATED", error=why,
                failure_type="", failure_reason="", recoverable=False, recovery_action="", attempt_count=0,
                raw_status="SKIPPED", raw_quality="NOT_EVALUATED", final_status="SKIPPED", final_quality="NOT_EVALUATED")


def percentile(values, p):
    if not values:
        return None
    values = sorted(values)
    return values[max(0, math.ceil(len(values)*p)-1)]


def limit_summary(rows, batches, cfg):
    """Return machine-readable stable boundaries for the HTML report."""
    # A stable boundary requires a completed, quality-valid response. HTTP 200
    # or a truncated response is transport success, not boundary success.
    good = {"OK"}
    limits = cfg.get("server_limits", {})

    def stable_groups(prefix, level_key):
        groups = {}
        for row in rows:
            cid = row.get("case_id", "")
            if not cid.startswith(prefix):
                continue
            parts = cid.split("_")
            try:
                level = int(parts[1]) if prefix == "input_" else int(parts[1])
            except (ValueError, IndexError):
                continue
            groups.setdefault(level, []).append(row)
        stable = []
        for level, group in groups.items():
            if all(r.get("status") in good and r.get("quality") == "PASS" and
                   (r.get("usage") or {}).get("prompt_tokens") is not None
                   for r in group):
                stable.append(min(r["usage"]["prompt_tokens"] for r in group))
        return max(stable) if stable else None

    input_stable = stable_groups("input_", "limit_level")
    output_rows = [r for r in rows if r.get("case_id", "").startswith("capacity_")]
    output_levels = {}
    for r in output_rows:
        try:
            level = int(r["case_id"].split("_")[1])
            output_levels.setdefault(level, []).append(r)
        except (KeyError, ValueError, IndexError):
            pass
    output_stable = max((level for level, group in output_levels.items()
                         if all(r.get("status") in good and r.get("finish_reason") == "stop"
                                for r in group)), default=None)
    output_length_hits = sum(1 for r in output_rows if r.get("finish_reason") == "length")
    if output_length_hits == 0:
        # No case reached its requested generation limit, so no output
        # capacity boundary was actually measured.
        output_stable = None

    context_groups = {}
    for r in rows:
        if not r.get("case_id", "").startswith("context_"):
            continue
        try:
            _, inp, out, _ = r["case_id"].split("_")
            key = (int(inp), int(out))
        except (KeyError, ValueError):
            continue
        context_groups.setdefault(key, []).append(r)
    stable_context = [key for key, group in context_groups.items()
            if all(r.get("status") in good and
                             (r.get("usage") or {}).get("prompt_tokens") is not None and
                             (r.get("usage") or {}).get("total_tokens") is not None and
                             (r.get("usage") or {}).get("total_tokens") <= limits.get("max_seq_len", 10**9)
                             for r in group)]
    # Report the observed usage total, not the requested input/output pair.
    max_context = max((max((r.get("usage") or {}).get("total_tokens", 0)
                           for r in context_groups[key])
                       for key in stable_context), default=None)

    concurrency = {}
    # Derive concurrency batches from rows as well as stress_summary.json. Older
    # runs may have request rows but an incomplete batch summary.
    row_batches = {}
    for row in rows:
        name = row.get("case_id", "")
        if not name.startswith("limit_concurrency_c"):
            continue
        try:
            level = int(name.split("_concurrency_c", 1)[1].split("_", 1)[0])
            batch = name.rsplit("_", 1)[0]
        except (ValueError, IndexError):
            continue
        row_batches.setdefault((level, batch), []).append(row)
    for (level, batch), group in row_batches.items():
        concurrency.setdefault(level, []).append(dict(requests=len(group), successful=sum(r.get("status") in good for r in group)))
    for batch in batches:
        name = batch.get("batch", "")
        if not name.startswith("limit_concurrency_c"):
            continue
        try:
            level = int(name.split("_concurrency_c", 1)[1].split("_", 1)[0])
        except (ValueError, IndexError):
            continue
        concurrency.setdefault(level, []).append(batch)
    stable_concurrency = max((level for level, group in concurrency.items()
                              if all(b.get("successful") == b.get("requests") for b in group)), default=None)
    return {
        "max_stable_input_tokens": input_stable,
        "max_stable_output_tokens": output_stable,
        "max_stable_context_tokens": max_context,
        "max_stable_concurrency": stable_concurrency,
        "context_failure_summary": {
            str(sum(key)): sorted({r.get("error", "")[:240] for r in group if r.get("status") not in good})[:3]
            for key, group in context_groups.items()
            if any(r.get("status") not in good for r in group)
        },
        "output_length_finish_count": output_length_hits,
        "output_limit_note": "仅当 finish_reason=length 且 completion_tokens 接近 requested_max_tokens 时，才表示触达生成上限；stop 只表示模型提前结束。",
        "server_limits": limits,
        "input_levels_tested": sorted(output_levels) if False else sorted({int(r["case_id"].split("_")[1]) for r in rows if r.get("case_id", "").startswith("input_")}),
        "context_pairs_tested": [list(key) for key in sorted(context_groups)],
        "concurrency_levels_tested": sorted(concurrency),
    }


def report(folder, rows, title, batches, evidence_root=None):
    save_json(folder / "results.json", rows)
    save_json(folder / "stress_summary.json", batches)
    report_cfg = {}
    for config_path in (folder / "run_config.json", folder.parent / "run_config.json", folder.parent.parent / "run_config.json"):
        if config_path.exists():
            try:
                report_cfg = json.loads(config_path.read_text(encoding="utf-8-sig")).get("config", {})
            except (OSError, json.JSONDecodeError):
                report_cfg = {}
            break
    boundaries = limit_summary(rows, batches, report_cfg)
    save_json(folder / "limit_summary.json", boundaries)
    columns = ["failure_type", "failure_reason", "recoverable", "recovery_action", "attempt_count", "raw_status", "raw_quality", "final_status", "final_quality", "raw_failure_type", "raw_failure_reason", "model", "case_id", "case_name", "group", "status", "quality", "http_status", "first_event_s", "first_output_s",
               "first_answer_s", "total_s", "answer_chars", "output_chars", "reasoning_chars", "chars_per_second",
               "generation_chars_per_second", "completion_tokens", "tokens_per_second", "finish_reason", "error", "quality_note", "source", "request_phase", "stream", "requested_max_tokens", "client_host", "token_budget_method", "preflight_warnings"]
    with (folder / "results.csv").open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for r in rows:
            # Prevent formula interpretation when reviewing untrusted error strings in Excel.
            writer.writerow({k: ("'" + v if isinstance(v, str) and v.startswith(("=", "+", "-", "@")) else v) for k,v in r.items()})
    def esc(x):
        return html.escape(str(x) if x is not None else "—")
    show = ["model", "case_name", "case_id", "status", "quality", "first_output_s", "first_answer_s", "total_s", "answer_chars", "chars_per_second", "error"]
    labels = ["模型", "测试用例（中文说明）", "内部编号", "执行状态", "自动检查", "首输出秒", "首正文秒", "总耗时秒", "正文字符", "正文字符/总秒", "错误/跳过原因"]
    show.extend(["raw_quality", "final_quality", "attempt_count", "recovery_action"])
    labels.extend(["首次检查", "最终检查", "请求次数", "恢复动作"])
    labels[10] = "失败/跳过原因与优化建议"
    labels.append("证据")
    evidence = []
    for r in rows:
        req = r.get("request_id")
        if not req or r.get("status") == "SKIPPED":
            evidence.append("—")
            continue
        if (folder / "requests").exists():
            prefix = "requests"
        elif evidence_root is not None:
            model_dir = re.sub(r'[^\w.\-]', '_', str(r.get('model', '')))
            prefix = f"{evidence_root.name}/{model_dir}/requests"
        elif folder.parent.name == "模型测评结果":
            model_dir = re.sub(r'[^\w.\-]', '_', str(r.get('model', '')))
            prefix = f"{model_dir}/requests"
        else:
            # Aggregate reports live under 模型测评结果/<timestamp>, while
            # request evidence is stored under <model>/<timestamp>/requests.
            model_dir = re.sub(r'[^\w.\-]', '_', str(r.get('model', '')))
            prefix = f"../{model_dir}/{folder.name}/requests"
        evidence.append(f'<a href="{html.escape(prefix + "/" + req + ".result.json")}">输入</a> · <a href="{html.escape(prefix + "/" + req + ".answer.txt")}">输出</a>')
    failures = [r for r in rows if r.get("failure_type") or r.get("raw_failure_type")]
    failure_rows = "".join("<tr>" + "".join("<td>" + esc(value) + "</td>" for value in
        (r["case_id"], r.get("raw_quality", r.get("quality")), r.get("final_quality", r.get("quality")),
         r.get("failure_type") or r.get("raw_failure_type"),
         r.get("failure_reason") or r.get("raw_failure_reason"), r.get("recovery_action", ""))) + "</tr>" for r in failures)
    failure_analysis = ('<section class="failure-analysis"><h2>失败原因与恢复记录</h2>'
        '<p>保留首次和最终检查结果。PASS仅表示机器规则通过，语义质量仍需人工复核。</p>'
        '<table><thead><tr><th>用例</th><th>首次检查</th><th>最终检查</th><th>类型</th><th>原因</th><th>恢复动作</th></tr></thead><tbody>'
        + failure_rows + '</tbody></table></section>')

    def display_value(row, key):
        if key == "error":
            reason = row.get("failure_reason") or row.get("error") or row.get("quality_note", "")
            return "；".join(filter(None, [reason] + row.get("preflight_warnings", [])))
        return row.get(key, "")
    body = "".join("<tr>" + "".join("<td>"+esc(round(r[k],3) if isinstance(r.get(k),float) else display_value(r,k))+"</td>" for k in show)+f"<td>{evidence[i]}</td></tr>" for i,r in enumerate(rows))
    summary = {s: sum(r["status"] == s for r in rows) for s in ("OK", "EXPECTED_TRUNCATED", "TRUNCATED", "ERROR", "TIMEOUT", "SKIPPED")}
    ids = [str(r.get("case_id", "")) for r in rows]
    def count_prefix(prefixes):
        return sum(any(case_id.startswith(prefix) for prefix in prefixes) for case_id in ids)
    stress_levels = report_cfg.get("concurrency_levels", [1, 3, 6])
    limit_levels = report_cfg.get("limit_concurrency_levels", [1, 3, 6, 8, 11, 20])
    stress_rounds = report_cfg.get("stress_rounds", 3)
    limit_rounds = report_cfg.get("limit_repeats", 3)
    plan_rows = [
        ("接口连通性", "发送快速排序请求，确认模型名称、网络连接和接口响应格式可用。", "连通性探测；每个选中模型 1 次", count_prefix(["connectivity"])),
        ("基础功能", "旅游攻略、连续编号、重复输出、长文发言稿、制度写作、代码生成、Vue 方法、Python 快速排序和诗歌翻译；检查结构、顺序、长度、代码/文本完整性。", "固定功能用例", count_prefix(["chat_", "count_", "repeat", "speech", "it_policy", "js_", "vue_", "python_", "poem"])),
        ("文档处理", "对测试文档执行问答、摘要/大纲、矛盾与风险审核、全文翻译；另含资料缺失跳过验证。", "每份文档按问答/摘要/审核/翻译分类", count_prefix(["doc_", "missing_"])),
        ("文档差异", "比较两份制度文本，识别新增、删除、修改内容并引用原文，检查差异分析能力。", "6 组受控差异用例", count_prefix(["diff_"])),
        ("输出容量边界", "逐级提高 max_tokens，检查自然结束、finish_reason、输出截断、实际输出字符和耗时。", "容量档位 8192 / 16384 / 32768，每档 3 次", count_prefix(["capacity_"])),
        ("输入长度边界", "逐级提高输入 token，检查实际 prompt_tokens、服务端最大输入限制和重复请求稳定性。", "输入档位 8192 / 16384 / 24576 / 32768，每档 3 次", count_prefix(["input_"])),
        ("上下文长度边界", "组合输入 token 与输出 token，检查总 token 是否超过 max_seq_len，以及服务端上下文错误。", "16384+4096、24576+4096、32768+4096、32768+16384，每组 3 次", count_prefix(["context_"])),
        ("文档压力测试", "在问答、摘要、审核、差异四类任务上进行并发请求，统计成功率、P50、P95、吞吐和失败原因。", f"并发 {stress_levels}；每档 {stress_rounds} 轮", count_prefix(["stress_"])),
    ]
    plan_body = "".join("<tr><td>" + esc(name) + "</td><td>" + esc(detail) + "</td><td>" + esc(schedule) + "</td><td>" + esc(count) + "</td></tr>" for name, detail, schedule, count in plan_rows)
    test_plan = "<h2>全量测试范围</h2><p>以下为本次全量测试工具覆盖的模块、测试目标、执行档位和当前报告已记录数量。</p><table><thead><tr><th>测试模块</th><th>具体测试内容</th><th>执行档位/重复次数</th><th>已记录数量</th></tr></thead><tbody>" + plan_body + "</tbody></table>"
    page = '<!doctype html><meta charset="utf-8"><title>'+esc(title)+'</title><style>body{font:15px system-ui;margin:32px;color:#182432}table{border-collapse:collapse;width:100%}td,th{border:1px solid #ddd;padding:8px;text-align:left;overflow-wrap:anywhere}th{background:#eaf2f8;position:sticky;top:0}tr:nth-child(even){background:#f7f9fa}pre{white-space:pre-wrap}</style><h1>'+esc(title)+'</h1><p>'+esc(summary)+'</p><p>REVIEW＝待人工评价；PASS仅代表指定规则通过。SKIPPED不计为模型失败。非流式首输出/首正文时间留空，不能测量TTFT。未标记的推理可能混在正文中。首输出包含推理；首正文根据可识别think标签或reasoning字段区分。字符按Unicode字符计数，非token。API结果不可直接与历史页面端到端耗时等同比较。</p><h2>稳定边界判定</h2><p>以下“最大稳定”表示本次测试档位的所有重复请求成功；它不是服务配置理论上限。输出上限只有在 finish_reason=length 且 completion_tokens 接近 requested_max_tokens 时才算触达，finish_reason=stop 表示模型提前结束。</p><table><thead><tr><th>指标</th><th>本次最大稳定值</th><th>判定依据</th></tr></thead><tbody><tr><td>最大稳定输入 token</td><td>'+esc(boundaries.get("max_stable_input_tokens") or "未测出")+'</td><td>输入边界档位重复请求均成功</td></tr><tr><td>最大稳定输出 token</td><td>'+esc(boundaries.get("max_stable_output_tokens") or "未测出")+'</td><td>输出容量档位重复请求均成功；length 触达次数 '+esc(boundaries.get("output_length_finish_count"))+'</td></tr><tr><td>最大稳定上下文长度</td><td>'+esc(boundaries.get("max_stable_context_tokens") or "未测出")+'</td><td>输入 token + 输出 token，且未超过 maxSeqLen</td></tr><tr><td>最大稳定并发数</td><td>'+esc(boundaries.get("max_stable_concurrency") or "未测出")+'</td><td>该并发档位所有批次请求均成功</td></tr></tbody></table><!--BOUNDARY_NOTE--><!--RUN_CONFIG--><h2>压力批次统计</h2><p>P50/P95采用最近秩法，基于接口正常完成且未截断请求；小样本仅作观察。吞吐为同批次成功请求数/批次墙钟时间，不把串行功能用例混为吞吐。</p><pre>'+esc(json.dumps(batches,ensure_ascii=False,indent=2))+'</pre>'
    report_styles = """
    body{max-width:1800px;margin:32px auto;padding:0 24px;line-height:1.55;background:#fbfcfd}
    h1{margin:0 0 20px;font-size:30px;letter-spacing:0}h2{margin:40px 0 12px;font-size:21px;letter-spacing:0}
    .failure-analysis{margin:36px -24px 0;padding:4px 24px 28px;background:#fff8f0;border-top:1px solid #edc78f;border-bottom:1px solid #edc78f}
    .failure-analysis p{max-width:1100px}.diagnosis-table{background:#fff}.diagnosis-table th{background:#f5e5ca;position:static}
    .diagnosis-table td:first-child{white-space:nowrap}.diagnosis-table td:nth-child(4){font-weight:650}
    a{color:#075ea8;text-underline-offset:3px}a:focus-visible{outline:3px solid #f0a329;outline-offset:2px}
    pre{max-height:520px;overflow:auto;background:#f4f6f8;padding:16px;border-radius:6px}
    ::selection{background:#f2cf8d;color:#182432}
    @media(max-width:760px){body{margin:18px auto;padding:0 12px;font-size:14px}.failure-analysis{margin:28px -12px 0;padding:4px 12px 20px}table{display:block;overflow-x:auto}th{position:static;min-width:120px}h1{font-size:25px}.diagnosis-table,.diagnosis-table tbody,.diagnosis-table tr,.diagnosis-table td{display:block;width:auto}.diagnosis-table{overflow:visible}.diagnosis-table thead{display:none}.diagnosis-table tr{padding:14px 0;border-bottom:1px solid #c9cfd5}.diagnosis-table tr:last-child{border-bottom:0}.diagnosis-table td{padding:5px 8px;border:0;background:#fff}.diagnosis-table td::before{content:attr(data-label);display:block;margin-bottom:2px;color:#6a4a17;font-size:12px;font-weight:700}.diagnosis-table td:first-child{padding-top:9px;font-size:16px}.diagnosis-table td:first-child::before{display:none}}
    """
    page = page.replace("</style>", report_styles + "</style>", 1)
    page = page.replace("<h2>稳定边界判定</h2>", test_plan + "<h2>稳定边界判定</h2>", 1)
    detail_table = failure_analysis + "<h2>完整测试明细</h2><table><thead><tr>" + "".join("<th>" + esc(label) + "</th>" for label in labels) + "</tr></thead><tbody>" + body + "</tbody></table>"
    page = page.replace("<h2>压力批次统计</h2>", detail_table + "<h2>压力批次统计</h2>", 1)
    context_failures = boundaries.get("context_failure_summary", {})
    context_note = "；".join(f"总长 {k}: {v[0]}" for k, v in context_failures.items() if v)
    tested_concurrency = boundaries.get("concurrency_levels_tested", [])
    boundary_note = ("<p><strong>边界诊断：</strong>"
                     + ("上下文边界未形成稳定档位，失败原因：" + esc(context_note) + "。" if context_note else "上下文边界档位均未满足稳定判定（需检查请求状态和 usage）。")
                     + ("并发边界已测试档位：" + esc(", ".join(map(str, tested_concurrency))) + "。" if tested_concurrency else "并发边界没有可用批次摘要，已从请求明细重新统计。")
                     + "</p>")
    page = page.replace("<!--BOUNDARY_NOTE-->", boundary_note, 1)
    note = "<p><strong>状态说明：</strong>EXPECTED_TRUNCATED 表示该用例允许触达 token 上限；TRUNCATED 表示非预期截断，应提高 max_tokens 后复测。报告中的‘证据’链接可打开逐请求结果和输出。</p>"
    page = page.replace("</h1>", "</h1>" + note, 1)
    for config_path in (folder / "run_config.json", folder.parent / "run_config.json", folder.parent.parent / "run_config.json"):
        if config_path.exists():
            config_html = "<h2>本次运行配置</h2><pre>" + esc(config_path.read_text(encoding="utf-8")) + "</pre>"
            page = page.replace("<!--RUN_CONFIG-->", config_html, 1)
            break
    page = page.replace("<!--RUN_CONFIG-->", "", 1)
    models_seen = sorted({str(r.get("model", "")) for r in rows if r.get("model")})
    failed_ids = [str(r.get("case_id", "")) for r in rows if r.get("quality") == "FAIL"]
    truncated_ids = [str(r.get("case_id", "")) for r in rows if r.get("status") == "TRUNCATED"]
    summary_html = (
        "<hr><h2>本次测评汇总</h2>"
        f"<p>模型：{esc(', '.join(models_seen))}；实际请求 {sum(r.get('attempt_count', int(r['status'] != 'SKIPPED')) for r in rows)} 次；"
        f"首次 FAIL {sum(r.get('raw_quality', r.get('quality')) == 'FAIL' for r in rows)} 项；"
        f"最终 FAIL {len(failed_ids)} 项：{esc(', '.join(failed_ids) or '无')}。</p>"
        "<p>OK表示接口正常完成，不代表内容质量通过。恢复后的结果与首次结果分别保留；count_5000不采用分段拼接。</p>"
    )
    page += summary_html
    (folder / "report.html").write_text(page, encoding="utf-8")


def main():
    parser = argparse.ArgumentParser(description="屹唐模型功能/压力/对比自动测试")
    parser.add_argument("--config", type=Path, default=HERE / "config.json")
    parser.add_argument("--mode", choices=["all", "full", "functional", "stress", "smoke", "extended", "capacity", "limits"], default="all")
    model_group = parser.add_mutually_exclusive_group()
    model_group.add_argument("--models", nargs="+", help="选择一个或多个模型，例如 --models deepseek14b")
    model_group.add_argument("--model", help="选择单个模型，例如 --model deepseek14b")
    parser.add_argument("--max-tokens", type=int, help="覆盖配置中的 max_tokens 和 5000 字长输出上限")
    parser.add_argument("--timeout", type=int, dest="request_timeout_seconds", help="覆盖单次请求总超时（秒）")
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--data-dir", type=Path, help="业务文档目录；相对路径基于工具根目录")
    parser.add_argument("--allow-no-documents", action="store_true", help="明确允许功能/完整测试缺少业务文档")
    args = parser.parse_args()
    cfg = json.loads(args.config.read_text(encoding="utf-8-sig"))
    if args.data_dir is not None:
        cfg["data_dir"] = str(args.data_dir)
    if args.max_tokens is not None:
        if args.max_tokens < 1:
            parser.error("--max-tokens 必须为正整数")
        cfg["max_tokens"] = args.max_tokens
        cfg["count_5000_max_tokens"] = args.max_tokens
    if args.request_timeout_seconds is not None:
        if args.request_timeout_seconds < 1:
            parser.error("--timeout 必须为正整数")
        cfg["request_timeout_seconds"] = args.request_timeout_seconds
    models = args.models or ([args.model] if args.model else cfg["models"])
    if not models or len(set(models)) != len(models):
        parser.error("模型列表为空或重复")
    print("本次选择模型：" + ", ".join(models), flush=True)
    if cfg["stress_rounds"] < 1 or any(not isinstance(n,int) or not 1 <= n <= 32 for n in cfg["concurrency_levels"]+cfg["extended_concurrency_levels"]):
        parser.error("并发必须为1到32的整数，轮数至少1")
    for key in ("connect_timeout_seconds", "idle_timeout_seconds", "request_timeout_seconds", "max_tokens", "count_5000_max_tokens"):
        if cfg[key] <= 0:
            parser.error(key + "必须为正数")
    server_limits = cfg.get("server_limits", {})
    max_iter_times = server_limits.get("max_iter_times")
    if max_iter_times and any(level > max_iter_times for level in cfg.get("capacity_token_levels", [])):
        parser.error("capacity_token_levels cannot exceed server_limits.max_iter_times")
    try:
        cfg["data_dir"] = str(resolve_data_dir(cfg))
        cases, sources = build_cases(cfg) if args.mode not in ("capacity", "limits") else ([], [])
    except ValueError as exc:
        parser.error(str(exc))
    if args.mode in ("all", "full", "functional", "stress", "extended") and not sources and not args.allow_no_documents:
        parser.error("未加载业务文档，不能生成完整测评。请用 --data-dir 指定资料目录；仅测基础能力可加 --allow-no-documents")
    print(f"资料目录：{cfg['data_dir']}；业务文档：{len(sources)}", flush=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    base = ROOT / "模型测评结果" / stamp
    base.mkdir(parents=True)
    if args.mode == "capacity":
        cases = build_capacity_cases(cfg)
        sources = []
    elif args.mode == "limits":
        cases = build_capacity_cases(cfg) + build_limit_cases(cfg)
        sources = []
    elif args.mode in ("all", "full"):
        cases.extend(build_capacity_cases(cfg))
        cases.extend(build_limit_cases(cfg))
    save_json(base / "cases.json", cases)
    save_json(base / "sources.json", sources)
    save_json(base / "run_config.json", dict(config=cfg, mode=args.mode, models=models, dry_run=args.dry_run))
    print(f"用例 {len(cases)}；结果目录 {base}", flush=True)
    all_rows, all_batches = [], []
    interrupted = False
    try:
        for model in models:
            safe = re.sub(r'[^\w.\-]', '_', model)
            if safe in (".", ".."):
                parser.error("模型名称不合法")
            folder = base / safe
            folder.mkdir(parents=True)
            rows, batches = [], []
            report(folder, rows, model_report_title([model]), batches)
            probe = dict(id="connectivity", group="连通性", prompt="请用 Python 写一段快速排序代码", source="用户已验证的快速排序请求；连通性探测", rule="manual", max_tokens=cfg["max_tokens"])
            active = [c for c in cases if c["id"] in ("python_sort", "doc_1_qa")] if args.mode == "smoke" else cases
            try:
                if args.dry_run:
                    rows.extend(skipped(model,c,"仅生成计划，未发送请求") for c in active)
                    continue
                print(f"[{model}] 检查接口…", flush=True)
                check = request_model(cfg, model, probe, folder / "requests", "connectivity", timeout=cfg["request_timeout_seconds"])
                rows.append(check)
                # A short probe may spend its entire budget on reasoning. A valid
                # completed SSE response still proves transport/model availability.
                available = (check["status"] in ("OK", "TRUNCATED") or
                             (check["http_status"] == 200 and (check["done_received"] or not cfg.get("stream", True)) and
                              check["finish_reason"] in ("stop", "length") and check["output_chars"] > 0))
                if not available:
                    rows.extend(skipped(model,c,"连通性探测失败："+check["error"]) for c in active)
                    rows.append(skipped(model, dict(id="stress_not_run",group="压力",source="历史并发场景"), "连通性失败，压力测试未启动"))
                    print(f"[{model}] {check['status']}: {check['error']}", flush=True)
                    continue
                if args.mode in ("all", "full", "functional", "smoke", "extended", "capacity", "limits"):
                    for i, case in enumerate(active,1):
                        row = skipped(model,case,case["skip"]) if case.get("skip") else request_model(cfg,model,case,folder / "requests",case["id"])
                        rows.append(row)
                        report(folder, rows, model_report_title([model]), batches)
                        print(f"[{model}] {i}/{len(active)} {case['id']}: {row['status']} / {row['quality']}", flush=True)
                if args.mode in ("all", "full", "stress", "extended"):
                    levels = cfg["concurrency_levels"] + (cfg["extended_concurrency_levels"] if args.mode == "extended" else [])
                    for kind in ("qa", "summary", "review", "diff"):
                        pool = [c for c in cases if c["id"].endswith("_"+kind) and not c.get("skip") and c.get("input_document_chars",0)<=6000] if kind != "diff" else [c for c in cases if c["group"]=="差异分析"][:3]
                        if kind != "diff":
                            pool.extend(c for c in cases if c["id"].endswith("_"+kind) and not c.get("skip") and c.get("input_document_chars", 0) > cfg.get("stress_max_document_chars", 6000))
                        else:
                            pool.extend(c for c in cases if c.get("id", "").startswith("diff_") and c not in pool)
                        if not pool:
                            continue
                        for level in levels:
                            for rep in range(cfg["stress_rounds"]):
                                batch_id = f"stress_{kind}_c{level}_r{rep+1}"
                                t0 = time.perf_counter()
                                with concurrent.futures.ThreadPoolExecutor(max_workers=level) as executor:
                                    futures = [executor.submit(request_model,cfg,model,{**pool[j%len(pool)],"id":batch_id+f"_{j+1}","group":"压力"},folder / "requests",batch_id+f"_{j+1}") for j in range(level)]
                                    batch = [f.result() for f in concurrent.futures.as_completed(futures)]
                                wall = time.perf_counter()-t0
                                ok = [r for r in batch if r["status"]=="OK"]
                                batches.append(dict(model=model,batch=batch_id,concurrency=level,requests=len(batch),successful=len(ok),success_rate=len(ok)/len(batch),wall_seconds=wall,requests_per_second=len(ok)/wall,p50_seconds=percentile([r["total_s"] for r in ok],.5),p95_seconds=percentile([r["total_s"] for r in ok],.95)))
                                rows.extend(batch)
                                report(folder, rows, model_report_title([model]), batches)
                                print(f"[{model}] {batch_id}: {len(ok)}/{len(batch)} 正常完成", flush=True)
                                if not ok:
                                    rows.append(skipped(model,dict(id=batch_id+"_stop",group="压力",source="保护性停止"),"此压力场景整批无成功，停止升高该场景并发"))
                                    break
                            if not ok:
                                break
                if args.mode in ("all", "full", "limits"):
                    probe, limit_levels = build_concurrency_probe(cfg)
                    for level in limit_levels:
                        for repeat in range(1, cfg.get("limit_repeats", 3) + 1):
                            batch_id = f"limit_concurrency_c{level}_r{repeat}"
                            t0 = time.perf_counter()
                            with concurrent.futures.ThreadPoolExecutor(max_workers=level) as executor:
                                futures = [executor.submit(request_model, cfg, model,
                                            {**probe, "id": f"{batch_id}_{j+1}"},
                                            folder / "requests", f"{batch_id}_{j+1}")
                                            for j in range(level)]
                                batch = [f.result() for f in concurrent.futures.as_completed(futures)]
                            wall = time.perf_counter() - t0
                            # Boundary stability is based on complete, exact replies;
                            # truncated responses remain visible but do not count as successful.
                            ok = [r for r in batch if r.get("status") == "OK" and r.get("quality") == "PASS"]
                            batches.append(dict(model=model, batch=batch_id, concurrency=level,
                                                requests=len(batch), successful=len(ok),
                                                success_rate=len(ok) / len(batch), wall_seconds=wall,
                                                requests_per_second=len(ok) / wall if wall else None,
                                                p50_seconds=percentile([r["total_s"] for r in ok], .5),
                                                p95_seconds=percentile([r["total_s"] for r in ok], .95)))
                            rows.extend(batch)
                            report(folder, rows, model_report_title([model]), batches)
                            print(f"[{model}] {batch_id}: {len(ok)}/{len(batch)} 稳定性请求", flush=True)
            finally:
                report(folder, rows, model_report_title([model]), batches)
                all_rows.extend(rows)
                all_batches.extend(batches)
                report(base, all_rows, model_report_title(models), all_batches)
    except KeyboardInterrupt:
        interrupted = True
        print("已中断，已完成请求和部分输出已保存。", flush=True)
    finally:
        report(ROOT / "模型测评结果", all_rows, model_report_title(models), all_batches, evidence_root=base)
        print(f"报告：{base / 'report.html'}", flush=True)
    fixed_report = ROOT / "模型测评结果" / "report.html"
    print(f"固定汇总报告：{fixed_report}", flush=True)
    if interrupted:
        return 130
    return 2 if not args.dry_run and any(r["status"] in ("ERROR","TIMEOUT","TRUNCATED") or r.get("quality")=="FAIL" for r in all_rows) else 0


if __name__ == "__main__":
    raise SystemExit(main())
