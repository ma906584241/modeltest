"""Deterministic validation, token budgets and bounded recovery policies."""
from collections import Counter
from functools import lru_cache
import math
import re


class PreflightError(ValueError):
    def __init__(self, message, category="CONFIG_LIMIT"):
        super().__init__(message)
        self.category = category


@lru_cache(maxsize=8)
def load_tokenizer(path):
    try:
        from transformers import AutoTokenizer
        return AutoTokenizer.from_pretrained(path, local_files_only=True, trust_remote_code=False)
    except Exception as exc:
        raise PreflightError(f"无法加载真实 tokenizer {path!r}：{exc}") from exc


def preflight(cfg, model, case, payload):
    if case.get("data_error"):
        raise PreflightError(case["data_error"], "DATA_QUALITY")
    document = case.get("source_text")
    if document is not None and (not document.strip() or "\ufffd" in document or "\x00" in document):
        raise PreflightError("原始文档为空、含替换字符或 NUL，需检查编码和转换结果", "DATA_QUALITY")
    path = cfg.get("tokenizer_paths", {}).get(model)
    tokenizer = load_tokenizer(path) if path else None
    exact_required = case.get("id") == "count_5000"
    if exact_required and tokenizer is None:
        raise PreflightError(f"count_5000 必须配置 tokenizer_paths.{model} 为服务模型对应的本地 tokenizer 目录")
    if tokenizer:
        try:
            prompt_tokens = len(tokenizer.apply_chat_template(payload["messages"], tokenize=True, add_generation_prompt=True))
        except Exception as exc:
            raise PreflightError(f"无法用模型 chat template 计算输入 token：{exc}") from exc
        method = "model_tokenizer_chat_template"
    else:
        # Explicit conservative fallback, never advertised as measured tokens.
        prompt_tokens = sum(len(m["content"].encode("utf-8")) + 32 for m in payload["messages"]) + 64
        method = "utf8_bytes_plus_template_margin_estimate"
    standard_tokens = None
    if exact_required:
        standard_tokens = len(tokenizer.encode("\n".join(f"{i} 测试" for i in range(5001)), add_special_tokens=False))
        required = math.ceil(standard_tokens * 1.1)
        if required > payload["max_tokens"]:
            raise PreflightError(f"标准答案 {standard_tokens} token，含10%余量需 {required}，超过用例额度 {payload['max_tokens']}")
        payload["max_tokens"] = required
        if cfg.get("supports_min_new_tokens", False):
            payload["min_new_tokens"] = standard_tokens
    limits = cfg.get("server_limits", {})
    maximum = payload["max_tokens"]
    margin = cfg.get("safety_margin", 1024)
    if not isinstance(maximum, int) or maximum < 1 or margin < 0:
        raise PreflightError("max_tokens 必须为正整数，safety_margin 不得为负数")
    if maximum > limits.get("max_iter_times", 32768):
        raise PreflightError(f"请求输出 {maximum} 超过服务端 max_iter_times={limits.get('max_iter_times', 32768)}")
    if prompt_tokens + maximum + margin > limits.get("max_seq_len", 65536):
        raise PreflightError(f"输入 {prompt_tokens} + 输出 {maximum} + 余量 {margin} 超过 max_seq_len={limits.get('max_seq_len', 65536)} ({method})")
    if prompt_tokens > limits.get("max_input_token_len", 32768):
        raise PreflightError(f"输入 {prompt_tokens} 超过 max_input_token_len ({method})")
    if payload.get("min_new_tokens", 0) > maximum:
        raise PreflightError("min_new_tokens 超过 max_tokens")
    return dict(prompt_tokens_budget=prompt_tokens, token_budget_method=method,
                standard_answer_tokens=standard_tokens, safety_margin=margin)


def chinese_number(value):
    if value.isdigit():
        return str(int(value))
    digits = dict(zip("零一二三四五六七八九", range(10)))
    total = current = 0
    for char in value:
        if char in digits:
            current = digits[char]
        elif char in "十百千":
            total += (current or 1) * {"十": 10, "百": 100, "千": 1000}[char]
            current = 0
    return str(total + current)


def numbering(text):
    """Extract ordered heading/clause keys, including Chinese chapter labels."""
    keys = []
    for line in text.splitlines():
        line = re.sub(r"^\s*(?:#{1,6}\s*|\*\*)?", "", line)
        match = re.match(r"第([零一二三四五六七八九十百千\d]+)(章|节|条)", line)
        if match:
            keys.append(({"章": "chapter", "节": "section", "条": "article"}[match[2]], chinese_number(match[1])))
            continue
        match = re.match(r"(?i)(chapter|section|article)\s+(\d+)\b", line)
        if match:
            keys.append((match[1].lower(), str(int(match[2]))))
            continue
        match = re.match(r"[（(]([零一二三四五六七八九十百千\d]+)[)）]", line)
        if match:
            keys.append(("item", chinese_number(match[1])))
            continue
        match = re.match(r"(\d+(?:\.\d+)*)(?:[、.)．]|\s)", line)
        if match:
            keys.append(("item", match[1]))
            continue
        match = re.match(r"([一二三四五六七八九十百]+)、", line)
        if match:
            keys.append(("item", chinese_number(match[1])))
    return keys


def validate(case, answer, finish):
    rule = case.get("rule")
    if rule == "count":
        numbers = [int(x) for x in re.findall(r"(?<!\d)(\d+)\s*测试", answer)]
        if case.get("end") == 5000:
            pure = all(re.fullmatch(r"\d+ 测试", line) for line in answer.splitlines())
        else:
            pure = re.fullmatch(r"\d+\s*测试(?:(?:\s+|[，,]\s*)\d+\s*测试)*", answer) is not None
        ok = pure and finish == "stop" and numbers == list(range(case["end"] + 1))
        return ("PASS" if ok else "FAIL", f"识别{len(numbers)}项；要求{case['end']+1}项、纯正文、严格连续且finish_reason=stop")
    if rule == "repeat":
        compact = re.sub(r"\s+", "", answer)
        count, rest = divmod(len(compact), 4)
        ok = count > 0 and compact == "测试输出" * count + "测试输出"[:rest]
        ok = ok and (finish == "length" or (rest == 0 and finish == "stop"))
        return ("PASS" if ok else "FAIL", "忽略空白，仅允许完整重复；length允许末尾合法前缀")
    if rule == "speech":
        headings = list(re.finditer(r"(?m)^\s*(?:#{1,6}\s*)?第([一二三四五六七八12345678]+)章[^\n]*\n", answer))
        ids = [chinese_number(m[1]) for m in headings]
        sections = [answer[m.end():headings[i+1].start() if i+1 < len(headings) else len(answer)] for i, m in enumerate(headings)]
        counts = [len(re.findall(r"[\u4e00-\u9fff]", s)) for s in sections]
        n = sum(counts)
        sentences = [re.sub(r"\s", "", s) for s in re.split(r"[。！？\n]", "\n".join(sections)) if len(s.strip()) >= 20]
        duplicates = sum((v-1)*len(k) for k, v in Counter(sentences).items())
        refusal = re.search(r"无法(?:直接)?(?:完成|提供|撰写)|不能(?:为您|为你)?(?:提供|撰写)|以下(?:是|为).{0,8}提纲|仅供参考的提纲", answer)
        repeated_run = re.search(r"(.{12,80})\1{3,}", re.sub(r"\s+", "", "\n".join(sections)))
        ok = (finish == "stop" and ids == list(map(str, range(1, 9))) and n >= case.get("minimum", 4000)
              and all(c >= 450 for c in counts) and not refusal and not repeated_run and duplicates < max(1, n * .15))
        return ("PASS" if ok else "FAIL", f"正文汉字{n}；章节{ids}；各章字数{counts}；重复句字符{duplicates}；拒绝/提纲={bool(refusal)}；语义质量需人工复核")
    if rule == "translate":
        required = numbering(case.get("source_text", ""))
        actual = numbering(answer)
        remaining = answer
        for name in case.get("allowed_chinese_names", []):
            remaining = remaining.replace(name, "")
        chinese = len(re.findall(r"[\u4e00-\u9fff]", remaining))
        ending = re.search(r'[.!?][\s\"\u201d\u2019\)\]]*$', answer) is not None
        ok = finish == "stop" and bool(answer) and required == actual and chinese == 0 and ending
        if ok and not required:
            return "REVIEW", "无可校验的原文章节/条款编号；语言和句尾检查通过，全文完整性需人工确认"
        return ("PASS" if ok else "FAIL", f"编号顺序完整={required == actual}；应有{len(required)}项/实际{len(actual)}项；中文残留{chinese}；完整句尾={ending}；仅验证结构和语言，译意需人工复核")
    return None


def translation_chunks(text, size=4000):
    """Keep paragraphs intact where possible; never drop any source characters."""
    chunks, current = [], ""
    for paragraph in text.splitlines(keepends=True):
        if current and len(current) + len(paragraph) > size:
            chunks.append(current)
            current = ""
        while len(paragraph) > size:
            boundary = max(paragraph.rfind("。", 0, size), paragraph.rfind("；", 0, size)) + 1
            boundary = boundary if boundary >= size // 2 else size
            chunks.append(paragraph[:boundary])
            paragraph = paragraph[boundary:]
        current += paragraph
    if current:
        chunks.append(current)
    return chunks


def classify(result):
    if result.get("failure_type"):
        return
    if result["status"] in ("ERROR", "TIMEOUT"):
        category = "NETWORK_ERROR"
    elif result.get("quality") == "FAIL":
        category = "CONFIG_LIMIT" if result.get("finish_reason") == "length" else "MODEL_BEHAVIOR"
    else:
        category = ""
    result.update(failure_type=category, failure_reason=(result.get("error") or result.get("quality_note", "")) if category else "")
