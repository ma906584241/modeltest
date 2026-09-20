"""Rebuild an existing report after changing report labels."""
from __future__ import annotations

import json
import sys
from pathlib import Path

from run_tests import case_description, report


def main() -> None:
    if len(sys.argv) != 2:
        raise SystemExit("用法：python rebuild_report.py <报告时间戳目录>")
    folder = Path(sys.argv[1]).resolve()
    cases = {item["id"]: item for item in json.loads((folder / "cases.json").read_text(encoding="utf-8"))}
    rows = json.loads((folder / "results.json").read_text(encoding="utf-8"))
    for row in rows:
        case = cases.get(row.get("case_id"), {"id": row.get("case_id", "")})
        row["case_name"] = case_description(case)
        if row.get("error") == "旧版二进制 DOC 未提取；另存同名 UTF-8 .txt 后自动纳入":
            row["error"] = "旧版 .doc 文件暂未解析。请将其另存为 UTF-8 编码的同名 .txt 文件，下一次测试会自动纳入。"
    batches = json.loads((folder / "stress_summary.json").read_text(encoding="utf-8"))
    models = sorted({str(row.get("model", "")).strip() for row in rows if row.get("model")})
    title = f"{models[0]}模型测试报告" if len(models) == 1 else f"模型测试报告（{', '.join(models) or '未指定模型'}）"
    report(folder, rows, title, batches)
    report(folder.parent, rows, title, batches, evidence_root=folder)
    print(f"已更新报告：{folder / 'report.html'}")


if __name__ == "__main__":
    main()
