#!/usr/bin/env python3
"""Conservative local structural validation of ONE homework delivery.
Does not assess answer accuracy, screenshot authenticity, or DOCX visual layout.
"""
import argparse
import json
import re
import sys
import zipfile
from pathlib import Path

PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--task", required=True, type=Path)
    p.add_argument("--report", required=True, type=Path)
    p.add_argument("--output-dir", required=True, type=Path)
    p.add_argument("--json-out", type=Path)
    a = p.parse_args()
    checks = []
    def check(name, condition, detail=""):
        checks.append({"name": name, "pass": bool(condition), "detail": str(detail)})
    check("TASK.md exists", a.task.is_file(), a.task)
    check("Delivery dir exists", a.output_dir.is_dir(), a.output_dir)
    check("Report exists", a.report.is_file(), a.report)
    if a.report.is_file():
        good_docx = a.report.suffix.lower() == ".docx" and zipfile.is_zipfile(a.report)
        check("Valid DOCX container", good_docx)
        if good_docx:
            with zipfile.ZipFile(a.report) as z:
                names = set(z.namelist())
                check("Word document content", "word/document.xml" in names)
                image_count = len([n for n in names if n.startswith("word/media/") and not n.endswith("/")])
                checks.append({"name": "Embedded images (informational)", "pass": True, "detail": image_count})
    if a.task.is_file():
        task_text = a.task.read_text(encoding="utf-8")
        # Only explicit Markdown checkboxes are treated as verification-gate items.
        pending = re.findall(r"^\s*[-*]\s*\[ \]\s*(.*)$", task_text, flags=re.M)
        check("No unchecked acceptance boxes", len(pending) == 0, f"pending={len(pending)}")
        check("Task has source page references", bool(re.search(r"页码", task_text)))
    pngs = list(a.output_dir.rglob("*.png")) if a.output_dir.exists() else []
    invalid_pngs = []
    for f in pngs:
        try:
            with f.open("rb") as handle:
                if handle.read(8) != PNG_SIGNATURE:
                    invalid_pngs.append(str(f))
        except OSError:
            invalid_pngs.append(str(f))
    check("All PNG headers valid", not invalid_pngs, invalid_pngs)
    checks.append({"name": "Found PNGs (informational; not evidence of authenticity)", "pass": True, "detail": len(pngs)})
    report = {"task": str(a.task), "report": str(a.report), "output_dir": str(a.output_dir),
              "checks": checks, "all_structural_checks_passed": all(c["pass"] for c in checks),
              "warning": "结构检查不能代替任务完整性、实验真实性、答案准确性或 DOCX 渲染逐页验收。"}
    data = json.dumps(report, ensure_ascii=False, indent=2)
    if a.json_out:
        a.json_out.parent.mkdir(parents=True, exist_ok=True)
        a.json_out.write_text(data, encoding="utf-8")
    print(data)
    sys.exit(0 if report["all_structural_checks_passed"] else 1)

if __name__ == "__main__":
    main()
