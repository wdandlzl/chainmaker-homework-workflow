#!/usr/bin/env python3
"""Index ONE explicitly selected lab PDF; writes page text and lightweight summary.
Requires: pip install pymupdf
Does not OCR, guess tasks, or read neighboring files.
"""
import argparse
import json
import re
from pathlib import Path

HEADINGS = ["实验目的", "课前预习与基础知识点", "实验环境", "实验原理", "实验步骤", "实验作业"]

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    args = parser.parse_args()
    if not args.pdf.is_file() or args.pdf.suffix.lower() != ".pdf":
        raise SystemExit("PDF not found or wrong extension")
    try:
        import fitz
    except ImportError as exc:
        raise SystemExit("Missing dependency: pip install pymupdf") from exc
    args.output_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = args.output_dir / "pdf_pages"
    pages_dir.mkdir(exist_ok=True)
    index = []
    with fitz.open(args.pdf) as doc:
        for i, page in enumerate(doc):
            number = i + 1
            content = page.get_text("text", sort=True)
            (pages_dir / f"page_{number:03d}.txt").write_text(content, encoding="utf-8")
            hits = [h for h in HEADINGS if h in re.sub(r"\s+", "", content)]
            idx = {
                "page": number,
                "chars": len(content),
                "images": len(page.get_images(full=True)),
                "headings": hits,
                "text_file": str(pages_dir / f"page_{number:03d}.txt"),
            }
            index.append(idx)
    payload = {"pdf": str(args.pdf.resolve()), "pages": index,
               "note": "页面命中仅为候选；【实验作业】可能跨页或藏于图片，须核实。"}
    (args.output_dir / "pdf_index.json").write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    lines = [f"# PDF 页码索引：{args.pdf.name}", "", "| 页码 | 文本字符 | 图片数 | 命中标题 |", "|---:|---:|---:|---|"]
    for p in index:
        lines.append(f"| {p['page']} | {p['chars']} | {p['images']} | {', '.join(p['headings']) or '-'} |")
    lines += ["", "注意：请核对命中【实验作业】的页面及其后续页；如果文本不足，请渲染页面检查真实内容。"]
    (args.output_dir / "PDF_INDEX.md").write_text("\n".join(lines), encoding="utf-8")
    print(f"Indexed {len(index)} pages -> {args.output_dir / 'PDF_INDEX.md'}")

if __name__ == "__main__":
    main()
