#!/usr/bin/env python3
"""Index ONE explicitly selected lab PDF; writes page text and lightweight summary.
Uses PyMuPDF when available, otherwise pypdf.
Does not OCR, guess tasks, or read neighboring files.
"""
import argparse
from contextlib import contextmanager
import json
import re
from pathlib import Path

HEADINGS = ["实验目的", "课前预习与基础知识点", "实验环境", "实验原理", "实验步骤", "实验作业"]
ASSIGNMENT = re.compile(r"(?m)^[ \t]*(?:【\s*实验作业\s*】|实验作业[ \t]*[:：]?[ \t]*(?=$))")
PREVIEW_LIMIT = 1600


@contextmanager
def open_pages(pdf):
    try:
        import fitz
    except ImportError:
        from pypdf import PdfReader

        with pdf.open("rb") as source:
            yield PdfReader(source).pages
    else:
        with fitz.open(pdf) as doc:
            yield doc


def page_text(page):
    if hasattr(page, "get_text"):
        return page.get_text("text", sort=True)
    return page.extract_text() or ""


def image_count(page):
    if hasattr(page, "get_images"):
        return len(page.get_images(full=True))
    return len(page.images)


def selected_pages(spec, total):
    pages = set()
    for part in spec.split(","):
        match = re.fullmatch(r"(\d+)(?:\s*-\s*(\d+))?", part.strip())
        if not match:
            raise ValueError(f"无效页码或范围：{part}")
        start, end = int(match[1]), int(match[2] or match[1])
        if not 1 <= start <= end <= total:
            raise ValueError(f"页码须在 1–{total} 内：{part}")
        pages.update(range(start, end + 1))
    return sorted(pages)


def show_assignment(index, candidates, index_file):
    print(f"已索引 {len(index)} 页：{index_file}")
    for heading in ("实验环境", "实验步骤"):
        pages = [str(page["page"]) for page in index if heading in page["headings"]]
        print(f"{heading}候选页：{', '.join(pages) if pages else '无'}")
    if not candidates:
        pages = [str(page["page"]) for page in index if "实验作业" in page["headings"]]
        print(f"未识别到独立的实验作业标题。宽松候选页：{', '.join(pages) if pages else '无'}。请检查目录和页面图像。")
        return
    print("实验作业候选预览。每个标题页及紧接的续页各最多 1600 字；请核对跨页及图片内容。")
    for number, start in candidates:
        end = min(number + 1, len(index))
        for page in index[number - 1:end]:
            content = Path(page["text_file"]).read_text(encoding="utf-8")
            if page["page"] == number:
                content = content[start:]
            print(f"--- PDF 第 {page['page']} 页 ---")
            print(content[:PREVIEW_LIMIT].rstrip())
            if len(content) > PREVIEW_LIMIT:
                print("…预览截断，使用 --pages 回读整页。")

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--output-dir", type=Path)
    parser.add_argument("--pages", help="只读指定的物理页码，例如 3,7-9；不重建索引")
    args = parser.parse_args()
    if not args.pages and args.output_dir is None:
        parser.error("全量索引需要 --output-dir")
    if not args.pdf.is_file() or args.pdf.suffix.lower() != ".pdf":
        raise SystemExit("PDF not found or wrong extension")
    if args.pages:
        with open_pages(args.pdf) as doc:
            try:
                pages = selected_pages(args.pages, len(doc))
            except ValueError as exc:
                parser.error(str(exc))
            for number in pages:
                print(f"--- PDF 第 {number} 页 ---")
                print(page_text(doc[number - 1]).rstrip())
        return
    args.output_dir.mkdir(parents=True, exist_ok=True)
    pages_dir = args.output_dir / "pdf_pages"
    pages_dir.mkdir(exist_ok=True)
    index = []
    candidates = []
    with open_pages(args.pdf) as doc:
        for i, page in enumerate(doc):
            number = i + 1
            content = page_text(page)
            match = ASSIGNMENT.search(content)
            if match:
                candidates.append((number, match.start()))
            (pages_dir / f"page_{number:03d}.txt").write_text(content, encoding="utf-8")
            hits = [h for h in HEADINGS if h in re.sub(r"\s+", "", content)]
            idx = {
                "page": number,
                "chars": len(content),
                "images": image_count(page),
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
    show_assignment(index, candidates, args.output_dir / "PDF_INDEX.md")

if __name__ == "__main__":
    main()
