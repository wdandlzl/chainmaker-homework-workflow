#!/usr/bin/env python3
"""Read an explicitly supplied .docx template; includes headers/footers and textbox XML text.
Requires: pip install python-docx
This does not replace a visual template inspection in Word/LibreOffice.
"""
import argparse
import re
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

NS_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

def xml_text(blob):
    root = ET.fromstring(blob)
    text = [node.text or "" for node in root.iter(f"{NS_W}t")]
    return " ".join(t for t in text if t.strip())

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--docx", required=True, type=Path)
    p.add_argument("--out", required=True, type=Path)
    a = p.parse_args()
    if not a.docx.is_file() or not zipfile.is_zipfile(a.docx):
        raise SystemExit("Template not found or not a valid DOCX")
    try:
        from docx import Document
    except ImportError as exc:
        raise SystemExit("Missing dependency: pip install python-docx") from exc
    d = Document(a.docx)
    sections = ["# Word 模板扫描", f"源文件：`{a.docx}`", "", "## 段落文本"]
    for i, para in enumerate(d.paragraphs, 1):
        if para.text.strip():
            sections.append(f"- P{i} [{para.style.name}]: {para.text.strip()}")
    sections += ["", f"## 表格（共 {len(d.tables)} 个）"]
    for i, table in enumerate(d.tables, 1):
        sections.append(f"### 表格 {i}")
        for row in table.rows:
            sections.append(" | ".join(c.text.replace("\n", " / ").strip() for c in row.cells))
    sections += ["", "## 全部可见文字候选（含页眉页脚与文字框）"]
    with zipfile.ZipFile(a.docx) as z:
        targets = [n for n in z.namelist() if re.fullmatch(r"word/(?:document|header\d+|footer\d+|footnotes|endnotes)\.xml", n)]
        for name in targets:
            sections.append(f"### {name}")
            sections.append(xml_text(z.read(name)) or "（未检测到文字）")
    sections += ["", "**注意**：该扫描不保证识别形状布局、图片内文字、文件命名要求的语义；必须在 Word/LibreOffice 中视觉核对模板后再填写。"]
    a.out.parent.mkdir(parents=True, exist_ok=True)
    a.out.write_text("\n".join(sections), encoding="utf-8")
    print(f"Template scanned -> {a.out}")

if __name__ == "__main__":
    main()
