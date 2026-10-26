#!/usr/bin/env python3
"""Generate a small sample .docx for trying out the formatter.

Usage: python scripts/make_sample.py [output_dir]
"""
import sys
from pathlib import Path

from docx import Document

OUT = Path(sys.argv[1] if len(sys.argv) > 1 else "sample") / "样例公文.docx"

PARAGRAPHS = [
    "某某市某某局",
    "关于开展2026年度公文格式自查工作的通知",
    "各科室、各下属单位:",
    "为规范公文处理工作,提升文件质量,现就2026年度公文格式自查工作通知如下.",
    "一、自查范围",
    "本次自查覆盖2026年1月至6月印发的各类通知、报告、请示等文件.",
    "（一）重点内容",
    "1.版式要素是否齐全,包括版心尺寸、字体字号、行距设置.",
    "（二）时间安排",
    "（1）6月15日前完成自查,形成书面报告报办公室.",
    "二、工作要求",
    "请各单位高度重视,指定专人负责,确保自查工作取得实效.",
]


def main() -> int:
    doc = Document()
    for text in PARAGRAPHS:
        doc.add_paragraph(text)
    OUT.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(OUT))
    print(f"已生成: {OUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
