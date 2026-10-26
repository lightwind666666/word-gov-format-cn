#!/usr/bin/env python3
"""中文公文版式排版（.docx）。

安全默认值
----------
- 不覆盖已存在的输出文件，需要显式 --force
- --in-place 前自动生成带时间戳的备份
- 支持 --dry-run 预演，不写任何文件
- 标点规范化只在含中文的段落生效，URL / 邮箱 / 英文缩写 / 版本号 / 时间会被保护
- 按 run 改写文本，尽量保留行内格式与图片
"""
import argparse
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    from docx import Document
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.opc.exceptions import PackageNotFoundError
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except Exception:
    print("缺少依赖：python-docx，请先安装：pip install python-docx", file=sys.stderr)
    raise


LEVEL1_RE = re.compile(r"^\s*[一二三四五六七八九十百]+、")
LEVEL2_RE = re.compile(r"^\s*（[一二三四五六七八九十百]+）")
LEVEL3_RE = re.compile(r"^\s*[0-9]+[.]")
LEVEL4_RE = re.compile(r"^\s*（[0-9]+）")
LEVEL3_PREFIX_RE = re.compile(r"^\s*[0-9]+[.]")

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff]")
_CJK_CHARS = "\u3000，。；：！？、（）【】“”‘’\u4e00-\u9fff"

# 受保护的“非中文标点语境”片段：URL、邮箱、点分英文串（域名/e.g./v1.2.3/文件名）、时间
PROTECT_RE = re.compile(
    r"(?:(?:https?|ftp)://|www\.)[^\s" + _CJK_CHARS + r"]+"
    r"|[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}"
    r"|[A-Za-z0-9][A-Za-z0-9+&*/%_#~=-]*(?:\.[A-Za-z0-9]+)+"
    r"|\d{1,2}:\d{2}(?::\d{2})?"
)

FULLWIDTH_MAP = {
    ",": "，",
    ";": "；",
    ":": "：",
    "!": "！",
    "?": "？",
    "(": "（",
    ")": "）",
    "[": "【",
    "]": "】",
}

# 改写 run.text 会重建 run 内容，这些元素必须原样保留
COMPLEX_RUN_TAGS = (
    "w:drawing",
    "w:pict",
    "w:object",
    "w:tab",
    "w:br",
    "w:cr",
    "w:fldChar",
    "w:instrText",
)


def _ascii_alnum(ch: str) -> bool:
    return bool(ch) and ch.isascii() and ch.isalnum()


def normalize_punctuation(text: str, state: dict) -> str:
    """把中文语境的半角标点转全角，英文/URL/邮箱语境保持原样。"""
    out = []
    i, n = 0, len(text)
    while i < n:
        m = PROTECT_RE.match(text, i)
        if m and m.end() > i:
            out.append(m.group(0))
            i = m.end()
            continue

        ch = text[i]
        prev = text[i - 1] if i > 0 else ""
        nxt = text[i + 1] if i + 1 < n else ""

        if ch == ".":
            if _ascii_alnum(prev) or _ascii_alnum(nxt):
                out.append(".")
            else:
                out.append("。")
        elif ch in ",;:":
            if _ascii_alnum(prev) or _ascii_alnum(nxt):
                out.append(ch)
            else:
                out.append(FULLWIDTH_MAP[ch])
        elif ch in "!?":
            if CJK_RE.match(prev) or CJK_RE.match(nxt):
                out.append(FULLWIDTH_MAP[ch])
            else:
                out.append(ch)
        elif ch == '"':
            if CJK_RE.match(prev) or CJK_RE.match(nxt):
                out.append("\u201c" if state["dq"] else "\u201d")
                state["dq"] = not state["dq"]
            else:
                out.append(ch)
        elif ch == "'":
            if _ascii_alnum(prev) and _ascii_alnum(nxt):
                out.append("'")          # 英文撇号，如 don't
            elif CJK_RE.match(prev) or CJK_RE.match(nxt):
                out.append("\u2018" if state["sq"] else "\u2019")
                state["sq"] = not state["sq"]
            else:
                out.append(ch)
        else:
            out.append(FULLWIDTH_MAP.get(ch, ch))
        i += 1
    return "".join(out)


def set_run_font(run, cn_font: str, en_font: str, size_pt: float, bold=None, color=None):
    """bold=None 表示保留原有加粗状态（正文行内强调不被清除）。"""
    run.font.name = en_font
    run._element.rPr.rFonts.set(qn("w:eastAsia"), cn_font)
    run.font.size = Pt(size_pt)
    if bold is not None:
        run.bold = bold
    if color is not None:
        run.font.color.rgb = color


def set_paragraph_spacing(p, line_pt: float, first_line_indent_pt: float = 0):
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.EXACTLY
    p.paragraph_format.line_spacing = Pt(line_pt)
    p.paragraph_format.first_line_indent = Pt(first_line_indent_pt)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)


def classify_paragraph(text: str) -> str:
    """按原始文本判定层级，必须在标点规范化之前调用。"""
    if LEVEL1_RE.match(text):
        return "level1"
    if LEVEL2_RE.match(text):
        return "level2"
    if LEVEL3_RE.match(text):
        return "level3"
    if LEVEL4_RE.match(text):
        return "level4"
    return "body"


def run_has_complex_content(run) -> bool:
    r = run._r
    return any(r.findall(qn(tag)) for tag in COMPLEX_RUN_TAGS)


def rewrite_paragraph_text(p, style: str, normalize: bool):
    """逐 run 规范化标点，避免重建 run 造成图片/行内格式丢失。"""
    has_cjk = bool(CJK_RE.search(p.text))
    if not (normalize and has_cjk):
        return
    state = {"dq": True, "sq": True}
    first_text_run = True
    for run in p.runs:
        if run_has_complex_content(run):
            continue
        text = run.text
        if not text:
            continue
        prefix = ""
        if style == "level3" and first_text_run:
            m = LEVEL3_PREFIX_RE.match(text)
            if m:
                prefix = m.group(0)      # 保护 “1.” 这类编号，避免被转成 “1。”
        run.text = prefix + normalize_punctuation(text[len(prefix):], state)
        if text.strip():
            first_text_run = False


def format_paragraph(p, en_font: str, is_issuer: bool = False, is_title: bool = False, normalize: bool = True):
    raw = p.text
    if not raw.strip():
        return

    style = classify_paragraph(raw)
    rewrite_paragraph_text(p, style, normalize)

    if is_issuer:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_line = 32 if len(p.text.strip()) > 28 else 38
        set_paragraph_spacing(p, line_pt=title_line, first_line_indent_pt=0)
        for r in p.runs:
            set_run_font(
                r,
                cn_font="方正小标宋简体",
                en_font=en_font,
                size_pt=22,
                bold=False,
                color=RGBColor(0xFF, 0x00, 0x00),
            )
        return

    if is_title:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_line = 32 if len(p.text.strip()) > 28 else 38
        set_paragraph_spacing(p, line_pt=title_line, first_line_indent_pt=0)
        for r in p.runs:
            set_run_font(r, "方正小标宋简体", en_font, 22, bold=False)
        return

    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if style == "level1":
        set_paragraph_spacing(p, line_pt=28, first_line_indent_pt=0)
        for r in p.runs:
            set_run_font(r, "黑体", en_font, 16, bold=False)
    elif style == "level2":
        set_paragraph_spacing(p, line_pt=28, first_line_indent_pt=0)
        for r in p.runs:
            set_run_font(r, "楷体_GB2312", en_font, 16, bold=False)
    elif style in ("level3", "level4"):
        set_paragraph_spacing(p, line_pt=28, first_line_indent_pt=0)
        for r in p.runs:
            set_run_font(r, "仿宋_GB2312", en_font, 16, bold=False)
    else:
        # 正文：仿宋_GB2312，三号，首行缩进 2 字
        set_paragraph_spacing(p, line_pt=28, first_line_indent_pt=32)
        for r in p.runs:
            set_run_font(r, "仿宋_GB2312", en_font, 16)


ISSUER_SUFFIXES = ("局", "部", "委", "办", "中心", "处", "室", "会", "厅")


def _looks_like_issuer(text: str) -> bool:
    return (
        len(text) <= 30
        and text.endswith(ISSUER_SUFFIXES)
        and "\n" not in text
        and text not in ("", )
    )


def detect_roles(doc):
    """返回 (issuer_index, title_index)；启发式判定，可能判错。"""
    non_empty = [(idx, p) for idx, p in enumerate(doc.paragraphs) if p.text.strip()]
    issuer_index = None
    title_index = None
    if not non_empty:
        return None, None

    first_idx, first_p = non_empty[0]
    if _looks_like_issuer(first_p.text.strip()):
        issuer_index = first_idx
        if len(non_empty) >= 2:
            title_index = non_empty[1][0]
    else:
        title_index = first_idx

    if issuer_index is None and len(non_empty) >= 2:
        second_idx, second_p = non_empty[1]
        if _looks_like_issuer(second_p.text.strip()):
            issuer_index = second_idx
            title_index = non_empty[0][0]

    return issuer_index, title_index


def format_document(file_path: Path, out_path: Path, en_font: str, normalize: bool = True):
    doc = Document(str(file_path))

    # 页面设置（A4 + 指定边距，对应版心 156mm x 225mm）
    for section in doc.sections:
        section.page_width = Cm(21)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(3.7)
        section.bottom_margin = Cm(3.5)
        section.left_margin = Cm(2.8)
        section.right_margin = Cm(2.6)

    issuer_index, title_index = detect_roles(doc)

    for idx, p in enumerate(doc.paragraphs):
        format_paragraph(
            p,
            en_font=en_font,
            is_issuer=(issuer_index is not None and idx == issuer_index),
            is_title=(title_index is not None and idx == title_index),
            normalize=normalize,
        )

    # 表格内容按正文处理
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    format_paragraph(p, en_font=en_font, is_issuer=False, is_title=False, normalize=normalize)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out_path))


def iter_docx(path: Path):
    if path.is_file() and path.suffix.lower() == ".docx":
        yield path
        return
    if path.is_dir():
        for p in sorted(path.glob("*.docx")):
            if p.name.startswith("~$"):
                continue
            yield p


def build_output_path(src: Path, output: Path | None, in_place: bool) -> Path:
    if in_place:
        return src
    if output is None:
        return src.with_name(f"{src.stem}_排版.docx")
    if output.suffix.lower() == ".docx":
        return output
    return output / f"{src.stem}_排版.docx"


def backup_path(target: Path) -> Path:
    stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    candidate = target.with_name(f"{target.name}.bak-{stamp}")
    n = 1
    while candidate.exists():
        candidate = target.with_name(f"{target.name}.bak-{stamp}-{n}")
        n += 1
    return candidate


def main():
    parser = argparse.ArgumentParser(description="中文公文版式排版（.docx）")
    parser.add_argument("input", help="输入文件或目录（仅 .docx）")
    parser.add_argument("-o", "--output", help="输出文件或输出目录（可选）")
    parser.add_argument("--in-place", action="store_true", help="原地覆盖（自动生成 .bak 备份）")
    parser.add_argument("--force", action="store_true", help="允许覆盖已存在的输出文件")
    parser.add_argument("--no-punct-normalize", action="store_true", help="不做标点全角化")
    parser.add_argument("--dry-run", action="store_true", help="只打印计划，不写任何文件")
    parser.add_argument("--en-font", default="Times New Roman", help="阿拉伯数字/英文默认字体")
    args = parser.parse_args()

    in_path = Path(args.input).expanduser().resolve()
    out_arg = Path(args.output).expanduser().resolve() if args.output else None

    if not in_path.exists():
        print(f"输入不存在：{in_path}", file=sys.stderr)
        return 1

    targets = list(iter_docx(in_path))
    if not targets:
        print("未找到可处理的 .docx 文件", file=sys.stderr)
        return 1

    if len(targets) > 1 and out_arg is not None and out_arg.suffix.lower() == ".docx":
        print(
            f"输入有 {len(targets)} 个文件，但 -o 指向单个 .docx 文件；"
            "请改为输出目录，否则多个文件会相互覆盖。",
            file=sys.stderr,
        )
        return 1

    normalize = not args.no_punct_normalize
    exit_code = 0

    for src in targets:
        out_path = build_output_path(src, out_arg, args.in_place)

        in_place = args.in_place
        if not in_place and out_path.exists() and not args.force:
            print(f"跳过（输出已存在，需 --force）：{out_path}", file=sys.stderr)
            exit_code = 2
            continue

        if args.dry_run:
            plan = "覆盖原件" if in_place else "写入"
            print(f"[dry-run] {src} -> {out_path}（{plan}）")
            continue

        backup = None
        if in_place:
            backup = backup_path(src)
            backup.write_bytes(src.read_bytes())
            print(f"已备份: {backup}")

        try:
            format_document(file_path=src, out_path=out_path, en_font=args.en_font, normalize=normalize)
        except PackageNotFoundError:
            print(f"跳过（不是有效的 .docx）: {src}", file=sys.stderr)
            exit_code = 1
            continue
        except Exception as exc:  # 单个文件失败不中断批量
            print(f"失败: {src} -> {type(exc).__name__}: {exc}", file=sys.stderr)
            exit_code = 1
            continue

        print(f"已生成: {out_path}")

    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
