---
name: word-gov-format-cn
description: Apply Chinese official-document Word formatting rules (A4 margins, fixed line spacing, heading/body font mapping, full-width punctuation) to .docx files.
---

# Word Gov Format CN

Use this skill to format Chinese `.docx` files according to your specified公文版式要求.

## 安装 / Installation

Requires Python 3.10 or newer. Clone this repository into your agent's skill directory, or use the script directly:

```bash
git clone https://github.com/lightwind666666/word-gov-format-cn.git
cd word-gov-format-cn
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell: activate with `.venv\\Scripts\\Activate.ps1`.

## Quick start

Single file:

```bash
python3 scripts/word_gov_format_cn.py <input.docx>
```

Batch process a folder:

```bash
python3 scripts/word_gov_format_cn.py <folder_path>
```

Write output to a specific directory:

```bash
python3 scripts/word_gov_format_cn.py <input_or_dir> -o <output_dir>
```

Overwrite original file (caution):

```bash
python3 scripts/word_gov_format_cn.py <input.docx> --in-place
```

## Implemented format rules

- Paper: A4 (210mm x 297mm)
- Margins: top 3.7cm, bottom 3.5cm, left 2.8cm, right 2.6cm
- Body area equivalent: 156mm x 225mm
- Body text: 仿宋_GB2312, 三号 (16pt), first-line indent 2 chars
- Level-1 heading (`一、`): 黑体, 三号
- Level-2 heading (`（一）`): 楷体_GB2312, 三号
- Level-3 heading (`1.`): 仿宋_GB2312, 三号
- Level-4 heading (`（1）`): 仿宋_GB2312, 三号
- Issuer mark: 方正小标宋简体, 二号, red
- Document title: 方正小标宋简体, 二号
- Line spacing: fixed 28pt (titles use 32/38pt heuristic)
- Paragraph spacing before/after: 0pt
- Arabic digits and Latin text use Times New Roman
- Half-width punctuation normalized to full-width punctuation where applicable
- **Red separator line under issuer mark**: NOT applied by the script. Manual step required — add a bottom border (single, red, ~1pt) to the first paragraph (issuer mark) after formatting. See `references/red-underline.md` for the code snippet.

## Notes

- Requires `python-docx` in the Python environment (`pip install python-docx`).
- `.doc` files are not supported; convert to `.docx` first.
- Font rendering depends on local font availability. If a font is missing, Word will substitute a fallback.
- Double-sided printing and left-side binding are printer/output settings; they are not forced by this script.

## Pitfalls

- Use Python ≥ 3.10; the script uses `Path | None` annotations. Install dependencies into the same Python environment used to run the script.
- The formatter replaces paragraph runs with normalized text. Use it on plain-text documents; inline pictures, hyperlinks, fields, and mixed run formatting in processed paragraphs may be lost.
- Review automatic issuer/title detection and punctuation changes. In particular, `1.` becomes `1。` before heading classification, so third-level numbering currently receives body indentation.
- When processing a directory, pass an output directory, not a single `.docx` output filename; otherwise outputs overwrite each other. Directory processing is not recursive.
- This applies a specific layout preset, not a complete official-document compliance check.

## Resource

### scripts/word_gov_format_cn.py

Main formatter script for Chinese official-document Word layout.
