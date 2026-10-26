---
name: word-gov-format-cn
description: Apply Chinese official-document Word formatting rules (A4 margins, fixed line spacing, heading/body font mapping, full-width punctuation) to .docx files. Use when asked to 公文排版、套公文格式、统一字体字号行距、把文档按党政机关公文格式整理。
---

# Word Gov Format CN

给中文 `.docx` 套用公文版式。安全默认：不覆盖已有文件、原地修改自动备份、可预演。

## 安装 / Installation

Requires Python 3.10 or newer. Clone this repository into your agent's skill directory, or use the script directly:

```bash
git clone https://github.com/lightwind666666/word-gov-format-cn.git
cd word-gov-format-cn
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

Windows PowerShell: activate with `.venv\Scripts\Activate.ps1`.

## Quick start

```bash
# 单文件 → 生成 xxx_排版.docx（同名存在时不会覆盖）
python3 scripts/word_gov_format_cn.py <input.docx>

# 批量处理目录（不递归）→ 输出到指定目录
python3 scripts/word_gov_format_cn.py <folder_path> -o <output_dir>

# 先预演，确认后再执行
python3 scripts/word_gov_format_cn.py <input.docx> --dry-run

# 覆盖已存在的输出
python3 scripts/word_gov_format_cn.py <input.docx> -o <out_dir> --force

# 原地修改（自动生成 .bak-<时间戳> 备份）
python3 scripts/word_gov_format_cn.py <input.docx> --in-place

# 只排版、不动标点
python3 scripts/word_gov_format_cn.py <input.docx> --no-punct-normalize
```

## 必须先确认的破坏性操作

- `--in-place` 会改写原件（虽自动备份）；`--force` 会覆盖已有文件。
- 执行前先 `--dry-run` 列出计划，并向用户确认目标路径。
- 未经用户明确同意，**不得**使用 `--in-place` 或 `--force`。

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
- Punctuation: half-width → full-width **only in paragraphs containing Chinese**, and only in Chinese context

## Punctuation scope

Protected and left untouched: URLs (`www.*`, `http(s)://`), emails, dotted Latin tokens (`gov.cn`, `e.g.`, `v1.2.3`, `file.docx`), times (`10:30`), ASCII apostrophes (`don't`), and list numbering (`1.`). Paragraphs with no Chinese characters are not normalized at all. Use `--no-punct-normalize` to disable entirely.

## 未处理的部分

- 发文机关标志下的红色分隔线需手动添加，见 `references/red-underline.md`。
- 页眉页脚、文本框/图形内文字不处理（避免误伤版式）。
- 不做完整公文合规校验，只套版式；成稿仍需人工核对。

## Pitfalls

- Use Python ≥ 3.10; the script uses `Path | None` annotations.
- 发文机关与标题靠启发式识别（首段 ≤30 字且以「局/部/委/办/中心/处/室/会/厅」结尾视为机关）。窄标题、被截断的标题行可能判错，生成后需肉眼确认。
- `--in-place` 的备份是 `<文件名>.bak-<时间戳>`，不是版本管理；重要文件仍建议先提交到 git 或另存副本。
- 目录批量不递归，子目录需单独处理；多个输入不要配 `-o 单个.docx`（脚本会报错拒绝）。
- 字体需本机安装：`方正小标宋简体` 为方正商用字体，`仿宋_GB2312`/`楷体_GB2312` 为中易字体，请自行确认授权；本仓库不附带任何字体文件。缺失时 Word 会回退替换。

## Resource

### scripts/word_gov_format_cn.py

Main formatter script for Chinese official-document Word layout.

### scripts/selftest.py

Regression check for the safety behaviours above: `python3 scripts/selftest.py`.
