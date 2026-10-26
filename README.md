# word-gov-format-cn

给中文 `.docx` 一键套用公文版式的 Agent Skill + 命令行脚本。

- A4 纸，页边距上 3.7cm / 下 3.5cm / 左 2.8cm / 右 2.6cm（版心约 156×225mm）
- 标题 `方正小标宋简体` 二号，一级标题 `黑体` 三号，二级标题 `楷体_GB2312` 三号，正文 `仿宋_GB2312` 三号首行缩进 2 字
- 固定行距 28pt，段前段后 0
- 阿拉伯数字 / 英文使用 `Times New Roman`
- 半角标点转为全角标点

## 安装

```bash
git clone https://github.com/lightwind666666/word-gov-format-cn.git
cd word-gov-format-cn
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

需要 Python 3.10+（脚本使用了 `Path | None` 类型注解）。

## 用法

```bash
# 单文件 → 生成 xxx_排版.docx
python scripts/word_gov_format_cn.py 通知.docx

# 批量处理目录（不递归）→ 输出到指定目录
python scripts/word_gov_format_cn.py ./稿件 -o ./输出

# 直接覆盖原文件（谨慎）
python scripts/word_gov_format_cn.py 通知.docx --in-place

# 自定义数字/英文字体
python scripts/word_gov_format_cn.py 通知.docx --en-font "Arial"
```

## 生成一份测试文档

```bash
python scripts/make_sample.py
python scripts/word_gov_format_cn.py sample/样例公文.docx -o sample/out
```

## 作为 Agent Skill 使用

本仓库的 `SKILL.md` 遵循 AgentSkills 规范。把整个目录放进 agent 的 skill 目录即可（例如 `~/.codex/skills/`、`~/.claude/skills/`），脚本路径相对于 skill 目录解析。

## 已知限制

- 只做版式，不做完整公文合规校验。
- 会重写段落文本以规范标点，段落内的图片、超链接、域、混合格式可能丢失。
- 发文机关标识（红色抬头）下的红色分隔线脚本不会自动加，见 `references/red-underline.md`。
- 批量模式请传输出目录，不要传单个 `.docx` 文件名，否则多个输入会写到同一路径相互覆盖。
- 字体需本机安装，缺失时 Word 会自动替换。

## 许可

MIT
