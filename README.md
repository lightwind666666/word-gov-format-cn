# word-gov-format-cn

给中文 `.docx` 一键套用公文版式的 Agent Skill + 命令行脚本。

- A4 纸，页边距上 3.7cm / 下 3.5cm / 左 2.8cm / 右 2.6cm（版心约 156×225mm）
- 标题 `方正小标宋简体` 二号，一级标题 `黑体` 三号，二级标题 `楷体_GB2312` 三号，正文 `仿宋_GB2312` 三号首行缩进 2 字
- 固定行距 28pt，段前段后 0
- 阿拉伯数字 / 英文使用 `Times New Roman`

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

# 先预演，不写任何文件
python scripts/word_gov_format_cn.py ./稿件 -o ./输出 --dry-run

# 覆盖已存在的输出文件
python scripts/word_gov_format_cn.py ./稿件 -o ./输出 --force

# 原地修改（自动生成 .bak-<时间戳> 备份）
python scripts/word_gov_format_cn.py 通知.docx --in-place

# 只排版、不改标点
python scripts/word_gov_format_cn.py 通知.docx --no-punct-normalize

# 自定义数字/英文字体
python scripts/word_gov_format_cn.py 通知.docx --en-font "Arial"
```

退出码：`0` 全部成功；`1` 有文件失败/参数错误；`2` 有输出已存在被跳过（未加 `--force`）。

## 安全行为

| 场景 | 行为 |
|---|---|
| 输出文件已存在 | 跳过并提示，需 `--force` 才覆盖（退出码 2） |
| `--in-place` | 先写 `.bak-<时间戳>` 备份，再改写原件 |
| `--dry-run` | 只打印计划路径，不创建任何文件 |
| 目录批量 + `-o 单个.docx` | 直接报错拒绝，避免多文件写同一路径 |
| 输入损坏 / 非 docx | 提示后跳过，继续处理其余文件 |

## 防止误提交业务材料

- `.gitignore` 默认忽略 Word、Excel、CSV、PDF、演示文稿、原件备份（含 `.bak-<时间戳>`）、`.env` 配置和常见私钥文件，文档扩展名兼容大小写。
- 已审核的虚构样例 `sample/样例公文.docx` 继续保留跟踪；不对整个 `sample/` 目录放行，新放入的业务文档仍被忽略。
- 真实材料和备份建议存放在仓库外的受控目录；不要用 `git add -f` 绕过忽略规则提交业务材料。
- 忽略规则不会清除已跟踪文件或历史记录，也不检查文档正文、作者属性、批注和修订。提交前仍需检查 `git diff --cached`；新增样例必须使用虚构内容并检查隐藏属性。

## 标点规范化的作用范围

只在**含中文的段落**做半角转全角，识别并保护以下常见形式：

```
www.hangzhou.gov.cn      → 不变
test@abc.com             → 不变
e.g.  /  v1.2.3          → 不变
don't                    → 不变
10:30                    → 不变
1. 列表编号               → 不变
各科室,各单位.            → 各科室，各单位。
他说"没问题"              → 他说“没问题”
纯英文段落                → 完全不处理
```

识别使用整段文字上下文，网址、邮箱、时间、编号和引号跨越多个 Word 格式片段（run）时也使用同一规则，再将结果写回原有片段。超链接自身和含复杂对象的 run 不改写文本。

标点判断使用启发式规则，成稿仍需核对。需要完全关闭：`--no-punct-normalize`。

## 内容保真

- 按原字符位置回填普通 run，保留行内图片及正文原有加粗，不重建整个段落。
- 含图片、制表符、换行符、域的 run 不做文本改写，只调整字体字号。
- 表格单元格按正文规则处理。

## 已知限制

- 只做版式，不做完整公文合规校验。
- 发文机关 / 标题靠启发式识别（首段 ≤30 字且以「局/部/委/办/中心/处/室/会/厅」结尾即视为机关），窄标题可能误判为红色抬头，生成后请核对。
- 页眉页脚、文本框/图形内文字不处理。
- 发文机关标志下的红色分隔线不自动添加，见 `references/red-underline.md`。
- 目录批量不递归。
- 字体需本机安装，缺失时 Word 会自动替换。

## 字体授权

本仓库**不附带任何字体文件**，仅引用字体名称。`方正小标宋简体` 为方正商用字体，`仿宋_GB2312` / `楷体_GB2312` 为中易字体，商业或对外发布场景请自行确认并取得授权，效果以本机已安装字体为准。

## 自测

```bash
# 跨 run 标点、内容保真和误提交规则检查（内存处理文档，只读检查 Git）
python -B -m unittest discover -s scripts -p 'test_*.py' -v

# 原有 CLI 自测（在临时目录验证输出、覆盖和备份）
python scripts/selftest.py
```

覆盖：跨格式片段的网址/邮箱/时间/编号/引号、超链接及图片保留、误提交规则，以及原有的版式参数、覆盖保护、备份、dry-run、批量冲突。

## 许可

代码 MIT。
