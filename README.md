# restore-zhongzhuhao

> 把 Word(docx) 转 Markdown 时被 pandoc 丢掉的 **着重号（字下加点）** 和 **字体颜色** 自动补回来。

[![License: Unlicense](https://img.shields.io/badge/license-Unlicense-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)](#)

---

## 它解决什么问题

Word（`.docx`）里会把字符标注存进运行属性（`<w:rPr>`）：

| 标注 | docx 中的表示 | pandoc 转 Markdown/HTML 时 |
| --- | --- | --- |
| 着重号（字下加点） | `<w:em w:val="dot"/>` | **被忽略**，退化为普通文本 |
| 字体颜色 | `<w:color w:val="FF0000"/>` | **被忽略**，退化为普通文本 |

结果就是转换后的文档丢失了强调信息。本工具读回 docx 的标注，在 Markdown 中
用内联 HTML `<span>` 复原（Markdown 支持内联 HTML，pandoc 再转 HTML 时原样透传），
也可一步直接产出成品 HTML。

```html
<span class="zhongzhuhao">着重文字</span>
<span style="color:#FF0000">红色文字</span>
<span class="zhongzhuhao" style="color:#0000FF">又蓝又着重</span>
```

配合注入的 CSS，浏览器即可正确渲染字下加点：

```css
.zhongzhuhao {
  text-emphasis-style: dot;
  text-emphasis-position: under;
  text-emphasis-color: currentColor;
}
```

## 特性

- ✅ 复原**着重号**（字下加点）与**字体颜色**，两者可叠加
- ✅ 默认忽略黑色 `#000000` / 白色 `#FFFFFF`，避免把 Word 默认正文色当成标注
- ✅ 输出 **Markdown**（内联 HTML）或 **HTML**（独立网页，样式自动注入 `<head>`）
- ✅ 跳过代码块、行内代码、链接，避免破坏 Markdown 语法
- ✅ 容忍 pandoc 软换行（短语被换行拆开也能匹配）
- ✅ 提供**图形界面**（tkinter，零额外依赖）
- ✅ 命令行 / 模块 / 安装后入口，一应俱全
- ✅ 结构清晰、易于扩展新标注类型（见 [CHANGELOG.md](CHANGELOG.md)）

## 开发环境

本项目在 **OpenCode + DeepSeek v4.1** 环境下开发完成。

## 环境要求

- Python **3.9+**
- [pandoc](https://pandoc.org/)（用于 `--run-pandoc` 转换，以及 HTML 输出）
- 依赖：`lxml`（唯一必需第三方库）

## 安装

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

或安装为本地包（提供命令入口）：

```bash
pip install -e .
# 之后可用：
restore-zhongzhuhao --help
restore-zhongzhuhao-gui
```

## 快速开始

### 图形界面

```bash
python gui.py
```

选择源 docx、输入 Markdown（或勾选「由 pandoc 生成」），选择**输出格式**，
填输出文件，勾选要处理的标注，点「开始处理」。

### 命令行

```bash
# 已有 pandoc 生成的 markdown -> 修复为 Markdown
python -m restore_zhongzhuhao --docx 原稿.docx --markdown pandoc输出.md -o 修复后.md

# 一键：先调 pandoc 转 markdown（内置 --wrap=none）再修复
python -m restore_zhongzhuhao --docx 原稿.docx --run-pandoc -o 修复后.md

# 直接输出成品 HTML
python -m restore_zhongzhuhao --docx 原稿.docx --run-pandoc --format html -o 成品.html
```

### 参数

| 参数 | 说明 |
| --- | --- |
| `--docx` | 源 Word 文件（须含 `w:em` / `w:color`）；必填 |
| `--markdown` | pandoc 生成的 Markdown 文件 |
| `-o, --out` | 输出文件；省略则按格式写回 / 同名（`.md` / `.html`） |
| `--format` | 输出格式 `markdown`（默认）或 `html` |
| `--run-pandoc [CMD]` | 先调用 pandoc（默认 `pandoc`）把 docx 转成 markdown 再修复 |
| `--pandoc-bin` | pandoc 可执行文件（HTML 转换用），默认 `pandoc` |
| `--to` | pandoc 输出格式，默认 `markdown` |
| `--no-emphasis` | 不处理着重号 |
| `--no-color` | 不处理字体颜色 |
| `--ignore-colors` | 忽略的颜色（逗号分隔），默认 `#000000,#FFFFFF`；传空串即不过滤 |
| `--emphasis-class` | 着重号 span 的 class 名，默认 `zhongzhuhao` |
| `--css CSS` | 自定义注入样式 |
| `--no-style` | 不注入样式块 |
| `--dry-run` | 只打印将处理的短语，不写文件 |
| `--encoding` | 文件读写编码，默认 `utf-8` |

## 工作原理

1. **提取**：用 `zipfile` + `lxml` 读 `word/document.xml`，遍历 `<w:p>/<w:r>`，
   取每个 run 的 `w:em` 与 `w:color`；同段落内相邻且标注相同的 run 合并为一段（`Segment`）。
2. **匹配**：读取 Markdown 全文，对每个标注段做正则定位；短语中的空格可匹配软换行；
   跳过围栏代码块、行内代码、链接/图片等受保护区域。
3. **包裹**：命中处以内联 HTML `<span>`（按需带 `class` / `style`）包裹。
4. **输出**：`--format html` 时再用 pandoc 转成独立 HTML 并把样式注入 `<head>`。

## 项目结构

```
zhongzhuhao/
├── gui.py                      # 便捷启动器：python gui.py
├── restore_zhongzhuhao/        # 主包
│   ├── __init__.py             # 对外 API
│   ├── __main__.py             # python -m restore_zhongzhuhao
│   ├── core.py                 # 核心：提取 / 匹配 / 渲染（无 UI 依赖）
│   ├── cli.py                  # 命令行入口
│   └── gui.py                  # tkinter 图形界面
├── tests/
│   └── test_core.py            # 端到端测试
├── README.md
├── CHANGELOG.md                # 迭代说明 / 修改点
├── LICENSE                     # The Unlicense（公共领域）
├── requirements.txt
└── pyproject.toml
```

## 测试

```bash
python tests/test_core.py
# 或
pytest -q
```

测试会手工构造含 `w:em` / `w:color` 的最小 docx，调用 pandoc 转 Markdown，
再运行修复并断言结果（含代码块/行内代码/链接跳过、软换行、颜色叠加等）。

## 已知限制

- 假设「短语在文档中首次出现的位置」即 docx 标注的位置（通常成立）。
- 若 pandoc 引入文本变形（智能引号、破折号、`\*` 转义等），对应短语会告警并跳过。
- 缩进式代码块（行首 4 空格）未特殊跳过，建议使用围栏代码块。
- 跨段落的标注不会合并。

## 迭代说明

改动记录、代码结构，以及「如何新增一种标注（例如粗体/高亮）」的三步指南，
见 **[CHANGELOG.md](CHANGELOG.md)**。

## 许可证

[**The Unlicense**](LICENSE) —— 本作品已释放至公共领域（Public Domain）。
你可以任意复制、修改、发布、商用，无需署名，无需承担任何担保。
这是最开放的开源许可之一。

## 致谢

- [pandoc](https://pandoc.org/) —— 强大的文档转换器
- [lxml](https://lxml.de/) —— 解析 OOXML
