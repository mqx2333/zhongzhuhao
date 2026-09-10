# restore_zhongzhuhao

辅助 **pandoc** 修复 Word(docx) 转换时丢失的字符标注，输出 **Markdown 或 HTML**：

- **着重号（字下加点）**：docx 的 `<w:em w:val="dot"/>`
- **字体颜色**：docx 的 `<w:color w:val="RRGGBB"/>`

pandoc 转换时会忽略这两者。本程序读回 docx 的标注，在 Markdown 中用内联 HTML
`<span>` 补回（Markdown 支持内联 HTML，pandoc 再转 HTML 时原样透传）；
也可直接输出成品 HTML（内部用 pandoc 转换并把样式注入 `<head>`）。
可选注入 `text-emphasis` 样式渲染字下加点。

```html
<span class="zhongzhuhao">着重文字</span>
<span style="color:#FF0000">红色文字</span>
<span class="zhongzhuhao" style="color:#0000FF">又蓝又着重</span>
```

## 工作流

```
src.docx ──pandoc──> 转出的.md（标注丢失）
   │
   └──(可选 --run-pandoc 由程序调用 pandoc)──┘
                 │
             修复：读 docx 提取 w:em / w:color → 在 md 中以内联 <span> 包裹 → 修复后.md
```

## 安装

```bash
pip install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
```

> 说明：仅 `lxml` 为必需依赖；核心逻辑已不再使用 BeautifulSoup。

## 使用

### 图形界面

```bash
python gui.py
```

选源 docx、输入 Markdown（或勾选「由 pandoc 生成」），
选择**输出格式（Markdown / HTML）**，填输出文件，勾选要处理的标注，点「开始处理」。
界面含渲染效果图例、状态栏与「打开输出位置」。

### 命令行

```bash
# 已有 pandoc 生成的 markdown -> 修复为 Markdown
python restore_zhongzhuhao.py --docx 原稿.docx --markdown pandoc输出.md -o 修复后.md

# 一键：先调 pandoc 转 markdown（内置 --wrap=none）再修复
python restore_zhongzhuhao.py --docx 原稿.docx --run-pandoc -o 修复后.md

# 直接输出成品 HTML（用 pandoc 转成独立网页）
python restore_zhongzhuhao.py --docx 原稿.docx --run-pandoc --format html -o 成品.html
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
| `--emphasis-class` | 着重号 span 的 class 名，默认 `zhongzhuhao` |
| `--css CSS` | 自定义注入样式 |
| `--no-style` | 不注入样式块 |
| `--dry-run` | 只打印将处理的短语，不写文件 |
| `--encoding` | 文件读写编码，默认 `utf-8` |

## 工作原理

1. 用 `zipfile` + `lxml` 读 `word/document.xml`，遍历 `<w:p>/<w:r>`，
   取每个 run 的 `w:em`（着重号）与 `w:color`（颜色），
   同段落内相邻且标注相同的 run 合并为一段（`Segment`）。
2. 读取 Markdown，整篇正则匹配各段文字；短语中的空格可匹配软换行；
   跳过围栏代码块、行内代码、链接/图片等受保护区域。
3. 命中处用内联 HTML `<span>`（按需带 `class` / `style`）包裹；可选注入样式。
4. `--format html` 时，把修复后的 Markdown 交给 pandoc 转成独立 HTML，
   再把样式注入 `<head>`，输出可直接打开的网页。

## 测试

```bash
python test/test_zhongzhuhao.py
```

## 迭代

改动记录、代码结构与「如何新增一种标注」见 [CHANGELOG.md](CHANGELOG.md)。

## 已知限制

- 假设文档中「该短语首次出现的位置」即 docx 标注的位置（通常成立）。
- 若 pandoc 引入文本变形（智能引号、破折号、`\*` 转义等），对应短语可能匹配失败并告警跳过。
- 缩进式代码块（行首 4 空格）未特殊跳过，建议用围栏代码块。
- 跨段落的标注不会合并。
