# 修改点与迭代记录

本文件记录每次改动，并说明代码结构与「如何新增一种标注」，方便后续迭代。

---

## v0.5.0 —— 打包重构 + 开源发布

### 变更
- **重构为包结构**：单文件脚本拆分为 `restore_zhongzhuhao/` 包
  （`core.py` 纯逻辑、`cli.py` 命令行、`gui.py` 界面），职责分离、便于测试与复用。
- **入口统一**：
  - `python -m restore_zhongzhuhao`（CLI，`__main__.py`）
  - `python gui.py`（顶层便捷启动器 → 包内 GUI）
  - 安装后：`restore-zhongzhuhao` / `restore-zhongzhuhao-gui`（见 `pyproject.toml`）
- **测试迁移**：`test/test_zhongzhuhao.py` → `tests/test_core.py`，兼容 pytest。
- **对外 API**：`__init__.py` 统一导出核心函数与 `__version__`。

### 新增
- `LICENSE`：采用 **The Unlicense**（公共领域，最开放）。
- `pyproject.toml`：打包元数据 + console/gui 脚本入口。
- 更完整的 `README.md`：特性、安装、用法、结构、原理、限制、环境说明。

### 说明
- 功能行为与 v0.4.0 一致（着重号 + 颜色，Markdown / HTML 输出）。

---

## v0.4.0 —— HTML 输出 + GUI 美化

### 新增
- **HTML 输出格式**：`--format html`（CLI）/ GUI 单选「HTML (.html)」。
  流程：先在 Markdown 上修复标注，再用 pandoc 转成独立网页（`markdown_to_html()`），
  最后把样式注入 `<head>`（`inject_html_style()`），得到可直接打开的成品 HTML。
- **最高层入口 `process(...)`**：按 `output_format` 分派 markdown / html，CLI 与 GUI 共用。
- **CLI 参数**：`--format {markdown,html}`、`--pandoc-bin`。

### 优化
- `markdown_to_html()` 与 `run_pandoc()` 均加 `--wrap=none`，避免 pandoc 换行把
  `<span …>` 标签拆到两行（虽然 HTML 合法，但不利于查看/断言）。
- 路径默认后缀随格式自动切换（html → `.html`，md → `.md`）。

### GUI 美化（gui.py 重写）
- 顶部蓝色标题栏 + 副标题；卡片式分区（文件 / 处理选项 / 效果示意 / 日志）。
- 统一配色与字体（自动选用 微软雅黑 / Segoe UI），自绘「强调」主按钮。
- 新增**输出格式单选**、**渲染效果示意图例**（Canvas 画字下加点与红字）、
  **状态栏**、**打开输出位置**按钮。
- 处理在后台线程执行，日志重定向回界面。

---

## v0.3.0 —— 字体颜色 + GUI + 处理逻辑重构

### 新增
- **字体颜色标注**：读取 docx 的 `<w:color w:val="RRGGBB"/>`，
  在 Markdown 中补为 `<span style="color:#RRGGBB">…</span>`。
- **颜色与着重号叠加**：同一段文字可同时带 class 与 style，
  如 `<span class="zhongzhuhao" style="color:#0000FF">…</span>`。
- **图形界面** `gui.py`（tkinter，无需额外依赖）：
  选文件、勾选处理项、试运行、后台线程执行、日志输出。
- **CLI 新参数**：`--no-emphasis`、`--no-color`、`--emphasis-class`。

### 重构（优化处理逻辑）
- 引入数据模型 `Annotation` / `Segment` / `Options`，把「标注」抽象成可扩展字段。
- `extract_emphasis()`（只取着重号）→ **`extract_segments()`**（统一取所有标注，
  同段落相邻且标注相同的 run 合并为一段）。
- 匹配器从「逐行」升级为「整篇正则匹配」：
  - 短语内的空格可匹配软换行 `[ \t]*(?:\n[ \t]*)?`，解决 pandoc 换行拆散短语。
  - 受保护区（围栏代码块、行内代码、链接/图片）改为整篇绝对偏移，命中时跳过。
  - 插入标签后用 `_shift_protected()` 平移保护区偏移，避免每段重算。
- `run_pandoc()` 在 markdown 输出时追加 `--wrap=none`，从源头避免软换行。
- 新增高层入口 `process_markdown()`，CLI 与 GUI 共用，消除重复逻辑。

### 提醒
- `extract_emphasis()` 已移除；如有外部调用请改用 `extract_segments(...)` 后取 `[s.text for s in ...]`。

---

## v0.2.0 —— 输出改为 Markdown

- 由「docx → HTML」改为「docx → Markdown」，因源文档转出为 Markdown。
- 用内联 HTML `<span class="zhongzhuhao">` 补着重号（Markdown 支持内联 HTML，pandoc 透传）。
- 跳过围栏代码块、行内代码、链接，避免破坏 Markdown 语法。

## v0.1.0 —— 初版（HTML）

- 读 docx `w:em`，在 pandoc 生成的 HTML 中回填 `<span>`，注入 text-emphasis 样式。

---

## 代码结构

```
gui.py                        便捷启动器：python gui.py
restore_zhongzhuhao/
├── __init__.py               对外 API + __version__
├── __main__.py               python -m restore_zhongzhuhao
├── core.py                   纯逻辑（无 UI 依赖）
│   ├── 数据模型              Annotation / Segment / Options
│   ├── 提取                  extract_segments()
│   ├── Markdown 匹配         mark_fenced_code / protected_ranges / compute_protected
│   │                         _pattern_for / _search / _shift_protected
│   ├── 渲染                  _span_open()      ← Annotation → <span ...>
│   ├── 包裹                  wrap_segments()
│   ├── 高层入口              process_markdown() / process()
│   └── 转换                  run_pandoc() / markdown_to_html() / inject_html_style()
├── cli.py                    命令行入口 main()
└── gui.py                    tkinter 界面，调用 process()
tests/test_core.py            端到端测试
```

## 如何新增一种标注（例如粗体、高亮底色）

以新增「粗体 `<w:b/>`」为例，只需三处：

1. **数据模型**：在 `Annotation` 加字段
   ```python
   bold: bool = False
   ```
   并更新 `is_empty()`（加入 `and not self.bold`）。

2. **提取**：在 `extract_segments()` 里读取 run 属性
   ```python
   bold = rpr is not None and rpr.find(wtag("b")) is not None
   ```
   放进 `Annotation(...)`。
   （注意：若与已有字段组合，相邻合并逻辑自动按整个 Annotation 是否相等判断。）

3. **渲染**：在 `_span_open()` 里输出对应属性
   ```python
   if annot.bold:
       attrs.append('style="font-weight:bold"')  # 或合并进已有 style
   ```
   > 注意目前 `style` 是单属性拼接；若同时有 color 与 bold，需要合并成一个
   > `style="color:...;font-weight:bold"`，建议改成「收集样式列表再 join」。

无需改动匹配、包裹、CLI、GUI —— 它们都基于 `Segment`/`Annotation` 通用运行。

## 已知限制

- 假设文档中「该短语首次出现的位置」即 docx 标注的位置（通常成立）。
- pandoc 文本变形（智能引号、破折号、转义 `\*` 等）会导致匹配失败并告警跳过。
- 缩进式代码块（行首 4 空格）未特殊跳过，建议用围栏代码块。
- 跨段落合并的标注不会合并。

## 待办 / 可迭代方向

- [ ] `_span_open()` 支持多 style 合并（color + 其他样式）。
- [ ] 支持更多标注：粗体、斜体、下划线、删除线、高亮底色、字号。
- [ ] GUI 增加「预览」与「按颜色筛选/统计」。
- [ ] 模糊匹配（相似度）以容忍 pandoc 的文本变形。
- [ ] 批量处理多个 docx。
- [ ] 打包为 exe（PyInstaller）。
