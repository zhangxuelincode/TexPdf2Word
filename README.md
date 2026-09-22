# TexPdf2Word

一个将 **LaTeX、PDF、Markdown 或粗转 DOCX** 转译成指定 Word 模板格式的 AI Agent Skill。

它面向毕业论文、学位论文、单位报告、标准文档等场景，尤其适合那些 **pandoc 默认 DOCX 输出不够用**、必须严格套用学校或机构 Word 模板的任务。

English: see the [Quick Start](#-quick-start) and [Usage](#-usage) sections below.

![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)
![Python](https://img.shields.io/badge/python-3.10%2B-3776AB.svg)
![Output](https://img.shields.io/badge/output-DOCX-2B579A.svg)

---

## 目录

- [这是什么](#-这是什么)
- [支持的功能](#-支持的功能)
- [工作原理](#-工作原理)
- [环境要求](#-环境要求)
- [Quick Start](#-quick-start)
- [完整操作步骤](#-完整操作步骤)
- [项目结构](#-项目结构)
- [已知限制](#-已知限制)
- [安全说明](#-安全说明)
- [许可证与致谢](#-许可证与致谢)

## 📖 这是什么

`pandoc --reference-doc` 只能复用 Word 模板的"样式定义"，不理解模板里的
页面语义和版式约束。对严格模板（学位论文、单位报告等）经常出现：
封面/声明页没填、图表标题丢编号、表格不是三线表、引用不是可跳转的上标、
目录字段是空的、正文页眉串成"致谢"……

TexPdf2Word 不是又一个"通用转换器"，而是一套 **AI 模板适配工作流**：

1. 先用脚本 **inspect** 你的 Word 模板，摸清样式、编号、分节和页眉页脚；
2. 让 AI Agent 基于检查结果，**生成一版项目专用的 Python 重建脚本**；
3. 用 `python-docx` + 底层 OOXML 把正文、图题、公式、表格映射进模板；
4. 用 Word COM 更新目录/页码字段并导出 PDF；
5. 用结构层 + 渲染层两套校验器和 PDF 预览拼图做 **自动质检**。

所有脚本都是确定性 Python，可独立于任何 AI Agent 手动使用。

## ✨ 支持的功能

### 输入 → 输出

| 输入 | 输出 |
| --- | --- |
| LaTeX 项目（多文件、图片、BibTeX） | 完整 DOCX（含模板封面/声明/目录） |
| Markdown 文件 | 同上 |
| 原生数字 PDF（有文本层） | 同上 |
| 已有粗转 DOCX | 适配模板后的成品 DOCX + PDF 预览 |

### 转换能力

- **模板前置页保护与填充**：封面、英文封面、原创性声明、授权页、签名页
  按"受保护区域"处理——只替换文字内容，不破坏段落样式、字体、字号、
  加粗、对齐、间距和分页结构。
- **正文样式映射**：按可见样式名重映射 `w:pStyle/w:rStyle/wtblStyle`，
  避免 pandoc 与模板的样式 ID 冲突导致 `Heading 1` 静默变成别的样式。
- **章节自动编号**：修复多级标题编号（`第N章` / `1.1` / `1.1.1`）、
  前后置标题（摘要/参考文献/致谢）取消编号。
- **图题/表题修复**：自动补 `图 X.Y` / `表 X.Y` 编号前缀并居中，
  支持从源 `.tex` 回读 caption 文本。
- **三线表**：可选把全部表格改写成规范三线表（含垂直合并单元格的
  底边修复）。
- **参考文献与引用**：作者-年份引用改写为 GB/T 7714 数字上标 `[N]`，
  并通过书签 + 内部超链接实现 **点击跳转** 到参考文献条目。
- **交叉引用导航**：正文里的 `图N-M` / `表N-M` / `式(N-M)` / 代码清单
  引用全部变成可点击跳转的内部链接。
- **公式保留**：pandoc 生成的 OMML 公式原样保留，不做破坏性重排。
- **页眉页脚修复**：清理删除样例章节后遗留的页眉引用，支持改写为
  `STYLEREF` 动态页眉（随章节标题变化）、正文页码从 1 重新开始。
- **目录（TOC）注入**：粗转产物"只有目录标题、没有 TOC 字段"时，
  幂等注入 `{ TOC \o "1-3" \h \z \u }` 字段。
- **Word 终处理**：通过 Word COM 更新全部字段/目录并导出 PDF
  （自动禁用宏，保证安全）。
- **两级自动质检**：
  - *结构层* `validate_docx_conversion.py`：模板占位符残留、正文/后置
    顺序、标题样式、继承页眉、图片/表格数量、受保护页格式漂移；
  - *渲染层* `validate_docx_render.py`：TOC 字段存在性、numId↔abstractNum
    一致性、多级标题格式串、参考文献计数器独立性、正文页眉静态文本
    残留、图片/表格/caption 数量与源文件比对、引用覆盖率、三线表样式、
    PDF 字段错误串（`错误!` / `Error!`）。
- **PDF 预览拼图**：把导出 PDF 的指定页渲染成一张拼图，肉眼快速复核
  封面、目录、图表、公式、参考文献。

### 兼容宿主

| Agent / IDE | 状态 | 说明 |
| --- | --- | --- |
| **Codex** | 原生支持 | 放到 `~/.codex/skills/`，`Use $texpdf2word …` 触发 |
| **Claude Code** | 包装后可用 | `SKILL.md` frontmatter 与 Claude Skills 规范一致 |
| **Cursor** | 手动 / Rules | 把 SKILL.md 内容放进 `.cursor/rules/*.mdc` |
| 其他能跑 shell 的 Agent | 可用 | 脚本是纯 CPython，无 runtime 依赖 |

## 🔍 工作原理

```
LaTeX / PDF / Markdown ──(pandoc / Word 导入 / pdf2docx)──▶ 粗转 body.docx
                                                                │
用户模板 template.docx ──(inspect_docx_template.py)──▶ 模板检查报告 template_report.json
                                                                │
                     AI Agent 结合两份报告，改造 adaptive_docx_pipeline.py
                                                                │
                                     ▼
                    final.docx（正文映射进模板 + 图题/表格/引用修复）
                                     │
                     (add_crossref_hyperlinks.py 交叉引用跳转)
                                     │
                     (finalize_word_docx.py 更新字段/目录 + 导出 PDF)
                                     │
                     (validate_docx_conversion.py + validate_docx_render.py)
                                     │
                     (render_pdf_preview.py 预览拼图，人工/Agent 复核)
```

核心思路：**输入文件是内容源，Word 模板是排版源**。不要指望 pandoc 或
PDF 导入自己推断模板语义；先检查模板，再生成/改造一版项目专用后处理
脚本，最后用自动校验 + 视觉复核闭环交付。

## 🧰 环境要求

- Python ≥ 3.10
- 核心依赖：

  ```bash
  pip install python-docx pywin32 pymupdf pillow
  ```

  （`pywin32` 仅 Windows 需要；非 Windows 先装 `python-docx pymupdf pillow`）

- 可选依赖：

  ```bash
  pip install pdf2docx      # 仅当需要把原生 PDF 转成粗转 body 时
  ```

- [pandoc](https://pandoc.org/installing.html)：LaTeX/Markdown 粗转推荐安装。
- **完整流程（更新字段 + 导出 PDF）** 需要 Windows + 本机 Microsoft Word；
  inspect / pipeline / 校验（docx 层）/ 预览在 macOS / Linux 也可运行。

## 🚀 Quick Start

```bash
# 1. 克隆仓库
git clone https://github.com/<your-name>/TexPdf2Word.git
cd TexPdf2Word

# 2. 安装依赖
pip install python-docx pywin32 pymupdf pillow

# 3. 跑通自带最小示例（无需任何真实模板）
python examples/minimal_markdown/run_example.py
```

示例会现场生成一个演示模板，依次执行 inspect → 粗转 → adaptive pipeline
→ finalize → PDF 预览，产物落在 `examples/minimal_markdown/`。
Windows + Word 环境会额外得到 `final.pdf` 和 `final.preview.png`。

## 📋 完整操作步骤

以"LaTeX 论文项目 + 学校 Word 模板"为例（其余输入同理）：

### 第 1 步：检查 Word 模板

```bash
python skills/texpdf2word/scripts/inspect_docx_template.py template.docx --out template_report.json
```

输出 JSON 包含模板全部样式（含定义未使用的）、段落、表格、分节设置、
编号提示和超链接颜色。

### 第 2 步：生成粗转 body DOCX

```bash
# LaTeX（推荐；Markdown 去掉 --citeproc 相关参数即可）
pandoc main.tex --citeproc --reference-doc template.docx -o body.docx

# 原生 PDF：用 Word COM 导入或 pdf2docx；仅在没有源文件时才走 PDF
```

### 第 3 步：运行 adaptive pipeline 生成 final.docx

```bash
python skills/texpdf2word/scripts/adaptive_docx_pipeline.py \
    --template template.docx \
    --body-docx body.docx \
    --out final.docx \
    --config skills/texpdf2word/presets/chinese_thesis.json \
    --three-line-tables
```

- `chinese_thesis.json` 是**中文学位论文的通用示例配置**，请按你的模板
  修改 `body_style`、`unnumbered_h1`、字体等字段（详见脚本头部注释）。
- 非中文模板去掉 `--three-line-tables`，并使用自己的 config。
- ⚠️ **重要**：对真实学位论文模板，不要指望 starter 脚本开箱即用——
  它只做"安全的最小重建"。正确姿势是把
  `adaptive_docx_pipeline.py` **复制到工作目录后按模板改造**（删样例页、
  填封面字段、替换摘要、重建目录/分节），这一步通常由 AI Agent 按
  [SKILL.md](skills/texpdf2word/SKILL.md) 的工作流完成。

### 第 4 步：加交叉引用跳转（推荐）

```bash
python skills/texpdf2word/scripts/add_crossref_hyperlinks.py final.docx --link-citations
```

正文中的 `图N-M` / `表N-M` / `式(N-M)` / `[N]` 引用变成可点击的内部链接。
必须在 finalize **之前**运行。

### 第 5 步：Word 终处理（Windows + Word）

```bash
python skills/texpdf2word/scripts/finalize_word_docx.py final.docx --pdf
```

自动更新目录/页码等全部字段并导出 `final.pdf`。脚本会先禁用宏再打开文档。

### 第 6 步：两级自动质检

```bash
# 结构层
python skills/texpdf2word/scripts/validate_docx_conversion.py final.docx \
    --template template.docx --protected-until "中 文 摘 要" \
    --pdf final.pdf --out validation.json

# 渲染层（与源 LaTeX 目录比对图片/表格/caption/引用数量）
python skills/texpdf2word/scripts/validate_docx_render.py final.docx \
    --pdf final.pdf --source-latex-dir <latex-project> \
    --expected-table-style three-line --out validation_render.json
```

任何 FAIL 都要回到第 3 步修复后重跑；把真实模板的第一个生成页标记
（如 `中 文 摘 要`）替换成你模板里的实际标记。

### 第 7 步：视觉复核

```bash
python skills/texpdf2word/scripts/render_pdf_preview.py final.pdf --pages 1-8
```

生成预览拼图，人工或 AI 复核封面、目录、图表、公式、参考文献页面。

### 单脚本速查

| 脚本 | 用途 |
| --- | --- |
| `inspect_docx_template.py` | 检查模板样式/段落/表格/分节/编号 |
| `adaptive_docx_pipeline.py` | 可改造的 DOCX 重建起点（config 驱动） |
| `add_crossref_hyperlinks.py` | 图/表/式/引用内部跳转链接 |
| `inject_toc_field.py` | 幂等注入 TOC 字段 |
| `set_styleref_header.py` | 页眉改写为 STYLEREF 动态章节标题 |
| `finalize_word_docx.py` | Word COM 更新字段/目录 + 导出 PDF |
| `validate_docx_conversion.py` | 结构层质检 |
| `validate_docx_render.py` | 渲染层质检（12+ 项检查） |
| `render_pdf_preview.py` | PDF 预览拼图 |

### 作为 AI Skill 安装

把 skill 复制到你的 Agent skills 目录：

```bash
# macOS / Linux（Codex）
mkdir -p ~/.codex/skills
cp -r skills/texpdf2word ~/.codex/skills/

# Windows（PowerShell）
New-Item -ItemType Directory -Force -Path "$HOME\.codex\skills" | Out-Null
Copy-Item -Recurse -Force skills\texpdf2word "$HOME\.codex\skills\"
```

然后对 Agent 说：

```text
Use $texpdf2word to convert my LaTeX thesis into Word using this .docx template.
```

AI 会按 [SKILL.md](skills/texpdf2word/SKILL.md) 的工作流：检查模板 →
生成粗转 body → 改造项目专用 pipeline → 交叉引用 → finalize →
两级质检 → PDF 预览复核。

## 📁 项目结构

```
TexPdf2Word/
├── README.md                          ← 本文件
├── LICENSE / NOTICE / SECURITY.md / CONTRIBUTING.md
├── requirements.txt
├── skills/texpdf2word/                ← Skill 本体
│   ├── SKILL.md                       ← AI 工作流 + 质量门 + 常见坑
│   ├── agents/openai.yaml
│   ├── presets/chinese_thesis.json    ← 中文学位论文通用示例配置
│   ├── references/
│   │   ├── pandoc-limitations.md      ← 与 pandoc 默认行为的差异
│   │   └── thesis-case-study.md       ← 匿名化的真实论文模板实战复盘
│   └── scripts/                       ← 9 个确定性 Python 脚本（见速查表）
└── examples/minimal_markdown/         ← 最小可跑端到端示例
```

## ⚠️ 已知限制

- **finalize 阶段仅支持 Windows。** 更新字段/目录并导出 PDF 依赖
  Word COM（pywin32）。其余步骤跨平台。
- **不支持扫描版 / 纯图片 PDF。** PDF 输入需要原生数字文本层；
  请先 OCR，或直接用 LaTeX / Markdown 源文件。
- **模板含宏：宏会被禁用。** finalize 阶段强制
  `AutomationSecurity = msoAutomationSecurityForceDisable`。
- **starter pipeline 默认行为是保守的。** 三线表、正文字体覆盖等
  强规则全部 opt-in（config 或 CLI flag），默认不会偷改正文字体。
- **样式名识别覆盖中英文模板。** 日文/韩文等本地化模板可能需要在
  config 里补充 `unnumbered_heading_styles` 和 `body_candidate_styles`。

## 🔒 安全说明

- finalize 阶段会用本机 Word 打开你提供的 `.docx`，脚本默认禁用宏；
  请勿在不可信来源的输入上放宽该设置。
- Skill 工作流允许 AI 改写 `adaptive_docx_pipeline.py` 生成项目专用脚本，
  **运行前请人工 review diff**。自带脚本不做任何网络 I/O，
  只写用户显式指定的输出路径。
- 完整威胁模型见 [SECURITY.md](SECURITY.md)。

## 📄 许可证与致谢

本项目采用 [Apache License 2.0](LICENSE) 开源许可。

本项目基于开源项目 [docx-template-translator-skill](https://github.com/zouchenzhen/docx-template-translator-skill)
（Apache-2.0）整理与衍生而来，原项目版权信息保留于 [NOTICE](NOTICE)。
感谢原作者的开源贡献。
