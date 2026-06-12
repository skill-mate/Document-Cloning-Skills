---
name: docx-format-clone
description: 为强格式重复型文档（投标书、政府公文、技术方案、白皮书）生成 Word。提供两种模式：（1）骨架模式 build_from_skeleton.py — 用样板 docx 当容器，保留全部 styles/numbering/页眉页脚/页面设置不动，只换正文，最精确；（2）Profile 模式 build_docx.py — 用预设或抓取的 profile.json 渲染，无需样板。流程对用户无感：内部先落 Markdown，再调脚本输出 .docx，强制实现「目录单独成节 / 一级标题另起一页 / 表格 cantSplit 不跨页 / 跨页表头重复 / 中英文字体分别设置」。当用户要求「严格按 XX 格式生成 / 复制 XX 模版做一份 / 套用上次的版式」时触发，**首选骨架模式**。
triggers:
  - "套用模版"
  - "套用样板"
  - "复刻格式"
  - "复制版式"
  - "严格按.*格式"
  - "按上次的版式"
  - "和.*格式一样"
  - "政府公文"
  - "投标书"
  - "技术方案"
  - "技术文档"
  - "标书"
  - "公文格式"
  - "GB/T 9704"
  - "目录单独"
  - "表格不跨页"
  - "表头重复"
  - "另起一页"
  - "生成word"
  - "生成docx"
  - "生成.*Word"
  - "导出word"
  - "导出docx"
allowed-tools: Read, Write, Edit, Bash
version: 1.0.0
---

# docx-format-clone — 强格式 Word 文档生成器

把 Markdown 严格、可重复地渲染成符合既定版式的 Word，专治三类高频场景：

1. **复刻型**：用户给一份 `.docx` 样板，下次新文档要"格式完全一样、只换名称和部分内容"。
2. **公文型**：政府/事业单位行文，`GB/T 9704-2012` 标准（仿宋_GB2312 三号、行距 28pt 固定、首行缩进 2 字符…）。
3. **方案型**：投标书/技术方案/白皮书，封面 + 目录 + 多部分章节 + 表格 + 图片。

## 工作流（用户无感）

不要把 md 中间产物当成最终交付物给用户看。最终交付一律是 `.docx`。

```
用户描述需求
   │
   ▼
[1] 决定格式来源 ──────► 用户给 reference.docx? ──► extract_style.py → profile.json
   │                       否
   │                       └──► 选预设：gov_gongwen.json / tech_doc.json
   ▼
[2] 写中间产物 input.md（含本 skill 自定义指令）
   │
   ▼
[3] 调 build_docx.py input.md --profile X.json -o output.docx
   │
   ▼
[4] 把 output.docx 路径告诉用户；不主动把 md 推给用户看
```

## 第一步：决定格式来源（**只问一次**）

- 用户已经说了"按上次/某个文件的格式"——**优先用骨架模式**（3.A），直接 reference 那份 docx，不需要 extract。
- 用户没说样板 ——根据语境默认走 Profile 模式（3.B）：
  - 政府/单位/红头/公文/通知/报告 → `presets/gov_gongwen.json`
  - 投标书/技术方案/白皮书/产品文档 → `presets/tech_doc.json`
  确定后**直接用**，不要反问"用哪个"，除非两者都明显不合适。
- 用户说"复刻一份字段配置以后能调"，再用 extract_style.py 生成 profile.json。

## 第二步：写 Markdown（带 skill 指令）

把内容落在 `<工作目录>/.dumate/inbox/<文档名>.md`（或类似路径，用户感知不到）。
本 skill 在标准 Markdown 之上识别下列**指令注释**，其它 Markdown 语法（标题/段落/列表/表格/图片/引用）按常规即可：

| 指令 | 作用 | 用法位置 |
|---|---|---|
| `<!-- cover ... -->` | 封面页（独立一节，正中排版），单 `-->` 内换行写字段 | 文档开头 |
| `<!-- toc depth=3 -->` | 目录页（自动单独成节，前后分页），depth 控制几级 | 封面之后 |
| `<!-- pagebreak -->` | 强制分页 | 任意位置 |
| `<!-- newsection -->` | 强制分节（用于页眉页脚不同） | 任意位置 |
| `<!-- table-caption: 表1 XXX --><br>(紧跟一个表)` | 给下一张表加题注，自动 cantSplit + tblHeader | 表前一行 |
| `<!-- image-caption: 图1 XXX -->` | 给下一张图加题注 | 图前一行 |
| `# 标题` 一级标题 | **默认自动在前面分页**（profile 里 `auto_h1_pagebreak: true`） | 任意 |

封面 cover 块字段示例：

```markdown
<!-- cover
title: 广汽一汽问题信息管理系统软件项目投标书
subtitle: 投标方案
org: 广州智能汽车科技有限公司
date: 2024年1月15日
-->
```

## 第三步：渲染 docx

有**两条路**，agent 按下面的优先级选：

### 3.A 骨架模式（最精确，**优先用**）

当用户给了样板 docx，且强调"格式完全一致"，用 build_from_skeleton.py。
它把样板当容器：保留全部 styles.xml / numbering.xml / theme/ / header*.xml / footer*.xml /
settings.xml / 节属性（页面/页眉页脚引用）—— 一字不改，只换正文。

```bash
python3 ~/.claude/skills/docx-format-clone/scripts/build_from_skeleton.py \
  /path/to/input.md \
  --reference /path/to/sample.docx \
  -o /path/to/output.docx
```

特点：
- 字体/字号/字色/段距/行距/对齐/编号定义/页眉红头/页脚水印/页边距，全部继承样板
- Markdown 标题按"Heading N / 标题 N"映射到样板的同名段落样式（首匹配）
- 列表自动绑到 `List Number / List Bullet / List Paragraph`（按样板里有哪个用哪个）
- 表格优先绑 `Table Grid` 等表格样式；样板里没有就**手动加单线黑边**保底

### 3.B Profile 模式（无样板时用）

```bash
python3 ~/.claude/skills/docx-format-clone/scripts/build_docx.py \
  /path/to/input.md \
  -o /path/to/output.docx \
  --profile /tmp/profile.json     # 或 --preset gov_gongwen / tech_doc
```

适用：没有样板，或样板已经被 extract_style.py 抓成 profile.json。

### 二者共同强制保证

build_*.py 都会在 XML 层强制做到：

- **封面**单独一节，垂直居中（section 设 `valign=center`），前后分页。
- **目录**单独一节，使用 Word `TOC` 域（用户首次打开按 F9 / 自动更新）；目录前后强制分页。
- **每个 H1 之前自动分页**（profile.auto_h1_pagebreak=true 时）。
- **表格**：每行写入 `w:cantSplit`（行不跨页）；表头行写入 `w:tblHeader`（跨页时自动重复表头）；表格设 `w:tblLayout=fixed` + 列宽，避免 Word 重排破版。
- **中英文字体分设**：每个 run 同时设置 `rFonts.ascii / hAnsi`（西文）和 `rFonts.eastAsia`（中文），杜绝 Word 字体回退。
- **段落格式逐级复刻**（v1.1）：profile 中每级（title/h1/h2/h3/h4/body/table/table_header）都带完整字段——字体、字号、加粗、字色、对齐、首行缩进、左右缩进、段前段后、行距规则与值、与下段同页（keep_with_next）。extract_style.py 从样板抓真实数值，build_docx.py 渲染时**优先使用 `levels[role]`**，缺字段才回退到预设。
- **页眉页脚**：profile 配置；默认页脚居中页码，封面/目录页码独立。
- **页边距**：从 profile 取（公文默认 上 3.7cm 下 3.5cm 左 2.8cm 右 2.6cm；技术文档 2.54cm 四周）。

## 文件清单

```
~/.claude/skills/docx-format-clone/
├── SKILL.md                          # 本文件
├── presets/
│   ├── gov_gongwen.json              # 政府公文（GB/T 9704）
│   └── tech_doc.json                 # 技术方案/投标书
├── scripts/
│   ├── build_from_skeleton.py        # 【骨架模式】reference.docx + md → docx（最精确）
│   ├── build_docx.py                 # 【Profile 模式】md + profile → docx
│   └── extract_style.py              # reference.docx → profile.json（v1.1，逐级）
└── examples/
    └── sample.md                     # 完整示例 md（可作为模板起点）
```

## 调用约定（给 agent）

1. **不要让用户看到 md 中间步骤**。把 md 写到不显眼的临时位置即可，最终只把 `.docx` 路径告诉用户。
2. 用户说"复刻 / 套用 / 严格按 XX.docx" 时，**首选 build_from_skeleton.py**（骨架模式），把那份 docx 当 `--reference`。这是最精确的复刻方式——样式表/编号/页眉页脚原封不动。
3. 用户没指定字体/字号、又没给样板时，按语境选 `gov_gongwen` / `tech_doc` 预设走 Profile 模式，**不要反问**。
4. 写 md 时表格放标准 Markdown pipe table 即可，build 脚本会自动套 cantSplit + tblHeader。
5. 表格列宽如果有诉求，加 `<!-- table-widths: 2,3,5,2 -->`（数字是相对比例）。
6. 大段公文条款建议用 `> ` 引用块；条文里的"第X条"用 `### 第一条` 等标题层。
7. 生成完调一次 `python3 -c "from docx import Document; Document('out.docx')"` 自检，确认文件没坏。

## 失败模式与回退

- 表格单行内容超过一页：`cantSplit` 会被 Word 忽略，强制拆分（这是 Word 正常行为，无解，提醒用户拆行）。
- 用户给的样板里嵌了图片或公式：extract_style 只抓"样式"不抓内容，拷贝样式不会带走图片，提醒用户图片需另行提供。
- profile.json 里某个字段缺失：build_docx 走该字段的内置回退（中文宋体 / 西文 Calibri / 小四 / 单倍行距），并在 stderr 打 `[fallback] xxx`，不阻断生成。

## 不要做的事

- 不要把 md 当成最终交付物展示给用户，除非用户明确说"我要 md"。
- 不要在 build 之前问"你确认要生成吗"——已经触发本 skill，就直接生成。
- 不要凭记忆复刻样板格式。永远先 extract_style.py 抓真实数值。
- 不要修改 `presets/*.json`（那是出厂默认）；用户要改格式，写到 profile.json 副本里改。
