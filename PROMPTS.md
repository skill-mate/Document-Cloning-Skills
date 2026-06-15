# Prompt Templates · docx-format-clone

How to talk to Claude when you want this skill to do the work.

The skill is triggered by **keywords**, not by a specific prompt format —
any natural-language request that contains a trigger word will activate it.
Below are copy-paste templates for the three most common scenarios, so you
don't need to think about wording every time.

> **Triggers** (any one is enough): 套用模版 · 套用样板 · 复刻格式 · 复制版式 ·
> 严格按 X 格式 · 按上次的版式 · 政府公文 · 投标书 · 技术方案 · 标书 · 公文格式 ·
> GB/T 9704 · 目录单独 · 表格不跨页 · 表头重复 · 另起一页 · 生成 word · 生成 docx ·
> 导出 word · 导出 docx

---

## Template A — Skeleton mode (clone an existing .docx)

The most precise mode. The reference `.docx` is treated as a container:
its styles, numbering, theme, header / footer and page setup are kept
intact; only the body is replaced.

### Fill-in-the-blank

```
请用 docx-format-clone skill，骨架模式，
【样板】<sample.docx 的绝对路径>
【需求】<内容描述：要换什么 / 保留什么 / 重点强调什么>
【封面】title=<>，subtitle=<>，org=<>，date=<>
【输出】<output.docx 的绝对路径>
```

### Concrete example

```
请用 docx-format-clone skill，骨架模式生成一份新投标书。
【样板】/Users/me/work/上一份投标书.docx
【需求】投标方改成"东风一汽"，项目名改成"东风一汽问题信息管理系统软件项目"，
合同金额改成 580 万，工期 6 个月，其他章节结构与上次一致；
技术方案部分补充一段微服务架构说明（约 400 字）。
【封面】title=东风一汽问题信息管理系统软件项目投标书，subtitle=投标方案，
org=广州智能汽车科技有限公司，date=2026年6月15日
【输出】~/Desktop/东风一汽投标书.docx
```

### Variations

- **沿用上次输出做新版**：`【样板】<上一次生成的 .docx>` — 输出会继续延续之前的版式
- **只改局部**：`【需求】只把第三章改成 XXX，其他章节保持原样` — Claude 会读旧版、保留没动的章节
- **多份连发**：分别给三家公司各做一份 → 在【需求】里说"分别给 A/B/C 三家公司各生成一份，A 改成…，B 改成…，C 改成…，输出三个文件"

---

## Template B — Profile mode (no sample, use a built-in preset)

When you don't have a reference `.docx`, pick one of two presets:

| Preset | 适用 |
|---|---|
| `gov_gongwen` | 政府/事业单位公文，GB/T 9704，仿宋_GB2312 三号、固定行距 28pt |
| `tech_doc` | 投标书 / 技术方案 / 白皮书，宋体 + 黑体，1.5 倍行距 |

### Fill-in-the-blank

```
请用 docx-format-clone skill，预设 <gov_gongwen | tech_doc>，
生成一份 <文档类型>。
【大纲】
1. <章节1>
2. <章节2>
   2.1 <小节>
   2.2 <小节>
3. <章节3>
【封面】title=<>，org=<>，date=<>
【输出】<output.docx 的绝对路径>
```

### Concrete example — government notice

```
请用 docx-format-clone skill，预设 gov_gongwen，生成一份关于办公室搬迁的通知。
【正文要点】
- 自 2026 年 7 月 1 日起，办公室从 XX 路 100 号搬迁至 YY 路 200 号
- 各部门联系电话不变
- 7 月 3 日前完成搬迁，期间业务不中断
- 联系人：王主任 13800138000
【封面】title=关于办公室搬迁的通知，org=XX 有限公司，date=2026年6月30日
【输出】~/Desktop/搬迁通知.docx
```

### Concrete example — technical proposal

```
请用 docx-format-clone skill，预设 tech_doc，生成一份某市智慧交通技术方案。
【大纲】
1. 项目背景
2. 技术架构（云原生 + 微服务）
3. 核心模块（信号优化 / 车流监测 / 应急指挥）
4. 实施进度（分三阶段，共 12 个月）
5. 团队配置
6. 报价（先空着）
【封面】title=XX 市智慧交通建设技术方案，org=ABC 智能科技，date=2026年6月20日
【输出】~/Desktop/智慧交通方案.docx
```

---

## Template C — Let Claude write the content first, then render

When you only have an outline and want Claude to draft the actual text
before producing the docx.

```
请用 docx-format-clone skill 骨架模式生成一份 <文档类型>。

【样板】<sample.docx 的路径>
【输出】<output.docx 的路径>

【内容要求】
- 章节大纲：<列出来>
- 风格：<专业 / 正式 / 通俗>
- 重点强调：<核心卖点 / 关键数据>
- 长度：<每章约 X 字>

请先帮我把每章正文写出来，确认无误后再调用 skill 生成 .docx。
```

This pattern lets you review the prose before it's locked into formatted
output, which saves a regeneration round-trip.

---

## Iteration prompts (after the first generation)

Once Claude has produced a `.docx`, point at the bits to fix:

```
第三章再扩到 800 字，强调云原生部署优势；
表格 2-1 第三行改成"完全响应：采用 Spring Cloud Alibaba 微服务架构"；
封面 date 改成 2026 年 6 月 18 日。
重新生成。
```

```
把所有的"乙方"统一改成"承建方"，重新生成到同一个路径。
```

```
按你刚生成的那份 docx 当骨架，再做一份给"南京 XX 集团"的版本。
```

---

## Tips

1. **Path always absolute** — `/Users/...` 或 `~/...` 都行；不要写 `./X.docx` 或纯中文相对路径。
2. **Sample must be `.docx`** — `.doc` 老格式 / PDF 都不能当样板；先用 Word/WPS 另存为 `.docx`。
3. **You don't write Markdown** — Claude writes the intermediate `input.md`
   internally, then runs the build script. You see only the final `.docx`.
4. **Don't ask "is this OK?" after the prompt** — once you trigger the skill
   it will run end-to-end. Reviewing happens after, not before.
5. **Cover fields are optional** — omit any of `title / subtitle / org / date`
   and that line is skipped on the cover page.
6. **Custom column widths**: in your prompt, mention "表格列宽 2:3:5:2" and
   Claude will translate it into the right `<!-- table-widths: ... -->` directive.

---

## One-liner quick reference

```
用 docx-format-clone skill，[骨架 | 预设 gov_gongwen | 预设 tech_doc] 模式，
[样板=<path> | 无样板]，写 [内容]，输出到 [path]。
```

Fill the four slots and Claude takes it from there.
