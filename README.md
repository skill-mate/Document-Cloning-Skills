# Document-Cloning-Skills · docx-format-clone

A [Claude Code](https://docs.claude.com/en/docs/agents-and-tools/claude-code/skills)
skill that renders Markdown into **strictly-formatted Word documents**.

Built for repetitive high-format work: tenders, government documents
(GB/T 9704), technical proposals, white papers — where you want the next
document to look identical to the previous one, with only the content
changed.

## Two modes

| Mode | Script | Use when |
|---|---|---|
| **Skeleton** (most precise) | `scripts/build_from_skeleton.py` | You have a reference `.docx` whose layout you want to clone byte-for-byte. The skeleton mode keeps `styles.xml` / `numbering.xml` / `theme/` / `header*.xml` / `footer*.xml` / `settings.xml` / page setup intact and **replaces only the body**. |
| **Profile** | `scripts/build_docx.py` | No reference doc. Pick a built-in preset (`gov_gongwen` / `tech_doc`) or feed it a `profile.json` extracted from a sample. |

## Hard guarantees (both modes)

The renderer always enforces — at the OOXML level, not via Word
auto-formatting — these layout rules:

- Cover page in its own section, vertically centered, no page number
- Table of contents in its own section with a Word `TOC` field (auto-update on open)
- Every H1 forces a page break
- Tables: every row gets `w:cantSplit` (row never spans pages); first row gets `w:tblHeader` (repeats on each new page); table is `tblLayout=fixed` so Word doesn't reflow it
- Chinese / Western fonts set independently per run (`rFonts/eastAsia` + `rFonts/ascii`) — no Word font fallback
- Page margins / size from profile or reference

## Install

```bash
git clone https://github.com/skill-mate/Document-Cloning-Skills.git \
  ~/.claude/skills/docx-format-clone
```

Requires Python 3.9+ and `python-docx`:

```bash
pip3 install python-docx
```

## Usage inside Claude Code

Just describe the task; Claude picks up the skill from its trigger keywords:

> 严格按 上一份投标书.docx 的格式，生成一份新公司的投标书

> Use the skeleton mode to clone last_proposal.docx and write a new tender for ABC Corp

## Direct CLI

```bash
# Skeleton mode — clone the reference layout
python3 scripts/build_from_skeleton.py input.md \
  --reference sample.docx -o output.docx

# Profile mode — built-in preset
python3 scripts/build_docx.py input.md \
  --preset gov_gongwen -o output.docx
# or:  --preset tech_doc

# Profile mode — custom profile extracted from a sample
python3 scripts/extract_style.py sample.docx -o profile.json
python3 scripts/build_docx.py input.md \
  --profile profile.json -o output.docx
```

## Markdown conventions

Standard Markdown (headings / paragraphs / lists / tables / images / blockquotes
/ fenced code) plus a few HTML-comment directives the renderer understands:

| Directive | Effect |
|---|---|
| `<!-- cover ...fields... -->` | Cover page (own section, vertically centered) |
| `<!-- toc depth=3 -->` | TOC page (own section, with Word `TOC` field) |
| `<!-- pagebreak -->` | Hard page break |
| `<!-- newsection -->` | Hard section break (for different headers/footers) |
| `<!-- table-caption: 表 1-1 ... -->` | Caption for the next table |
| `<!-- table-widths: 2,3,5,2 -->` | Relative column widths for the next table |
| `<!-- image-caption: 图 1 ... -->` | Caption for the next image |

See `examples/sample.md` for a full example.

## Repo layout

```
.
├── SKILL.md                    # Skill manifest + trigger keywords (read by Claude Code)
├── presets/
│   ├── gov_gongwen.json        # GB/T 9704 government document
│   └── tech_doc.json           # Tender / technical proposal / white paper
├── scripts/
│   ├── build_from_skeleton.py  # Skeleton mode — reference.docx + md → docx
│   ├── build_docx.py           # Profile mode — md + profile/preset → docx
│   └── extract_style.py        # reference.docx → profile.json (per-level)
└── examples/
    └── sample.md
```

## License

MIT — see [LICENSE](LICENSE).
