---
name: westwell-ppt
description: Generate Westwell-branded PowerPoint presentations — from structured .md/outline inputs or from open-ended business questions. Combines McKinsey Problem Solving (Issue Tree / Hypotheses / Dummy Pages with page-dependency tracking / per-page evidence trail) with Westwell's corporate visual language and consulting-grade storyline (SCQA / BCG / Bain). Invoke whenever the user wants a PPT, presentation, slides, or deck — especially for Westwell product/solution proposals, market analyses, strategy decks, and cross-session resumptions of earlier decks.
---

# Westwell PPT Skill

Generate Westwell-branded `.pptx` decks. Two modes of operation:

- **Light mode** — user provides a structured `.md` or outline → go straight to Dummy + generate
- **Full mode** — user provides an open-ended business question → run McKinsey Problem Solving first (Issue Tree → Hypotheses → Dummy), then generate

Both modes share:
- **Visual style** — Westwell corporate template (large image · large text · large whitespace)
- **Narrative structure** — consulting-grade storyline (conclusion-first, MECE, "So What")
- **Builder** — `scripts/pptx_builder.py` · always use this, never raw python-pptx

---

## Quick Reference

```python
import sys
sys.path.insert(0, '~/.claude/skills/westwell-ppt')
from scripts.pptx_builder import WestwellPPT, preview_pptx

ppt = WestwellPPT(
    template='~/.claude/skills/westwell-ppt/assets/PPTTemplate.potx',
    output='/path/to/output.pptx'
)
# ... add slides ...
ppt.save()
preview_pptx('/path/to/output.pptx', '/tmp/westwell_preview/', dpi_scale=1.5)
```

---

## Reference Files (lazy-load on demand)

Keep this SKILL.md the only thing loaded by default. Pull in a reference **the first time** you enter the step that needs it, then let it fall out of context.

| File | When to load |
|------|-------------|
| `references/design-system.md` | Anytime you need colors, fonts, geometry |
| `references/visual-grammar.md` | Step 3 — 2-slide showcase, rhythm rules, Westwell-compatible style directions |
| `references/asset-protocol.md` | Step 3/5 — required logo/product/UI/site-photo asset protocol for concrete brands, customers, and products |
| `references/checklist.md` | Step 6 — final QA checklist + 5-dimension design review |
| `references/layouts-guide.md` | Step 3 (Dummy design) — narrative layouts (cover, chapter, bullets, stats, text_image, two_col, three_col, table, image, text, end) |
| `references/layouts-analytic.md` | Step 3 — analytic layouts (single chart, 2×2 matrix, waterfall, timeline, insight summary) for logic flow and data reasoning |
| `references/layouts-composition.md` | Step 3 — **composition layouts + editorial framing** (eyebrow/subtitle/footnote/bottom_callout/notes kwargs) (pyramid, value_chain, control_matrix, not_list, before_after, value_ladder, big_number, quote, number_list, step_grid) |
| `references/exemplar-strategy-memo.md` | Step 3 — **5 叙事节拍 for strategy memos** (核心结论 · 紧迫感 · 分层战略 · 能力边界 · Bottom line). Teaches the reasoning, not a page-count template. Load when user asks for a 战略思考 / 内部咨询 / 董事会 / IPO 内部讨论 deck |
| `references/methodology.md` | Phase 1 Step 1–2 — Is/Isn't boundary, MECE Issue Tree, Hypotheses (only if running Full mode) |
| `references/dummy-pages-spec.md` | Step 3 — Dummy.md format, page dependency types, dependency check dialogue |
| `references/data-collection.md` | Step 5 — per-page loop, 6-item self-check, data-trail.md format, context management |
| `references/troubleshooting.md` | Step 7 — when iterating on an issue |

All paths relative to `~/.claude/skills/westwell-ppt/`.

---

## Westwell Visual Design Principles

Derived from studying actual Westwell solution decks (airport, factory, seaport). These define what "Westwell quality" looks like — do not violate.

**① Titles: top-left, left-aligned, navy, ~22pt bold.** Every content slide title lives in the top-left corner, paired with "Make a WELL Change" in the top-right (inherited from template). The title IS the conclusion of the slide — not a topic label.

**② Big image · Big text · Big whitespace.** Westwell decks are "presentation-style", not "document-style". At most 2–3 information layers per slide. A full-bleed product photo with 2 lines of text beats 8 bullets. Never crowd.

**③ Dark slides for impact moments.** Dark blue (`#2743C6`) backgrounds for cover, chapter separators, statement slides, and occasionally KPI slides. They create rhythm and signal importance.

**④ Chapter separators = full-bleed background + centered text.** Full-bleed scene + dark overlay + chapter number + large title. Clean, dramatic, no decorative boxes.

**⑤ TOC = image background + rounded card grid.** Full-bleed scene photo + 2×3 grid of semi-transparent rounded cards. Active chapter highlighted in white.

**⑥ Content pages: title + horizontal rule + content.** Left-aligned title top, thin teal/navy horizontal rule beneath, then content (left-text + right-image, architecture diagram, bullet rows, or 2–3 column layout).

**⑦ Avoid these mistakes:**
- Never center a content slide title
- Never exceed ~60 words of body text on a single slide
- Never use a dense "colored header + 4 packed boxes" McKinsey layout — Westwell's aesthetic is airier
- Never add decorative shapes or borders that aren't inherited from the template
- **Never leave a dead band between the last content block and the template footer art at y≈5.933".** Content bottom should hug 5.85" (0.05–0.10" margin). A 0.3–0.5" empty strip below the callout/footnote reads as "this page isn't finished" even when `density='compact'`. If a column (e.g. `value_ladder` stage, `two_col` card) renders with obvious internal emptiness, prefer adding a short `body` line / extending the card over letting the space stay blank. See `design-system.md` → *Bottom-padding discipline* for the exact constants.

**Balance with analytic layouts:** McKinsey analytic layouts (2×2 matrix, waterfall, timeline, etc. — see `layouts-analytic.md`) are welcome **as logic tools**, but no more than 2–3 per chapter. Use `statement` / chapter separator / KPI `stats` slides to breathe between them.

**⑧ Visual Grammar Pass for 5+ page decks.** Before batch generation, create
two representative preview slides: one anchor slide (`cover` / `hero` /
chapter / statement) and one evidence slide (`big_numbers` / `lead_image` /
`pipeline` / `image_grid` / dense analytic). Render them, confirm the
grammar, then scale to the full deck. This avoids 12-slide rewrites caused by
wrong visual direction.

**⑨ Real assets beat generic decoration.** When a slide is about a concrete
customer, product, UI, vehicle, port, airport, or factory, collect or request
the actual logo / product image / UI screenshot / site photo. If the asset is
missing, use a clear placeholder and list it in data needs; never fill the
slot with unrelated stock or decorative AI imagery.

**Three Westwell-compatible style directions (layout only, not palette):**
- **Pentagram / Data Editorial** — data proof, benchmarks, market analysis;
  use strong grids, large numbers, `big_numbers()`, `rowlines()`, chart images.
- **Build / Executive Minimal** — board, strategy, and conclusion pages; use
  extreme restraint, `hero()`, `quote_editorial()`, sparse `text()`.
- **Takram / Soft Systems** — platform architecture, ecosystem, energy
  coordination; use calm systems diagrams, `pipeline()`, `lead_image()`,
  `image_grid()`.

---

## Consulting Storyline Principles

Draws from McKinsey, BCG, Bain — they differ in emphasis but share discipline.

### Universal rules (all three agree)

**Conclusion first, always.** Every slide title states the takeaway — what the audience should believe after seeing this slide. Reading only the titles in sequence should tell the whole argument.

| Weak (topic label) | Strong (conclusion title) |
|--------------------|--------------------------|
| 市场现状 | AI 调度在港口已进入规模商用,机场是下一个突破口 |
| 产品架构 | 四级决策路由在毫秒级响应与复杂推理之间精确分流 |
| 竞争格局 | 西井在端侧算力与场景专精上建立了不易复制的壁垒 |
| 实施计划 | 四阶段孵化路径:每个节点有明确的可测量成效 |

**MECE.** Every list (slides in a chapter, bullets in a slide) must be Mutually Exclusive, Collectively Exhaustive. If you can't make it MECE, restructure.

**"So What?" test.** After drafting a slide: "so what does this mean for the audience?" If the answer adds information not in the title, revise the title.

**Horizontal & vertical logic.** Titles in sequence = coherent story. Every bullet / visual on a slide must directly support the slide's title.

### Three firms, three framings (pick what fits)

**McKinsey SCQA** — use when you need to win buy-in from a skeptical audience. Lead them through Situation → Complication → Question → Answer.

**BCG value tree + 2×2 choice** — use when you need rigorous analytical framing. Decompose the core metric into drivers, then make an explicit strategic choice between alternatives. Implications pyramid: every data point → one implication → one so-what.

**Bain brutal simplicity + proof points** — use for action-oriented, time-compressed audiences. Every section ends with a crisp "what does this mean for you?". Claims need evidence: "We deployed X in Y scenario in 83 seconds with 93% task success rate" beats "we have strong capability in X".

Phase 1's methodology (Issue Tree → Hypotheses) produces the raw material; the firm you pick just shapes how that material is staged.

---

## Workflow

```
Phase 1 · Problem Solving  (Full mode only — skip if user gave structured input)
  Step 1   Define the boundary (Is / Isn't)
  Step 2   Issue Tree + Hypotheses                       → load methodology.md

Phase 2 · Design
  Step 3   Dummy Pages with page dependencies            → load dummy-pages-spec.md
                                                          + visual-grammar.md
                                                          + asset-protocol.md
                                                          + layouts-guide.md
                                                          + layouts-analytic.md
                                                          + layouts-composition.md  (strategy-memo composites)
                                                          + exemplar-strategy-memo.md  (ONLY if user asked for 战略/董事会/IPO 内部讨论)
            Output: <project>_DummyPages_<YYYYMMDD>.md
  Step 4   For 5+ page decks: generate 2-slide visual grammar preview;
           user confirms Dummy + visual grammar; choose generation mode A / B

Phase 3 · Generate & Deliver
  Step 5   Per-page loop (dep check → data → chart → slide → preview → pause)
                                                          → load data-collection.md
  Step 6   Deliver final .pptx + preview                  → load checklist.md
  Step 7   Iterate                                        → load troubleshooting.md (on demand)
```

---

## Entry Point — Which Mode?

Decide before you load anything:

**Full mode (Phase 1 + 2 + 3)** — trigger on:
- "帮我分析 X 市场 / 场景 / 机会"
- "给 X 做一份战略 / 方案 / 白皮书"
- "用 [方法论] 做一份 PPT"
- User has a question and no structure yet

**Light mode (Phase 2 + 3)** — trigger on:
- User uploads a structured `.md` or detailed outline
- "把这份材料做成 PPT"
- "按这个大纲生成"
- It's a product intro or periodic report with an existing template structure

If ambiguous, ask **one** clarifying question: "这是已有结构的材料转 PPT,还是需要我先帮你把问题拆清楚?" Then proceed.

**Cross-session resumption** — user uploads an existing `*_DummyPages_*.md`: treat as mid-flow. Read the Dummy, check the page dependency overview, ask what page to continue from, go directly into Step 5.

---

## Phase 1 · Problem Solving (Full mode)

### Step 1 — Define the boundary (Is / Isn't)

Output a short scope definition. No reference file needed.

```markdown
## 问题定义
### 是 ✅
- 核心问题: [一句话]
- 范围: [时间/地域/场景]
- 交付物: [PPT 页数、风格]
- 受众: [决策层/技术评审/运营]
### 不是 ❌
- [排除内容]
```

Ask the user to confirm, then continue.

### Step 2 — Issue Tree + Hypotheses

**First time entering this step**, load:

```python
Read("~/.claude/skills/westwell-ppt/references/methodology.md")
```

Then:
1. Build an MECE Issue Tree using the frameworks in methodology.md
2. Run 5–10 quick `web_search` calls to establish basic ground truth (plus pull from `~/Desktop/DatAi/` for internal context when relevant)
3. Form 4–8 hypotheses, each mapped to an Issue Tree node
4. Show the user the tree + hypotheses, get sign-off
5. **Release methodology.md from working context** before moving on

The Hypothesis Tree becomes the raw input for Phase 2's slide structure. Save a compact version in the project directory if the user wants to persist it separately (helpful for cross-session resumption of the executive summary).

---

## Phase 2 · Design

### Step 3 — Build the Dummy

**First time entering this step**, load:

```python
Read("~/.claude/skills/westwell-ppt/references/dummy-pages-spec.md")
Read("~/.claude/skills/westwell-ppt/references/visual-grammar.md")
Read("~/.claude/skills/westwell-ppt/references/asset-protocol.md")
Read("~/.claude/skills/westwell-ppt/references/layouts-guide.md")
Read("~/.claude/skills/westwell-ppt/references/layouts-analytic.md")
Read("~/.claude/skills/westwell-ppt/references/layouts-composition.md")
# ONLY load the exemplar if the user asked for a strategy memo / 董事会 / IPO 内部讨论:
Read("~/.claude/skills/westwell-ppt/references/exemplar-strategy-memo.md")
```

Then produce `<project>_DummyPages_<YYYYMMDD>.md` with:

1. **Project info** — title, date, total pages, audience, narrative mode (SCQA/BCG/Bain), core conclusion
2. **Page dependency overview** — explicit batches (第一轮 independent → 第二轮 forward-dependent → 第三轮 backward-dependent / last)
3. **Cross-session resumption note** — how to pick up where we left off
4. **Visual grammar table** — page rhythm, style direction, dark/light cadence,
   and which two pages will be generated as the preview showcase when total
   page count is 5+
5. **Every page** with: dependency label, slide type (mapped to builder method), content points, data needs, information sources
6. **Every page's four visual questions**: narrative role, audience distance,
   visual temperature, capacity estimate
7. **Asset needs** for concrete customers/products/scenes: logo, product/site
   photos, UI screenshots, diagrams, charts, and fallback if missing
8. **For non-independent pages**, also: 前置条件, 依赖页面, 缺失时对策

Use the 5 dependency labels from `dummy-pages-spec.md`:
```
✅ 独立    ⏩ 依赖前页    ⏪ 依赖后页    📄 需要文档    ⏸️ 最后生成
```

Pick slide types from **three** reference files:
- Narrative layouts → `layouts-guide.md` (cover, chapter, bullets, stats, text_image, etc.)
- Analytic layouts → `layouts-analytic.md` (single chart, 2×2 matrix, waterfall, timeline, insight summary)
- Composition layouts → `layouts-composition.md` (pyramid, value_chain, control_matrix, not_list, before_after, value_ladder, big_number, quote, number_list) — strategy-memo patterns, use for 战略 / 董事会 / IPO 内部讨论 type decks
- Visual grammar layouts → `layouts-guide.md` / `visual-grammar.md`
  (`hero`, `big_numbers`, `image_grid`, `pipeline`, `rowlines`,
  `lead_image`, `quote_editorial`) — use these to bring in
  Guizang/Huashu-style rhythm while staying PPTX-native and Westwell-branded

For a **strategy memo** specifically (战略思考 / 内部咨询版 / 董事会材料 / IPO 讨论稿), also load `exemplar-strategy-memo.md` — it gives you a battle-tested 18-page storyline (SCQA × 三层战略) you can adapt rather than designing from scratch.

Remember:
- **Slide titles must be论点, not topic labels** (see the Universal Rules table above)
- **Analytic layouts requiring charts need PNG pre-generation** — prefer the `diagram` skill for waterfalls / 2×2 / timelines; fall back to matplotlib with Westwell colors only if diagram doesn't cover the need. Charts go to `/tmp/westwell_charts/slide_XX.png`, then insert via `image()` or `text_image()`.
- **For 5+ page decks, do not batch-generate all slides before the 2-slide
  visual grammar preview is rendered and accepted.**
- **Missing assets must stay explicit** (`[ CUSTOMER LOGO ]`,
  `[ SITE PHOTO ]`, `[ UI SCREENSHOT ]`) until real assets are supplied.
- **Every slide type named in the Dummy must map to a real builder method**. If unsure, re-read `layouts-guide.md` or check `scripts/pptx_builder.py` directly.

Output the full Dummy as a file (not just inline), then release the reference files.

### Step 4 — Confirm & Choose Mode

Show the user the Dummy (or use `preview-md` skill to render it). Ask:

```
这份 Dummy 和依赖关系总览确认吗?

确认后请选择生成方式:

A) 逐页确认(推荐,首次使用或重要项目)
   每页生成后暂停,确认风格和内容再继续。

B) 一次性生成
   直接生成完整 PPTX。适合风格已确认后。

请选择 A 或 B:
```

**Do not write any slide code until the user confirms the Dummy.**

---

## Phase 3 · Generate & Deliver

### Step 5 — Per-page loop

**First time entering this step**, load:

```python
Read("~/.claude/skills/westwell-ppt/references/data-collection.md")
```

Then for each page, run the 9-step loop defined in `data-collection.md`:

```
0. 依赖检查 (read Dummy page header; walk through dep dialogue if needed)
1. 查看 Dummy 设计
2. 数据收集 (2–5 searches + internal docs → data-trail.md)
3. 图表生成 (if needed: diagram skill > matplotlib → /tmp/westwell_charts/)
4. 生成本页 PPT 代码
5. 预览本页
6. 自检 6 项
7. 告知用户 + 贴预览
8. 等待确认 (mode A) OR 继续 (mode B, pause only between chapters)
9. 清空本页搜索结果上下文
```

**Mode A** — pause after every page, show preview, ask "第 X 页确认,继续吗?"
**Mode B** — run through, but still pause between chapters for structural check-ins.

**Evidence discipline:** Every non-trivial data point in a slide must correspond to a row in `<project>_data-trail_<YYYYMMDD>.md`. If you can't cite it, downgrade it to qualitative. Never fabricate numbers.

**Script pattern:** Write a standalone `generate_<project>_<date>.py`:

```python
#!/usr/bin/env python3
import sys
sys.path.insert(0, '~/.claude/skills/westwell-ppt')
from scripts.pptx_builder import WestwellPPT, preview_pptx

TEMPLATE = '~/.claude/skills/westwell-ppt/assets/PPTTemplate.potx'
OUTPUT   = '/path/to/output.pptx'

ppt = WestwellPPT(template=TEMPLATE, output=OUTPUT)

# slides here — one per page in Dummy order
# (or dependency order if using the 3-batch generation strategy)

path = ppt.save()
print(f"Saved: {path}")
```

Preview command:

```bash
python3 -c "
import sys; sys.path.insert(0, '~/.claude/skills/westwell-ppt')
from scripts.pptx_builder import preview_pptx
preview_pptx('/tmp/output.pptx', '/tmp/westwell_preview/', dpi_scale=1.5)
"
```

### Step 6 — Deliver

1. Run final script → produce `.pptx`
2. Report output path
3. List the deliverables the user now has:
   - `<output>.pptx`
   - `<project>_DummyPages_<YYYYMMDD>.md`
   - `<project>_data-trail_<YYYYMMDD>.md`
   - `/tmp/westwell_charts/*.png` (if any)
4. Offer: `open <output.pptx>`

### Step 7 — Iterate

When the user asks for fixes, **only load troubleshooting on demand**:

```python
Read("~/.claude/skills/westwell-ppt/references/troubleshooting.md")
```

Focus the fix on the smallest surface that solves the complaint. Re-preview. If the change affects the structure, also update the Dummy and data-trail so the two stay in sync with the deck.

---

## Lazy Loading Policy

- **Default**: only this SKILL.md is in context
- **On entering a step**: load the file(s) listed for that step, use them, then **release**
- **Never preload** all references at the start — they waste tokens and crowd out slide content
- **Never re-read** a file you already used in the same step — keep a note of key decisions instead
- **Cross-session**: when picking up an existing project, re-read the Dummy and data-trail first, then load step-specific references only when you actually enter that step

This policy is what lets the skill scale to 20–30 page decks without burning context on documentation.

---

## Content Transformation Rules

(Shortened — the full rules live distributed across `methodology.md`, `data-collection.md`, and the storyline section above.)

**Insight titles, not topic labels** — see the Universal Rules table.

**One claim per slide** — if a section has 3 sub-arguments, that's 3 slides, not 3 bullets.

**Proof points over assertions** — replace "我们有丰富经验" with the actual number or case.

**Trim to speaker-supported text** — speaker-verbal content ≠ slide text. Each bullet ≤ 60 chars. Body ≤ 40 words / slide.

**Visual > text when possible** — comparison → `two_col`. Numbers → `stats`. Architecture / diagram → PNG + `image` / `text_image`. Process flow → analytic timeline or `bullets` with icons.

**Chapter rhythm** — every 3–5 content slides needs a chapter separator. Audiences lose track without visible structure.

---

## File Paths

```
~/.claude/skills/westwell-ppt/
├── SKILL.md                        ← this file
├── assets/
│   └── PPTTemplate.potx            ← Westwell corporate template
├── scripts/
│   └── pptx_builder.py             ← WestwellPPT builder class
└── references/
    ├── design-system.md            ← colors, fonts, geometry
    ├── layouts-guide.md            ← narrative layouts (Westwell built-ins)
    ├── layouts-analytic.md         ← analytic layouts (McKinsey-style, logic flow)
    ├── layouts-composition.md      ← strategy-memo composites (pyramid, value_chain, control_matrix, not_list, before_after, value_ladder, big_number, quote, number_list)
    ├── exemplar-strategy-memo.md   ← 5 叙事节拍 for strategy memos (思路,not 页数模板)
    ├── methodology.md              ← Problem Solving (Is/Isn't, Issue Tree, Hypotheses)
    ├── dummy-pages-spec.md         ← Dummy.md format + page dependency system
    ├── data-collection.md          ← per-page loop, self-check, data-trail
    └── troubleshooting.md          ← common issues & fixes
```

Project-level artifacts (written by Claude while running this skill, live next to the user's PPT):

```
<project-dir>/
├── <project>.pptx
├── <project>_DummyPages_<YYYYMMDD>.md
└── <project>_data-trail_<YYYYMMDD>.md
```

---

## Dependencies

- `python-pptx` — slide generation
- `LibreOffice` (`soffice`) — `.potx` → `.pptx` conversion
- `pymupdf` (`fitz`) — PPTX → PDF → PNG preview
- Optional: `matplotlib` / `plotly` for analytic charts (or use the `diagram` skill)

Sanity check:
```bash
soffice --version && python3 -c "import pptx, fitz; print('OK')"
```

---

## Known Issues & Fixes

Short list — full catalog lives in `references/troubleshooting.md`:

- **Cover "click to add text"**: Remove all layout placeholders before adding content as free textboxes. Title y=1.70", subtitle area y=4.60".
- **`agenda()` with mixed list**: Accepts plain strings (auto-numbers 01, 02…) OR `(num, title, subtitle)` tuples. Don't mix.
- **Analytic layouts without chart PNGs**: All of §1, §3, §5, §6 in `layouts-analytic.md` require you to pre-generate a PNG (via `diagram` skill or matplotlib) and insert via `image()` / `text_image()`. The builder does not generate charts on its own.
