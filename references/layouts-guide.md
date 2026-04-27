# Westwell PPT Layouts Guide

When to use each slide type in `pptx_builder.py`, with content guidelines.

---

## Slide Type Decision Tree

```
Is this the first slide?
  → cover()

Is this a visual grammar preview, chapter thesis, or executive anchor?
  → hero()

Is this a section divider between chapters?
  → chapter()

Is this the last slide?
  → end()

Is this a table of contents?
  → agenda()

Does the slide feature a prominent number/metric/KPI?
  + classic Westwell KPI cards → stats()
  + data-editorial / benchmark proof → big_numbers()

Does this slide need an executive quote or bottom-line pause?
  → quote_editorial()

Does the slide have an image as the main content?
  + with a text explanation beside it → lead_image() or text_image()
  + with a text explanation above it  → image_below()
  + image fills the whole content area → image()
  + multiple real assets / screenshots → image_grid()

Does this slide show a process or implementation sequence?
  → pipeline()

Does this slide list principles, protocols, or findings in rows?
  → rowlines()

Does the slide have two parallel topics to compare?
  → two_col()

Does the slide have three parallel categories?
  → three_col()

Does the slide have a list of 3–7 points?
  → bullets()

Does the slide have a data table?
  → table()

Does the slide have flowing prose or a single insight?
  → text()
```

---

## Visual Grammar Layouts

These layouts absorb Guizang/Huashu editorial discipline while staying
Westwell-branded and PPTX-native. They are preferred during the 2-slide
visual grammar pass.

## `hero(title, kicker='', lead='', dark=True, variant='chapter', meta_left='', meta_right='', foot_left='', foot_right='')`

**When:** Visual grammar preview, chapter thesis, executive anchor, or a major
strategic pivot.

**Use:** 1 strong title, optional short kicker, optional lead. Keep body text
off the slide unless it is a one-line thesis.

## `big_numbers(title, metrics, kicker='', lead='', dark=False)`

**When:** Benchmark, KPI, market proof, adoption results.

**Input:** `metrics=[(value, label, note), ...]`, 2-4 items.

**Rules:** Values must be short and memorable. Notes should be evidence lines,
not explanations. Use `stats()` when you want classic Westwell KPI cards;
use `big_numbers()` when the number itself is the visual anchor.

## `pipeline(title, steps, kicker='', lead='', dark=False)`

**When:** Rollout path, product workflow, platform logic, capability maturity.

**Input:** `steps=[{'title': '...', 'body': '...'}, ...]`, 3-6 steps.

**Rules:** Each step title is a noun phrase or short verb phrase. Body stays
under 25 Chinese characters where possible.

## `rowlines(title, rows, kicker='', lead='', dark=False)`

**When:** Asset protocol, design principles, key findings, risks, decisions.

**Input:** `rows=[(key, value, meta), ...]` or dicts with `key/value/meta`.

**Rules:** Use rows when card grids feel too generic or too decorative.

## `quote_editorial(text, attribution='', title='', kicker='', dark=True)`

**When:** Executive pause, quote, bottom-line, "so what" moment.

**Rules:** The quote should fit in 1-3 lines. Use it to create rhythm, not as
filler.

## `lead_image(title, lead, img_path=None, kicker='', caption='', dark=False)`

**When:** A single photo, UI screenshot, product image, or diagram is the proof.

**Rules:** Missing image shows an explicit placeholder. Real images should
preserve top-critical information.

## `image_grid(title, images, kicker='', lead='', dark=False)`

**When:** Multiple assets prove the point: site photos, UI states, product
details, before/after screenshots.

**Input:** `images=[{'path': '/abs/file.png', 'label': 'Dispatch UI'}, ...]`.
2×2 is best; 3×2 is the maximum.

---

## `cover(title, subtitle, context, date)`

**When:** First slide only.

**Layout:** `标题幻灯片` — dark blue background, large right-side art composition.

> **Implementation note:** Remove all layout placeholders before adding content as free textboxes — otherwise empty placeholders show "Click to add text". Title y=1.70", subtitle area y=4.60".

**Parameters:**
- `title`: The deck title — concise, max 2 lines (e.g. "Westwell Airport\nIntelligent Scheduling")
- `subtitle`: One-line subtitle (e.g. "ATC Agent Product Proposal")
- `context`: Client or context line, shown small (e.g. "Westwell Technology · Airport Solutions")
- `date`: Presentation date (e.g. "2026年4月")

**Content rules:** Title is always white, left-aligned in the left half only (right side reserved for the WMF art decoration).

---

## `chapter(num, title, subtitle, highlight)`

**When:** Section separator between major chapters. Do not use for content.

**Layout:** `agenda slide 2` — dark blue background, colorful right-side geometric art. Creates a visual pause and signals a new narrative chapter.

**Parameters:**
- `num`: Chapter number string (e.g. `"01"`, `"02"`)
- `title`: Chapter title (e.g. `"Problem & Opportunity"`) — supports `\n` for two lines
- `subtitle`: Optional supporting line in smaller text
- `highlight`: Optional one-line key insight/teaser shown in teal

**Content rules:** Keep chapter title to max 2 lines. The large art on the right is the visual anchor — don't compete with it.

---

## `agenda(items)`

**When:** Table of contents / roadmap slide. Typically slide 2 after cover.

**Layout:** `agenda slide 1` — dark blue background, teal circle number badges on left.

**Parameters:**
- `items`: List of strings; each becomes a numbered row (e.g. `["Problem & Opportunity", "Solution Architecture", "Timeline"]`)

**Content rules:** 3–6 items ideal. Each item is one line. Don't nest sub-items here.

---

## `text(title, body, dark=False)`

**When:** A slide with a single, flowing explanation — no list structure needed.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `title`: Insight-first title (the WHAT, not the topic label)
- `body`: Prose text at 20pt. Can be 2–4 sentences. Use `\n\n` for paragraph breaks.
- `dark`: Use dark variant for emphasis slides or when following a chapter separator

**Content rules:** Body text is 20pt, max ~200 characters. If you have 5+ discrete points, use `bullets()` instead.

---

## `bullets(title, points, dark=False, body_size=18)`

**When:** A list of parallel, discrete points — decisions, features, criteria, findings.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `title`: Insight-first title
- `points`: List of strings, 3–7 items. Each item is one bullet line.
- `body_size`: Font size for bullet text (default 18pt, can increase to 20pt for short lists)

**Content rules:**
- Each bullet should be one idea, max ~80 characters
- Teal `■` marker is added automatically — don't add `•` in the text strings
- If points have sub-bullets, combine into one line or use `text()` with prose instead
- More than 7 bullets → split into two slides

---

## `table(title, headers, rows, dark=False, bold_col=None)`

**When:** Comparative data, specifications, a feature matrix, or structured facts.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `headers`: List of column header strings (3–5 columns ideal)
- `rows`: List of row lists
- `bold_col`: Index (0-based) of a column to bold for visual emphasis (e.g. `0` for row labels)

**Content rules:**
- Table header row uses navy background + white bold text
- Rows alternate `C_WHITE` / `C_ALTROW` for readability
- Max 6–8 rows before the table becomes unreadable
- Keep cell text concise: one value or short phrase per cell, not sentences

---

## `stats(title, stats, dark=False)`

**When:** 2–4 key metrics or KPIs deserve to be the focal point of a slide.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `stats`: List of `(value, label, description)` tuples, e.g.:
  ```python
  [("95%", "调度准确率", "AI建议采纳率超过业界均值30%"),
   ("<1s",  "决策响应时间", "四级路由最快路径50ms"),
   ("4阶段", "试点计划",   "6个月完成全场景覆盖")]
  ```

**Content rules:**
- Value: big number, ratio, or brief metric — rendered huge in teal
- Label: short name for the metric (2–6 chars ideally)
- Description: one supporting sentence (~30 chars)
- 2–4 stats per slide; 3 is the visual sweet spot

---

## `image(title, img_path, caption, dark=False)`

**When:** An image IS the content — a screenshot, diagram, map, or photo. Text is secondary.

**Layout:** `custom slide1-light` or `custom slide1-dark`. Image fills the content safe zone.

**Parameters:**
- `img_path`: Absolute path to image file
- `caption`: Short description below the image

**Content rules:** The image should be clean and high-resolution. Caption is small (12pt) and positioned below the image.

---

## `image_left(title, body, img_path, dark=False)`

**When:** You need to explain an image with supporting text — equal importance between text and image.

**Layout:** `content1-text&image` — left half text, right half image.

**Parameters:**
- `body`: Explanation text for the left column (2–4 sentences or a short bulleted list with `\n` between items)
- `img_path`: Absolute path to image

**Content rules:** Keep body text concise — the left column is narrow (~5.5"). Text is 18pt. This layout works well for architecture screenshots with annotations, or UI previews with feature callouts.

---

## `image_below(title, body, img_path, dark=False)`

**When:** A brief lead-in text introduces an image — text is secondary to the visual.

**Layout:** `content2-text&image` — top-third text, bottom-two-thirds image.

**Parameters:**
- `body`: Short 1–2 sentence lead-in
- `img_path`: Absolute path to image

**Content rules:** Body text area is shallow — keep to 1–2 lines at 18pt. The image dominates this layout.

---

## `two_col(title, left_head, left_body, right_head, right_body, dark=False)`

**When:** Comparing or contrasting two parallel concepts — before/after, option A/B, current/future.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `left_head` / `right_head`: Column heading in navy (or white on dark)
- `left_body` / `right_body`: Body text for each column, 18pt

**Content rules:** Each body should be 3–5 short lines or a brief bulleted list. Headings are visually emphasized with a teal underline rule.

---

## `three_col(title, columns, dark=False)`

**When:** Three parallel categories — pillars, phases, or product dimensions.

**Layout:** `custom slide1-light` or `custom slide1-dark`.

**Parameters:**
- `columns`: List of 3 dicts: `[{"head": "...", "body": "..."}, ...]`

**Content rules:** Each column body should be very concise — 2–3 short lines at 16–18pt. Three-column slides read dense; use only when the three-way parallel is the insight.

---

## `end(text)`

**When:** Final slide of the deck.

**Layout:** `end slide` — dark blue background, colorful left-side geometric art.

**Parameters:**
- `text`: Closing message (e.g. `"THANK YOU"`, `"Q & A"`, `"期待与您深入探讨"`)

**Content rules:** Text is large (40pt+), centered on the right two-thirds of the slide, in white. The colorful art anchors the left side.

---

## Dark vs Light Slide Decision

Most content slides can be light or dark. Use dark (`dark=True`) when:
- The slide immediately follows a chapter separator (visual continuity)
- The slide contains a particularly powerful statement or conclusion
- You want to add dramatic emphasis to a critical data point

Keep the pattern consistent within a chapter — don't alternate dark/light randomly.

**Typical pattern:**
- Cover → dark
- Agenda → dark
- Chapter 1 separator → dark; first 1–2 slides in chapter → dark; rest → light
- Conclusion slide → dark
- End → dark
