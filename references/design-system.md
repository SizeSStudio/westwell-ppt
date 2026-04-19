# Westwell PPT Design System

Reference for colors, fonts, geometry, and design principles used in `pptx_builder.py`.

---

## Colour Palette

| Name       | Hex       | Python constant | Use                                         |
|------------|-----------|-----------------|---------------------------------------------|
| Dark Blue  | `#2743C6` | `C_DARK`        | Dark slide background (cover, chapter, dark slides) |
| Navy       | `#001BA6` | `C_NAVY`        | **All text on light slides** — titles, headings, KPI values, table cell emphasis |
| Teal       | `#00CAD4` | `C_TEAL`        | **Non-text accents only** — rules, bullet markers, card top strips, dividers |
| White      | `#FFFFFF` | `C_WHITE`       | **All text on dark slides**                 |
| Light Gray | `#F0F4FF` | `C_LGRAY`       | Card backgrounds, light panel fills         |
| Body Gray  | `#555566` | `C_GRAY`        | Supporting caption text on light slides     |
| Near-Black | `#1A1A2E` | `C_BLACK`       | Body prose / table cells on light slides    |
| Alt Row    | `#E4ECFB` | `C_ALTROW`      | Alternating table row background            |

**Text-color rule (strict):** Every piece of rendered text must use one of
**Navy / Black / White**. Teal is reserved for non-text accents (rules,
markers, top strips). Never use teal for text — including stats values,
chapter numbers, column headings, or bold table cells.

---

## Typography

| Role              | Font              | Size          | Weight | Color (light/dark)     |
|-------------------|-------------------|---------------|--------|------------------------|
| Slide title        | Encode Sans       | 28–34 pt      | Bold   | `C_NAVY` / `C_WHITE`   |
| Chapter title      | Encode Sans       | 40 pt         | Bold   | `C_WHITE`              |
| Cover title        | Encode Sans       | 40 pt         | Bold   | `C_WHITE`              |
| Body text (CN)     | 思源黑体           | 18–22 pt      | Regular| `C_BLACK` / `C_WHITE`  |
| Body text (EN)     | Encode Sans       | 18–20 pt      | Regular| `C_BLACK` / `C_WHITE`  |
| Bullet text        | 思源黑体           | 18–20 pt      | Regular| `C_BLACK` / `C_WHITE`  |
| Table header       | Encode Sans / 思源黑体 | 16 pt    | Bold   | `C_WHITE`              |
| Table cell         | 思源黑体           | 14–16 pt      | Regular| `C_BLACK`              |
| KPI value          | Encode Sans       | 48–56 pt      | Bold   | `C_TEAL`               |
| KPI label          | 思源黑体           | 14 pt         | Regular| `C_GRAY`               |
| Caption            | 思源黑体           | 12 pt         | Regular| `C_GRAY`               |

**Hard rule: no body text below 18 pt.** Small text looks unprofessional in Westwell's clean style.

---

## Slide Geometry (inches)

```
Slide canvas:  13.333 × 7.500"

Title placeholder (all content slides):
  Left=0.906  Top=0.394  Width=11.500  Height=1.449

Content safe zone (below title, above WMF decoration at y=5.933"):
  Left=0.906   Top=1.800          ← CT (no subtitle)  / 2.00 CT_WITH_SUB
  Right=12.427  Bottom=5.850      ← CB (no footnote)
  Width=11.521  Height=4.050

With editorial framing (subtitle / footnote / bottom_callout):
  CT_WITH_SUB  = 2.00
  FOOTNOTE_Y   = 5.55   FOOTNOTE_H = 0.30   (ends 5.85)
  CB_WITH_FN   = 5.48   (content ends 0.07" above footnote top)
  bottom_callout cb = 5.05   callout at y=5.12, h=0.73  (ends 5.85)

Cover title placeholder (left half only):
  Left=0.689  Top=2.069  Width=8.754  Height=2.705

Cover subtitle / date:
  Left=0.689  Top=5.070  Width=8.754  Height=1.220
```

### Bottom-padding discipline

The template's WMF footer art starts at **y ≈ 5.933"** — this is the hard
ceiling. Content, footnote strips, and bottom-callout strips should **hug
that line (0.05–0.10" margin)**, not leave a half-inch of empty band above
it. Users read that unused band as "the page is too sparse" even on
density=compact layouts.

Earlier revisions left 0.20–0.40" of empty space below every footnote or
callout; current constants (CB=5.85, CB_WITH_FN=5.48, FOOTNOTE_Y=5.55,
bottom_callout at 5.12/0.73) collapse that gap to ~0.08". If you ever
raise these, raise them together — do not re-introduce the dead band.

**Rule:** never place content below y = 5.85"; never place a callout strip
whose bottom exceeds y = 5.90".

---

## Layout Index

The template has 21 named layouts. Key ones used by `pptx_builder.py`:

| Layout name              | Used for                        | Background     |
|--------------------------|---------------------------------|----------------|
| `标题幻灯片`              | `cover()`                       | Dark blue + right-side art |
| `agenda slide 1`         | `agenda()`                      | Dark blue, teal row accents |
| `agenda slide 2`         | `chapter()`                     | Dark blue + colorful right-side geometric art |
| `custom slide1-light`    | `text()`, `bullets()`, `table()`, `stats()`, `two_col()`, `three_col()` (light) | White |
| `custom slide1-dark`     | same slide types with `dark=True` | Dark blue |
| `content1-text&image`    | `image_left()`                  | White, image right half |
| `content2-text&image`    | `image_below()`                 | White, image lower half |
| `end slide`              | `end()`                         | Dark blue + left-side colorful art |

---

## Decoration System (WMF Images)

The template uses WMF (Windows Metafile) vector decorations **inherited from slide layouts**. Do not attempt to add or modify these programmatically — they are part of the layout master and appear automatically when the correct layout is chosen.

| Pattern                  | Appears on                                |
|--------------------------|-------------------------------------------|
| Large central composition | Cover slide (right side)                 |
| Colorful geometric art   | Chapter slides (`agenda slide 2`), End slide |
| Bottom circles strip     | Standard content layouts                  |
| Right-side circles strip | Some light content layouts                |
| TOC circles              | Agenda slide                              |

---

## Design Philosophy (derived from actual Westwell decks)

Westwell solution decks are **presentation-style, not document-style**. The visual language is:

**Do:**
- Title at top-left, left-aligned, navy (~22pt bold) — never centered on content slides
- "Make a WELL Change" / logo inherited top-right from template
- Large product images or scene photos as primary visual anchors
- 2–3 information layers maximum per slide
- Generous whitespace — if a slide feels "full", remove something
- Teal horizontal rule below title to separate title from content area
- Dark blue slides (`C_DARK`) for: cover, chapter separators, key statement slides
- On dark slides: white body text, teal accent numbers

**Don't:**
- Dense multi-box card layouts (McKinsey-style information packing)
- Font sizes below 18pt for any readable body text
- More than ~40 words of body text on a single slide
- Decorative colored borders around content boxes
- More than 3 columns
- More than 5–6 bullet points per slide

**Observed layout patterns in Westwell decks (in order of frequency):**
1. Title + horizontal rule + left-text / right-image (most common)
2. Full-bleed image + dark overlay + centered large text (chapter, statement)
3. Title + horizontal rule + 2–3 column sections with icon + heading + body
4. Title + architecture/flow diagram (image fills content area)
5. Dark slide + large KPI numbers + supporting labels
6. TOC: full-bleed background photo + 2×3 rounded card grid
