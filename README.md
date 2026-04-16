# westwell-ppt

A [Claude Code](https://claude.ai/code) skill that generates professional, branded PowerPoint decks from `.md` documents, outlines, or open-ended business questions.

Combines **McKinsey Problem Solving methodology** (Issue Tree, Hypotheses, Dummy Pages with page-dependency tracking) with a **corporate visual template** and **consulting-grade storyline** (SCQA / BCG / Bain).

## What it does

- **Light mode**: hand it a structured `.md` or outline, get a polished `.pptx` back
- **Full mode**: give it an open-ended question like "analyze the AI scheduling market opportunity", and it runs Issue Tree + Hypotheses before designing slides
- **Cross-session resumption**: upload a `DummyPages.md` from a prior session and continue from any page

## Prerequisites

| Tool | Purpose | Install |
|------|---------|---------|
| `python-pptx` | Slide generation | `pip install python-pptx` |
| `LibreOffice` | `.potx` to `.pptx` conversion | `brew install --cask libreoffice` |
| `pymupdf` (fitz) | PPTX to PDF to PNG preview | `pip install pymupdf` |

Verify:

```bash
soffice --version && python3 -c "import pptx, fitz; print('OK')"
```

## Installation

### As a Claude Code skill

Copy the entire folder to your skills directory:

```bash
cp -r westwell-ppt ~/.claude/skills/westwell-ppt
```

Then invoke in Claude Code with `/westwell-ppt` or let it trigger automatically when you mention creating a PPT, presentation, or slides.

### Standalone usage

```python
import sys, os
SKILL_DIR = os.path.expanduser('~/.claude/skills/westwell-ppt')
sys.path.insert(0, SKILL_DIR)
from scripts.pptx_builder import WestwellPPT, preview_pptx

ppt = WestwellPPT(
    template=os.path.join(SKILL_DIR, 'assets', 'PPTTemplate.potx'),
    output='my_deck.pptx'
)
ppt.cover('Title', 'Subtitle', 'Context', '2026')
ppt.bullets('Conclusion-first title', ['Point A', 'Point B', 'Point C'])
ppt.end('Thank you')
ppt.save()
```

## Slide types

| Method | Use case |
|--------|----------|
| `cover()` | Title slide with brand art |
| `agenda()` | Table of contents with numbered cards |
| `chapter()` | Dark separator between major sections |
| `statement()` | Single bold claim, centered (impact moment) |
| `bullets()` | 3-7 bulleted points |
| `text_image()` | Left text + right image (side by side) |
| `image()` | Full-width image with optional narrative body |
| `stats()` | 2-4 KPI cards with auto-sizing values |
| `two_col()` | Two-column comparison |
| `three_col()` | Three-column insight cards |
| `table()` | Data table with navy header row |
| `text()` | Flowing prose |
| `end()` | Closing slide with brand art |

Every content method supports `dark=True` for dark-background variant and `density='compact'` (default) or `density='standard'` for different bottom decoration weights.

## Design system

- **Text colors**: Navy (`#001BA6`) / Near-Black (`#1A1A2E`) / White only. No colored text.
- **Accent color**: Teal (`#00CAD4`) for non-text elements only (rules, bullet markers, card strips).
- **Title rule**: tight to title text (RULE_Y = 1.15"), capped at 9" width to avoid brand-zone collision.
- **Titles**: must be insight conclusions, not topic labels. Soft max 25 characters per line.
- **Density**: `compact` (default) uses `custom slide2-*` with slim bottom decoration. `standard` uses `custom slide1-*` with full decorative footer.

See `references/design-system.md` for full color palette, typography, and geometry.

## Workflow (when used as a Claude Code skill)

```
Phase 1 (Full mode only)
  Step 1  Define boundary (Is / Isn't)
  Step 2  Issue Tree + Hypotheses        -> methodology.md

Phase 2
  Step 3  Design Dummy Pages             -> dummy-pages-spec.md
  Step 4  User confirms Dummy

Phase 3
  Step 5  Per-page generation loop       -> data-collection.md
  Step 6  Deliver .pptx
  Step 7  Iterate                        -> troubleshooting.md
```

## Template

The included `assets/PPTTemplate.potx` provides 21 named slide layouts. Replace it with your own `.potx` or `.pptx` template to adapt the visual identity. The builder references layouts by name (e.g., `custom slide1-light`, `agenda slide 2`), so your template must include matching layout names or you'll need to update the `_layout()` fallback logic.

## File structure

```
westwell-ppt/
├── SKILL.md                    <- Skill manifest (loaded by Claude Code)
├── README.md                   <- This file
├── assets/
│   └── PPTTemplate.potx        <- PowerPoint template
├── scripts/
│   └── pptx_builder.py         <- WestwellPPT builder class (~960 lines)
└── references/
    ├── design-system.md        <- Colors, fonts, geometry
    ├── layouts-guide.md        <- Narrative slide types
    ├── layouts-analytic.md     <- McKinsey-style analytic layouts
    ├── methodology.md          <- Problem Solving (Issue Tree, Hypotheses)
    ├── dummy-pages-spec.md     <- Dummy page format + dependency system
    ├── data-collection.md      <- Per-page generation loop
    └── troubleshooting.md      <- Common issues & fixes
```

## License

Code is MIT. The template artwork (`assets/PPTTemplate.potx`) contains proprietary Westwell brand assets and is NOT covered by the MIT license. See [LICENSE](LICENSE) for details.
