# Westwell PPT QA Checklist

Run this before delivery and after any major visual iteration.

## P0 · Must Pass

- Titles are conclusion-first, not topic labels.
- Content title does not collide with the top-right brand zone.
- Body text is readable at projection distance; do not use tiny filler text.
- No slide exceeds the bottom safe line near `y=5.85"` unless it is a cover,
  chapter, or intentionally full-bleed image page.
- Missing real assets are shown as explicit placeholders, not decorative
  substitutes.
- Images preserve top-critical information; screenshots do not crop off the
  browser/app header unless intentionally focused.
- Teal is used as accent/rule/marker, not as the main body text color.

## P1 · Rhythm

- For 5+ page decks, a 2-slide visual grammar preview was generated first.
- No more than three similar light content pages in a row.
- Dense analytic pages are separated by hero, quote, KPI, image-led, or
  summary pages.
- Data pages have one strong numeric or chart anchor.
- Strategy decks include executive pauses, not just boxes of analysis.

## P2 · Craft

- Repeated elements align to the same left edge and spacing rhythm.
- Card heights and image frames are consistent within the same slide.
- Page chrome (`meta_*`, `foot_*`) is quiet and does not repeat the title.
- Tables have fewer columns/rows than a spreadsheet; dense backup data belongs
  in an appendix or separate document.
- Quotes and bottom lines are short enough to read in one glance.

## P3 · Verification

- Run the smoke test for builder API compatibility.
- Render the deck with `preview_pptx()` and inspect PNGs.
- Check the first, middle, and last slide manually in PowerPoint/Keynote when
  possible.
- For new layouts, include at least one dark and one light example in the
  smoke deck.

## 5-Dimension Review

Score 1-10 and fix the lowest dimension first:

| Dimension | Prompt |
|---|---|
| Brand consistency | Does this still feel like Westwell? |
| Visual hierarchy | Does the eye know where to go first? |
| Craft quality | Are alignment, spacing, crop, and type disciplined? |
| Functionality | Does every element support the title's claim? |
| Originality | Is the slide distinctive without becoming off-brand? |
