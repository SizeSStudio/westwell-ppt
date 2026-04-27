# Core Asset Protocol

Westwell PPT quality depends on real assets. Do not use generic decoration
when a logo, product image, UI screenshot, or site photo is the actual proof.

## When It Applies

Use this protocol whenever a deck mentions a specific:

- customer, partner, port, airport, factory, or market site
- Westwell product, platform, vehicle, or module
- competitor or named benchmark
- digital interface, dashboard, or app workflow

## Asset Priority

| Priority | Asset | Rule |
|---|---|---|
| 1 | Westwell + customer logos | Use official files where available; do not redraw. |
| 2 | Product / vehicle / site photos | Prefer real Westwell or client context over stock imagery. |
| 3 | UI screenshots | Required for digital platform claims. Mask sensitive data. |
| 4 | Architecture / flow diagrams | Use when the system relationship is the claim. |
| 5 | Brand colors / typography | Supporting signal, not a substitute for real assets. |

## If Assets Are Missing

- Use a clear placeholder frame, e.g. `[ CUSTOMER SITE PHOTO ]`.
- State the missing asset in the dummy page's data needs.
- Do not fill the slot with random stock art, decorative SVG, or abstract AI
  imagery unless the user explicitly wants a conceptual mood page.

## Minimum Quality Bar

For photos and screenshots:

- large enough to fill its frame without blur
- top-critical information preserved by the crop
- copyright/source is clear enough for the intended use
- visually consistent with neighboring assets
- directly supports the slide title

For charts and diagrams:

- labels readable after insertion into PPTX
- no default matplotlib colors unless intentionally restyled
- Westwell navy/teal/gray palette used for emphasis
- chart answers the slide's "so what" directly

## Dummy Page Fields

Every page that relies on assets should include:

- asset type: logo / product photo / UI / site photo / chart / diagram
- source: user-provided / internal / official web / generated placeholder
- fallback: placeholder / simplified diagram / split page / remove claim
