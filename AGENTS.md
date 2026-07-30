# shopifycode / Primal Win

Static HTML marketing site for the Primal Win supplement brand, plus Shopify Liquid theme fragments under `theme/`.

## Cursor Cloud specific instructions

- **Static site:** Hand-coded `.html` pages at the repo root, image/SVG assets, and optional `theme/` fragments. No package manager or build step.
- **Run locally:** `python3 serve.py` → `http://127.0.0.1:4322`
- **Reference HTML:** Use root `*.html` files as the design source of truth when matching Shopify sections.
- **Ingredients page:** `ingredients.html` supports deep links via `?ing=<slug>`; same behavior is implemented in `theme/sections/primal-ingredients.liquid`.
- **Shopify theme:** Not a full Dawn replacement. Copy `theme/assets/primal.css`, `theme/snippets/*`, and `theme/sections/*` into your live Dawn theme. Page JSON templates live under `theme/templates/`. See `theme/README.md` and `INGREDIENTS-PASTE-GUIDE.txt` for paste workflow.
- **Alignment standard:** 1440px frame, 1280px content line — see `theme/snippets/pw-align-vars.liquid`.
