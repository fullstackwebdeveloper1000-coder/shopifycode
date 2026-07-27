# shopifycode / Primal Win

Static HTML marketing site for the "Primal Win" supplement brand, plus partial Shopify Liquid theme fragments under `theme/`.

## Cursor Cloud specific instructions

- This is a pure static front-end repo: hand-coded `.html` pages at the root, image/SVG assets, `theme/assets/primal.css`, and `theme/sections/*.liquid`. There is no backend, database, package manager, or build step, and nothing to install (Python stdlib only).
- Run the site locally with `python3 serve.py`, which serves the repo root at `http://127.0.0.1:4322` (see `serve.py`). This is the only runnable service. There is no lint, automated test, or build tooling.
- Interactive functionality is inline vanilla JS. On `ingredients.html`, clicking an ingredient card selects it and updates the detail panel; it deep-links via `ingredients.html?ing=<name>`.
- `theme/` contains only Shopify `sections/` + one CSS asset (no `layout/`, `templates/`, `config/`, `locales/`, `snippets/`), so it is not a complete/servable theme on its own and requires a Shopify CLI + host store to preview. This is optional and not reproducible from this repo alone.
