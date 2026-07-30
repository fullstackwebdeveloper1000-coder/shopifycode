# Primal Win — Shopify theme fragments

These files extend **Shopify Dawn** (or similar). They are maintained against the static HTML in the repo root.

## Setup in Shopify

1. Upload **`assets/primal.css`** and link it in `layout/theme.liquid` (or Dawn’s custom CSS).
2. Copy **`snippets/pw-align-vars.liquid`** and **`snippets/pw-align-global.liquid`** into your theme snippets.
3. Add section files from **`sections/`** via **Edit code** (large files: use `INGREDIENTS-PASTE-GUIDE.txt` at repo root).
4. For multi-section pages, create or edit templates under **`templates/`** (e.g. `page.the-science.json`).

## Section map (main)

| Section file | Use |
|--------------|-----|
| `primal-hero-perfect.liquid` | Homepage hero (620px) |
| `primal-homeingredient.liquid` | Homepage ingredient cards + deep links |
| `primal-ingredient-grid.liquid` | Alternate homepage grid |
| `primal-trust-bar.liquid` / `primal-trust-fixed.liquid` | Trust strip (desktop / mobile swipe) |
| `primal-ingredients.liquid` | Full ingredients page |
| `primal-science-*.liquid`, `pw-science-*.liquid` | The Science page (see `templates/page.the-science.json`) |
| `primal-story-hero.liquid` | Our Story |
| `primal-gpw-hero.liquid` | Get Primal Win |
| `primal-returns-hero.liquid` | Returns policy |
| `primal-footer.liquid` | Footer block |
| `primal-exact-header.liquid` | Header alignment helper |

Store: `primal-win.myshopify.com`
