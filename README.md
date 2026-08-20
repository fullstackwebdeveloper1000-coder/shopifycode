# shopifycode

Two e-commerce projects in one repo:

## 1. JUBELAY — Pakistan Shopify Store (active)

**Brand:** JUBELAY · *Style Made Simple*  
**Product:** Premium Instant Jersey Hijab Sets @ **PKR 1,799**  
**Market:** Pakistan (COD-first)

| Resource | Path |
|----------|------|
| Store setup + live preview URLs | [`jubelay-shopify-store/README.md`](jubelay-shopify-store/README.md) |
| Brand kit + Shopify guide | [`jubelay-brand-kit/`](jubelay-brand-kit/) |
| Local storefront preview | [`jubelay-shopify-store/preview/index.html`](jubelay-shopify-store/preview/index.html) |
| Product catalog CSV | [`jubelay-shopify-store/products.csv`](jubelay-shopify-store/products.csv) |

**Preview locally:**
```bash
python3 serve.py
# Open http://127.0.0.1:4322/jubelay-shopify-store/preview/index.html
```

**Live Shopify preview store:** https://8rgs5j-x9.myshopify.com  
Claim instructions in `jubelay-shopify-store/store-access.json`.

---

## 2. Primal Win — UK Supplement Brand (separate project)

Static HTML pages for a UK supplement brand (not Pakistan e-commerce).

| Page | File |
|------|------|
| Homepage | `index.html` |
| Product | `get-primal-win.html` |
| Science / Ingredients / Our Story / Returns | `*.html` |

---

## Research decision (Pakistan)

From product research chat: user rejected gadgets → chose **fashion/beauty/dressing** → final brand **JUBELAY** with Dawn/Horizon Shopify theme, COD, free shipping above Rs. 2,000.
