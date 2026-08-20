# JUBELAY — Shopify Store (Pakistan)

Research chat decisions applied: **Premium Instant Hijabs & Modest Fashion** for Pakistan, brand **JUBELAY** (*Style Made Simple*).

## Live preview store

| | |
|---|---|
| **Name** | JUBELAY |
| **Domain** | https://8rgs5j-x9.myshopify.com |
| **Currency / market** | PKR · Pakistan |
| **Theme** | Horizon (brand colors + logo applied) |

Open `store-access.json` for:

- **accessUrl** — enter the preview admin / password page
- **saveUrl** — claim this temporary store into your permanent Shopify account (do this soon — preview stores expire)

```bash
# Or from CLI:
shopify store open --store 8rgs5j-x9.myshopify.com
shopify store info --store 8rgs5j-x9.myshopify.com --json
```

## Already set up on the store

- **8 products** (5 Instant Jersey Hijab Sets @ Rs 1,799 + pins, caps, starter bundle)
- **5 collections:** Instant Hijabs, Accessories, Bundles & Sets, New Arrivals, Best Sellers
- **7 pages:** About, Contact, FAQ, Shipping, Returns, Privacy, Terms
- **Menus:** Main (Home / Shop / Collections / About / Contact) + Footer help links
- **Logo** uploaded and set on theme
- **Brand colors:** black `#2C2C2C`, gold `#C9A96E`, cream `#F5F0EB`
- **Fonts:** Playfair Display + Assistant
- **Announcement:** COD nationwide + free shipping above Rs 2,000
- **Hero:** “Style Made Simple” + SHOP NOW → Instant Hijabs
- **Discount:** `JUBELAY10` = 10% off once per customer

## Local preview (no Shopify login needed)

```bash
python3 serve.py
# Open http://127.0.0.1:4322/jubelay-shopify-store/preview/index.html
```

## Repo package

| Path | Purpose |
|------|---------|
| `preview/index.html` | Local storefront preview with brand colors + product grid |
| `products.csv` | Re-import catalog into any Shopify store |
| `pages/*.html` | Page body HTML (already pushed live) |
| `scripts/setup_store.py` | Re-run product/collection/page setup via CLI |
| `theme-config/` | Horizon settings + homepage/header snapshots |
| `store-state.json` | Product / collection / page GIDs |
| `../jubelay-brand-kit/` | Full brand guide + mockups from research |

## Next steps (you)

1. Open **saveUrl** in `store-access.json` and claim the store to your Shopify account.
2. Settings → Payments: enable **COD** + JazzCash / EasyPaisa.
3. Settings → Shipping: Pakistan zone, free above Rs 2,000 / Rs 200 below.
4. Add WhatsApp business number on Contact page + a WhatsApp chat app.
5. Add real product photos (hijab flat-lays / lifestyle).
6. Connect domain `jubelay.pk` when ready.
7. Password page: remove / publish when you go live.

## Brand reference (from research)

- Niche: Premium Instant Hijabs & Modest Fashion Accessories  
- Hero SKU: Instant Jersey Hijab Set — Rs **1,799**  
- Audience: Pakistani women 18–35  
- Stack: Shopify · COD · PostEx/Leopards/TCS · Instagram/TikTok/WhatsApp  
