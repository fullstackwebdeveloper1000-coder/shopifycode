# JUBELAY — Shopify Store Package

Pakistan modest fashion store: **Premium Instant Hijabs**  
Brand: **JUBELAY** · Tagline: *Style Made Simple* · Market: COD + JazzCash + EasyPaisa

Built from the Pakistan e-commerce research chat + `jubelay-brand-kit`.

## What's included

| Path | Purpose |
|------|---------|
| `theme/` | Shopify Online Store 2.0 theme (upload / `shopify theme push`) |
| `store-data/products.csv` | 8 launch SKUs ready to import |
| `store-data/pages/` | About, FAQ, Shipping, Returns, Contact, Privacy, Terms |
| `preview/` | Static homepage preview (no Shopify account needed) |
| `../jubelay-brand-kit/` | Logos, mockups, full setup guide |

## Quick start (Shopify Admin)

1. Create a store at [shopify.com/free-trial](https://www.shopify.com/free-trial) (Pakistan) — name it **JUBELAY**.
2. **Online Store → Themes → Add theme → Upload** → zip the `theme/` folder.
3. **Products → Import** → upload `store-data/products.csv`.
4. **Online Store → Pages** → paste content from `store-data/pages/`.
5. Theme settings: colors/fonts already match brand kit. Upload logo from `jubelay-brand-kit/images/jubelay-logo-mens-womens-primary.png`.
6. Enable **Manual payment → Cash on Delivery**. Add JazzCash / EasyPaisa apps or manual methods.
7. Shipping: Free above **Rs. 2,000**, else **Rs. 200**.

## CLI (after store exists)

```bash
cd jubelay-shopify/theme
npx @shopify/cli@latest theme push --store YOUR-STORE.myshopify.com
```

## Launch catalog (PKR)

| Product | Price |
|---------|-------|
| Instant Jersey Hijab Set — Black / Beige / Dusty Rose / Olive / Navy | 1,799 |
| Magnetic Hijab Pin Set (6 pcs) | 499 |
| Hijab Underscarf Cap — 2 Pack | 399 |
| Starter Bundle (Hijab + Pins + Cap) | 2,199 |

## Preview locally

```bash
cd jubelay-shopify/preview && python3 -m http.server 8080
```

Open http://localhost:8080
