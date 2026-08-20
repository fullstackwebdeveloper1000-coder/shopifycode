# JUBELAY — Pakistan Shopify Store

**Style Made Simple** | Premium Instant Hijabs & Modest Fashion

Yeh folder Pakistan e-commerce product research ke baad banaya gaya hai. Research chat mein **JUBELAY** brand finalize hua — modest fashion / instant hijab niche.

---

## Research Summary (Kya Decide Hua)

| Item | Detail |
|------|--------|
| **Brand** | JUBELAY |
| **Tagline** | Style Made Simple |
| **Product** | Instant Jersey Hijab Set (PKR 1,799) |
| **Market** | Pakistan — COD + JazzCash + EasyPaisa |
| **Platform** | Shopify + Dawn Theme |
| **Domain** | jubelay.pk ya shopjubelay.com |
| **Target** | Pakistani women 18-35, modest fashion |

### Launch Products (8 SKUs)
1. Instant Jersey Hijab — Black, Beige, Dusty Rose, Olive, Navy (PKR 1,799 each)
2. Magnetic Hijab Pin Set — 6 pcs (PKR 499)
3. Hijab Underscarf Cap — 2 Pack (PKR 399)
4. Starter Bundle — Hijab + Pins + Cap (PKR 2,199)

---

## Folder Structure

```
jubelay-shopify/
├── import/
│   ├── products.csv      ← Shopify Admin se import karein
│   ├── collections.csv   ← Collections manually banaein
│   └── pages.json        ← About, FAQ, Shipping, etc.
├── scripts/
│   ├── setup-store.sh    ← Automated setup (store connect ke baad)
│   └── create-pages.mjs  ← Pages create karta hai via CLI
└── README.md

jubelay-theme/            ← Customized Dawn theme (JUBELAY colors/fonts)
jubelay-brand-kit/        ← Logo, mockups, complete guide
```

---

## Quick Start — Store Banana (6 Steps)

### Step 1: Shopify Account Banayein
1. [shopify.com/pk](https://www.shopify.com/pk) par jayein
2. Free trial start karein (store name: **JUBELAY**)
3. Admin URL copy karein: `your-store.myshopify.com`

### Step 2: Store Connect Karein
```bash
cd /workspace
npm install
npx shopify store auth --store YOUR-STORE.myshopify.com
```

### Step 3: Theme Push Karein
```bash
npx shopify theme push --store YOUR-STORE.myshopify.com --path jubelay-theme
```

### Step 4: Products Import Karein
1. Shopify Admin → **Products** → **Import**
2. Upload: `jubelay-shopify/import/products.csv`
3. Review karein → Products **Active** karein

### Step 5: Pages + Settings
```bash
chmod +x jubelay-shopify/scripts/setup-store.sh
./jubelay-shopify/scripts/setup-store.sh YOUR-STORE.myshopify.com
```

Ya manually pages copy karein: `jubelay-brand-kit/JUBELAY-SHOPIFY-COMPLETE-GUIDE.md` (Section 7)

### Step 6: Pakistan Setup
| Setting | Value |
|---------|-------|
| **Payments** | COD + JazzCash + EasyPaisa |
| **Shipping** | Free above Rs. 2,000 / Rs. 200 below |
| **Courier** | PostEx / Leopards / TCS |
| **WhatsApp** | Chat button app install karein |
| **Announcement** | "COD Available Nationwide \| Free Ship Rs.2000+" |

---

## Brand Colors (Theme Settings)

| Color | Hex | Use |
|-------|-----|-----|
| JUBELAY Black | `#2C2C2C` | Buttons, headers |
| JUBELAY Gold | `#C9A96E` | Accents, sale tags |
| JUBELAY Cream | `#F5F0EB` | Section backgrounds |
| JUBELAY Text | `#333333` | Body text |

**Fonts:** Playfair Display (headings) + Assistant (body)

---

## Budget (Month 1)

| Tier | PKR |
|------|-----|
| Minimum test | 14,000 – 16,000 |
| Proper launch | 58,000 – 86,000 |
| Profit per order (Rs. 1,799 sale) | ~627 |

---

## Agla Step

**Aap apna Shopify store URL share karein** (e.g. `jubelay-store.myshopify.com`) — main products import, pages create, aur theme setup kar dunga.

Agar abhi account nahi hai, pehle [shopify.com/pk](https://www.shopify.com/pk) par trial start karein.

---

*Complete guide: `jubelay-brand-kit/JUBELAY-SHOPIFY-COMPLETE-GUIDE.md`*
