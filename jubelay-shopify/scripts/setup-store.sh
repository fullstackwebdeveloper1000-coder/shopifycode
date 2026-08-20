#!/usr/bin/env bash
# JUBELAY Shopify Store Setup Script
# Usage: ./jubelay-shopify/scripts/setup-store.sh your-store.myshopify.com

set -euo pipefail

STORE="${1:-}"
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
ROOT_DIR="$(dirname "$SCRIPT_DIR")"
WORKSPACE_DIR="$(dirname "$ROOT_DIR")"

if [ -z "$STORE" ]; then
  echo "Usage: $0 <store-handle.myshopify.com>"
  echo ""
  echo "Example: $0 jubelay-store.myshopify.com"
  echo ""
  echo "Pehle Shopify account banayein: https://www.shopify.com/pk"
  echo "Phir store URL yahan pass karein."
  exit 1
fi

# Normalize store handle
STORE="${STORE%.myshopify.com}.myshopify.com"

SHOPIFY="npx shopify"
cd "$WORKSPACE_DIR"

echo "=========================================="
echo "  JUBELAY Shopify Store Setup"
echo "  Store: $STORE"
echo "=========================================="
echo ""

# Step 1: Authenticate
echo "[1/5] Store authentication..."
echo "Browser khulega — Shopify CLI Connector App install karein."
$SHOPIFY store auth --store "$STORE" \
  --scopes "read_products,write_products,read_inventory,write_inventory,read_locations,read_files,write_files,read_themes,write_themes,read_content,write_content,read_online_store_pages"

# Step 2: Push theme
echo ""
echo "[2/5] JUBELAY theme push kar rahe hain..."
$SHOPIFY theme push --store "$STORE" --path jubelay-theme --unpublished

# Step 3: Import products via CSV (manual instruction)
echo ""
echo "[3/5] Products import..."
echo "Shopify Admin > Products > Import > Upload file:"
echo "  $ROOT_DIR/import/products.csv"
echo ""
echo "Products DRAFT status mein import honge — review ke baad Active karein."

# Step 4: Create pages
echo ""
echo "[4/5] Store pages create kar rahe hain..."
node "$SCRIPT_DIR/create-pages.mjs" "$STORE"

# Step 5: Final checklist
echo ""
echo "[5/5] Setup complete! Ab ye steps manually karein:"
echo ""
echo "  PAYMENTS (Settings > Payments):"
echo "    - Cash on Delivery enable karein"
echo "    - JazzCash / EasyPaisa add karein"
echo ""
echo "  SHIPPING (Settings > Shipping):"
echo "    - Pakistan zone: Free above Rs. 2,000 | Rs. 200 below"
echo ""
echo "  APPS:"
echo "    - WhatsApp Chat Button install karein"
echo "    - Judge.me reviews (optional)"
echo ""
echo "  THEME:"
echo "    - Online Store > Themes > JUBELAY theme publish karein"
echo "    - Logo upload: jubelay-brand-kit/images/jubelay-logo-primary.png"
echo "    - Announcement bar: 'COD Available Nationwide | Free Ship Rs.2000+'"
echo ""
echo "  DOMAIN:"
echo "    - jubelay.pk connect karein (Settings > Domains)"
echo ""
echo "  GO LIVE:"
echo "    - Test COD order place karein"
echo "    - Password protection hata kar launch karein"
echo ""
echo "=========================================="
echo "  JUBELAY — Style Made Simple"
echo "=========================================="
