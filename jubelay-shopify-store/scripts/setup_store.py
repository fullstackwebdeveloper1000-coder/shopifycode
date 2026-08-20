#!/usr/bin/env python3
"""Populate the JUBELAY Shopify preview store from research brand kit."""

from __future__ import annotations

import json
import os
import subprocess
import sys
from pathlib import Path

STORE = os.environ.get("SHOPIFY_STORE", "8rgs5j-x9.myshopify.com")
PUBLICATION_ID = "gid://shopify/Publication/322001994020"
LOCATION_ID = "gid://shopify/Location/117401682212"
ROOT = Path(__file__).resolve().parents[1]

AGENT_ENV = {
    **os.environ,
    "PATH": f"{os.path.expanduser('~/.local/bin')}:{os.environ.get('PATH', '')}",
    "SHOPIFY_CLI_AGENT_INFO": "n:cursor-cloud-agent|v:1.0.0|p:cursor",
    "SHOPIFY_CLI_AGENT_IDS": "s:bc-01037d27-a95e-4ff3-934c-cd1d88c147bb|r:jubelay-store-setup|i:pakistan-shopify",
}

HIJAB_COLORS = [
    ("Black", "BLK", "classic Black"),
    ("Beige", "BEG", "soft Beige"),
    ("Dusty Rose", "ROS", "Dusty Rose"),
    ("Olive Green", "OLV", "Olive Green"),
    ("Navy Blue", "NVY", "Navy Blue"),
]


def gql(query: str, variables: dict | None = None, mutate: bool = False) -> dict:
    cmd = [
        "shopify",
        "store",
        "execute",
        "--store",
        STORE,
        "--json",
        "--query",
        query,
    ]
    if variables is not None:
        cmd.extend(["--variables", json.dumps(variables)])
    if mutate:
        cmd.append("--allow-mutations")
    result = subprocess.run(cmd, capture_output=True, text=True, env=AGENT_ENV)
    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr, file=sys.stderr)
        raise RuntimeError(f"GraphQL failed: {result.returncode}")
    # CLI may print progress lines before JSON
    text = result.stdout.strip()
    start = text.find("{")
    if start == -1:
        raise RuntimeError(f"No JSON in output: {text}")
    return json.loads(text[start:])


def product_description(color_label: str) -> str:
    return f"""<p>Slip on elegance in seconds with the <strong>JUBELAY Instant Jersey Hijab Set</strong> in {color_label}.</p>
<p>No pins needed, no slipping, no fuss — just pure effortless style.</p>
<p><strong>What's included:</strong></p>
<ul>
<li>1x Instant Jersey Hijab with Built-in Cap ({color_label.split()[-1] if ' ' in color_label else color_label})</li>
<li>2x Magnetic Hijab Pins (Gold finish)</li>
<li>1x Branded JUBELAY Pouch</li>
</ul>
<p><strong>Why you'll love it:</strong></p>
<ul>
<li>Ready in seconds — just slip on and go</li>
<li>Premium stretch jersey — soft, breathable</li>
<li>Built-in cap — no underscarf needed</li>
<li>Magnetic pins — secure without damage</li>
<li>Free size — fits comfortably</li>
</ul>
<p><strong>Fabric &amp; care:</strong> 95% Cotton, 5% Spandex · Hand wash, hang dry</p>
<p><strong>Shipping:</strong> FREE delivery above Rs. 2,000 · 2–5 days across Pakistan · COD available</p>
<p>Style Made Simple. Shop JUBELAY.</p>"""


def create_product(
    title: str,
    description: str,
    product_type: str,
    tags: list[str],
    price: str,
    compare_at: str | None,
    sku: str,
    inventory_qty: int = 25,
) -> str:
    existing = gql(
        """
        query($q: String!) {
          products(first: 1, query: $q) {
            nodes { id title handle }
          }
        }
        """,
        {"q": f'title:"{title}"'},
    )
    nodes = existing.get("products", {}).get("nodes", [])
    if nodes:
        print(f"  exists: {title}")
        product_id = nodes[0]["id"]
    else:
        created = gql(
            """
            mutation($product: ProductCreateInput!) {
              productCreate(product: $product) {
                product {
                  id
                  handle
                  variants(first: 1) { nodes { id inventoryItem { id } } }
                }
                userErrors { field message }
              }
            }
            """,
            {
                "product": {
                    "title": title,
                    "descriptionHtml": description,
                    "vendor": "JUBELAY",
                    "productType": product_type,
                    "tags": tags,
                    "status": "ACTIVE",
                }
            },
            mutate=True,
        )
        payload = created["productCreate"]
        if payload["userErrors"]:
            raise RuntimeError(payload["userErrors"])
        product_id = payload["product"]["id"]
        variant_id = payload["product"]["variants"]["nodes"][0]["id"]
        inv_item_id = payload["product"]["variants"]["nodes"][0]["inventoryItem"]["id"]
        print(f"  created: {title} -> {product_id}")

        variant_input: dict = {
            "id": variant_id,
            "price": price,
            "inventoryItem": {"sku": sku},
        }
        if compare_at:
            variant_input["compareAtPrice"] = compare_at
        upd = gql(
            """
            mutation($productId: ID!, $variants: [ProductVariantsBulkInput!]!) {
              productVariantsBulkUpdate(productId: $productId, variants: $variants) {
                productVariants { id price sku }
                userErrors { field message }
              }
            }
            """,
            {"productId": product_id, "variants": [variant_input]},
            mutate=True,
        )
        if upd["productVariantsBulkUpdate"]["userErrors"]:
            raise RuntimeError(upd["productVariantsBulkUpdate"]["userErrors"])

        # Inventory tracked + quantity
        gql(
            """
            mutation($id: ID!) {
              inventoryItemUpdate(id: $id, input: { tracked: true }) {
                inventoryItem { id tracked }
                userErrors { field message }
              }
            }
            """,
            {"id": inv_item_id},
            mutate=True,
        )
        try:
            inv_set = gql(
                """
                mutation($input: InventorySetQuantitiesInput!, $key: String!) {
                  inventorySetQuantities(input: $input) @idempotent(key: $key) {
                    userErrors { field message }
                  }
                }
                """,
                {
                    "key": f"jubelay-inv-{sku}",
                    "input": {
                        "name": "available",
                        "reason": "correction",
                        "quantities": [
                            {
                                "inventoryItemId": inv_item_id,
                                "locationId": LOCATION_ID,
                                "quantity": inventory_qty,
                                "changeFromQuantity": 0,
                            }
                        ],
                    },
                },
                mutate=True,
            )
            if inv_set["inventorySetQuantities"]["userErrors"]:
                print(f"  inventory warn: {inv_set['inventorySetQuantities']['userErrors']}")
        except Exception as exc:
            print(f"  inventory skipped: {exc}")

    # Publish to Online Store
    pub = gql(
        """
        mutation($id: ID!, $input: [PublicationInput!]!) {
          publishablePublish(id: $id, input: $input) {
            userErrors { field message }
          }
        }
        """,
        {"id": product_id, "input": [{"publicationId": PUBLICATION_ID}]},
        mutate=True,
    )
    if pub["publishablePublish"]["userErrors"]:
        # Already published is fine
        msgs = " ".join(e["message"] for e in pub["publishablePublish"]["userErrors"])
        if "already" not in msgs.lower():
            print(f"  publish warn: {msgs}")

    return product_id


def create_collection(title: str, handle: str, description: str, product_ids: list[str]) -> str:
    existing = gql(
        """
        query($q: String!) {
          collections(first: 1, query: $q) {
            nodes { id title handle }
          }
        }
        """,
        {"q": f"handle:{handle}"},
    )
    nodes = existing.get("collections", {}).get("nodes", [])
    if nodes:
        collection_id = nodes[0]["id"]
        print(f"  collection exists: {title}")
    else:
        created = gql(
            """
            mutation($input: CollectionInput!) {
              collectionCreate(input: $input) {
                collection { id handle }
                userErrors { field message }
              }
            }
            """,
            {
                "input": {
                    "title": title,
                    "handle": handle,
                    "descriptionHtml": description,
                    "products": product_ids,
                }
            },
            mutate=True,
        )
        payload = created["collectionCreate"]
        if payload["userErrors"]:
            raise RuntimeError(payload["userErrors"])
        collection_id = payload["collection"]["id"]
        print(f"  collection created: {title}")

    gql(
        """
        mutation($id: ID!, $input: [PublicationInput!]!) {
          publishablePublish(id: $id, input: $input) {
            userErrors { field message }
          }
        }
        """,
        {"id": collection_id, "input": [{"publicationId": PUBLICATION_ID}]},
        mutate=True,
    )
    return collection_id


def create_page(title: str, handle: str, body: str) -> str:
    existing = gql(
        """
        query($q: String!) {
          pages(first: 1, query: $q) {
            nodes { id title handle }
          }
        }
        """,
        {"q": f"handle:{handle}"},
    )
    nodes = existing.get("pages", {}).get("nodes", [])
    if nodes:
        print(f"  page exists: {title}")
        return nodes[0]["id"]

    created = gql(
        """
        mutation($page: PageCreateInput!) {
          pageCreate(page: $page) {
            page { id handle }
            userErrors { field message code }
          }
        }
        """,
        {
            "page": {
                "title": title,
                "handle": handle,
                "body": body,
                "isPublished": True,
            }
        },
        mutate=True,
    )
    payload = created["pageCreate"]
    if payload["userErrors"]:
        raise RuntimeError(payload["userErrors"])
    print(f"  page created: {title}")
    return payload["page"]["id"]


def main() -> None:
    print(f"Setting up store: {STORE}")

    print("\n== Products ==")
    hijab_ids: list[str] = []
    for color, code, label in HIJAB_COLORS:
        pid = create_product(
            title=f"JUBELAY Instant Jersey Hijab Set — {color}",
            description=product_description(label),
            product_type="Instant Hijab",
            tags=[
                "hijab",
                "instant hijab",
                "jersey hijab",
                "jubelay",
                "modest fashion",
                color.lower(),
                "COD",
            ],
            price="1799.00",
            compare_at="2299.00",
            sku=f"JUBELAY-HIJAB-{code}-001",
        )
        hijab_ids.append(pid)

    pin_id = create_product(
        title="JUBELAY Magnetic Hijab Pin Set (6 pcs)",
        description="""<p>Secure your hijab without holes or damage. Set of 6 magnetic pins in gold/silver finish.</p>
<ul><li>Strong magnets</li><li>Fabric-safe</li><li>Reusable</li></ul>
<p>COD available · Free shipping above Rs. 2,000</p>""",
        product_type="Hijab Accessories",
        tags=["hijab pins", "magnetic pins", "accessories", "jubelay"],
        price="499.00",
        compare_at="699.00",
        sku="JUBELAY-PIN-GLD-001",
        inventory_qty=50,
    )

    cap_id = create_product(
        title="JUBELAY Hijab Underscarf Cap — 2 Pack",
        description="""<p>Soft underscarf caps for extra coverage and grip. Pack of 2 (Black + Beige).</p>
<p>Free size · Breathable fabric · Perfect with any hijab</p>""",
        product_type="Hijab Accessories",
        tags=["underscarf", "cap", "accessories", "jubelay"],
        price="399.00",
        compare_at="549.00",
        sku="JUBELAY-CAP-2PK-001",
        inventory_qty=40,
    )

    bundle_id = create_product(
        title="JUBELAY Starter Bundle (Hijab + Pins + Cap)",
        description="""<p>Everything you need to start — Instant Jersey Hijab Set + Magnetic Pins + Underscarf Cap.</p>
<p><strong>Bundle savings</strong> vs buying separately. Choose your hijab color at checkout via WhatsApp note.</p>
<p>Style Made Simple. Shop JUBELAY.</p>""",
        product_type="Bundles",
        tags=["bundle", "starter", "hijab", "jubelay", "best seller"],
        price="2199.00",
        compare_at="2697.00",
        sku="JUBELAY-BND-STR-001",
        inventory_qty=20,
    )

    print("\n== Collections ==")
    collections = {
        "instant-hijabs": create_collection(
            "Instant Hijabs",
            "instant-hijabs",
            "<p>Premium ready-to-wear instant jersey hijabs — Style Made Simple.</p>",
            hijab_ids,
        ),
        "hijab-accessories": create_collection(
            "Hijab Accessories",
            "hijab-accessories",
            "<p>Magnetic pins, underscarf caps, and essentials.</p>",
            [pin_id, cap_id],
        ),
        "bundles-sets": create_collection(
            "Bundles & Sets",
            "bundles-sets",
            "<p>Save more with JUBELAY starter bundles.</p>",
            [bundle_id],
        ),
        "new-arrivals": create_collection(
            "New Arrivals",
            "new-arrivals",
            "<p>Latest from JUBELAY.</p>",
            hijab_ids + [pin_id, cap_id, bundle_id],
        ),
        "best-sellers": create_collection(
            "Best Sellers",
            "best-sellers",
            "<p>Customer favourites.</p>",
            [hijab_ids[0], hijab_ids[2], bundle_id, pin_id],
        ),
    }

    print("\n== Pages ==")
    pages_dir = ROOT / "pages"
    page_files = [
        ("About Us", "about-us", "about-us.html"),
        ("Contact Us", "contact", "contact.html"),
        ("FAQ", "faq", "faq.html"),
        ("Shipping & Delivery", "shipping", "shipping.html"),
        ("Return & Exchange", "returns", "returns.html"),
        ("Privacy Policy", "privacy-policy", "privacy-policy.html"),
        ("Terms & Conditions", "terms", "terms.html"),
    ]
    page_ids: dict[str, str] = {}
    for title, handle, filename in page_files:
        body = (pages_dir / filename).read_text(encoding="utf-8")
        page_ids[handle] = create_page(title, handle, body)

    # Save IDs for theme/menu scripts
    out = {
        "store": STORE,
        "products": {
            "hijabs": hijab_ids,
            "pins": pin_id,
            "cap": cap_id,
            "bundle": bundle_id,
        },
        "collections": collections,
        "pages": page_ids,
        "publicationId": PUBLICATION_ID,
        "locationId": LOCATION_ID,
    }
    out_path = ROOT / "store-state.json"
    out_path.write_text(json.dumps(out, indent=2), encoding="utf-8")
    print(f"\nWrote {out_path}")
    print("Done.")


if __name__ == "__main__":
    main()
