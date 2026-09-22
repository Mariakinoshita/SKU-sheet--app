"""
app.py - Streamlit web interface for the SKU sheet generator.
"""

import io
import os
from datetime import datetime

import requests
import streamlit as st
import pycountry
from openpyxl import Workbook
from openpyxl.styles import Font

API_VERSION = "2025-01"
RELEASE_DATE_NAMESPACE = "custom"
RELEASE_DATE_KEY = "release_date"


def get_credentials(store_choice, env_file):
    """Reads Shopify credentials for the selected store from Streamlit
    secrets (when deployed, using a [StoreName] section per store) or
    from the given local .env file (when run locally)."""
    try:
        store_secrets = st.secrets[store_choice]
        return (
            store_secrets["SHOPIFY_STORE_DOMAIN"],
            store_secrets["SHOPIFY_CLIENT_ID"],
            store_secrets["SHOPIFY_CLIENT_SECRET"],
        )
    except Exception:
        from dotenv import load_dotenv
        load_dotenv(env_file, override=True)
        return (
            os.getenv("SHOPIFY_STORE_DOMAIN"),
            os.getenv("SHOPIFY_CLIENT_ID"),
            os.getenv("SHOPIFY_CLIENT_SECRET"),
        )


def get_access_token(domain, client_id, client_secret):
    url = f"https://{domain}/admin/oauth/access_token"
    resp = requests.post(url, data={
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": "client_credentials",
    })
    resp.raise_for_status()
    return resp.json()["access_token"]


def country_name(code):
    if not code:
        return ""
    country = pycountry.countries.get(alpha_2=code)
    return country.name if country else code


SEARCH_QUERY = """
query searchProduct($search: String!) {
  products(first: 5, query: $search) {
    edges {
      node {
        title
        featuredImage {
          url
        }
        metafield(namespace: "%s", key: "%s") {
          value
        }
        variants(first: 100) {
          edges {
            node {
              sku
              title
              price
              image {
                url
              }
              inventoryItem {
                unitCost {
                  amount
                }
                harmonizedSystemCode
                countryCodeOfOrigin
              }
            }
          }
        }
      }
    }
  }
}
""" % (RELEASE_DATE_NAMESPACE, RELEASE_DATE_KEY)


def find_product(domain, token, title_query):
    url = f"https://{domain}/admin/api/{API_VERSION}/graphql.json"
    headers = {
        "X-Shopify-Access-Token": token,
        "Content-Type": "application/json",
    }
    resp = requests.post(url, headers=headers, json={
        "query": SEARCH_QUERY,
        "variables": {"search": f"title:*{title_query}*"},
    })
    resp.raise_for_status()
    data = resp.json()

    if "errors" in data:
        return None, data["errors"]

    return data["data"]["products"]["edges"], None


def build_rows(product):
    rows = []
    title = product["title"]
    release_date = product["metafield"]["value"] if product.get("metafield") else ""
    product_image = product["featuredImage"]["url"] if product.get("featuredImage") else ""

    for edge in product["variants"]["edges"]:
        v = edge["node"]
        variant_label = v["title"]
        sku_name = title if variant_label == "Default Title" else f"{title} {variant_label}"

        cost = ""
        item = v.get("inventoryItem")
        if item and item.get("unitCost"):
            cost = item["unitCost"]["amount"]

        hs_code = item["harmonizedSystemCode"] if item else ""
        origin = country_name(item["countryCodeOfOrigin"] if item else "")
        image = v["image"]["url"] if v.get("image") else product_image

        rows.append([
            v["sku"], sku_name, cost, v["price"],
            release_date, hs_code, origin, image,
        ])
    return rows


st.set_page_config(page_title="SKU Sheet Generator", page_icon=":package:")
st.title("SKU Sheet Generator")
st.write("Pick a store, paste in product titles (one per line), then generate a downloadable Excel SKU sheet.")

STORES = {
    "McFly": "store.env",
    "Wunderhorse UK": "wunderhorse.env",
}
store_choice = st.selectbox("Store", list(STORES.keys()))
env_file = STORES[store_choice]

titles_input = st.text_area(
    "Product titles",
    height=200,
    placeholder="Team McFly Black Hoodie\nPower To Play | CD\n...",
)

if st.button("Generate SKU Sheet", type="primary"):
    titles = [t.strip() for t in titles_input.splitlines() if t.strip()]

    if not titles:
        st.warning("Add at least one product title first.")
    else:
        domain, client_id, client_secret = get_credentials(store_choice, env_file)
        if not all([domain, client_id, client_secret]):
            st.error("Missing Shopify credentials - check your secrets/env setup.")
        else:
            with st.spinner("Fetching product data from Shopify..."):
                token = get_access_token(domain, client_id, client_secret)

                all_rows = []
                not_found = []
                ambiguous = []

                for title in titles:
                    matches, errors = find_product(domain, token, title)
                    if errors:
                        st.error(f"Error looking up '{title}': {errors}")
                        continue
                    if not matches:
                        not_found.append(title)
                        continue
                    if len(matches) > 1:
                        ambiguous.append((title, [m["node"]["title"] for m in matches]))
                    product = matches[0]["node"]
                    all_rows.extend(build_rows(product))

            if not_found:
                st.warning(f"No match found for: {(', ').join(not_found)}")
            for title, options in ambiguous:
                st.info(f"'{title}' matched multiple products - used the first ({options[0]}). Other matches: {(', ').join(options[1:])}")

            if all_rows:
                wb = Workbook()
                ws = wb.active
                ws.title = "SKU Sheet"
                headers = ["SKU", "SKU Name", "Purchase Price", "Retail Price",
                           "Release Date", "HS Tariff Code", "Country of Manufacture", "IMAGE"]
                ws.append(headers)
                for cell in ws[1]:
                    cell.font = Font(bold=True)
                for row in all_rows:
                    ws.append(row)
                for col in ws.columns:
                    values = [str(c.value) for c in col if c.value is not None]
                    max_len = max((len(v) for v in values), default=10)
                    ws.column_dimensions[col[0].column_letter].width = min(max_len + 2, 60)

                buffer = io.BytesIO()
                wb.save(buffer)
                buffer.seek(0)

                st.success(f"Generated {len(all_rows)} rows across {len(titles) - len(not_found)} product(s).")
                st.download_button(
                    "Download SKU Sheet (.xlsx)",
                    data=buffer,
                    file_name=f"sku_sheet_{domain.split('.')[0]}_{datetime.now():%Y%m%d}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("No rows generated - check the product titles above.")
