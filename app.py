import html
import io
import random
import re
from typing import Dict, List, Optional, Tuple
from urllib.parse import quote

import pandas as pd
import streamlit as st
import xmltodict
from bs4 import BeautifulSoup
from openpyxl.styles import Font, PatternFill

# ==============================================================================
# CONTACT & APP CONFIGURATION
# ==============================================================================
APP_VERSION = "v3.2"
CONTACT_EMAIL = "salarkhan.data@gmail.com"
LINKEDIN_URL = "https://www.linkedin.com/in/salarkhan-data"
FREE_EXPORT_LIMIT = 500
# ==============================================================================

st.set_page_config(
    page_title=f"FeedMedic {APP_VERSION} — Shopify Supplier Feed Auditor & Cleaner",
    layout="wide",
)


def build_gmail_url(email: str, subject: str, body: str) -> str:
    return (
        f"https://mail.google.com/mail/?view=cm&fs=1"
        f"&to={quote(email)}&su={quote(subject)}&body={quote(body)}"
    )


# --- 1. REALISTIC 650-ROW MULTI-OPTION DEMO SUPPLIER FEED GENERATOR ---
def generate_demo_feed() -> pd.DataFrame:
    rng = random.Random(42)
    # (base_name, p_type, vendor, opt1_name, opt2_name, [(opt1_val, opt2_val), ...])
    categories = [
        ("Wireless Earbuds", "Audio", "SoundCore Wholesale", "Color", "", [("Black", ""), ("White", ""), ("Navy", "")]),
        (
            "Cotton Crew T-Shirt",
            "Apparel",
            "NorthTex Apparel",
            "Size",
            "Color",
            [("Small", "Black"), ("Small", "White"), ("Medium", "Black"), ("Medium", "White"), ("Large", "Black")],
        ),
        ("Insulated Water Bottle", "Home & Kitchen", "HydroPeak", "Capacity", "", [("500ml", ""), ("750ml", ""), ("1L", "")]),
        ("USB-C Fast Charger", "Electronics", "VoltGrid", "Wattage", "", [("30W", ""), ("65W", ""), ("100W", "")]),
        (
            "Leather Card Wallet",
            "Accessories",
            "CraftHide",
            "Material",
            "Color",
            [("Full-Grain", "Brown"), ("Full-Grain", "Black"), ("Suede", "Tan")],
        ),
        ("Ceramic Coffee Mug", "Home & Kitchen", "KilnWorks", "Finish", "", [("Matte Black", ""), ("Cream", ""), ("Olive", "")]),
        ("LED Desk Lamp", "Lighting", "LumaTech", "Light Tone", "", [("Warm White", ""), ("Cool White", "")]),
        ("Silicone Phone Case", "Accessories", "ShieldPro", "Color", "", [("Midnight", ""), ("Clear", ""), ("Forest", "")]),
    ]

    rows = []
    sku_counter = 1000

    # Inject 1 product with 105 variants to test legacy 100-variant vs expanded 2048-variant behavior
    for v_idx in range(1, 106):
        sku_counter += 1
        rows.append(
            {
                "Handle": "bulk-braided-usb-cable-pack",
                "Title": "Bulk Braided USB-C Cable",
                "Body (HTML)": "<p>Heavy-duty nylon braided charging cable.</p>",
                "Vendor": "VoltGrid",
                "Type": "Electronics",
                "Tags": "cable, usb-c, wholesale",
                "Option1 Name": "Length & Color",
                "Option1 Value": f"Variant-{v_idx}",
                "Option2 Name": "",
                "Option2 Value": "",
                "Option3 Name": "",
                "Option3 Value": "",
                "Variant SKU": f"VG-{sku_counter}",
                "Variant Price": "14.99",
                "Variant Compare At Price": "19.99",
                "Variant Inventory Qty": "40",
                "Image Src": "",
                "Status": "active",
            }
        )

    # Generate remaining 545 realistic rows with controlled supplier defects
    prod_num = 1
    while len(rows) < 650:
        base_name, p_type, vendor, o1_name, o2_name, option_pairs = categories[prod_num % len(categories)]
        title = f"{base_name} Series {prod_num}"
        handle = re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")

        if prod_num % 8 == 0:
            handle = f"{base_name} & Series {prod_num}"
        elif prod_num % 29 == 0:
            handle = ""

        base_price = round(rng.uniform(12.0, 189.0), 2)
        body = f"<p>Premium {base_name.lower()} designed for daily retail durability.</p>"

        if prod_num % 7 == 0:
            body = (
                f"<script>trackSupplier()</script><style>.p{{color:red}}</style>"
                f"<p style='font-family:Calibri'>Premium {base_name.lower()}.&amp;nbsp;Wholesale grade.</p>"
            )

        for o_idx, (o1_val, o2_val) in enumerate(option_pairs):
            if len(rows) >= 650:
                break
            sku_counter += 1
            sku = f"SUP-{sku_counter}"
            price_str = f"{base_price:.2f}"
            compare_str = f"{round(base_price * 1.25, 2):.2f}"
            qty_str = str(rng.randint(5, 120))
            status_str = "active"
            img_str = ""
            opt1_v, opt2_v = o1_val, o2_val

            roll = rng.random()
            if roll < 0.14:
                price_str = f"${base_price:,.2f}"
            elif roll < 0.21:
                price_str = f"EUR {str(base_price).replace('.', ',')}"
            elif roll < 0.24:
                price_str = "0.00"
            elif roll < 0.26:
                price_str = "-9.50"

            if rng.random() < 0.06:
                compare_str = f"{max(1.0, round(base_price - 5.0, 2)):.2f}"

            if rng.random() < 0.05 and sku_counter > 1110:
                sku = f"SUP-{sku_counter - 3}"

            q_roll = rng.random()
            if q_roll < 0.04:
                qty_str = "In Stock"
            elif q_roll < 0.07:
                qty_str = "Out of Stock"
            elif q_roll < 0.09:
                qty_str = "-4"

            s_roll = rng.random()
            if s_roll < 0.05:
                status_str = "Enabled"
            elif s_roll < 0.08:
                status_str = "Published"
            elif s_roll < 0.10:
                status_str = "Disabled"

            if rng.random() < 0.07:
                img_str = f"/uploads/catalog/img_{sku_counter}.jpg"
            elif rng.random() < 0.03:
                img_str = f"http://cdn.supplier.local/my image {sku_counter}.jpg"

            # Inject 3 genuine duplicate option-combination collisions
            if prod_num in (11, 22, 33) and o_idx == 1:
                opt1_v, opt2_v = option_pairs[0]

            rows.append(
                {
                    "Handle": handle,
                    "Title": title,
                    "Body (HTML)": body,
                    "Vendor": vendor,
                    "Type": p_type,
                    "Tags": f"{p_type.lower()}, {vendor.lower()}",
                    "Option1 Name": o1_name,
                    "Option1 Value": opt1_v,
                    "Option2 Name": o2_name,
                    "Option2 Value": opt2_v,
                    "Option3 Name": "",
                    "Option3 Value": "",
                    "Variant SKU": sku,
                    "Variant Price": price_str,
                    "Variant Compare At Price": compare_str,
                    "Variant Inventory Qty": qty_str,
                    "Image Src": img_str,
                    "Status": status_str,
                }
            )
        prod_num += 1

    return pd.DataFrame(rows)


# --- 2. UNIVERSAL IN-MEMORY FILE PARSER ---
def find_records_in_nested(obj) -> List[dict]:
    best: List[dict] = []
    if isinstance(obj, list):
        dicts = [x for x in obj if isinstance(x, dict)]
        if len(dicts) > len(best):
            best = dicts
        for item in obj:
            cand = find_records_in_nested(item)
            if len(cand) > len(best):
                best = cand
    elif isinstance(obj, dict):
        for val in obj.values():
            cand = find_records_in_nested(val)
            if len(cand) > len(best):
                best = cand
    return best


def find_single_record_dict(obj) -> Optional[dict]:
    if isinstance(obj, dict):
        scalar_vals = sum(1 for v in obj.values() if isinstance(v, (str, int, float)))
        if scalar_vals >= 2:
            return obj
        for val in obj.values():
            found = find_single_record_dict(val)
            if found:
                return found
    return None


def read_uploaded_file(uploaded_file) -> pd.DataFrame:
    raw_bytes = uploaded_file.getvalue()
    if not raw_bytes:
        raise ValueError("Uploaded file is empty (0 bytes).")
    name = uploaded_file.name.lower()

    if name.endswith(".csv"):
        try:
            text = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="replace")
        lines = [ln for ln in text.splitlines() if ln.strip()]
        if not lines:
            raise ValueError("CSV file contains no readable rows.")
        header = lines[0]
        sep = ";" if header.count(";") > header.count(",") else ","
        if header.count("\t") > max(header.count(","), header.count(";")):
            sep = "\t"
        return pd.read_csv(io.StringIO(text), sep=sep, dtype=str, keep_default_na=False)

    if name.endswith(".xlsx") or name.endswith(".xls"):
        return pd.read_excel(io.BytesIO(raw_bytes), dtype=str, keep_default_na=False)

    if name.endswith(".xml"):
        try:
            text = raw_bytes.decode("utf-8-sig")
        except UnicodeDecodeError:
            text = raw_bytes.decode("latin-1", errors="replace")
        parsed = xmltodict.parse(text)
        records = find_records_in_nested(parsed)
        if not records:
            single = find_single_record_dict(parsed)
            if single:
                records = [single]
        if not records:
            raise ValueError("No product records found inside this XML file.")
        return pd.json_normalize(records, sep="_").fillna("").astype(str)

    raise ValueError("Unsupported format. Please upload CSV, XLSX, XLS, or XML.")


# --- 3. SMART COLUMN AUTO-MAPPER (INCLUDING OPTION 1 / 2 / 3) ---
ALIASES: Dict[str, List[str]] = {
    "Handle": ["handle", "slug", "url_handle", "product_handle", "product_id", "group_id"],
    "Title": ["title", "product_name", "name", "item_name", "product_title"],
    "Body (HTML)": ["body (html)", "body_html", "description", "desc", "product_description"],
    "Vendor": ["vendor", "brand", "manufacturer", "supplier"],
    "Type": ["type", "product_type", "category", "product_category"],
    "Tags": ["tags", "keywords", "labels"],
    "Option1 Name": ["option1 name", "option1_name", "variant_name", "attr1_name"],
    "Option1 Value": ["option1 value", "option1_value", "variant", "size", "attr1_value"],
    "Option2 Name": ["option2 name", "option2_name", "attr2_name"],
    "Option2 Value": ["option2 value", "option2_value", "color", "colour", "material", "attr2_value"],
    "Option3 Name": ["option3 name", "option3_name", "attr3_name"],
    "Option3 Value": ["option3 value", "option3_value", "style", "finish", "attr3_value"],
    "Variant SKU": ["variant sku", "sku", "item_code", "part_number", "mpn", "code"],
    "Variant Price": ["variant price", "price", "retail_price", "sell_price", "unit_price", "rrp"],
    "Variant Compare At Price": ["variant compare at price", "compare_at_price", "msrp", "list_price", "old_price"],
    "Variant Inventory Qty": ["variant inventory qty", "qty", "quantity", "stock", "inventory", "available"],
    "Image Src": ["image src", "image", "image_url", "img", "img_url", "picture", "main_image"],
    "Status": ["status", "state", "published", "active"],
}


def norm_col(c: str) -> str:
    return re.sub(r"[^a-z0-9]+", "_", str(c).strip().lower()).strip("_")


def auto_map(columns: List[str]) -> Dict[str, Optional[str]]:
    norm_lookup = {norm_col(c): c for c in columns}
    mapping: Dict[str, Optional[str]] = {}
    used = set()

    for target, names in ALIASES.items():
        found = None
        for alias in names:
            key = norm_col(alias)
            if key in norm_lookup and norm_lookup[key] not in used:
                found = norm_lookup[key]
                break
        if not found:
            for n_col, orig_col in norm_lookup.items():
                if orig_col in used:
                    continue
                if any(len(norm_col(a)) >= 5 and norm_col(a) in n_col for a in names[:2]):
                    found = orig_col
                    break
        if found:
            used.add(found)
        mapping[target] = found
    return mapping


# --- 4. DETERMINISTIC CLEANERS ---
def slugify(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "-", str(text).strip().lower()).strip("-")


def clean_price(raw_val: str) -> Optional[float]:
    s = str(raw_val).strip()
    if not s:
        return None
    s = re.sub(r"(USD|EUR|GBP|CAD|AUD|PKR|[$€£¥₹\s])", "", s, flags=re.IGNORECASE)
    if re.match(r"^-?\d{1,3}(\.\d{3})*,\d{1,2}$", s) or re.match(r"^-?\d+,\d{1,2}$", s):
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", "")
    try:
        return round(float(s), 2)
    except ValueError:
        return None


def clean_html(raw_html: str) -> str:
    text = str(raw_html).strip()
    if not text:
        return ""
    unescaped = html.unescape(text)
    soup = BeautifulSoup(unescaped, "html.parser")
    for bad in soup(["script", "style", "iframe", "form", "object", "embed"]):
        bad.decompose()
    allowed = {"p", "br", "ul", "ol", "li", "strong", "b", "em", "i", "h3", "h4"}
    for tag in soup.find_all(True):
        if tag.name not in allowed:
            tag.unwrap()
        else:
            tag.attrs = {}
    return re.sub(r"\s{2,}", " ", str(soup)).strip()


def clean_image_url(raw_url: str) -> Tuple[str, Optional[str], bool]:
    """Returns (clean_https_url, error_msg_or_None, resolved_cleanly)."""
    u = str(raw_url).strip()
    if not u:
        return "", None, True
    if u.startswith("https://"):
        if " " in u:
            return u.replace(" ", "%20"), "Spaces in HTTPS URL (Auto-Encoded)", True
        return u, None, True
    if u.startswith("http://"):
        return "", "Insecure HTTP Image URL (Stripped — Needs Manual HTTPS Fix)", False
    return "", "Invalid/Relative Image Path (Stripped — Needs Manual URL Fix)", False


def get_val(row: pd.Series, col_map: Dict[str, Optional[str]], key: str) -> str:
    col = col_map.get(key)
    if not col or col not in row.index:
        return ""
    return str(row[col]).strip()


def slice_complete_handles(clean_df: pd.DataFrame, max_rows: int) -> Tuple[pd.DataFrame, bool]:
    """Slices clean_df up to max_rows without ever splitting a product's variants in half."""
    if clean_df.empty or len(clean_df) <= max_rows:
        return clean_df, False
    cut_handle = clean_df.iloc[max_rows]["Handle"]
    end_idx = max_rows
    while end_idx > 0 and clean_df.iloc[end_idx - 1]["Handle"] == cut_handle:
        end_idx -= 1
    if end_idx == 0:
        # Edge case: first single product itself exceeds max_rows
        return clean_df.head(max_rows), True
    return clean_df.iloc[:end_idx], False


# --- 5. SYMMETRICAL NON-OVERLAPPING HEALTH SCORE CALCULATOR ---
SEV_RANK = {"CRITICAL": 3, "HIGH": 2, "MEDIUM": 1}


def compute_honest_scores(
    total_rows: int,
    issues_df: pd.DataFrame,
    blocker_row_nums: set,
) -> Tuple[int, int, dict]:
    n = max(total_rows, 1)
    if issues_df.empty:
        return 100, 100, {"raw_crit": 0, "raw_high": 0, "raw_med": 0, "rem_blockers": 0, "rem_high": 0}

    row_worst_raw: Dict[int, int] = {}
    for _, r in issues_df.iterrows():
        rn = int(r["Row"])
        rk = SEV_RANK.get(str(r["Severity"]), 0)
        if rk > row_worst_raw.get(rn, 0):
            row_worst_raw[rn] = rk

    raw_crit = sum(1 for v in row_worst_raw.values() if v == 3)
    raw_high = sum(1 for v in row_worst_raw.values() if v == 2)
    raw_med = sum(1 for v in row_worst_raw.values() if v == 1)

    raw_penalty = (raw_crit / n) * 80.0 + (raw_high / n) * 40.0 + (raw_med / n) * 15.0
    raw_score = max(0, min(100, round(100.0 - raw_penalty)))

    unresolved = issues_df[issues_df["Resolved"] == "No"]
    row_worst_post: Dict[int, int] = {}
    for rn in blocker_row_nums:
        row_worst_post[rn] = 3

    for _, r in unresolved.iterrows():
        rn = int(r["Row"])
        rk = 3 if rn in blocker_row_nums else 2
        if rk > row_worst_post.get(rn, 0):
            row_worst_post[rn] = rk

    rem_blockers = sum(1 for v in row_worst_post.values() if v == 3)
    rem_high = sum(1 for v in row_worst_post.values() if v == 2)

    clean_penalty = (rem_blockers / n) * 80.0 + (rem_high / n) * 40.0
    clean_score = max(raw_score, min(100, round(100.0 - clean_penalty)))

    return raw_score, clean_score, {
        "raw_crit": raw_crit,
        "raw_high": raw_high,
        "raw_med": raw_med,
        "rem_blockers": rem_blockers,
        "rem_high": rem_high,
    }


# --- 6. TWO-PASS SHOPIFY AUDIT + QUARANTINE + MULTI-OPTION CLEAN ENGINE ---
def run_audit_and_clean(
    df_raw: pd.DataFrame,
    col_map: Dict[str, Optional[str]],
    markup_pct: float,
    clear_inverted_compare: bool,
    quarantine_blockers: bool,
    variant_limit: Optional[int],
):
    issues = []
    manual_review_rows = []
    before_after_diffs = []
    parsed_records = []

    sku_col = col_map.get("Variant SKU")
    sku_counts = {}
    if sku_col and sku_col in df_raw.columns:
        skus = df_raw[sku_col].astype(str).str.strip()
        sku_counts = skus[skus != ""].value_counts().to_dict()

    handle_counts: Dict[str, int] = {}
    handle_order: Dict[str, int] = {}
    handles_forced_draft = set()
    handle_first_title: Dict[str, str] = {}
    handle_opt1_name: Dict[str, str] = {}
    handle_opt2_name: Dict[str, str] = {}
    handle_opt3_name: Dict[str, str] = {}
    handle_has_opt2: Dict[str, bool] = {}
    handle_has_opt3: Dict[str, bool] = {}

    # Pass 1: Normalize rows & collect row-level defects
    for idx, row in df_raw.iterrows():
        row_num = idx + 2
        raw_handle = get_val(row, col_map, "Handle")
        raw_title = get_val(row, col_map, "Title")
        raw_body = get_val(row, col_map, "Body (HTML)")
        raw_vendor = get_val(row, col_map, "Vendor")
        raw_type = get_val(row, col_map, "Type")
        raw_tags = get_val(row, col_map, "Tags")
        raw_opt1_name = get_val(row, col_map, "Option1 Name")
        raw_opt1_val = get_val(row, col_map, "Option1 Value")
        raw_opt2_name = get_val(row, col_map, "Option2 Name")
        raw_opt2_val = get_val(row, col_map, "Option2 Value")
        raw_opt3_name = get_val(row, col_map, "Option3 Name")
        raw_opt3_val = get_val(row, col_map, "Option3 Value")
        raw_sku = get_val(row, col_map, "Variant SKU")
        raw_price = get_val(row, col_map, "Variant Price")
        raw_compare = get_val(row, col_map, "Variant Compare At Price")
        raw_qty = get_val(row, col_map, "Variant Inventory Qty")
        raw_img = get_val(row, col_map, "Image Src")
        raw_status = get_val(row, col_map, "Status") or "active"

        is_blocker = False
        blocker_reasons = []

        # R01 & R02: Handle
        if not raw_handle:
            if raw_title:
                final_handle = slugify(raw_title)
                issues.append((row_num, raw_sku, final_handle, "R01", "CRITICAL", "Missing Handle", "(blank)", f"Auto-Fixed -> {final_handle}", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Handle", "(blank)", final_handle))
            else:
                final_handle = ""
                is_blocker = True
                blocker_reasons.append("Missing both Handle and Title")
                issues.append((row_num, raw_sku, "", "R01", "CRITICAL", "Missing Handle & Title", "(blank)", "Quarantined (Manual Fix Required)", "No"))
        else:
            final_handle = slugify(raw_handle)
            if final_handle != raw_handle:
                issues.append((row_num, raw_sku, final_handle, "R02", "CRITICAL", "Invalid Handle Characters", raw_handle, f"Auto-Fixed -> {final_handle}", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Handle", raw_handle, final_handle))

        if final_handle:
            handle_counts[final_handle] = handle_counts.get(final_handle, 0) + 1
            if final_handle not in handle_order:
                handle_order[final_handle] = len(handle_order)
            if final_handle not in handle_first_title and raw_title:
                handle_first_title[final_handle] = raw_title
            if final_handle not in handle_opt1_name and raw_opt1_name:
                handle_opt1_name[final_handle] = raw_opt1_name
            if final_handle not in handle_opt2_name and raw_opt2_name:
                handle_opt2_name[final_handle] = raw_opt2_name
            if final_handle not in handle_opt3_name and raw_opt3_name:
                handle_opt3_name[final_handle] = raw_opt3_name
            if raw_opt2_val:
                handle_has_opt2[final_handle] = True
            if raw_opt3_val:
                handle_has_opt3[final_handle] = True

        # R04: Duplicate SKU
        if raw_sku and sku_counts.get(raw_sku, 0) > 1:
            issues.append((row_num, raw_sku, final_handle, "R04", "HIGH", "Duplicate SKU", raw_sku, "Flagged (Merchant Must Deduplicate)", "No"))

        # R05 & R06: Price
        parsed_price = clean_price(raw_price)
        if parsed_price is None:
            issues.append((row_num, raw_sku, final_handle, "R05", "CRITICAL", "Unparseable or Blank Price", raw_price, "Flagged -> Product Forced to Draft", "No"))
            final_price_num = None
            final_price_str = "0.00"
            if final_handle:
                handles_forced_draft.add(final_handle)
        else:
            if raw_price not in (str(parsed_price), f"{parsed_price:.2f}"):
                issues.append((row_num, raw_sku, final_handle, "R05", "CRITICAL", "Contaminated Price Format", raw_price, f"Auto-Fixed -> {parsed_price:.2f}", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Variant Price", raw_price, f"{parsed_price:.2f}"))
            if parsed_price <= 0:
                issues.append((row_num, raw_sku, final_handle, "R06", "HIGH", "Zero or Negative Price (Revenue Leak)", raw_price, "Flagged -> Product Forced to Draft", "No"))
                final_price_num = max(0.0, parsed_price)
                final_price_str = f"{final_price_num:.2f}"
                if final_handle:
                    handles_forced_draft.add(final_handle)
            else:
                adj_price = round(parsed_price * (1.0 + markup_pct / 100.0), 2)
                final_price_num = adj_price
                final_price_str = f"{adj_price:.2f}"

        # R07: Compare At Price
        final_compare_str = ""
        if raw_compare:
            parsed_comp = clean_price(raw_compare)
            if parsed_comp is None:
                issues.append((row_num, raw_sku, final_handle, "R07", "MEDIUM", "Invalid Compare-At Price", raw_compare, "Auto-Cleared", "Yes"))
            elif final_price_num is not None and parsed_comp <= final_price_num:
                if clear_inverted_compare:
                    issues.append((row_num, raw_sku, final_handle, "R07", "HIGH", "Inverted Compare-At Price (<= Price)", raw_compare, "Auto-Cleared (Inverted)", "Yes"))
                    before_after_diffs.append((row_num, raw_sku, "Compare At Price", raw_compare, "(cleared)"))
                    final_compare_str = ""
                else:
                    issues.append((row_num, raw_sku, final_handle, "R07", "HIGH", "Inverted Compare-At Price (<= Price)", raw_compare, "Flagged Only (Unresolved)", "No"))
                    final_compare_str = f"{parsed_comp:.2f}"
            else:
                adj_comp = round(parsed_comp * (1.0 + markup_pct / 100.0), 2)
                final_compare_str = f"{adj_comp:.2f}"

        # R10: Stock Quantity
        qty_low = raw_qty.lower()
        if qty_low in ("out of stock", "no", "false", "0", ""):
            final_qty = "0"
            if qty_low in ("out of stock", "no", "false"):
                issues.append((row_num, raw_sku, final_handle, "R10", "HIGH", "Text in Stock Column", raw_qty, "Auto-Fixed -> 0", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Inventory Qty", raw_qty, "0"))
        elif re.match(r"^-?\d+$", raw_qty):
            n_qty = int(raw_qty)
            if n_qty < 0:
                issues.append((row_num, raw_sku, final_handle, "R10", "HIGH", "Negative Stock Quantity", raw_qty, "Auto-Fixed -> 0", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Inventory Qty", raw_qty, "0"))
                final_qty = "0"
            else:
                final_qty = str(n_qty)
        else:
            issues.append((row_num, raw_sku, final_handle, "R10", "HIGH", "Ambiguous Stock Text", raw_qty, "Flagged (Set to 0 — Verify Real Qty)", "No"))
            final_qty = "0"

        # R11: Dirty HTML
        final_body = clean_html(raw_body)
        if raw_body and final_body != raw_body:
            if any(tok in raw_body.lower() for tok in ("<script", "<style", "<iframe", "style=", "class=", "&amp;")):
                issues.append((row_num, raw_sku, final_handle, "R11", "HIGH", "Dirty HTML / Script / Inline Style", raw_body[:60], "Auto-Sanitized", "Yes"))
                before_after_diffs.append((row_num, raw_sku, "Body (HTML)", raw_body[:80], final_body[:80]))

        # R12: Image URL
        final_img, img_err, img_resolved = clean_image_url(raw_img)
        if img_err:
            sev = "MEDIUM" if img_resolved else "HIGH"
            issues.append((row_num, raw_sku, final_handle, "R12", sev, "Image URL Defect", raw_img[:60], img_err, "Yes" if img_resolved else "No"))

        # R13: Status
        status_map = {
            "active": "active", "enabled": "active", "published": "active", "live": "active", "1": "active",
            "draft": "draft", "disabled": "draft", "hidden": "draft", "0": "draft",
            "archived": "archived",
        }
        final_status = status_map.get(raw_status.lower(), "draft")
        if raw_status.lower() not in ("active", "draft", "archived"):
            issues.append((row_num, raw_sku, final_handle, "R13", "CRITICAL", "Invalid Shopify Status", raw_status, f"Auto-Fixed -> {final_status}", "Yes"))
            before_after_diffs.append((row_num, raw_sku, "Status", raw_status, final_status))

        parsed_records.append(
            {
                "row_num": row_num,
                "handle": final_handle,
                "title": raw_title,
                "body": final_body,
                "vendor": raw_vendor,
                "type": raw_type,
                "tags": raw_tags,
                "opt1_val": raw_opt1_val,
                "opt2_val": raw_opt2_val,
                "opt3_val": raw_opt3_val,
                "sku": raw_sku,
                "price": final_price_str,
                "compare": final_compare_str,
                "qty": final_qty,
                "img": final_img,
                "status": final_status,
                "is_blocker": is_blocker,
                "blocker_reasons": blocker_reasons,
            }
        )

    # Group scattered rows with the same Handle contiguously (Shopify CSV requirement)
    parsed_records.sort(key=lambda r: (handle_order.get(r["handle"], 10**9), r["row_num"]))

    # Pass 2: Cross-row Shopify structural validation (Option1 + Option2 + Option3)
    over_limit_handles = (
        {h for h, cnt in handle_counts.items() if cnt > variant_limit}
        if variant_limit is not None
        else set()
    )
    seen_handles = set()
    seen_handle_options = set()
    shopify_rows = []
    blocker_row_nums = set()

    for rec in parsed_records:
        h = rec["handle"]
        row_num = rec["row_num"]
        blocker = rec["is_blocker"]
        reasons = list(rec["blocker_reasons"])

        if h in over_limit_handles:
            blocker = True
            reasons.append(f"Handle exceeds configured {variant_limit}-variant limit ({handle_counts[h]} rows)")
            issues.append((row_num, rec["sku"], h, "R08", "CRITICAL", f"Variant Limit Exceeded (>{variant_limit})", f"{handle_counts[h]} variants", "Quarantined (Split Product or Adjust Limit)", "No"))

        if h and not handle_first_title.get(h):
            blocker = True
            reasons.append("No Title found on any row for this Handle")
            issues.append((row_num, rec["sku"], h, "R03", "CRITICAL", "Missing Parent Product Title", "(blank)", "Quarantined (Manual Fix Required)", "No"))

        is_multi = handle_counts.get(h, 0) > 1
        needs_opt2 = handle_has_opt2.get(h, False)
        needs_opt3 = handle_has_opt3.get(h, False)

        o1 = rec["opt1_val"] if rec["opt1_val"] else ("Default Title" if not is_multi else "")
        o2 = rec["opt2_val"] if needs_opt2 else ""
        o3 = rec["opt3_val"] if needs_opt3 else ""

        if is_multi and h not in over_limit_handles:
            missing_parts = []
            if not o1:
                missing_parts.append("Option1")
            if needs_opt2 and not o2:
                missing_parts.append("Option2")
            if needs_opt3 and not o3:
                missing_parts.append("Option3")

            combo_label = " / ".join(p for p in (o1, o2, o3) if p) or "(blank)"
            combo_key = (h, o1.lower(), o2.lower(), o3.lower())

            if missing_parts:
                blocker = True
                reasons.append(f"Missing {' & '.join(missing_parts)} Value on multi-variant Handle '{h}'")
                issues.append((row_num, rec["sku"], h, "R09", "CRITICAL", "Missing Variant Option Value", combo_label, "Quarantined (Prevents Shopify Option Mismatch)", "No"))
            elif combo_key in seen_handle_options:
                blocker = True
                reasons.append(f"Duplicate variant combination ('{combo_label}') under Handle '{h}'")
                issues.append((row_num, rec["sku"], h, "R09", "CRITICAL", "Duplicate Variant Option Combination", combo_label, "Quarantined (Prevents 'Variant Already Exists')", "No"))
            else:
                seen_handle_options.add(combo_key)

        if blocker:
            blocker_row_nums.add(row_num)
            manual_review_rows.append(
                {
                    "Source Row": row_num,
                    "Handle": h,
                    "SKU": rec["sku"],
                    "Title": rec["title"],
                    "Price": rec["price"],
                    "Why Quarantined": "; ".join(reasons),
                }
            )
            if quarantine_blockers:
                continue

        is_first = h not in seen_handles
        if h:
            seen_handles.add(h)

        group_status = "draft" if h in handles_forced_draft else rec["status"]
        opt1_name_out = handle_opt1_name.get(h) or ("Option 1" if is_multi else "Title")
        opt2_name_out = (handle_opt2_name.get(h) or "Option 2") if needs_opt2 else ""
        opt3_name_out = (handle_opt3_name.get(h) or "Option 3") if needs_opt3 else ""

        shopify_rows.append(
            {
                "Handle": h,
                "Title": handle_first_title.get(h, rec["title"]) if is_first else "",
                "Body (HTML)": rec["body"] if is_first else "",
                "Vendor": rec["vendor"] if is_first else "",
                "Type": rec["type"] if is_first else "",
                "Tags": rec["tags"] if is_first else "",
                "Published": ("TRUE" if group_status == "active" else "FALSE") if is_first else "",
                "Option1 Name": opt1_name_out if is_first else "",
                "Option1 Value": o1,
                "Option2 Name": opt2_name_out if is_first else "",
                "Option2 Value": o2,
                "Option3 Name": opt3_name_out if is_first else "",
                "Option3 Value": o3,
                "Variant SKU": rec["sku"],
                "Variant Grams": "0",
                "Variant Inventory Tracker": "shopify",
                "Variant Inventory Qty": rec["qty"],
                "Variant Inventory Policy": "deny",
                "Variant Fulfillment Service": "manual",
                "Variant Price": rec["price"],
                "Variant Compare At Price": rec["compare"],
                "Variant Requires Shipping": "TRUE",
                "Variant Taxable": "TRUE",
                "Image Src": rec["img"],
                "Status": group_status if is_first else "",
            }
        )

    issues_df = pd.DataFrame(
        issues,
        columns=["Row", "SKU", "Handle", "Rule", "Severity", "Issue Type", "Raw Value", "Action Taken", "Resolved"],
    )
    clean_df = pd.DataFrame(shopify_rows)
    manual_df = pd.DataFrame(manual_review_rows)
    diff_df = pd.DataFrame(before_after_diffs, columns=["Row", "SKU", "Column", "Before (Raw)", "After (Auto-Fixed)"])

    raw_score, clean_score, breakdown = compute_honest_scores(len(df_raw), issues_df, blocker_row_nums)
    return issues_df, clean_df, manual_df, diff_df, raw_score, clean_score, breakdown


def build_styled_excel(issues_df: pd.DataFrame, manual_df: pd.DataFrame, summary_rows: List[dict]) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Executive_Summary", index=False)

        if issues_df.empty:
            pd.DataFrame([{"Status": "Zero issues detected"}]).to_excel(writer, sheet_name="All_Detected_Issues", index=False)
        else:
            issues_df.to_excel(writer, sheet_name="All_Detected_Issues", index=False)

        if manual_df.empty:
            pd.DataFrame([{"Status": "No quarantined rows — 100% passed structural checks"}]).to_excel(
                writer, sheet_name="Needs_Manual_Review", index=False
            )
        else:
            manual_df.to_excel(writer, sheet_name="Needs_Manual_Review", index=False)

        header_fill = PatternFill(start_color="0F766E", end_color="0F766E", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        for ws in writer.book.worksheets:
            ws.freeze_panes = "A2"
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
            for col in ws.columns:
                max_len = max((len(str(c.value or "")) for c in col[:200]), default=10)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 48)

    return buf.getvalue()


# --- 7. STREAMLIT WEB UI ---
st.title(f"FeedMedic {APP_VERSION} — Shopify Supplier Feed Auditor & Cleaner")
st.caption(
    "Catch silent Shopify import crashes, $0.00 price leaks, duplicate SKUs, multi-option variant collisions, and dirty HTML before touching your live store."
)
st.info(
    "Zero Data Storage: Files are processed strictly in RAM and permanently discarded when you close this tab. No Shopify login or API access required."
)

if "df_raw" not in st.session_state:
    st.session_state.df_raw = None
    st.session_state.source_name = ""
    st.session_state.last_upload_token = None

c_up, c_demo = st.columns([2, 1])
with c_up:
    uploaded = st.file_uploader(
        "Upload Supplier Feed (CSV, Excel, or XML — max 25 MB)",
        type=["csv", "xlsx", "xls", "xml"],
    )
    if uploaded is None:
        st.session_state.last_upload_token = None
    else:
        token = f"{uploaded.name}:{uploaded.size}"
        if token != st.session_state.last_upload_token:
            try:
                st.session_state.df_raw = read_uploaded_file(uploaded)
                st.session_state.source_name = uploaded.name
                st.session_state.last_upload_token = token
            except Exception as e:
                st.error(f"Could not read file: {e}")

with c_demo:
    st.write("")
    st.write("")
    if st.button("Load Messy Demo Feed (650 SKUs)", type="primary", use_container_width=True):
        st.session_state.df_raw = generate_demo_feed()
        st.session_state.source_name = "demo_messy_supplier_650_skus.csv (Synthetic Multi-Option Demo)"

if st.session_state.df_raw is not None:
    df_raw = st.session_state.df_raw
    st.success(f"Loaded `{st.session_state.source_name}` — **{len(df_raw):,} rows** & **{len(df_raw.columns)} columns** detected.")

    detected = auto_map(list(df_raw.columns))
    col_options = ["(None)"] + list(df_raw.columns)

    with st.expander("Step 1: Verify Column Mapping & Cleanup Rules", expanded=False):
        user_map: Dict[str, Optional[str]] = {}
        cols_grid = st.columns(3)
        for i, target in enumerate(ALIASES.keys()):
            default_col = detected.get(target)
            default_idx = col_options.index(default_col) if default_col in col_options else 0
            chosen = cols_grid[i % 3].selectbox(f"Shopify: {target}", col_options, index=default_idx)
            user_map[target] = None if chosen == "(None)" else chosen

        opt1, opt2, opt3, opt4 = st.columns(4)
        markup_pct = opt1.number_input("Optional Markup % on Valid Prices", min_value=0.0, max_value=500.0, value=0.0, step=5.0)
        clear_inv = opt2.checkbox("Clear Inverted Compare-At Prices", value=True)
        quarantine_on = opt3.checkbox("Quarantine Import-Blocking Rows", value=True)
        v_limit_label = opt4.selectbox(
    "Max Variants / Product (R08)",
    ["100 (Legacy Safe)", "2048 (Shopify Expanded)", "No Limit"],
    index=1,
)
        variant_limit_map = {"100 (Legacy Safe)": 100, "2048 (Shopify Expanded)": 2048, "No Limit": None}
        variant_limit = variant_limit_map[v_limit_label]

    with st.spinner("Auditing Shopify CSV rules and sanitizing feed..."):
        issues_df, clean_df, manual_df, diff_df, raw_score, clean_score, bd = run_audit_and_clean(
            df_raw, user_map, markup_pct, clear_inv, quarantine_on, variant_limit
        )

    export_df, sliced_oversized = slice_complete_handles(clean_df, FREE_EXPORT_LIMIT)

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Raw Feed Health Score", f"{raw_score} / 100")
    m2.metric("Post-Clean Health Score", f"{clean_score} / 100", f"+{clean_score - raw_score} pts")
    m3.metric("Total Issues Caught", f"{len(issues_df):,}")
    m4.metric("Quarantined Rows", f"{len(manual_df):,}")

    if clean_score < 100:
        st.caption(
            f"Honest Score Breakdown: Post-Clean Score is {clean_score}/100 (not 100/100) because "
            f"{bd['rem_blockers']} import-blocking rows require manual structural fixes and "
            f"{bd['rem_high']} clean rows still carry merchant-decision flags (duplicate SKUs, $0 draft prices, ambiguous stock, or stripped broken image URLs)."
        )

    t1, t2, t3, t4, t5 = st.tabs(
        [
            "Summary by Rule",
            "Before vs. After Diff",
            "All Detected Issues",
            "Quarantined (Manual Review)",
            "Clean Shopify Preview",
        ]
    )

    with t1:
        if issues_df.empty:
            st.success("Zero errors detected. Your feed passed all checks.")
        else:
            rule_summary = (
                issues_df.groupby(["Rule", "Severity", "Issue Type", "Resolved"], as_index=False)
                .size()
                .rename(columns={"size": "Affected Rows"})
                .sort_values(by=["Severity", "Affected Rows"], ascending=[True, False])
            )
            st.dataframe(rule_summary, use_container_width=True, hide_index=True)

    with t2:
        if diff_df.empty:
            st.info("No automatic transformations were needed on this file.")
        else:
            st.caption("Exact side-by-side log of deterministic fixes applied safely by FeedMedic:")
            st.dataframe(diff_df.head(500), use_container_width=True, hide_index=True)

    with t3:
        if issues_df.empty:
            st.success("No individual row issues found.")
        else:
            sev_filter = st.multiselect("Filter by Severity", ["CRITICAL", "HIGH", "MEDIUM"], default=["CRITICAL", "HIGH", "MEDIUM"])
            filtered = issues_df[issues_df["Severity"].isin(sev_filter)]
            st.dataframe(filtered.head(1000), use_container_width=True, hide_index=True)

    with t4:
        if manual_df.empty:
            st.success("No rows were quarantined.")
        else:
            st.warning(
                "These rows would risk crashing Shopify's CSV importer (e.g., variant limit exceeded or colliding Option1/Option2/Option3 values). "
                "They are excluded from the Clean CSV and saved in the Excel Audit Report."
            )
            st.dataframe(manual_df, use_container_width=True, hide_index=True)

    with t5:
        st.caption(f"Formatted strictly to Shopify's official Product CSV parent-child structure ({len(export_df):,} complete-product rows in free export).")
        st.dataframe(export_df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Download Outputs")

    csv_bytes = export_df.to_csv(index=False).encode("utf-8")

    summary_metrics = [
        {"Metric": "FeedMedic Version", "Value": APP_VERSION},
        {"Metric": "Source File", "Value": st.session_state.source_name},
        {"Metric": "Total Raw Rows", "Value": len(df_raw)},
        {"Metric": "Raw Feed Health Score", "Value": f"{raw_score} / 100"},
        {"Metric": "Post-Clean Health Score", "Value": f"{clean_score} / 100"},
        {"Metric": "Total Issues Detected", "Value": len(issues_df)},
        {"Metric": "Quarantined Import-Blocking Rows", "Value": len(manual_df)},
        {"Metric": "Clean Rows With Unresolved Merchant Flags", "Value": bd["rem_high"]},
        {"Metric": "Total Clean Shopify-Ready Rows", "Value": len(clean_df)},
        {"Metric": "Free Export Rows (Complete Products Only)", "Value": len(export_df)},
    ]
    excel_bytes = build_styled_excel(issues_df, manual_df, summary_metrics)

    d1, d2 = st.columns(2)
    d1.download_button(
        label=f"Download Clean Shopify CSV ({len(export_df):,} Rows — Free)",
        data=csv_bytes,
        file_name="feedmedic_clean_shopify.csv",
        mime="text/csv",
        type="primary",
        use_container_width=True,
        disabled=export_df.empty,
    )
    d2.download_button(
        label="Download Full 3-Sheet Excel Audit Report (.xlsx — Free)",
        data=excel_bytes,
        file_name="feedmedic_audit_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # Actionable Conversion Trigger right where the user hits the cap or quarantine blockers
    remaining_clean = len(clean_df) - len(export_df)
    if remaining_clean > 0 or len(manual_df) > 0:
        msg_parts = []
        if remaining_clean > 0:
            msg_parts.append(
                f"**{remaining_clean:,} more clean rows** remain beyond the free {FREE_EXPORT_LIMIT}-row preview (capped cleanly at complete product boundaries)"
            )
        if len(manual_df) > 0:
            msg_parts.append(f"**{len(manual_df):,} quarantined rows** need manual variant/title restructuring")
        st.warning(
            "Next Step: " + " and ".join(msg_parts) + ". "
            "Want the complete catalog cleaned, restructured, and verified in a Shopify staging store within 48 hours (starting at $95 for up to 2,500 SKUs)?"
        )
        if sliced_oversized:
            st.caption("Note: Your first single product alone exceeded 500 variants, so the preview was capped at row 500.")

        quick_subject = f"Feed Rescue Quote — {len(df_raw):,} Row Supplier Feed"
        quick_body = (
            f"Hi Salar,\n\nI just ran my feed through FeedMedic {APP_VERSION}.\n"
            f"- Total rows: {len(df_raw):,}\n"
            f"- Quarantined rows: {len(manual_df):,}\n"
            f"- Post-Clean Score: {clean_score}/100\n\n"
            "I'm attaching a 30-50 row sample (or the full file) for a quote.\n"
        )
        st.link_button(
            "Send a 30-50 Row Sample via Gmail for a Free Diagnostic & Quote",
            build_gmail_url(CONTACT_EMAIL.strip(), quick_subject, quick_body),
            use_container_width=True,
        )

# --- 8. CLEAN B2B SERVICES & ZERO-MAILTO CONTACT FOOTER ---
st.divider()
st.subheader("Need Full Catalog Rescue or a Reusable Supplier Cleaner?")

f1, f2 = st.columns(2)
with f1:
    st.markdown("**1. Emergency Feed Rescue — $95 (up to 2,500 SKUs)**")
    st.caption(
        "One-off cleanup for 1 supplier format (up to 2,500 SKUs): multi-option variant restructuring, custom markup rules, "
        "deduplication guidance, and live Shopify staging import verification. 48-hour delivery. Larger catalogs (2,500+ SKUs) quoted after a free 30–50 row sample check."
    )
with f2:
    st.markdown("**2. Private 1-Click Supplier Cleaner — $250**")
    st.caption(
        "Built for stores & agencies receiving recurring weekly/monthly XML, CSV, or Excel updates from the same supplier. "
        "Custom reusable tool + automated Excel error report. Run it yourself anytime with zero coding."
    )

c_btn1, c_btn2 = st.columns(2)
with c_btn1:
    clean_email = CONTACT_EMAIL.strip()
    if clean_email and "@" in clean_email and not clean_email.startswith("http"):
        footer_url = build_gmail_url(
            clean_email,
            "Shopify Supplier Feed Cleanup — 30-50 Row Sample",
            "Hi Salar,\n\nI want to share a 30-50 row sample of my messy supplier feed for a quick diagnostic.\n\nTotal SKUs in full catalog: \nRecurring supplier? (Yes/No): \n",
        )
        st.link_button("Send a 30-50 Row Sample via Gmail", footer_url, use_container_width=True)
        st.caption("Using Outlook, Apple Mail, or another client? Copy email in 1 click:")
        st.code(clean_email, language=None)
    else:
        st.error("Config Error: CONTACT_EMAIL must be a plain email address (no http/https).")

with c_btn2:
    clean_li = LINKEDIN_URL.strip()
    if clean_li and clean_li.startswith("https://"):
        st.link_button("Connect & Message Me on LinkedIn", clean_li, use_container_width=True)
        st.caption("Direct LinkedIn profile link (click copy icon on the right):")
        st.code(clean_li, language=None)
    else:
        st.error("Config Error: LINKEDIN_URL must start with https://")
