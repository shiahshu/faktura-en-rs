#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fakturaen.py — native text PDF invoice extraction by pure rules (no LLM)
=======================================================================
Styled after invoice.py / racun.py output; reuses match_invoices.py pdfplumber rule parsing:
- Tkinter file dialog selects a single native (text-layer) PDF invoice file
- Pure regex + table-row rules: header (invoice id / seller & buyer PIB / Datum prometa / official amount) + line items
- Per-page split: a PDF with multiple invoices (e.g. *_sef.pdf batch files) yields one invoice per page
- Prepayment deduction: detects avans lines (Ukupan iznos osnovice umanjen za osnovicu po avansu etc.),
  outputs Prepayment_Deduction and Actual_Payment columns (avans3.py semantics)
- No LLM calls, extraction only
- Output: invoice_summary.xlsx + PowerShell console table (racun.py style, no Chinese translation column)

Usage:
    python fakturaen.py
"""

import os
import re
import time
import unicodedata

import pandas as pd
import pdfplumber

from tkinter import filedialog, Tk

EXCEL_FILE = "invoice_summary.xlsx"

# ===================== Utility functions =====================


def safe_str(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).replace("\u00a0", " ").strip()


def universal_clean(val):
    """Number cleaning: unify thousands-separator formats to float"""
    s = str(val).strip().replace(" ", "").replace("　", "")
    if not s or s == "nan" or s == "0":
        return 0.0
    s = re.sub(r"[^-0-9,.]", "", s)
    if "." in s and "," in s:
        if s.rfind(".") > s.rfind(","):
            s = s.replace(",", "")
        else:
            s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def parse_numeric_or_none(val):
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except Exception:
        pass
    if isinstance(val, (int, float)):
        return float(val)
    s = safe_str(val)
    if not s or not re.search(r"\d", s):
        return None
    return universal_clean(s)


def normalize_text(text):
    text = safe_str(text)
    if not text:
        return ""
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("\u200b", "")
    text = re.sub(r"\s+", " ", text)
    return text.strip().lower()


def find_first_match(patterns, text):
    for pattern in patterns:
        m = re.search(pattern, text, re.I | re.M)
        if m:
            return safe_str(m.group(1))
    return ""


def is_number_token(token):
    token = safe_str(token)
    return bool(token) and bool(re.fullmatch(r"[\-+]?\d[\d\.,]*", token))


def is_summary_line(line):
    text = normalize_text(line)
    return any(
        token in text
        for token in [
            "zbir stavki",
            "ukupna osnovica",
            "ukupan pdv",
            "ukupan iznos fakture",
            "iznos za placanje",
            "ukupan iznos osnovice",
            "ukupan pdv umanjen",
            "napomena",
            "dan nastanka pdv obaveze",
        ]
    )


def is_header_line(line):
    text = normalize_text(line)
    # requires opis + kolicin + cena present;
    # single-column compatibility abbreviations: jedinica mere / jed. mere / jed mere / j.m. / mere
    has_unit = any(
        u in text
        for u in ("jedinica mere", "jed. mere", "jed mere", "j.m.", "j m", "mere")
    )
    return (
        "opis" in text
        and "kolicin" in text
        and "cena" in text
        and has_unit
    )


def looks_like_continuation(line):
    text = safe_str(line)
    if not text:
        return False
    if normalize_text(text).startswith("sifra"):
        return False
    if is_summary_line(text) or is_header_line(text):
        return False
    if ":" in text:
        return False
    if len(text) > 60:
        return False
    if re.fullmatch(r"[A-Za-z0-9čćžšđČĆŽŠĐ\s().,/%\-]+", text) is None:
        return False
    return True


def parse_item_line(line, has_discount=True):
    """
    Parse an item line, supporting three column layouts (determined by whether the header has a discount column):
      with discount standard: desc qty price unit discount net tax_rate        (numeric tax rate)
      with discount special: desc qty price unit discount net tax_status       (tax status text, e.g. special taxation)
      no discount: desc qty price unit net tax_rate
    Trailing non-numeric tokens are treated as tax status text (special taxation).
    """
    tokens = [tok for tok in safe_str(line).split() if tok]
    if len(tokens) < 6:
        return None

    # strip trailing non-numeric tokens → tax status text (special taxation)
    i = len(tokens) - 1
    tax_status_parts = []
    while i >= 0 and not is_number_token(tokens[i]):
        tax_status_parts.insert(0, tokens[i])
        i -= 1
    tax_status = " ".join(tax_status_parts).strip()
    body = tokens[: i + 1]

    if has_discount:
        # layout: qty price unit discount net [tax_rate|tax_status]
        # 6 fields (standard, numeric tax rate)
        if (
            len(body) >= 6
            and is_number_token(body[-1])
            and is_number_token(body[-2])
            and is_number_token(body[-3])
            and not is_number_token(body[-4])
            and is_number_token(body[-5])
            and is_number_token(body[-6])
        ):
            desc = " ".join(body[:-6]).strip()
            if not desc:
                return None
            return {
                "desc": desc,
                "qty": universal_clean(body[-6]),
                "price": universal_clean(body[-5]),
                "unit": body[-4],
                "discount": universal_clean(body[-3]),
                "net": universal_clean(body[-2]),
                "tax_rate": universal_clean(body[-1]),
                "tax_status": tax_status,
            }
        # 5 fields (special taxation, tax status text)
        if (
            len(body) >= 5
            and is_number_token(body[-1])
            and is_number_token(body[-2])
            and not is_number_token(body[-3])
            and is_number_token(body[-4])
            and is_number_token(body[-5])
        ):
            desc = " ".join(body[:-5]).strip()
            if not desc:
                return None
            return {
                "desc": desc,
                "qty": universal_clean(body[-5]),
                "price": universal_clean(body[-4]),
                "unit": body[-3],
                "discount": universal_clean(body[-2]),
                "net": universal_clean(body[-1]),
                "tax_rate": 0.0,
                "tax_status": tax_status or "posebni postupci oporezivanja",
            }
        return None

    # no discount: qty price unit net [tax_rate|tax_status]
    if (
        len(body) >= 5
        and is_number_token(body[-1])
        and is_number_token(body[-2])
        and not is_number_token(body[-3])
        and is_number_token(body[-4])
        and is_number_token(body[-5])
    ):
        desc = " ".join(body[:-5]).strip()
        if not desc:
            return None
        return {
            "desc": desc,
            "qty": universal_clean(body[-5]),
            "price": universal_clean(body[-4]),
            "unit": body[-3],
            "discount": 0.0,
            "net": universal_clean(body[-2]),
            "tax_rate": universal_clean(body[-1]),
            "tax_status": tax_status,
        }
    return None


# ===================== Prepayment deduction extraction (avans3.py semantics) =====================

def extract_avans_deduction(page_text):
    """
    Extract the prepayment (avans) deduction from an eFaktura page,
    dynamically iterating over all tax-rate brackets.

    Relevant lines (one group per tax rate):
      Ukupna osnovica - stopa 20%: 17.837,55                 (base before deduction)
      Ukupan iznos osnovice umanjen za osnovicu po avansu     (base after deduction)
      Ukupan PDV - stopa 20%: 3.567,51                        (VAT before deduction)
      Ukupan PDV umanjen za PDV po avansu                     (VAT after deduction)

    Deduction = Σ per rate [(base before - base after) + (VAT before - VAT after)].
    The system always prints these lines; when avans is zero the deduction is 0.
    """
    text = normalize_text(page_text)

    rates = sorted({int(r) for r in re.findall(r"stopa[^\d]*(\d+)", text)})
    if not rates:
        return 0.0

    total_ded = 0.0
    for rate in rates:
        def grab(pattern):
            m = re.search(pattern, text)
            return universal_clean(m.group(1)) if m else 0.0

        base_total = grab(rf"ukupna osnovica[^\d]*stopa[^\d]*{rate}[^\d]*([0-9][0-9.,]*)")
        avans_base = grab(
            rf"osnovice umanjen za osnovicu po avansu[^\d]*stopa[^\d]*{rate}[^\d]*([0-9][0-9.,]*)"
        )
        pdv_total = grab(rf"ukupan pdv[^\d]*stopa[^\d]*{rate}[^\d]*([0-9][0-9.,]*)")
        avans_pdv = grab(
            rf"pdv umanjen za pdv po avansu[^\d]*stopa[^\d]*{rate}[^\d]*([0-9][0-9.,]*)"
        )

        total_ded += round((base_total - avans_base) + (pdv_total - avans_pdv), 2)

    return round(total_ded, 2)


# ===================== Header extraction (seller/buyer PIB) =====================

def extract_header_fields(page_text):
    invoice_id = find_first_match([
        r"(?:broj\s*fakture|broj\s*racuna|broj\s*računa|faktura\s*br\.?|račun\s*br\.?|racun\s*br\.?)[:\s]*([A-Z0-9\-\/]+)",
        # eFaktura continuation pages often only retain: Generisao sistem eFaktura pod brojem:00084/26
        r"generisao\s+sistem\s+eFaktura\s+pod\s+brojem[:\s]*([A-Z0-9\-\/]+)",
        r"\b(IF\d{2}-\d{3,})\b",
        r"\b([0-9]{2,4}(?:-[0-9]{2,}){1,4})\b",
    ], page_text)

    def validate_pib(val):
        s = safe_str(val)
        return s if re.fullmatch(r"[0-9]{9}", s) else ""

    # buyer PIB: PIB kupca; (?!\d) avoids truncating the first 9 digits of a longer number
    buyer_pib = find_first_match([
        r"\bPIB\s+kupca[:\s]*([0-9]{9})(?!\d)",
        r"\bPIB\s+customer[:\s]*([0-9]{9})(?!\d)",
    ], page_text)
    buyer_pib = validate_pib(buyer_pib)

    # seller PIB: PIB inside the Prodavac block; fallback to the non-buyer PIB
    seller_pib = ""
    prodavac_idx = re.search(r"\bProdavac\b", page_text, re.I)
    if prodavac_idx:
        seg = page_text[prodavac_idx.start(): prodavac_idx.start() + 500]
        seller_pib = find_first_match([
            r"\bPIB\b[:\s]*([0-9]{9})(?!\d)",
            r"\bPIB\s+([0-9]{9})(?!\d)",
        ], seg)
        seller_pib = validate_pib(seller_pib)
    if not seller_pib:
        for m in re.finditer(r"\bPIB\b[^0-9]{0,3}([0-9]{9})(?!\d)", page_text, re.I):
            cand = safe_str(m.group(1))
            cand = validate_pib(cand)
            if cand and cand != buyer_pib:
                seller_pib = cand
                break

    datum_prometa = find_first_match([
        r"Datum prometa[:\s]*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})",
        r"Date of supply[:\s]*([0-9]{1,2}[./-][0-9]{1,2}[./-][0-9]{2,4})",
    ], page_text)

    # official amount: prefer Iznos za plaćanje, fallback Ukupan iznos fakture / Grand total
    # [^\r\n]*? keeps the match on the current line, avoiding swallowing the next line's number
    grand_total = find_first_match([
        r"Iznos za plaćanje[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Iznos za placanje[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Ukupan iznos fakture[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Grand total[^\r\n]*?([0-9][0-9\., ]*)\s*$",
    ], page_text)

    # official footer net / vat / gross (shown when reconciling differences against extracted data)
    footer_net = find_first_match([
        r"Ukupna osnovica[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Total net[^\r\n]*?([0-9][0-9\., ]*)\s*$",
    ], page_text)
    footer_vat = find_first_match([
        r"Ukupan PDV[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Ukupan porez[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Total VAT[^\r\n]*?([0-9][0-9\., ]*)\s*$",
    ], page_text)
    footer_total = find_first_match([
        r"Ukupan iznos fakture[^\r\n]*?([0-9][0-9\., ]*)\s*$",
    ], page_text)

    return {
        "invoice_id": safe_str(invoice_id),
        "seller_pib": safe_str(seller_pib),
        "buyer_pib": safe_str(buyer_pib),
        "datum_prometa": safe_str(datum_prometa),
        "grand_total": universal_clean(grand_total),
        "prep_deduction": extract_avans_deduction(page_text),
        "footer_net": universal_clean(footer_net),
        "footer_vat": universal_clean(footer_vat),
        "footer_total": universal_clean(footer_total),
    }


# ===================== Line item extraction (from match_invoices.py) =====================

def extract_items_from_lines(lines):
    items = []
    started = False
    has_discount = True
    current_item = None
    continuations = []

    for line in lines:
        if not started:
            if is_header_line(line):
                started = True
                has_discount = "umanjenje" in normalize_text(line)
            continue

        if is_summary_line(line):
            break

        if normalize_text(line).startswith("sifra"):
            continue

        parsed = parse_item_line(line, has_discount=has_discount)
        if parsed:
            if current_item:
                if continuations:
                    current_item["desc"] = " ".join(
                        [current_item["desc"]] + continuations
                    ).strip()
                items.append(current_item)
            current_item = parsed
            continuations = []
            continue

        if current_item and looks_like_continuation(line):
            if current_item.get("tax_status"):
                # special taxation tax status text may span multiple lines (e.g. Posebni / postupci / oporezivanja)
                current_item["tax_status"] = (
                    current_item["tax_status"] + " " + line
                ).strip()
            else:
                continuations.append(line)

    if current_item:
        if continuations:
            current_item["desc"] = " ".join(
                [current_item["desc"]] + continuations
            ).strip()
        items.append(current_item)

    return items


# ===================== Single PDF extraction (per-page multi-invoice split) =====================

def process_pdf(file_path):
    invoices = []
    with pdfplumber.open(file_path) as pdf:
        for idx, page in enumerate(pdf.pages):
            page_text = page.extract_text(x_tolerance=2, y_tolerance=2) or ""
            lines = [safe_str(l) for l in page_text.splitlines() if safe_str(l)]
            if not lines:
                continue

            header = extract_header_fields(page_text)
            items = extract_items_from_lines(lines)

            if not header["invoice_id"]:
                print(f"  ⏭️ Page {idx+1} skipped (no invoice id found)")
                continue

            if not items:
                print(f"  ⚠️ Page {idx+1} {header['invoice_id']} yielded no line items (header/row format did not match):")
                for ln in lines[:30]:
                    print(f"     | {ln}")

            inv_id = header["invoice_id"]

            # continuation pages: same invoice id on the next page is merged or skipped
            if invoices and inv_id == invoices[-1]["header"]["invoice_id"]:
                if items:
                    print(f"  🔁 Page {idx+1} is a continuation of invoice {inv_id}, appended {len(items)} item rows")
                    invoices[-1]["items"].extend(items)
                else:
                    print(f"  🔁 Page {idx+1} is a continuation of invoice {inv_id} (no items), skipped")
                # a continuation page may also carry duplicate footer or prepayment deduction info
                for key in ("grand_total", "footer_total", "prep_deduction", "seller_pib", "buyer_pib", "datum_prometa"):
                    if header.get(key):
                        invoices[-1]["header"][key] = header[key]
                continue

            invoices.append({"header": header, "items": items})

    if not invoices:
        print(f"⚠️ No invoice pages recognized: {os.path.basename(file_path)}")
        return []

    return invoices


# ===================== Credit note keywords =====================

CREDIT_NOTE_KW = [
    "knjizno odobrenje", "knjižno odobrenje",
    "storno", "storniranje", "storno račun", "storno faktura",
    "credit note", "credit memo", "credit memorandum",
    "odobrenje", "umanjenje", "smanjenje",
]


# ===================== Build report (racun.py style, no Chinese translation column) =====================

def build_report(invoices):
    rows = []
    g = {"qty": 0.0, "net": 0.0, "tax": 0.0, "prep": 0.0, "actual": 0.0}

    for inv in invoices:
        header = inv["header"]
        items = inv["items"]
        inv_id = header["invoice_id"]
        seller_pib = header["seller_pib"]
        buyer_pib = header["buyer_pib"]
        datum = header["datum_prometa"]
        f_total = header["grand_total"]
        footer_total = header.get("footer_total", 0.0)

        # prepayment deduction is derived from the footer gross total vs the final amount due.
        # only when both exist and the footer gross is higher is this treated as a real deduction;
        # this avoids mistaking the post-deduction base/vat for a deduction, or double-counting.
        if footer_total > 0 and f_total > 0 and footer_total >= f_total:
            prep_ded = round(max(0.0, footer_total - f_total), 2)
        else:
            prep_ded = 0.0
            if header.get("prep_deduction", 0.0) > 0:
                print(
                    f"  ⚠️ [{inv_id}] Footer did not form a valid deduction closure, "
                    f"ignoring suspected deduction {header['prep_deduction']:,.2f}"
                )

        sum_qty = 0.0
        sum_net = 0.0
        sum_tax = 0.0

        for itm in items:
            net = round(itm["net"], 2)
            rate = itm["tax_rate"]
            tax = round(net * rate / 100, 2)
            qty = round(itm["qty"], 2)
            price = round(itm["price"], 2)

            row = {
                "Invoice_ID": inv_id,
                "Seller_PIB": "",
                "Buyer_PIB": "",
                "Datum_prometa": "",
                "Description": itm["desc"],
                "Quantity": qty,
                "Jedinica_Mere": itm["unit"],
                "Price": price,
                "Net_Amount": net,
                "PDV_Rate": f"{rate:g}%",
                "VAT_Amount": tax,
                "Total_Amount": round(net + tax, 2),
                "Prepayment_Deduction": 0.0,
                "Actual_Payment": round(net + tax, 2),
                "Status": "",
            }

            # credit note -> convert to negative
            desc_lower = str(itm["desc"]).lower()
            if any(kw in desc_lower for kw in CREDIT_NOTE_KW):
                print(f"🔴 Credit note detected: {inv_id} | {itm['desc']} → converted to negative")
                row["Quantity"] = -abs(qty) if qty != 0 else -1
                row["Price"] = -abs(price) if price != 0 else 0
                row["Net_Amount"] = -abs(net)
                row["VAT_Amount"] = -abs(tax)
                row["Total_Amount"] = -abs(net + tax)

            sum_qty += row["Quantity"]
            sum_net += row["Net_Amount"]
            sum_tax += row["VAT_Amount"]
            rows.append(row)

        audit_net = round(sum_net, 2)
        audit_tax = round(sum_tax, 2)
        audit_total = round(audit_net + audit_tax, 2)
        expected = round(audit_total - prep_ded, 2)

        # reconciliation breakdown, to locate the source of any difference (net/vat/deduction/official)
        print(
            f"  ├ Net={audit_net:,.2f} VAT={audit_tax:,.2f} Gross={audit_total:,.2f} "
            f"Prepay={prep_ded:,.2f} Due={expected:,.2f} Official={f_total:,.2f}"
        )

        if f_total and f_total > 0:
            diff = round(expected - f_total, 2)
            if diff == 0:
                st, tag = "OK", "✅ Reconciliation OK"
            elif abs(diff) <= 0.01:
                st, tag = "REVIEW", "⚠️ 1-cent difference, review"
            else:
                st, tag = "ERR", "❌ Mismatch"
            print(
                f"{tag} [{inv_id}]: actual due={expected:,.2f} "
                f"vs official={f_total:,.2f} (diff={diff:,.2f})"
            )
        else:
            st, tag = "NO_OFFICIAL", "⚠️ No official amount found"
            f_total = expected
            print(f"⚠️ No official amount [{inv_id}], using item total {audit_total:,.2f}")

        for r in rows[len(rows) - len(items):]:
            r["Status"] = st

        # invoice summary row (Total_Amount=pre-deduction gross, Actual_Payment=post-deduction due)
        rows.append({
            "Invoice_ID": inv_id,
            "Seller_PIB": seller_pib,
            "Buyer_PIB": buyer_pib,
            "Datum_prometa": datum,
            "Description": "--- summary ---",
            "Quantity": sum_qty,
            "Jedinica_Mere": "",
            "Price": "",
            "Net_Amount": audit_net,
            "PDV_Rate": "",
            "VAT_Amount": audit_tax,
            "Total_Amount": audit_total,
            "Prepayment_Deduction": prep_ded,
            "Actual_Payment": f_total,
            "Status": st,
        })

        g["qty"] += sum_qty
        g["net"] += audit_net
        g["tax"] += audit_tax
        g["prep"] += prep_ded
        g["actual"] += f_total

    # grand total row
    rows.append({
        "Invoice_ID": f"TOTAL ({len(invoices)})",
        "Seller_PIB": "",
        "Buyer_PIB": "",
        "Datum_prometa": "",
        "Description": "=== TOTAL ===",
        "Quantity": g["qty"],
        "Jedinica_Mere": "",
        "Price": "",
        "Net_Amount": g["net"],
        "PDV_Rate": "",
        "VAT_Amount": g["tax"],
        "Total_Amount": round(g["net"] + g["tax"], 2),
        "Prepayment_Deduction": g["prep"],
        "Actual_Payment": g["actual"],
        "Status": "done",
    })

    return rows


# ===================== Console table (racun.py style) =====================

def fmt_money(v):
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return ""


def fmt_qty(v):
    try:
        return f"{float(v):,.2f}"
    except Exception:
        return ""


def clip_text(text, width):
    text = safe_str(text)
    if len(text) <= width:
        return text
    return text[:width - 3] + "..."


def auto_width(series, default=10, max_width=40):
    if series.empty:
        return default
    longest = max(len(str(x)) for x in series.fillna(""))
    return min(max(longest + 2, default), max_width)


def print_console_table(df):
    invoice_w = auto_width(df["Invoice_ID"], default=12, max_width=18)
    desc_w = auto_width(df["Description"], default=26, max_width=52)
    seller_pib_w = 12
    buyer_pib_w = 12
    datum_w = auto_width(df["Datum_prometa"], default=14, max_width=16)
    qty_w = 12
    unit_w = 10
    price_w = 15
    money_w = 15
    status_w = 8

    total_width = invoice_w + desc_w + seller_pib_w + buyer_pib_w + datum_w + qty_w + unit_w + price_w + money_w * 5 + status_w + 16

    print("\n" + "=" * total_width)
    header = (
        f"{'Invoice_ID':<{invoice_w}}"
        f"{'Description':<{desc_w}}"
        f"{'Seller PIB':<{seller_pib_w}}"
        f"{'Buyer PIB':<{buyer_pib_w}}"
        f"{'DATUM PROMETA':<{datum_w}}"
        f"{'Qty':>{qty_w}}"
        f"{'Jed':>{unit_w}}"
        f"{'Price':>{price_w}}"
        f"{'Net':>{money_w}}"
        f"{'VAT':>{money_w}}"
        f"{'Total':>{money_w}}"
        f"{'Prepay':>{money_w}}"
        f"{'Actual':>{money_w}}"
        f"{'Status':>{status_w}}"
    )
    print(header)
    print("-" * total_width)

    for _, row in df.iterrows():
        if str(row.get("Invoice_ID", "")).startswith("TOTAL"):
            print("-" * total_width)

        invoice = clip_text(row.get("Invoice_ID", ""), invoice_w - 1)
        desc = clip_text(row.get("Description", ""), desc_w - 1)
        seller_pib = clip_text(row.get("Seller_PIB", ""), seller_pib_w - 1)
        buyer_pib = clip_text(row.get("Buyer_PIB", ""), buyer_pib_w - 1)
        datum = clip_text(row.get("Datum_prometa", ""), datum_w - 1)
        qty = fmt_qty(row.get("Quantity", ""))
        unit = clip_text(row.get("Jedinica_Mere", ""), unit_w - 1)

        if str(row.get("Invoice_ID", "")).startswith("TOTAL"):
            price = ""
        else:
            price = fmt_money(row.get("Price", ""))

        net = fmt_money(row.get("Net_Amount", ""))
        vat = fmt_money(row.get("VAT_Amount", ""))
        total = fmt_money(row.get("Total_Amount", ""))
        prep = fmt_money(row.get("Prepayment_Deduction", ""))
        actual = fmt_money(row.get("Actual_Payment", ""))
        status = clip_text(row.get("Status", ""), status_w - 1)

        print(
            f"{invoice:<{invoice_w}}"
            f"{desc:<{desc_w}}"
            f"{seller_pib:<{seller_pib_w}}"
            f"{buyer_pib:<{buyer_pib_w}}"
            f"{datum:<{datum_w}}"
            f"{qty:>{qty_w}}"
            f"{unit:>{unit_w}}"
            f"{price:>{price_w}}"
            f"{net:>{money_w}}"
            f"{vat:>{money_w}}"
            f"{total:>{money_w}}"
            f"{prep:>{money_w}}"
            f"{actual:>{money_w}}"
            f"{status:>{status_w}}"
        )

    print("=" * total_width)


# ===================== Main =====================

def main():
    task_started_at = time.time()

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    files = filedialog.askopenfilenames(
        title="Select native PDF invoice(s)",
        filetypes=[("PDF", "*.pdf")]
    )
    root.destroy()

    if not files:
        print("No file selected")
        return

    all_invoices = []
    for file_path in files:
        print(f"\n📂 {os.path.basename(file_path)}")
        invoices = process_pdf(file_path)
        if invoices:
            print(f"  ✅ Recognized {len(invoices)} invoice(s)")
            all_invoices.extend(invoices)
        else:
            print(f"  ⚠️ No invoices recognized")

    if not all_invoices:
        print(f"\n⏱️ Total task time: {time.time() - task_started_at:.1f} s")
        return

    print(f"\n📊 {len(files)} file(s), {len(all_invoices)} invoice(s) recognized")
    rows = build_report(all_invoices)

    df = pd.DataFrame(rows)

    cols = [
        "Invoice_ID",
        "Seller_PIB",
        "Buyer_PIB",
        "Datum_prometa",
        "Description",
        "Quantity",
        "Jedinica_Mere",
        "Price",
        "Net_Amount",
        "PDV_Rate",
        "VAT_Amount",
        "Total_Amount",
        "Prepayment_Deduction",
        "Actual_Payment",
        "Status",
    ]
    df = df[[c for c in cols if c in df.columns]]

    out_file = EXCEL_FILE
    try:
        df.to_excel(out_file, index=False)
    except PermissionError:
        out_file = "invoice_summary_FIX_" + time.strftime("%H%M%S") + ".xlsx"
        df.to_excel(out_file, index=False)
        print("\n⚠️ File in use, saved to backup file")

    print("\n✅ Done:", os.path.abspath(out_file))
    print_console_table(df.fillna(""))

    # if any invoice is not in OK status, print a footer data comparison table
    bad_ids = set()
    for _, row in df.iterrows():
        st = str(row.get("Status", ""))
        if st and st not in ("OK", "done", ""):
            inv_id = str(row.get("Invoice_ID", ""))
            if inv_id and not inv_id.startswith("TOTAL"):
                bad_ids.add(inv_id)

    if bad_ids:
        # build {invoice_id: header} map
        header_map = {}
        for inv in all_invoices:
            hid = str(inv["header"]["invoice_id"])
            header_map[hid] = inv["header"]

        print("\n\n" + "=" * 110)
        print("📋 Footer data comparison table (reconciliation-difference invoices only, all from raw PDF footer)")
        print("=" * 110)
        fhdr = f"{'Broj Fakture':<18}{'Net':>15}{'VAT':>15}{'Total':>15}{'Prepay':>15}{'Actual':>15}"
        print(fhdr)
        print("-" * 110)

        s_net = s_vat = s_total = s_prep = s_actual = 0.0
        for hid in sorted(bad_ids):
            hdr = header_map.get(hid, {})
            f_net = hdr.get("footer_net", 0.0) or 0.0
            f_vat = hdr.get("footer_vat", 0.0) or 0.0
            f_total = hdr.get("footer_total", 0.0) or 0.0
            f_prep = hdr.get("prep_deduction", 0.0) or 0.0
            f_actual = hdr.get("grand_total", 0.0) or 0.0
            s_net += f_net
            s_vat += f_vat
            s_total += f_total
            s_prep += f_prep
            s_actual += f_actual
            print(
                f"{hid:<18}"
                f"{f_net:>15,.2f}"
                f"{f_vat:>15,.2f}"
                f"{f_total:>15,.2f}"
                f"{f_prep:>15,.2f}"
                f"{f_actual:>15,.2f}"
            )
        print("-" * 110)
        print(
            f"{'SUM':<18}"
            f"{s_net:>15,.2f}"
            f"{s_vat:>15,.2f}"
            f"{s_total:>15,.2f}"
            f"{s_prep:>15,.2f}"
            f"{s_actual:>15,.2f}"
        )
        print("=" * 110)
        print("(data above comes from the raw PDF footer and is not used in the item-level calculation)")

    elapsed = time.time() - task_started_at
    print(f"⏱️ Total task time: {elapsed:.1f} s (~{elapsed / 60:.2f} min)")


if __name__ == "__main__":
    main()
