#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fakturars.py — izdvajanje faktura iz nativnih PDF dokumenata po čistim pravilima (bez LLM)
===========================================================================================
Po uzoru na izlaz invoice.py / racun.py; koristi pdfplumber parsiranje iz match_invoices.py:
- Tkinter dijalog za izbor jednog nativnog (sa tekstualnim slojem) PDF fajla fakture
- Čisti regex + pravila tabelarnih redova: zaglavlje (broj fakture / PIB prodavca i kupca / Datum prometa / zvanični iznos) + stavke
- Razdvajanje po stranici: PDF sa više faktura (npr. *_sef.pdf batch) → jedna faktura po stranici
- Umanjenje po avansu: detektuje avans linije (Ukupan iznos osnovice umanjen za osnovicu po avansu itd.),
  izlazne kolone Prepayment_Deduction i Actual_Payment (semantika avans3.py)
- Bez poziva LLM, samo izdvajanje
- Izlaz: Zbir_Faktura.xlsx + PowerShell tabela (stil racun.py, bez kolone sa prevodom na kineski)

Upotreba:
    python fakturars.py
"""

import os
import re
import time
import unicodedata

import pandas as pd
import pdfplumber

from tkinter import filedialog, Tk

EXCEL_FILE = "Zbir_Faktura.xlsx"

# ===================== Pomocne funkcije =====================


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
    """Čišćenje brojeva: objedinjavanje raznih formata hiljada u float"""
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
    # zahteva: opis + kolicin + cena moraju biti prisutni;
    # jednokolonni kompatibilni skraćenici: jedinica mere / jed. mere / jed mere / j.m. / mere
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
    Parsira red stavke, podržava tri rasporeda kolona (prema tome da li zaglavlje ima kolonu popusta):
      sa popustom standardno: desc qty price unit discount net tax_rate        (brojčana stopa)
      sa popustom posebno: desc qty price unit discount net tax_status          (tekstualni status, npr. posebni postupci)
      bez popusta: desc qty price unit net tax_rate
    Prateći ne-brojčani tokeni se tretiraju kao tekstualni status poreza (posebni postupci oporezivanja).
    """
    tokens = [tok for tok in safe_str(line).split() if tok]
    if len(tokens) < 6:
        return None

    # skini prateće ne-brojčane tokene  → tekstualni status poreza (posebni postupci)
    i = len(tokens) - 1
    tax_status_parts = []
    while i >= 0 and not is_number_token(tokens[i]):
        tax_status_parts.insert(0, tokens[i])
        i -= 1
    tax_status = " ".join(tax_status_parts).strip()
    body = tokens[: i + 1]

    if has_discount:
        # raspored: qty price unit discount net [tax_rate|tax_status]
        # 6 polja (standardno, brojčana stopa)
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
        # 5 polja (posebni postupci, tekstualni status)
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

    # bez popusta: qty price unit net [tax_rate|tax_status]
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


# ===================== Izdvajanje umanjenja po avansu (semantika avans3.py) =====================

def extract_avans_deduction(page_text):
    """
    Izdvaja umanjenje po avansu sa eFaktura stranice,
    dinamički prolazi kroz sve stope PDV-a.

    Relevantne linije (jedna grupa po stopi):
      Ukupna osnovica - stopa 20%: 17.837,55                 (osnovica pre umanjenja)
      Ukupan iznos osnovice umanjen za osnovicu po avansu     (osnovica posle umanjenja)
      Ukupan PDV - stopa 20%: 3.567,51                        (PDV pre umanjenja)
      Ukupan PDV umanjen za PDV po avansu                     (PDV posle umanjenja)

    Umanjenje = Σ po stopi [(osnovica pre - osnovica posle) + (PDV pre - PDV posle)].
    Sistem uvek štampa ove linije; kada je avans nula, umanjenje je 0.
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


# ===================== Izdvajanje zaglavlja (PIB prodavca/kupca) =====================

def extract_header_fields(page_text):
    invoice_id = find_first_match([
        r"(?:broj\s*fakture|broj\s*racuna|broj\s*računa|faktura\s*br\.?|račun\s*br\.?|racun\s*br\.?)[:\s]*([A-Z0-9\-\/]+)",
        # eFaktura nastavne stranice često zadržavaju samo: Generisao sistem eFaktura pod brojem:00084/26
        r"generisao\s+sistem\s+eFaktura\s+pod\s+brojem[:\s]*([A-Z0-9\-\/]+)",
        r"\b(IF\d{2}-\d{3,})\b",
        r"\b([0-9]{2,4}(?:-[0-9]{2,}){1,4})\b",
    ], page_text)

    def validate_pib(val):
        s = safe_str(val)
        return s if re.fullmatch(r"[0-9]{9}", s) else ""

    # PIB kupca: PIB kupca; (?!\d) sprečava odsecanje prvih 9 cifara dužeg broja
    buyer_pib = find_first_match([
        r"\bPIB\s+kupca[:\s]*([0-9]{9})(?!\d)",
        r"\bPIB\s+customer[:\s]*([0-9]{9})(?!\d)",
    ], page_text)
    buyer_pib = validate_pib(buyer_pib)

    # PIB prodavca: PIB u Prodavac bloku; rezerva je PIB koji nije kupčev
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

    # zvanični iznos: prvo Iznos za plaćanje, rezerva Ukupan iznos fakture / Grand total
    # [^\r\n]*? drži poklapanje u okviru tekuće linije, sprečava gutanje broja iz sledeće linije
    grand_total = find_first_match([
        r"Iznos za plaćanje[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Iznos za placanje[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Ukupan iznos fakture[^\r\n]*?([0-9][0-9\., ]*)\s*$",
        r"Grand total[^\r\n]*?([0-9][0-9\., ]*)\s*$",
    ], page_text)

    # zvanični footer: neto / pdv / ukupan iznos (prikazuje se pri usaglašavanju razlike)
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


# ===================== Izdvajanje stavki (iz match_invoices.py) =====================

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
                # tekstualni status poreza (posebni postupci) može se protezati preko više linija
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


# ===================== Izdvajanje jednog PDF-a (više faktura po stranicama) =====================

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
                print(f"  ⏭️ Stranica {idx+1} preskočena (broj fakture nije prepoznat)")
                continue

            if not items:
                print(f"  ⚠️ Stranica {idx+1} {header['invoice_id']} — nisu izdvojene stavke (zaglavlje/redovi nisu odgovarali):")
                for ln in lines[:30]:
                    print(f"     | {ln}")

            inv_id = header["invoice_id"]

            # nastavci: ista faktura na narednoj stranici se spaja ili preskače
            if invoices and inv_id == invoices[-1]["header"]["invoice_id"]:
                if items:
                    print(f"  🔁 Stranica {idx+1} je nastavak fakture {inv_id}, dodato {len(items)} stavki")
                    invoices[-1]["items"].extend(items)
                else:
                    print(f"  🔁 Stranica {idx+1} je nastavak fakture {inv_id} (bez stavki), preskočeno")
                # nastavna stranica može ponovo nositi footer ili podatke o avansu
                for key in ("grand_total", "footer_total", "prep_deduction", "seller_pib", "buyer_pib", "datum_prometa"):
                    if header.get(key):
                        invoices[-1]["header"][key] = header[key]
                continue

            invoices.append({"header": header, "items": items})

    if not invoices:
        print(f"⚠️ Nema prepoznatih stranica faktura: {os.path.basename(file_path)}")
        return []

    return invoices


# ===================== Ključne reči kreditne note =====================

CREDIT_NOTE_KW = [
    "knjizno odobrenje", "knjižno odobrenje",
    "storno", "storniranje", "storno račun", "storno faktura",
    "credit note", "credit memo", "credit memorandum",
    "odobrenje", "umanjenje", "smanjenje",
]


# ===================== Izgradnja izveštaja (stil racun.py, bez kolone sa prevodom) =====================

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

        # umanjenje po avansu se izvodi iz footer ukupnog iznosa i konačnog duga.
        # tretira se tek kada oba postoje i kada je footer ukupni iznos veći od duga,
        # čime se izbegava tumačenje umanjenih osnovice/pdv-a kao avansa ili dvostruko umanjenje.
        if footer_total > 0 and f_total > 0 and footer_total >= f_total:
            prep_ded = round(max(0.0, footer_total - f_total), 2)
        else:
            prep_ded = 0.0
            if header.get("prep_deduction", 0.0) > 0:
                print(
                    f"  ⚠️ [{inv_id}] Footer ne formira validan krug umanjenja, "
                    f"ignoriše se sumnja za avans {header['prep_deduction']:,.2f}"
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

            # kreditna nota → pretvoriti u negativno
            desc_lower = str(itm["desc"]).lower()
            if any(kw in desc_lower for kw in CREDIT_NOTE_KW):
                print(f"🔴 Otkrivena kreditna nota: {inv_id} | {itm['desc']} → prevedeno u negativno")
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

        # prikaz obračuna radi lociranja izvora razlike (neto/pdv/avans/zvanični iznos)
        print(
            f"  ├ Neto={audit_net:,.2f} PDV={audit_tax:,.2f} Bruto={audit_total:,.2f} "
            f"Avans={prep_ded:,.2f} Dug={expected:,.2f} Zvanično={f_total:,.2f}"
        )

        if f_total and f_total > 0:
            diff = round(expected - f_total, 2)
            if diff == 0:
                st, tag = "OK", "✅ Usaglašeno"
            elif abs(diff) <= 0.01:
                st, tag = "REVIEW", "⚠️ Razlika od 1 cent, proveriti"
            else:
                st, tag = "ERR", "❌ Nije usaglašeno"
            print(
                f"{tag} [{inv_id}]: stvarni dug={expected:,.2f} "
                f"vs zvanični={f_total:,.2f} (razlika={diff:,.2f})"
            )
        else:
            st, tag = "NO_OFFICIAL", "⚠️ Zvanični iznos nije pronađen"
            f_total = expected
            print(f"⚠️ Zvanični iznos nije pronađen [{inv_id}], koristi se zbir stavki {audit_total:,.2f}")

        for r in rows[len(rows) - len(items):]:
            r["Status"] = st

        # rezime fakture (Total_Amount=bruto pre umanjenja, Actual_Payment=dug posle umanjenja)
        rows.append({
            "Invoice_ID": inv_id,
            "Seller_PIB": seller_pib,
            "Buyer_PIB": buyer_pib,
            "Datum_prometa": datum,
            "Description": "--- rezime ---",
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

    # red ukupnog zbira
    rows.append({
        "Invoice_ID": f"UKUPNO ({len(invoices)})",
        "Seller_PIB": "",
        "Buyer_PIB": "",
        "Datum_prometa": "",
        "Description": "=== UKUPNO ===",
        "Quantity": g["qty"],
        "Jedinica_Mere": "",
        "Price": "",
        "Net_Amount": g["net"],
        "PDV_Rate": "",
        "VAT_Amount": g["tax"],
        "Total_Amount": round(g["net"] + g["tax"], 2),
        "Prepayment_Deduction": g["prep"],
        "Actual_Payment": g["actual"],
        "Status": "završeno",
    })

    return rows


# ===================== Tabela u konzoli (stil racun.py) =====================

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
        f"{'Prodavac PIB':<{seller_pib_w}}"
        f"{'Kupac PIB':<{buyer_pib_w}}"
        f"{'DATUM PROMETA':<{datum_w}}"
        f"{'Količina':>{qty_w}}"
        f"{'Jed':>{unit_w}}"
        f"{'Cena':>{price_w}}"
        f"{'Neto':>{money_w}}"
        f"{'PDV':>{money_w}}"
        f"{'Ukupno':>{money_w}}"
        f"{'Avans':>{money_w}}"
        f"{'Plaćeno':>{money_w}}"
        f"{'Status':>{status_w}}"
    )
    print(header)
    print("-" * total_width)

    for _, row in df.iterrows():
        if str(row.get("Invoice_ID", "")).startswith(("TOTAL", "UKUPNO")):
            print("-" * total_width)

        invoice = clip_text(row.get("Invoice_ID", ""), invoice_w - 1)
        desc = clip_text(row.get("Description", ""), desc_w - 1)
        seller_pib = clip_text(row.get("Seller_PIB", ""), seller_pib_w - 1)
        buyer_pib = clip_text(row.get("Buyer_PIB", ""), buyer_pib_w - 1)
        datum = clip_text(row.get("Datum_prometa", ""), datum_w - 1)
        qty = fmt_qty(row.get("Quantity", ""))
        unit = clip_text(row.get("Jedinica_Mere", ""), unit_w - 1)

        if str(row.get("Invoice_ID", "")).startswith(("TOTAL", "UKUPNO")):
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


# ===================== Glavni program =====================

def main():
    task_started_at = time.time()

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    files = filedialog.askopenfilenames(
        title="Izaberite nativne PDF fajlove faktura (više fajlova)",
        filetypes=[("PDF", "*.pdf")]
    )
    root.destroy()

    if not files:
        print("Nije izabran nijedan fajl")
        return

    all_invoices = []
    for file_path in files:
        print(f"\n📂 {os.path.basename(file_path)}")
        invoices = process_pdf(file_path)
        if invoices:
            print(f"  ✅ Prepoznato {len(invoices)} faktura")
            all_invoices.extend(invoices)
        else:
            print(f"  ⚠️ Nije prepoznata nijedna faktura")

    if not all_invoices:
        print(f"\n⏱️ Ukupno vreme zadatka: {time.time() - task_started_at:.1f} s")
        return

    print(f"\n📊 {len(files)} fajlova, prepoznato {len(all_invoices)} faktura")
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
        out_file = "Zbir_Faktura_FIX_" + time.strftime("%H%M%S") + ".xlsx"
        df.to_excel(out_file, index=False)
        print("\n⚠️ Fajl je zauzet, sačuvano u rezervni fajl")

    print("\n✅ Završeno:", os.path.abspath(out_file))
    print_console_table(df.fillna(""))

    # ako postoje fakture bez OK statusa, ispiši tabelu poređenja footer podataka
    bad_ids = set()
    for _, row in df.iterrows():
        st = str(row.get("Status", ""))
        if st and st not in ("OK", "završeno", ""):
            inv_id = str(row.get("Invoice_ID", ""))
            if inv_id and not inv_id.startswith("TOTAL"):
                bad_ids.add(inv_id)

    if bad_ids:
        # mapiranje {broj_fakture: header}
        header_map = {}
        for inv in all_invoices:
            hid = str(inv["header"]["invoice_id"])
            header_map[hid] = inv["header"]

        print("\n\n" + "=" * 110)
        print("📋 Tabela poređenja footer podataka (samo fakture sa razlikom u usaglašavanju, svi podaci iz originalnog PDF footera)")
        print("=" * 110)
        fhdr = f"{'Broj Fakture':<18}{'Net':>15}{'PDV':>15}{'Total':>15}{'Avans':>15}{'Plaćeno':>15}"
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
        print("(iznadni podaci dolaze iz originalnog PDF footera i ne učestvuju u računanju stavki)")

    elapsed = time.time() - task_started_at
    print(f"⏱️ Ukupno vreme zadatka: {elapsed:.1f} s (~{elapsed / 60:.2f} min)")


if __name__ == "__main__":
    main()
