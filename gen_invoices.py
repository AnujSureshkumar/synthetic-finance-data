"""
gen_invoices.py - synthetic vendor invoices for the Invoice -> Journal Entry bot.

Project 2 (Invoice -> JE bot) needs two things:
  1. A pile of realistic *purchase* invoices as PDFs - the bot's raw input. It
     reads each PDF, extracts the fields, picks the right expense ledger and GST
     ledgers, works out TDS, and proposes a balanced journal entry.
  2. A structured "ground truth" register of exactly what each PDF contains, so
     the bot's extraction and JE proposal can be scored against the right answer.

What this generator produces (in ./output):
    vendor_master.csv / .xlsx          15 vendors with valid-format GSTINs/PANs
    purchase_register.csv / .xlsx      ~60 invoices, header-level ground truth
    invoices_pdf/INV-*.pdf             one PDF per invoice (a readable tax invoice)

Modelling choices that keep it coherent:
  - The buyer is our fictional ITeS company in Karnataka (state code 29).
    Intra-state purchases (vendor also in 29) attract CGST+SGST; inter-state
    vendors attract IGST. Place of supply drives the split.
  - Each vendor sells one expense category, which fixes the expense ledger
    (from chart_of_accounts.csv), the GST rate, and the TDS section.
  - TDS is deducted on the taxable value (not on GST), per Indian practice.
  - A few deliberate edge cases are sprinkled in so the bot has something to
    handle: RCM (reverse charge, no GST charged by vendor), an MSME vendor
    (payment-terms flag), and a couple of round-sum invoices.

Synthetic only. GSTINs pass the real checksum but are registered to no one.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

from common import (STATE_CODES, ensure_output_dir, full_name, get_rng,
                    random_gstin, random_pan)
from excel_style import write_branded_excel

# --------------------------------------------------------------------------- #
# The buyer (our company) and the invoice window
# --------------------------------------------------------------------------- #

COMPANY = {
    "name": "Indus Novate Technologies Pvt Ltd",
    "gstin": None,          # filled with a Karnataka GSTIN at runtime
    "state_code": "29",     # Karnataka (Bengaluru)
    "address": "4th Floor, Prestige Tech Park, Outer Ring Road, Bengaluru 560103",
}

N_INVOICES = 60
N_PDFS = 12                 # how many of the invoices also get a PDF
INVOICE_MONTHS = ["2026-04", "2026-05", "2026-06"]  # Q1 FY26-27

# Vendor categories: (label, expense_code, expense_name, gst_rate, tds_section)
# expense_code/name match chart_of_accounts.csv exactly.
VENDOR_CATEGORIES = [
    ("Cloud hosting",        5101, "Cloud Hosting (AWS/Azure/GCP)",        18, "194J"),
    ("SaaS tools",           5102, "Software Subscriptions (SaaS Tools)",  18, "194J"),
    ("Subcontracting",       5100, "Subcontracting Charges",               18, "194C"),
    ("Legal & professional", 6100, "Legal & Professional Fees",            18, "194J"),
    ("Statutory audit",      6101, "Statutory Audit Fees",                 18, "194J"),
    ("Office rent",          6000, "Rent - Office Premises",               18, "194I"),
    ("Housekeeping/security", 6005, "Housekeeping & Security",             18, "194C"),
    ("Telephone & internet", 6009, "Telephone & Internet",                18, None),
    ("Office supplies",      6006, "Office Supplies & Stationery",         18, None),
    ("Repairs - equipment",  6004, "Repairs & Maintenance - Equipment",    18, "194C"),
    ("Advertising",          6205, "Advertising & Marketing",              18, "194C"),
    ("Travel & hotel",       6200, "Travelling - Domestic",                12, None),
]

# Vendor name building blocks (so names read like real Indian B2B suppliers).
VENDOR_PREFIX = ["Apex", "Zenith", "Nimbus", "Quantum", "Vertex", "Pioneer",
                 "Stellar", "Orbit", "Catalyst", "Meridian", "Sterling", "Lumina",
                 "Trinity", "Pinnacle", "Vega"]
VENDOR_SUFFIX = {
    "Cloud hosting": "Cloud Services Pvt Ltd",
    "SaaS tools": "Software Labs Pvt Ltd",
    "Subcontracting": "Tech Solutions Pvt Ltd",
    "Legal & professional": "Consulting LLP",
    "Statutory audit": "& Associates, Chartered Accountants",
    "Office rent": "Estates & Holdings Pvt Ltd",
    "Housekeeping/security": "Facility Management Pvt Ltd",
    "Telephone & internet": "Telecom Pvt Ltd",
    "Office supplies": "Stationers & Supplies",
    "Repairs - equipment": "Engineering Services",
    "Advertising": "Media & Communications Pvt Ltd",
    "Travel & hotel": "Travels Pvt Ltd",
}

# Per-category invoice value bands (taxable amount, INR).
VALUE_BANDS = {
    "Cloud hosting": (180_000, 900_000),
    "SaaS tools": (40_000, 350_000),
    "Subcontracting": (150_000, 1_200_000),
    "Legal & professional": (25_000, 300_000),
    "Statutory audit": (150_000, 450_000),
    "Office rent": (350_000, 350_000),     # fixed monthly rent
    "Housekeeping/security": (60_000, 180_000),
    "Telephone & internet": (15_000, 90_000),
    "Office supplies": (8_000, 60_000),
    "Repairs - equipment": (12_000, 120_000),
    "Advertising": (50_000, 600_000),
    "Travel & hotel": (20_000, 200_000),
}


# --------------------------------------------------------------------------- #
# Vendor master
# --------------------------------------------------------------------------- #

def build_vendors(rng) -> pd.DataFrame:
    rows = []
    other_states = [s for s in STATE_CODES if s != COMPANY["state_code"]]
    used_prefix = set()
    for i, (label, ecode, ename, rate, tds) in enumerate(VENDOR_CATEGORIES):
        # roughly half the vendors are local (Karnataka), half out-of-state
        if label in ("Office rent", "Housekeeping/security", "Office supplies"):
            state = COMPANY["state_code"]            # local services
        elif label in ("Cloud hosting", "SaaS tools", "Advertising"):
            state = rng.choice(other_states)          # typically out-of-state
        else:
            state = rng.choice([COMPANY["state_code"]] + other_states)

        prefix = rng.choice([p for p in VENDOR_PREFIX if p not in used_prefix])
        used_prefix.add(prefix)
        name = f"{prefix} {VENDOR_SUFFIX[label]}"

        pan = random_pan(rng, entity="C")
        gstin = random_gstin(rng, state_code=state, pan=pan)
        msme = rng.random() < 0.35
        rows.append({
            "vendor_id": f"VEND{100 + i}",
            "vendor_name": name,
            "category": label,
            "gstin": gstin,
            "pan": pan,
            "state_code": state,
            "state": STATE_CODES[state],
            "is_intra_state": state == COMPANY["state_code"],
            "expense_code": ecode,
            "expense_account": ename,
            "gst_rate": rate,
            "tds_section": tds or "",
            "msme_registered": msme,
            "payment_terms_days": 45 if msme else rng.choice([30, 45, 60]),
            "contact": full_name(rng),
        })
    return pd.DataFrame(rows)


# --------------------------------------------------------------------------- #
# Invoice register (ground truth)
# --------------------------------------------------------------------------- #

# TDS rates by section (the common rate for each).
TDS_RATE = {"194C": 2.0, "194J": 10.0, "194I": 10.0, "": 0.0}


def build_invoices(vendors: pd.DataFrame, rng) -> pd.DataFrame:
    rows = []
    seq = 0
    # Guarantee a handful of reverse-charge invoices (legal/professional under
    # RCM) so the bot always has the edge case to handle - inject at fixed slots.
    rcm_slots = set(rng.sample(range(N_INVOICES), 3))
    legal_vendor = vendors[vendors["category"] == "Legal & professional"].iloc[0]
    for n in range(N_INVOICES):
        force_rcm = n in rcm_slots
        if force_rcm:
            v = legal_vendor
        else:
            v = vendors.sample(1, random_state=rng.randint(0, 1_000_000)).iloc[0]
        cat = v["category"]
        lo, hi = VALUE_BANDS[cat]
        if lo == hi:
            taxable = lo
        else:
            taxable = int(round(rng.uniform(lo, hi) / 100) * 100)

        rate = int(v["gst_rate"])
        intra = bool(v["is_intra_state"])

        # Reverse charge: vendor charges no GST, buyer self-assesses.
        rcm = force_rcm

        if rcm:
            cgst = sgst = igst = 0
        elif intra:
            cgst = round(taxable * rate / 200)   # half each
            sgst = round(taxable * rate / 200)
            igst = 0
        else:
            cgst = sgst = 0
            igst = round(taxable * rate / 100)

        total = taxable + cgst + sgst + igst

        tds_sec = v["tds_section"]
        tds_rate = TDS_RATE.get(tds_sec, 0.0)
        # TDS only when taxable crosses a nominal threshold (simplified)
        tds_amt = round(taxable * tds_rate / 100) if taxable >= 30_000 else 0
        net_payable = total - tds_amt

        month = rng.choice(INVOICE_MONTHS)
        day = rng.randint(1, 28)
        inv_date = f"{month}-{day:02d}"
        seq += 1
        inv_no = f"INV-{month.replace('-', '')}-{seq:03d}"

        rows.append({
            "invoice_no": inv_no,
            "invoice_date": inv_date,
            "vendor_id": v["vendor_id"],
            "vendor_name": v["vendor_name"],
            "vendor_gstin": v["gstin"],
            "place_of_supply": v["state"],
            "supply_type": "Intra-state" if intra else "Inter-state",
            "reverse_charge": "Yes" if rcm else "No",
            "expense_code": int(v["expense_code"]),
            "expense_account": v["expense_account"],
            "description": f"{cat} - {month}",
            "taxable_value": taxable,
            "gst_rate": rate,
            "cgst": cgst,
            "sgst": sgst,
            "igst": igst,
            "invoice_total": total,
            "tds_section": tds_sec,
            "tds_rate": tds_rate,
            "tds_amount": tds_amt,
            "net_payable": net_payable,
            "msme": "Yes" if v["msme_registered"] else "No",
        })
    df = pd.DataFrame(rows).sort_values(["invoice_date", "invoice_no"]).reset_index(drop=True)
    return df


# --------------------------------------------------------------------------- #
# PDF rendering (one tax invoice per row)
# --------------------------------------------------------------------------- #

def render_pdfs(invoices: pd.DataFrame, vendors: pd.DataFrame, out_dir: Path,
                n: int) -> int:
    try:
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.units import mm
        from reportlab.pdfgen import canvas
    except ImportError:
        print("  (reportlab not installed - skipping PDF rendering; "
              "run: pip install reportlab --break-system-packages)")
        return 0

    from common import BRAND, inr

    pdf_dir = out_dir / "invoices_pdf"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    vmap = vendors.set_index("vendor_id").to_dict("index")
    sea = colors.HexColor("#" + BRAND["sea_green_deep"])
    ink = colors.HexColor("#" + BRAND["ink"])

    made = 0
    for _, r in invoices.head(n).iterrows():
        v = vmap[r["vendor_id"]]
        path = pdf_dir / f"{r['invoice_no']}.pdf"
        c = canvas.Canvas(str(path), pagesize=A4)
        w, h = A4
        y = h - 25 * mm

        # Vendor header
        c.setFillColor(sea)
        c.setFont("Helvetica-Bold", 16)
        c.drawString(20 * mm, y, r["vendor_name"])
        c.setFillColor(ink)
        c.setFont("Helvetica", 9)
        y -= 6 * mm
        c.drawString(20 * mm, y, f"GSTIN: {v['gstin']}   PAN: {v['pan']}   State: {v['state']} ({v['state_code']})")
        y -= 5 * mm
        c.drawString(20 * mm, y, f"Contact: {v['contact']}")

        # Title
        c.setFillColor(sea)
        c.setFont("Helvetica-Bold", 13)
        c.drawRightString(w - 20 * mm, h - 25 * mm, "TAX INVOICE")
        c.setFillColor(ink)
        c.setFont("Helvetica", 9)
        c.drawRightString(w - 20 * mm, h - 31 * mm, f"Invoice No: {r['invoice_no']}")
        c.drawRightString(w - 20 * mm, h - 36 * mm, f"Date: {r['invoice_date']}")
        if r["reverse_charge"] == "Yes":
            c.setFillColor(colors.red)
            c.drawRightString(w - 20 * mm, h - 41 * mm, "Reverse Charge: YES")
            c.setFillColor(ink)

        # Bill-to box
        y -= 12 * mm
        c.setStrokeColor(sea)
        c.setLineWidth(0.5)
        c.line(20 * mm, y, w - 20 * mm, y)
        y -= 6 * mm
        c.setFont("Helvetica-Bold", 9)
        c.drawString(20 * mm, y, "Bill To:")
        c.setFont("Helvetica", 9)
        y -= 5 * mm
        c.drawString(20 * mm, y, COMPANY["name"])
        y -= 5 * mm
        c.drawString(20 * mm, y, f"GSTIN: {COMPANY['gstin']}")
        y -= 5 * mm
        c.drawString(20 * mm, y, COMPANY["address"])

        # Line item table
        y -= 12 * mm
        c.setFillColor(sea)
        c.rect(20 * mm, y - 2 * mm, w - 40 * mm, 7 * mm, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(22 * mm, y, "Description")
        c.drawRightString(w - 60 * mm, y, "Rate %")
        c.drawRightString(w - 22 * mm, y, "Amount (INR)")
        c.setFillColor(ink)
        c.setFont("Helvetica", 9)
        y -= 9 * mm
        c.drawString(22 * mm, y, f"{r['expense_account']}  [{r['description']}]")
        c.drawRightString(w - 60 * mm, y, str(r["gst_rate"]))
        c.drawRightString(w - 22 * mm, y, inr(r["taxable_value"]))

        # Tax breakup
        def line(label, amount):
            nonlocal y
            y -= 6 * mm
            c.drawRightString(w - 60 * mm, y, label)
            c.drawRightString(w - 22 * mm, y, inr(amount))

        y -= 3 * mm
        c.line(20 * mm, y, w - 20 * mm, y)
        line("Taxable Value", r["taxable_value"])
        if r["cgst"]:
            line(f"CGST @ {r['gst_rate']/2:g}%", r["cgst"])
            line(f"SGST @ {r['gst_rate']/2:g}%", r["sgst"])
        if r["igst"]:
            line(f"IGST @ {r['gst_rate']:g}%", r["igst"])
        y -= 2 * mm
        c.line(w - 90 * mm, y, w - 20 * mm, y)
        c.setFont("Helvetica-Bold", 10)
        line("Invoice Total", r["invoice_total"])
        c.setFont("Helvetica", 9)
        if r["tds_amount"]:
            line(f"Less: TDS {r['tds_section']} @ {r['tds_rate']:g}%", -r["tds_amount"])
        c.setFont("Helvetica-Bold", 10)
        line("Net Payable", r["net_payable"])

        # Footer
        c.setFont("Helvetica-Oblique", 7)
        c.setFillColor(colors.HexColor("#" + BRAND["muted"]))
        c.drawString(20 * mm, 18 * mm,
                     "Synthetic invoice for demonstration only - not a real "
                     "commercial document. GSTIN format-valid, not registered.")
        c.showPage()
        c.save()
        made += 1
    return made


# --------------------------------------------------------------------------- #
# Main
# --------------------------------------------------------------------------- #

def main() -> None:
    rng = get_rng()
    out = ensure_output_dir()

    # give the company a valid Karnataka GSTIN
    COMPANY["gstin"] = random_gstin(rng, state_code=COMPANY["state_code"])

    vendors = build_vendors(rng)
    invoices = build_invoices(vendors, rng)

    vendor_num = ["expense_code", "gst_rate", "payment_terms_days"]
    inv_num = ["expense_code", "taxable_value", "gst_rate", "cgst", "sgst",
               "igst", "invoice_total", "tds_rate", "tds_amount", "net_payable"]

    vendors.to_csv(out / "vendor_master.csv", index=False)
    write_branded_excel(vendors, out / "vendor_master.xlsx",
                        sheet_name="Vendor Master", numeric_cols=vendor_num,
                        title="Vendor master - purchase ledger (synthetic)")

    invoices.to_csv(out / "purchase_register.csv", index=False)
    write_branded_excel(invoices, out / "purchase_register.xlsx",
                        sheet_name="Purchase Register", numeric_cols=inv_num,
                        title="Purchase register - invoice ground truth (synthetic)")

    n_pdf = render_pdfs(invoices, vendors, out, N_PDFS)

    # --- integrity checks ---
    from common import validate_gstin
    assert vendors["gstin"].map(validate_gstin).all(), "A vendor GSTIN fails checksum"
    assert validate_gstin(COMPANY["gstin"]), "Company GSTIN fails checksum"
    recon = (invoices["taxable_value"] + invoices["cgst"] + invoices["sgst"]
             + invoices["igst"] - invoices["invoice_total"]).abs().max()
    assert recon == 0, f"Invoice total does not reconcile (gap {recon})"

    print("Vendors: {}  Invoices: {}  PDFs: {}".format(
        len(vendors), len(invoices), n_pdf))
    print("Company GSTIN: {} (valid: {})".format(
        COMPANY["gstin"], validate_gstin(COMPANY["gstin"])))
    print("Supply mix:")
    print(invoices["supply_type"].value_counts().to_string())
    print("Reverse-charge invoices: {}".format(
        (invoices["reverse_charge"] == "Yes").sum()))
    print("Invoices with TDS: {}".format((invoices["tds_amount"] > 0).sum()))
    print("Total taxable value: Rs {:,}".format(invoices["taxable_value"].sum()))
    print("Wrote vendor_master, purchase_register (csv + xlsx) and "
          "{} PDFs in invoices_pdf/".format(n_pdf))


if __name__ == "__main__":
    main()
