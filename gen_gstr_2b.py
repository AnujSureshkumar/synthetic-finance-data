"""
gen_gstr_2b.py - GSTR-2B (auto-drafted ITC statement) + a books purchase
register, built as a matched pair for the GSTR-2B reconciler (Project 5).

The reconciler's whole job is to compare what the GST portal shows (GSTR-2B,
auto-populated from suppliers' GSTR-1 filings) against what the company has
booked in its own purchase register, and to classify every line:

    MATCHED            same invoice, same values in both
    VALUE_MISMATCH     same invoice, but taxable / tax differs
    GSTIN_MISMATCH     same invoice booked against a wrong/typo'd GSTIN
    ONLY_IN_2B         supplier filed it; we have not booked it
    ONLY_IN_BOOKS      we booked it; supplier has not filed (or wrong period)

So this generator deliberately *manufactures* each of those situations from a
common base set, and also writes the ground-truth classification so the
reconciler can be scored.

Outputs (in ./output):
    gstr2b_062026.json            CBIC-style GSTR-2B JSON (return period 06-2026)
    purchase_register_books.csv / .xlsx   the company's books
    gstr2b_recon_truth.csv / .xlsx        expected classification per invoice

Return period: June 2026 (062026), FY 2026-27.
Synthetic only; all GSTINs pass the checksum but are registered to no one.
"""

from __future__ import annotations

import json

import pandas as pd

from common import (STATE_CODES, ensure_output_dir, get_rng, random_gstin,
                    random_pan, validate_gstin)
from excel_style import write_branded_excel

RETURN_PERIOD = "062026"        # MMYYYY (June 2026)
PERIOD_MONTH = "2026-06"
BUYER_STATE = "29"              # Karnataka
N_BASE = 55                     # invoices in the common base set

VENDOR_POOL = [
    ("Apex Cloud Services Pvt Ltd", 5101, 18),
    ("Zenith Software Labs Pvt Ltd", 5102, 18),
    ("Nimbus Tech Solutions Pvt Ltd", 5100, 18),
    ("Quantum Consulting LLP", 6100, 18),
    ("Vertex Facility Management Pvt Ltd", 6005, 18),
    ("Pioneer Telecom Pvt Ltd", 6009, 18),
    ("Stellar Stationers & Supplies", 6006, 18),
    ("Orbit Engineering Services", 6004, 18),
    ("Catalyst Media & Communications Pvt Ltd", 6205, 18),
    ("Meridian Travels Pvt Ltd", 6200, 12),
    ("Sterling Estates & Holdings Pvt Ltd", 6000, 18),
    ("Lumina & Associates, Chartered Accountants", 6101, 18),
]


def build_vendors(rng):
    """Stable vendor list with valid GSTINs; ~half local (Karnataka)."""
    others = [s for s in STATE_CODES if s != BUYER_STATE]
    vendors = []
    for i, (name, ecode, rate) in enumerate(VENDOR_POOL):
        state = BUYER_STATE if i % 2 == 0 else rng.choice(others)
        pan = random_pan(rng, entity="C")
        vendors.append({
            "name": name, "gstin": random_gstin(rng, state, pan),
            "state": state, "expense_code": ecode, "rate": rate,
            "intra": state == BUYER_STATE,
        })
    return vendors


def make_base(vendors, rng):
    """One coherent invoice line, before we split into books vs 2B."""
    base = []
    for n in range(N_BASE):
        v = rng.choice(vendors)
        rate = v["rate"]
        taxable = int(round(rng.uniform(15_000, 800_000) / 100) * 100)
        if v["intra"]:
            camt = round(taxable * rate / 200)
            samt = round(taxable * rate / 200)
            iamt = 0
        else:
            camt = samt = 0
            iamt = round(taxable * rate / 100)
        total = taxable + camt + samt + iamt
        day = rng.randint(1, 28)
        base.append({
            "inum": f"{v['name'][:3].upper()}/26-27/{1000 + n}",
            "dt": f"{day:02d}-06-2026",
            "ctin": v["gstin"],
            "trdnm": v["name"],
            "pos": v["state"],
            "rate": rate,
            "txval": taxable,
            "camt": camt, "samt": samt, "iamt": iamt, "csamt": 0,
            "val": total,
            "expense_code": v["expense_code"],
        })
    return base


def split_with_discrepancies(base, rng):
    """
    Turn the base set into (books_rows, b2b_rows, truth_rows) with a realistic
    mix of reconciliation outcomes.
    """
    books, b2b, truth = [], [], []

    # Decide a status for each base invoice.
    # Weighting: mostly matched, with a meaningful tail of exceptions.
    statuses = (["MATCHED"] * 34 + ["VALUE_MISMATCH"] * 6 + ["GSTIN_MISMATCH"] * 4
                + ["ONLY_IN_2B"] * 6 + ["ONLY_IN_BOOKS"] * 5)
    rng.shuffle(statuses)
    # pad/truncate to len(base)
    while len(statuses) < len(base):
        statuses.append("MATCHED")
    statuses = statuses[:len(base)]

    for inv, status in zip(base, statuses):
        b_row = {
            "invoice_no": inv["inum"],
            "invoice_date": inv["dt"],
            "vendor_gstin": inv["ctin"],
            "vendor_name": inv["trdnm"],
            "expense_code": inv["expense_code"],
            "taxable_value": inv["txval"],
            "cgst": inv["camt"], "sgst": inv["samt"], "igst": inv["iamt"],
            "invoice_total": inv["val"],
        }
        portal = dict(inv)  # what the 2B shows

        if status == "MATCHED":
            books.append(b_row)
            b2b.append(portal)
        elif status == "VALUE_MISMATCH":
            # books has a slightly different (lower) taxable - data-entry error
            factor = rng.choice([0.9, 0.95, 1.08])
            nt = int(round(b_row["taxable_value"] * factor / 100) * 100)
            ratio = nt / b_row["taxable_value"]
            b_row["taxable_value"] = nt
            b_row["cgst"] = round(b_row["cgst"] * ratio)
            b_row["sgst"] = round(b_row["sgst"] * ratio)
            b_row["igst"] = round(b_row["igst"] * ratio)
            b_row["invoice_total"] = (b_row["taxable_value"] + b_row["cgst"]
                                      + b_row["sgst"] + b_row["igst"])
            books.append(b_row)
            b2b.append(portal)
        elif status == "GSTIN_MISMATCH":
            # books booked against a wrong GSTIN (different state digit/typo)
            wrong = random_gstin(rng)  # unrelated valid GSTIN
            b_row["vendor_gstin"] = wrong
            books.append(b_row)
            b2b.append(portal)
        elif status == "ONLY_IN_2B":
            # supplier filed; not booked -> appears only on the portal
            b2b.append(portal)
        elif status == "ONLY_IN_BOOKS":
            # booked; supplier hasn't filed (or filed wrong period)
            books.append(b_row)

        truth.append({
            "invoice_no": inv["inum"],
            "vendor_name": inv["trdnm"],
            "expected_status": status,
            "taxable_2b": inv["txval"],
            "taxable_books": b_row["taxable_value"] if status != "ONLY_IN_2B" else None,
        })

    return books, b2b, truth


def to_gstr2b_json(b2b_rows, buyer_gstin):
    """Assemble a CBIC-style GSTR-2B JSON from the portal-visible rows."""
    by_vendor = {}
    for r in b2b_rows:
        by_vendor.setdefault((r["ctin"], r["trdnm"]), []).append(r)

    b2b = []
    for (ctin, trdnm), invs in by_vendor.items():
        inv_list = []
        for r in invs:
            inv_list.append({
                "inum": r["inum"],
                "dt": r["dt"],
                "val": r["val"],
                "pos": r["pos"],
                "rev": "N",
                "itcavl": "Y",
                "rsn": "",
                "items": [{
                    "rt": r["rate"],
                    "txval": r["txval"],
                    "igst": r["iamt"],
                    "cgst": r["camt"],
                    "sgst": r["samt"],
                    "cess": r["csamt"],
                }],
            })
        b2b.append({"ctin": ctin, "trdnm": trdnm,
                    "supprd": RETURN_PERIOD, "inv": inv_list})

    return {
        "chksum": "SYNTHETIC-DEMO-DATA",
        "data": {
            "gstin": buyer_gstin,
            "rtnprd": RETURN_PERIOD,
            "version": "GSTR2B-1.0",
            "gendt": "12-07-2026",
            "itcsumm": {
                "itcavl": {
                    "igst": sum(r["iamt"] for r in b2b_rows),
                    "cgst": sum(r["camt"] for r in b2b_rows),
                    "sgst": sum(r["samt"] for r in b2b_rows),
                    "cess": 0,
                },
            },
            "docdata": {"b2b": b2b},
        },
    }


def main() -> None:
    rng = get_rng()
    out = ensure_output_dir()
    buyer_gstin = random_gstin(rng, state_code=BUYER_STATE)

    vendors = build_vendors(rng)
    base = make_base(vendors, rng)
    books, b2b_rows, truth = split_with_discrepancies(base, rng)

    books_df = pd.DataFrame(books).sort_values("invoice_no").reset_index(drop=True)
    truth_df = pd.DataFrame(truth).sort_values("invoice_no").reset_index(drop=True)
    gstr2b = to_gstr2b_json(b2b_rows, buyer_gstin)

    # write JSON
    json_path = out / f"gstr2b_{RETURN_PERIOD}.json"
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(gstr2b, f, indent=2)

    # write books + truth (csv + styled xlsx)
    books_num = ["expense_code", "taxable_value", "cgst", "sgst", "igst",
                 "invoice_total"]
    books_df.to_csv(out / "purchase_register_books.csv", index=False)
    write_branded_excel(books_df, out / "purchase_register_books.xlsx",
                        sheet_name="Books", numeric_cols=books_num,
                        title="Purchase register (books) - Jun-2026 (synthetic)")

    truth_df.to_csv(out / "gstr2b_recon_truth.csv", index=False)
    write_branded_excel(truth_df, out / "gstr2b_recon_truth.xlsx",
                        sheet_name="Recon Truth",
                        numeric_cols=["taxable_2b", "taxable_books"],
                        title="GSTR-2B recon - expected classification (synthetic)")

    # --- integrity / summary ---
    assert validate_gstin(buyer_gstin)
    n_2b_inv = sum(len(v["inv"]) for v in gstr2b["data"]["docdata"]["b2b"])
    print(f"Buyer GSTIN: {buyer_gstin} (valid: {validate_gstin(buyer_gstin)})")
    print(f"Base invoices: {len(base)}")
    print(f"In books: {len(books_df)}   On portal (2B): {n_2b_inv}")
    print("Expected reconciliation mix:")
    print(truth_df["expected_status"].value_counts().to_string())
    print(f"Wrote {json_path.name}, purchase_register_books, gstr2b_recon_truth")


if __name__ == "__main__":
    main()
