"""
gen_salary_slips.py — 50 synthetic employees of an Indian ITeS firm.

Produces two datasets:

1. salary_master   — one row per employee: full annual CTC structure plus the
   investment / exemption declarations a payroll team collects each April.
   This is the direct input to the **Tax Regime Optimizer** (Project 1): every
   field the old-vs-new comparison needs is here (Basic, HRA component, rent
   paid, 80C/80CCD(1B)/80D/80E/24(b), employer NPS under 80CCD(2), etc.).

2. salary_register — a single month's payslip register (earnings + statutory
   deductions + indicative TDS + net pay). Feeds the payroll journal entry in
   the Invoice->JE bot and the headcount cost lines in the BvA dashboard.

The FY26 new-regime slabs used for the register's indicative TDS are the same
ones the optimizer will use on its new-regime side, so the number is internally
consistent — but it is clearly indicative, not a payroll-grade computation.

CTC structure (employer practice modelled):
    basic           = 45% of CTC
    hra component   = 50% of basic
    employer PF     = 12% of basic, capped at Rs 1,800/month (Rs 21,600/yr)
    gratuity        = 4.81% of basic
    LTA             = 5% of basic
    employer NPS    = 10% of basic for ~1 in 4 staff (80CCD(2)), else nil
    special allow.  = balancing figure so components sum to CTC

Synthetic only. Names, PANs and figures are invented.

Outputs (in ./output):
    salary_master.csv / .xlsx
    salary_register_<month>.csv / .xlsx
"""

from __future__ import annotations

import pandas as pd

from common import (STATE_CODES, ensure_output_dir, full_name, get_rng,
                    random_pan, round_to)
from excel_style import write_branded_excel

N_EMPLOYEES = 50
PAY_MONTH = "2026-06"  # June 2026 payslip register

# Role bands: (title, min_ctc, max_ctc, weight)
ROLE_BANDS = [
    ("Associate", 400_000, 750_000, 22),
    ("Senior Associate", 700_000, 1_200_000, 18),
    ("Team Lead", 1_100_000, 1_800_000, 14),
    ("Project Manager", 1_600_000, 2_600_000, 12),
    ("Senior Manager", 2_400_000, 3_800_000, 8),
    ("Associate Director", 3_500_000, 5_500_000, 4),
    ("Director", 5_000_000, 8_000_000, 2),
]

METRO_CITIES = ["Mumbai", "Delhi", "Bengaluru", "Chennai", "Kolkata", "Hyderabad"]
NON_METRO_CITIES = ["Pune", "Coimbatore", "Mysuru", "Indore", "Kochi", "Jaipur",
                    "Nagpur", "Vadodara"]

# Professional tax differs slightly by state; Rs 2,400/yr is the common ceiling.
ANNUAL_PT = 2_400


# --------------------------------------------------------------------------- #
# FY26 (AY 2026-27) new-regime tax — used only for indicative monthly TDS
# --------------------------------------------------------------------------- #

def fy26_new_regime_tax(gross_salary_income: float) -> float:
    """
    Indicative annual tax under the FY26 new regime (default regime).
    Standard deduction Rs 75,000; 87A rebate makes taxable income up to
    Rs 12,00,000 tax-free; 4% health & education cess on top.
    Surcharge ignored (immaterial for this synthetic salary range; the
    optimizer handles surcharge properly).
    """
    taxable = max(0.0, gross_salary_income - 75_000)
    slabs = [
        (400_000, 0.00),
        (800_000, 0.05),
        (1_200_000, 0.10),
        (1_600_000, 0.15),
        (2_000_000, 0.20),
        (2_400_000, 0.25),
        (float("inf"), 0.30),
    ]
    tax = 0.0
    lower = 0.0
    for upper, rate in slabs:
        if taxable > lower:
            tax += (min(taxable, upper) - lower) * rate
            lower = upper
        else:
            break
    # 87A rebate: nil tax if taxable income <= 12,00,000
    if taxable <= 1_200_000:
        tax = 0.0
    return round(tax * 1.04)  # add 4% cess


# --------------------------------------------------------------------------- #
# Build one employee
# --------------------------------------------------------------------------- #

def pick_role(rng):
    titles = [r[0] for r in ROLE_BANDS]
    weights = [r[3] for r in ROLE_BANDS]
    title = rng.choices(titles, weights=weights, k=1)[0]
    band = next(r for r in ROLE_BANDS if r[0] == title)
    return title, band


def build_employee(rng, idx: int) -> dict:
    title, band = pick_role(rng)
    ctc = round_to(rng.uniform(band[1], band[2]), 1_000)

    metro = rng.random() < 0.6
    city = rng.choice(METRO_CITIES if metro else NON_METRO_CITIES)

    # Earnings structure
    basic = round_to(0.45 * ctc, 100)
    hra_component = round_to(0.50 * basic, 100)
    employer_pf = min(round_to(0.12 * basic, 12), 21_600)
    gratuity = round_to(0.0481 * basic, 1)
    lta = round_to(0.05 * basic, 100)
    employer_nps = round_to(0.10 * basic, 100) if rng.random() < 0.25 else 0
    special_allowance = (ctc - basic - hra_component - employer_pf
                         - gratuity - lta - employer_nps)
    # Guard: if rounding pushed special negative, absorb into special and rebuild
    if special_allowance < 0:
        special_allowance = 0

    # Employee statutory deductions
    employee_pf = employer_pf  # mirror the employer cap
    professional_tax = ANNUAL_PT

    # --- Investment / exemption declarations (old-regime relevance) ----------
    # ~30% are "non-declarers" (lean new regime), rest declare a realistic mix.
    declarer = rng.random() < 0.70

    owns_home = declarer and rng.random() < 0.30
    if owns_home:
        rent_paid_annual = 0
        home_loan_interest = round_to(rng.uniform(60_000, 200_000), 1_000)
    else:
        # monthly rent scales loosely with CTC; metros cost more
        base_rent = ctc / 12 * rng.uniform(0.12, 0.22) * (1.25 if metro else 1.0)
        rent_paid_annual = round_to(base_rent * 12, 1_000)
        home_loan_interest = 0

    if declarer:
        room_80c = max(0, 150_000 - employee_pf)
        decl_80c_other = round_to(rng.uniform(0, room_80c), 1_000)
        decl_80ccd_1b = rng.choice([0, 0, 25_000, 50_000])
        decl_80d_self = rng.choice([0, 15_000, 25_000])
        decl_80d_parents = rng.choice([0, 0, 25_000, 50_000])
        decl_80e = round_to(rng.uniform(0, 80_000), 1_000) if rng.random() < 0.15 else 0
    else:
        decl_80c_other = 0
        decl_80ccd_1b = 0
        decl_80d_self = 0
        decl_80d_parents = 0
        decl_80e = 0

    return {
        "emp_id": f"INF{1000 + idx}",
        "name": full_name(rng),
        "pan": random_pan(rng, entity="P"),
        "gender": rng.choice(["M", "F"]),
        "age": rng.randint(23, 58),
        "date_of_joining": f"{rng.randint(2015, 2025)}-{rng.randint(1, 12):02d}-01",
        "department": rng.choice(["Engineering", "Delivery", "QA", "Product",
                                  "Support", "Finance", "HR", "Sales"]),
        "designation": title,
        "location": city,
        "metro": metro,
        "gross_ctc": ctc,
        "basic": basic,
        "hra_component": hra_component,
        "special_allowance": special_allowance,
        "lta": lta,
        "employer_pf": employer_pf,
        "gratuity": gratuity,
        "employer_nps_80ccd2": employer_nps,
        "employee_pf_80c": employee_pf,
        "professional_tax": professional_tax,
        "rent_paid_annual": rent_paid_annual,
        "decl_80c_other": decl_80c_other,
        "decl_80ccd_1b_nps": decl_80ccd_1b,
        "decl_80d_self": decl_80d_self,
        "decl_80d_parents": decl_80d_parents,
        "decl_80e_edu_loan_int": decl_80e,
        "decl_24b_home_loan_int": home_loan_interest,
    }


# --------------------------------------------------------------------------- #
# Monthly payslip register derived from the master
# --------------------------------------------------------------------------- #

def build_register(master: pd.DataFrame, month: str) -> pd.DataFrame:
    rows = []
    for _, e in master.iterrows():
        gross_monthly = round(e["gross_ctc"] / 12)
        basic_m = round(e["basic"] / 12)
        hra_m = round(e["hra_component"] / 12)
        special_m = round(e["special_allowance"] / 12)
        lta_m = round(e["lta"] / 12)
        # earnings paid to employee exclude gratuity & employer contributions
        earnings = basic_m + hra_m + special_m + lta_m
        epf_m = round(e["employee_pf_80c"] / 12)
        pt_m = 200  # Rs 200/month
        # indicative TDS: annual new-regime tax / 12 on cash salary income
        annual_cash = e["basic"] + e["hra_component"] + e["special_allowance"] + e["lta"]
        tds_m = round(fy26_new_regime_tax(annual_cash) / 12)
        net_pay = earnings - epf_m - pt_m - tds_m
        rows.append({
            "pay_month": month,
            "emp_id": e["emp_id"],
            "name": e["name"],
            "designation": e["designation"],
            "department": e["department"],
            "basic": basic_m,
            "hra": hra_m,
            "special_allowance": special_m,
            "lta": lta_m,
            "gross_earnings": earnings,
            "employee_pf": epf_m,
            "professional_tax": pt_m,
            "tds": tds_m,
            "total_deductions": epf_m + pt_m + tds_m,
            "net_pay": net_pay,
            "employer_pf": round(e["employer_pf"] / 12),
            "ctc_monthly": gross_monthly,
        })
    return pd.DataFrame(rows)


def main() -> None:
    rng = get_rng()
    out = ensure_output_dir()

    master = pd.DataFrame(build_employee(rng, i) for i in range(N_EMPLOYEES))

    # integrity: components reconcile to CTC within a small rounding tolerance
    recon = (master["basic"] + master["hra_component"] + master["special_allowance"]
             + master["lta"] + master["employer_pf"] + master["gratuity"]
             + master["employer_nps_80ccd2"])
    max_gap = int((master["gross_ctc"] - recon).abs().max())
    assert max_gap <= 12, f"CTC components fail to reconcile (gap {max_gap})"

    register = build_register(master, PAY_MONTH)

    master_num = [
        "gross_ctc", "basic", "hra_component", "special_allowance", "lta",
        "employer_pf", "gratuity", "employer_nps_80ccd2", "employee_pf_80c",
        "professional_tax", "rent_paid_annual", "decl_80c_other",
        "decl_80ccd_1b_nps", "decl_80d_self", "decl_80d_parents",
        "decl_80e_edu_loan_int", "decl_24b_home_loan_int",
    ]
    reg_num = [
        "basic", "hra", "special_allowance", "lta", "gross_earnings",
        "employee_pf", "professional_tax", "tds", "total_deductions",
        "net_pay", "employer_pf", "ctc_monthly",
    ]

    master.to_csv(out / "salary_master.csv", index=False)
    write_branded_excel(master, out / "salary_master.xlsx",
                        sheet_name="Salary Master", numeric_cols=master_num,
                        title="Salary master — FY26 (synthetic, 50 employees)")

    reg_name = f"salary_register_{PAY_MONTH.replace('-', '')}"
    register.to_csv(out / f"{reg_name}.csv", index=False)
    write_branded_excel(register, out / f"{reg_name}.xlsx",
                        sheet_name="Payslip Register", numeric_cols=reg_num,
                        title=f"Payslip register — {PAY_MONTH} (synthetic)")

    print(f"Employees: {len(master)}")
    print(f"CTC range: Rs {master['gross_ctc'].min():,} - Rs {master['gross_ctc'].max():,}")
    print(f"Median CTC: Rs {int(master['gross_ctc'].median()):,}")
    print(f"Declarers (any 80C/80D/24b/rent): "
          f"{(master[['decl_80c_other','decl_80d_self','decl_24b_home_loan_int','rent_paid_annual']].sum(axis=1) > 0).sum()}/{len(master)}")
    print(f"Max CTC reconciliation gap: Rs {max_gap}")
    print("Wrote salary_master + " + reg_name + " (csv + xlsx)")


if __name__ == "__main__":
    main()
