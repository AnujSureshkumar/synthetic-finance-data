"""
gen_lease_schedule.py - 10 IND-AS 116 leases with full ROU asset and lease
liability amortisation schedules (Project 6: lease accounting / disclosures).

Under IND-AS 116 a lessee, at commencement, recognises:
  - a lease liability = present value of remaining lease payments, discounted at
    the incremental borrowing rate (IBR); and
  - a right-of-use (ROU) asset, here taken as liability + initial direct costs.

Thereafter, each period:
  - the liability accrues interest (opening x monthly IBR) and is reduced by the
    payment -> interest unwinds, principal pays down;
  - the ROU asset is depreciated straight-line over the lease term.

This generator builds a master of 10 leases (offices, vehicles, equipment) with
varied terms, rents, annual escalations and IBRs, then expands each into a
month-by-month schedule. The schedules are the source data for IND-AS 116
disclosures (maturity analysis, interest expense, ROU depreciation) and tie out:
each lease's liability amortises to zero at the end of its term.

Outputs (in ./output):
    lease_master.csv / .xlsx       one row per lease (commencement-date figures)
    lease_schedule.csv / .xlsx     one row per lease x month

Synthetic only.
"""

from __future__ import annotations

import pandas as pd

from common import ensure_output_dir, get_rng
from excel_style import write_branded_excel

# (asset class, description, term band months, monthly rent band, annual
#  escalation choices, IBR band) - sampled per lease.
LEASE_TYPES = [
    ("Office premises", ["Office - Bengaluru HQ", "Office - Pune Delivery Centre",
                         "Office - Chennai Branch", "Office - Hyderabad Branch"],
     (60, 108), (350_000, 1_200_000), [0.05, 0.05, 0.07, 0.10], (0.09, 0.11)),
    ("Vehicles", ["Vehicle - Director (SUV)", "Vehicle - Pool Car"],
     (36, 60), (45_000, 90_000), [0.0, 0.05], (0.09, 0.105)),
    ("Equipment", ["Servers & Storage (Colo)", "Networking Equipment",
                   "Office Equipment - Copiers", "Laptops - Fleet Lease"],
     (24, 48), (60_000, 250_000), [0.0, 0.03, 0.05], (0.095, 0.115)),
]

N_LEASES = 10
START_BAND = ("2024-04", "2025-10")  # commencement dates spread across ~18 months


def month_add(ym: str, k: int) -> str:
    y, m = int(ym[:4]), int(ym[5:7])
    idx = (y * 12 + (m - 1)) + k
    return f"{idx // 12}-{idx % 12 + 1:02d}"


def pv_of_lease(monthly_rent: float, term: int, monthly_ibr: float,
                annual_esc: float) -> tuple[float, list[float]]:
    """
    Present value of the payment stream, with rent escalating once per 12 months.
    Returns (PV, list of the actual monthly payments).
    """
    payments = []
    pv = 0.0
    for t in range(1, term + 1):
        years_elapsed = (t - 1) // 12
        pay = monthly_rent * ((1 + annual_esc) ** years_elapsed)
        payments.append(pay)
        pv += pay / ((1 + monthly_ibr) ** t)
    return pv, payments


def build_leases(rng):
    masters, schedules = [], []
    for i in range(N_LEASES):
        cls, names, term_band, rent_band, esc_choices, ibr_band = \
            LEASE_TYPES[i % len(LEASE_TYPES)]
        term = rng.randint(*term_band)
        term = (term // 6) * 6 or 6           # round to a tidy 6-month multiple
        monthly_rent = int(round(rng.uniform(*rent_band) / 1000) * 1000)
        annual_esc = rng.choice(esc_choices)
        annual_ibr = round(rng.uniform(*ibr_band), 4)
        monthly_ibr = annual_ibr / 12
        start_idx = rng.randint(0, 18)
        commence = month_add(START_BAND[0], start_idx)

        pv, payments = pv_of_lease(monthly_rent, term, monthly_ibr, annual_esc)
        liability0 = round(pv)
        idc = int(round(rng.uniform(0, 0.02) * liability0 / 1000) * 1000)  # initial direct costs
        rou0 = liability0 + idc
        monthly_dep = rou0 / term

        masters.append({
            "lease_id": f"LSE{101 + i}",
            "asset_class": cls,
            "description": rng.choice(names),
            "commencement": commence,
            "term_months": term,
            "end_period": month_add(commence, term - 1),
            "monthly_rent_start": monthly_rent,
            "annual_escalation_pct": round(annual_esc * 100, 1),
            "ibr_annual_pct": round(annual_ibr * 100, 2),
            "lease_liability_initial": liability0,
            "initial_direct_costs": idc,
            "rou_asset_initial": rou0,
        })

        # expand the schedule
        liab = float(liability0)
        rou = float(rou0)
        for t in range(1, term + 1):
            period = month_add(commence, t - 1)
            interest = liab * monthly_ibr
            pay = payments[t - 1]
            principal = pay - interest
            closing_liab = liab - principal
            dep = monthly_dep
            rou_close = rou - dep
            # snap tiny float residue to zero on the final period
            if t == term:
                closing_liab = 0.0
                rou_close = 0.0
            schedules.append({
                "lease_id": f"LSE{101 + i}",
                "description": masters[-1]["description"],
                "period_no": t,
                "period": period,
                "opening_liability": round(liab),
                "lease_payment": round(pay),
                "interest_expense": round(interest),
                "principal_repayment": round(principal),
                "closing_liability": round(closing_liab),
                "rou_opening": round(rou),
                "rou_depreciation": round(dep),
                "rou_closing": round(rou_close),
            })
            liab = closing_liab
            rou = rou_close

    return pd.DataFrame(masters), pd.DataFrame(schedules)


def main() -> None:
    rng = get_rng()
    out = ensure_output_dir()
    master, schedule = build_leases(rng)

    master_num = ["term_months", "monthly_rent_start", "lease_liability_initial",
                  "initial_direct_costs", "rou_asset_initial"]
    sched_num = ["period_no", "opening_liability", "lease_payment",
                 "interest_expense", "principal_repayment", "closing_liability",
                 "rou_opening", "rou_depreciation", "rou_closing"]

    master.to_csv(out / "lease_master.csv", index=False)
    write_branded_excel(master, out / "lease_master.xlsx",
                        sheet_name="Lease Master", numeric_cols=master_num,
                        title="Lease master - IND-AS 116 (synthetic, 10 leases)")

    schedule.to_csv(out / "lease_schedule.csv", index=False)
    write_branded_excel(schedule, out / "lease_schedule.xlsx",
                        sheet_name="Amortisation", numeric_cols=sched_num,
                        title="Lease amortisation - ROU & liability (synthetic)")

    # --- integrity checks ---
    # each lease's liability must amortise to zero at the end of its term
    last = schedule.sort_values("period_no").groupby("lease_id").tail(1)
    assert (last["closing_liability"] == 0).all(), "A lease liability does not close to zero"
    assert (last["rou_closing"] == 0).all(), "An ROU asset does not fully depreciate"
    # principal + interest should equal the payment each row (within rounding)
    chk = (schedule["interest_expense"] + schedule["principal_repayment"]
           - schedule["lease_payment"]).abs()
    assert chk.max() <= 2, f"Payment split fails to reconcile (gap {chk.max()})"

    print(f"Leases: {len(master)}   Schedule rows: {len(schedule)}")
    print(f"Total initial lease liability: Rs {master['lease_liability_initial'].sum():,}")
    print(f"Total ROU asset (incl IDC):    Rs {master['rou_asset_initial'].sum():,}")
    print(f"Total interest over all terms: Rs {schedule['interest_expense'].sum():,}")
    print("By asset class:")
    print(master.groupby("asset_class")["lease_liability_initial"].agg(["count", "sum"]).to_string())
    print("All liabilities close to zero:",
          bool((last["closing_liability"] == 0).all()))
    print("Wrote lease_master and lease_schedule (csv + xlsx)")


if __name__ == "__main__":
    main()
