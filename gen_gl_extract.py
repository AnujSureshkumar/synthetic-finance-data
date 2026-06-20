"""
gen_gl_extract.py - 24 months x 3 business units of budget-vs-actual P&L data.

This is the dataset behind the Budget vs Actual + BU Performance dashboard
(Project 3). It reuses the chart of accounts so the line items match every other
deliverable.

Model (kept economically coherent so the numbers survive scrutiny):
  - 24 months: FY2024-25 and FY2025-26 (Apr-2024 .. Mar-2026).
  - 3 business units of different scale, each with its own growth trend and
    Indian-FY seasonality (Q4 Jan-Mar strongest).
  - ~60 P&L accounts pulled from chart_of_accounts.csv.
  - Revenue accounts share each BU's revenue by fixed weights.
  - Variable cost accounts are a % of revenue; fixed accounts step up annually.
  - Budget is the smooth plan; actual = plan adjusted for the revenue beat/miss
    plus per-account efficiency noise.
  - A few deliberate variances are injected so the dashboard's auto-generated
    commentary has real stories to tell (a cloud-cost overrun, a revenue miss).

Long-format output (one row per month x BU x account) makes pivoting trivial:
    fy, month, period, bu, account_code, account_name, account_type,
    group, sub_group, budget, actual, variance, variance_pct

Synthetic only.

Outputs (in ./output):
    gl_extract.csv / .xlsx
"""

from __future__ import annotations

import pandas as pd

from common import ensure_output_dir, get_rng
from excel_style import write_branded_excel

# --------------------------------------------------------------------------- #
# Configuration
# --------------------------------------------------------------------------- #

# 24 months: Apr-2024 .. Mar-2026 (two complete Indian financial years)
START_YEAR, START_MONTH = 2024, 4
N_MONTHS = 24
N_ACCOUNTS = 60

# Business units: (name, base monthly revenue in INR, monthly growth rate)
BUSINESS_UNITS = [
    ("Digital Engineering", 9_500_000, 0.013),
    ("Cloud & Infrastructure", 6_200_000, 0.018),
    ("Product Licensing", 4_000_000, 0.009),
]

# Indian-FY monthly seasonality multipliers, indexed Apr(0)..Mar(11).
SEASONALITY = [0.92, 0.95, 0.98, 1.00, 0.97, 1.00,
               1.03, 1.02, 0.99, 1.05, 1.08, 1.12]

# Target cost structure as % of revenue for variable cost sub-groups.
VARIABLE_COST_TARGETS = {
    "Employee benefit expense": 0.50,   # dominant cost for an ITeS firm
    "Cost of services": 0.16,
}

# Fixed monthly cost as a fraction of the BU's base revenue (spread across
# the fixed accounts in each sub-group).
FIXED_COST_TARGETS = {
    "Other expenses": 0.10,
    "Finance costs": 0.02,
    "Depreciation & amortisation": 0.035,
}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #

def month_list():
    months = []
    y, m = START_YEAR, START_MONTH
    for _ in range(N_MONTHS):
        fy_start = y if m >= 4 else y - 1
        fy = "FY{}-{}".format(str(fy_start)[2:], str(fy_start + 1)[2:])
        months.append({
            "fy": fy,
            "month": "{}-{:02d}".format(y, m),
            "period": pd.Timestamp(year=y, month=m, day=1).strftime("%b-%Y"),
            "fy_index": (m - 4) % 12,      # 0=Apr .. 11=Mar
            "abs_index": len(months),       # 0..23 across the whole series
            "fy_year": fy_start,
        })
        m += 1
        if m > 12:
            m = 1
            y += 1
    return months


def load_pnl_accounts() -> pd.DataFrame:
    """Load the chart of accounts; (re)generate it if the CSV isn't there yet."""
    csv = ensure_output_dir() / "chart_of_accounts.csv"
    if not csv.exists():
        import gen_chart_of_accounts
        gen_chart_of_accounts.main()
    coa = pd.read_csv(csv)
    pnl = coa[coa["account_type"].isin(["Income", "Expense"])].copy()
    # Tax expense is booked at entity level, not per BU - exclude it.
    pnl = pnl[pnl["sub_group"] != "Tax expense"]

    # Pick a representative 60 across the P&L rather than just the lowest codes,
    # so finance costs and depreciation (high codes) stay in the dashboard.
    priority = [
        "Operating revenue", "Other income",
        "Employee benefit expense", "Cost of services",
        "Finance costs", "Depreciation & amortisation",
        "Other expenses",
    ]
    order = {s: i for i, s in enumerate(priority)}
    pnl["_p"] = pnl["sub_group"].map(order).fillna(len(priority))
    pnl = (pnl.sort_values(["_p", "account_code"])
              .head(N_ACCOUNTS)
              .drop(columns="_p")
              .sort_values("account_code")
              .reset_index(drop=True))
    return pnl


def assign_account_params(pnl: pd.DataFrame, rng):
    """
    Decide how each account behaves: revenue share, % of revenue, or fixed.
    Returns a dict keyed by account_code.
    """
    params = {}

    rev_codes = pnl[pnl["sub_group"] == "Operating revenue"]["account_code"].tolist()
    raw = {c: rng.uniform(0.5, 2.0) for c in rev_codes}
    total = sum(raw.values())
    rev_weights = {c: raw[c] / total for c in rev_codes}

    # variable cost accounts: split each sub-group's target % across its accounts
    for sub, target in VARIABLE_COST_TARGETS.items():
        codes = pnl[pnl["sub_group"] == sub]["account_code"].tolist()
        if not codes:
            continue
        raw = {c: rng.uniform(0.5, 2.0) for c in codes}
        s = sum(raw.values())
        for c in codes:
            params[c] = {"driver": "pct_rev", "pct": target * raw[c] / s}

    # fixed cost accounts
    for sub, target in FIXED_COST_TARGETS.items():
        codes = pnl[pnl["sub_group"] == sub]["account_code"].tolist()
        if not codes:
            continue
        raw = {c: rng.uniform(0.5, 2.0) for c in codes}
        s = sum(raw.values())
        for c in codes:
            params[c] = {"driver": "fixed", "frac": target * raw[c] / s}

    for c in rev_codes:
        params[c] = {"driver": "revenue", "weight": rev_weights[c]}

    # anything left (e.g. Other income, Sales Returns) -> small variable line
    for c in pnl["account_code"]:
        if c not in params:
            params[c] = {"driver": "pct_rev", "pct": rng.uniform(0.002, 0.01)}

    return params


def deliberate_modifier(bu_name, sub_group, account_name, fy_index, fy_year):
    """
    Inject a handful of named, explainable variances so the dashboard commentary
    has genuine stories. Returns a multiplicative factor applied to actual.
    """
    # Cloud & Infrastructure: cloud-hosting cost overrun in H2 FY25-26.
    if (bu_name == "Cloud & Infrastructure"
            and "Cloud Hosting" in account_name
            and fy_year == 2025 and fy_index >= 6):
        return 1.28
    # Product Licensing: revenue miss across Q3 FY25-26 (lost a renewal).
    if (bu_name == "Product Licensing"
            and sub_group == "Operating revenue"
            and fy_year == 2025 and fy_index in (6, 7, 8)):
        return 0.82
    # Digital Engineering: travel back above budget in FY25-26 (return to office).
    if (bu_name == "Digital Engineering"
            and account_name.startswith("Travelling")
            and fy_year == 2025):
        return 1.35
    return 1.0


# --------------------------------------------------------------------------- #
# Generation
# --------------------------------------------------------------------------- #

def generate() -> pd.DataFrame:
    rng = get_rng()
    months = month_list()
    pnl = load_pnl_accounts()
    params = assign_account_params(pnl, rng)

    rows = []
    for bu_name, base_rev, growth in BUSINESS_UNITS:
        budget_rev, actual_rev = {}, {}
        for mo in months:
            i = mo["abs_index"]
            plan = base_rev * ((1 + growth) ** i) * SEASONALITY[mo["fy_index"]]
            budget_rev[i] = plan
            beat = rng.gauss(0.01, 0.05)  # slight average beat, real spread
            actual_rev[i] = plan * (1 + beat)

        for _, acc in pnl.iterrows():
            code = acc["account_code"]
            p = params[code]
            for mo in months:
                i = mo["abs_index"]
                b_rev, a_rev = budget_rev[i], actual_rev[i]

                if p["driver"] == "revenue":
                    budget = b_rev * p["weight"]
                    actual = a_rev * p["weight"]
                elif p["driver"] == "pct_rev":
                    budget = b_rev * p["pct"]
                    eff = rng.gauss(0.0, 0.06)  # efficiency noise vs plan
                    actual = a_rev * p["pct"] * (1 + eff)
                else:  # fixed
                    step = 1.0 + 0.06 * (mo["fy_year"] - START_YEAR)
                    budget = base_rev * p["frac"] * step
                    actual = budget * (1 + rng.gauss(0.0, 0.03))

                actual *= deliberate_modifier(
                    bu_name, acc["sub_group"], acc["account_name"],
                    mo["fy_index"], mo["fy_year"])

                budget = round(budget)
                actual = round(actual)
                variance = actual - budget
                rows.append({
                    "fy": mo["fy"],
                    "month": mo["month"],
                    "period": mo["period"],
                    "bu": bu_name,
                    "account_code": code,
                    "account_name": acc["account_name"],
                    "account_type": acc["account_type"],
                    "group": acc["group"],
                    "sub_group": acc["sub_group"],
                    "budget": budget,
                    "actual": actual,
                    "variance": variance,
                    "variance_pct": round(variance / budget * 100, 1) if budget else 0.0,
                })

    return pd.DataFrame(rows)


def main() -> None:
    out = ensure_output_dir()
    df = generate()

    df.to_csv(out / "gl_extract.csv", index=False)
    write_branded_excel(
        df, out / "gl_extract.xlsx", sheet_name="GL Extract (BvA)",
        numeric_cols=["account_code", "budget", "actual", "variance"],
        title="GL extract - budget vs actual, 3 BUs x 24 months (synthetic)",
    )

    n_months = df["month"].nunique()
    n_bu = df["bu"].nunique()
    n_acc = df["account_code"].nunique()
    print("Rows: {}  ({} months x {} BUs x {} accounts)".format(
        len(df), n_months, n_bu, n_acc))
    print("Date span: {} .. {}".format(df["month"].min(), df["month"].max()))

    inc = df[df["account_type"] == "Income"].groupby("bu")["actual"].sum()
    exp = df[df["account_type"] == "Expense"].groupby("bu")["actual"].sum()
    print("BU operating margin (24-month actuals):")
    for bu in inc.index:
        margin = (inc[bu] - exp[bu]) / inc[bu] * 100
        print("  {:<26} revenue Rs {:5.1f} Cr  margin {:5.1f}%".format(
            bu, inc[bu] / 1e7, margin))

    print("Sub-groups represented: {}".format(df["sub_group"].nunique()))
    print("Wrote gl_extract.csv and gl_extract.xlsx")


if __name__ == "__main__":
    main()
