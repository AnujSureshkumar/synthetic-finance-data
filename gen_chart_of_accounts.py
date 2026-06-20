"""
gen_chart_of_accounts.py — a realistic Indian SME / ITeS chart of accounts.

Produces ~200 ledgers across the standard five heads, with the India-specific
accounts a controller actually uses: GST input/output by rate, TDS payable by
section, PF/ESI/PT, gratuity, and so on. Numeric coding follows a simple
block scheme so the GL extract and other generators can reference accounts by
range.

Code blocks (codes unchanged; chart presents in Schedule III balance-sheet order)
    3000-3999  Equity        (presented first)
    2000-2999  Liabilities   (presented second)
    1000-1999  Assets        (presented third)
    4000-4999  Income        (presented fourth)
    5000-8999  Expenses      (presented fifth)

Outputs (in ./output):
    chart_of_accounts.csv
    chart_of_accounts.xlsx   (brand-styled)

Consumed by: GL extract, invoice -> JE bot, GSTR-2B reconciler, dashboard.
"""

from __future__ import annotations

import pandas as pd

from common import ensure_output_dir
from excel_style import write_branded_excel

# Each tuple: (code, name, account_type, group, sub_group, normal_balance, is_control)
# is_control marks ledgers that summarise a subledger (debtors, creditors, etc.).

ACCOUNTS: list[tuple] = []


def add(code, name, atype, group, sub, normal, control=False):
    ACCOUNTS.append((code, name, atype, group, sub, normal, control))


# --------------------------------------------------------------------------- #
# 1000-1999  ASSETS
# --------------------------------------------------------------------------- #
# Non-current assets — Property, plant & equipment (gross + accumulated dep)
ppe = [
    ("Leasehold Improvements", 1010),
    ("Furniture & Fixtures", 1020),
    ("Office Equipment", 1030),
    ("Computers & Laptops", 1040),
    ("Servers & Networking Equipment", 1050),
    ("Plant & Machinery", 1060),
    ("Electrical Installations", 1070),
    ("Motor Vehicles", 1080),
]
for name, code in ppe:
    add(code, name, "Asset", "Non-current assets", "Property, plant & equipment", "Debit")
    add(code + 5, f"Accumulated Depreciation - {name}", "Asset",
        "Non-current assets", "Property, plant & equipment", "Credit")

# Intangibles & ROU
add(1100, "Software Licences (Intangible)", "Asset", "Non-current assets", "Intangible assets", "Debit")
add(1105, "Accumulated Amortisation - Software", "Asset", "Non-current assets", "Intangible assets", "Credit")
add(1110, "Goodwill", "Asset", "Non-current assets", "Intangible assets", "Debit")
add(1120, "Right-of-Use Asset - Office Premises", "Asset", "Non-current assets", "Right-of-use assets", "Debit")
add(1121, "Right-of-Use Asset - Equipment", "Asset", "Non-current assets", "Right-of-use assets", "Debit")
add(1125, "Accumulated Depreciation - ROU Assets", "Asset", "Non-current assets", "Right-of-use assets", "Credit")

# Non-current financial assets
add(1150, "Security Deposits - Rent", "Asset", "Non-current assets", "Other financial assets", "Debit")
add(1151, "Security Deposits - Utilities", "Asset", "Non-current assets", "Other financial assets", "Debit")
add(1160, "Deferred Tax Asset", "Asset", "Non-current assets", "Deferred tax", "Debit")
add(1170, "Investment in Subsidiary - Middle East", "Asset", "Non-current assets", "Investments", "Debit")

# Current assets — Trade receivables (control)
add(1200, "Trade Receivables - Domestic", "Asset", "Current assets", "Trade receivables", "Debit", control=True)
add(1201, "Trade Receivables - Export", "Asset", "Current assets", "Trade receivables", "Debit", control=True)
add(1205, "Allowance for Expected Credit Loss", "Asset", "Current assets", "Trade receivables", "Credit")
add(1210, "Unbilled Revenue", "Asset", "Current assets", "Trade receivables", "Debit")

# Cash & bank
banks = [
    "HDFC Bank - Current A/c", "ICICI Bank - Current A/c", "Axis Bank - Current A/c",
    "State Bank of India - Current A/c", "Kotak Mahindra Bank - Payroll A/c",
    "HDFC Bank - EEFC A/c (USD)",
]
for i, b in enumerate(banks):
    add(1300 + i, b, "Asset", "Current assets", "Cash & cash equivalents", "Debit")
add(1320, "Cash in Hand", "Asset", "Current assets", "Cash & cash equivalents", "Debit")
add(1325, "Petty Cash", "Asset", "Current assets", "Cash & cash equivalents", "Debit")
add(1330, "Fixed Deposits (< 3 months)", "Asset", "Current assets", "Cash & cash equivalents", "Debit")
add(1335, "Fixed Deposits (3-12 months)", "Asset", "Current assets", "Bank balances - other", "Debit")

# GST input credit (by rate) + electronic ledgers
# igst_r = full slab; half_r = CGST = SGST (each half the slab)
gst_slabs = [("0%", "0%"), ("5%", "2.5%"), ("12%", "6%"), ("18%", "9%"), ("28%", "14%")]
for i, (igst_r, half_r) in enumerate(gst_slabs):
    add(1400 + i * 3, f"Input CGST {half_r}", "Asset", "Current assets", "GST input credit", "Debit")
    add(1401 + i * 3, f"Input SGST {half_r}", "Asset", "Current assets", "GST input credit", "Debit")
    add(1402 + i * 3, f"Input IGST {igst_r}", "Asset", "Current assets", "GST input credit", "Debit")
add(1450, "GST Electronic Cash Ledger", "Asset", "Current assets", "GST input credit", "Debit")
add(1451, "GST Input Credit (Provisional / ineligible)", "Asset", "Current assets", "GST input credit", "Debit")

# Other taxes recoverable
add(1460, "TDS Receivable (26AS)", "Asset", "Current assets", "Other tax assets", "Debit")
add(1461, "TCS Receivable", "Asset", "Current assets", "Other tax assets", "Debit")
add(1462, "Advance Income Tax", "Asset", "Current assets", "Other tax assets", "Debit")
add(1463, "Income Tax Refund Due", "Asset", "Current assets", "Other tax assets", "Debit")

# Inventory (light, ITeS holds little)
add(1500, "Inventory - Consumables", "Asset", "Current assets", "Inventories", "Debit")
add(1501, "Inventory - Spares", "Asset", "Current assets", "Inventories", "Debit")

# Loans, advances & prepaids
add(1550, "Advance to Vendors", "Asset", "Current assets", "Other current assets", "Debit", control=True)
add(1551, "Advance to Employees", "Asset", "Current assets", "Other current assets", "Debit")
add(1552, "Prepaid Insurance", "Asset", "Current assets", "Other current assets", "Debit")
add(1553, "Prepaid Software Subscriptions", "Asset", "Current assets", "Other current assets", "Debit")
add(1554, "Prepaid Rent", "Asset", "Current assets", "Other current assets", "Debit")
add(1555, "Prepaid AMC", "Asset", "Current assets", "Other current assets", "Debit")
add(1560, "GST Refund Receivable (Export/LUT)", "Asset", "Current assets", "Other current assets", "Debit")
add(1561, "Receivable from Group Companies", "Asset", "Current assets", "Other current assets", "Debit", control=True)

# --------------------------------------------------------------------------- #
# 2000-2999  LIABILITIES
# --------------------------------------------------------------------------- #
# Borrowings
add(2000, "Term Loan - HDFC Bank", "Liability", "Non-current liabilities", "Borrowings", "Credit")
add(2001, "Working Capital Loan (CC)", "Liability", "Current liabilities", "Borrowings", "Credit")
add(2002, "Vehicle Loan", "Liability", "Non-current liabilities", "Borrowings", "Credit")

# Lease liabilities (IND-AS 116)
add(2010, "Lease Liability - Non-current", "Liability", "Non-current liabilities", "Lease liabilities", "Credit")
add(2011, "Lease Liability - Current", "Liability", "Current liabilities", "Lease liabilities", "Credit")

# Trade payables (control)
add(2100, "Trade Payables - Goods", "Liability", "Current liabilities", "Trade payables", "Credit", control=True)
add(2101, "Trade Payables - Services", "Liability", "Current liabilities", "Trade payables", "Credit", control=True)
add(2102, "Trade Payables - MSME", "Liability", "Current liabilities", "Trade payables", "Credit", control=True)
add(2103, "Trade Payables - Import", "Liability", "Current liabilities", "Trade payables", "Credit", control=True)
add(2110, "Creditors for Capital Goods", "Liability", "Current liabilities", "Trade payables", "Credit", control=True)

# GST output (by rate) + RCM
for i, (igst_r, half_r) in enumerate(gst_slabs):
    add(2200 + i * 3, f"Output CGST {half_r}", "Liability", "Current liabilities", "GST payable", "Credit")
    add(2201 + i * 3, f"Output SGST {half_r}", "Liability", "Current liabilities", "GST payable", "Credit")
    add(2202 + i * 3, f"Output IGST {igst_r}", "Liability", "Current liabilities", "GST payable", "Credit")
add(2250, "GST Payable under RCM", "Liability", "Current liabilities", "GST payable", "Credit")
add(2251, "GST Liability - Net Payable", "Liability", "Current liabilities", "GST payable", "Credit")

# TDS / TCS payable by section
tds_sections = [
    ("192", "Salary"), ("194C", "Contractors"), ("194J", "Professional Fees"),
    ("194I", "Rent"), ("194H", "Commission"), ("194Q", "Purchase of Goods"),
    ("195", "Non-resident Payments"),
]
for i, (sec, desc) in enumerate(tds_sections):
    add(2300 + i, f"TDS Payable - {sec} ({desc})", "Liability",
        "Current liabilities", "Statutory dues", "Credit")
add(2320, "TCS Payable - 206C", "Liability", "Current liabilities", "Statutory dues", "Credit")

# Payroll statutory
add(2350, "PF Payable (Employee + Employer)", "Liability", "Current liabilities", "Statutory dues", "Credit")
add(2351, "ESI Payable", "Liability", "Current liabilities", "Statutory dues", "Credit")
add(2352, "Professional Tax Payable", "Liability", "Current liabilities", "Statutory dues", "Credit")
add(2353, "Labour Welfare Fund Payable", "Liability", "Current liabilities", "Statutory dues", "Credit")
add(2354, "NPS Payable (Employer)", "Liability", "Current liabilities", "Statutory dues", "Credit")

# Employee-related
add(2400, "Salaries Payable", "Liability", "Current liabilities", "Employee benefits", "Credit")
add(2401, "Bonus Payable", "Liability", "Current liabilities", "Employee benefits", "Credit")
add(2402, "Reimbursements Payable", "Liability", "Current liabilities", "Employee benefits", "Credit")
add(2410, "Provision for Gratuity", "Liability", "Non-current liabilities", "Employee benefits", "Credit")
add(2411, "Provision for Leave Encashment", "Liability", "Non-current liabilities", "Employee benefits", "Credit")

# Other current liabilities & provisions
add(2500, "Accrued Expenses", "Liability", "Current liabilities", "Other current liabilities", "Credit")
add(2501, "Advance from Customers", "Liability", "Current liabilities", "Other current liabilities", "Credit", control=True)
add(2502, "Deferred Revenue (Unearned)", "Liability", "Current liabilities", "Other current liabilities", "Credit")
add(2503, "Payable to Group Companies", "Liability", "Current liabilities", "Other current liabilities", "Credit", control=True)
add(2510, "Provision for Income Tax", "Liability", "Current liabilities", "Provisions", "Credit")
add(2511, "Provision for Audit Fees", "Liability", "Current liabilities", "Provisions", "Credit")
add(2512, "Provision for Expenses", "Liability", "Current liabilities", "Provisions", "Credit")
add(2520, "Deferred Tax Liability", "Liability", "Non-current liabilities", "Deferred tax", "Credit")

# --------------------------------------------------------------------------- #
# 3000-3999  EQUITY
# --------------------------------------------------------------------------- #
add(3000, "Equity Share Capital", "Equity", "Equity", "Share capital", "Credit")
add(3001, "Preference Share Capital", "Equity", "Equity", "Share capital", "Credit")
add(3010, "Securities Premium", "Equity", "Equity", "Reserves & surplus", "Credit")
add(3020, "Retained Earnings", "Equity", "Equity", "Reserves & surplus", "Credit")
add(3021, "General Reserve", "Equity", "Equity", "Reserves & surplus", "Credit")
add(3030, "Current Year Profit & Loss", "Equity", "Equity", "Reserves & surplus", "Credit")
add(3040, "Foreign Currency Translation Reserve", "Equity", "Equity", "Other reserves", "Credit")
add(3050, "Remeasurement of Defined Benefit Plans (OCI)", "Equity", "Equity", "Other reserves", "Credit")

# --------------------------------------------------------------------------- #
# 4000-4999  INCOME
# --------------------------------------------------------------------------- #
add(4000, "Revenue - Software Services (Domestic)", "Income", "Revenue", "Operating revenue", "Credit")
add(4001, "Revenue - Software Services (Export)", "Income", "Revenue", "Operating revenue", "Credit")
add(4002, "Revenue - Annual Maintenance Contracts", "Income", "Revenue", "Operating revenue", "Credit")
add(4003, "Revenue - Product Licences", "Income", "Revenue", "Operating revenue", "Credit")
add(4004, "Revenue - SaaS Subscriptions", "Income", "Revenue", "Operating revenue", "Credit")
add(4005, "Revenue - Implementation & Consulting", "Income", "Revenue", "Operating revenue", "Credit")
add(4006, "Revenue - Managed Services", "Income", "Revenue", "Operating revenue", "Credit")
add(4010, "Sales Returns & Allowances", "Income", "Revenue", "Operating revenue", "Debit")
add(4020, "Interest Income - Deposits", "Income", "Other income", "Other income", "Credit")
add(4021, "Foreign Exchange Gain", "Income", "Other income", "Other income", "Credit")
add(4022, "Profit on Sale of Assets", "Income", "Other income", "Other income", "Credit")
add(4023, "Liabilities Written Back", "Income", "Other income", "Other income", "Credit")
add(4024, "Miscellaneous Income", "Income", "Other income", "Other income", "Credit")
add(4025, "Export Incentives", "Income", "Other income", "Other income", "Credit")

# --------------------------------------------------------------------------- #
# 5000-8999  EXPENSES
# --------------------------------------------------------------------------- #
# 5000s Direct / cost of services — employee benefits
add(5000, "Salaries & Wages", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5001, "Bonus & Incentives", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5002, "Employer PF Contribution", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5003, "Employer ESI Contribution", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5004, "Gratuity Expense", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5005, "Leave Encashment Expense", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5006, "Staff Welfare Expenses", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5007, "Recruitment Expenses", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5008, "Training & Development", "Expense", "Expenses", "Employee benefit expense", "Debit")
add(5009, "Contract Staff Cost", "Expense", "Expenses", "Employee benefit expense", "Debit")

# 5100s Technology / subcontracting (cost of revenue for ITeS)
add(5100, "Subcontracting Charges", "Expense", "Expenses", "Cost of services", "Debit")
add(5101, "Cloud Hosting (AWS/Azure/GCP)", "Expense", "Expenses", "Cost of services", "Debit")
add(5102, "Software Subscriptions (SaaS Tools)", "Expense", "Expenses", "Cost of services", "Debit")
add(5103, "Software Licence Fees", "Expense", "Expenses", "Cost of services", "Debit")
add(5104, "Data & Connectivity Charges", "Expense", "Expenses", "Cost of services", "Debit")
add(5105, "Third-party API / Platform Fees", "Expense", "Expenses", "Cost of services", "Debit")

# 6000s Administrative
add(6000, "Rent - Office Premises", "Expense", "Expenses", "Other expenses", "Debit")
add(6001, "Rates & Taxes", "Expense", "Expenses", "Other expenses", "Debit")
add(6002, "Electricity & Water", "Expense", "Expenses", "Other expenses", "Debit")
add(6003, "Repairs & Maintenance - Building", "Expense", "Expenses", "Other expenses", "Debit")
add(6004, "Repairs & Maintenance - Equipment", "Expense", "Expenses", "Other expenses", "Debit")
add(6005, "Housekeeping & Security", "Expense", "Expenses", "Other expenses", "Debit")
add(6006, "Office Supplies & Stationery", "Expense", "Expenses", "Other expenses", "Debit")
add(6007, "Printing", "Expense", "Expenses", "Other expenses", "Debit")
add(6008, "Postage & Courier", "Expense", "Expenses", "Other expenses", "Debit")
add(6009, "Telephone & Internet", "Expense", "Expenses", "Other expenses", "Debit")
add(6010, "Insurance - General", "Expense", "Expenses", "Other expenses", "Debit")
add(6011, "Insurance - Group Health (Employees)", "Expense", "Expenses", "Other expenses", "Debit")

# 6100s Professional & compliance
add(6100, "Legal & Professional Fees", "Expense", "Expenses", "Other expenses", "Debit")
add(6101, "Statutory Audit Fees", "Expense", "Expenses", "Other expenses", "Debit")
add(6102, "Internal Audit Fees", "Expense", "Expenses", "Other expenses", "Debit")
add(6103, "Tax & GST Consultancy", "Expense", "Expenses", "Other expenses", "Debit")
add(6104, "Company Secretarial / ROC Fees", "Expense", "Expenses", "Other expenses", "Debit")
add(6105, "Director Sitting Fees", "Expense", "Expenses", "Other expenses", "Debit")

# 6200s Travel & marketing
add(6200, "Travelling - Domestic", "Expense", "Expenses", "Other expenses", "Debit")
add(6201, "Travelling - Overseas", "Expense", "Expenses", "Other expenses", "Debit")
add(6202, "Conveyance", "Expense", "Expenses", "Other expenses", "Debit")
add(6203, "Vehicle Running & Fuel", "Expense", "Expenses", "Other expenses", "Debit")
add(6204, "Business Promotion", "Expense", "Expenses", "Other expenses", "Debit")
add(6205, "Advertising & Marketing", "Expense", "Expenses", "Other expenses", "Debit")
add(6206, "Conference & Events", "Expense", "Expenses", "Other expenses", "Debit")
add(6207, "Subscriptions & Memberships", "Expense", "Expenses", "Other expenses", "Debit")

# 6300s Other operating
add(6300, "Bank Charges", "Expense", "Expenses", "Other expenses", "Debit")
add(6301, "Foreign Exchange Loss", "Expense", "Expenses", "Other expenses", "Debit")
add(6302, "Rounding Off", "Expense", "Expenses", "Other expenses", "Debit")
add(6303, "Bad Debts Written Off", "Expense", "Expenses", "Other expenses", "Debit")
add(6304, "Provision for Doubtful Debts", "Expense", "Expenses", "Other expenses", "Debit")
add(6305, "Donations & CSR", "Expense", "Expenses", "Other expenses", "Debit")
add(6306, "Loss on Sale of Assets", "Expense", "Expenses", "Other expenses", "Debit")
add(6307, "GST Expense (ineligible ITC)", "Expense", "Expenses", "Other expenses", "Debit")
add(6308, "Miscellaneous Expenses", "Expense", "Expenses", "Other expenses", "Debit")

# 7000s Finance cost
add(7000, "Interest on Term Loan", "Expense", "Finance costs", "Finance costs", "Debit")
add(7001, "Interest on Working Capital", "Expense", "Finance costs", "Finance costs", "Debit")
add(7002, "Interest on Lease Liability", "Expense", "Finance costs", "Finance costs", "Debit")
add(7003, "Interest on Statutory Dues (late)", "Expense", "Finance costs", "Finance costs", "Debit")
add(7004, "Loan Processing Charges", "Expense", "Finance costs", "Finance costs", "Debit")

# 8000s Depreciation & amortisation, tax
add(8000, "Depreciation - PPE", "Expense", "Depreciation & amortisation", "Depreciation & amortisation", "Debit")
add(8001, "Depreciation - ROU Assets", "Expense", "Depreciation & amortisation", "Depreciation & amortisation", "Debit")
add(8002, "Amortisation - Intangibles", "Expense", "Depreciation & amortisation", "Depreciation & amortisation", "Debit")
add(8100, "Current Tax Expense", "Expense", "Tax expense", "Tax expense", "Debit")
add(8101, "Deferred Tax Expense", "Expense", "Tax expense", "Tax expense", "Debit")
add(8102, "Prior Period Tax Adjustments", "Expense", "Tax expense", "Tax expense", "Debit")


def build_df() -> pd.DataFrame:
    df = pd.DataFrame(
        ACCOUNTS,
        columns=[
            "account_code", "account_name", "account_type",
            "group", "sub_group", "normal_balance", "is_control",
        ],
    )
    # Present in Schedule III balance-sheet order; codes are unchanged
    type_order = {"Equity": 0, "Liability": 1, "Asset": 2, "Income": 3, "Expense": 4}
    df["_ord"] = df["account_type"].map(type_order)
    df = df.sort_values(["_ord", "account_code"]).drop(columns="_ord").reset_index(drop=True)
    # integrity checks
    assert df["account_code"].is_unique, "Duplicate account codes detected"
    return df


def main() -> None:
    out = ensure_output_dir()
    df = build_df()
    csv_path = out / "chart_of_accounts.csv"
    xlsx_path = out / "chart_of_accounts.xlsx"
    df.to_csv(csv_path, index=False)
    write_branded_excel(
        df, xlsx_path, sheet_name="Chart of Accounts",
        numeric_cols=["account_code"],
        title="Chart of Accounts — Indian SME / ITeS (synthetic)",
    )
    print(f"Chart of accounts: {len(df)} ledgers")
    print(df["account_type"].value_counts().to_string())
    print(f"Wrote {csv_path.name} and {xlsx_path.name}")


if __name__ == "__main__":
    main()
