# Synthetic data generators — Project Zero

Shared mock datasets for the *AI for Indian CFOs* portfolio. Every later project
draws on these so all six demos tell a consistent story with the same fictional
ITeS company, the same chart of accounts, and the same people.

**Data discipline:** everything here is synthetic. Names are assembled from
common-name pools; PANs and GSTINs are format-valid (GSTIN passes its real
checksum) but are not registered to any person or business. No office data ever
touches these files.

## Layout

```
synthetic-data/
  common.py                 shared helpers (names, PAN/GSTIN, INR format, RNG, brand colours)
  excel_style.py            one function to write brand-styled .xlsx (Sea Green / Tahoma)
  gen_chart_of_accounts.py  ~200-ledger Indian SME / ITeS chart of accounts
  gen_salary_slips.py       50 employees: annual CTC master + monthly payslip register
  gen_gl_extract.py         24 months x 3 BUs x 60 accounts of budget-vs-actual P&L
  gen_invoices.py           12 vendors, ~60 purchase invoices + PDF tax invoices
  gen_gstr_2b.py            GSTR-2B JSON + books register with seeded mismatches
  gen_lease_schedule.py     10 IND-AS 116 leases with ROU + liability schedules
  requirements.txt          pandas, openpyxl, reportlab
  output/                   generated CSV + XLSX + PDF + JSON (re-creatable; safe to delete)
```

## Run

```bash
cd synthetic-data
pip install -r requirements.txt
python gen_chart_of_accounts.py
python gen_salary_slips.py
python gen_gl_extract.py
python gen_invoices.py
python gen_gstr_2b.py
python gen_lease_schedule.py
```

Output lands in `output/`. Generators are **seeded** (`SEED` in `common.py`), so
reruns reproduce identical files. Change the seed for a fresh draw.

## What each generator produces

### `gen_chart_of_accounts.py`

A 208-ledger chart of accounts in Indian SME / ITeS format, coded in blocks:

| Block | Head |
|------:|------|
| 1000–1999 | Equity |
| 2000–2999 | Liabilities |
| 3000–3999 | Assets |
| 4000–4999 | Income |
| 5000–8999 | Expenses |

Includes the India-specific ledgers a controller actually posts to: GST
input/output by rate, TDS payable by section (192, 194C, 194J, 194I, 194Q…),
PF/ESI/PT, gratuity and leave provisions, IND-AS 116 ROU assets and lease
liabilities, EEFC/export ledgers, and deferred tax.

Columns: `account_code, account_name, account_type, group, sub_group,
normal_balance, is_control`.

*Consumed by:* GL extract (Project 3), Invoice→JE bot (Project 2), GSTR-2B
reconciler (Project 5), BvA dashboard (Project 3).

### `gen_salary_slips.py`

**`salary_master`** — one row per employee, annual figures. This is the direct
input to the **Tax Regime Optimizer (Project 1)**. Data dictionary:

| Field | Meaning |
|-------|---------|
| `emp_id, name, pan, gender, age` | identity (age drives senior-citizen slabs) |
| `date_of_joining, department, designation, location, metro` | context; `metro` drives HRA 50% vs 40% cap |
| `gross_ctc` | total cost to company (annual) |
| `basic` | 45% of CTC |
| `hra_component` | 50% of basic (the *paid* HRA; exemption is computed by the optimizer) |
| `special_allowance, lta` | balancing + LTA components |
| `employer_pf` | 12% of basic, capped ₹21,600/yr |
| `gratuity` | 4.81% of basic |
| `employer_nps_80ccd2` | employer NPS (10% of basic for ~25% of staff) — deductible in **both** regimes |
| `employee_pf_80c` | employee PF — counts toward the ₹1.5L 80C cap |
| `professional_tax` | ₹2,400/yr (old-regime deduction) |
| `rent_paid_annual` | rent paid (0 if employee owns home) — HRA exemption input |
| `decl_80c_other` | ELSS/LIC/PPF/tuition beyond PF (kept within the 80C cap) |
| `decl_80ccd_1b_nps` | additional NPS, ₹50k ceiling |
| `decl_80d_self, decl_80d_parents` | health insurance premia |
| `decl_80e_edu_loan_int` | education-loan interest (no cap) |
| `decl_24b_home_loan_int` | self-occupied home-loan interest, ₹2L ceiling |

Roughly 30% of employees are modelled as "non-declarers" (lean new regime); the
rest declare a realistic mix — so the optimizer's old-vs-new result varies
across the population instead of always favouring one side.

**`salary_register_<month>`** — one month's payslip register derived from the
master: per-employee earnings (basic/HRA/special/LTA), statutory deductions
(employee PF, PT), an **indicative** TDS, and net pay. The TDS uses the FY26
new-regime slabs (standard deduction ₹75,000; nil tax up to ₹12L taxable via
87A; 4% cess) — indicative only, not a payroll-grade computation.

*Consumed by:* Tax Regime Optimizer (Project 1, master), payroll journal entry
in Invoice→JE bot (Project 2, register), headcount cost lines in the BvA
dashboard (Project 3, register).

### `gen_gl_extract.py`

24 months (Apr-2024 .. Mar-2026, two full Indian FYs) × 3 business units × 60
P&L accounts, in **long format** (one row per month × BU × account) so pivoting
into a budget-vs-actual view is trivial. 4,320 rows.

The three BUs differ in scale, growth and margin (Digital Engineering ~₹27.8 Cr
/ 22.8%, Cloud & Infrastructure ~₹19.4 Cr / 23.6%, Product Licensing ~₹10.9 Cr /
20.9% over the 24 months). Revenue follows Indian-FY seasonality (Q4 Jan–Mar
strongest). Budget is the smooth plan; actual = plan adjusted for the revenue
beat/miss plus per-account efficiency noise.

Three **deliberate, named variances** are injected so the dashboard's
auto-commentary has real stories: a Cloud-hosting cost overrun (+~28%) in H2
FY25-26, a Product Licensing revenue miss (−~18%) across Q3 FY25-26 (lost
renewal), and Digital Engineering travel back above budget (+~35%) in FY25-26
(return to office).

Columns: `fy, month, period, bu, account_code, account_name, account_type,
group, sub_group, budget, actual, variance, variance_pct`.

*Consumed by:* Budget vs Actual + BU Performance dashboard (Project 3).

### `gen_invoices.py`

A **vendor master** of 12 suppliers (one per expense category, valid-format
GSTINs/PANs, MSME flags, payment terms) and a **purchase register** of ~60
invoices that serve as the *ground truth* for the Invoice → Journal Entry bot
(Project 2). The buyer is our fictional ITeS company in Karnataka (state 29):
intra-state vendors attract CGST+SGST, out-of-state vendors attract IGST, driven
by place of supply. Each vendor's category fixes the expense ledger (matched to
`chart_of_accounts.csv`), the GST rate, and the TDS section; TDS is computed on
the taxable value. Three invoices are deliberately booked under **reverse charge**
(vendor charges no GST, buyer self-assesses) so the bot must handle RCM.

It also renders the first 12 invoices as **PDF tax invoices** (`invoices_pdf/`,
brand-styled via reportlab) — the bot's raw OCR/extraction input. Integrity
checks confirm every GSTIN passes its checksum and every invoice total
reconciles to taxable + CGST + SGST + IGST.

*Consumed by:* Invoice → JE bot (Project 2).

### `gen_gstr_2b.py`

A matched pair for the **GSTR-2B reconciler** (Project 5): a CBIC-style
**GSTR-2B JSON** (`gstr2b_062026.json`, return period June 2026, b2b section with
per-supplier invoices and an ITC summary) alongside the company's own **books
purchase register**. Both are derived from one base set, then deliberately pulled
apart so every reconciliation outcome appears — MATCHED, VALUE_MISMATCH,
GSTIN_MISMATCH, ONLY_IN_2B (supplier filed, not booked), and ONLY_IN_BOOKS
(booked, supplier hasn't filed). The expected classification per invoice is
written to `gstr2b_recon_truth.csv` so the reconciler can be scored.

*Consumed by:* GSTR-2B reconciler (Project 5).

### `gen_lease_schedule.py`

Ten **IND-AS 116 leases** (offices, vehicles, equipment) with varied terms,
rents, annual escalations and incremental borrowing rates. A **lease master**
holds commencement-date figures (lease liability = PV of payments; ROU asset =
liability + initial direct costs) and the generator expands each lease into a
month-by-month **amortisation schedule**: opening liability, interest unwind,
principal repayment, closing liability, and straight-line ROU depreciation.
Integrity checks confirm every lease's liability amortises to zero at end of term
and interest + principal ties to the payment each period. This is the source data
for IND-AS 116 disclosures (maturity analysis, interest expense, ROU depreciation).

*Consumed by:* Lease accounting / disclosures (Project 6).

## FY26 tax note

`fy26_new_regime_tax()` in `gen_salary_slips.py` encodes the FY26 (AY 2026-27)
new-regime slabs and is reused by the register for indicative TDS. The Tax
Regime Optimizer will own the authoritative engine (both regimes, surcharge,
marginal relief, all exemptions); this helper is deliberately simplified.

## Status

All six generators (roadmap §4) are built and verified: chart of accounts,
salary slips, GL extract, invoices, GSTR-2B, and lease schedules. Every dataset
is seeded and reproducible, draws on the shared chart of accounts, and passes its
own integrity checks on each run.
