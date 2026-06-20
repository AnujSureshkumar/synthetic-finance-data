"""
common.py — shared helpers for the AI-for-Indian-CFOs synthetic data generators.

Everything in this file is reused across all six project datasets (chart of
accounts, salary slips, invoices, GL extracts, GSTR-2B, lease schedules).

Design rules:
- Deterministic by default. Every generator seeds from SEED so reruns are stable
  and diffs in version control are meaningful. Pass a different seed for variety.
- Synthetic only. No real people, no real GSTINs tied to real businesses, no real
  office data. Names are assembled from common-name pools; identifiers pass their
  format checksums but are not registered to anyone.
- Indian context. Names, states, vendor pools, and number formatting target an
  Indian SME / ITeS setting.

Brand palette (from D:\\Claude Projects\\instructions.md) is exposed here so any
generator that writes a styled Excel file pulls the same colours.
"""

from __future__ import annotations

import random
import string
from pathlib import Path

# --------------------------------------------------------------------------- #
# Reproducibility
# --------------------------------------------------------------------------- #

SEED = 20260606  # session date; bump or override for a fresh draw


def get_rng(seed: int | None = None) -> random.Random:
    """Return an isolated RNG so generators don't perturb each other's streams."""
    return random.Random(SEED if seed is None else seed)


# --------------------------------------------------------------------------- #
# Output location
# --------------------------------------------------------------------------- #

# All generated files land in ./output next to this file.
OUTPUT_DIR = Path(__file__).resolve().parent / "output"


def ensure_output_dir() -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    return OUTPUT_DIR


# --------------------------------------------------------------------------- #
# Brand colours (hex, no leading '#') for openpyxl PatternFill / Font
# --------------------------------------------------------------------------- #

BRAND = {
    "sea_green": "4DBBAE",
    "sea_green_light": "B5E2D9",
    "sea_green_wash": "E1EFEC",
    "sea_green_deep": "2A8676",
    "cream": "F8F6EF",
    "ink": "2E3A36",
    "muted": "6B7570",
    "hairline": "D6DBD7",
    "row_band": "F2EFE6",
}
BRAND_FONT = "Tahoma"


# --------------------------------------------------------------------------- #
# Name pools
# --------------------------------------------------------------------------- #

FIRST_NAMES = [
    "Aarav", "Vivaan", "Aditya", "Vihaan", "Arjun", "Sai", "Reyansh", "Krishna",
    "Ishaan", "Rohan", "Kabir", "Ayaan", "Dhruv", "Aryan", "Karthik", "Nikhil",
    "Rahul", "Siddharth", "Manish", "Anil", "Suresh", "Ramesh", "Vikram", "Deepak",
    "Ananya", "Diya", "Aadhya", "Saanvi", "Pari", "Anika", "Navya", "Myra",
    "Sneha", "Pooja", "Priya", "Divya", "Kavya", "Meera", "Lakshmi", "Sunita",
    "Neha", "Shreya", "Aishwarya", "Ritu", "Geeta", "Asha", "Nisha", "Swathi",
    "Farhan", "Zoya", "Imran", "Ayesha", "Thomas", "Maria", "Joseph", "Grace",
]

LAST_NAMES = [
    "Sharma", "Verma", "Gupta", "Iyer", "Nair", "Menon", "Reddy", "Rao",
    "Naidu", "Patel", "Shah", "Mehta", "Desai", "Joshi", "Kulkarni", "Deshpande",
    "Bhat", "Hegde", "Shetty", "Pillai", "Krishnan", "Subramanian", "Chatterjee",
    "Banerjee", "Mukherjee", "Das", "Bose", "Singh", "Yadav", "Kumar", "Mishra",
    "Pandey", "Tiwari", "Agarwal", "Jain", "Khan", "Sheikh", "Pereira", "Fernandes",
]


def full_name(rng: random.Random) -> str:
    return f"{rng.choice(FIRST_NAMES)} {rng.choice(LAST_NAMES)}"


# --------------------------------------------------------------------------- #
# Indian states (code, name) — GST state codes
# --------------------------------------------------------------------------- #

STATE_CODES = {
    "27": "Maharashtra",
    "29": "Karnataka",
    "33": "Tamil Nadu",
    "07": "Delhi",
    "36": "Telangana",
    "24": "Gujarat",
    "09": "Uttar Pradesh",
    "19": "West Bengal",
    "32": "Kerala",
    "06": "Haryana",
}


# --------------------------------------------------------------------------- #
# Identifier generators (format-valid, not registered to anyone)
# --------------------------------------------------------------------------- #

def random_pan(rng: random.Random, entity: str = "P") -> str:
    """
    PAN format: AAAAA9999A.
    - chars 1-3: random A-Z (alphabetic series)
    - char 4: entity type ('P' individual, 'C' company, 'F' firm, ...)
    - char 5: first letter of surname (random here)
    - chars 6-9: digits
    - char 10: alphabetic check char (random; real check is undocumented)
    """
    first3 = "".join(rng.choice(string.ascii_uppercase) for _ in range(3))
    fifth = rng.choice(string.ascii_uppercase)
    digits = "".join(rng.choice(string.digits) for _ in range(4))
    check = rng.choice(string.ascii_uppercase)
    return f"{first3}{entity}{fifth}{digits}{check}"


# GSTIN checksum uses base-36 over a 36-char code set, factor alternating 1/2.
_GSTIN_CODE = string.digits + string.ascii_uppercase  # '0'..'9','A'..'Z' -> 0..35


def _gstin_check_digit(first14: str) -> str:
    factor = 1
    total = 0
    mod = len(_GSTIN_CODE)  # 36
    for ch in first14:
        val = _GSTIN_CODE.index(ch)
        prod = val * factor
        total += prod // mod + prod % mod
        factor = 2 if factor == 1 else 1
    check_val = (mod - (total % mod)) % mod
    return _GSTIN_CODE[check_val]


def random_gstin(rng: random.Random, state_code: str | None = None,
                 pan: str | None = None) -> str:
    """
    GSTIN: SS + PAN(10) + entity(1) + 'Z' + checksum(1) = 15 chars.
    The 15th char is a real, verifiable checksum over the first 14.
    """
    if state_code is None:
        state_code = rng.choice(list(STATE_CODES.keys()))
    if pan is None:
        pan = random_pan(rng, entity="C")
    entity_digit = str(rng.randint(1, 9))
    first14 = f"{state_code}{pan}{entity_digit}Z"
    return first14 + _gstin_check_digit(first14)


def validate_gstin(gstin: str) -> bool:
    """True if the 15th char matches the checksum of the first 14."""
    if len(gstin) != 15:
        return False
    return gstin[14] == _gstin_check_digit(gstin[:14])


# --------------------------------------------------------------------------- #
# Money helpers
# --------------------------------------------------------------------------- #

def inr(amount: float) -> str:
    """Format a number in the Indian grouping system, e.g. 1234567 -> 12,34,567."""
    amount = round(amount)
    sign = "-" if amount < 0 else ""
    s = str(abs(amount))
    if len(s) <= 3:
        return sign + s
    last3 = s[-3:]
    rest = s[:-3]
    # group the remaining digits in pairs
    parts = []
    while len(rest) > 2:
        parts.insert(0, rest[-2:])
        rest = rest[:-2]
    if rest:
        parts.insert(0, rest)
    return sign + ",".join(parts) + "," + last3


def round_to(value: float, step: int = 100) -> int:
    """Round to the nearest `step` — payroll figures are rarely to the rupee."""
    return int(round(value / step) * step)
