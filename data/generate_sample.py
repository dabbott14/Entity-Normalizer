"""Generate a synthetic wire-activity workbook for demoing the normalizer.

Every name here is fictional. The data reproduces the problems real wire
data has: SWIFT "1/" prefixes, inconsistent case, punctuation and legal
suffixes, abbreviations, word-order flips, and typos.

Usage:
    python data/generate_sample.py            # writes data/sample_case.xlsx
    python data/generate_sample.py out.xlsx
"""

from __future__ import annotations

import random
import sys
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

SEED = 2026

# canonical name -> spellings that appear in the "bank data"
ENTITIES: dict[str, list[str]] = {
    "Harbor Point Logistics LLC": [
        "HARBOR POINT LOGISTICS LLC", "Harbor Point Logistics, LLC",
        "HARBOR PT LOGISTICS", "harbor point logistics llc", "HARBOUR POINT LOGISTICS LLC",
    ],
    "Bluewater Trading Co": [
        "BLUEWATER TRADING CO", "Bluewater Trading Company", "BLUE WATER TRADING CO.",
        "BLUEWATER TRADNG CO",
    ],
    "Meridian Coastal Bank NA": [
        "MERIDIAN COASTAL BANK NA", "MERIDIAN COASTAL BK N.A.", "Meridian Coastal Bank, N.A.",
        "MERIDIAN COASTAL BANK",
    ],
    "Silverline Holdings Inc": [
        "SILVERLINE HOLDINGS INC", "Silverline Holdings, Inc.", "SILVERLINE HLDGS INC",
        "SILVER LINE HOLDINGS",
    ],
    "Northgate International Services": [
        "NORTHGATE INTL SVCS", "Northgate International Services", "NORTHGATE INTERNATIONAL SVCS LTD",
    ],
    "Ridgeview Property Management": [
        "RIDGEVIEW PROPERTY MGMT", "Ridgeview Property Management LLC", "RIDGEVIEW PROP MGMT",
    ],
    "Unit 21/B Holdings LLC": [  # contains "1/" mid-name: must NOT be altered
        "UNIT 21/B HOLDINGS LLC", "Unit 21/B Holdings, LLC",
    ],
    "Marcus T Delaney": [
        "MARCUS T DELANEY", "DELANEY, MARCUS T", "Marcus Delaney", "MARCUS T. DELANY",
    ],
    "Priya Raman": ["PRIYA RAMAN", "RAMAN, PRIYA", "Priya Raman"],
    "Elena Kovac": ["ELENA KOVAC", "KOVAC ELENA", "ELENA KOVACS"],
    "Coastal Freight Partners": ["COASTAL FREIGHT PARTNERS", "Coastal Freight Partners LP"],
    "Granite Peak Consulting": ["GRANITE PEAK CONSULTING", "Granite Peak Consulting Group"],
}

# A case-only collision: the original VBA listed these as one name but only
# renamed the exact spelling shown, so the other spelling survived.
ENTITIES["Acme Industrial Supply"] = ["ACME INDUSTRIAL SUPPLY", "Acme Industrial Supply"]

CUSTOMER = "Halvorsen Marine Imports LLC"  # the case subject


def _pick(rng: random.Random) -> str:
    canonical = rng.choice(list(ENTITIES))
    spelling = rng.choice(ENTITIES[canonical])
    if rng.random() < 0.35 and "1/" not in spelling:
        spelling = "1/" + spelling  # structured SWIFT name line
    return spelling


def build(rng: random.Random | None = None) -> dict[str, pd.DataFrame]:
    rng = rng or random.Random(SEED)
    start = date(2026, 1, 2)

    wire_rows = []
    for i in range(320):
        incoming = rng.random() < 0.5
        cp = _pick(rng)
        wire_rows.append(
            {
                "Txn ID": f"W{100000 + i}",
                "Date": start + timedelta(days=rng.randint(0, 270)),
                "Direction": "Incoming" if incoming else "Outgoing",
                "Orig Acct Name": cp if incoming else CUSTOMER,
                "Bene Acct Name": CUSTOMER if incoming else cp,
                "Amount": round(rng.choice([rng.uniform(1_000, 9_900), rng.uniform(10_000, 250_000)]), 2),
            }
        )
    wire = pd.DataFrame(wire_rows)

    cash = pd.DataFrame(
        {
            "Txn ID": [f"C{200000 + i}" for i in range(60)],
            "Date": [start + timedelta(days=rng.randint(0, 270)) for _ in range(60)],
            "Type": [rng.choice(["Cash Deposit", "Cash Withdrawal"]) for _ in range(60)],
            "Amount": [round(rng.uniform(500, 9_800), 2) for _ in range(60)],
        }
    )

    check_rows = []
    for i in range(40):
        check_rows.append(
            {
                "Txn ID": f"K{300000 + i}",
                "Date": start + timedelta(days=rng.randint(0, 270)),
                "Check #": 4100 + i,
                # Payees on checks can share names with wire counterparties.
                "Payee": rng.choice(ENTITIES[rng.choice(list(ENTITIES))]),
                "Amount": round(rng.uniform(250, 40_000), 2),
            }
        )
    check = pd.DataFrame(check_rows)

    combined = pd.concat(
        [
            pd.DataFrame(
                {
                    "Tab Name": "Wire",
                    "Txn ID": wire["Txn ID"],
                    "Date": wire["Date"],
                    "Description": wire["Orig Acct Name"],
                    "Account Name": wire["Bene Acct Name"],
                    "Amount": wire["Amount"],
                }
            ),
            pd.DataFrame(
                {
                    "Tab Name": "Cash",
                    "Txn ID": cash["Txn ID"],
                    "Date": cash["Date"],
                    "Description": cash["Type"],
                    "Account Name": CUSTOMER,
                    "Amount": cash["Amount"],
                }
            ),
            pd.DataFrame(
                {
                    "Tab Name": "Check",
                    "Txn ID": check["Txn ID"],
                    "Date": check["Date"],
                    "Description": check["Payee"],
                    "Account Name": CUSTOMER,
                    "Amount": check["Amount"],
                }
            ),
        ],
        ignore_index=True,
    ).sort_values(["Date", "Txn ID"]).reset_index(drop=True)

    return {"Combined": combined, "Wire": wire, "Cash": cash, "Check": check}


def main(out: str | None = None) -> Path:
    path = Path(out) if out else Path(__file__).with_name("sample_case.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for name, df in build().items():
            df.to_excel(writer, sheet_name=name, index=False)
    return path


if __name__ == "__main__":
    print(f"Wrote {main(sys.argv[1] if len(sys.argv) > 1 else None)}")
