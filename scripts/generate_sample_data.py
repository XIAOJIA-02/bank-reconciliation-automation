"""
Generate the synthetic dataset used by this project.

Everything here is FICTIONAL: the company, funds, vendors, bank and amounts are
invented. Generating the data with a script (instead of shipping hand-typed
CSVs) makes it transparent which reconciling items were planted, so the
reconciliation engine can be tested against a known answer.

Entity : Harbourview Asset Management Pte Ltd (fictional)
Account: SGD operating account
Period : 1-31 October 2026, plus a 1-10 November 2026 bank statement
         used for subsequent-clearance testing.

Run:  python scripts/generate_sample_data.py
"""

from pathlib import Path
import csv
import json

OPENING_BALANCE = 186_420.50  # agreed to the 30 Sep 2026 reconciliation (no reconciling items b/f)
DATA_DIR = Path(__file__).resolve().parent.parent / "data"

# ---------------------------------------------------------------------------
# Transactions that appear in BOTH the general ledger and the bank statement.
# (gl_date, bank_date, gl_ref, gl_description, bank_description, amount)
# amount > 0 = money in, amount < 0 = money out.
# ---------------------------------------------------------------------------
COMMON = [
    ("2026-10-01", "2026-10-01", "PV-2610-001", "Office rent - October", "GIRO HARBOURVIEW PROPERTY PL RENT OCT", -12_800.00),
    ("2026-10-02", "2026-10-02", "RV-2610-001", "Mgmt fee Sep - Harbour Global Equity Fund", "INWARD TT HARBOUR GLOBAL EQ FD MGT FEE", 48_215.30),
    ("2026-10-02", "2026-10-02", "RV-2610-002", "Mgmt fee Sep - Harbour Income Fund", "INWARD TT HARBOUR INCOME FD MGT FEE", 21_664.85),
    ("2026-10-05", "2026-10-07", "CHQ100231", "Cheque 100231 - Brightline Office Supplies", "CHEQUE 100231", -486.20),
    ("2026-10-06", "2026-10-08", "CHQ100232", "Cheque 100232 - Keystone Cleaning Services", "CHEQUE 100232", -1_150.00),
    ("2026-10-08", "2026-10-08", "PV-2610-002", "IT managed services - Nimbustech (Oct)", "TT NIMBUSTECH SOLUTIONS PL", -4_500.00),
    ("2026-10-08", "2026-10-08", "PV-2610-003", "Market data subscription - Dataquill", "TT DATAQUILL ANALYTICS", -3_960.00),
    ("2026-10-09", "2026-10-12", "CHQ100233", "Cheque 100233 - Redwood Legal LLP", "CHEQUE 100233", -7_420.00),
    ("2026-10-12", "2026-10-12", "RV-2610-003", "Advisory fee - Client Pinecrest Family Office", "INWARD TT PINECREST FO ADVISORY", 15_000.00),
    ("2026-10-13", "2026-10-15", "CHQ100234", "Cheque 100234 - Lumen Courier", "CHEQUE 100234", -268.40),
    ("2026-10-14", "2026-10-14", "RV-2610-004", "Advisory fee cheque - Client Orchard Trust", "CHEQUE DEPOSIT", 3_850.00),
    ("2026-10-15", "2026-10-16", "CHQ100235", "Cheque 100235 - Summit Training Academy", "CHEQUE 100235", -2_200.00),
    ("2026-10-15", "2026-10-15", "PV-2610-004", "Utilities - October", "GIRO CITY UTILITIES", -642.35),
    ("2026-10-16", "2026-10-16", "RV-2610-005", "Performance fee Q3 - Harbour Global Equity Fund", "INWARD TT HARBOUR GLOBAL EQ FD PERF FEE", 96_480.00),
    ("2026-10-19", "2026-10-19", "PV-2610-005", "Fund admin fee - Coastal Fund Services", "TT COASTAL FUND SERVICES", -8_750.00),
    ("2026-10-20", "2026-10-22", "CHQ100237", "Cheque 100237 - Atlas Printing", "CHEQUE 100237", -930.00),
    ("2026-10-21", "2026-10-21", "PV-2610-006", "Professional indemnity insurance top-up", "TT MERIDIAN INSURANCE BROKERS", -5_310.00),
    ("2026-10-22", "2026-10-22", "PV-2610-007", "IT project milestone - Nimbustech", "TT NIMBUSTECH SOLUTIONS PL", -4_500.00),  # same amount as 8 Oct: tests ambiguity handling
    ("2026-10-23", "2026-10-26", "CHQ100238", "Cheque 100238 - Brightline Office Supplies", "CHEQUE 100238", -312.75),
    ("2026-10-26", "2026-10-26", "RV-2610-006", "Advisory fee - Client Pinecrest Family Office", "INWARD TT PINECREST FO ADVISORY", 15_000.00),
    ("2026-10-27", "2026-10-27", "PV-2610-008", "Payroll - October", "GIRO PAYROLL OCT 2026", -68_450.00),
    ("2026-10-27", "2026-10-27", "PV-2610-009", "Employer CPF contributions - October", "GIRO CPF CONTRIBUTION", -11_942.60),
    ("2026-10-28", "2026-10-28", "PV-2610-010", "Audit fee interim billing", "TT NORTHGATE ASSURANCE PAC", -9_800.00),
    ("2026-10-29", "2026-10-29", "PV-2610-011", "Cloud software - CloudSoft annual licence", "TT CLOUDSOFT PTE LTD", -2_400.00),
    ("2026-10-30", "2026-10-30", "RV-2610-007", "Mgmt fee Oct (in advance) - Harbour Income Fund", "INWARD TT HARBOUR INCOME FD MGT FEE", 22_105.40),
]

# ---------------------------------------------------------------------------
# Planted reconciling items (the "answer key").
# ---------------------------------------------------------------------------
# 1) Transposition error: cheque written and cleared for 1,350.00 but booked as 1,530.00.
TRANSPOSITION = ("2026-10-19", "2026-10-21", "CHQ100236", "Cheque 100236 - Keystone Cleaning Services", "CHEQUE 100236", -1_530.00, -1_350.00)

# 2) Duplicate posting in the GL: CloudSoft payment (PV-2610-011) entered twice.
DUPLICATE_GL = ("2026-10-29", "PV-2610-011", "Cloud software - CloudSoft annual licence", -2_400.00)

# 3) GL-only items at month end (timing differences) -> cleared or not in November.
GL_ONLY = [
    ("2026-10-20", "CHQ100239", "Cheque 100239 - Vantage Event Management", -3_120.00),   # never clears -> follow up
    ("2026-10-29", "CHQ100245", "Cheque 100245 - Redwood Legal LLP", -6_800.00),          # clears 3 Nov
    ("2026-10-30", "CHQ100246", "Cheque 100246 - Atlas Printing", -1_275.00),             # clears 4 Nov
    ("2026-10-31", "RV-2610-008", "Advisory fee cheque - Client Orchard Trust", 12_500.00),   # banked 2 Nov
    ("2026-10-31", "RV-2610-009", "Advisory fee - Client Tanjong Holdings", 8_000.00),        # banked 9 Nov -> cut-off risk
]

# 4) Bank-only items (not in the GL at month end).
BANK_ONLY = [
    ("2026-10-08", "TT CHARGES", -20.00),
    ("2026-10-19", "TT CHARGES", -20.00),
    ("2026-10-21", "RETURNED CHEQUE - INSUFFICIENT FUNDS ORCHARD TRUST", -3_850.00),
    ("2026-10-23", "CHEQUE 100299", -2_750.00),                         # not in cheque register -> bank error / unauthorised
    ("2026-10-26", "GIRO MERIDIAN INSURANCE PREMIUM", -960.00),          # direct debit not recorded
    ("2026-10-28", "INWARD TT REF 88213 NO DETAILS", 5_000.00),          # unidentified receipt
    ("2026-10-31", "INTEREST CREDIT", 214.66),
    ("2026-10-31", "SERVICE CHARGE - ACCOUNT MAINTENANCE", -30.00),
]

# 5) November bank statement (1-10 Nov) for subsequent-clearance testing.
NOV_BANK = [
    ("2026-11-02", "CHEQUE DEPOSIT", 12_500.00),
    ("2026-11-02", "GIRO HARBOURVIEW PROPERTY PL RENT NOV", -12_800.00),
    ("2026-11-03", "CHEQUE 100245", -6_800.00),
    ("2026-11-04", "CHEQUE 100246", -1_275.00),
    ("2026-11-05", "INWARD TT HARBOUR GLOBAL EQ FD MGT FEE", 49_002.10),
    ("2026-11-06", "TT NIMBUSTECH SOLUTIONS PL", -4_500.00),
    ("2026-11-09", "INWARD TT TANJONG HOLDINGS", 8_000.00),
    ("2026-11-10", "CHEQUE 100248", -720.00),
]


def build():
    gl_rows, bank_rows = [], []

    for gl_date, bank_date, ref, gl_desc, bank_desc, amt in COMMON:
        gl_rows.append([gl_date, ref, gl_desc, amt])
        bank_rows.append([bank_date, bank_desc, amt])

    gl_date, bank_date, ref, gl_desc, bank_desc, gl_amt, bank_amt = TRANSPOSITION
    gl_rows.append([gl_date, ref, gl_desc, gl_amt])
    bank_rows.append([bank_date, bank_desc, bank_amt])

    gl_rows.append(list(DUPLICATE_GL))
    for row in GL_ONLY:
        gl_rows.append(list(row))
    for row in BANK_ONLY:
        bank_rows.append(list(row))

    # Sort and number
    gl_rows.sort(key=lambda r: r[0])
    bank_rows.sort(key=lambda r: r[0])

    gl_out = []
    for i, (d, ref, desc, amt) in enumerate(gl_rows, start=1):
        gl_out.append({
            "date": d,
            "je_id": f"JE{i:04d}",
            "reference": ref,
            "description": desc,
            "debit": f"{amt:.2f}" if amt > 0 else "",
            "credit": f"{-amt:.2f}" if amt < 0 else "",
        })

    def bank_statement(rows, opening, prefix):
        out, bal = [], opening
        for i, (d, desc, amt) in enumerate(rows, start=1):
            bal = round(bal + amt, 2)
            out.append({
                "date": d,
                "bank_ref": f"{prefix}{i:04d}",
                "description": desc,
                "withdrawal": f"{-amt:.2f}" if amt < 0 else "",
                "deposit": f"{amt:.2f}" if amt > 0 else "",
                "balance": f"{bal:.2f}",
            })
        return out, bal

    bank_out, oct_close = bank_statement(bank_rows, OPENING_BALANCE, "BK2610-")
    nov_out, _ = bank_statement([tuple(r) for r in NOV_BANK], oct_close, "BK2611-")
    return gl_out, bank_out, nov_out


def write_csv(path, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main():
    DATA_DIR.mkdir(exist_ok=True)
    gl, bank, nov = build()
    write_csv(DATA_DIR / "general_ledger_oct2026.csv", gl)
    write_csv(DATA_DIR / "bank_statement_oct2026.csv", bank)
    write_csv(DATA_DIR / "bank_statement_nov2026_1-10.csv", nov)
    config = {
        "entity": "Harbourview Asset Management Pte Ltd (fictional)",
        "account": "Operating account",
        "period_end": "2026-10-31",
        "currency": "SGD",
        "opening_balance_bank": OPENING_BALANCE,
        "opening_balance_book": OPENING_BALANCE,
        "cheque_book_range": [100231, 100250],
        "note": "Synthetic data generated by scripts/generate_sample_data.py",
    }
    with open(DATA_DIR / "config.json", "w") as f:
        json.dump(config, f, indent=2)
    print(f"Wrote {len(gl)} GL lines, {len(bank)} Oct bank lines, {len(nov)} Nov bank lines to {DATA_DIR}")


if __name__ == "__main__":
    main()
