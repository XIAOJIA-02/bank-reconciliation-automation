"""
Bank reconciliation engine.

Pipeline
--------
1. Load the general ledger (GL) cash account and the bank statement.
2. Flag possible duplicate postings in the GL (same reference + same amount).
3. Match GL lines to bank lines in three passes:
     Pass 1  cheque number + exact amount
     Pass 2  exact amount + date within a tolerance window (closest date wins)
     Pass 3  cheque number agrees but amount differs  -> "amount mismatch"
4. Classify everything left unmatched into reconciling-item categories.
5. Test month-end timing items against the next month's bank statement
   (subsequent clearance - a standard audit procedure for cut-off).
6. Build the reconciliation statement and prove that
   adjusted bank balance == adjusted book balance.

Sign convention: every amount is signed from the company's point of view.
    +  money into the bank account (GL debit / bank deposit)
    -  money out of the bank account (GL credit / bank withdrawal)
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
DATE_WINDOW_DAYS = 3          # Pass 2 tolerance between GL date and bank date
LATE_CLEARANCE_DAYS = 5       # receipts banked later than this after period end -> cut-off risk
DEFAULT_THRESHOLD = 5_000.00  # items at or above this amount are escalated (set per engagement)

CHEQUE_RE = re.compile(r"(?:CHQ|CHEQUE)\s*(\d{6})", re.IGNORECASE)

# Reconciling-item categories: (name, side adjusted, normal priority)
BANK_SIDE = "Bank"
BOOK_SIDE = "Book"


@dataclass
class ReconResult:
    config: dict
    gl: pd.DataFrame
    bank: pd.DataFrame
    next_bank: pd.DataFrame
    matches: pd.DataFrame
    exceptions: pd.DataFrame
    clearance: pd.DataFrame
    summary: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# 1. Load
# ---------------------------------------------------------------------------
def _num(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors="coerce").fillna(0.0)


def cheque_no(text: str) -> str:
    """Return the 6-digit cheque number found in a reference/description, if any."""
    m = CHEQUE_RE.search(str(text))
    return m.group(1) if m else ""


def load_gl(path: Path) -> pd.DataFrame:
    gl = pd.read_csv(path, dtype=str)
    gl["date"] = pd.to_datetime(gl["date"])
    gl["amount"] = (_num(gl["debit"]) - _num(gl["credit"])).round(2)
    gl["cheque_no"] = gl["reference"].map(cheque_no).fillna("")
    return gl


def load_bank(path: Path) -> pd.DataFrame:
    bank = pd.read_csv(path, dtype=str)
    bank["date"] = pd.to_datetime(bank["date"])
    bank["amount"] = (_num(bank["deposit"]) - _num(bank["withdrawal"])).round(2)
    bank["balance"] = _num(bank["balance"])
    bank["cheque_no"] = bank["description"].map(cheque_no).fillna("")
    return bank


def load_config(path: Path) -> dict:
    with open(path) as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# 2. Duplicate detection (GL)
# ---------------------------------------------------------------------------
def find_gl_duplicates(gl: pd.DataFrame) -> set[str]:
    """
    Return je_ids of the SECOND and later postings that share a reference and
    amount with an earlier posting. The first posting is treated as genuine.
    """
    dup_mask = gl.sort_values(["date", "je_id"]).duplicated(subset=["reference", "amount"], keep="first")
    return set(gl.sort_values(["date", "je_id"]).loc[dup_mask, "je_id"])


# ---------------------------------------------------------------------------
# 3. Matching
# ---------------------------------------------------------------------------
def match_transactions(gl: pd.DataFrame, bank: pd.DataFrame, exclude_gl: set[str]) -> pd.DataFrame:
    """One-to-one matching. Returns a table of matched pairs with the method used."""
    matches = []
    used_gl = set(exclude_gl)
    used_bank: set[str] = set()

    # Pass 1: cheque number + exact amount
    for _, g in gl[gl["cheque_no"] != ""].iterrows():
        if g["je_id"] in used_gl:
            continue
        cand = bank[(bank["cheque_no"] == g["cheque_no"]) & (bank["amount"] == g["amount"])
                    & (~bank["bank_ref"].isin(used_bank))]
        if not cand.empty:
            b = cand.iloc[0]
            matches.append(_pair(g, b, "1 - Cheque no. + amount"))
            used_gl.add(g["je_id"]); used_bank.add(b["bank_ref"])

    # Pass 2: exact amount + closest date within the window (non-cheque items)
    for _, g in gl.sort_values("date").iterrows():
        if g["je_id"] in used_gl or g["cheque_no"]:
            continue
        cand = bank[(bank["amount"] == g["amount"]) & (bank["cheque_no"] == "")
                    & (~bank["bank_ref"].isin(used_bank))].copy()
        cand["gap"] = (cand["date"] - g["date"]).dt.days.abs()
        cand = cand[cand["gap"] <= DATE_WINDOW_DAYS].sort_values(["gap", "date"])
        if not cand.empty:
            b = cand.iloc[0]
            matches.append(_pair(g, b, "2 - Amount + date window"))
            used_gl.add(g["je_id"]); used_bank.add(b["bank_ref"])

    # Pass 3: same cheque number, different amount -> matched but flagged
    for _, g in gl[gl["cheque_no"] != ""].iterrows():
        if g["je_id"] in used_gl:
            continue
        cand = bank[(bank["cheque_no"] == g["cheque_no"]) & (~bank["bank_ref"].isin(used_bank))]
        if not cand.empty:
            b = cand.iloc[0]
            matches.append(_pair(g, b, "3 - Cheque no. only (AMOUNT MISMATCH)"))
            used_gl.add(g["je_id"]); used_bank.add(b["bank_ref"])

    return pd.DataFrame(matches)


def _pair(g: pd.Series, b: pd.Series, method: str) -> dict:
    return {
        "je_id": g["je_id"], "gl_date": g["date"], "gl_reference": g["reference"],
        "gl_description": g["description"], "gl_amount": g["amount"],
        "bank_ref": b["bank_ref"], "bank_date": b["date"], "bank_description": b["description"],
        "bank_amount": b["amount"], "difference": round(b["amount"] - g["amount"], 2),
        "days_apart": int((b["date"] - g["date"]).days), "match_method": method,
    }


# ---------------------------------------------------------------------------
# 4-5. Classification + subsequent clearance
# ---------------------------------------------------------------------------
def find_clearance(item: pd.Series, next_bank: pd.DataFrame, used: set[str]) -> pd.Series | None:
    """Find the line in next month's statement that clears an outstanding GL item."""
    if item["cheque_no"]:
        cand = next_bank[(next_bank["cheque_no"] == item["cheque_no"]) & (next_bank["amount"] == item["amount"])]
    else:
        cand = next_bank[(next_bank["amount"] == item["amount"]) & (next_bank["cheque_no"] == "")]
    cand = cand[~cand["bank_ref"].isin(used)].sort_values("date")
    return None if cand.empty else cand.iloc[0]


def classify(gl, bank, next_bank, matches, duplicates, config, threshold) -> tuple[pd.DataFrame, pd.DataFrame]:
    period_end = pd.Timestamp(config["period_end"])
    lo, hi = config["cheque_book_range"]
    rows, clearance_rows, used_next = [], [], set()

    def add(source, ref_id, date, reference, description, amount, category, side, effect,
            aje, action, priority, cleared_on=None):
        # Large items are escalated, except timing differences already proven by subsequent clearance
        if abs(amount) >= threshold and priority != "High" and cleared_on is None:
            priority = "High"
            action += f" Amount >= threshold ({threshold:,.0f})."
        rows.append({
            "source": source, "source_id": ref_id, "date": date, "reference": reference,
            "description": description, "amount": amount, "category": category,
            "adjust_side": side, "recon_effect": round(effect, 2), "proposed_entry": aje,
            "follow_up_action": action, "priority": priority, "cleared_on": cleared_on,
        })

    matched_gl = set(matches["je_id"]) if not matches.empty else set()
    matched_bank = set(matches["bank_ref"]) if not matches.empty else set()

    # (a) Amount mismatches on matched pairs (Pass 3)
    for _, m in matches[matches["difference"] != 0].iterrows():
        diff = m["difference"]
        hint = " Difference divisible by 9 -> likely transposition error." if round(abs(diff) * 100) % 9 == 0 else ""
        add("GL vs Bank", m["je_id"], m["gl_date"], m["gl_reference"],
            f"{m['gl_description']} | booked {abs(m['gl_amount']):,.2f}, bank cleared {abs(m['bank_amount']):,.2f}",
            diff, "Book error - amount mismatch", BOOK_SIDE, diff,
            f"Dr Bank {diff:,.2f} / Cr Expense {diff:,.2f}" if diff > 0 else f"Dr Expense {-diff:,.2f} / Cr Bank {-diff:,.2f}",
            "Agree to cheque image and invoice; bank pays the cheque's face value, so correct the book entry." + hint,
            "Medium")

    # (b) Unmatched GL lines
    for _, g in gl[~gl["je_id"].isin(matched_gl)].iterrows():
        if g["je_id"] in duplicates:
            add("GL", g["je_id"], g["date"], g["reference"], g["description"], g["amount"],
                "Book error - duplicate posting", BOOK_SIDE, -g["amount"],
                f"Reverse duplicate {g['je_id']}: Dr Bank {abs(g['amount']):,.2f} / Cr Expense {abs(g['amount']):,.2f}",
                "Confirm only one invoice/payment exists; reverse the duplicate journal.", "Medium")
            continue

        hit = find_clearance(g, next_bank, used_next)
        cleared_on = hit["date"] if hit is not None else None
        if hit is not None:
            used_next.add(hit["bank_ref"])
        days_after = (cleared_on - period_end).days if cleared_on is not None else None

        if g["amount"] > 0:
            category = "Deposit in transit"
            if cleared_on is None:
                action, priority = "Not banked in subsequent period - verify receipt exists (existence/cut-off).", "High"
            elif days_after > LATE_CLEARANCE_DAYS:
                action, priority = (f"Banked {days_after} days after period end - cut-off risk; inspect remittance "
                                    "and confirm receipt belongs to October."), "High"
            else:
                action, priority = f"Cleared {cleared_on:%d %b %Y} - timing difference only.", "Low"
        else:
            category = "Outstanding cheque" if g["cheque_no"] else "Unpresented payment"
            if cleared_on is None:
                action, priority = ("Not presented by end of subsequent period - confirm with payee; "
                                    "consider whether cheque is lost, void or should be written back."), "High"
            else:
                action, priority = f"Cleared {cleared_on:%d %b %Y} - timing difference only.", "Low"

        add("GL", g["je_id"], g["date"], g["reference"], g["description"], g["amount"],
            category, BANK_SIDE, g["amount"], "No entry - timing difference", action, priority, cleared_on)
        clearance_rows.append({
            "je_id": g["je_id"], "reference": g["reference"], "description": g["description"],
            "gl_date": g["date"], "amount": g["amount"], "category": category,
            "cleared_on": cleared_on, "cleared_bank_ref": hit["bank_ref"] if hit is not None else None,
            "days_after_period_end": days_after, "status": "Cleared" if cleared_on is not None else "NOT CLEARED",
        })

    # (c) Unmatched bank lines
    for _, b in bank[~bank["bank_ref"].isin(matched_bank)].iterrows():
        desc, amt = b["description"].upper(), b["amount"]
        if b["cheque_no"] and not (lo <= int(b["cheque_no"]) <= hi):
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Bank error / unauthorised cheque", BANK_SIDE, -amt, "No book entry - bank to reverse",
                f"Cheque {b['cheque_no']} is outside the company's cheque book ({lo}-{hi}). "
                "Request cheque image, notify bank, escalate as potential fraud.", "High")
        elif b["cheque_no"]:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Unrecorded cheque", BOOK_SIDE, amt, f"Dr Expense/Payable {abs(amt):,.2f} / Cr Bank {abs(amt):,.2f}",
                "Company cheque cleared but not recorded - obtain support and record.", "Medium")
        elif "RETURN" in desc or "DISHONOUR" in desc:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Dishonoured cheque", BOOK_SIDE, amt, f"Dr Receivable {abs(amt):,.2f} / Cr Bank {abs(amt):,.2f}",
                "Reinstate receivable; contact client; assess recoverability (impairment).", "Medium")
        elif "INTEREST" in desc and amt > 0:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Interest not recorded", BOOK_SIDE, amt, f"Dr Bank {amt:,.2f} / Cr Interest income {amt:,.2f}",
                "Agree rate to bank confirmation / facility letter.", "Low")
        elif ("CHARGE" in desc or "FEE" in desc) and amt < 0:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Bank charges not recorded", BOOK_SIDE, amt, f"Dr Bank charges {abs(amt):,.2f} / Cr Bank {abs(amt):,.2f}",
                "Agree to bank tariff.", "Low")
        elif amt < 0:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Unrecorded direct debit", BOOK_SIDE, amt, f"Dr Expense {abs(amt):,.2f} / Cr Bank {abs(amt):,.2f}",
                "Obtain invoice / GIRO mandate; confirm authorised before recording.", "Medium")
        else:
            add("Bank", b["bank_ref"], b["date"], b["bank_ref"], b["description"], amt,
                "Unidentified receipt", BOOK_SIDE, amt, f"Dr Bank {amt:,.2f} / Cr Suspense {amt:,.2f}",
                "Identify remitter with bank; clear suspense; consider client money / AML implications.", "High")

    exceptions = pd.DataFrame(rows).sort_values(["adjust_side", "category", "date"]).reset_index(drop=True)
    exceptions.insert(0, "item_id", [f"R{i:02d}" for i in range(1, len(exceptions) + 1)])
    return exceptions, pd.DataFrame(clearance_rows)


# ---------------------------------------------------------------------------
# 6. Reconciliation statement
# ---------------------------------------------------------------------------
def build_summary(gl, bank, exceptions, config) -> dict:
    book_close = round(config["opening_balance_book"] + gl["amount"].sum(), 2)
    bank_close = round(bank["balance"].iloc[-1], 2)
    # Integrity check: bank running balance agrees to opening + movements
    bank_calc = round(config["opening_balance_bank"] + bank["amount"].sum(), 2)

    bank_adj = exceptions.loc[exceptions["adjust_side"] == BANK_SIDE, "recon_effect"].sum()
    book_adj = exceptions.loc[exceptions["adjust_side"] == BOOK_SIDE, "recon_effect"].sum()
    adj_bank = round(bank_close + bank_adj, 2)
    adj_book = round(book_close + book_adj, 2)

    return {
        "bank_closing": bank_close,
        "bank_closing_recalculated": bank_calc,
        "book_closing": book_close,
        "bank_side_adjustments": round(bank_adj, 2),
        "book_side_adjustments": round(book_adj, 2),
        "adjusted_bank": adj_bank,
        "adjusted_book": adj_book,
        "unreconciled_difference": round(adj_bank - adj_book, 2),
        "gl_lines": len(gl), "bank_lines": len(bank),
        "by_category": exceptions.groupby("category")["recon_effect"].agg(["count", "sum"]).round(2).to_dict("index"),
        "high_priority_items": int((exceptions["priority"] == "High").sum()),
    }


def run(data_dir: Path, threshold: float = DEFAULT_THRESHOLD) -> ReconResult:
    data_dir = Path(data_dir)
    config = load_config(data_dir / "config.json")
    gl = load_gl(data_dir / "general_ledger_oct2026.csv")
    bank = load_bank(data_dir / "bank_statement_oct2026.csv")
    next_bank = load_bank(data_dir / "bank_statement_nov2026_1-10.csv")

    duplicates = find_gl_duplicates(gl)
    matches = match_transactions(gl, bank, exclude_gl=duplicates)
    exceptions, clearance = classify(gl, bank, next_bank, matches, duplicates, config, threshold)
    summary = build_summary(gl, bank, exceptions, config)
    summary["matched_pairs"] = len(matches)
    summary["threshold"] = threshold
    return ReconResult(config, gl, bank, next_bank, matches, exceptions, clearance, summary)
