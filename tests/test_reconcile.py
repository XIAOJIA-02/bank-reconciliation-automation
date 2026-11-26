"""
Tests: the synthetic data has a known answer key (see scripts/generate_sample_data.py),
so the engine must find every planted item and reconcile to nil.

Run:  python -m pytest -v
"""

from pathlib import Path

import shutil

import pandas as pd
import pytest

from bankrec.reconcile import InputDataError, cheque_no, find_gl_duplicates, load_bank, load_gl, run

DATA = Path(__file__).resolve().parent.parent / "data"


@pytest.fixture(scope="module")
def result():
    return run(DATA)


def items(result, category):
    return result.exceptions[result.exceptions["category"] == category]


def test_reconciles_to_nil(result):
    assert result.summary["unreconciled_difference"] == 0
    assert result.summary["adjusted_bank"] == pytest.approx(271_203.41)


def test_bank_statement_running_balance_is_internally_consistent(result):
    assert result.summary["bank_closing"] == result.summary["bank_closing_recalculated"]


def test_cheque_number_extraction():
    assert cheque_no("CHQ100231") == "100231"
    assert cheque_no("CHEQUE 100299") == "100299"
    assert cheque_no("CHEQUE DEPOSIT") == ""


def test_same_amount_transfers_matched_to_closest_date(result):
    # Two Nimbustech payments of 4,500 on 8 Oct and 22 Oct must not be cross-matched
    m = result.matches.set_index("gl_reference")
    assert m.loc["PV-2610-002", "days_apart"] == 0
    assert m.loc["PV-2610-007", "days_apart"] == 0


def test_transposition_error_detected(result):
    row = items(result, "Book error - amount mismatch")
    assert len(row) == 1
    assert row.iloc[0]["recon_effect"] == pytest.approx(180.00)
    assert "divisible by 9" in row.iloc[0]["follow_up_action"]


def test_duplicate_posting_detected(result):
    assert len(find_gl_duplicates(result.gl)) == 1
    row = items(result, "Book error - duplicate posting")
    assert row.iloc[0]["recon_effect"] == pytest.approx(2_400.00)


def test_timing_differences_and_subsequent_clearance(result):
    dit = items(result, "Deposit in transit")
    assert dit["amount"].sum() == pytest.approx(20_500.00)
    late = dit[dit["amount"] == 8_000.00].iloc[0]
    assert late["priority"] == "High" and "cut-off" in late["follow_up_action"]

    osc = items(result, "Outstanding cheque")
    assert osc["amount"].sum() == pytest.approx(-11_195.00)
    not_cleared = result.clearance[result.clearance["status"] == "NOT CLEARED"]
    assert list(not_cleared["reference"]) == ["CHQ100239"]


def test_bank_only_items_classified(result):
    assert items(result, "Bank charges not recorded")["amount"].sum() == pytest.approx(-70.00)
    assert items(result, "Interest not recorded")["amount"].sum() == pytest.approx(214.66)
    assert items(result, "Dishonoured cheque")["amount"].sum() == pytest.approx(-3_850.00)
    assert items(result, "Unrecorded direct debit")["amount"].sum() == pytest.approx(-960.00)
    assert items(result, "Unidentified receipt")["priority"].iloc[0] == "High"


def test_cheque_outside_cheque_book_is_provisional_and_high_priority(result):
    row = items(result, "Unverified cheque - outside cheque book").iloc[0]
    assert row["adjust_side"] == "Bank"
    assert row["recon_effect"] == pytest.approx(2_750.00)
    assert row["priority"] == "High"
    # It must not be presented as a confirmed bank error
    assert row["follow_up_action"].startswith("PROVISIONAL")
    assert "No entry yet" in row["proposed_entry"]


def test_arithmetic_reconciliation_is_not_an_audit_conclusion(result):
    assert result.summary["unreconciled_difference"] == 0
    assert result.summary["open_items"] == 4   # High items stay open until evidenced


def test_amount_only_clearance_is_marked_preliminary(result):
    c = result.clearance.set_index("reference")
    assert c.loc["CHQ100245", "evidence_basis"] == "Cheque number + amount"
    assert "preliminary" in c.loc["RV-2610-008", "evidence_basis"]


def test_every_line_is_accounted_for(result):
    """Each GL and bank line is either matched or appears as a reconciling item."""
    gl_ids = set(result.matches["je_id"]) | set(result.exceptions.loc[result.exceptions["source"] == "GL", "source_id"])
    bank_ids = set(result.matches["bank_ref"]) | set(result.exceptions.loc[result.exceptions["source"] == "Bank", "source_id"])
    assert gl_ids == set(result.gl["je_id"])
    assert bank_ids == set(result.bank["bank_ref"])


def test_threshold_changes_escalation():
    low = run(DATA, threshold=1_000)
    high = run(DATA, threshold=1_000_000)
    assert low.summary["high_priority_items"] > high.summary["high_priority_items"]
    # threshold never changes the numbers, only the follow-up priority
    assert low.summary["adjusted_bank"] == high.summary["adjusted_bank"]


def test_workpaper_is_created(tmp_path, result):
    from bankrec.workpaper import export_workpaper
    path = export_workpaper(result, tmp_path / "wp.xlsx")
    sheets = pd.ExcelFile(path).sheet_names
    assert sheets[:4] == ["Summary", "Exceptions", "Proposed AJEs", "Subsequent Clearance"]


# ---------------------------------------------------------------------------
# Input validation and edge cases: bad data must stop the run, not slip through
# ---------------------------------------------------------------------------
GL_HEADER = "date,je_id,reference,description,debit,credit\n"
BANK_HEADER = "date,bank_ref,description,withdrawal,deposit,balance\n"


def write(tmp_path, name, text):
    p = tmp_path / name
    p.write_text(text)
    return p


def test_missing_column_is_rejected(tmp_path):
    p = write(tmp_path, "gl.csv", "date,je_id,reference,description,debit\n2026-10-01,JE1,X,Y,10\n")
    with pytest.raises(InputDataError, match="missing column"):
        load_gl(p)


def test_non_numeric_amount_is_rejected(tmp_path):
    p = write(tmp_path, "gl.csv", GL_HEADER + "2026-10-01,JE1,X,Y,12O.00,\n")   # letter O, not zero
    with pytest.raises(InputDataError, match="non-numeric"):
        load_gl(p)


def test_invalid_date_is_rejected(tmp_path):
    p = write(tmp_path, "gl.csv", GL_HEADER + "31/10/2026,JE1,X,Y,10,\n")
    with pytest.raises(InputDataError, match="date"):
        load_gl(p)


def test_line_with_both_debit_and_credit_is_rejected(tmp_path):
    p = write(tmp_path, "gl.csv", GL_HEADER + "2026-10-01,JE1,X,Y,10,10\n")
    with pytest.raises(InputDataError, match="both debit and credit"):
        load_gl(p)


def test_duplicate_bank_reference_is_rejected(tmp_path):
    p = write(tmp_path, "bank.csv", BANK_HEADER +
              "2026-10-01,BK1,A,10,,90\n2026-10-02,BK1,B,10,,80\n")
    with pytest.raises(InputDataError, match="duplicate bank_ref"):
        load_bank(p)


def test_amounts_with_thousand_separators_are_accepted(tmp_path):
    p = write(tmp_path, "gl.csv", GL_HEADER + '2026-10-01,JE1,X,Y,"1,250.00",\n')
    assert load_gl(p)["amount"].iloc[0] == pytest.approx(1_250.00)


def test_ambiguous_same_amount_lines_match_one_to_one(tmp_path):
    """Two identical GL receipts and two identical bank deposits must pair up once each, never twice."""
    d = tmp_path / "data"
    shutil.copytree(DATA, d)
    gl = pd.read_csv(d / "general_ledger_oct2026.csv", dtype=str)
    bank = pd.read_csv(d / "bank_statement_oct2026.csv", dtype=str)
    gl.loc[len(gl)] = ["2026-10-15", "JE9001", "RV-X1", "Extra receipt", "777.00", ""]
    gl.loc[len(gl)] = ["2026-10-15", "JE9002", "RV-X2", "Extra receipt", "777.00", ""]
    last = float(bank["balance"].iloc[-1])
    bank.loc[len(bank)] = ["2026-10-16", "BK-X1", "INWARD TT", "", "777.00", f"{last + 777:.2f}"]
    bank.loc[len(bank)] = ["2026-10-16", "BK-X2", "INWARD TT", "", "777.00", f"{last + 1554:.2f}"]
    gl.to_csv(d / "general_ledger_oct2026.csv", index=False)
    bank.to_csv(d / "bank_statement_oct2026.csv", index=False)

    r = run(d)
    m = r.matches.set_index("je_id")
    assert {m.loc["JE9001", "bank_ref"], m.loc["JE9002", "bank_ref"]} == {"BK-X1", "BK-X2"}
    assert r.matches["je_id"].is_unique and r.matches["bank_ref"].is_unique
    assert r.summary["unreconciled_difference"] == 0
