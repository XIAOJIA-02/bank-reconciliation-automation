"""
Export the reconciliation as an audit-style Excel workpaper.

The Summary sheet is built with live Excel formulas (SUMIFS / COUNTIFS that
point at the Exceptions and raw-data sheets), so a reviewer can trace every
number on the reconciliation statement back to its source line.

Colour convention: blue = hard-coded input, green = link to another sheet,
black = calculation.
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .reconcile import BANK_SIDE, BOOK_SIDE, ReconResult

FONT = "Arial"
NUM = '#,##0.00;(#,##0.00);"-"'
DATE = "dd-mmm-yyyy"
NAVY = "1F3864"

F_BASE = Font(name=FONT, size=10)
F_BOLD = Font(name=FONT, size=10, bold=True)
F_TITLE = Font(name=FONT, size=14, bold=True, color=NAVY)
F_HEAD = Font(name=FONT, size=10, bold=True, color="FFFFFF")
F_INPUT = Font(name=FONT, size=10, color="0000FF")
F_LINK = Font(name=FONT, size=10, color="008000")
F_NOTE = Font(name=FONT, size=9, italic=True, color="595959")
FILL_HEAD = PatternFill("solid", fgColor=NAVY)
FILL_TOTAL = PatternFill("solid", fgColor="D9E1F2")
FILL_SIGN = PatternFill("solid", fgColor="FFF2CC")
THIN = Side(style="thin", color="808080")
TOP = Border(top=THIN)
DOUBLE = Border(top=THIN, bottom=Side(style="double", color="000000"))

# Counter-account used in the proposed adjusting journal entry, by category
AJE_ACCOUNT = {
    "Bank charges not recorded": "Bank charges expense",
    "Interest not recorded": "Interest income",
    "Dishonoured cheque": "Trade receivables - Orchard Trust",
    "Unrecorded direct debit": "Insurance expense (pending support)",
    "Unrecorded cheque": "Accounts payable / expense (pending support)",
    "Unidentified receipt": "Suspense account",
    "Book error - amount mismatch": "Cleaning expense (original posting account)",
    "Book error - duplicate posting": "Software expense (original posting account)",
}

BANK_ORDER = ["Deposit in transit", "Outstanding cheque", "Unpresented payment", "Unverified cheque - outside cheque book"]
BANK_LABEL = {
    "Deposit in transit": "Add: Deposits in transit",
    "Outstanding cheque": "Less: Outstanding cheques",
    "Unpresented payment": "Less: Unpresented payments",
    "Unverified cheque - outside cheque book": "Add: Cheque outside cheque book (provisional)",
}
BOOK_ORDER = ["Interest not recorded", "Unidentified receipt", "Book error - duplicate posting",
              "Book error - amount mismatch", "Bank charges not recorded", "Dishonoured cheque",
              "Unrecorded direct debit", "Unrecorded cheque"]
BOOK_LABEL = {
    "Interest not recorded": "Add: Interest credited by bank",
    "Unidentified receipt": "Add: Unidentified receipt (to suspense)",
    "Book error - duplicate posting": "Add: Reverse duplicate payment posting",
    "Book error - amount mismatch": "Add/(less): Correct transposition error",
    "Bank charges not recorded": "Less: Bank charges",
    "Dishonoured cheque": "Less: Dishonoured customer cheque",
    "Unrecorded direct debit": "Less: Unrecorded GIRO direct debit",
    "Unrecorded cheque": "Less: Unrecorded cheque",
}


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def _table(ws, df: pd.DataFrame, headers: list[str], widths: list[int], money_cols=(), date_cols=(),
           start_row: int = 1):
    for c, (h, w) in enumerate(zip(headers, widths), start=1):
        cell = ws.cell(row=start_row, column=c, value=h)
        cell.font, cell.fill = F_HEAD, FILL_HEAD
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        ws.column_dimensions[get_column_letter(c)].width = w
    for r, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for c, val in enumerate(row, start=1):
            if isinstance(val, pd.Timestamp):
                val = val.to_pydatetime()
            if val is pd.NaT or (isinstance(val, float) and pd.isna(val)):
                val = None
            cell = ws.cell(row=r, column=c, value=val)
            cell.font = F_BASE
            cell.alignment = Alignment(vertical="top", wrap_text=headers[c - 1] in ("Description", "Follow-up action", "Proposed entry"))
            if c in money_cols:
                cell.number_format = NUM
            if c in date_cols:
                cell.number_format = DATE
    last = start_row + len(df)
    ws.auto_filter.ref = f"A{start_row}:{get_column_letter(len(headers))}{max(last, start_row + 1)}"
    ws.freeze_panes = ws.cell(row=start_row + 1, column=1)
    return last


def _raw_sheet(wb, title, df, cols, headers, widths, money, dates):
    ws = wb.create_sheet(title)
    _table(ws, df[cols], headers, widths, money, dates)
    return ws


# ---------------------------------------------------------------------------
# main export
# ---------------------------------------------------------------------------
def export_workpaper(result: ReconResult, path: Path, preparer: str = "[Your name]") -> Path:
    cfg, s = result.config, result.summary
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"

    # ----- raw data sheets (built first so we know their row ranges) -----
    gl_cols = ["date", "je_id", "reference", "description", "debit", "credit"]
    gl = result.gl.copy()
    gl["debit"] = pd.to_numeric(gl["debit"], errors="coerce")
    gl["credit"] = pd.to_numeric(gl["credit"], errors="coerce")
    bank_cols = ["date", "bank_ref", "description", "withdrawal", "deposit", "balance"]
    bank = result.bank.copy()
    nxt = result.next_bank.copy()
    for d in (bank, nxt):
        d["withdrawal"] = pd.to_numeric(d["withdrawal"], errors="coerce")
        d["deposit"] = pd.to_numeric(d["deposit"], errors="coerce")

    # ----- Exceptions -----
    ex = result.exceptions.copy()
    ex_ws = wb.create_sheet("Exceptions")
    ex_cols = ["item_id", "priority", "category", "adjust_side", "source", "source_id", "date", "reference",
               "description", "amount", "recon_effect", "proposed_entry", "follow_up_action", "cleared_on"]
    ex_head = ["Item", "Priority", "Category", "Adjust side", "Source", "Source ID", "Date", "Reference",
               "Description", "Amount (SGD)", "Effect on recon (SGD)", "Proposed entry", "Follow-up action",
               "Cleared in Nov"]
    ex_last = _table(ex_ws, ex[ex_cols], ex_head, [6, 9, 28, 10, 10, 12, 12, 13, 42, 13, 13, 40, 60, 13],
                     money_cols=(10, 11), date_cols=(7, 14))
    ex_ws.conditional_formatting.add(f"B2:B{ex_last}", CellIsRule(operator="equal", formula=['"High"'],
                                     fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, bold=True, color="9C0006")))
    ex_ws.conditional_formatting.add(f"B2:B{ex_last}", CellIsRule(operator="equal", formula=['"Medium"'],
                                     fill=PatternFill("solid", fgColor="FFE699")))
    CAT = f"Exceptions!$C$2:$C${ex_last}"
    SIDE = f"Exceptions!$D$2:$D${ex_last}"
    EFF = f"Exceptions!$K$2:$K${ex_last}"
    PRI = f"Exceptions!$B$2:$B${ex_last}"

    # ----- Proposed AJEs -----
    aje_ws = wb.create_sheet("Proposed AJEs")
    aje_rows = []
    for _, e in ex[ex["adjust_side"] == BOOK_SIDE].iterrows():
        amt, acct = abs(e["recon_effect"]), AJE_ACCOUNT.get(e["category"], "To be determined")
        dr, cr = ("Cash at bank", acct) if e["recon_effect"] > 0 else (acct, "Cash at bank")
        aje_rows.append([e["item_id"], dr, amt, None, e["category"]])
        aje_rows.append([e["item_id"], f"   {cr}", None, amt, e["follow_up_action"]])
    aje_df = pd.DataFrame(aje_rows, columns=["item", "account", "dr", "cr", "note"])
    aje_last = _table(aje_ws, aje_df, ["Item", "Account", "Debit (SGD)", "Credit (SGD)", "Narrative / support required"],
                      [7, 44, 14, 14, 70], money_cols=(3, 4))
    t = aje_last + 1
    aje_ws.cell(row=t, column=2, value="Total").font = F_BOLD
    for col in ("C", "D"):
        c = aje_ws[f"{col}{t}"]
        c.value, c.font, c.number_format, c.border = f"=SUM({col}2:{col}{aje_last})", F_BOLD, NUM, DOUBLE
    aje_ws.cell(row=t + 1, column=2, value="Debits = credits?").font = F_BASE
    aje_ws[f"C{t+1}"] = f'=IF(ABS(C{t}-D{t})<0.005,"Balanced","NOT BALANCED")'
    aje_ws.cell(row=t + 2, column=2, value="Net effect on Cash at bank").font = F_BASE
    aje_ws[f"C{t+2}"] = (f'=SUMIF(B2:B{aje_last},"Cash at bank",C2:C{aje_last})'
                         f'-SUMIF(B2:B{aje_last},"   Cash at bank",D2:D{aje_last})')
    aje_ws[f"C{t+2}"].number_format = NUM
    aje_ws.cell(row=t + 3, column=2, value="Proposed entries only - require client approval before posting.").font = F_NOTE
    AJE_NET = f"'Proposed AJEs'!$C${t+2}"

    # ----- Subsequent clearance -----
    cl_ws = wb.create_sheet("Subsequent Clearance")
    cl_cols = ["je_id", "reference", "description", "gl_date", "amount", "category", "cleared_on",
               "cleared_bank_ref", "days_after_period_end", "status", "evidence_basis"]
    cl_last = _table(cl_ws, result.clearance[cl_cols],
                     ["JE ID", "Reference", "Description", "GL date", "Amount (SGD)", "Category", "Cleared on",
                      "Nov bank ref", "Days after period end", "Status", "Evidence basis"],
                     [9, 13, 42, 12, 13, 20, 12, 13, 11, 13, 34], money_cols=(5,), date_cols=(4, 7))
    cl_ws.conditional_formatting.add(f"J2:J{cl_last}", CellIsRule(operator="equal", formula=['"NOT CLEARED"'],
                                     fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(name=FONT, bold=True, color="9C0006")))
    cl_ws.cell(row=cl_last + 2, column=1, value=(
        "Procedure: month-end timing items were traced to the 1-10 Nov 2026 bank statement. Cheques are matched on "
        "cheque number; receipts on amount only, which is preliminary evidence to be agreed to deposit slips. "
        "Receipts banked > 5 days after period end are flagged as a cut-off risk.")).font = F_NOTE

    # ----- Matched -----
    m_ws = wb.create_sheet("Matched")
    m_cols = ["je_id", "gl_date", "gl_reference", "gl_description", "gl_amount", "bank_ref", "bank_date",
              "bank_description", "bank_amount", "difference", "days_apart", "match_method"]
    _table(m_ws, result.matches[m_cols],
           ["JE ID", "GL date", "GL ref", "GL description", "GL amount", "Bank ref", "Bank date",
            "Bank description", "Bank amount", "Difference", "Days apart", "Match method"],
           [9, 12, 13, 40, 13, 13, 12, 40, 13, 11, 8, 34], money_cols=(5, 9, 10), date_cols=(2, 7))

    # ----- Raw data -----
    _raw_sheet(wb, "GL_Oct", gl, gl_cols, ["Date", "JE ID", "Reference", "Description", "Debit (SGD)", "Credit (SGD)"],
               [12, 9, 13, 46, 14, 14], (5, 6), (1,))
    gl_last = len(gl) + 1
    bank_ws = _raw_sheet(wb, "Bank_Oct", bank, bank_cols,
                         ["Date", "Bank ref", "Description", "Withdrawal (SGD)", "Deposit (SGD)", "Balance (SGD)"],
                         [12, 13, 46, 14, 14, 15], (4, 5, 6), (1,))
    bank_last = len(bank) + 1
    _raw_sheet(wb, "Bank_Nov", nxt, bank_cols,
               ["Date", "Bank ref", "Description", "Withdrawal (SGD)", "Deposit (SGD)", "Balance (SGD)"],
               [12, 13, 46, 14, 14, 15], (4, 5, 6), (1,))

    # =====================================================================
    # Summary sheet
    # =====================================================================
    ws.sheet_view.showGridLines = False
    for col, w in zip("ABCDEF", [3, 50, 18, 4, 30, 18]):
        ws.column_dimensions[col].width = w

    ws["B1"] = "Bank Reconciliation Workpaper"; ws["B1"].font = F_TITLE
    ws["B2"] = "SYNTHETIC DATA - fictional entity created for a portfolio project"; ws["B2"].font = F_NOTE

    meta = [("Client", cfg["entity"]), ("Account", f"{cfg['account']} ({cfg['currency']})"),
            ("Period end", pd.Timestamp(cfg["period_end"]).to_pydatetime()),
            ("Prepared by / date", f"{preparer}"), ("Reviewed by / date", "")]
    for i, (k, v) in enumerate(meta, start=4):
        ws[f"B{i}"], ws[f"C{i}"] = k, v
        ws[f"B{i}"].font = F_BOLD
        ws[f"C{i}"].font = F_BASE
        if k == "Period end":
            ws[f"C{i}"].number_format = DATE
            ws[f"C{i}"].alignment = Alignment(horizontal="left")
        if k.startswith(("Prepared", "Reviewed")):
            ws[f"C{i}"].fill = FILL_SIGN

    ws["B10"] = "Objective"; ws["B10"].font = F_BOLD
    ws["B11"] = ("Reconcile the cash book to the bank statement at period end, identify and evaluate "
                 "reconciling items, and test timing items for subsequent clearance (existence, completeness, cut-off).")
    ws.merge_cells("B11:F11"); ws["B11"].alignment = Alignment(wrap_text=True, vertical="top")
    ws["B11"].font = F_BASE; ws.row_dimensions[11].height = 28

    r = 13
    def hdr(text):
        nonlocal r
        for col in "BC":
            ws[f"{col}{r}"].fill = FILL_HEAD
        ws[f"B{r}"], ws[f"C{r}"] = text, "SGD"
        ws[f"B{r}"].font = ws[f"C{r}"].font = F_HEAD
        ws[f"C{r}"].alignment = Alignment(horizontal="right")
        r += 1

    def line(label, formula, font=F_BASE, border=None, fill=None):
        nonlocal r
        ws[f"B{r}"], ws[f"C{r}"] = label, formula
        ws[f"B{r}"].font = F_BOLD if font is F_BOLD else F_BASE
        ws[f"C{r}"].font, ws[f"C{r}"].number_format = font, NUM
        if border:
            ws[f"C{r}"].border = border
        if fill:
            ws[f"B{r}"].fill = ws[f"C{r}"].fill = fill
        r += 1
        return r - 1

    # Bank side
    hdr("Reconciliation statement - bank side")
    bank_start = line("Balance per bank statement", f"=Bank_Oct!F{bank_last}", font=F_LINK)
    bank_cats = [c for c in BANK_ORDER if c in set(ex.loc[ex.adjust_side == BANK_SIDE, "category"])]
    for c in bank_cats:
        line(BANK_LABEL[c], f'=SUMIFS({EFF},{CAT},"{c}",{SIDE},"{BANK_SIDE}")')
    adj_bank = line("Adjusted bank balance", f"=SUM(C{bank_start}:C{r-1})", F_BOLD, DOUBLE, FILL_TOTAL)
    r += 1

    # Book side
    hdr("Reconciliation statement - book side")
    ws[f"B{r}"], ws[f"C{r}"] = "Opening cash book balance (1 Oct 2026)", cfg["opening_balance_book"]
    ws[f"B{r}"].font, ws[f"C{r}"].font, ws[f"C{r}"].number_format = F_BASE, F_INPUT, NUM
    open_row = r; r += 1
    line("Add: October debits per GL", f"=SUM(GL_Oct!E2:E{gl_last})", font=F_LINK)
    line("Less: October credits per GL", f"=-SUM(GL_Oct!F2:F{gl_last})", font=F_LINK)
    book_close = line("Balance per cash book", f"=SUM(C{open_row}:C{r-1})", F_BOLD, TOP)
    adj_start = r
    book_cats = [c for c in BOOK_ORDER if c in set(ex.loc[ex.adjust_side == BOOK_SIDE, "category"])]
    for c in book_cats:
        line(BOOK_LABEL[c], f'=SUMIFS({EFF},{CAT},"{c}",{SIDE},"{BOOK_SIDE}")')
    adj_book = line("Adjusted book balance", f"=C{book_close}+SUM(C{adj_start}:C{r-1})", F_BOLD, DOUBLE, FILL_TOTAL)
    r += 1

    # Proof
    hdr("Proof & integrity checks")
    diff = line("Unreconciled difference (adjusted bank - adjusted book)", f"=C{adj_bank}-C{adj_book}", F_BOLD)
    ws[f"B{r}"], ws[f"C{r}"] = "Arithmetic reconciliation", f'=IF(ABS(C{diff})<0.005,"RECONCILED","INVESTIGATE")'
    ws[f"B{r}"].font = F_BASE; ws[f"C{r}"].font = F_BOLD; ws[f"C{r}"].alignment = Alignment(horizontal="right")
    status = r; r += 1
    ws[f"B{r}"], ws[f"C{r}"] = "Open items requiring evidence (High priority)", f'=COUNTIF({PRI},"High")'
    ws[f"B{r}"].font = F_BASE; ws[f"C{r}"].font = F_BOLD; ws[f"C{r}"].alignment = Alignment(horizontal="right")
    open_r = r; r += 1
    ws[f"B{r}"], ws[f"C{r}"] = "Audit status", f'=IF(C{open_r}>0,"OPEN ITEMS","CONCLUDED")'
    ws[f"B{r}"].font = F_BASE; ws[f"C{r}"].font = F_BOLD; ws[f"C{r}"].alignment = Alignment(horizontal="right")
    audit_r = r; r += 1
    ws[f"B{r}"], ws[f"C{r}"] = "Opening bank balance per statement (1 Oct 2026)", cfg["opening_balance_bank"]
    ws[f"B{r}"].font, ws[f"C{r}"].font, ws[f"C{r}"].number_format = F_BASE, F_INPUT, NUM
    bank_open = r; r += 1
    ws[f"B{r}"] = "Bank statement: opening + movements = closing balance?"
    ws[f"C{r}"] = (f'=IF(ABS(C{bank_open}+SUM(Bank_Oct!E2:E{bank_last})'
                   f'-SUM(Bank_Oct!D2:D{bank_last})-Bank_Oct!F{bank_last})<0.005,"OK","CHECK")')
    ws[f"B{r}"].font = F_BASE; ws[f"C{r}"].font = F_BOLD; ws[f"C{r}"].alignment = Alignment(horizontal="right"); r += 1
    ws[f"B{r}"] = "Book-side adjustments agree to proposed AJEs?"
    ws[f"C{r}"] = f'=IF(ABS(C{adj_book}-C{book_close}-{AJE_NET})<0.005,"OK","CHECK")'
    ws[f"B{r}"].font = F_BASE; ws[f"C{r}"].font = F_BOLD; ws[f"C{r}"].alignment = Alignment(horizontal="right"); r += 1
    line("Adjusted book balance after AJEs = adjusted bank balance", f"=C{book_close}+{AJE_NET}")

    green = PatternFill("solid", fgColor="C6EFCE"); red = PatternFill("solid", fgColor="FFC7CE")
    ws.conditional_formatting.add(f"C{status}:C{r}", CellIsRule(operator="equal", formula=['"RECONCILED"'], fill=green))
    ws.conditional_formatting.add(f"C{status}:C{r}", CellIsRule(operator="equal", formula=['"OK"'], fill=green))
    ws.conditional_formatting.add(f"C{status}:C{r}", FormulaRule(formula=[f'OR(C{status}="INVESTIGATE",C{status}="CHECK")'], fill=red))
    ws.conditional_formatting.add(f"C{audit_r}", CellIsRule(operator="equal", formula=['"OPEN ITEMS"'], fill=PatternFill("solid", fgColor="FFE699")))

    # Exceptions overview (right-hand panel)
    p = 13
    for col in "EF":
        ws[f"{col}{p}"].fill = FILL_HEAD
    ws[f"E{p}"], ws[f"F{p}"] = "Exception overview", "Count"
    ws[f"E{p}"].font = ws[f"F{p}"].font = F_HEAD
    ws[f"F{p}"].alignment = Alignment(horizontal="right")
    p += 1
    for label, f in [("Matched pairs", f"=COUNTA(Matched!A2:A{len(result.matches)+1})"),
                     ("Reconciling items", f"=COUNTA(Exceptions!A2:A{ex_last})"),
                     ("  High priority", f'=COUNTIF({PRI},"High")'),
                     ("  Medium priority", f'=COUNTIF({PRI},"Medium")'),
                     ("  Low priority", f'=COUNTIF({PRI},"Low")'),
                     ("Timing items not cleared by 10 Nov",
                      f'=COUNTIF(\'Subsequent Clearance\'!J2:J{cl_last},"NOT CLEARED")')]:
        ws[f"E{p}"], ws[f"F{p}"] = label, f
        ws[f"E{p}"].font = ws[f"F{p}"].font = F_BASE
        p += 1

    p += 1
    ws[f"E{p}"] = "Items for senior / manager attention"; ws[f"E{p}"].font = F_BOLD; p += 1
    for _, e in ex[ex.priority == "High"].iterrows():
        ws[f"E{p}"] = f"{e['item_id']} {e['category']}: {e['amount']:,.2f}"
        ws[f"E{p}"].font = F_BASE
        ws.merge_cells(f"E{p}:F{p}")
        p += 1

    r = max(r, p) + 1
    ws[f"B{r}"] = "Conclusion"; ws[f"B{r}"].font = F_BOLD; r += 1
    ws[f"B{r}"] = ("The cash book reconciles arithmetically to the bank statement with nil unreconciled difference. "
                   "This is not yet an audit conclusion: the high-priority items listed above remain open until supporting "
                   "evidence is obtained. Book-side items are proposed as adjusting journal entries (see 'Proposed AJEs'). "
                   "The cheque outside the entity's cheque book is a provisional bank-side item pending the bank's response, "
                   "and timing items matched on amount only are preliminary until agreed to deposit slips.")
    ws.merge_cells(f"B{r}:F{r}"); ws[f"B{r}"].alignment = Alignment(wrap_text=True, vertical="top")
    ws[f"B{r}"].font = F_BASE; ws.row_dimensions[r].height = 54
    r += 2
    ws[f"B{r}"] = "Legend: blue = input, green = link to source sheet, black = formula. Yellow = sign-off."
    ws[f"B{r}"].font = F_NOTE

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path
