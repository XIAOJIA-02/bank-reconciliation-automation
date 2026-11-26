"""
Run the bank reconciliation end to end.

    python main.py                       # uses ./data, writes ./output
    python main.py --threshold 10000     # change the escalation threshold
    python main.py --preparer "Your Name"
"""

import argparse
from pathlib import Path

from bankrec.reconcile import DEFAULT_THRESHOLD, run
from bankrec.workpaper import export_workpaper


def main():
    ap = argparse.ArgumentParser(description="Automated bank reconciliation with audit-style workpaper output")
    ap.add_argument("--data-dir", default="data", help="folder with the GL, bank statements and config.json")
    ap.add_argument("--output-dir", default="output", help="where to write the workpaper and CSV reports")
    ap.add_argument("--threshold", type=float, default=DEFAULT_THRESHOLD,
                    help="items at or above this amount are escalated (default 5,000)")
    ap.add_argument("--preparer", default="[Your name]", help="name shown in the workpaper sign-off box")
    args = ap.parse_args()

    result = run(Path(args.data_dir), threshold=args.threshold)
    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)

    wp = export_workpaper(result, out / "bank_reconciliation_workpaper_oct2026.xlsx", preparer=args.preparer)
    result.exceptions.to_csv(out / "reconciling_items.csv", index=False)
    result.matches.to_csv(out / "matched_transactions.csv", index=False)

    s = result.summary
    print("=" * 64)
    print(f"Bank reconciliation - {result.config['entity']}")
    print(f"Period end {result.config['period_end']}  |  threshold SGD {s['threshold']:,.0f}")
    print("=" * 64)
    print(f"GL lines {s['gl_lines']}, bank lines {s['bank_lines']}, matched pairs {s['matched_pairs']}")
    print(f"Balance per bank statement   {s['bank_closing']:>14,.2f}")
    print(f"  bank-side adjustments      {s['bank_side_adjustments']:>14,.2f}")
    print(f"Adjusted bank balance        {s['adjusted_bank']:>14,.2f}")
    print(f"Balance per cash book        {s['book_closing']:>14,.2f}")
    print(f"  book-side adjustments      {s['book_side_adjustments']:>14,.2f}")
    print(f"Adjusted book balance        {s['adjusted_book']:>14,.2f}")
    print(f"Unreconciled difference      {s['unreconciled_difference']:>14,.2f}")
    print("-" * 64)
    for cat, v in s["by_category"].items():
        print(f"  {cat:<36}{int(v['count']):>3}  {v['sum']:>12,.2f}")
    print("-" * 64)
    print(f"High-priority items: {s['high_priority_items']}")
    status = "RECONCILED" if abs(s["unreconciled_difference"]) < 0.005 else "DIFFERENCE - INVESTIGATE"
    print(f"Arithmetic status: {status}")
    audit = "OPEN ITEMS - evidence required" if s["open_items"] else "CONCLUDED"
    print(f"Audit status: {audit} ({s['open_items']} open)")
    print(f"Workpaper: {wp}")


if __name__ == "__main__":
    main()
