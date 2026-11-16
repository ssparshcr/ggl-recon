#!/usr/bin/env python3
import os
import sys
import argparse

sys.path.append('/Users/sparshmehta/.gemini/antigravity/scratch/python_packages')
sys.path.append('/Users/sparshmehta/.gemini/antigravity/scratch/ggl_recon_app')

from pdf_parser import parse_invoice_pdf, parse_credit_note_pdf, scan_all_data
from engine import ReconEngine
from excel_generator import generate_reconciliation_workbook

EXCEL_PATH = '/Users/sparshmehta/.gemini/antigravity/scratch/GGL_Master_Reconciliation.xlsx'

def fmt_c(val):
    if val is None:
        return '-'
    return f"₹{val:,.2f}"

def print_report(engine):
    monthly = engine.get_monthly_reconciliation()
    print("=" * 115)
    print("   MEHTA IMPEX — GUJARAT GUARDIAN LTD (GGL) RECONCILIATION REPORT (FY 2026–27)")
    print("=" * 115)
    header = f"{'Month':<9} | {'Bills':<5} | {'SQMM Total':<12} | {'Billed (₹)':<14} | {'Expected (₹)':<14} | {'Overbilled':<12} | {'RD Settled':<12} | {'RD Pending':<12} | {'CD Pending':<10} | {'NET OWED':<12}"
    print(header)
    print("-" * 115)
    
    tot_billed = 0.0
    tot_exp = 0.0
    tot_diff = 0.0
    tot_settled = 0.0
    tot_rd_pending = 0.0
    tot_cd_pending = 0.0
    tot_net_owed = 0.0
    
    for m in monthly:
        b_cnt = len(m['invoices'])
        sqmm = f"{m['total_sqmm']:,.1f}"
        billed = fmt_c(m['billed_total'])
        exp = fmt_c(m['actual_total']) if not m['has_unrated_invoices'] else 'Needs Rates'
        diff = fmt_c(m['overbilled_total']) if not m['has_unrated_invoices'] else '-'
        settled = fmt_c(m['rd_settled_total'])
        rd_pend = fmt_c(m['rd_pending'])
        cd_pend = fmt_c(m['cd_pending'])
        net_owed = fmt_c(m['net_pending_from_ggl'])
        
        tot_billed += m['billed_total']
        if not m['has_unrated_invoices']:
            tot_exp += m['actual_total'] or 0.0
            tot_diff += m['overbilled_total'] or 0.0
            tot_rd_pending += m['rd_pending'] or 0.0
        tot_settled += m['rd_settled_total'] or 0.0
        tot_cd_pending += m['cd_pending'] or 0.0
        tot_net_owed += m['net_pending_from_ggl'] or 0.0
        
        print(f"{m['month_key']:<9} | {b_cnt:<5} | {sqmm:<12} | {billed:<14} | {exp:<14} | {diff:<12} | {settled:<12} | {rd_pend:<12} | {cd_pend:<10} | {net_owed:<12}")
        
    print("-" * 115)
    print(f"{'TOTAL':<9} | {' ':5} | {' ':12} | {fmt_c(tot_billed):<14} | {fmt_c(tot_exp):<14} | {fmt_c(tot_diff):<12} | {fmt_c(tot_settled):<12} | {fmt_c(tot_rd_pending):<12} | {fmt_c(tot_cd_pending):<10} | {fmt_c(tot_net_owed):<12}")
    print("=" * 115)
    print(f"Master Excel file: {EXCEL_PATH}\n")

def main():
    parser = argparse.ArgumentParser(description="GGL Rate Difference & Credit Note Tracker")
    parser.add_argument('--report', action='store_true', help="Print monthly reconciliation report")
    parser.add_argument('--scan', type=str, help="Scan root directory for bills & CNs")
    parser.add_argument('--add-bill', type=str, help="Parse and add a GGL invoice PDF")
    parser.add_argument('--add-cn', type=str, help="Parse and add a GGL credit note PDF")
    parser.add_argument('--set-rate', nargs=3, metavar=('INV_NO', 'SR', 'RATE'), help="Set agreed rate for an item")
    parser.add_argument('--export', action='store_true', help="Regenerate master Excel file")
    parser.add_argument('--serve', action='store_true', help="Run the Web Dashboard")

    args = parser.parse_args()
    engine = ReconEngine()

    if args.scan:
        invs, cns = scan_all_data(args.scan)
        for inv in invs:
            engine.add_invoice(inv)
        for cn in cns:
            engine.add_credit_note(cn)
        generate_reconciliation_workbook(engine, EXCEL_PATH)
        print(f"Scanned & loaded {len(invs)} invoices and {len(cns)} credit notes.")
        print_report(engine)
        return

    if args.add_bill:
        inv = parse_invoice_pdf(args.add_bill)
        engine.add_invoice(inv)
        generate_reconciliation_workbook(engine, EXCEL_PATH)
        print(f"Added invoice #{inv['invoice_no']} ({inv['invoice_date']}) with {len(inv['items'])} items.")
        for it in inv['items']:
            print(f"  #{it['sr']} {it['desc']} ({it['mm']}mm, {it['width']}x{it['length']}) Qty:{it['qty']} SQMM:{it['sqmm']} BilledRate:₹{it['b_rate']}")
        return

    if args.add_cn:
        cn = parse_credit_note_pdf(args.add_cn)
        engine.add_credit_note(cn)
        generate_reconciliation_workbook(engine, EXCEL_PATH)
        print(f"Added CN #{cn['cn_no']} ({cn['cn_date']}) Type:{cn['cn_type']} Month:{cn['target_month']} Total:₹{cn['total_amount']:,.2f}")
        return

    if args.set_rate:
        inv_no, sr, rate = args.set_rate
        engine.set_agreed_rate(inv_no, int(sr), float(rate))
        generate_reconciliation_workbook(engine, EXCEL_PATH)
        print(f"Set rate for Inv {inv_no} item #{sr} to ₹{rate}")
        return

    if args.export:
        p = generate_reconciliation_workbook(engine, EXCEL_PATH)
        print(f"Exported to {p}")
        return

    if args.serve:
        import server
        server.run()
        return

    # Default action: print report
    print_report(engine)

if __name__ == '__main__':
    main()
