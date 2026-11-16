import os
import re
import glob
from datetime import datetime

try:
    import pypdf
except ImportError:
    import sys
    sys.path.append('/Users/sparshmehta/.gemini/antigravity/scratch/python_packages')
    import pypdf

MONTH_MAP = {
    'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
    'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
    'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
}

def normalize_date(date_str):
    """Normalize date strings like '26 Apr, 2026' or '2026-04-26' to YYYY-MM-DD."""
    if not date_str:
        return ''
    date_str = date_str.strip()
    m = re.search(r'(\d{1,2})\s*([A-Za-z]{3}),?\s*(\d{4})', date_str)
    if m:
        day = int(m.group(1))
        mon = MONTH_MAP.get(m.group(2).lower()[:3], '01')
        yr = m.group(3)
        return f'{yr}-{mon}-{day:02d}'
    m2 = re.search(r'(\d{4})[/-](\d{1,2})[/-](\d{1,2})', date_str)
    if m2:
        return f'{m2.group(1)}-{int(m2.group(2)):02d}-{int(m2.group(3)):02d}'
    return date_str

def parse_invoice_pdf(pdf_path):
    """Parse a GGL Tax Invoice PDF."""
    reader = pypdf.PdfReader(pdf_path)
    full_text = '\n'.join([p.extract_text() for p in reader.pages])
    
    # 1. Invoice Number & Date
    inv_m = re.search(r'TAX INVOICE\s*:\s*(\w+)\s+INVOICE DATE\s*:\s*(\d{1,2}\s+[A-Za-z]{3},?\s+\d{4})', full_text)
    if not inv_m:
        inv_m = re.search(r'TAX INVOICE\s*:\s*(\w+)\s+INVOICE DATE\s*:\s*([^\n\r]+?)(?:\s+ORIGINAL|\s+DUPLICATE|\s+TRIPLICATE|$)', full_text)
    inv_no = inv_m.group(1) if inv_m else ''
    inv_date_raw = inv_m.group(2).strip() if inv_m else ''
    inv_date = normalize_date(inv_date_raw)
    
    # Month Key e.g. '2026-04'
    month_key = inv_date[:7] if len(inv_date) >= 7 else ''
    
    # Customer and dispatch info
    truck_m = re.search(r'Truck No\.\s*:\s*([^\n]+)', full_text)
    truck_no = truck_m.group(1).strip() if truck_m else ''
    lr_m = re.search(r'LR No\.\s*:\s*([^\n]+)', full_text)
    lr_no = lr_m.group(1).strip() if lr_m else ''
    
    # Surcharges
    fe_m = re.search(r'FUEL AND ENERGY SURCHARGE.*?([\d,]+)\s*$', full_text, re.MULTILINE)
    pms_m = re.search(r'PRECIOUS METAL SURCHARGE.*?([\d,]+)\s*$', full_text, re.MULTILINE)
    ins_m = re.search(r'DOMESTIC INSURANCE CHARGES.*?([\d,]+)\s*$', full_text, re.MULTILINE)
    tot_m = re.search(r'Totals.*?([\d,]+)\s*$', full_text, re.MULTILINE)
    
    fe = float(fe_m.group(1).replace(',', '')) if fe_m else 0.0
    pms = float(pms_m.group(1).replace(',', '')) if pms_m else 0.0
    ins = float(ins_m.group(1).replace(',', '')) if ins_m else 0.0
    
    # Subtotal and GST
    subtot_m = re.search(r'Invoice Subtotal\s+([\d,]+)', full_text)
    cgst_m = re.search(r'CGST@9%\s+([\d,]+)', full_text)
    sgst_m = re.search(r'SGST@9%\s+([\d,]+)', full_text)
    
    subtotal = float(subtot_m.group(1).replace(',', '')) if subtot_m else 0.0
    cgst = float(cgst_m.group(1).replace(',', '')) if cgst_m else 0.0
    sgst = float(sgst_m.group(1).replace(',', '')) if sgst_m else 0.0
    
    if tot_m:
        total = float(tot_m.group(1).replace(',', ''))
    else:
        total = round(subtotal + cgst + sgst, 2)
        
    # Line items
    lines = full_text.splitlines()
    items = []
    i = 0
    while i < len(lines):
        line = lines[i].strip()
        m = re.match(r'^(\d+)\s+(700\d{5})\s+(\d+(?:\.\d+)?)\s*mm\s*/\s*(\d+)[Xx](\d+)', line)
        if m:
            item_sr = int(m.group(1))
            hsn = m.group(2)
            thickness = float(m.group(3))
            width = int(m.group(4))
            length = int(m.group(5))
            
            i += 1
            desc = lines[i].strip() if i < len(lines) else ''
            
            cases = 0
            sqm = 0.0
            price_sqm = 0.0
            amount = 0.0
            qty = 0
            
            while i < len(lines) - 1:
                i += 1
                curr = lines[i].strip()
                if re.match(r'^\d+\s+700\d{5}', curr) or 'FUEL AND ENERGY' in curr or 'Invoice Subtotal' in curr:
                    i -= 1
                    break
                
                # Check for sqm line
                sqm_m = re.match(r'^(\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+([\d,]+)', curr)
                if sqm_m:
                    cases = int(sqm_m.group(1))
                    sqm = float(sqm_m.group(2))
                    price_sqm = float(sqm_m.group(3))
                    amount = float(sqm_m.group(4).replace(',', ''))
                
                # Check for container / qty line: e.g. '10006626 - 65.000'
                qty_m = re.search(r'-\s*(\d+(?:\.\d+)?)\s*$', curr)
                if qty_m and not sqm_m:
                    qty += int(float(qty_m.group(1)))
            
            if qty == 0 and sqm > 0 and width > 0 and length > 0:
                sheet_sqm = (width * length) / 1000000.0
                qty = int(round(sqm / sheet_sqm))
                
            sqmm = round(sqm * thickness, 4)
            b_rate = round(price_sqm / thickness, 2)
            item_code = 'MIR' if 'Mirror' in desc else 'CLR'
            
            items.append({
                'sr': item_sr,
                'hsn': hsn,
                'item_type': item_code,
                'desc': desc,
                'mm': thickness,
                'width': width,
                'length': length,
                'qty': qty,
                'sqm': sqm,
                'sqmm': sqmm,
                'price_sqm': price_sqm,
                'b_rate': b_rate,
                'billed_amount': amount,
                'agreed_rate': None  # To be input by user
            })
        i += 1

    return {
        'invoice_no': inv_no,
        'invoice_date': inv_date,
        'month_key': month_key,
        'truck_no': truck_no,
        'lr_no': lr_no,
        'fe': fe,
        'pms': pms,
        'ins': ins,
        'subtotal': subtotal,
        'cgst': cgst,
        'sgst': sgst,
        'total_amount': total,
        'items': items,
        'source_file': os.path.basename(pdf_path),
        'file_path': pdf_path
    }

def parse_credit_note_pdf(pdf_path):
    """Parse a GGL Credit Note PDF."""
    reader = pypdf.PdfReader(pdf_path)
    full_text = '\n'.join([p.extract_text() for p in reader.pages])
    
    # 1. CN Number and Date
    cn_m = re.search(r'(?:Invoice no\s+Invoice date|CN No\.\s+CN Date)\s*\n\s*(\d{10,14})\s+(\d{1,2}\s+[A-Za-z]{3},?\s+\d{4})', full_text)
    if not cn_m:
        cn_m = re.search(r'(\d{13})\s+(\d{1,2}\s+[A-Za-z]{3},?\s+\d{4})', full_text)
    if not cn_m:
        cn_m = re.search(r'(?:Invoice no\s+Invoice date|CN No\.\s+CN Date)\s*\n\s*(\d+)\s+([^\n]+)', full_text)
        
    cn_no = cn_m.group(1) if cn_m else ''
    cn_date_raw = cn_m.group(2).strip() if cn_m else ''
    cn_date = normalize_date(cn_date_raw)
    
    # 2. Total
    tot_m = re.search(r'Total\s*:\s*([\d,]+)', full_text)
    total_amount = float(tot_m.group(1).replace(',', '')) if tot_m else 0.0
    
    # 3. Description & Basic Amount
    desc_m = re.search(r'DESCRIPTION\s+AMOUNT\s*\n\s*([^\n]+)', full_text)
    desc_line = desc_m.group(1).strip() if desc_m else ''
    
    amt_m = re.search(r'([\d,]+(?:\.\d+)?)\s*$', desc_line)
    basic_amount = float(amt_m.group(1).replace(',', '')) if amt_m else round(total_amount / 1.18, 2)
    clean_desc = re.sub(r'[\d,]+(?:\.\d+)?\s*$', '', desc_line).strip()
    
    lines = full_text.splitlines()
    desc_start = False
    extra_lines = []
    for l in lines:
        if 'DESCRIPTION AMOUNT' in l or ('DESCRIPTION' in l and 'AMOUNT' in l):
            desc_start = True
            continue
        if desc_start:
            if any(k in l for k in ['SGST', 'CGST', 'IGST', 'TOEMAIL', 'PRINTER', 'Page', 'Total :']):
                break
            line_s = l.strip()
            if line_s and not any(k in line_s for k in ['BILL TO', 'Customer Order', 'Place of Supply', 'PAN Number', 'GSTN No', 'MEHTA IMPEX', 'VADODARA']):
                if line_s != clean_desc and line_s != desc_line:
                    extra_lines.append(line_s)
    if extra_lines:
        clean_desc += ' | ' + ' '.join(extra_lines)
    
    # 4. Accurate Classification (Quarterly/Annual Tonnage Incentives vs Monthly Invoice RD)
    full_upper = (full_text + ' ' + clean_desc).upper()
    is_cd = any(k in full_upper for k in ['CASH DISCOUNT', 'CDCN', 'C D '])
    is_qtr_tonnage = any(k in full_upper for k in ['QTR', 'QTRLY', 'QUARTERLY', 'JULTOSEP', 'JUL TO SEP', 'JAN-MAR', 'TONNAGE', 'ANNUAL'])
    is_fe = any(k in full_upper for k in ['F&E', 'ENERGY', 'EDIFF'])
    is_pms = any(k in full_upper for k in ['PRECIOUS METAL', 'PMS'])

    if is_cd:
        cn_type = 'CD'
    elif is_qtr_tonnage:
        cn_type = 'TONNAGE'
    elif is_fe:
        cn_type = 'FE_ADJ'
    elif is_pms:
        cn_type = 'PMS_RET'
    elif 'ADDITIONAL RATE DIFFERENCE' in full_upper or 'ADD. R D' in full_upper or 'ARD' in full_upper:
        cn_type = 'ARD'
    elif 'RATE DIFF' in full_upper or 'RATE DIFFERENCE' in full_upper or ' R D' in full_upper:
        cn_type = 'RD'
    elif 'DAMAGE' in full_upper:
        cn_type = 'DAMAGE'
    else:
        cn_type = 'OTHER'

    # 5. Target Month & Earned Period Extraction
    # If clean_desc has a '|', the part after '|' has the specific line-item reference (e.g. "ARD JUL26")
    # while the part before '|' is just GGL's boilerplate header (e.g. "Additional Rate Difference for Aug-2026").
    parts = clean_desc.split('|', 1)
    search_texts = [parts[1].upper(), parts[0].upper()] if len(parts) > 1 else [clean_desc.upper()]

    earned_period = ''
    for text in search_texts:
        if 'JAN-MAR' in text or 'JAN-MAR26' in text:
            earned_period = 'FY25-26-Q4'; break
        elif 'APR-JUN' in text or 'APRTOJUN' in text:
            earned_period = '2026-Q1'; break
        elif 'JULTOSEP' in text or 'JUL-SEP' in text or 'JUL TO SEP' in text:
            earned_period = '2026-Q2'; break
        elif any(k in text for k in ['OCT26', 'OCT 26', 'OCT-2026', 'OCTOBER 26', 'OCTOBER-2026']):
            earned_period = '2026-10'; break
        elif any(k in text for k in ['SEP26', 'SEP 26', 'SEP-2026', 'SEPT 26', 'SEPT-2026']):
            earned_period = '2026-09'; break
        elif any(k in text for k in ['AUG26', 'AUG 26', 'AUG-2026', 'AUGUST 26']):
            earned_period = '2026-08'; break
        elif any(k in text for k in ['JUL26', 'JUL -2026', 'JULY 26', 'JUL-2026', 'JULY-2026', 'MONTHLY JUL']):
            earned_period = '2026-07'; break
        elif any(k in text for k in ['JUN26', 'JUN 26', 'JUNE 26', 'JUN-2026', 'JUNE-2026', 'JUN 2', 'JUN2']):
            earned_period = '2026-06'; break
        elif any(k in text for k in ['MAY26', 'MAY 26', 'MAY-2026']):
            earned_period = '2026-05'; break
        elif any(k in text for k in ['APR26', 'APR 26', 'APR-2026', 'APRIL 26', 'APRIL-2026']):
            earned_period = '2026-04'; break

    target_month = ''
    # Check invoice reference e.g. "F&E200001303" or "EDIFF200002138" or "2026200004977"
    ref_inv = ''
    ref_m = re.search(r'(?:F&E|EDIFF|INV)?(20\d{7,11})', clean_desc)
    if ref_m:
        raw_ref = ref_m.group(1)
        # Normalize to 13-digit invoice number if known
        known_inv_months = {
            '2026200000536': '2026-04', '200000536': '2026-04',
            '2026200001303': '2026-05', '200001303': '2026-05',
            '2026200002138': '2026-06', '200002138': '2026-06',
            '2026200003174': '2026-07', '200003174': '2026-07',
            '2026200003536': '2026-07', '200003536': '2026-07',
            '2026200004977': '2026-09', '200004977': '2026-09',
        }
        for k_inv, k_month in known_inv_months.items():
            if raw_ref in k_inv or k_inv in raw_ref:
                ref_inv = k_inv if len(k_inv) == 13 else f"2026{k_inv}"
                target_month = k_month
                earned_period = k_month
                break
        if not ref_inv:
            ref_inv = raw_ref

    # If target_month still empty, use earned_period if it represents a specific month (YYYY-MM)
    if not target_month and earned_period and re.match(r'^\d{4}-\d{2}$', earned_period):
        target_month = earned_period
    elif not target_month and is_qtr_tonnage and earned_period:
        target_month = earned_period
    elif not target_month:
        m_match = re.search(r'(?:for\s+)([A-Za-z]{3})[- ]?(\d{4}|\d{2})', clean_desc, re.IGNORECASE)
        if not m_match:
            m_match = re.search(r'\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*[- ]?(\d{4}|\d{2})\b', clean_desc, re.IGNORECASE)
        if m_match:
            mon_str = m_match.group(1).lower()[:3]
            yr_str = m_match.group(2)
            if len(yr_str) == 2:
                yr_str = '20' + yr_str
            if mon_str in MONTH_MAP:
                target_month = f'{yr_str}-{MONTH_MAP[mon_str]}'

    # If target_month still empty, default to CN issue month
    if not target_month and cn_date:
        target_month = cn_date[:7]
    if not earned_period:
        earned_period = target_month

    return {
        'cn_no': cn_no,
        'cn_date': cn_date,
        'target_month': target_month,
        'earned_period': earned_period,
        'ref_invoice': ref_inv,
        'cn_type': cn_type,
        'desc': clean_desc,
        'basic_amount': basic_amount,
        'total_amount': total_amount,
        'source_file': os.path.basename(pdf_path),
        'file_path': pdf_path
    }

def parse_ledger_pdf(pdf_path):
    """Parse a GGL Statement of Account / Ledger PDF."""
    reader = pypdf.PdfReader(pdf_path)
    full_text = '\n'.join([p.extract_text() for p in reader.pages])
    
    entries = []
    opening_bal = 0.0
    closing_bal = 0.0
    
    # Opening balance
    ob_m = re.search(r'Opening Balance\s+([-\d,]+\.?\d*)', full_text)
    if ob_m:
        opening_bal = float(ob_m.group(1).replace(',', ''))
        
    # Closing balance
    cb_m = re.search(r'Closing Balance:\s*([-\d,]+\.?\d*)', full_text)
    if cb_m:
        closing_bal = float(cb_m.group(1).replace(',', ''))
        
    lines = full_text.splitlines()
    for line in lines:
        m = re.match(r'^\s*(\d{2}-\d{2}-\d{4})\s+(CNT|INV|PAY)\s+([^\s]+)\s+([\d,]+\.?\d*)\s+([-\d,]+\.?\d*)\s*(.*)', line)
        if m:
            dt, doc_type, doc_no, amt, bal, rem = m.groups()
            dr = float(amt.replace(',', '')) if doc_type == 'INV' else None
            cr = float(amt.replace(',', '')) if doc_type in ['CNT', 'PAY'] else None
            bal_val = float(bal.replace(',', ''))
            entries.append({
                'date': dt,
                'type': doc_type,
                'doc_no': doc_no,
                'dr': dr,
                'cr': cr,
                'balance': bal_val,
                'remarks': rem.strip()
            })
            
    return {
        'opening_balance': opening_bal,
        'closing_balance': closing_bal,
        'entries': entries,
        'source_file': os.path.basename(pdf_path)
    }

def scan_all_data(root_dir):
    """Scans all monthly folders for bills and credit notes."""
    invoices = []
    credit_notes = []
    
    bill_paths = glob.glob(os.path.join(root_dir, '**/BILL/*.pdf'), recursive=True) + \
                 glob.glob(os.path.join(root_dir, '**/BILL/*.PDF'), recursive=True)
    for bp in sorted(bill_paths):
        try:
            inv = parse_invoice_pdf(bp)
            if inv['invoice_no']:
                invoices.append(inv)
        except Exception as e:
            print(f"Error parsing invoice {bp}: {e}")
            
    cn_paths = glob.glob(os.path.join(root_dir, '**/CN/*.pdf'), recursive=True) + \
               glob.glob(os.path.join(root_dir, '**/CN/*.PDF'), recursive=True)
    for cp in sorted(cn_paths):
        try:
            cn = parse_credit_note_pdf(cp)
            if cn['cn_no']:
                credit_notes.append(cn)
        except Exception as e:
            print(f"Error parsing CN {cp}: {e}")
            
    invoices.sort(key=lambda x: x['invoice_date'])
    credit_notes.sort(key=lambda x: x['cn_date'])
    
    return invoices, credit_notes
