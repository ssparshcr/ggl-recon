import os
import json
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    import sys
    sys.path.append('/Users/sparshmehta/.gemini/antigravity/scratch/python_packages')
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

NAVY = '1B365D'
LIGHT_NAVY = '2B548F'
HEADER_FILL = PatternFill(start_color=NAVY, end_color=NAVY, fill_type='solid')
SUBHEADER_FILL = PatternFill(start_color=LIGHT_NAVY, end_color=LIGHT_NAVY, fill_type='solid')
ZEBRA_FILL = PatternFill(start_color='F4F7FB', end_color='F4F7FB', fill_type='solid')
PENDING_FILL = PatternFill(start_color='FDE8E8', end_color='FDE8E8', fill_type='solid')
SUCCESS_FILL = PatternFill(start_color='E8F5E9', end_color='E8F5E9', fill_type='solid')
SETTLED_FILL = SUCCESS_FILL
TOTAL_FILL = PatternFill(start_color='E9ECEF', end_color='E9ECEF', fill_type='solid')

FONT_HEADER = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
FONT_SUBHEADER = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
FONT_BOLD = Font(name='Segoe UI', size=10, bold=True)
FONT_REGULAR = Font(name='Segoe UI', size=10)
FONT_TITLE = Font(name='Segoe UI', size=15, bold=True, color=NAVY)
FONT_ALERT = Font(name='Segoe UI', size=10, bold=True, color='C53030')
FONT_OK = Font(name='Segoe UI', size=10, bold=True, color='2F855A')

GRAY_HEADER = PatternFill(start_color='EDF2F7', end_color='EDF2F7', fill_type='solid')
SUMMARY_FILL = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
HIGHLIGHT_FILL = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
FONT_NOTE = Font(name='Segoe UI', size=8.5, italic=True, color='4A5568')
FONT_GREEN = Font(name='Segoe UI', size=9.5, bold=True, color='15803D')
FONT_RED = Font(name='Segoe UI', size=9.5, bold=True, color='B91C1C')

AMBER_FILL = PatternFill(start_color='FEF3C7', end_color='FEF3C7', fill_type='solid')
FONT_AMBER = Font(name='Segoe UI', size=9.5, bold=True, color='B45309')

THIN_BORDER = Border(
    left=Side(style='thin', color='D0D7DE'),
    right=Side(style='thin', color='D0D7DE'),
    top=Side(style='thin', color='D0D7DE'),
    bottom=Side(style='thin', color='D0D7DE')
)

DOUBLE_BOTTOM_BORDER = Border(
    left=Side(style='thin', color='D0D7DE'),
    right=Side(style='thin', color='D0D7DE'),
    top=Side(style='thin', color='D0D7DE'),
    bottom=Side(style='double', color='1B365D')
)

def format_currency_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right', vertical='center')

def format_qty_cell(cell, value=None):
    if value is not None:
        cell.value = value
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal='right', vertical='center')

def auto_fit_columns(ws, min_width=12):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if '\n' in val_str:
                val_str = max(val_str.split('\n'), key=len)
            max_len = max(max_len, len(val_str))
        ws.column_dimensions[col_letter].width = max(max_len + 3, min_width)



def add_historical_main_sheet(wb, engine):
    ws = wb.create_sheet(title='MAIN (Historical Format)')
    ws.title = 'MAIN (Historical Format)'
    ws.views.sheetView[0].showGridLines = True
    
    # Document Header
    ws['A1'] = "MEHTA IMPEX — GUJARAT GUARDIAN LIMITED (GGL)"
    ws['A1'].font = FONT_TITLE
    ws['A2'] = "Historical Item-wise & Bill-wise Rate Difference Audit (Replicating Original Excel Format 'MAIN')"
    ws['A2'].font = Font(name='Segoe UI', size=10, italic=True, color='475569')
    
    ws['A3'] = "Includes Dedicated Fuel & Energy (F&E) and Precious Metal Surcharge (PMS) Credit Note Tracking alongside RD and CD"
    ws['A3'].font = FONT_NOTE

    # Explicit CN Mapping for FY 26-27 Invoices
    cn_map = {
        '2026200000536': {
            'rd_cns': ['2026600000898', '2026600001433', '2026600000990'],
            'cd_cns': ['2026600001590'],
            'fe_cns': [],
            'pms_cns': []
        },
        '2026200001303': {
            'rd_cns': ['2026600001881', '2026600002285', '2026600002420'],
            'cd_cns': ['2026600002137'],
            'fe_cns': ['2026600003242'],
            'pms_cns': []
        },
        '2026200002138': {
            'rd_cns': ['2026600004371', '2026600004784'],
            'cd_cns': ['2026600004600'],
            'fe_cns': ['2026600006041'],
            'pms_cns': []
        },
        '2026200003174': {
            'rd_cns': ['2026600006714', '2026600007765'],
            'cd_cns': [{'cn_no': '2026600006986', 'amt': 33181.98}],
            'fe_cns': [],
            'pms_cns': []
        },
        '2026200003536': {
            'rd_cns': ['2026600006825', '2026600007892'],
            'cd_cns': [{'cn_no': '2026600006986', 'amt': 35947.14}],
            'fe_cns': [],
            'pms_cns': []
        },
        '2026200004977': {
            'rd_cns': ['2026600009498', '2026600010213'],
            'cd_cns': ['2026600009732'],
            'fe_cns': [],
            'pms_cns': []
        }
    }

    invoices = sorted(engine.invoices.values(), key=lambda x: x.get('invoice_date', ''))
    
    current_row = 5
    summary_rows = []
    bill_summary_info = []

    headers = [
        ('A', 'DATE/INV'),
        ('B', 'SR.'),
        ('C', 'ITEM'),
        ('D', 'MM'),
        ('E', 'WIDTH'),
        ('F', 'LENGTH'),
        ('G', 'QTY`'),
        ('H', 'SQM'),
        ('I', 'SQMM'),
        ('J', 'B. RATE'),
        ('K', 'RATE'),
        ('L', 'AMOUNT'),
        ('M', 'F&E'),
        ('N', 'PMS'),
        ('O', 'INS'),
        ('P', 'CGST'),
        ('Q', 'SGST'),
        ('R', 'ACTUAL AMT'),
        ('S', 'BILL AMOUNT'),
        ('T', 'DIFF'),
        ('U', 'RD RCVD 1'),
        ('V', 'RD RCVD 2'),
        ('W', 'RD RCVD 3'),
        ('X', 'TOTAL RD'),
        ('Y', 'PENDING RD'),
        ('Z', 'F&E CN NO.'),
        ('AA', 'F&E CN RCVD'),
        ('AB', 'PMS CN NO.'),
        ('AC', 'PMS CN RCVD'),
        ('AD', 'CD SHOULD'),
        ('AE', 'CD RCVD'),
        ('AF', 'PENDING CD'),
        ('AG', 'AUDIT REMARK')
    ]

    for inv_idx, inv in enumerate(invoices):
        inv_no = inv.get('invoice_no')
        inv_date = inv.get('invoice_date', '')
        try:
            d_parts = inv_date.split('-')
            fmt_date = f"{d_parts[2]}/{d_parts[1]}/{d_parts[0]}"
        except:
            fmt_date = inv_date

        items = inv.get('items', [])
        num_items = len(items)
        if num_items == 0:
            continue

        header_row = current_row
        for col_letter, h_title in headers:
            cell = ws[f"{col_letter}{header_row}"]
            cell.value = h_title
            cell.font = FONT_HEADER
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = THIN_BORDER
        ws.row_dimensions[header_row].height = 24

        item_start_row = header_row + 1
        item_end_row = item_start_row + num_items - 1
        
        # Populate Items
        for idx, it in enumerate(items):
            r = item_start_row + idx
            ws.row_dimensions[r].height = 19
            
            if idx == 0:
                ws[f"A{r}"].value = fmt_date
                ws[f"A{r}"].font = FONT_BOLD
                ws[f"A{r}"].alignment = Alignment(horizontal='center')
            elif idx == 1:
                ws[f"A{r}"].value = int(inv_no) if inv_no.isdigit() else inv_no
                ws[f"A{r}"].font = FONT_BOLD
                ws[f"A{r}"].alignment = Alignment(horizontal='center')
            elif idx == 2 and inv.get('truck_no'):
                ws[f"A{r}"].value = f"{inv.get('truck_no')} (LR:{inv.get('lr_no','')})"
                ws[f"A{r}"].font = FONT_NOTE
                ws[f"A{r}"].alignment = Alignment(horizontal='center')
                
            ws[f"B{r}"].value = it.get('sr')
            ws[f"B{r}"].alignment = Alignment(horizontal='center')
            
            desc_str = it.get('desc', '')
            t_short = 'MIR' if it.get('item_type') == 'MIR' or 'MIRROR' in desc_str.upper() else 'CLR'
            ws[f"C{r}"].value = f"{t_short} ({it.get('mm',0)}mm)"
            ws[f"C{r}"].font = FONT_REGULAR
            
            ws[f"D{r}"].value = it.get('mm')
            ws[f"D{r}"].alignment = Alignment(horizontal='center')
            
            ws[f"E{r}"].value = it.get('width')
            ws[f"E{r}"].alignment = Alignment(horizontal='right')
            ws[f"E{r}"].number_format = '#,##0'
            
            ws[f"F{r}"].value = it.get('length')
            ws[f"F{r}"].alignment = Alignment(horizontal='right')
            ws[f"F{r}"].number_format = '#,##0'
            
            ws[f"G{r}"].value = it.get('qty')
            ws[f"G{r}"].alignment = Alignment(horizontal='right')
            ws[f"G{r}"].number_format = '#,##0'
            
            ws[f"H{r}"].value = f"=E{r}*F{r}*G{r}/1000000"
            ws[f"H{r}"].alignment = Alignment(horizontal='right')
            ws[f"H{r}"].number_format = '#,##0.0000'
            
            ws[f"I{r}"].value = f"=H{r}*D{r}"
            ws[f"I{r}"].alignment = Alignment(horizontal='right')
            ws[f"I{r}"].number_format = '#,##0.0000'
            ws[f"I{r}"].font = FONT_BOLD
            
            ws[f"J{r}"].value = it.get('b_rate')
            ws[f"J{r}"].alignment = Alignment(horizontal='right')
            ws[f"J{r}"].number_format = '#,##0.00'
            
            ws[f"K{r}"].value = it.get('agreed_rate')
            ws[f"K{r}"].alignment = Alignment(horizontal='right')
            ws[f"K{r}"].number_format = '#,##0.00'
            ws[f"K{r}"].font = FONT_BOLD
            
            ws[f"L{r}"].value = f"=K{r}*I{r}"
            ws[f"L{r}"].alignment = Alignment(horizontal='right')
            ws[f"L{r}"].number_format = '#,##0.00'
            
            for c_idx in range(1, len(headers) + 1):
                c_let = get_column_letter(c_idx)
                if not ws[f"{c_let}{r}"].border.left.style:
                    ws[f"{c_let}{r}"].border = THIN_BORDER

        # Credit Notes for this Invoice
        inv_cns = cn_map.get(inv_no, {'rd_cns': [], 'cd_cns': [], 'fe_cns': [], 'pms_cns': []})
        rd_cn_list = list(inv_cns.get('rd_cns', []))
        cd_cn_list = list(inv_cns.get('cd_cns', []))
        fe_cn_list = list(inv_cns.get('fe_cns', []))
        pms_cn_list = list(inv_cns.get('pms_cns', []))

        # Dynamic discovery of CNs from engine
        for c_no, c in sorted(engine.credit_notes.items(), key=lambda x: (x[1].get('cn_date',''), x[0])):
            ref = c.get('ref_invoice', '')
            c_type = c.get('cn_type', '')
            desc = (c.get('desc') or '').upper()
            m_target = c.get('target_month') or c.get('cn_date', '')[:7]
            inv_m = inv.get('month_key') or inv_date[:7]
            
            matches_inv = (ref and (ref in inv_no or inv_no in ref)) or (not ref and m_target == inv_m and not cn_map.get(inv_no))
            if matches_inv:
                is_fe = c_type == 'FE_ADJ' or any(k in desc for k in ['F&E', 'ENERGY', 'EDIFF'])
                is_pms = c_type == 'PMS_RET' or 'PMS' in desc
                if is_fe:
                    if c_no not in fe_cn_list: fe_cn_list.append(c_no)
                elif is_pms:
                    if c_no not in pms_cn_list: pms_cn_list.append(c_no)
                elif c_type in ['RD', 'ARD'] or any(k in desc for k in ['RATE DIFF', 'ARD', 'BOOSTER']):
                    existing_rd = [x if isinstance(x, str) else x.get('cn_no') for x in rd_cn_list]
                    if c_no not in existing_rd: rd_cn_list.append(c_no)
                elif c_type == 'CD' or 'CASH DISCOUNT' in desc:
                    existing_cd = [x if isinstance(x, str) else x.get('cn_no') for x in cd_cn_list]
                    if c_no not in existing_cd: cd_cn_list.append(c_no)

        note_row = item_start_row
        if num_items >= 2:
            note_row = item_start_row + 1
            
        ws[f"T{note_row}"].value = "NOTE NO:"
        ws[f"T{note_row}"].font = FONT_BOLD
        ws[f"T{note_row}"].alignment = Alignment(horizontal='center')
        
        for cn_i, cn_ref in enumerate(rd_cn_list[:3]):
            col_target = ['U', 'V', 'W'][cn_i]
            c_no_str = cn_ref if isinstance(cn_ref, str) else cn_ref.get('cn_no')
            ws[f"{col_target}{note_row}"].value = int(c_no_str) if c_no_str.isdigit() else c_no_str
            ws[f"{col_target}{note_row}"].font = FONT_BOLD
            ws[f"{col_target}{note_row}"].alignment = Alignment(horizontal='center')

        # F&E CN Note
        if fe_cn_list:
            c_fe = fe_cn_list[0]
            fe_no_str = c_fe if isinstance(c_fe, str) else c_fe.get('cn_no')
            ws[f"Z{note_row}"].value = int(fe_no_str) if fe_no_str.isdigit() else fe_no_str
            ws[f"Z{note_row}"].font = FONT_AMBER
            ws[f"Z{note_row}"].alignment = Alignment(horizontal='center')
        else:
            ws[f"Z{note_row}"].value = "-"
            ws[f"Z{note_row}"].alignment = Alignment(horizontal='center')
            ws[f"Z{note_row}"].font = FONT_NOTE

        # PMS CN Note
        if pms_cn_list:
            c_pms = pms_cn_list[0]
            pms_no_str = c_pms if isinstance(c_pms, str) else c_pms.get('cn_no')
            ws[f"AB{note_row}"].value = int(pms_no_str) if pms_no_str.isdigit() else pms_no_str
            ws[f"AB{note_row}"].font = FONT_BOLD
            ws[f"AB{note_row}"].alignment = Alignment(horizontal='center')
        else:
            ws[f"AB{note_row}"].value = "-"
            ws[f"AB{note_row}"].alignment = Alignment(horizontal='center')
            ws[f"AB{note_row}"].font = FONT_NOTE
            
        # CD CN Note
        if cd_cn_list:
            c_cd = cd_cn_list[0]
            cd_no_str = c_cd if isinstance(c_cd, str) else c_cd.get('cn_no')
            ws[f"AD{note_row}"].value = "CD NOTE:"
            ws[f"AD{note_row}"].font = FONT_BOLD
            ws[f"AD{note_row}"].alignment = Alignment(horizontal='center')
            ws[f"AE{note_row}"].value = int(cd_no_str) if cd_no_str.isdigit() else cd_no_str
            ws[f"AE{note_row}"].font = FONT_BOLD
            ws[f"AE{note_row}"].alignment = Alignment(horizontal='center')

        # Summary Row
        sum_row = item_end_row + 1
        summary_rows.append(sum_row)
        ws.row_dimensions[sum_row].height = 21
        
        ws[f"A{sum_row}"].value = f"TOTAL #{inv_no[-4:]}"
        ws[f"A{sum_row}"].font = FONT_BOLD
        ws[f"A{sum_row}"].fill = SUMMARY_FILL
        
        ws[f"G{sum_row}"].value = f"=SUM(G{item_start_row}:G{item_end_row})"
        ws[f"G{sum_row}"].number_format = '#,##0'
        
        ws[f"H{sum_row}"].value = f"=SUM(H{item_start_row}:H{item_end_row})"
        ws[f"H{sum_row}"].number_format = '#,##0.0000'
        
        ws[f"I{sum_row}"].value = f"=SUM(I{item_start_row}:I{item_end_row})"
        ws[f"I{sum_row}"].number_format = '#,##0.0000'
        ws[f"I{sum_row}"].font = FONT_BOLD
        
        ws[f"L{sum_row}"].value = f"=SUM(L{item_start_row}:L{item_end_row})"
        ws[f"L{sum_row}"].number_format = '#,##0.00'
        ws[f"L{sum_row}"].font = FONT_BOLD
        
        ws[f"M{sum_row}"].value = inv.get('agreed_fe') if inv.get('fe_is_custom') else inv.get('fe', 0.0)
        ws[f"M{sum_row}"].number_format = '#,##0.00'
        
        ws[f"N{sum_row}"].value = inv.get('agreed_pms') if inv.get('pms_is_custom') else inv.get('pms', 0.0)
        ws[f"N{sum_row}"].number_format = '#,##0.00'
        
        ws[f"O{sum_row}"].value = inv.get('ins', 0.0)
        ws[f"O{sum_row}"].number_format = '#,##0.00'
        
        ws[f"P{sum_row}"].value = f"=(L{sum_row}+M{sum_row}+N{sum_row}+O{sum_row})*0.09"
        ws[f"P{sum_row}"].number_format = '#,##0.00'
        
        ws[f"Q{sum_row}"].value = f"=(L{sum_row}+M{sum_row}+N{sum_row}+O{sum_row})*0.09"
        ws[f"Q{sum_row}"].number_format = '#,##0.00'
        
        ws[f"R{sum_row}"].value = f"=L{sum_row}+M{sum_row}+N{sum_row}+O{sum_row}+P{sum_row}+Q{sum_row}"
        ws[f"R{sum_row}"].number_format = '#,##0.00'
        ws[f"R{sum_row}"].font = FONT_BOLD
        
        ws[f"S{sum_row}"].value = inv.get('total_amount', 0.0)
        ws[f"S{sum_row}"].number_format = '#,##0.00'
        ws[f"S{sum_row}"].font = FONT_BOLD
        
        ws[f"T{sum_row}"].value = f"=S{sum_row}-R{sum_row}"
        ws[f"T{sum_row}"].number_format = '#,##0.00'
        ws[f"T{sum_row}"].font = FONT_RED
        
        # Credit Note Values (RD)
        rd_cols = ['U', 'V', 'W']
        for cn_i, cn_ref in enumerate(rd_cn_list[:3]):
            col_target = rd_cols[cn_i]
            c_no_str = cn_ref if isinstance(cn_ref, str) else cn_ref.get('cn_no')
            cn_obj = engine.credit_notes.get(c_no_str)
            cn_val = cn_obj.get('total_amount', 0.0) if cn_obj else 0.0
            ws[f"{col_target}{sum_row}"].value = cn_val
            ws[f"{col_target}{sum_row}"].number_format = '#,##0.00'
            ws[f"{col_target}{sum_row}"].font = FONT_GREEN
            
        ws[f"X{sum_row}"].value = f"=SUM(U{sum_row}:W{sum_row})"
        ws[f"X{sum_row}"].number_format = '#,##0.00'
        ws[f"X{sum_row}"].font = FONT_GREEN
        
        ws[f"Y{sum_row}"].value = f"=T{sum_row}-X{sum_row}"
        ws[f"Y{sum_row}"].number_format = '#,##0.00'
        ws[f"Y{sum_row}"].font = FONT_BOLD

        # F&E CN Amount & Document No.
        fe_val = 0.0
        fe_no_display = "-"
        if fe_cn_list:
            c_fe = fe_cn_list[0]
            fe_no_display = c_fe if isinstance(c_fe, str) else c_fe.get('cn_no')
            for c_fe_item in fe_cn_list:
                c_fe_str = c_fe_item if isinstance(c_fe_item, str) else c_fe_item.get('cn_no')
                fe_obj = engine.credit_notes.get(c_fe_str)
                if fe_obj: fe_val += fe_obj.get('total_amount', 0.0)
        ws[f"Z{sum_row}"].value = int(fe_no_display) if str(fe_no_display).isdigit() else fe_no_display
        ws[f"Z{sum_row}"].alignment = Alignment(horizontal='center')
        if fe_val > 0:
            ws[f"Z{sum_row}"].font = FONT_AMBER
        else:
            ws[f"Z{sum_row}"].font = FONT_NOTE

        ws[f"AA{sum_row}"].value = fe_val
        ws[f"AA{sum_row}"].number_format = '#,##0.00'
        if fe_val > 0:
            ws[f"AA{sum_row}"].font = FONT_AMBER
            ws[f"AA{sum_row}"].fill = AMBER_FILL

        # PMS CN Amount & Document No.
        pms_val = 0.0
        pms_no_display = "-"
        if pms_cn_list:
            c_pms = pms_cn_list[0]
            pms_no_display = c_pms if isinstance(c_pms, str) else c_pms.get('cn_no')
            for c_pms_item in pms_cn_list:
                c_pms_str = c_pms_item if isinstance(c_pms_item, str) else c_pms_item.get('cn_no')
                pms_obj = engine.credit_notes.get(c_pms_str)
                if pms_obj: pms_val += pms_obj.get('total_amount', 0.0)
        ws[f"AB{sum_row}"].value = int(pms_no_display) if str(pms_no_display).isdigit() else pms_no_display
        ws[f"AB{sum_row}"].alignment = Alignment(horizontal='center')
        if pms_val > 0:
            ws[f"AB{sum_row}"].font = FONT_GREEN
        else:
            ws[f"AB{sum_row}"].font = FONT_NOTE

        ws[f"AC{sum_row}"].value = pms_val
        ws[f"AC{sum_row}"].number_format = '#,##0.00'
        if pms_val > 0:
            ws[f"AC{sum_row}"].font = FONT_GREEN
            ws[f"AC{sum_row}"].fill = SUCCESS_FILL

        # Cash Discount
        ws[f"AD{sum_row}"].value = f"=(S{sum_row}-X{sum_row})*0.03"
        ws[f"AD{sum_row}"].number_format = '#,##0.00'
        
        cd_val = 0.0
        if cd_cn_list:
            c_cd = cd_cn_list[0]
            if isinstance(c_cd, dict):
                cd_val = c_cd.get('amt', 0.0)
            else:
                cd_obj = engine.credit_notes.get(c_cd)
                cd_val = cd_obj.get('total_amount', 0.0) if cd_obj else 0.0
        ws[f"AE{sum_row}"].value = cd_val
        ws[f"AE{sum_row}"].number_format = '#,##0.00'
        
        ws[f"AF{sum_row}"].value = f"=AD{sum_row}-AE{sum_row}"
        ws[f"AF{sum_row}"].number_format = '#,##0.00'
        ws[f"AF{sum_row}"].font = FONT_BOLD
        
        remarks = []
        if inv_no == '2026200001303':
            remarks.append("May Boosters #2285 & #2420 (₹60.6K)")
            remarks.append("F&E CN #2026600003242 (₹15,233.00)")
        elif inv_no == '2026200002138':
            remarks.append("June Booster #4784 (₹26.3K)")
            remarks.append("F&E CN #2026600006041 (₹10,047.00)")
        elif inv_no in ['2026200003174', '2026200003536']:
            remarks.append("Aug ARD settlement CNs #7765 / #7892")
        elif inv_no == '2026200004977':
            remarks.append("Exact bill CN #9498 + extra ARD CN #10213")

        ws[f"AG{sum_row}"].value = " | ".join(remarks)
        ws[f"AG{sum_row}"].font = FONT_NOTE

        for c_idx in range(1, len(headers) + 1):
            c_let = get_column_letter(c_idx)
            cell = ws[f"{c_let}{sum_row}"]
            cell.border = THIN_BORDER
            if cell.fill.fill_type is None:
                cell.fill = SUMMARY_FILL

        # Basic Row
        b_row = sum_row + 1
        ws.row_dimensions[b_row].height = 18
        ws[f"A{b_row}"].value = "BASIC (/1.18)"
        ws[f"A{b_row}"].font = FONT_NOTE
        ws[f"A{b_row}"].alignment = Alignment(horizontal='center')
        
        ws[f"T{b_row}"].value = f"=T{sum_row}/1.18"
        ws[f"T{b_row}"].number_format = '#,##0.00'
        ws[f"T{b_row}"].font = FONT_NOTE
        
        for c_let in ['U', 'V', 'W']:
            ws[f"{c_let}{b_row}"].value = f"=IF({c_let}{sum_row}>0, {c_let}{sum_row}/1.18, \"\")"
            ws[f"{c_let}{b_row}"].number_format = '#,##0.00'
            ws[f"{c_let}{b_row}"].font = FONT_NOTE
            
        ws[f"X{b_row}"].value = f"=X{sum_row}/1.18"
        ws[f"X{b_row}"].number_format = '#,##0.00'
        ws[f"X{b_row}"].font = FONT_NOTE
        
        ws[f"Y{b_row}"].value = f"=Y{sum_row}/1.18"
        ws[f"Y{b_row}"].number_format = '#,##0.00'
        ws[f"Y{b_row}"].font = FONT_NOTE
        
        # Rate Row
        r_row = sum_row + 2
        ws.row_dimensions[r_row].height = 18
        ws[f"A{r_row}"].value = "RATE (/SQMM)"
        ws[f"A{r_row}"].font = FONT_NOTE
        ws[f"A{r_row}"].alignment = Alignment(horizontal='center')
        
        ws[f"K{r_row}"].value = "ASSUMED"
        ws[f"K{r_row}"].font = FONT_NOTE
        ws[f"K{r_row}"].alignment = Alignment(horizontal='right')
        
        ws[f"T{r_row}"].value = f"=T{b_row}/I{sum_row}"
        ws[f"T{r_row}"].number_format = '₹#,##0.00'
        ws[f"T{r_row}"].font = FONT_BOLD
        
        ws[f"X{r_row}"].value = f"=X{b_row}/I{sum_row}"
        ws[f"X{r_row}"].number_format = '₹#,##0.00'
        ws[f"X{r_row}"].font = FONT_GREEN
        
        ws[f"Y{r_row}"].value = f"=Y{b_row}/I{sum_row}"
        ws[f"Y{r_row}"].number_format = '₹#,##0.00'
        ws[f"Y{r_row}"].font = FONT_BOLD

        # Save info for PENDING SUMMARY
        bill_summary_info.append({
            'inv_no': inv_no,
            'inv_date': fmt_date,
            'items_str': ' & '.join(list(dict.fromkeys([('MIR' if it.get('item_type')=='MIR' or 'MIRROR' in it.get('desc','').upper() else 'CLR') for it in items]))),
            'discussed_rate_str': ' | '.join(list(dict.fromkeys([f"₹{it.get('agreed_rate',0):g}" for it in items if it.get('agreed_rate')]))),
            'sum_row': sum_row,
            'rate_row': r_row,
            'has_fe_cn': len(fe_cn_list) > 0,
            'fe_cn_no': fe_cn_list[0] if fe_cn_list else '-',
            'has_pms_cn': len(pms_cn_list) > 0,
            'pms_cn_no': pms_cn_list[0] if pms_cn_list else '-'
        })
        
        current_row = r_row + 2

    # -------------------------------------------------------------
    # GRAND TOTAL ROW
    # -------------------------------------------------------------
    tot_row = current_row + 1
    ws.row_dimensions[tot_row].height = 24
    
    ws[f"A{tot_row}"].value = "GRAND TOTAL (FY 2026-27)"
    ws[f"A{tot_row}"].font = Font(name='Segoe UI', size=10, bold=True, color='FFFFFF')
    ws[f"A{tot_row}"].fill = HEADER_FILL
    
    sum_rows_str = lambda col: '+'.join([f"{col}{sr}" for sr in summary_rows])
    
    ws[f"G{tot_row}"].value = f"={sum_rows_str('G')}"
    ws[f"H{tot_row}"].value = f"={sum_rows_str('H')}"
    ws[f"I{tot_row}"].value = f"={sum_rows_str('I')}"
    ws[f"L{tot_row}"].value = f"={sum_rows_str('L')}"
    ws[f"M{tot_row}"].value = f"={sum_rows_str('M')}"
    ws[f"N{tot_row}"].value = f"={sum_rows_str('N')}"
    ws[f"O{tot_row}"].value = f"={sum_rows_str('O')}"
    ws[f"P{tot_row}"].value = f"={sum_rows_str('P')}"
    ws[f"Q{tot_row}"].value = f"={sum_rows_str('Q')}"
    ws[f"R{tot_row}"].value = f"={sum_rows_str('R')}"
    ws[f"S{tot_row}"].value = f"={sum_rows_str('S')}"
    ws[f"T{tot_row}"].value = f"={sum_rows_str('T')}"
    ws[f"X{tot_row}"].value = f"={sum_rows_str('X')}"
    ws[f"Y{tot_row}"].value = f"={sum_rows_str('Y')}"
    ws[f"Z{tot_row}"].value = "TOTAL F&E:"
    ws[f"Z{tot_row}"].font = FONT_NOTE
    ws[f"Z{tot_row}"].alignment = Alignment(horizontal='right')
    ws[f"AA{tot_row}"].value = f"={sum_rows_str('AA')}"
    ws[f"AB{tot_row}"].value = "TOTAL PMS:"
    ws[f"AB{tot_row}"].font = FONT_NOTE
    ws[f"AB{tot_row}"].alignment = Alignment(horizontal='right')
    ws[f"AC{tot_row}"].value = f"={sum_rows_str('AC')}"
    ws[f"AD{tot_row}"].value = f"={sum_rows_str('AD')}"
    ws[f"AE{tot_row}"].value = f"={sum_rows_str('AE')}"
    ws[f"AF{tot_row}"].value = f"={sum_rows_str('AF')}"
    
    ws[f"G{tot_row}"].number_format = '#,##0'
    ws[f"H{tot_row}"].number_format = '#,##0.0000'
    ws[f"I{tot_row}"].number_format = '#,##0.0000'
    ws[f"I{tot_row}"].font = FONT_BOLD
    
    for c_let in ['L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'X', 'Y', 'AA', 'AC', 'AD', 'AE', 'AF']:
        cell = ws[f"{c_let}{tot_row}"]
        cell.number_format = '₹#,##0.00'
        cell.font = FONT_BOLD
        cell.fill = TOTAL_FILL
        cell.border = DOUBLE_BOTTOM_BORDER

    # -------------------------------------------------------------
    # PENDING SUMMARY TABLE (Replicating 'PENDING SUMMARY FOR 25-26')
    # -------------------------------------------------------------
    r_pend_head = tot_row + 3
    ws.merge_cells(f"C{r_pend_head}:M{r_pend_head}")
    p_title = ws[f"C{r_pend_head}"]
    p_title.value = "📋 PENDING SUMMARY (FY 2026–27) — REPLICATING ORIGINAL EXCEL FORMAT"
    p_title.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    p_title.fill = HEADER_FILL
    p_title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[r_pend_head].height = 24

    p_headers = [
        ('C', 'SR.'),
        ('D', 'DATE'),
        ('E', 'BILL NO'),
        ('F', 'ITEMS'),
        ('G', 'DISCUSSED RATE'),
        ('H', 'RD RCVD (IN RATE)'),
        ('I', 'RD PENDING (IN RATE)'),
        ('J', 'RD PENDING (INC GST)'),
        ('K', 'F&E CN RCVD (INC GST)'),
        ('L', 'PMS CN RCVD (INC GST)'),
        ('M', 'CD PENDING (INC GST)')
    ]
    
    r_p_sub = r_pend_head + 1
    ws.row_dimensions[r_p_sub].height = 22
    for p_col, p_txt in p_headers:
        cell = ws[f"{p_col}{r_p_sub}"]
        cell.value = p_txt
        cell.font = FONT_BOLD
        cell.fill = GRAY_HEADER
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = THIN_BORDER

    p_row_start = r_p_sub + 1
    for p_i, b_info in enumerate(bill_summary_info, start=1):
        curr_p = p_row_start + p_i - 1
        ws.row_dimensions[curr_p].height = 20
        
        ws[f"C{curr_p}"].value = f"{p_i})"
        ws[f"C{curr_p}"].alignment = Alignment(horizontal='center')
        
        ws[f"D{curr_p}"].value = b_info['inv_date']
        ws[f"D{curr_p}"].alignment = Alignment(horizontal='center')
        
        ws[f"E{curr_p}"].value = int(b_info['inv_no']) if b_info['inv_no'].isdigit() else b_info['inv_no']
        ws[f"E{curr_p}"].alignment = Alignment(horizontal='center')
        ws[f"E{curr_p}"].font = FONT_BOLD
        
        ws[f"F{curr_p}"].value = b_info['items_str']
        ws[f"F{curr_p}"].alignment = Alignment(horizontal='left')
        
        ws[f"G{curr_p}"].value = b_info['discussed_rate_str']
        ws[f"G{curr_p}"].alignment = Alignment(horizontal='center')
        
        sr_r = b_info['sum_row']
        rt_r = b_info['rate_row']
        
        ws[f"H{curr_p}"].value = f"=X{rt_r}"
        ws[f"H{curr_p}"].number_format = '₹#,##0.00'
        ws[f"H{curr_p}"].alignment = Alignment(horizontal='right')
        
        ws[f"I{curr_p}"].value = f"=Y{rt_r}"
        ws[f"I{curr_p}"].number_format = '₹#,##0.00'
        ws[f"I{curr_p}"].alignment = Alignment(horizontal='right')
        
        ws[f"J{curr_p}"].value = f"=Y{sr_r}"
        ws[f"J{curr_p}"].number_format = '₹#,##0.00'
        ws[f"J{curr_p}"].alignment = Alignment(horizontal='right')
        ws[f"J{curr_p}"].font = FONT_BOLD
        
        ws[f"K{curr_p}"].value = f"=AA{sr_r}"
        ws[f"K{curr_p}"].number_format = '₹#,##0.00'
        ws[f"K{curr_p}"].alignment = Alignment(horizontal='right')
        if b_info['has_fe_cn']:
            ws[f"K{curr_p}"].font = FONT_AMBER
            ws[f"K{curr_p}"].fill = AMBER_FILL

        ws[f"L{curr_p}"].value = f"=AC{sr_r}"
        ws[f"L{curr_p}"].number_format = '₹#,##0.00'
        ws[f"L{curr_p}"].alignment = Alignment(horizontal='right')

        ws[f"M{curr_p}"].value = f"=AF{sr_r}"
        ws[f"M{curr_p}"].number_format = '₹#,##0.00'
        ws[f"M{curr_p}"].alignment = Alignment(horizontal='right')
        ws[f"M{curr_p}"].font = FONT_BOLD
        
        for p_col, _ in p_headers:
            ws[f"{p_col}{curr_p}"].border = THIN_BORDER

    p_tot_row = p_row_start + len(bill_summary_info)
    ws.row_dimensions[p_tot_row].height = 22
    ws[f"C{p_tot_row}"].value = "TOTAL DUE (FY 2026-27):"
    ws[f"C{p_tot_row}"].font = FONT_BOLD
    ws.merge_cells(f"C{p_tot_row}:I{p_tot_row}")
    
    ws[f"J{p_tot_row}"].value = f"=SUM(J{p_row_start}:J{p_tot_row-1})"
    ws[f"J{p_tot_row}"].number_format = '₹#,##0.00'
    ws[f"J{p_tot_row}"].font = FONT_BOLD
    ws[f"J{p_tot_row}"].fill = TOTAL_FILL
    ws[f"J{p_tot_row}"].border = DOUBLE_BOTTOM_BORDER
    
    ws[f"K{p_tot_row}"].value = f"=SUM(K{p_row_start}:K{p_tot_row-1})"
    ws[f"K{p_tot_row}"].number_format = '₹#,##0.00'
    ws[f"K{p_tot_row}"].font = FONT_AMBER
    ws[f"K{p_tot_row}"].fill = TOTAL_FILL
    ws[f"K{p_tot_row}"].border = DOUBLE_BOTTOM_BORDER

    ws[f"L{p_tot_row}"].value = f"=SUM(L{p_row_start}:L{p_tot_row-1})"
    ws[f"L{p_tot_row}"].number_format = '₹#,##0.00'
    ws[f"L{p_tot_row}"].font = FONT_BOLD
    ws[f"L{p_tot_row}"].fill = TOTAL_FILL
    ws[f"L{p_tot_row}"].border = DOUBLE_BOTTOM_BORDER

    ws[f"M{p_tot_row}"].value = f"=SUM(M{p_row_start}:M{p_tot_row-1})"
    ws[f"M{p_tot_row}"].number_format = '₹#,##0.00'
    ws[f"M{p_tot_row}"].font = FONT_BOLD
    ws[f"M{p_tot_row}"].fill = TOTAL_FILL
    ws[f"M{p_tot_row}"].border = DOUBLE_BOTTOM_BORDER

    # -------------------------------------------------------------
    # RECEIVED SUMMARY & AUDIT EXPLANATION
    # -------------------------------------------------------------
    r_aud = p_tot_row + 3
    ws.merge_cells(f"C{r_aud}:Y{r_aud}")
    aud_title = ws[f"C{r_aud}"]
    aud_title.value = "🔍 COMMERCIAL AUDIT & COMPREHENSIVE SETTLEMENT BREAKDOWN (RD, F&E, PMS, CD)"
    aud_title.font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
    aud_title.fill = SUBHEADER_FILL
    aud_title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[r_aud].height = 24
    
    audit_lines = [
        (f"1. Total Billed Turnover ({len(invoices)} Invoices)", f"=S{tot_row}", "Gross billed amount from GGL plant invoices (inc. 18% GST)"),
        ("2. Total Actual Discussed Cost", f"=R{tot_row}", "Calculated based on agreed discussed rates (inc. 18% GST)"),
        ("3. Total Discussed Rate Difference (DIFF)", f"=T{tot_row}", "Total gross overbilling across all dispatches (Billed - Actual)"),
        ("4. Total RD / ARD Settled via Credit Notes", f"=X{tot_row}", "Sum of all monthly RD/ARD credit notes issued by GGL"),
        ("5. Net Rate Difference Position (Settled - DIFF)", f"=X{tot_row}-T{tot_row}", "Surplus credit in Mehta Impex favor (Extra Credit if positive / Pending if negative)"),
        ("6. Total Fuel & Energy (F&E) CNs Credited", f"=AA{tot_row}", "F&E Surcharge refund credit notes issued by GGL (CN #3242 & #6041)"),
        ("7. Total Precious Metal Surcharge (PMS) CNs Credited", f"=AC{tot_row}", "PMS Surcharge refund credit notes issued by GGL"),
        ("8. Total Cash Discount (CD 3%) Payable", f"=AD{tot_row}", "3% Cash Discount on net purchases (Turnaround prompt payment)"),
        ("9. Total Cash Discount (CD) Received", f"=AE{tot_row}", "Total CD credit notes issued by GGL"),
        ("10. Net Cash Discount Shortfall Due from GGL", f"=AF{tot_row}", "Actual prompt payment recovery due from GGL (= CD Payable - CD Received)"),
    ]

    for a_idx, (a_label, a_formula, a_desc) in enumerate(audit_lines):
        row_n = r_aud + 1 + a_idx
        ws.row_dimensions[row_n].height = 20
        
        ws[f"C{row_n}"].value = a_label
        ws[f"C{row_n}"].font = FONT_BOLD
        ws[f"C{row_n}"].border = THIN_BORDER
        ws.merge_cells(f"C{row_n}:I{row_n}")
        
        c_val = ws[f"J{row_n}"]
        c_val.value = a_formula
        c_val.font = FONT_BOLD if a_idx not in [4, 5, 9] else (FONT_GREEN if a_idx in [4, 5] else FONT_RED)
        c_val.number_format = '₹#,##0.00'
        c_val.alignment = Alignment(horizontal='right')
        c_val.border = THIN_BORDER
        c_val.fill = SUCCESS_FILL if a_idx in [4, 5] else (HIGHLIGHT_FILL if a_idx == 9 else PatternFill(fill_type=None))
        ws.merge_cells(f"J{row_n}:L{row_n}")
        
        c_desc = ws[f"M{row_n}"]
        c_desc.value = a_desc
        c_desc.font = FONT_NOTE
        c_desc.border = THIN_BORDER
        ws.merge_cells(f"M{row_n}:Y{row_n}")

    # Detailed Audit Findings Box
    r_exp = r_aud + len(audit_lines) + 2
    ws.merge_cells(f"C{r_exp}:Y{r_exp}")
    ws[f"C{r_exp}"].value = "💡 DETAILED SURCHARGE & INCENTIVE AUDIT FINDINGS (F&E, PMS & BOOSTER CNS)"
    ws[f"C{r_exp}"].font = Font(name='Segoe UI', size=10, bold=True, color='1E293B')
    ws[f"C{r_exp}"].fill = HIGHLIGHT_FILL
    ws.row_dimensions[r_exp].height = 22

    explanations = [
        "A. FUEL & ENERGY (F&E) SURCHARGE CREDIT NOTES (₹25,280.00 Total): GGL credited two separate credit notes for Fuel & Energy surcharge adjustments: (1) CN #2026600003242 of ₹15,233.00 dated 19-Jun-2026 for Invoice #2026200001303 ('F&E200001303'), and (2) CN #2026600006041 of ₹10,047.00 dated 14-Jul-2026 for Invoice #2026200002138 ('JUN26 EDIFF200002138'). These are tracked separately from RD/ARD.",
        "B. PRECIOUS METAL SURCHARGE (PMS): PMS was billed on silvering mirrors (e.g., ₹103.6K in May, ₹101.3K in June, ₹101.7K in July). No separate PMS reversal credit notes have been issued by GGL in FY 2026–27; rates remain at billed list.",
        "C. MAY 2026 BOOSTER CREDIT NOTES (+₹39,114.38 Surplus): On Invoice #2026200001303, discussed overbilling was ₹82,263.62. GGL issued normal ARD CN #1881 of ₹60,732.00, PLUS two Booster CNs (#2285 for ₹14,162.00 & #2420 for ₹46,484.00), giving total RD credit of ₹1,21,378.00! These Booster incentives exceeded standard discussed diff.",
        "D. AUGUST 2026 JULY SETTLEMENT CNs (₹63,279.00): In July (Invs #3174 & #3536), overbilling was ₹457,507.60. GGL issued ₹393,791.00 in July, and then issued CN #7765 (₹29,255) & #7892 (₹34,024) in August titled 'ARD JUL26', bringing total July settlement to ₹457,070.00 (leaving only ₹437.60 pending).",
        "E. NET RECONCILIATION SUMMARY: Glass RD is 100% settled (+₹7,609.75 net extra credit pool). Fuel & Energy surcharges are settled via ₹25,280.00 in CNs. Cash Discount has a remaining recoverable claim of ₹4,109.49."
    ]

    for e_i, e_txt in enumerate(explanations):
        curr_e = r_exp + 1 + e_i
        ws.row_dimensions[curr_e].height = 28
        ws.merge_cells(f"C{curr_e}:Y{curr_e}")
        ws[f"C{curr_e}"].value = e_txt
        ws[f"C{curr_e}"].font = FONT_NOTE
        ws[f"C{curr_e}"].alignment = Alignment(wrap_text=True, vertical='center')

    # Auto-fit columns
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            v_str = str(cell.value or '')
            if '\n' in v_str:
                v_str = max(v_str.split('\n'), key=len)
            max_len = max(max_len, len(v_str))
        ws.column_dimensions[col_letter].width = max(max_len + 3, 11)
        
    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['Z'].width = 15
    ws.column_dimensions['AA'].width = 14
    ws.column_dimensions['AB'].width = 15
    ws.column_dimensions['AC'].width = 14
    ws.column_dimensions['AG'].width = 38


    return ws


def generate_reconciliation_workbook(engine, output_path='/Users/sparshmehta/.gemini/antigravity/scratch/GGL_Master_Reconciliation.xlsx'):
    wb = openpyxl.Workbook()
    
    # -------------------------------------------------------------
    # SHEET 1: MONTH-WISE RATE DIFFERENCE RECONCILIATION
    # -------------------------------------------------------------
    ws_dash = wb.active
    ws_dash.title = 'Month-wise RD Recon'
    ws_dash.views.sheetView[0].showGridLines = True
    
    ws_dash['A1'] = "MEHTA IMPEX — GUJARAT GUARDIAN LTD (GGL) RECONCILIATION"
    ws_dash['A1'].font = FONT_TITLE
    ws_dash['A2'] = "FY 2026–27 Rate Difference (RD/ARD) & Credit Note Tracking (Excluding Historical Pre-2026 Balance)"
    ws_dash['A2'].font = Font(name='Segoe UI', size=10, italic=True, color='4A5568')
    
    # KPI Summary Cards (Rows 4-5)
    # KPI Summary Cards (Rows 4-5)
    monthly_data = engine.get_monthly_reconciliation()
    start_row = 10
    num_months = len(monthly_data)
    end_row = start_row + num_months
    
    kpis = [
        ("Total Billed (FY 26-27)", f"=SUM(F{start_row}:F{end_row-1})", "B4", "C4", "B5", "C5", NAVY),
        ("Actual Expected Cost", f"=SUM(G{start_row}:G{end_row-1})", "D4", "E4", "D5", "E5", NAVY),
        ("Overbilled Rate Diff", f"=SUM(H{start_row}:H{end_row-1})", "F4", "G4", "F5", "G5", 'C53030'),
        ("Total Settled via CNs", f"=SUM(M{start_row}:M{end_row-1})", "H4", "I4", "H5", "I5", '2F855A'),
        ("Net Extra Credit (Mehta Impex)", f"=M{end_row}-H{end_row}", "J4", "K4", "J5", "K5", '2F855A'),
    ]
    
    for title, formula, top_l, top_r, bot_l, bot_r, val_color in kpis:
        ws_dash.merge_cells(f"{top_l}:{top_r}")
        ws_dash.merge_cells(f"{bot_l}:{bot_r}")
        c_title = ws_dash[top_l]
        c_title.value = title
        c_title.font = Font(name='Segoe UI', size=9, bold=True, color='718096')
        c_title.alignment = Alignment(horizontal='center', vertical='center')
        c_title.fill = PatternFill(start_color='EDF2F7', end_color='EDF2F7', fill_type='solid')
        
        c_val = ws_dash[bot_l]
        c_val.value = formula
        c_val.font = Font(name='Segoe UI', size=13, bold=True, color=val_color)
        c_val.alignment = Alignment(horizontal='center', vertical='center')
        c_val.number_format = '₹#,##0.00'
        c_val.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        
        for row in ws_dash[f"{top_l}:{bot_r}"]:
            for cell in row:
                cell.border = THIN_BORDER

    # Headers at Row 9
    headers_dash = [
        "Month", "Invoice IDs", "Total SQMM",
        "Mirror SQMM", "Clear SQMM", "Billed Total (₹)", "Actual Cost (₹)",
        "Overbilled DIFF (₹)", "CN Document IDs Credited",
        "RD Received (₹)", "ARD Received (₹)", "F&E Refund (₹)", "Total Settled (₹)",
        "MONTH NET DIFF (₹)", "CUMULATIVE RUNNING BAL (₹)", "Status"
    ]
    
    row_idx = 9
    for col_idx, h in enumerate(headers_dash, start=1):
        cell = ws_dash.cell(row=row_idx, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
        
    ws_dash.row_dimensions[row_idx].height = 28
    
    for idx, m in enumerate(monthly_data):
        curr_row = start_row + idx
        ws_dash.row_dimensions[curr_row].height = 22
        fill = ZEBRA_FILL if idx % 2 == 1 else PatternFill(fill_type=None)
        
        # Month
        c1 = ws_dash.cell(row=curr_row, column=1, value=m['month_key'])
        c1.alignment = Alignment(horizontal='center', vertical='center')
        c1.font = FONT_BOLD
        
        # Invoice IDs
        inv_ids_str = ", ".join([str(inv['invoice_no']) for inv in m['invoices']]) or "-"
        c2 = ws_dash.cell(row=curr_row, column=2, value=inv_ids_str)
        c2.alignment = Alignment(horizontal='center', vertical='center')
        c2.font = Font(name='Segoe UI', size=9)
        
        # SQMM
        c3 = ws_dash.cell(row=curr_row, column=3, value=m['total_sqmm'])
        format_qty_cell(c3)
        c3.font = FONT_BOLD
        c4 = ws_dash.cell(row=curr_row, column=4, value=m['mirror_sqmm'])
        format_qty_cell(c4)
        c5 = ws_dash.cell(row=curr_row, column=5, value=m['clear_sqmm'])
        format_qty_cell(c5)
        
        # Billed & Actual
        c6 = ws_dash.cell(row=curr_row, column=6, value=m['billed_total'])
        format_currency_cell(c6)
        
        c7 = ws_dash.cell(row=curr_row, column=7, value=m['actual_total'] or 0.0)
        format_currency_cell(c7)
        
        c8 = ws_dash.cell(row=curr_row, column=8, value=f"=F{curr_row}-G{curr_row}")
        format_currency_cell(c8)
        c8.font = FONT_BOLD
        
        # CN Document IDs credited
        rd_cns = [str(cn['cn_no']) for cn in m['credit_notes'] if cn.get('cn_type') in ['RD', 'ARD']]
        cn_ids_str = ", ".join(rd_cns) or "-"
        c9 = ws_dash.cell(row=curr_row, column=9, value=cn_ids_str)
        c9.alignment = Alignment(horizontal='center', vertical='center')
        c9.font = Font(name='Segoe UI', size=8)

        # Credits
        c10 = ws_dash.cell(row=curr_row, column=10, value=m['rd_received'])
        format_currency_cell(c10)
        c11 = ws_dash.cell(row=curr_row, column=11, value=m['ard_received'])
        format_currency_cell(c11)
        c12 = ws_dash.cell(row=curr_row, column=12, value=m['fe_received'])
        format_currency_cell(c12)
        
        c13 = ws_dash.cell(row=curr_row, column=13, value=f"=J{curr_row}+K{curr_row}")
        format_currency_cell(c13)
        c13.font = FONT_BOLD
        
        # MONTH NET DIFFERENCE (Column 14)
        # Positive = Pending Claim (Red); Negative = Extra Credit (Green)
        c14 = ws_dash.cell(row=curr_row, column=14, value=f"=H{curr_row}-M{curr_row}")
        format_currency_cell(c14)
        net_diff = m.get('month_net_diff', 0.0)
        if net_diff < -10:
            c14.font = Font(name='Segoe UI', size=10, bold=True, color='2F855A')
            c14.fill = SUCCESS_FILL
        elif net_diff > 10:
            c14.font = Font(name='Segoe UI', size=10, bold=True, color='C53030')
            c14.fill = PENDING_FILL
        else:
            c14.font = Font(name='Segoe UI', size=10, bold=True, color='2F855A')
            c14.fill = SUCCESS_FILL

        # CUMULATIVE RUNNING BALANCE (Column 15)
        if idx == 0:
            c15 = ws_dash.cell(row=curr_row, column=15, value=f"=N{curr_row}")
        else:
            c15 = ws_dash.cell(row=curr_row, column=15, value=f"=O{curr_row-1}+N{curr_row}")
        format_currency_cell(c15)
        cum_bal = m.get('cumulative_rd_balance', 0.0)
        if cum_bal <= 10: # Extra credit or balanced
            c15.font = Font(name='Segoe UI', size=10, bold=True, color='2F855A')
            c15.fill = SUCCESS_FILL
        else:
            c15.font = Font(name='Segoe UI', size=10, bold=True, color='C53030')
            c15.fill = PENDING_FILL
        
        # Status (Column 16)
        c16 = ws_dash.cell(row=curr_row, column=16)
        if net_diff < -10:
            c16.value = "EXTRA CREDIT (Settles Prior Months)"
            c16.font = FONT_OK
            c16.fill = SUCCESS_FILL
        elif -10 <= net_diff <= 10:
            c16.value = "100% SETTLED"
            c16.font = FONT_OK
            c16.fill = SUCCESS_FILL
        elif cum_bal <= 0:
            c16.value = "COVERED BY ADVANCE POOL"
            c16.font = FONT_OK
            c16.fill = SUCCESS_FILL
        else:
            c16.value = "PENDING CLAIM"
            c16.font = FONT_ALERT
            c16.fill = PENDING_FILL
        c16.alignment = Alignment(horizontal='center', vertical='center')

        for col_i in range(1, len(headers_dash) + 1):
            ws_dash.cell(row=curr_row, column=col_i).border = THIN_BORDER

    # Total Row
    ws_dash.row_dimensions[end_row].height = 25
    
    t1 = ws_dash.cell(row=end_row, column=1, value="TOTAL (FY 26-27)")
    t1.font = FONT_BOLD
    t1.alignment = Alignment(horizontal='center', vertical='center')
    
    ws_dash.cell(row=end_row, column=2, value="")
    format_qty_cell(ws_dash.cell(row=end_row, column=3, value=f"=SUM(C{start_row}:C{end_row-1})"))
    format_qty_cell(ws_dash.cell(row=end_row, column=4, value=f"=SUM(D{start_row}:D{end_row-1})"))
    format_qty_cell(ws_dash.cell(row=end_row, column=5, value=f"=SUM(E{start_row}:E{end_row-1})"))
    
    format_currency_cell(ws_dash.cell(row=end_row, column=6, value=f"=SUM(F{start_row}:F{end_row-1})"))
    format_currency_cell(ws_dash.cell(row=end_row, column=7, value=f"=SUM(G{start_row}:G{end_row-1})"))
    format_currency_cell(ws_dash.cell(row=end_row, column=8, value=f"=SUM(H{start_row}:H{end_row-1})"))
    
    ws_dash.cell(row=end_row, column=9, value="")
    format_currency_cell(ws_dash.cell(row=end_row, column=10, value=f"=SUM(J{start_row}:J{end_row-1})"))
    format_currency_cell(ws_dash.cell(row=end_row, column=11, value=f"=SUM(K{start_row}:K{end_row-1})"))
    format_currency_cell(ws_dash.cell(row=end_row, column=12, value=f"=SUM(L{start_row}:L{end_row-1})"))
    format_currency_cell(ws_dash.cell(row=end_row, column=13, value=f"=SUM(M{start_row}:M{end_row-1})"))
    
    # Net Total
    t_net = ws_dash.cell(row=end_row, column=14, value=f"=SUM(N{start_row}:N{end_row-1})")
    format_currency_cell(t_net)
    t_net.font = Font(name='Segoe UI', size=11, bold=True, color='2F855A')
    t_net.fill = SUCCESS_FILL

    # Final Cumulative Net
    t_cum = ws_dash.cell(row=end_row, column=15, value=f"=N{end_row}")
    format_currency_cell(t_cum)
    t_cum.font = Font(name='Segoe UI', size=11, bold=True, color='2F855A')
    t_cum.fill = SUCCESS_FILL

    t_st = ws_dash.cell(row=end_row, column=16, value=f'=IF(O{end_row}<-10, "NET EXTRA CREDIT", IF(O{end_row}>10, "PENDING CLAIM", "100% SETTLED"))')
    t_st.font = FONT_OK
    t_st.fill = SUCCESS_FILL
    t_st.alignment = Alignment(horizontal='center', vertical='center')
    
    for col_i in range(1, len(headers_dash) + 1):
        c = ws_dash.cell(row=end_row, column=col_i)
        c.border = DOUBLE_BOTTOM_BORDER

    auto_fit_columns(ws_dash)

    # -------------------------------------------------------------
    # SHEET 2: MAIN (HISTORICAL FORMAT)
    # -------------------------------------------------------------
    add_historical_main_sheet(wb, engine)

    # -------------------------------------------------------------
    # SHEET 2: BILL-WISE BREAKDOWN & LINKED CN IDS
    # -------------------------------------------------------------
    ws_inv = wb.create_sheet(title='Bill-wise Breakdown')
    ws_inv.views.sheetView[0].showGridLines = True
    
    headers_inv = [
        "Month", "Invoice Date", "Invoice No", "Item Sr", "Type", "Description",
        "MM", "Width", "Length", "Qty", "SQM", "SQMM",
        "Billed Rate (₹)", "Discussed Rate (₹)", "Agreed Base Amt (₹)",
        "F&E (₹)", "PMS (₹)", "INS (₹)", "CGST (9%)", "SGST (9%)",
        "ACTUAL AMT (₹)", "BILL AMOUNT (₹)", "OVERBILLING (DIFF ₹)",
        "Linked CN Document IDs", "Status"
    ]
    
    ws_inv.row_dimensions[1].height = 26
    for col_idx, h in enumerate(headers_inv, start=1):
        cell = ws_inv.cell(row=1, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = THIN_BORDER
        
    curr_row = 2
    for inv_no, inv in sorted(engine.invoices.items(), key=lambda x: x[1].get('invoice_date', '')):
        m_key = inv.get('month_key', '')
        inv_dt = inv.get('invoice_date', '')
        items = inv.get('items', [])
        num_items = len(items)
        
        # Find matched CNs for this invoice (RD, ARD, F&E, PMS, CD)
        m_cns = []
        for c_no, cn in sorted(engine.credit_notes.items(), key=lambda x: (x[1].get('cn_date',''), x[0])):
            ref = cn.get('ref_invoice', '')
            c_type = cn.get('cn_type', '')
            desc = (cn.get('desc') or '').upper()
            m_target = cn.get('target_month') or cn.get('cn_date', '')[:7]
            if (ref and (ref in inv_no or inv_no in ref)) or (not ref and m_target == m_key):
                if c_type in ['RD', 'ARD']:
                    m_cns.append(f"#{c_no}")
                elif c_type == 'FE_ADJ' or any(k in desc for k in ['F&E', 'ENERGY', 'EDIFF']):
                    m_cns.append(f"⛽ F&E: #{c_no} (₹{cn.get('total_amount',0):,.0f})")
                elif c_type == 'PMS_RET' or 'PMS' in desc:
                    m_cns.append(f"🥈 PMS: #{c_no}")
                elif c_type == 'CD':
                    m_cns.append(f"CD: #{c_no}")
        cn_str = " | ".join(m_cns) if m_cns else "-" 
        
        for item_idx, it in enumerate(items):
            ws_inv.row_dimensions[curr_row].height = 20
            
            ws_inv.cell(row=curr_row, column=1, value=m_key).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=2, value=inv_dt).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=3, value=inv_no).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=4, value=it.get('sr')).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=5, value=it.get('item_type')).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=6, value=it.get('desc'))
            
            ws_inv.cell(row=curr_row, column=7, value=it.get('mm')).alignment = Alignment(horizontal='center')
            ws_inv.cell(row=curr_row, column=8, value=it.get('width')).alignment = Alignment(horizontal='right')
            ws_inv.cell(row=curr_row, column=9, value=it.get('length')).alignment = Alignment(horizontal='right')
            
            ws_inv.cell(row=curr_row, column=10, value=it.get('qty')).alignment = Alignment(horizontal='right')
            format_qty_cell(ws_inv.cell(row=curr_row, column=11, value=it.get('sqm')))
            format_qty_cell(ws_inv.cell(row=curr_row, column=12, value=it.get('sqmm')))
            
            format_currency_cell(ws_inv.cell(row=curr_row, column=13, value=it.get('b_rate')))
            
            c_agr = ws_inv.cell(row=curr_row, column=14, value=it.get('agreed_rate'))
            format_currency_cell(c_agr)
            c_agr.font = FONT_BOLD
            c_agr.fill = PatternFill(start_color='FFF9C4', end_color='FFF9C4', fill_type='solid')
            
            c_base = ws_inv.cell(row=curr_row, column=15, value=f"=L{curr_row}*N{curr_row}")
            format_currency_cell(c_base)
            
            if item_idx == 0:
                format_currency_cell(ws_inv.cell(row=curr_row, column=16, value=inv.get('agreed_fe') if inv.get('fe_is_custom') else inv.get('fe', 0.0)))
                format_currency_cell(ws_inv.cell(row=curr_row, column=17, value=inv.get('agreed_pms') if inv.get('pms_is_custom') else inv.get('pms', 0.0)))
                format_currency_cell(ws_inv.cell(row=curr_row, column=18, value=inv.get('ins', 0.0)))
                
                r_start = curr_row
                r_end = curr_row + num_items - 1
                
                ws_inv.cell(row=curr_row, column=19, value=f"=ROUND((SUM(O{r_start}:O{r_end})+P{curr_row}+Q{curr_row}+R{curr_row})*0.09, 2)")
                format_currency_cell(ws_inv.cell(row=curr_row, column=19))
                
                ws_inv.cell(row=curr_row, column=20, value=f"=ROUND((SUM(O{r_start}:O{r_end})+P{curr_row}+Q{curr_row}+R{curr_row})*0.09, 2)")
                format_currency_cell(ws_inv.cell(row=curr_row, column=20))
                
                ws_inv.cell(row=curr_row, column=21, value=f"=SUM(O{r_start}:O{r_end})+P{curr_row}+Q{curr_row}+R{curr_row}+S{curr_row}+T{curr_row}")
                format_currency_cell(ws_inv.cell(row=curr_row, column=21))
                ws_inv.cell(row=curr_row, column=21).font = FONT_BOLD
                
                format_currency_cell(ws_inv.cell(row=curr_row, column=22, value=inv.get('total_amount', 0.0)))
                ws_inv.cell(row=curr_row, column=22).font = FONT_BOLD
                
                ws_inv.cell(row=curr_row, column=23, value=f"=V{curr_row}-U{curr_row}")
                format_currency_cell(ws_inv.cell(row=curr_row, column=23))
                ws_inv.cell(row=curr_row, column=23).font = FONT_BOLD
                ws_inv.cell(row=curr_row, column=23).fill = PENDING_FILL
                
                c_cns = ws_inv.cell(row=curr_row, column=24, value=cn_str)
                c_cns.font = Font(name='Segoe UI', size=8)
                c_cns.alignment = Alignment(horizontal='center', vertical='center')
                
                c_st = ws_inv.cell(row=curr_row, column=25, value="Active")
                c_st.font = FONT_OK
                c_st.alignment = Alignment(horizontal='center')
            else:
                for c_blank in range(16, 26):
                    ws_inv.cell(row=curr_row, column=c_blank, value="")

            for col_i in range(1, len(headers_inv) + 1):
                ws_inv.cell(row=curr_row, column=col_i).border = THIN_BORDER
                
            curr_row += 1

    auto_fit_columns(ws_inv)

    # -------------------------------------------------------------
    # SHEET 3: CASH DISCOUNT (CD) TRACKER (BASE X & PROMPT PAYMENT)
    # -------------------------------------------------------------
    ws_cd = wb.create_sheet(title='Cash Discount Tracker')
    ws_cd.views.sheetView[0].showGridLines = True
    
    headers_cd = [
        "Month", "Invoice No", "Invoice Date", "Invoice Gross (₹)",
        "1st RD CN No.", "1st RD CN Amt (₹)", "Base 'X' (Gross - RD CN) (₹)",
        "Payment Date", "Payment Ref / NFT", "Turnaround", "CD Rate (%)",
        "Eligible CD on Base X (₹)", "GGL CD CN No.", "Actual CD Received (₹)", "CD Shortfall (₹)", "Status"
    ]
    
    ws_cd.row_dimensions[1].height = 26
    for col_idx, h in enumerate(headers_cd, start=1):
        cell = ws_cd.cell(row=1, column=col_idx, value=h)
        cell.font = FONT_HEADER
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = THIN_BORDER
        
    cd_row = 2
    cd_tracker_list = engine.get_cd_tracker_data()
    
    for c in cd_tracker_list:
        ws_cd.row_dimensions[cd_row].height = 20
        
        ws_cd.cell(row=cd_row, column=1, value=c.get('month_key', '')).alignment = Alignment(horizontal='center')
        ws_cd.cell(row=cd_row, column=2, value=c.get('inv_no', '')).alignment = Alignment(horizontal='center')
        ws_cd.cell(row=cd_row, column=3, value=c.get('inv_date', '')).alignment = Alignment(horizontal='center')
        format_currency_cell(ws_cd.cell(row=cd_row, column=4, value=c.get('inv_total', 0.0)))
        
        ws_cd.cell(row=cd_row, column=5, value=c.get('rd_cn_no', '-')).alignment = Alignment(horizontal='center')
        format_currency_cell(ws_cd.cell(row=cd_row, column=6, value=c.get('rd_cn_amt', 0.0)))
        
        # Base X = Gross - RD CN
        c_basex = ws_cd.cell(row=cd_row, column=7, value=f"=ROUND(D{cd_row}-F{cd_row}, 2)")
        format_currency_cell(c_basex)
        c_basex.font = FONT_BOLD
        
        ws_cd.cell(row=cd_row, column=8, value=c.get('pay_date', '-')).alignment = Alignment(horizontal='center')
        ws_cd.cell(row=cd_row, column=9, value=c.get('pay_ref', '-')).alignment = Alignment(horizontal='left')
        ws_cd.cell(row=cd_row, column=10, value=c.get('turnaround_label', '-')).alignment = Alignment(horizontal='center')
        
        c_rate = ws_cd.cell(row=cd_row, column=11, value=c.get('cd_rate_pct', 3.0))
        c_rate.alignment = Alignment(horizontal='center')
        c_rate.font = FONT_BOLD
        
        # Eligible CD on Base X = G * (K / 100)
        c_elig = ws_cd.cell(row=cd_row, column=12, value=f"=ROUND(G{cd_row}*(K{cd_row}/100), 2)")
        format_currency_cell(c_elig)
        c_elig.font = FONT_BOLD
        
        ws_cd.cell(row=cd_row, column=13, value=c.get('cd_cn_no', '-')).alignment = Alignment(horizontal='center')
        
        c_rcvd = ws_cd.cell(row=cd_row, column=14, value=c.get('cd_received', 0.0))
        format_currency_cell(c_rcvd)
        c_rcvd.font = FONT_BOLD
        
        # Shortfall = L - N
        c_short = ws_cd.cell(row=cd_row, column=15, value=f"=ROUND(L{cd_row}-N{cd_row}, 2)")
        format_currency_cell(c_short)
        c_short.font = FONT_BOLD
        if c.get('shortfall', 0.0) > 10.0:
            c_short.fill = PENDING_FILL
        else:
            c_short.fill = SETTLED_FILL
            
        c_st = ws_cd.cell(row=cd_row, column=16, value=c.get('status', ''))
        c_st.alignment = Alignment(horizontal='center')
        
        for col_i in range(1, len(headers_cd) + 1):
            ws_cd.cell(row=cd_row, column=col_i).border = THIN_BORDER
        cd_row += 1

    # Total CD row
    ws_cd.cell(row=cd_row, column=1, value="TOTAL (FY 2026–27)").font = FONT_BOLD
    format_currency_cell(ws_cd.cell(row=cd_row, column=4, value=f"=SUM(D2:D{cd_row-1})"))
    format_currency_cell(ws_cd.cell(row=cd_row, column=6, value=f"=SUM(F2:F{cd_row-1})"))
    
    t_basex = ws_cd.cell(row=cd_row, column=7, value=f"=SUM(G2:G{cd_row-1})")
    format_currency_cell(t_basex)
    t_basex.font = FONT_BOLD
    
    t_elig = ws_cd.cell(row=cd_row, column=12, value=f"=SUM(L2:L{cd_row-1})")
    format_currency_cell(t_elig)
    t_elig.font = FONT_BOLD
    
    t_rcvd = ws_cd.cell(row=cd_row, column=14, value=f"=SUM(N2:N{cd_row-1})")
    format_currency_cell(t_rcvd)
    t_rcvd.font = Font(name='Segoe UI', size=11, bold=True, color='2F855A')
    
    t_cd = ws_cd.cell(row=cd_row, column=15, value=f"=SUM(O2:O{cd_row-1})")
    format_currency_cell(t_cd)
    t_cd.font = Font(name='Segoe UI', size=11, bold=True, color='C53030')
    t_cd.fill = PENDING_FILL
    
    ws_cd.cell(row=cd_row, column=16, value="PENDING CLAIM").font = FONT_BOLD
    
    for col_i in range(1, len(headers_cd) + 1):
        ws_cd.cell(row=cd_row, column=col_i).border = DOUBLE_BOTTOM_BORDER

    auto_fit_columns(ws_cd)

    # -------------------------------------------------------------
    # SHEET 4: QUARTERLY & YEARLY TONNAGE ARD INCENTIVES
    # -------------------------------------------------------------
    ws_ton = wb.create_sheet(title='Quarterly & Yearly ARD')
    ws_ton.views.sheetView[0].showGridLines = True
    
    ws_ton['A1'] = "GUJARAT GUARDIAN LIMITED — TONNAGE VOLUME INCENTIVES (ARD)"
    ws_ton['A1'].font = FONT_TITLE
    ws_ton['A2'] = "*Note: Quarterly ARDs (earned on quarterly tonnage) and Yearly ARDs (earned on annual tonnage) are volume incentives and are tracked separately from monthly invoice rate difference overbilling."
    ws_ton['A2'].font = Font(name='Segoe UI', size=9.5, italic=True, color='4A5568')
    
    headers_ton = ["Period / Scope", "Credit Note Date", "Document No.", "Description / Tonnage Remark", "Eligible Tonnage", "Rate Incentive", "Credit Amount (₹)", "Status"]
    ws_ton.row_dimensions[4].height = 26
    for c_i, h in enumerate(headers_ton, start=1):
        c = ws_ton.cell(row=4, column=c_i, value=h)
        c.fill = HEADER_FILL
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = THIN_BORDER
        
    tonnage_records = [
        # FY 2025-26 Annual Tonnage Settlement (from user Excel & Ledger 30/01/2026)
        ("FY 2024–25 / 2025 Annual", "30-01-2026", "4588832", "Float discount 0.5 Rs on Ton 58.3 for year 2025", "58.30 Tons", "₹0.50/kg (₹500/ton)", 14275.00, "SETTLED (FY25-26)"),
        ("FY 2024–25 / 2025 Annual", "30-01-2026", "4588835", "Mirror Ton 76.7 Rs 1 discount for year 2025", "76.70 Tons", "₹1.00/kg (₹1000/ton)", 37561.00, "SETTLED (FY25-26)"),
        ("FY 2024–25 / 2025 Annual", "30-01-2026", "4588838", "Float discount 0.5 Rs on Ton 58.3 for year 2025", "58.30 Tons", "₹0.50/kg (₹500/ton)", 14275.00, "SETTLED (FY25-26)"),
        ("FY 2024–25 / 2025 Annual", "30-01-2026", "4588841", "Mirror Ton 76.7 Rs 1 discount for year 2025", "76.70 Tons", "₹1.00/kg (₹1000/ton)", 37561.00, "SETTLED (FY25-26)"),
        # FY 2026-27 Quarterly Tonnage ARDs
        ("FY 2025–26 Q4 (Jan–Mar)", "10-04-2026", "2026600000447", "MIRROR JAN-MAR26 QTRLY ARD", "Jan-Mar 26 Q4 Tonnage", "QTR Booster", 11998.00, "SETTLED (FY26-27)"),
        ("FY 2026–27 Q1 (Apr–Jun)", "16-07-2026", "2026600005291", "QTR FCL ARD APR-JUN26 (Float Clear)", "Apr-Jun 26 Q1 Tonnage", "QTR Booster", 14200.00, "SETTLED (FY26-27)"),
        ("FY 2026–27 Q1 (Apr–Jun)", "16-07-2026", "2026600005413", "QTR MIRROR ARD APR-JUN26 (Mirror)", "Apr-Jun 26 Q1 Tonnage", "QTR Booster", 19354.00, "SETTLED (FY26-27)"),
    ]
    
    t_row = 5
    for sc, dt, doc, desc, ton, r_inc, amt, st in tonnage_records:
        ws_ton.row_dimensions[t_row].height = 20
        ws_ton.cell(row=t_row, column=1, value=sc).alignment = Alignment(horizontal='center')
        ws_ton.cell(row=t_row, column=2, value=dt).alignment = Alignment(horizontal='center')
        ws_ton.cell(row=t_row, column=3, value=doc).alignment = Alignment(horizontal='center')
        ws_ton.cell(row=t_row, column=4, value=desc)
        ws_ton.cell(row=t_row, column=5, value=ton).alignment = Alignment(horizontal='right')
        ws_ton.cell(row=t_row, column=6, value=r_inc).alignment = Alignment(horizontal='center')
        format_currency_cell(ws_ton.cell(row=t_row, column=7, value=amt))
        c_st = ws_ton.cell(row=t_row, column=8, value=st)
        c_st.alignment = Alignment(horizontal='center')
        c_st.font = FONT_OK
        c_st.fill = SUCCESS_FILL
        for col_i in range(1, len(headers_ton) + 1):
            ws_ton.cell(row=t_row, column=col_i).border = THIN_BORDER
        t_row += 1
        
    # Total Tonnage row
    ws_ton.cell(row=t_row, column=1, value="TOTAL TONNAGE INCENTIVES RECEIVED").font = FONT_BOLD
    ws_ton.merge_cells(f"A{t_row}:F{t_row}")
    format_currency_cell(ws_ton.cell(row=t_row, column=7, value=f"=SUM(G5:G{t_row-1})"))
    ws_ton.cell(row=t_row, column=7).font = Font(name='Segoe UI', size=11, bold=True, color='2F855A')
    ws_ton.cell(row=t_row, column=7).fill = SUCCESS_FILL
    for col_i in range(1, len(headers_ton) + 1):
        ws_ton.cell(row=t_row, column=col_i).border = DOUBLE_BOTTOM_BORDER
    auto_fit_columns(ws_ton)

    # -------------------------------------------------------------
    # SHEET 5: GGL OFFICIAL STATEMENT OF ACCOUNT (FY 2026-27)
    # -------------------------------------------------------------
    ws_led = wb.create_sheet(title='GGL Ledger FY 26-27')
    ws_led.views.sheetView[0].showGridLines = True
    
    ws_led['A1'] = "GUJARAT GUARDIAN LIMITED — STATEMENT OF ACCOUNT (FY 2026–27)"
    ws_led['A1'].font = FONT_TITLE
    ws_led['A2'] = "Customer No: 2201046 | MEHTA IMPEX | Period: 01-04-2026 To 16-09-2026"
    ws_led['A2'].font = Font(name='Segoe UI', size=10, italic=True, color='4A5568')
    
    ledger_headers = ["Date", "Type", "Document No.", "Debit Amt (₹)", "Credit Amt (₹)", "Balance (₹)", "Remarks"]
    
    l_row = 4
    for c_i, h in enumerate(ledger_headers, start=1):
        c = ws_led.cell(row=l_row, column=c_i, value=h)
        c.fill = HEADER_FILL
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = THIN_BORDER
        
    entries_26 = (engine.ledger.get('entries') if engine.ledger else []) or []
    curr_l = 5
    for l in entries_26:
        ws_led.row_dimensions[curr_l].height = 20
        ws_led.cell(row=curr_l, column=1, value=l.get('date')).alignment = Alignment(horizontal='center')
        ws_led.cell(row=curr_l, column=2, value=l.get('type')).alignment = Alignment(horizontal='center')
        ws_led.cell(row=curr_l, column=3, value=l.get('doc_no')).alignment = Alignment(horizontal='center')
        format_currency_cell(ws_led.cell(row=curr_l, column=4, value=l.get('dr')))
        format_currency_cell(ws_led.cell(row=curr_l, column=5, value=l.get('cr')))
        c_bal = ws_led.cell(row=curr_l, column=6, value=l.get('balance'))
        format_currency_cell(c_bal)
        c_bal.font = FONT_BOLD
        if (l.get('balance') or 0) < 0:
            c_bal.fill = SUCCESS_FILL
        ws_led.cell(row=curr_l, column=7, value=l.get('remarks'))
        for col_i in range(1, len(ledger_headers) + 1):
            ws_led.cell(row=curr_l, column=col_i).border = THIN_BORDER
        curr_l += 1

    # Closing Balance Highlight Row
    ws_led.cell(row=curr_l, column=1, value="CLOSING BALANCE AS ON 16-09-2026:").font = FONT_TITLE
    ws_led.merge_cells(f"A{curr_l}:E{curr_l}")
    c_close = ws_led.cell(row=curr_l, column=6, value=-318315.72)
    format_currency_cell(c_close)
    c_close.font = Font(name='Segoe UI', size=13, bold=True, color='2F855A')
    c_close.fill = SUCCESS_FILL
    ws_led.cell(row=curr_l, column=7, value="CREDIT BALANCE (GGL OWES MEHTA IMPEX)").font = FONT_BOLD
    for col_i in range(1, len(ledger_headers) + 1):
        ws_led.cell(row=curr_l, column=col_i).border = DOUBLE_BOTTOM_BORDER
    auto_fit_columns(ws_led)

    # -------------------------------------------------------------
    # SHEET 6: GGL OFFICIAL STATEMENT OF ACCOUNT (FY 2025-26)
    # -------------------------------------------------------------
    ws_led25 = wb.create_sheet(title='GGL Ledger FY 25-26')
    ws_led25.views.sheetView[0].showGridLines = True
    
    ws_led25['A1'] = "GUJARAT GUARDIAN LIMITED — STATEMENT OF ACCOUNT (FY 2025–26)"
    ws_led25['A1'].font = FONT_TITLE
    ws_led25['A2'] = "Customer No: D5010928 | MEHTA IMPEX | Period: 01-04-2025 To 31-03-2026"
    ws_led25['A2'].font = Font(name='Segoe UI', size=10, italic=True, color='4A5568')
    
    l_row25 = 4
    for c_i, h in enumerate(ledger_headers, start=1):
        c = ws_led25.cell(row=l_row25, column=c_i, value=h)
        c.fill = HEADER_FILL
        c.font = FONT_HEADER
        c.alignment = Alignment(horizontal='center', vertical='center')
        c.border = THIN_BORDER
        
    entries_25 = []
    if os.path.exists(engine.data_file):
        try:
            with open(engine.data_file) as f:
                d_json = json.load(f)
                entries_25 = d_json.get('ledger_25_26', {}).get('entries', [])
        except:
            pass

    curr_25 = 5
    for l in entries_25:
        ws_led25.row_dimensions[curr_25].height = 20
        ws_led25.cell(row=curr_25, column=1, value=l.get('date')).alignment = Alignment(horizontal='center')
        ws_led25.cell(row=curr_25, column=2, value=l.get('type')).alignment = Alignment(horizontal='center')
        ws_led25.cell(row=curr_25, column=3, value=l.get('doc_no')).alignment = Alignment(horizontal='center')
        format_currency_cell(ws_led25.cell(row=curr_25, column=4, value=l.get('dr')))
        format_currency_cell(ws_led25.cell(row=curr_25, column=5, value=l.get('cr')))
        c_bal = ws_led25.cell(row=curr_25, column=6, value=l.get('balance'))
        format_currency_cell(c_bal)
        c_bal.font = FONT_BOLD
        if (l.get('balance') or 0) < 0:
            c_bal.fill = SUCCESS_FILL
        ws_led25.cell(row=curr_25, column=7, value=l.get('remarks'))
        for col_i in range(1, len(ledger_headers) + 1):
            ws_led25.cell(row=curr_25, column=col_i).border = THIN_BORDER
        curr_25 += 1

    # Closing Balance Highlight Row for FY 25-26
    ws_led25.cell(row=curr_25, column=1, value="CLOSING BALANCE AS ON 31-03-2026:").font = FONT_TITLE
    ws_led25.merge_cells(f"A{curr_25}:E{curr_25}")
    c_close25 = ws_led25.cell(row=curr_25, column=6, value=-113604.00)
    format_currency_cell(c_close25)
    c_close25.font = Font(name='Segoe UI', size=13, bold=True, color='2F855A')
    c_close25.fill = SUCCESS_FILL
    ws_led25.cell(row=curr_25, column=7, value="CARRIED FORWARD TO FY 2026-27 OPENING BALANCE").font = FONT_BOLD
    for col_i in range(1, len(ledger_headers) + 1):
        ws_led25.cell(row=curr_25, column=col_i).border = DOUBLE_BOTTOM_BORDER
    auto_fit_columns(ws_led25)

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)
    return output_path
