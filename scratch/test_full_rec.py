import sys, os
sys.path.insert(0, os.path.abspath("."))
import pandas as pd
from app_tk13895 import parse_gl094_excel

open_bal, df = parse_gl094_excel("GL_094_TK13895 (010126-311026).xlsx")

# Import and test modified reconcile
import numpy as np
from app_tk13895 import extract_invoice_number

def test_reconcile(df, open_bal):
    mask_dnb = (
        df['dien_giai'].str.contains(r'điện mua nội bộ|điện mua nb|cp điện mua nội bộ|cp điện mua nb', case=False, na=False) |
        (df['dien_giai'].str.contains(r'khai thuế gtgt tháng 2/2026', case=False, na=False) & (df['co'] > 1e10)) |
        (df['dien_giai'].str.contains(r'hđ1335', case=False, na=False) & (df['co'].abs() > 1e10)) |
        (df['dien_giai'].str.contains(r'hạch toán cp điện mua nội bộ', case=False, na=False))
    )
    df_dnb = df[mask_dnb].copy()
    df_work = df[~mask_dnb].copy()
    
    drop_indices_t1 = set()
    
    # 1a. Trực tiếp cùng số CT phụ AP (+No / -No)
    ap_groups = df_work[df_work['so_ct_phu'] != ''].groupby('so_ct_phu')
    for ap, grp in ap_groups:
        if len(grp) >= 2:
            pos = grp[grp['no'] > 0]
            neg = grp[grp['no'] < 0]
            for _, r_neg in neg.iterrows():
                amt = abs(r_neg['no'])
                match_pos = pos[(pos['no'] == amt) & (~pos.index.isin(drop_indices_t1))]
                if len(match_pos) > 0:
                    drop_indices_t1.add(r_neg.name)
                    drop_indices_t1.add(match_pos.index[0])
                    
    # 1b. Bút toán điều chỉnh cấp đối tượng (GL 1366 sum == 0)
    gl_groups = df_work[(df_work['so_ct_gl'] != '') & (~df_work.index.isin(drop_indices_t1))].groupby('so_ct_gl')
    for gl, grp in gl_groups:
        if len(grp) >= 2 and (grp['no'] < 0).any() and abs(grp['no'].sum()) < 1.0:
            drop_indices_t1.update(grp.index.tolist())
            
    # 1c. AR Credit Memo netting vs Điều chỉnh TK 33193 (sum == 0)
    ar_dc = df_work[df_work['dien_giai'].str.contains(r'đ/c thuế gtgt hóa đơn tiền.*t07/25|đ/c tk 33193 & 13895', case=False, na=False)]
    if len(ar_dc) >= 2 and abs(ar_dc['co'].sum()) < 1.0:
        drop_indices_t1.update(ar_dc.index.tolist())
        
    # 1d. Điều chỉnh lệch lẻ 1 đồng (HĐ 1398: +1 / -1)
    r1398_plus = df_work[(df_work['dien_giai'].str.contains('1398')) & (df_work['no'] == 1)]
    r1398_minus = df_work[(df_work['dien_giai'].str.contains('1398')) & (df_work['no'] == -1)]
    if len(r1398_plus) > 0 and len(r1398_minus) > 0:
        drop_indices_t1.add(r1398_plus.index[0])
        drop_indices_t1.add(r1398_minus.index[0])
        
    # 1e. Bút toán điều chỉnh sau kiểm toán tự bù trừ (CT25010: +1.075.985 / -1.075.985)
    r_audit_p = df_work[(df_work['dien_giai'].str.contains('25010')) & (df_work['no'] == 1075985)]
    r_audit_m = df_work[(df_work['dien_giai'].str.contains('25010')) & (df_work['no'] == -1075985)]
    if len(r_audit_p) > 0 and len(r_audit_m) > 0:
        drop_indices_t1.add(r_audit_p.index[0])
        drop_indices_t1.add(r_audit_m.index[0])
        
    df_t1_netoff = df_work.loc[list(drop_indices_t1)].copy()
    df_active = df_work.drop(index=list(drop_indices_t1)).copy()
    
    # Tầng 2 & 3: Khớp HĐ + Tiền
    df_no = df_active[df_active['no'] > 0].copy()
    df_co = df_active[df_active['co'] > 0].copy()
    
    matched_no = set()
    matched_co = set()
    
    # Tầng 2: 1-1
    for hd, grp_no in df_no.groupby('so_hd'):
        if not hd:
            continue
        grp_co = df_co[(df_co['so_hd'] == hd) & (~df_co.index.isin(matched_co))]
        for no_idx, r_no in grp_no.iterrows():
            amt = r_no['no']
            match = grp_co[(grp_co['co'] == amt) & (~grp_co.index.isin(matched_co))]
            if len(match) > 0:
                matched_no.add(no_idx)
                matched_co.add(match.index[0])
                
    # Tầng 3: Gộp AP
    unmatched_no_t2 = df_no[~df_no.index.isin(matched_no)]
    for hd, grp_no in unmatched_no_t2.groupby('so_hd'):
        if not hd:
            continue
        grp_co = df_co[(df_co['so_hd'] == hd) & (~df_co.index.isin(matched_co))]
        tot_no = grp_no['no'].sum()
        match_co = grp_co[grp_co['co'] == tot_no]
        if len(match_co) > 0:
            matched_no.update(grp_no.index.tolist())
            matched_co.add(match_co.index[0])
            
    # Tầng 4: Gõ nhầm số HĐ
    rem_no = df_no[~df_no.index.isin(matched_no)]
    rem_co = df_co[~df_co.index.isin(matched_co)]
    mistyped_pairs = [
        (130, 212, "HĐ 1404 gõ nhầm thành 460"),
        (449, 570, "HĐ 1296 copy nhầm thành 977"),
        (314, 571, "HĐ 978 copy nhầm thành 977"),
        (510, 578, "HĐ 1423 gõ nhầm thành 1373"),
        (453, 583, "HĐ 485 gõ nhầm thành 458")
    ]
    mistyped_cases = []
    for r_no_row, r_co_row, ghi_chu in mistyped_pairs:
        m_no = rem_no[rem_no['excel_row'] == r_no_row]
        m_co = rem_co[rem_co['excel_row'] == r_co_row]
        if len(m_no) > 0 and len(m_co) > 0:
            matched_no.add(m_no.index[0])
            matched_co.add(m_co.index[0])
            r_no = m_no.iloc[0]
            r_co = m_co.iloc[0]
            amt = r_no['no']
            so_hd_no_display = r_no['so_hd'] if r_no['so_hd'] else "Bóc từ AP"
            so_hd_co_display = r_co['so_hd'] if r_co['so_hd'] else "Nhầm trên GL"
            mistyped_cases.append({
                'no_row': r_no['excel_row'],
                'co_row': r_co['excel_row'],
                'ngay_no': r_no['ngay'].strftime('%d/%m/%Y') if hasattr(r_no['ngay'], 'strftime') else str(r_no['ngay'])[:10],
                'ngay_co': r_co['ngay'].strftime('%d/%m/%Y') if hasattr(r_co['ngay'], 'strftime') else str(r_co['ngay'])[:10],
                'so_hd_no': so_hd_no_display,
                'so_hd_co': so_hd_co_display,
                'so_tien': amt,
                'nguoi_lap_no': r_no['nguoi_lap'],
                'nguoi_lap_co': r_co['nguoi_lap'],
                'dien_giai_no': r_no['dien_giai'],
                'dien_giai_co': r_co['dien_giai'],
                'ghi_chu': ghi_chu
            })
            
    df_mistyped = pd.DataFrame(mistyped_cases)
    df_ton = df_no[~df_no.index.isin(matched_no)].copy()
    df_2025 = df_co[~df_co.index.isin(matched_co)].copy()
    df_da_gt = df_active[df_active.index.isin(matched_no) | df_active.index.isin(matched_co)].copy()
    
    return df_ton, df_2025, df_da_gt, df_mistyped

df_ton, df_2025, df_da_gt, df_mistyped = test_reconcile(df, open_bal)
print(f"df_ton count: {len(df_ton)}, sum: {df_ton['no'].sum():,.0f} đ")
print(f"df_2025 count: {len(df_2025)}, sum: {df_2025['co'].sum():,.0f} đ")
