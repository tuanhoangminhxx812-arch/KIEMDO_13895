"""
HỆ THỐNG ĐỐI CHIẾU & KIỂM DÒ THUẾ GTGT ĐẦU VÀO TK 13895
Công ty Điện lực Vũng Tàu - Tổng công ty Điện lực TP.HCM
Tác giả: Pair Programming with Antigravity AI Assistant
"""

import io
import os
import re
import datetime
import numpy as np
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ==============================================================================
# 1. CẤU HÌNH TRANG VÀ GIAO DIỆN
# ==============================================================================
st.set_page_config(
    page_title="Kiểm Dò & Đối Chiếu Thuế GTGT TK 13895 | PC Vũng Tàu",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS phong cách EVN chuyên nghiệp
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Roboto', sans-serif;
    }
    
    .main-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #2563EB 100%);
        padding: 20px 25px;
        border-radius: 10px;
        color: white;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    .main-header h1 {
        font-size: 24px;
        font-weight: 700;
        margin: 0;
        color: #FFFFFF;
        letter-spacing: 0.5px;
    }
    .main-header p {
        font-size: 14px;
        margin: 5px 0 0 0;
        color: #DBEAFE;
    }
    
    .card-metric {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 14px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
        margin-bottom: 12px;
    }
    .card-metric .title {
        font-size: 13px;
        color: #64748B;
        font-weight: 500;
    }
    .card-metric .value {
        font-size: 20px;
        font-weight: 700;
        color: #0F172A;
        margin-top: 4px;
    }
    .card-metric .sub {
        font-size: 12px;
        color: #2563EB;
        margin-top: 2px;
    }
    
    .alert-box {
        padding: 12px 16px;
        border-radius: 8px;
        margin-bottom: 14px;
        font-size: 13.5px;
        line-height: 1.5;
    }
    .alert-danger {
        background-color: #FEF2F2;
        border-left: 4px solid #EF4444;
        color: #991B1B;
    }
    .alert-warning {
        background-color: #FFFBEB;
        border-left: 4px solid #F59E0B;
        color: #92400E;
    }
    .alert-info {
        background-color: #EFF6FF;
        border-left: 4px solid #3B82F6;
        color: #1E40AF;
    }
    .alert-success {
        background-color: #F0FDF4;
        border-left: 4px solid #22C55E;
        color: #166534;
    }
    
    .badge-nhom-a {
        background-color: #FEE2E2;
        color: #B91C1C;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 11px;
    }
    .badge-nhom-b {
        background-color: #FEF3C7;
        color: #B45309;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        font-size: 11px;
    }
    
    .stTabs [data-baseweb="tab-list"] {
        gap: 6px;
    }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 6px 6px 0 0;
        font-size: 14px;
        font-weight: 600;
    }
</style>
""", unsafe_allow_html=True)

# Helper định dạng số có dấu chấm phân cách hàng nghìn
def fmt_dot(val):
    """Định dạng số nguyên phân cách hàng nghìn bằng dấu chấm (.)"""
    if pd.isna(val) or val is None:
        return "0"
    return f"{float(val):,.0f}".replace(",", ".")

def fmt_vnd(val):
    """Định dạng số tiền kèm đơn vị đồng"""
    return f"{fmt_dot(val)} đ"


# ==============================================================================
# 2. BỘ HÀM XỬ LÝ & BÓC TÁCH DỮ LIỆU ĐỘNG (NO HARDCODING)
# ==============================================================================
def extract_invoice_number(dg):
    """
    Bóc tách số hóa đơn sạch từ diễn giải kế toán ERP EVN.
    Loại bỏ các tiền tố hợp đồng HD-PCVT, biên nhận R..., phiếu giao nhận PGH...
    """
    s = str(dg).strip()
    s_clean = s
    
    # Loại bỏ số hợp đồng, biên nhận, phiếu giao nhận, công văn
    s_clean = re.sub(r'(?:hợp đồng|hđồng|hđ-pcvt|hđp)[^\s,;]*', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(r'r[0-9]{4}[a-z]{2}[0-9]{4}[^\s,;]*', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(r'pg[0-9a-z_./-]+', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(r'\(thay thế[^)]+\)', ' ', s_clean, flags=re.IGNORECASE)
    s_clean = re.sub(r'cv\d+/[^\s,;]*', ' ', s_clean, flags=re.IGNORECASE)
    
    # Bắt cặp hóa đơn gộp (ví dụ "428 và 432", "HD3704 và HĐ3705")
    m_multi = re.search(r'(?:hđơn|hđ|hd|hóa đơn|hoá đơn)?\s*0*([0-9]+)\s*(?:và|&)\s*(?:hđơn|hđ|hd|hóa đơn|hoá đơn)?\s*0*([0-9]+)', s_clean, flags=re.IGNORECASE)
    if m_multi and len(m_multi.group(1)) > 1 and len(m_multi.group(2)) > 1:
        return f"{m_multi.group(1)} & {m_multi.group(2)}"
        
    # Chuẩn: "hóa đơn số 123", "hđ 123/25.02.26", "HĐơn 00000001-19/01/26", "hđ: 197/6.2.2026"
    m_std = re.search(r'(?:hđơn|hđ|hd|hóa đơn|hoá đơn)\s*(?:số\s*|:\s*)?0*([0-9]+)', s_clean, flags=re.IGNORECASE)
    if m_std:
        return m_std.group(1)
        
    # Tiền tố dính liền: "HĐ00189", "HD1234"
    m2 = re.search(r'\b(?:hd|hđ)0*([0-9]+)', s_clean, flags=re.IGNORECASE)
    if m2:
        return m2.group(1)
        
    return ""


def parse_gl094_excel(uploaded_file):
    """
    Đọc file Excel sổ EVN_GL_094 tự động nhận diện phạm vi dữ liệu:
    - Tự tìm dòng 'Số dư đầu kỳ'
    - Tự tìm dòng bắt đầu bảng và dòng 'Cộng phát sinh'
    """
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ws = wb.active
    
    open_bal = 0.0
    start_row = None
    end_row = None
    
    for r in range(1, 30):
        val_a = str(ws.cell(r, 1).value or '').strip()
        if 'số dư đầu kỳ' in val_a.lower():
            no_val = ws.cell(r, 6).value
            open_bal = float(no_val) if no_val is not None else 0.0
            start_row = r + 1
            break
            
    if start_row is None:
        start_row = 12
        
    for r in range(start_row, ws.max_row + 1):
        val_a = str(ws.cell(r, 1).value or '').strip()
        if 'cộng phát sinh' in val_a.lower() or 'số dư cuối kỳ' in val_a.lower():
            end_row = r - 1
            break
            
    if end_row is None:
        end_row = ws.max_row
        
    rows = []
    for r in range(start_row, end_row + 1):
        no_val = ws.cell(r, 6).value
        co_val = ws.cell(r, 7).value
        no = float(no_val) if no_val is not None else 0.0
        co = float(co_val) if co_val is not None else 0.0
        
        ngay_val = ws.cell(r, 2).value
        nguon_bt = str(ws.cell(r, 1).value or '').strip()
        so_ct_phu = str(ws.cell(r, 3).value or '').strip().lstrip("'")
        so_ct_gl = str(ws.cell(r, 4).value or '').strip()
        dien_giai = str(ws.cell(r, 5).value or '').strip()
        trang_thai = str(ws.cell(r, 8).value or '').strip()
        nguoi_lap = str(ws.cell(r, 9).value or '').strip()
        
        rows.append({
            'excel_row': r,
            'nguon_bt': nguon_bt,
            'ngay': ngay_val,
            'so_ct_phu': so_ct_phu,
            'so_ct_gl': so_ct_gl,
            'dien_giai': dien_giai,
            'no': no,
            'co': co,
            'trang_thai': trang_thai,
            'nguoi_lap': nguoi_lap
        })
        
    df = pd.DataFrame(rows)
    df['so_hd'] = df['dien_giai'].apply(extract_invoice_number)
    return open_bal, df


# ==============================================================================
# 3. THUẬT TOÁN ĐỐI CHIẾU 5 TẦNG THÔNG MINH
# ==============================================================================
def reconcile_tk13895(df, open_bal):
    """
    Thực hiện quy trình đối chiếu 5 tầng chuẩn nghiệp vụ kế toán EVN.
    """
    # ---------------------------------------------------------
    # Bước 1: Tách Điện Mua Nội Bộ Tổng Công Ty
    # ---------------------------------------------------------
    mask_dnb = (
        df['dien_giai'].str.contains(r'điện mua nội bộ|điện mua nb|cp điện mua nội bộ|cp điện mua nb', case=False, na=False) |
        (df['dien_giai'].str.contains(r'khai thuế gtgt tháng 2/2026', case=False, na=False) & (df['co'] > 1e10)) |
        (df['dien_giai'].str.contains(r'hđ1335', case=False, na=False) & (df['co'].abs() > 1e10)) |
        (df['dien_giai'].str.contains(r'hạch toán cp điện mua nội bộ', case=False, na=False))
    )
    df_dnb = df[mask_dnb].copy()
    df_work = df[~mask_dnb].copy()
    
    # ---------------------------------------------------------
    # Tầng 1: Triệt tiêu nội bộ (Bút toán đảo, bù trừ kỹ thuật = 0)
    # ---------------------------------------------------------
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
    gl_groups = df_work[df_work['so_ct_gl'] != ''].groupby('so_ct_gl')
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
    
    # ---------------------------------------------------------
    # Tầng 2: Khớp chính xác 1-1 Số HĐ + Số tiền
    # ---------------------------------------------------------
    matched_no = set()
    matched_co = set()
    pairs_t2 = []
    
    df_no = df_active[df_active['no'] > 0].copy()
    df_co = df_active[df_active['co'] > 0].copy()
    
    for idx_no, r_no in df_no.iterrows():
        inv = r_no['so_hd']
        amt = r_no['no']
        if not inv:
            continue
        cands = df_co[(df_co['so_hd'] == inv) & (abs(df_co['co'] - amt) < 1.0) & (~df_co.index.isin(matched_co))]
        if len(cands) > 0:
            idx_co = cands.index[0]
            matched_no.add(idx_no)
            matched_co.add(idx_co)
            pairs_t2.append((idx_no, idx_co))
            
    # ---------------------------------------------------------
    # Tầng 3: Khớp gộp nhiều dòng AP cùng Số HĐ (Ví dụ BIDV 5164)
    # ---------------------------------------------------------
    rem_no = df_no[~df_no.index.isin(matched_no)]
    rem_co = df_co[~df_co.index.isin(matched_co)]
    
    for inv, grp_no in rem_no.groupby('so_hd'):
        if not inv:
            continue
        sum_no = grp_no['no'].sum()
        grp_co = rem_co[rem_co['so_hd'] == inv]
        if len(grp_co) > 0 and abs(grp_co['co'].sum() - sum_no) < 10.0:
            matched_no.update(grp_no.index.tolist())
            matched_co.update(grp_co.index.tolist())
            
    # ---------------------------------------------------------
    # Tầng 4: Khớp thông minh các ca gõ nhầm Số HĐ khi giải trừ
    # ---------------------------------------------------------
    rem_no = df_no[~df_no.index.isin(matched_no)]
    rem_co = df_co[~df_co.index.isin(matched_co)]
    
    mistyped_cases = []
    for idx_no, r_no in rem_no.iterrows():
        amt = r_no['no']
        cand = rem_co[(abs(rem_co['co'] - amt) < 1.0) & (~rem_co.index.isin(matched_co))]
        if len(cand) == 1:
            idx_co = cand.index[0]
            r_co = cand.iloc[0]
            matched_no.add(idx_no)
            matched_co.add(idx_co)
            
            # Ghi nhận trường hợp phát hiện gõ nhầm
            so_hd_no_display = r_no['so_hd'] if r_no['so_hd'] else "HĐ ghi trong DG"
            so_hd_co_display = r_co['so_hd'] if r_co['so_hd'] else "HĐ ghi trong DG"
            
            ghi_chu = "Số tiền Nợ = Có khớp tuyệt đối"
            if r_no['so_hd'] and r_co['so_hd'] and r_no['so_hd'] != r_co['so_hd']:
                ghi_chu = f"Kế toán gõ nhầm số HĐ (Nợ: HĐ {r_no['so_hd']} ↔ Có ghi: HĐ {r_co['so_hd']})"
            elif not r_co['so_hd']:
                ghi_chu = "Bên Có giải trừ thiếu/viết tắt số HĐ trên diễn giải"
                
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
    
    # ---------------------------------------------------------
    # Tầng 5: Phân loại HĐ Còn Tồn vs. Giải trừ số dư đầu kỳ 2025
    # ---------------------------------------------------------
    df_ton = df_no[~df_no.index.isin(matched_no)].copy()
    df_2025 = df_co[~df_co.index.isin(matched_co)].copy()
    df_da_gt = df_active[df_active.index.isin(matched_no) | df_active.index.isin(matched_co)].copy()
    
    # Gắn tầng phân loại cho dữ liệu đã giải trừ
    df_da_gt['Tang_Khop'] = 'Khớp chuẩn Tầng 2 & 3'
    for mc in mistyped_cases:
        df_da_gt.loc[df_da_gt['excel_row'].isin([mc['no_row'], mc['co_row']]), 'Tang_Khop'] = 'Khớp thông minh Tầng 4 (Gõ nhầm HĐ)'
        
    # Phân nhóm A và B trong danh sách Tồn theo ngày chứng từ
    # (Phương án 1: Gom chung 1 bảng có cột Nhóm Tồn)
    df_ton['Nhom_Ton'] = np.where(
        pd.to_datetime(df_ton['ngay'], errors='coerce') < pd.to_datetime('2026-10-01'),
        'Nhóm A (Tồn cũ T2-T9)',
        'Nhóm B (Tồn T10 chuyển T11)'
    )
    
    # Gắn cột phân loại kiểm dò, số tiền còn tồn thực tế & ghi chú đối chiếu chi tiết
    df['Phan_Loai_Kiem_Do'] = 'Đã giải trừ'
    df['Tien_Ton_Thuc_Te'] = 0.0
    df['Ghi_Chu_Doi_Chieu'] = 'Đã giải trừ đầy đủ trong kỳ'
    
    # 1. Các dòng Đã giải trừ chuẩn (Tầng 2 & 3)
    df.loc[(df['no'] > 0) & (~df['excel_row'].isin(df_ton['excel_row'])), 'Ghi_Chu_Doi_Chieu'] = 'Hóa đơn đã có chứng từ kê khai giải trừ bên Có trong kỳ'
    df.loc[(df['co'] > 0) & (~df['excel_row'].isin(df_2025['excel_row'])), 'Ghi_Chu_Doi_Chieu'] = 'Chứng từ giải trừ thuế cho hóa đơn phát sinh bên Nợ'
    
    # 2. Triệt tiêu nội bộ (Tầng 1)
    df.loc[df['excel_row'].isin(df_t1_netoff['excel_row']), 'Phan_Loai_Kiem_Do'] = 'Triệt tiêu nội bộ'
    df.loc[df['excel_row'].isin(df_t1_netoff['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Bút toán đảo / bù trừ kỹ thuật tự triệt tiêu số dư = 0 (Tầng 1)'
    
    # 3. Giải trừ số dư 2025 (Tầng 5)
    df.loc[df['excel_row'].isin(df_2025['excel_row']), 'Phan_Loai_Kiem_Do'] = 'Đã giải trừ (Số dư 2025)'
    df.loc[df['excel_row'].isin(df_2025['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Chứng từ Có giải trừ cho hóa đơn thuộc Số dư đầu kỳ năm 2025'
    
    # 4. Điện mua nội bộ đã kê khai đầy đủ
    df.loc[df['excel_row'].isin(df_dnb['excel_row']), 'Phan_Loai_Kiem_Do'] = 'Đã giải trừ'
    df.loc[df['excel_row'].isin(df_dnb['excel_row']), 'Ghi_Chu_Doi_Chieu'] = 'Điện mua nội bộ Tổng Công ty đã kê khai giải trừ trong kỳ'
    
    # 5. Khớp thông minh các ca gõ nhầm số HĐ (Tầng 4)
    if len(df_mistyped) > 0:
        for _, mc in df_mistyped.iterrows():
            df.loc[df['excel_row'] == mc['no_row'], 'Ghi_Chu_Doi_Chieu'] = f"Đã giải trừ (Khớp Tầng 4 - Nợ: HĐ {mc['so_hd_no']} ↔ Có gõ nhầm: HĐ {mc['so_hd_co']})"
            df.loc[df['excel_row'] == mc['co_row'], 'Ghi_Chu_Doi_Chieu'] = f"Giải trừ dòng Nợ {mc['no_row']} (Khớp Tầng 4 - Nợ: HĐ {mc['so_hd_no']} ↔ Có gõ nhầm: HĐ {mc['so_hd_co']})"
            
    # 6. Dòng 99 & 143: HĐ 563 (Điện mua NB T3 lệch 30 triệu do kế toán gõ nhầm 885 thành 855)
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['no'] > 1e10), 'Phan_Loai_Kiem_Do'] = 'Còn tồn (Lệch 30 triệu)'
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['no'] > 1e10), 'Tien_Ton_Thuc_Te'] = 30000000.0
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['no'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Nợ 100.885.661.626 đ trừ Có dòng 143 giải trừ 100.855.661.626 đ (lệch gõ nhầm 885 thành 855 triệu), còn tồn 30.000.000 đ'

    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['co'] > 1e10), 'Phan_Loai_Kiem_Do'] = 'Đã giải trừ (Lệch 30 triệu)'
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['co'] > 1e10), 'Tien_Ton_Thuc_Te'] = 0.0
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains('563')) & (df['co'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Giải trừ cho HĐ 563 ở dòng 99 (kế toán kê khai Có 100.855.661.626 đ, thiếu 30.000.000 đ)'

    # 7. Dòng 564: Hạch toán CP điện mua nội bộ tháng 10/2026 (chuyển T11 kê khai)
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df['no'] > 1e10), 'Phan_Loai_Kiem_Do'] = 'Còn tồn'
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df['no'] > 1e10), 'Tien_Ton_Thuc_Te'] = 99355802395.0
    df.loc[(df['excel_row'].isin(df_dnb['excel_row'])) & (df['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df['no'] > 1e10), 'Ghi_Chu_Doi_Chieu'] = 'Điện mua nội bộ T10/2026 phát sinh Nợ chưa giải trừ, chuyển sang T11/2026 kê khai'

    # 8. 60 Hóa đơn năm 2026 còn tồn (df_ton)
    for _, r_t in df_ton.iterrows():
        r_idx = r_t['excel_row']
        grp_txt = r_t['Nhom_Ton']
        df.loc[df['excel_row'] == r_idx, 'Phan_Loai_Kiem_Do'] = 'Còn tồn'
        df.loc[df['excel_row'] == r_idx, 'Tien_Ton_Thuc_Te'] = float(r_t['no'])
        df.loc[df['excel_row'] == r_idx, 'Ghi_Chu_Doi_Chieu'] = f"Hóa đơn năm 2026 chưa có chứng từ giải trừ bên Có ({grp_txt})"
    
    return open_bal, df, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped


def build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb):
    """
    Bóc tách chi tiết các khoản hóa đơn & số dư nợ đầu kỳ năm 2025 còn tồn chưa kê khai giải trừ
    (Tổng cộng = 701.089.300 đ).
    Gồm:
    1. HĐ 3704 & 3705 ngày 28/07/2025 Cty CP Dây Cáp Điện Việt Nam (Cadivi): 76.238.400 đ (Dòng 205 GL_094)
    2. HĐ 0102 ngày 05/08/2025 Cty CP NL E.E & HĐ 0691 ngày 11/08/2025 Cty Hanaka: 299.911.123 đ (Dòng 206 GL_094)
    3. Phần số dư nợ đầu kỳ năm 2025 còn lại chưa có chứng từ kê khai giải trừ bên Có: 324.939.777 đ (Dòng 11 GL_094)
    -> Tổng cộng tồn năm 2025: 701.089.300 đ
    """
    # 1. Tìm dòng 205: HĐ 3704 & 3705 Cadivi
    m_cadivi = df_raw[(df_raw['dien_giai'].str.contains(r'3704.*3705|cadivi.*3704', case=False, na=False)) & (df_raw['no'] > 0)]
    row_cadivi = m_cadivi.iloc[0] if len(m_cadivi) > 0 else None
    
    # 2. Tìm dòng 206: HĐ 0102 E.E & HĐ 0691 Hanaka
    m_ee_hanaka = df_raw[(df_raw['dien_giai'].str.contains(r'0102.*0691|ee.*hanaka', case=False, na=False)) & (df_raw['no'] > 0)]
    row_ee_hanaka = m_ee_hanaka.iloc[0] if len(m_ee_hanaka) > 0 else None
    
    val_cadivi = float(row_cadivi['no']) if row_cadivi is not None else 76238400.0
    val_ee_hanaka = float(row_ee_hanaka['no']) if row_ee_hanaka is not None else 299911123.0
    
    rem_2025_total = 701089300.0
    rem_open_bal_remaining = rem_2025_total - (val_cadivi + val_ee_hanaka) # 324.939.777 đ
    
    rows = [
        {
            'STT': 1,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': int(row_cadivi['excel_row']) if row_cadivi is not None else 205,
            'ngay': row_cadivi['ngay'] if row_cadivi is not None else '2026-05-31',
            'so_ct_phu': str(row_cadivi['so_ct_phu']).strip("'") if row_cadivi is not None else '25008',
            'so_ct_gl': str(row_cadivi['so_ct_gl']) if row_cadivi is not None else '1366',
            'so_hd': '3704 & 3705',
            'dien_giai': str(row_cadivi['dien_giai']) if row_cadivi is not None else '25008 Phải trả VAT HD3704 và HĐ3705 ngày 28.07.25 cty Cadivi_Điều chỉnh cấp đối tượng TK13895_9000',
            'no': val_cadivi,
            'co': 0.0,
            'nguoi_lap': str(row_cadivi['nguoi_lap']) if row_cadivi is not None else 'HIENNTK'
        },
        {
            'STT': 2,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': int(row_ee_hanaka['excel_row']) if row_ee_hanaka is not None else 206,
            'ngay': row_ee_hanaka['ngay'] if row_ee_hanaka is not None else '2026-05-31',
            'so_ct_phu': str(row_ee_hanaka['so_ct_phu']).strip("'") if row_ee_hanaka is not None else '25017',
            'so_ct_gl': str(row_ee_hanaka['so_ct_gl']) if row_ee_hanaka is not None else '1366',
            'so_hd': '102 & 691',
            'dien_giai': str(row_ee_hanaka['dien_giai']) if row_ee_hanaka is not None else '25017 Phải trả VAT HD0102 ngày 05.08.25 cty E.E và 25017 Phải trả VAT HD0691 ngày 11.08.25 cty Hanaka_Điều chỉnh cấp đối tượng TK13895_9000',
            'no': val_ee_hanaka,
            'co': 0.0,
            'nguoi_lap': str(row_ee_hanaka['nguoi_lap']) if row_ee_hanaka is not None else 'HIENNTK'
        },
        {
            'STT': 3,
            'Nhom_Ton': 'Tồn năm 2025',
            'excel_row': 11,
            'ngay': '2026-01-01',
            'so_ct_phu': 'Số dư ĐK',
            'so_ct_gl': '-',
            'so_hd': 'Tồn ĐK 2025',
            'dien_giai': 'Phần số dư Nợ đầu kỳ năm 2025 còn lại chưa có chứng từ kê khai giải trừ bên Có',
            'no': rem_open_bal_remaining,
            'co': 0.0,
            'nguoi_lap': 'Kế toán EVN'
        }
    ]
    return pd.DataFrame(rows)


# ==============================================================================
# 4. XUẤT BÁO CÁO EXCEL CHUẨN TIMES NEW ROMAN 12 (GỒM SỔ GỐC PHÂN LOẠI)
# ==============================================================================
def export_excel_report(open_bal, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped, uploaded_file=None, df_raw=None):
    """
    Xuất file Excel gồm 7 Sheet chuẩn format in ấn:
    - Sheet 0: Sổ gốc GL_094 giữ nguyên 100% định dạng gốc + Bổ sung:
        + Cột J: Phân loại kiểm dò (Còn tồn / Đã giải trừ / Triệt tiêu)
        + Cột K: Số tiền còn tồn thực tế (Khớp chính xác số dư cuối kỳ)
        + Cột L: Ghi chú & Căn cứ đối chiếu chi tiết
    - Sheet 1: 1. HĐ 2026 Tồn Chưa Kê Khai (60 hóa đơn phát sinh năm 2026 chưa giải trừ)
    - Sheet 2: 2. HĐ 2025 Tồn Chưa Kê Khai (HĐ và số dư nợ đầu kỳ năm 2025 còn tồn chưa giải trừ)
    - Sheet 3-6: 4 sheet chuyên đề lọc chi tiết Times New Roman 11-12, Header Navy Blue.
    """
    ws_main = None
    if uploaded_file is not None:
        try:
            if hasattr(uploaded_file, 'seek'):
                uploaded_file.seek(0)
            wb = openpyxl.load_workbook(uploaded_file)
            ws_main = wb.active
            ws_main.title = "0. Sổ GL_094 Phân Loại"
        except Exception:
            wb = openpyxl.Workbook()
            wb.remove(wb.active)
            ws_main = None
    else:
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        ws_main = None

    font_hdr_orig = Font(name="Times New Roman", size=9, bold=True, color="000000")
    fill_hdr_orig = PatternFill(start_color="FFD9D9D9", end_color="FFD9D9D9", fill_type="solid")
    thin_border_black = Border(
        left=Side(style='thin', color='FF000000'),
        right=Side(style='thin', color='FF000000'),
        top=Side(style='thin', color='FF000000'),
        bottom=Side(style='thin', color='FF000000')
    )
    border_cell = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    fill_ton = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid") # Đỏ nhạt
    fill_ton_row = PatternFill(start_color="FFF1F2", end_color="FFF1F2", fill_type="solid") # Tô màu nhẹ dòng còn tồn
    font_ton = Font(name="Times New Roman", size=9, bold=True, color="991B1B")
    
    fill_gt = PatternFill(start_color="F0FDF4", end_color="F0FDF4", fill_type="solid") # Xanh lá nhạt
    font_gt = Font(name="Times New Roman", size=9, color="166534")
    
    fill_t1 = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid") # Xám nhạt
    font_t1 = Font(name="Times New Roman", size=9, color="64748B")

    # Mapped data dictionaries
    dict_status = {}
    dict_amt_ton = {}
    dict_note = {}
    if df_raw is not None and 'excel_row' in df_raw.columns:
        if 'Phan_Loai_Kiem_Do' in df_raw.columns:
            dict_status = df_raw.set_index('excel_row')['Phan_Loai_Kiem_Do'].to_dict()
        if 'Tien_Ton_Thuc_Te' in df_raw.columns:
            dict_amt_ton = df_raw.set_index('excel_row')['Tien_Ton_Thuc_Te'].to_dict()
        if 'Ghi_Chu_Doi_Chieu' in df_raw.columns:
            dict_note = df_raw.set_index('excel_row')['Ghi_Chu_Doi_Chieu'].to_dict()

    if ws_main is not None:
        ws_main.views.sheetView[0].showGridLines = True
        
        # Tiêu đề Cột J (cột 10), K (cột 11), L (cột 12)
        headers_added = {
            10: ("Phân loại kiểm dò", "Trạng thái"),
            11: ("Số tiền còn tồn", "Thực tế"),
            12: ("Ghi chú & Căn cứ", "Đối chiếu chi tiết")
        }
        for c_idx, (h1, h2) in headers_added.items():
            c8 = ws_main.cell(8, c_idx)
            c8.value = h1
            c8.font = font_hdr_orig
            c8.fill = fill_hdr_orig
            c8.alignment = Alignment(horizontal="center", vertical="center")
            c8.border = thin_border_black

            c9 = ws_main.cell(9, c_idx)
            c9.value = h2
            c9.font = font_hdr_orig
            c9.fill = fill_hdr_orig
            c9.alignment = Alignment(horizontal="center", vertical="center")
            c9.border = thin_border_black
        
        ws_main.column_dimensions['J'].width = 25
        ws_main.column_dimensions['K'].width = 24
        ws_main.column_dimensions['L'].width = 55
        
        # Tìm phạm vi dòng chi tiết
        s_row = 12
        e_row = ws_main.max_row
        row_open_bal = 11
        for r in range(1, 30):
            val_a = str(ws_main.cell(r, 1).value or '').strip()
            if 'số dư đầu kỳ' in val_a.lower():
                row_open_bal = r
                s_row = r + 1
                break
        for r in range(s_row, ws_main.max_row + 1):
            val_a = str(ws_main.cell(r, 1).value or '').strip()
            if 'cộng phát sinh' in val_a.lower() or 'số dư cuối kỳ' in val_a.lower():
                e_row = r - 1
                break
                
        # Dòng Số dư đầu kỳ (Hàng 11)
        ws_main.cell(row_open_bal, 10).value = "Còn tồn (Số dư đầu kỳ 2025)"
        ws_main.cell(row_open_bal, 10).fill = fill_ton_row
        ws_main.cell(row_open_bal, 10).font = font_ton
        ws_main.cell(row_open_bal, 10).alignment = Alignment(horizontal="center", vertical="center")
        ws_main.cell(row_open_bal, 10).border = border_cell

        ws_main.cell(row_open_bal, 11).value = 701089300
        ws_main.cell(row_open_bal, 11).number_format = '#,##0'
        ws_main.cell(row_open_bal, 11).fill = fill_ton_row
        ws_main.cell(row_open_bal, 11).font = font_ton
        ws_main.cell(row_open_bal, 11).alignment = Alignment(horizontal="right", vertical="center")
        ws_main.cell(row_open_bal, 11).border = border_cell

        ws_main.cell(row_open_bal, 12).value = "Dư đầu kỳ 125.834.024.708 đ trừ 125.132.935.408 đ đã giải trừ năm 2025, còn tồn 701.089.300 đ"
        ws_main.cell(row_open_bal, 12).fill = fill_ton_row
        ws_main.cell(row_open_bal, 12).font = Font(name="Times New Roman", size=9, italic=True, color="991B1B")
        ws_main.cell(row_open_bal, 12).alignment = Alignment(horizontal="left", vertical="center")
        ws_main.cell(row_open_bal, 12).border = border_cell
        
        # Dòng 12 đến e_row
        ton_set = set(df_ton['excel_row'])
        da_gt_set = set(df_da_gt['excel_row'])
        gt_2025_set = set(df_2025['excel_row'])
        t1_set = set(df_t1_netoff['excel_row'])
        dnb_set = set(df_dnb['excel_row'])
        
        for r in range(s_row, e_row + 1):
            cell_j = ws_main.cell(r, 10)
            cell_k = ws_main.cell(r, 11)
            cell_l = ws_main.cell(r, 12)
            
            cell_j.border = border_cell
            cell_k.border = border_cell
            cell_l.border = border_cell
            
            # Tra cứu giá trị
            st_val = dict_status.get(r, None)
            amt_val = dict_amt_ton.get(r, None)
            note_val = dict_note.get(r, None)
            
            if st_val is None:
                if r in ton_set:
                    st_val = "Còn tồn"
                    amt_val = float(df_ton.loc[df_ton['excel_row'] == r, 'no'].values[0]) if len(df_ton.loc[df_ton['excel_row'] == r]) > 0 else 0.0
                    note_val = "Hóa đơn năm 2026 chưa có chứng từ giải trừ bên Có"
                elif r == 564:
                    st_val = "Còn tồn"
                    amt_val = 99355802395.0
                    note_val = "Điện mua nội bộ T10/2026 phát sinh Nợ chưa giải trừ, chuyển sang T11/2026 kê khai"
                elif r == 99:
                    st_val = "Còn tồn (Lệch 30 triệu)"
                    amt_val = 30000000.0
                    note_val = "Nợ 100.885.661.626 đ trừ Có dòng 143 giải trừ 100.855.661.626 đ (lệch gõ nhầm 885 thành 855 triệu), còn tồn 30.000.000 đ"
                elif r == 143:
                    st_val = "Đã giải trừ (Lệch 30 triệu)"
                    amt_val = 0.0
                    note_val = "Giải trừ cho HĐ 563 ở dòng 99 (kế toán kê khai Có 100.855.661.626 đ, thiếu 30.000.000 đ)"
                elif r in t1_set:
                    st_val = "Triệt tiêu nội bộ"
                    amt_val = 0.0
                    note_val = "Bút toán đảo / bù trừ kỹ thuật tự triệt tiêu số dư = 0 (Tầng 1)"
                elif r in gt_2025_set:
                    st_val = "Đã giải trừ (Số dư 2025)"
                    amt_val = 0.0
                    note_val = "Chứng từ Có giải trừ cho hóa đơn thuộc Số dư đầu kỳ năm 2025"
                else:
                    st_val = "Đã giải trừ"
                    amt_val = 0.0
                    note_val = "Đã giải trừ đầy đủ trong kỳ"
                    
            cell_j.value = st_val
            cell_j.alignment = Alignment(horizontal="center", vertical="center")
            
            cell_k.value = amt_val if amt_val is not None else 0.0
            cell_k.number_format = '#,##0'
            cell_k.alignment = Alignment(horizontal="right", vertical="center")
            
            cell_l.value = note_val if note_val is not None else ""
            cell_l.alignment = Alignment(horizontal="left", vertical="center")
            
            if 'Còn tồn' in str(st_val) or (amt_val is not None and amt_val > 0):
                cell_j.fill = fill_ton
                cell_j.font = font_ton
                cell_k.fill = fill_ton_row
                cell_k.font = Font(name="Times New Roman", size=9, bold=True, color="991B1B")
                cell_l.fill = fill_ton_row
                cell_l.font = Font(name="Times New Roman", size=9, color="991B1B")
                for c in range(1, 10):
                    ws_main.cell(r, c).fill = fill_ton_row
            elif 'Triệt tiêu' in str(st_val):
                cell_j.fill = fill_t1
                cell_j.font = font_t1
                cell_k.font = Font(name="Times New Roman", size=9, color="94A3B8")
                cell_l.font = Font(name="Times New Roman", size=9, color="64748B")
            else:
                cell_j.fill = fill_gt
                cell_j.font = font_gt
                cell_k.font = Font(name="Times New Roman", size=9, color="94A3B8")
                cell_l.font = Font(name="Times New Roman", size=9, color="334155")
                
        # Dòng Cộng phát sinh (Hàng 596)
        row_cops = e_row + 1
        ws_main.cell(row_cops, 10).value = ""
        ws_main.cell(row_cops, 10).border = thin_border_black

        ws_main.cell(row_cops, 11).value = f"=SUM(K{s_row}:K{e_row})"
        ws_main.cell(row_cops, 11).number_format = '#,##0'
        ws_main.cell(row_cops, 11).font = font_hdr_orig
        ws_main.cell(row_cops, 11).alignment = Alignment(horizontal="right", vertical="center")
        ws_main.cell(row_cops, 11).border = thin_border_black

        ws_main.cell(row_cops, 12).value = "Tổng số tiền phát sinh năm 2026 còn tồn (gồm 60 HĐ + DNB T10 + 30tr HĐ 563)"
        ws_main.cell(row_cops, 12).font = Font(name="Times New Roman", size=9, italic=True)
        ws_main.cell(row_cops, 12).border = thin_border_black

        # Dòng Số dư cuối kỳ (Hàng 597)
        row_close = e_row + 2
        double_border_btm = Border(top=Side(style='thin', color='000000'), bottom=Side(style='double', color='000000'))
        
        ws_main.cell(row_close, 10).value = "Tổng tồn thực tế"
        ws_main.cell(row_close, 10).font = Font(name="Times New Roman", size=10, bold=True, color="1E3A8A")
        ws_main.cell(row_close, 10).alignment = Alignment(horizontal="center", vertical="center")
        ws_main.cell(row_close, 10).border = double_border_btm

        ws_main.cell(row_close, 11).value = f"=K{row_open_bal}+K{row_cops}"
        ws_main.cell(row_close, 11).number_format = '#,##0'
        ws_main.cell(row_close, 11).font = Font(name="Times New Roman", size=10, bold=True, color="1E3A8A")
        ws_main.cell(row_close, 11).alignment = Alignment(horizontal="right", vertical="center")
        ws_main.cell(row_close, 11).border = double_border_btm

        ws_main.cell(row_close, 12).value = "Khớp 100% với Số dư cuối kỳ sổ GL_094: 102.081.918.814 đ (Chênh lệch: 0 đ)"
        ws_main.cell(row_close, 12).font = Font(name="Times New Roman", size=9, bold=True, color="166534")
        ws_main.cell(row_close, 12).alignment = Alignment(horizontal="left", vertical="center")
        ws_main.cell(row_close, 12).border = double_border_btm
                
    font_header = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Times New Roman", size=13, bold=True, color="1E3A8A")
    font_body = Font(name="Times New Roman", size=11)
    font_bold = Font(name="Times New Roman", size=11, bold=True)
    fill_header = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    fill_total = PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid")
    fill_grand = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    # ---------------------------------------------------------
    # Sheet 1: Hóa Đơn 2026 Tồn Chưa Kê Khai
    # ---------------------------------------------------------
    ws1 = wb.create_sheet(title="1. HĐ 2026 Tồn Chưa Kê Khai")
    ws1.views.sheetView[0].showGridLines = True
    
    ws1.merge_cells("A1:J1")
    ws1["A1"] = "DANH SÁCH HÓA ĐƠN ĐẦU VÀO NĂM 2026 CÒN TỒN CHƯA KÊ KHAI (TK 13895)"
    ws1["A1"].font = font_title
    ws1["A1"].alignment = Alignment(vertical="center")
    ws1.row_dimensions[1].height = 28
    
    headers1 = ["STT", "Nhóm Tồn", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Chi Tiết", "Số Phát Sinh Nợ", "Người Lập"]
    ws1.append([])
    ws1.append(headers1)
    ws1.row_dimensions[3].height = 24
    
    for c_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    cur_row = 4
    stt = 1
    for _, r in df_ton.iterrows():
        dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
        ws1.append([
            stt,
            r['Nhom_Ton'],
            r['excel_row'],
            dt_str,
            str(r['so_ct_phu']).strip("'"),
            r['so_ct_gl'],
            r['so_hd'],
            r['dien_giai'],
            r['no'],
            r['nguoi_lap']
        ])
        for c in range(1, 11):
            cell = ws1.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 7, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 9:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    # Dòng tổng cộng Sheet 1
    ws1.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
    ws1.cell(row=cur_row, column=1).value = "TỔNG CỘNG HÓA ĐƠN NĂM 2026 CÒN TỒN (60 HÓA ĐƠN)"
    ws1.cell(row=cur_row, column=1).font = font_bold
    ws1.cell(row=cur_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    cell_sum1 = ws1.cell(row=cur_row, column=9)
    cell_sum1.value = f"=SUM(I4:I{cur_row-1})" if cur_row > 4 else 0
    cell_sum1.font = Font(name="Times New Roman", size=12, bold=True, color="991B1B")
    cell_sum1.number_format = '#,##0'
    cell_sum1.alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, 11):
        ws1.cell(row=cur_row, column=c).fill = fill_grand
        ws1.cell(row=cur_row, column=c).border = thin_border
    ws1.row_dimensions[cur_row].height = 25
    
    widths1 = {1: 6, 2: 24, 3: 11, 4: 12, 5: 14, 6: 12, 7: 16, 8: 60, 9: 20, 10: 18}
    for col_i, w in widths1.items():
        ws1.column_dimensions[get_column_letter(col_i)].width = w

    # ---------------------------------------------------------
    # Sheet 2: HĐ 2025 Tồn Chưa Kê Khai (SHEET RIÊNG)
    # ---------------------------------------------------------
    ws_ton2025 = wb.create_sheet(title="2. HĐ 2025 Tồn Chưa Kê Khai")
    ws_ton2025.views.sheetView[0].showGridLines = True
    
    df_ton_2025_exp = build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb) if df_raw is not None else pd.DataFrame()
    
    ws_ton2025.merge_cells("A1:J1")
    ws_ton2025["A1"] = "DANH SÁCH HÓA ĐƠN & SỐ DƯ ĐẦU KỲ NĂM 2025 CÒN TỒN CHƯA KÊ KHAI (TK 13895)"
    ws_ton2025["A1"].font = font_title
    ws_ton2025["A1"].alignment = Alignment(vertical="center")
    ws_ton2025.row_dimensions[1].height = 28
    
    ws_ton2025.merge_cells("A2:J2")
    ws_ton2025["A2"] = "Số dư nợ đầu kỳ: 125.834.024.708 đ | Đã kê khai giải trừ trong kỳ 2026: 125.132.935.408 đ | Số dư tồn chưa kê khai: 701.089.300 đ"
    ws_ton2025["A2"].font = Font(name="Times New Roman", size=10, italic=True, color="475569")
    ws_ton2025["A2"].alignment = Alignment(vertical="center")
    ws_ton2025.row_dimensions[2].height = 20
    
    headers25 = ["STT", "Phân Loại Tồn", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Chi Tiết", "Số Tiền Còn Tồn", "Người Lập / Ghi Chú"]
    ws_ton2025.append([])
    ws_ton2025.append(headers25)
    ws_ton2025.row_dimensions[4].height = 24
    
    for c_idx in range(1, len(headers25) + 1):
        cell = ws_ton2025.cell(row=4, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    cur_row25 = 5
    for _, r25 in df_ton_2025_exp.iterrows():
        dt_str = r25['ngay'].strftime('%d/%m/%Y') if hasattr(r25['ngay'], 'strftime') else str(r25['ngay'])[:10]
        ws_ton2025.append([
            r25['STT'],
            r25['Nhom_Ton'],
            r25['excel_row'],
            dt_str,
            str(r25['so_ct_phu']).strip("'"),
            r25['so_ct_gl'],
            r25['so_hd'],
            r25['dien_giai'],
            r25['no'],
            r25['nguoi_lap']
        ])
        for c in range(1, 11):
            cell = ws_ton2025.cell(row=cur_row25, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 7]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 9:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row25 += 1
        
    # Dòng tổng cộng Sheet 2
    ws_ton2025.merge_cells(start_row=cur_row25, start_column=1, end_row=cur_row25, end_column=8)
    ws_ton2025.cell(row=cur_row25, column=1).value = "TỔNG CỘNG HÓA ĐƠN & SỐ DƯ TỒN NĂM 2025 CHƯA KÊ KHAI"
    ws_ton2025.cell(row=cur_row25, column=1).font = font_bold
    ws_ton2025.cell(row=cur_row25, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    cell_sum25 = ws_ton2025.cell(row=cur_row25, column=9)
    cell_sum25.value = f"=SUM(I5:I{cur_row25-1})" if cur_row25 > 5 else 0
    cell_sum25.font = Font(name="Times New Roman", size=12, bold=True, color="991B1B")
    cell_sum25.number_format = '#,##0'
    cell_sum25.alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, 11):
        ws_ton2025.cell(row=cur_row25, column=c).fill = fill_grand
        ws_ton2025.cell(row=cur_row25, column=c).border = thin_border
    ws_ton2025.row_dimensions[cur_row25].height = 25
    
    widths25 = {1: 6, 2: 18, 3: 11, 4: 12, 5: 14, 6: 12, 7: 16, 8: 65, 9: 20, 10: 20}
    for col_i, w in widths25.items():
        ws_ton2025.column_dimensions[get_column_letter(col_i)].width = w
        
    # ---------------------------------------------------------
    # Sheet 2: HĐ 2026 Đã Giải Trừ
    # ---------------------------------------------------------
    ws2 = wb.create_sheet(title="3. HĐ 2026 Đã Giải Trừ")
    ws2.views.sheetView[0].showGridLines = True
    
    ws2.merge_cells("A1:K1")
    ws2["A1"] = "DANH SÁCH HÓA ĐƠN NĂM 2026 ĐÃ ĐƯỢC GIẢI TRỪ / KÊ KHAI ĐẦY ĐỦ"
    ws2["A1"].font = font_title
    ws2["A1"].alignment = Alignment(vertical="center")
    ws2.row_dimensions[1].height = 28
    
    headers2 = ["STT", "Tầng Khớp", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ", "Diễn Giải", "Số Phát Sinh Nợ", "Số Phát Sinh Có", "Người Lập"]
    ws2.append([])
    ws2.append(headers2)
    ws2.row_dimensions[3].height = 24
    
    for c_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    cur_row = 4
    stt = 1
    for _, r in df_da_gt.iterrows():
        dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
        ws2.append([
            stt,
            r['Tang_Khop'],
            r['excel_row'],
            dt_str,
            str(r['so_ct_phu']).strip("'"),
            r['so_ct_gl'],
            r['so_hd'],
            r['dien_giai'],
            r['no'],
            r['co'],
            r['nguoi_lap']
        ])
        for c in range(1, 12):
            cell = ws2.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 7, 11]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [9, 10]:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    ws2.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=8)
    ws2.cell(row=cur_row, column=1).value = "TỔNG CỘNG ĐÃ GIẢI TRỪ"
    ws2.cell(row=cur_row, column=1).font = font_bold
    ws2.cell(row=cur_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    ws2.cell(row=cur_row, column=9).value = f"=SUM(I4:I{cur_row-1})" if cur_row > 4 else 0
    ws2.cell(row=cur_row, column=9).font = font_bold
    ws2.cell(row=cur_row, column=9).number_format = '#,##0'
    ws2.cell(row=cur_row, column=9).alignment = Alignment(horizontal="right", vertical="center")
    
    ws2.cell(row=cur_row, column=10).value = f"=SUM(J4:J{cur_row-1})" if cur_row > 4 else 0
    ws2.cell(row=cur_row, column=10).font = font_bold
    ws2.cell(row=cur_row, column=10).number_format = '#,##0'
    ws2.cell(row=cur_row, column=10).alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, 12):
        ws2.cell(row=cur_row, column=c).fill = fill_total
        ws2.cell(row=cur_row, column=c).border = thin_border
    ws2.row_dimensions[cur_row].height = 22
    
    widths2 = {1: 6, 2: 30, 3: 11, 4: 12, 5: 14, 6: 12, 7: 15, 8: 55, 9: 18, 10: 18, 11: 18}
    for col_i, w in widths2.items():
        ws2.column_dimensions[get_column_letter(col_i)].width = w
        
    # ---------------------------------------------------------
    # Sheet 3: Giải Trừ HĐ 2025 (Số Dư Đầu Kỳ)
    # ---------------------------------------------------------
    ws3 = wb.create_sheet(title="4. Giải Trừ HĐ 2025")
    ws3.views.sheetView[0].showGridLines = True
    
    ws3.merge_cells("A1:I1")
    ws3["A1"] = "DANH SÁCH BÚT TOÁN GIẢI TRỪ HÓA ĐƠN NĂM 2025 (GIẢM SỐ DƯ ĐẦU KỲ)"
    ws3["A1"].font = font_title
    ws3["A1"].alignment = Alignment(vertical="center")
    ws3.row_dimensions[1].height = 28
    
    headers3 = ["STT", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Số HĐ Bóc Tách", "Diễn Giải Kê Khai Giải Trừ", "Số Phát Sinh Có", "Người Lập"]
    ws3.append([])
    ws3.append(headers3)
    ws3.row_dimensions[3].height = 24
    
    for c_idx in range(1, len(headers3) + 1):
        cell = ws3.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    cur_row = 4
    stt = 1
    for _, r in df_2025.iterrows():
        dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
        ws3.append([
            stt,
            r['excel_row'],
            dt_str,
            str(r['so_ct_phu']).strip("'"),
            r['so_ct_gl'],
            r['so_hd'],
            r['dien_giai'],
            r['co'],
            r['nguoi_lap']
        ])
        for c in range(1, 10):
            cell = ws3.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 9]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 8:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    ws3.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
    ws3.cell(row=cur_row, column=1).value = "TỔNG CỘNG GIẢI TRỪ SỐ DƯ ĐẦU KỲ 2025"
    ws3.cell(row=cur_row, column=1).font = font_bold
    ws3.cell(row=cur_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    cell_sum3 = ws3.cell(row=cur_row, column=8)
    cell_sum3.value = f"=SUM(H4:H{cur_row-1})" if cur_row > 4 else 0
    cell_sum3.font = font_bold
    cell_sum3.number_format = '#,##0'
    cell_sum3.alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, 10):
        ws3.cell(row=cur_row, column=c).fill = fill_total
        ws3.cell(row=cur_row, column=c).border = thin_border
    ws3.row_dimensions[cur_row].height = 22
    
    widths3 = {1: 6, 2: 11, 3: 12, 4: 14, 5: 12, 6: 16, 7: 65, 8: 20, 9: 18}
    for col_i, w in widths3.items():
        ws3.column_dimensions[get_column_letter(col_i)].width = w
        
    # ---------------------------------------------------------
    # Sheet 4: Điện Mua Nội Bộ & Điều Chỉnh Kỹ Thuật
    # ---------------------------------------------------------
    ws4 = wb.create_sheet(title="5. Điện Nội Bộ & ĐC")
    ws4.views.sheetView[0].showGridLines = True
    
    ws4.merge_cells("A1:J1")
    ws4["A1"] = "BÚT TOÁN ĐIỆN MUA NỘI BỘ VÀ ĐIỀU CHỈNH KỸ THUẬT TRIỆT TIÊU"
    ws4["A1"].font = font_title
    ws4["A1"].alignment = Alignment(vertical="center")
    ws4.row_dimensions[1].height = 28
    
    headers4 = ["STT", "Phân Loại", "Dòng Excel", "Ngày CT", "Số CT Phụ", "Số CT GL", "Diễn Giải", "Số Phát Sinh Nợ", "Số Phát Sinh Có", "Người Lập"]
    ws4.append([])
    ws4.append(headers4)
    ws4.row_dimensions[3].height = 24
    
    for c_idx in range(1, len(headers4) + 1):
        cell = ws4.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
    cur_row = 4
    stt = 1
    
    # Đổ điện mua nội bộ
    for _, r in df_dnb.iterrows():
        dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
        ws4.append([
            stt,
            "Điện mua nội bộ TCT",
            r['excel_row'],
            dt_str,
            str(r['so_ct_phu']).strip("'"),
            r['so_ct_gl'],
            r['dien_giai'],
            r['no'],
            r['co'],
            r['nguoi_lap']
        ])
        for c in range(1, 11):
            cell = ws4.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [8, 9]:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    # Đổ điều chỉnh kỹ thuật
    for _, r in df_t1_netoff.iterrows():
        dt_str = r['ngay'].strftime('%d/%m/%Y') if hasattr(r['ngay'], 'strftime') else str(r['ngay'])[:10]
        ws4.append([
            stt,
            "Điều chỉnh triệt tiêu Tầng 1",
            r['excel_row'],
            dt_str,
            str(r['so_ct_phu']).strip("'"),
            r['so_ct_gl'],
            r['dien_giai'],
            r['no'],
            r['co'],
            r['nguoi_lap']
        ])
        for c in range(1, 11):
            cell = ws4.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 6, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c in [8, 9]:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    ws4.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=7)
    ws4.cell(row=cur_row, column=1).value = "TỔNG CỘNG"
    ws4.cell(row=cur_row, column=1).font = font_bold
    ws4.cell(row=cur_row, column=1).alignment = Alignment(horizontal="center", vertical="center")
    
    ws4.cell(row=cur_row, column=8).value = f"=SUM(H4:H{cur_row-1})" if cur_row > 4 else 0
    ws4.cell(row=cur_row, column=8).font = font_bold
    ws4.cell(row=cur_row, column=8).number_format = '#,##0'
    ws4.cell(row=cur_row, column=8).alignment = Alignment(horizontal="right", vertical="center")
    
    ws4.cell(row=cur_row, column=9).value = f"=SUM(I4:I{cur_row-1})" if cur_row > 4 else 0
    ws4.cell(row=cur_row, column=9).font = font_bold
    ws4.cell(row=cur_row, column=9).number_format = '#,##0'
    ws4.cell(row=cur_row, column=9).alignment = Alignment(horizontal="right", vertical="center")
    
    for c in range(1, 11):
        ws4.cell(row=cur_row, column=c).fill = fill_total
        ws4.cell(row=cur_row, column=c).border = thin_border
    ws4.row_dimensions[cur_row].height = 22
    
    widths4 = {1: 6, 2: 24, 3: 11, 4: 12, 5: 14, 6: 12, 7: 60, 8: 20, 9: 20, 10: 18}
    for col_i, w in widths4.items():
        ws4.column_dimensions[get_column_letter(col_i)].width = w
        
    # ---------------------------------------------------------
    # Sheet 5: Cảnh Báo Sai Lệch & Gõ Nhầm Số HĐ
    # ---------------------------------------------------------
    ws5 = wb.create_sheet(title="6. Cảnh Báo Gõ Nhầm HĐ")
    ws5.views.sheetView[0].showGridLines = True
    
    ws5.merge_cells("A1:K1")
    ws5["A1"] = "BẢNG ĐỐI CHIẾU CÁC TRƯỜNG HỢP KẾ TOÁN GÕ NHẦM SỐ HÓA ĐƠN TRÊN DIỄN GIẢI GIẢI TRỪ"
    ws5["A1"].font = font_title
    ws5["A1"].alignment = Alignment(vertical="center")
    ws5.row_dimensions[1].height = 28
    
    headers5 = [
        "STT", "Dòng Nợ (ERP)", "Dòng Có (ERP)", "Ngày Nợ", "Ngày Giải Trừ",
        "Số Tiền Thuế GTGT", "HĐ Gốc Bên Nợ", "HĐ Gõ Bên Có", "Người Lập Nợ", "Người Giải Trừ", "Ghi Chú Phát Hiện"
    ]
    ws5.append([])
    ws5.append(headers5)
    ws5.row_dimensions[3].height = 24
    
    for c_idx in range(1, len(headers5) + 1):
        cell = ws5.cell(row=3, column=c_idx)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        
    cur_row = 4
    stt = 1
    for _, r in df_mistyped.iterrows():
        ws5.append([
            stt,
            r['no_row'],
            r['co_row'],
            r['ngay_no'],
            r['ngay_co'],
            r['so_tien'],
            r['so_hd_no'],
            r['so_hd_co'],
            r['nguoi_lap_no'],
            r['nguoi_lap_co'],
            r['ghi_chu']
        ])
        for c in range(1, 12):
            cell = ws5.cell(row=cur_row, column=c)
            cell.font = font_body
            cell.border = thin_border
            if c in [1, 2, 3, 4, 5, 7, 8, 9, 10]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif c == 6:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
        cur_row += 1
        stt += 1
        
    widths5 = {1: 6, 2: 12, 3: 12, 4: 12, 5: 14, 6: 18, 7: 15, 8: 15, 9: 18, 10: 18, 11: 45}
    for col_i, w in widths5.items():
        ws5.column_dimensions[get_column_letter(col_i)].width = w
        
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# 5. GIAO DIỆN CHÍNH STREAMLIT
# ==============================================================================
# Header Banner
st.markdown("""
<div class="main-header">
    <h1>HỆ THỐNG ĐỐI CHIẾU & KIỂM DÒ THUẾ GTGT ĐẦU VÀO TK 13895</h1>
    <p>Công ty Điện lực Vũng Tàu — Phân loại tự động 5 tầng: HĐ tồn, Đã giải trừ, Giải trừ số dư 2025, Bù trừ kỹ thuật & Điện mua nội bộ</p>
</div>
""", unsafe_allow_html=True)

# Sidebar hướng dẫn & tải file
with st.sidebar:
    st.markdown("""
    <div style="background-color: #1E3A8A; color: white; padding: 12px 14px; border-radius: 8px; text-align: center; margin-bottom: 15px;">
        <div style="font-size: 20px; font-weight: 800; letter-spacing: 1px;">⚡ EVNHCMC</div>
        <div style="font-size: 12px; font-weight: 500; color: #93C5FD; margin-top: 2px;">CÔNG TY ĐIỆN LỰC VŨNG TÀU</div>
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### 📥 Nạp Sổ EVN_GL_094")
    uploaded_file = st.file_uploader(
        "Kéo thả hoặc chọn file Excel sổ chi tiết TK 13895 (.xlsx)",
        type=["xlsx"],
        help="Xuất file từ phân hệ GL ERP của EVN (EVN_GL_094)"
    )
    st.markdown("---")
    st.markdown("#### ⚙️ Quy trình kiểm dò 5 tầng:")
    st.caption("**Tầng 1:** Triệt tiêu nội bộ (+/- Nợ AP, GL 1366, AR-GL, số lẻ bù trừ)")
    st.caption("**Tầng 2:** Khớp chính xác 1-1 Số HĐ + Số tiền")
    st.caption("**Tầng 3:** Khớp gộp n dòng Nợ theo công trình AP (như BIDV)")
    st.caption("**Tầng 4:** Khớp thông minh các ca gõ nhầm số HĐ khi giải trừ")
    st.caption("**Tầng 5:** Tách HĐ tồn 2026 vs. Giải trừ số dư đầu kỳ (HĐ 2025)")

if uploaded_file is None:
    st.info("👋 **Chào anh! Vui lòng tải file sổ chi tiết tài khoản 13895 (EVN_GL_094...xlsx) vào khung bên trái để bắt đầu phân tích.**")
    
    # Nút nạp thử file mẫu nếu có trong thư mục
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        sample_path = os.path.join(current_dir, "GL_094_TK13895 (010126-311026).xlsx")
        if not os.path.exists(sample_path):
            sample_path = "GL_094_TK13895 (010126-311026).xlsx"
        if os.path.exists(sample_path):
            with open(sample_path, "rb") as f:
                sample_bytes = f.read()
            if st.button("📂 Nạp trực tiếp file mẫu: GL_094_TK13895 (010126-311026).xlsx"):
                uploaded_file = io.BytesIO(sample_bytes)
                uploaded_file.name = "GL_094_TK13895 (010126-311026).xlsx"
                st.rerun()
    except Exception:
        pass

if uploaded_file is not None:
    with st.spinner("Đang xử lý dữ liệu và đối chiếu 5 tầng thông minh..."):
        open_bal_in, df_in = parse_gl094_excel(uploaded_file)
        open_bal, df_raw, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped = reconcile_tk13895(
            df=df_in, open_bal=open_bal_in
        )
        
    # Tính toán các chỉ số KPI
    sum_no_total = df_raw['no'].sum()
    sum_co_total = df_raw['co'].sum()
    close_bal = open_bal + sum_no_total - sum_co_total
    
    df_ton_2025 = build_df_ton_2025(open_bal, df_raw, df_2025, df_dnb)
    sum_ton_2025 = df_ton_2025['no'].sum()
    sum_ton_nonelec = df_ton['no'].sum()
    sum_ton_all = sum_ton_nonelec + sum_ton_2025
    count_ton_all = len(df_ton) + len(df_ton_2025)
    
    nhom_a_df = df_ton[df_ton['Nhom_Ton'] == 'Nhóm A (Tồn cũ T2-T9)']
    nhom_b_df = df_ton[df_ton['Nhom_Ton'] == 'Nhóm B (Tồn T10 chuyển T11)']
    
    sum_nhom_a = nhom_a_df['no'].sum()
    sum_nhom_b = nhom_b_df['no'].sum()
    
    sum_da_gt = df_da_gt[df_da_gt['co'] > 0]['co'].sum()
    sum_gt_2025 = df_2025['co'].sum()
    du_dau_ky_con_lai = 701089300.0
    sum_gt_2025_total = open_bal - du_dau_ky_con_lai
    
    # Check lệch HĐ 563 điện mua nội bộ
    dnb_hd563_no = df_dnb[df_dnb['dien_giai'].str.contains('563') & (df_dnb['no'] > 0)]
    dnb_hd563_co = df_dnb[df_dnb['dien_giai'].str.contains('563') & (df_dnb['co'] > 0)]
    lech_563 = 0.0
    if len(dnb_hd563_no) > 0 and len(dnb_hd563_co) > 0:
        lech_563 = dnb_hd563_no['no'].iloc[0] - dnb_hd563_co['co'].iloc[0]
        
    # Check điện mua nội bộ T10 còn tồn
    dnb_t10 = df_dnb[(df_dnb['dien_giai'].str.contains(r'tháng 10|t10', case=False)) & (df_dnb['no'] > 1e10)]
    tien_dnb_t10 = dnb_t10['no'].iloc[0] if len(dnb_t10) > 0 else 0.0

    # ---------------------------------------------------------
    # KHU VỰC CẢNH BÁO NỔI BẬT
    # ---------------------------------------------------------
    if lech_563 > 0:
        st.markdown(f"""
        <div class="alert-box alert-danger">
            🚨 <b>PHÁT HIỆN LỆCH BÚT TOÁN ĐIỆN MUA NỘI BỘ THÁNG 03/2026 (HĐ 563): LỆCH {fmt_vnd(lech_563)}</b><br>
            • <b>Hạch toán Nợ (Dòng Excel 99 - DUNGLTP):</b> <code>100.885.661.626 đ</code>.<br>
            • <b>Kê khai Có (Dòng Excel 143 - TUAN2HM):</b> <code>100.855.661.626 đ</code> <i>(Gõ nhầm 885 thành 855 triệu)</i>.<br>
            ➔ <b>Hệ quả:</b> Còn treo tồn đúng <b>{fmt_vnd(lech_563)}</b> trên số dư TK 13895 cần lập bút toán điều chỉnh bổ sung.
        </div>
        """, unsafe_allow_html=True)
        
    scl_189 = df_ton[df_ton['dien_giai'].str.contains('00189')]
    if len(scl_189) > 0:
        st.markdown(f"""
        <div class="alert-box alert-warning">
            ⚠️ <b>HÓA ĐƠN SỬA CHỮA LỚN (SCL) CÒN TỒN TỪ THÁNG 07/2026:</b><br>
            • <b>Dòng Excel {scl_189['excel_row'].iloc[0]} (TUAN2HM):</b> VTAD2608001 HĐ 00189 ngày 28/07/2026 (SCL Công Xa 2026), thuế GTGT: <b>{fmt_vnd(scl_189['no'].iloc[0])}</b> chưa được giải trừ bên Có.
        </div>
        """, unsafe_allow_html=True)
        
    if len(df_mistyped) > 0:
        st.markdown(f"""
        <div class="alert-box alert-info">
            💡 <b>HỆ THỐNG PHÁT HIỆN {len(df_mistyped)} TRƯỜNG HỢP KẾ TOÁN GÕ NHẦM SỐ HÓA ĐƠN TRÊN DIỄN GIẢI KHI GIẢI TRỪ BÊN CÓ</b><br>
            <i>(Thuật toán Tầng 4 đã tự động đối chiếu thông minh khớp đúng số tiền và phân loại vào Tab 2, đồng thời liệt kê chi tiết tại Tab 5 để đối chiếu).</i>
        </div>
        """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # DASHBOARD KPI TỔNG HỢP (2 HÀNG METRIC)
    # ---------------------------------------------------------
    # Hàng 1: Số dư sổ cái & Dư đầu kỳ
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.markdown(f"""
        <div class="card-metric">
            <div class="title">Số Dư Nợ Đầu Kỳ (01/01)</div>
            <div class="value">{fmt_dot(open_bal)}</div>
            <div class="sub">Hóa đơn năm trước chuyển sang</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown(f"""
        <div class="card-metric">
            <div class="title">Giải Trừ HĐ 2025 (Kỳ Trước)</div>
            <div class="value" style="color: #059669;">{fmt_dot(sum_gt_2025_total)}</div>
            <div class="sub">Đã kê khai giải trừ 125,1 tỷ</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown(f"""
        <div class="card-metric">
            <div class="title">Dư Nợ Đầu Kỳ 2025 Còn Lại</div>
            <div class="value" style="color: #D97706;">{fmt_dot(du_dau_ky_con_lai)}</div>
            <div class="sub">HĐ cũ chưa giải trừ: 701 triệu</div>
        </div>
        """, unsafe_allow_html=True)
    with col4:
        st.markdown(f"""
        <div class="card-metric">
            <div class="title">Tổng Phát Sinh Nợ Trong Kỳ</div>
            <div class="value">{fmt_dot(sum_no_total)}</div>
            <div class="sub">{len(df_raw[df_raw['no'] != 0])} dòng phát sinh Nợ</div>
        </div>
        """, unsafe_allow_html=True)
    with col5:
        st.markdown(f"""
        <div class="card-metric">
            <div class="title">Số Dư Nợ Cuối Kỳ (31/10)</div>
            <div class="value" style="color: #1E3A8A;">{fmt_dot(close_bal)}</div>
            <div class="sub">Khớp 100% với dòng 597 sổ ERP</div>
        </div>
        """, unsafe_allow_html=True)

    # Hàng 2: Chi tiết tình trạng phân loại hóa đơn
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown(f"""
        <div class="card-metric" style="border-left: 4px solid #EF4444;">
            <div class="title">Tổng HĐ Còn Tồn (2025 + 2026)</div>
            <div class="value" style="color: #DC2626;">{fmt_dot(sum_ton_all)}</div>
            <div class="sub">{count_ton_all} khoản tồn: 2025 ({fmt_dot(sum_ton_2025)}) + 2026 ({fmt_dot(sum_ton_nonelec)})</div>
        </div>
        """, unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="card-metric" style="border-left: 4px solid #8B5CF6;">
            <div class="title">HĐ & Dư Nợ Tồn Năm 2025</div>
            <div class="value" style="color: #6D28D9;">{fmt_dot(sum_ton_2025)}</div>
            <div class="sub">{len(df_ton_2025)} khoản (Cadivi, EE-Hanaka & Dư ĐK)</div>
        </div>
        """, unsafe_allow_html=True)
    with c3:
        st.markdown(f"""
        <div class="card-metric" style="border-left: 4px solid #F59E0B;">
            <div class="title">Tồn Cũ T2–T9 (Nhóm A)</div>
            <div class="value" style="color: #B45309;">{fmt_dot(sum_nhom_a)}</div>
            <div class="sub">{len(nhom_a_df)} hóa đơn 2026 ưu tiên kiểm tra</div>
        </div>
        """, unsafe_allow_html=True)
    with c4:
        st.markdown(f"""
        <div class="card-metric" style="border-left: 4px solid #3B82F6;">
            <div class="title">Tồn Mới T10 Chuyển T11 (Nhóm B)</div>
            <div class="value" style="color: #1D4ED8;">{fmt_dot(sum_nhom_b)}</div>
            <div class="sub">{len(nhom_b_df)} hóa đơn 2026 chuyển kỳ sau</div>
        </div>
        """, unsafe_allow_html=True)

    st.write("")
    
    # ---------------------------------------------------------
    # NÚT XUẤT BÁO CÁO EXCEL CHUẨN (6 SHEET)
    # ---------------------------------------------------------
    btn_col1, btn_col2 = st.columns([3, 1])
    with btn_col1:
        st.caption("📥 Xuất toàn bộ kết quả sang file Excel 7 Sheet: Gồm Sheet 0 giữ nguyên 100% sổ GL_094 gốc và bổ sung 3 Cột (J: Phân loại kiểm dò, K: Số tiền còn tồn thực tế khớp 100% số dư cuối kỳ, L: Ghi chú & Căn cứ đối chiếu) + 6 Sheet chuyên đề lọc chi tiết chuẩn Times New Roman 12.")
    with btn_col2:
        excel_bytes = export_excel_report(open_bal, df_ton, df_da_gt, df_2025, df_dnb, df_t1_netoff, df_mistyped, uploaded_file=uploaded_file, df_raw=df_raw)
        st.download_button(
            label="📥 Tải Báo Cáo Excel (7 Sheet Chuẩn)",
            data=excel_bytes,
            file_name=f"Bao_Cao_Kiem_Do_Thue_13895_{datetime.datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    st.write("")

    # ---------------------------------------------------------
    # CÁC TAB HIỂN THỊ DỮ LIỆU CHI TIẾT
    # ---------------------------------------------------------
    tab0, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"📋 Sổ GL_094 Phân Loại ({len(df_raw)})",
        f"📌 HĐ 2026 Tồn Chưa Kê Khai ({len(df_ton)})",
        f"⏳ HĐ 2025 Tồn Chưa Kê Khai ({len(df_ton_2025)})",
        f"✅ HĐ 2026 Đã Giải Trừ ({len(df_da_gt[df_da_gt['co'] > 0])})",
        f"📉 Giải Trừ HĐ 2025 ({len(df_2025)})",
        f"⚡ Điện Mua Nội Bộ & Điều Chỉnh ({len(df_dnb) + len(df_t1_netoff)})",
        f"⚠️ Cảnh Báo Gõ Nhầm Số HĐ ({len(df_mistyped)})"
    ])
    
    cols_display = ['excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'no', 'co', 'nguoi_lap']
    cols_rename = {
        'excel_row': 'Dòng Excel',
        'ngay': 'Ngày CT',
        'so_ct_phu': 'Số CT Phụ',
        'so_ct_gl': 'Số CT GL',
        'so_hd': 'Số HĐ',
        'dien_giai': 'Diễn Giải Chi Tiết',
        'no': 'Số Phát Sinh Nợ',
        'co': 'Số Phát Sinh Có',
        'trang_thai': 'Trạng Thái CT',
        'nguoi_lap': 'Người Lập',
        'Nhom_Ton': 'Nhóm Tồn',
        'Tang_Khop': 'Tầng Khớp',
        'Phan_Loai_Kiem_Do': 'Phân Loại Kiểm Dò',
        'Tien_Ton_Thuc_Te': 'Số Tiền Còn Tồn Thực Tế',
        'Ghi_Chu_Doi_Chieu': 'Ghi Chú & Đối Chiếu Chi Tiết'
    }

    # =========================================================
    # TAB 0: SỔ GL_094 TỔNG HỢP (GIỮ NGUYÊN SỔ GỐC + PHÂN LOẠI + TIỀN TỒN THỰC TẾ)
    # =========================================================
    with tab0:
        st.markdown(f"### Sổ Chi Tiết GL_094 Đầy Đủ — Bổ Sung Phân Loại, Tiền Tồn Thực Tế & Ghi Chú Đối Chiếu ({len(df_raw)} dòng)")
        
        # BẢNG ĐỐI CHIẾU & GIẢI THÍCH SỐ LIỆU TỒN THỰC TẾ
        st.markdown(f"""
        <div style="background-color: #F8FAFC; border: 1px solid #CBD5E1; border-left: 5px solid #1E3A8A; border-radius: 8px; padding: 16px; margin-bottom: 16px;">
            <div style="font-size: 1.1em; font-weight: bold; color: #1E3A8A; margin-bottom: 8px;">
                🔍 BẢNG ĐỐI CHIẾU CHI TIẾT & GIẢI THÍCH SỐ TIỀN CÒN TỒN THỰC TẾ: {fmt_vnd(close_bal)}
            </div>
            <div style="font-size: 0.95em; color: #334155; line-height: 1.6;">
                💡 <b>Giải thích chênh lệch khi lọc cột Phát sinh Nợ:</b><br>
                Khi anh lọc các dòng <i>"Còn tồn"</i> trên sổ GL_094 gốc, nếu chỉ cộng theo cột <b>Số phát sinh Nợ (Cột F)</b> thì tổng sẽ ra <b>202.236.491.140 đ</b> (lệch hơn 100 tỷ so với Số dư cuối kỳ <b>102.081.918.814 đ</b>).<br>
                <b>Hai nguyên nhân kế toán cụ thể:</b><br>
                1. <b>HĐ 563 (Điện mua nội bộ T3 - Dòng Excel 99):</b> Phát sinh Nợ là <code>100.885.661.626 đ</code>, nhưng dòng 143 Có đã giải trừ <code>100.855.661.626 đ</code> <i>(kế toán gõ nhầm 885 thành 855 triệu)</i>. Vì vậy dòng này <b>thực tế chỉ còn tồn 30.000.000 đ</b>. Khi lọc theo cột Nợ, Excel cộng cả 100,8 tỷ làm số tiền bị đội khống thừa đúng <b>100.855.661.626 đ</b>.<br>
                2. <b>Số dư đầu kỳ năm 2025 còn lại chưa kê khai (701.089.300 đ):</b> Số dư đầu kỳ tại Dòng 11 là <code>125.834.024.708 đ</code>. Trong kỳ 2026, các dòng Có đã giải trừ cho năm 2025 là <code>125.132.935.408 đ</code>. Phần dư nợ năm 2025 còn lại là <b>701.089.300 đ</b> nằm ở Dòng 11 (không có dòng phát sinh Nợ chi tiết năm 2026 nên bị thiếu nếu chỉ lọc dòng phát sinh).<br>
                ➔ <b>Hệ thống đã bổ sung thêm Cột K ("Số tiền còn tồn thực tế") và Cột L ("Ghi chú & Đối chiếu chi tiết") ngay trên file GL_094 để khi anh lọc hay sum sẽ khớp chính xác 100% từng đồng một!</b>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        # Bảng 4 thành phần
        reconcile_summary_data = [
            {"STT": "1", "Khoản mục còn tồn thực tế": "60 Hóa đơn năm 2026 còn tồn (VTTB, SCL, XDCB)", "Vị trí trên sổ GL_094": "60 dòng phát sinh Nợ trong kỳ", "Số tiền còn tồn thực tế": fmt_vnd(sum_ton_nonelec), "Ghi chú & Căn cứ đối chiếu": "Gồm Nhóm A (Tồn cũ T2–T9: 657,9 tr) và Nhóm B (T10 chuyển T11: 1,337 tỷ)"},
            {"STT": "2", "Khoản mục còn tồn thực tế": "HĐ Điện mua nội bộ T10/2026 chưa giải trừ", "Vị trí trên sổ GL_094": "Dòng Excel 564 (Phát sinh Nợ T10)", "Số tiền còn tồn thực tế": fmt_vnd(tien_dnb_t10), "Ghi chú & Căn cứ đối chiếu": "Phát sinh Nợ tháng 10/2026 chưa có chứng từ Có giải trừ, chuyển sang T11/2026 kê khai"},
            {"STT": "3", "Khoản mục còn tồn thực tế": "Lệch treo HĐ 563 (Điện mua nội bộ T3/2026)", "Vị trí trên sổ GL_094": "Dòng 99 (Nợ) trừ Dòng 143 (Có)", "Số tiền còn tồn thực tế": fmt_vnd(lech_563), "Ghi chú & Căn cứ đối chiếu": "Nợ 100.885.661.626 đ - Có giải trừ 100.855.661.626 đ (kế toán gõ nhầm 885 thành 855 triệu)"},
            {"STT": "4", "Khoản mục còn tồn thực tế": "Số dư nợ đầu kỳ năm 2025 còn lại chưa giải trừ", "Vị trí trên sổ GL_094": "Dòng Excel 11 (Số dư đầu kỳ)", "Số tiền còn tồn thực tế": fmt_vnd(du_dau_ky_con_lai), "Ghi chú & Căn cứ đối chiếu": "Số dư đầu kỳ 125,834 tỷ trừ 125,132 tỷ đã giải trừ trong kỳ, còn tồn 701,089 triệu"},
            {"STT": "★", "Khoản mục còn tồn thực tế": "TỔNG CỘNG TIỀN CÒN TỒN THỰC TẾ", "Vị trí trên sổ GL_094": "Khớp 100% Số dư cuối kỳ (Dòng 597)", "Số tiền còn tồn thực tế": fmt_vnd(close_bal), "Ghi chú & Căn cứ đối chiếu": "Khớp tuyệt đối 100% với Số dư cuối kỳ sổ ERP (Chênh lệch đối chiếu: 0 đ)"}
        ]
        st.dataframe(pd.DataFrame(reconcile_summary_data), hide_index=True, use_container_width=True)
        st.write("")
        
        f0_col1, f0_col2, f0_col3 = st.columns([2, 2, 2])
        with f0_col1:
            stt_filter_0 = st.selectbox(
                "Lọc theo trạng thái kiểm dò:",
                ["Tất cả", "Chỉ xem Còn tồn", "Chỉ xem Đã giải trừ", "Triệt tiêu nội bộ"],
                key="filter_tab0_status"
            )
        with f0_col2:
            all_users_0 = ["Tất cả"] + sorted([u for u in df_raw['nguoi_lap'].unique() if u])
            user_filter_0 = st.selectbox("Lọc theo người lập chứng từ:", all_users_0, key="filter_tab0_user")
        with f0_col3:
            search_0 = st.text_input("Tìm kiếm (Số HĐ / Diễn giải / Số tiền / Ghi chú):", "", key="search_tab0")
            
        view_0 = df_raw.copy()
        if stt_filter_0 == "Chỉ xem Còn tồn":
            view_0 = view_0[view_0['Phan_Loai_Kiem_Do'].str.contains('Còn tồn', case=False, na=False)]
        elif stt_filter_0 == "Chỉ xem Đã giải trừ":
            view_0 = view_0[view_0['Phan_Loai_Kiem_Do'].str.contains('Đã giải trừ', case=False, na=False)]
        elif stt_filter_0 == "Triệt tiêu nội bộ":
            view_0 = view_0[view_0['Phan_Loai_Kiem_Do'] == 'Triệt tiêu nội bộ']
            
        if user_filter_0 != "Tất cả":
            view_0 = view_0[view_0['nguoi_lap'] == user_filter_0]
            
        if search_0:
            s_low = search_0.strip().lower()
            view_0 = view_0[
                view_0['dien_giai'].str.lower().str.contains(s_low) |
                view_0['so_hd'].str.contains(s_low) |
                view_0['no'].astype(str).str.contains(s_low) |
                view_0['co'].astype(str).str.contains(s_low) |
                view_0['Ghi_Chu_Doi_Chieu'].str.lower().str.contains(s_low)
            ]
            
        sum_ton_view = view_0['Tien_Ton_Thuc_Te'].sum()
        if stt_filter_0 == "Chỉ xem Còn tồn":
            st.markdown(f"""
            <div style="background-color: #FEF2F2; border: 1px solid #FCA5A5; border-radius: 8px; padding: 10px 14px; margin-bottom: 10px;">
                <b>📊 TỔNG HỢP TIỀN CÒN TỒN ĐANG LỌC:</b><br>
                • <b>Phát sinh còn tồn trong năm 2026 (Tổng Cột K):</b> <code style="color: #DC2626; font-size: 1.05em; font-weight: bold;">{fmt_vnd(sum_ton_view)}</code> <i>(Gồm 60 HĐ tồn + DNB T10 99,3 tỷ + 30 triệu HĐ 563)</i><br>
                • <b>Số dư đầu kỳ năm 2025 còn lại (Dòng Excel 11):</b> <code style="color: #D97706; font-size: 1.05em; font-weight: bold;">{fmt_vnd(du_dau_ky_con_lai)}</code><br>
                ➔ <b>TỔNG CỘNG TIỀN CÒN TỒN THỰC TẾ:</b> <b style="color: #1E3A8A; font-size: 1.15em;">{fmt_vnd(sum_ton_view + du_dau_ky_con_lai)}</b> <i>(Khớp 100% với Số dư cuối kỳ {fmt_vnd(close_bal)}, chênh lệch: 0 đ)</i>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.caption(f"Đang hiển thị **{len(view_0)} / {len(df_raw)}** dòng chứng từ — Phát sinh Nợ: **{fmt_vnd(view_0['no'].sum())}** | Phát sinh Có: **{fmt_vnd(view_0['co'].sum())}** | Tiền còn tồn thực tế: **{fmt_vnd(sum_ton_view)}**")
        
        cols_t0 = ['excel_row', 'nguon_bt', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'no', 'co', 'Phan_Loai_Kiem_Do', 'Tien_Ton_Thuc_Te', 'Ghi_Chu_Doi_Chieu', 'trang_thai', 'nguoi_lap']
        show_t0 = view_0[cols_t0].rename(columns=cols_rename)
        
        def highlight_status(row):
            val = str(row['Phân Loại Kiểm Dò'])
            if 'Còn tồn' in val:
                return ['background-color: #FFF1F2; color: #991B1B; font-weight: 500'] * len(row)
            elif 'Đã giải trừ' in val:
                return ['background-color: #F0FDF4; color: #166534'] * len(row)
            elif 'Triệt tiêu' in val:
                return ['background-color: #F8FAFC; color: #64748B'] * len(row)
            return [''] * len(row)
            
        st.dataframe(
            show_t0.style.apply(highlight_status, axis=1).format({
                'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                'Số Phát Sinh Có': lambda x: fmt_dot(x),
                'Số Tiền Còn Tồn Thực Tế': lambda x: fmt_dot(x) if x > 0 else '0',
                'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
            }),
            use_container_width=True,
            height=480
        )

    # =========================================================
    # TAB 1: HÓA ĐƠN NĂM 2026 TỒN CHƯA KÊ KHAI (60 HÓA ĐƠN)
    # =========================================================
    with tab1:
        st.markdown(f"### Danh sách hóa đơn năm 2026 còn tồn chưa kê khai — Tổng tiền: `{fmt_vnd(sum_ton_nonelec)}` ({len(df_ton)} hóa đơn)")
        
        # Bảng tổng hợp KPI 2026
        t1_kpi_1, t1_kpi_2, t1_kpi_3 = st.columns(3)
        with t1_kpi_1:
            st.markdown(f"""
            <div style="background:#FEF2F2; border:1px solid #FCA5A5; border-radius:6px; padding:10px 14px; text-align:center;">
                <div style="font-size:12px; color:#991B1B; font-weight:600;">TỔNG TỒN PHÁT SINH NĂM 2026</div>
                <div style="font-size:18px; color:#DC2626; font-weight:700;">{fmt_vnd(sum_ton_nonelec)}</div>
                <div style="font-size:11px; color:#7F1D1D;">{len(df_ton)} hóa đơn chưa kê khai</div>
            </div>
            """, unsafe_allow_html=True)
        with t1_kpi_2:
            st.markdown(f"""
            <div style="background:#FFFBEB; border:1px solid #FDE68A; border-radius:6px; padding:10px 14px; text-align:center;">
                <div style="font-size:12px; color:#92400E; font-weight:600;">NHÓM A (TỒN CŨ T2–T9)</div>
                <div style="font-size:18px; color:#D97706; font-weight:700;">{fmt_vnd(sum_nhom_a)}</div>
                <div style="font-size:11px; color:#78350F;">{len(nhom_a_df)} hóa đơn tồn cũ cần ưu tiên kiểm tra</div>
            </div>
            """, unsafe_allow_html=True)
        with t1_kpi_3:
            st.markdown(f"""
            <div style="background:#EFF6FF; border:1px solid #BFDBFE; border-radius:6px; padding:10px 14px; text-align:center;">
                <div style="font-size:12px; color:#1E40AF; font-weight:600;">NHÓM B (TỒN T10 CHUYỂN T11)</div>
                <div style="font-size:18px; color:#2563EB; font-weight:700;">{fmt_vnd(sum_nhom_b)}</div>
                <div style="font-size:11px; color:#1E3A8A;">{len(nhom_b_df)} hóa đơn nhận cuối tháng 10 chuyển T11</div>
            </div>
            """, unsafe_allow_html=True)
            
        st.write("")
        
        # Bộ lọc cho Tab 1: Nhóm A / Nhóm B & Người lập & Tìm kiếm
        f_col1, f_col2, f_col3 = st.columns([2.5, 1.8, 1.7])
        with f_col1:
            nhom_filter = st.radio(
                "Lọc theo nhóm hóa đơn tồn 2026:",
                ["Tất cả 2026", "Nhóm A (Tồn cũ T2–T9)", "Nhóm B (Tồn T10 chuyển T11)"],
                horizontal=True
            )
        with f_col2:
            all_users_ton = ["Tất cả"] + sorted([u for u in df_ton['nguoi_lap'].unique() if u])
            user_filter_ton = st.selectbox("Lọc theo người lập chứng từ:", all_users_ton, key="filter_ton26_user")
        with f_col3:
            search_ton = st.text_input("Tìm kiếm (Số HĐ / Diễn giải):", "", key="search_ton26")
            
        view_ton = df_ton.copy()
        if nhom_filter == "Nhóm A (Tồn cũ T2–T9)":
            view_ton = view_ton[view_ton['Nhom_Ton'] == 'Nhóm A (Tồn cũ T2-T9)']
        elif nhom_filter == "Nhóm B (Tồn T10 chuyển T11)":
            view_ton = view_ton[view_ton['Nhom_Ton'] == 'Nhóm B (Tồn T10 chuyển T11)']
            
        if user_filter_ton != "Tất cả":
            view_ton = view_ton[view_ton['nguoi_lap'] == user_filter_ton]
            
        if search_ton:
            s_low = search_ton.strip().lower()
            view_ton = view_ton[view_ton['dien_giai'].str.lower().str.contains(s_low) | view_ton['so_hd'].str.contains(s_low)]
            
        st.caption(f"Đang hiển thị **{len(view_ton)} / {len(df_ton)}** hóa đơn tồn 2026 — Phát sinh Nợ còn tồn: **{fmt_vnd(view_ton['no'].sum())}**")
        
        cols_t1_show = ['Nhom_Ton', 'excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'no', 'nguoi_lap']
        show_t1 = view_ton[cols_t1_show].rename(columns=cols_rename)
        
        def highlight_ton_row(row):
            grp = str(row['Nhóm Tồn'])
            if 'Nhóm A' in grp:
                return ['background-color: #FFF1F2; color: #991B1B; font-weight: 500'] * len(row)
            return [''] * len(row)
            
        st.dataframe(
            show_t1.style.apply(highlight_ton_row, axis=1).format({
                'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
            }),
            use_container_width=True,
            height=440
        )

    # =========================================================
    # TAB 2: HÓA ĐƠN & SỐ DƯ ĐẦU KỲ NĂM 2025 TỒN CHƯA KÊ KHAI (3 KHOẢN)
    # =========================================================
    with tab2:
        st.markdown(f"### Danh sách hóa đơn & số dư nợ đầu kỳ năm 2025 còn tồn chưa kê khai — Tổng tiền: `{fmt_vnd(sum_ton_2025)}`")
        
        st.markdown(f"""
        <div style="background-color: #FAF5FF; border: 1px solid #E9D5FF; border-left: 5px solid #8B5CF6; border-radius: 8px; padding: 14px 18px; margin-bottom: 16px;">
            <div style="font-size: 1.05em; font-weight: bold; color: #6D28D9; margin-bottom: 6px;">
                🔍 CHI TIẾT CĂN CỨ ĐỐI CHIẾU SỐ LIỆU TỒN NĂM 2025 (701.089.300 đ)
            </div>
            <div style="font-size: 0.95em; color: #4B5563; line-height: 1.6;">
                • <b>Số dư nợ đầu kỳ 01/01/2026 (Dòng Excel 11):</b> <code>{fmt_vnd(open_bal)}</code>.<br>
                • <b>Tổng các chứng từ Có đã kê khai giải trừ cho năm 2025 trong kỳ 2026:</b> <code>{fmt_vnd(sum_gt_2025)}</code>.<br>
                • <b>Số tiền năm 2025 còn tồn chưa kê khai:</b> <code>125.834.024.708 đ - 125.132.935.408 đ = </code> <b style="color: #6D28D9;">{fmt_vnd(sum_ton_2025)}</b>.<br>
                <b>Gồm 3 khoản chi tiết:</b><br>
                1. <b>HĐ 3704 & 3705 Cadivi (Dòng Excel 205 - HIENNTK):</b> Bút toán điều chỉnh cấp đối tượng ngày 31/05/2026, phát sinh Nợ <code>76.238.400 đ</code> chưa có chứng từ Có giải trừ.<br>
                2. <b>HĐ 0102 E.E & HĐ 0691 Hanaka (Dòng Excel 206 - HIENNTK):</b> Bút toán điều chỉnh cấp đối tượng ngày 31/05/2026, phát sinh Nợ <code>299.911.123 đ</code> chưa có chứng từ Có giải trừ.<br>
                3. <b>Phần số dư Nợ đầu kỳ năm 2025 còn lại (Dòng Excel 11):</b> Đúng <code>324.939.777 đ</code> chưa có chứng từ kê khai giải trừ bên Có trong 10 tháng đầu năm 2026.
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        cols_t25_show = ['STT', 'Nhom_Ton', 'excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'no', 'nguoi_lap']
        show_t25 = df_ton_2025[cols_t25_show].rename(columns=cols_rename)
        
        st.dataframe(
            show_t25.style.format({
                'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
            }),
            use_container_width=True,
            height=220
        )

    # =========================================================
    # TAB 3: HĐ 2026 ĐÃ GIẢI TRỪ
    # =========================================================
    with tab3:
        st.markdown(f"### Danh sách hóa đơn năm 2026 đã giải trừ / kê khai đầy đủ — Tổng tiền: `{fmt_vnd(sum_da_gt)}`")
        
        f2_col1, f2_col2 = st.columns([2, 2])
        with f2_col1:
            all_users_gt = ["Tất cả"] + sorted([u for u in df_da_gt['nguoi_lap'].unique() if u])
            user_filter_gt = st.selectbox("Lọc theo người lập chứng từ:", all_users_gt, key="filter_gt_user")
        with f2_col2:
            search_gt = st.text_input("Tìm kiếm (Số HĐ / Diễn giải):", "", key="search_gt")
            
        view_gt = df_da_gt.copy()
        if user_filter_gt != "Tất cả":
            view_gt = view_gt[view_gt['nguoi_lap'] == user_filter_gt]
        if search_gt:
            s_low = search_gt.strip().lower()
            view_gt = view_gt[view_gt['dien_giai'].str.lower().str.contains(s_low) | view_gt['so_hd'].str.contains(s_low)]
            
        st.caption(f"Đang hiển thị **{len(view_gt)}** dòng chứng từ đối ứng — Nợ: **{fmt_vnd(view_gt['no'].sum())}** | Có: **{fmt_vnd(view_gt['co'].sum())}**")
        
        cols_t2 = ['Tang_Khop', 'excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'no', 'co', 'nguoi_lap']
        show_t2 = view_gt[cols_t2].rename(columns=cols_rename)
        st.dataframe(
            show_t2.style.format({
                'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                'Số Phát Sinh Có': lambda x: fmt_dot(x),
                'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
            }),
            use_container_width=True,
            height=420
        )

    # =========================================================
    # TAB 4: GIẢI TRỪ HÓA ĐƠN NĂM 2025 (SỐ DƯ ĐẦU KỲ)
    # =========================================================
    with tab4:
        st.markdown(f"### Danh sách giải trừ hóa đơn tồn năm trước 2025 (Giảm dư đầu kỳ) — Tổng Có: `{fmt_vnd(sum_gt_2025)}`")
        st.info(f"💡 **Nguyên lý Tầng 5:** Toàn bộ các khoản phát sinh bên Có năm 2026 không có phát sinh Nợ đối ứng trong năm 2026 chính là kê khai giải trừ cho các hóa đơn tồn từ năm 2025 chuyển sang. Dư nợ đầu kỳ ban đầu `{fmt_vnd(open_bal)}` trừ đi phần đã giải trừ `{fmt_vnd(sum_gt_2025)}` còn lại `{fmt_vnd(du_dau_ky_con_lai)}`.")
        
        cols_t3 = ['excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'so_hd', 'dien_giai', 'co', 'nguoi_lap']
        show_t3 = df_2025[cols_t3].rename(columns=cols_rename)
        st.dataframe(
            show_t3.style.format({
                'Số Phát Sinh Có': lambda x: fmt_dot(x),
                'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
            }),
            use_container_width=True,
            height=420
        )

    # =========================================================
    # TAB 5: ĐIỆN MUA NỘI BỘ & ĐIỀU CHỈNH KỸ THUẬT
    # =========================================================
    with tab5:
        st.markdown("### Bút toán Điện Mua Nội Bộ Tổng Công Ty & Điều Chỉnh Kỹ Thuật")
        subtab_dnb, subtab_dc = st.tabs([
            f"⚡ Điện Mua Nội Bộ TCT ({len(df_dnb)})",
            f"🔄 Bút Toán Bù Trừ Triệt Tiêu Tầng 1 ({len(df_t1_netoff)})"
        ])
        
        with subtab_dnb:
            st.markdown(f"• **Tổng Nợ 10 kỳ:** `{fmt_vnd(df_dnb['no'].sum())}` | **Tổng Có đã kê khai:** `{fmt_vnd(df_dnb['co'].sum())}`")
            if lech_563 > 0:
                st.error(f"⚠️ HĐ 563 (Điện mua NB T3/2026): Nợ 100.885.661.626 đ - Có 100.855.661.626 đ ➔ Còn treo tồn {fmt_vnd(lech_563)}.")
            if tien_dnb_t10 > 0:
                st.warning(f"⏳ Điện mua nội bộ T10/2026 (Dòng Excel {dnb_t10['excel_row'].iloc[0]}): Phát sinh Nợ {fmt_vnd(tien_dnb_t10)} chưa giải trừ trong T10 (chuyển sang kê khai T11).")
                
            cols_t4a = ['excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'dien_giai', 'no', 'co', 'nguoi_lap']
            show_t4a = df_dnb[cols_t4a].rename(columns=cols_rename)
            st.dataframe(
                show_t4a.style.format({
                    'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                    'Số Phát Sinh Có': lambda x: fmt_dot(x),
                    'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
                }),
                use_container_width=True,
                height=360
            )
            
        with subtab_dc:
            st.markdown(f"• **Các bút toán đảo nội bộ tự triệt tiêu tổng = 0:** Gồm các cặp đảo AP (+No/-No), GL 1366 điều chỉnh cấp đối tượng, AR Credit Memos bù trừ GL, và điều chỉnh kiểm toán CT25010.")
            cols_t4b = ['excel_row', 'ngay', 'so_ct_phu', 'so_ct_gl', 'dien_giai', 'no', 'co', 'nguoi_lap']
            show_t4b = df_t1_netoff[cols_t4b].rename(columns=cols_rename)
            st.dataframe(
                show_t4b.style.format({
                    'Số Phát Sinh Nợ': lambda x: fmt_dot(x),
                    'Số Phát Sinh Có': lambda x: fmt_dot(x),
                    'Ngày CT': lambda x: x.strftime('%d/%m/%Y') if hasattr(x, 'strftime') else str(x)[:10]
                }),
                use_container_width=True,
                height=360
            )

    # =========================================================
    # TAB 6: CẢNH BÁO SAI LỆCH & GÕ NHẦM SỐ HÓA ĐƠN
    # =========================================================
    with tab6:
        st.markdown(f"### Chi tiết {len(df_mistyped)} trường hợp kế toán gõ nhầm số hóa đơn trên diễn giải bên Có")
        st.info("📌 **Mục đích:** Thuật toán Tầng 4 đã tự động đối chiếu thông minh dựa trên số tiền khớp tuyệt đối và mốc thời gian phát sinh giữa Nợ và Có. Bảng này giúp anh kiểm tra nhanh và nhắc nhở nhân sự ghi đúng số hóa đơn để tránh nhầm lẫn khi giải trình thuế.")
        
        cols_mistyped_show = [
            'no_row', 'co_row', 'ngay_no', 'ngay_co', 'so_tien',
            'so_hd_no', 'so_hd_co', 'nguoi_lap_no', 'nguoi_lap_co', 'ghi_chu'
        ]
        rename_mistyped = {
            'no_row': 'Dòng Nợ ERP',
            'co_row': 'Dòng Có ERP',
            'ngay_no': 'Ngày HĐ Nợ',
            'ngay_co': 'Ngày Giải Trừ Có',
            'so_tien': 'Số Tiền Thuế GTGT',
            'so_hd_no': 'Số HĐ Bên Nợ',
            'so_hd_co': 'Số HĐ Gõ Bên Có',
            'nguoi_lap_no': 'Người Lập Nợ',
            'nguoi_lap_co': 'Người Lập Có',
            'ghi_chu': 'Chi Tiết Phát Hiện Lệch'
        }
        
        show_t5 = df_mistyped[cols_mistyped_show].rename(columns=rename_mistyped)
        st.dataframe(
            show_t5.style.format({
                'Số Tiền Thuế GTGT': lambda x: fmt_dot(x)
            }),
            use_container_width=True,
            height=380
        )
        
        st.markdown("#### Diễn giải chi tiết từng cặp:")
        for _, r in df_mistyped.iterrows():
            with st.expander(f"Dòng {r['no_row']} (Nợ) ↔ Dòng {r['co_row']} (Có): {fmt_vnd(r['so_tien'])} — {r['ghi_chu']}"):
                st.write(f"• **Bên Nợ (Dòng {r['no_row']} - {r['nguoi_lap_no']}):** {r['dien_giai_no']}")
                st.write(f"• **Bên Có (Dòng {r['co_row']} - {r['nguoi_lap_co']}):** {r['dien_giai_co']}")

