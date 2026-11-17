import io
import re

with open("app_tk13895.py", "r", encoding="utf-8") as f:
    content = f.read()

# 1. Update docstring of export_excel_report
content = content.replace(
    'Xuất file Excel gồm 6 Sheet chuẩn format in ấn:',
    'Xuất file Excel gồm 7 Sheet chuẩn format in ấn:'
)
content = content.replace(
    '- Sheet 1-5: 5 sheet chuyên đề lọc chi tiết Times New Roman 11-12, Header Navy Blue.',
    '- Sheet 1: 1. HĐ 2026 Tồn Chưa Kê Khai (60 hóa đơn phát sinh năm 2026 chưa giải trừ)\n    - Sheet 2: 2. HĐ 2025 Tồn Chưa Kê Khai (HĐ và số dư nợ đầu kỳ năm 2025 còn tồn chưa giải trừ)\n    - Sheet 3-6: 4 sheet chuyên đề lọc chi tiết Times New Roman 11-12, Header Navy Blue.'
)

# 2. Update Sheet 1 & Add Sheet 2
old_sheet1_pattern = re.compile(
    r'    # ---------------------------------------------------------\n'
    r'    # Sheet 1: Hóa Đơn Tồn Chưa Kê Khai \(Cả 2025 & 2026\)\n'
    r'    # ---------------------------------------------------------\n'
    r'    ws1 = wb\.create_sheet\(title="1\. HĐ Tồn Chưa Kê Khai"\).*?'
    r'    widths1 = \{1: 6, 2: 24, 3: 11, 4: 12, 5: 14, 6: 12, 7: 16, 8: 60, 9: 20, 10: 18\}\n'
    r'    for col_i, w in widths1\.items\(\):\n'
    r'        ws1\.column_dimensions\[get_column_letter\(col_i\)\]\.width = w\n',
    re.DOTALL
)

new_sheet1_and_sheet2 = '''    # ---------------------------------------------------------
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
'''

assert old_sheet1_pattern.search(content) is not None, "old_sheet1_pattern not found!"
content = old_sheet1_pattern.sub(new_sheet1_and_sheet2, content)

# 3. Update titles of remaining sheets
content = content.replace(
    'ws2 = wb.create_sheet(title="2. HĐ 2026 Đã Giải Trừ")',
    'ws2 = wb.create_sheet(title="3. HĐ 2026 Đã Giải Trừ")'
)
content = content.replace(
    'ws3 = wb.create_sheet(title="3. Giải Trừ HĐ 2025")',
    'ws3 = wb.create_sheet(title="4. Giải Trừ HĐ 2025")'
)
content = content.replace(
    'ws4 = wb.create_sheet(title="4. Điện Nội Bộ & ĐC")',
    'ws4 = wb.create_sheet(title="5. Điện Nội Bộ & ĐC")'
)
content = content.replace(
    'ws5 = wb.create_sheet(title="5. Cảnh Báo Gõ Nhầm HĐ")',
    'ws5 = wb.create_sheet(title="6. Cảnh Báo Gõ Nhầm HĐ")'
)

# 4. Update Download Button Text in UI
content = content.replace(
    'Xuất toàn bộ kết quả sang file Excel 6 Sheet: Gồm Sheet 0 giữ nguyên 100% sổ GL_094 gốc và bổ sung 3 Cột (J: Phân loại kiểm dò, K: Số tiền còn tồn thực tế khớp 100% số dư cuối kỳ, L: Ghi chú & Căn cứ đối chiếu) + 5 Sheet chuyên đề lọc chi tiết chuẩn Times New Roman 12.',
    'Xuất toàn bộ kết quả sang file Excel 7 Sheet: Gồm Sheet 0 giữ nguyên 100% sổ GL_094 gốc và bổ sung 3 Cột (J: Phân loại kiểm dò, K: Số tiền còn tồn thực tế khớp 100% số dư cuối kỳ, L: Ghi chú & Căn cứ đối chiếu) + 6 Sheet chuyên đề lọc chi tiết chuẩn Times New Roman 12.'
)
content = content.replace(
    'label="📥 Tải Báo Cáo Excel (6 Sheet Chuẩn)",',
    'label="📥 Tải Báo Cáo Excel (7 Sheet Chuẩn)",'
)

# 5. Update st.tabs in UI
old_tabs = '''    tab0, tab1, tab2, tab3, tab4, tab5 = st.tabs([
        f"📋 Sổ GL_094 Phân Loại ({len(df_raw)})",
        f"📌 Hóa Đơn Tồn Chưa Kê Khai ({count_ton_all})",
        f"✅ HĐ 2026 Đã Giải Trừ ({len(df_da_gt[df_da_gt['co'] > 0])})",
        f"⏳ Giải Trừ HĐ 2025 ({len(df_2025)})",
        f"⚡ Điện Mua Nội Bộ & Điều Chỉnh ({len(df_dnb) + len(df_t1_netoff)})",
        f"⚠️ Cảnh Báo Gõ Nhầm Số HĐ ({len(df_mistyped)})"
    ])'''

new_tabs = '''    tab0, tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        f"📋 Sổ GL_094 Phân Loại ({len(df_raw)})",
        f"📌 HĐ 2026 Tồn Chưa Kê Khai ({len(df_ton)})",
        f"⏳ HĐ 2025 Tồn Chưa Kê Khai ({len(df_ton_2025)})",
        f"✅ HĐ 2026 Đã Giải Trừ ({len(df_da_gt[df_da_gt['co'] > 0])})",
        f"📉 Giải Trừ HĐ 2025 ({len(df_2025)})",
        f"⚡ Điện Mua Nội Bộ & Điều Chỉnh ({len(df_dnb) + len(df_t1_netoff)})",
        f"⚠️ Cảnh Báo Gõ Nhầm Số HĐ ({len(df_mistyped)})"
    ])'''

assert old_tabs in content, "old_tabs not found!"
content = content.replace(old_tabs, new_tabs)

# 6. Update tab1 implementation and add tab2
old_tab1_pattern = re.compile(
    r'    # =========================================================\n'
    r'    # TAB 1: HÓA ĐƠN TỒN CHƯA KÊ KHAI \(NĂM 2025 & NĂM 2026\)\n'
    r'    # =========================================================\n'
    r'    with tab1:.*?'
    r'    # =========================================================\n'
    r'    # TAB 2: HĐ 2026 ĐÃ GIẢI TRỪ\n',
    re.DOTALL
)

new_tab1_and_tab2 = '''    # =========================================================
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
'''

assert old_tab1_pattern.search(content) is not None, "old_tab1_pattern not found!"
content = old_tab1_pattern.sub(new_tab1_and_tab2, content)

# 7. Update tab2, tab3, tab4, tab5 variable names to tab3, tab4, tab5, tab6
content = content.replace('    with tab2:\n        st.markdown(f"### Danh sách hóa đơn năm 2026 đã giải trừ',
                          '    with tab3:\n        st.markdown(f"### Danh sách hóa đơn năm 2026 đã giải trừ')

content = content.replace('    # TAB 3: GIẢI TRỪ HÓA ĐƠN NĂM 2025 (SỐ DƯ ĐẦU KỲ)\n    # =========================================================\n    with tab3:',
                          '    # TAB 4: GIẢI TRỪ HÓA ĐƠN NĂM 2025 (SỐ DƯ ĐẦU KỲ)\n    # =========================================================\n    with tab4:')

content = content.replace('    # TAB 4: ĐIỆN MUA NỘI BỘ & ĐIỀU CHỈNH KỸ THUẬT\n    # =========================================================\n    with tab4:',
                          '    # TAB 5: ĐIỆN MUA NỘI BỘ & ĐIỀU CHỈNH KỸ THUẬT\n    # =========================================================\n    with tab5:')

content = content.replace('    # TAB 5: CẢNH BÁO SAI LỆCH & GÕ NHẦM SỐ HÓA ĐƠN\n    # =========================================================\n    with tab5:',
                          '    # TAB 6: CẢNH BÁO SAI LỆCH & GÕ NHẦM SỐ HÓA ĐƠN\n    # =========================================================\n    with tab6:')

# Verify syntax with AST
import ast
ast.parse(content)
print("AST validation passed! Writing to app_tk13895.py...")

with open("app_tk13895.py", "w", encoding="utf-8") as f:
    f.write(content)

print("Updated app_tk13895.py successfully!")
