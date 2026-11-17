@echo off
chcp 65001 > nul
title Hệ thống Kiểm dò Thuế GTGT TK 13895 - Điện lực Vũng Tàu
echo =====================================================================
echo    HỆ THỐNG KIỂM DÒ & ĐỐI CHIẾU THUẾ GTGT TK 13895 - PC VŨNG TÀU
echo =====================================================================
echo.
echo Đang khởi chạy ứng dụng Streamlit trên trình duyệt...
echo.
python -m streamlit run app_tk13895.py
pause
