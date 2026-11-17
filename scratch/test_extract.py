import sys, os
sys.path.insert(0, os.path.abspath("."))
from app_tk13895 import extract_invoice_number

s390 = "VTAD2607001_Tạm ứng 20% HĐ0022_22/08/26_SCL NMĐ AN Hội Côn Đảo"
s409 = "Kê khai thuế HĐ0022_20/08/26_VTAD2607001_SCL NMĐ An Hội Côn Đảo"

print("s390:", extract_invoice_number(s390))
print("s409:", extract_invoice_number(s409))
