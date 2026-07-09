#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Export báo giá từ dữ liệu app sang Excel + PDF (có hình ảnh)
Cách dùng:
  1. Mở app_quote_embedded.html → Tạo báo giá → Click "Xuất dữ liệu"
  2. Copy JSON dữ liệu vừa được copy vào clipboard
  3. Paste vào file data.json hoặc trực tiếp paste khi script yêu cầu
  4. Script sẽ xuất: BaoGia_[tên_khách]_[ngày].xlsx + .pdf
"""

import json
import sys
from datetime import datetime
from pathlib import Path

# xlsx_lib.py và prep_image.py nằm cùng thư mục scripts/ này
sys.path.insert(0, str(Path(__file__).parent))

try:
    from xlsx_lib import XlsxQuote
    from prep_image import prep_image
except ImportError:
    print("❌ Không tìm thấy xlsx_lib hoặc prep_image.")
    print("Vui lòng đảm bảo file nằm trong cùng thư mục.")
    sys.exit(1)

def download_image(url, local_path):
    """Download ảnh từ URL về local"""
    if not url or url.startswith('data:'):
        return local_path if url.startswith('data:') else None

    try:
        import urllib.request
        urllib.request.urlretrieve(url, local_path)
        return local_path
    except Exception as e:
        print(f"⚠️  Không thể tải ảnh: {url} - {e}")
        return None

def export_quote(data_json_str):
    """
    Xuất báo giá từ dữ liệu JSON

    Args:
        data_json_str: JSON string từ app
    """
    try:
        data = json.loads(data_json_str)
    except json.JSONDecodeError as e:
        print(f"❌ JSON không hợp lệ: {e}")
        return False

    # Kiểm tra dữ liệu
    if 'items' not in data or len(data['items']) == 0:
        print("❌ Không có sản phẩm trong báo giá")
        return False

    print("=" * 70)
    print("📋 XUẤT BÁOGIÁ")
    print("=" * 70)

    # Thông tin báo giá
    quote_no = data.get('quoteNo', f"BG-{datetime.now().strftime('%d%m%Y%H%M%S')}")
    quote_date = data.get('quoteDate', datetime.now().strftime('%Y-%m-%d'))
    customer = data.get('customer', 'Khách hàng')
    address = data.get('address', '')
    phone = data.get('phone', '')

    # Parse ngày
    try:
        date_obj = datetime.strptime(quote_date, '%Y-%m-%d')
        date_str = f"ngày {date_obj.day} tháng {date_obj.month}  năm {date_obj.year}"
    except:
        date_str = quote_date

    print(f"✅ Báo giá: {quote_no}")
    print(f"✅ Ngày: {date_str}")
    print(f"✅ Khách: {customer}")
    print(f"✅ Sản phẩm: {len(data['items'])}")

    # Mở mẫu (fixed version)
    mau_file = Path(__file__).parent / "Bao_Gia_Mau_FIX.xlsx"
    if not mau_file.exists():
        mau_file = Path(__file__).parent / "Bao_Gia_Mẫu.xlsx"
    if not mau_file.exists():
        print(f"❌ Không tìm thấy mẫu")
        return False

    wb = XlsxQuote(str(mau_file))

    # Cập nhật thông tin
    wb.set_text("B7", f"Số: {quote_no}/2026")
    wb.set_text("B9", f"Kính gửi: {customer}")
    wb.set_text("B10", f"Địa chỉ: {address}")
    wb.set_text("B11", f"Điện thoại: {phone}")
    wb.replace_in_strings("ngày 20 tháng 3  năm 2026", date_str)

    # Điền sản phẩm
    temp_dir = Path(__file__).parent / "temp_images"
    temp_dir.mkdir(exist_ok=True)

    for idx, item in enumerate(data['items'][:8], start=15):  # Max 8 dòng
        row = idx
        wb.set_text(f"C{row}", item.get('name', ''))
        wb.set_text(f"D{row}", item.get('description', ''))
        wb.set_text(f"F{row}", item.get('unit', ''))
        wb.set_number(f"G{row}", item.get('qty', 1))
        wb.set_number(f"H{row}", item.get('price', 0))

        # Xử lý ảnh
        image_url = item.get('image', '')
        if image_url:
            temp_img = temp_dir / f"product_{row}.jpg"
            if download_image(image_url, str(temp_img)):
                try:
                    prep_image(str(temp_img), str(temp_img), bg="white")
                    wb.embed_image(f"E{row}", str(temp_img), width_px=100)
                    print(f"  ✅ Row {row}: {item['name']} (có ảnh)")
                except Exception as e:
                    print(f"  ⚠️  Row {row}: {item['name']} (ảnh lỗi: {e})")
            else:
                print(f"  ✅ Row {row}: {item['name']} (không có ảnh)")
        else:
            print(f"  ✅ Row {row}: {item['name']}")

    # Lưu Excel
    out_name = f"BaoGia_{customer.replace(' ', '_')}_{datetime.now().strftime('%d%m%Y%H%M%S')}.xlsx"
    out_file = Path(__file__).parent / out_name

    wb.save(str(out_file))
    wb.cleanup()

    print(f"\n✅ Lưu Excel: {out_file}")

    # Xuất PDF
    try:
        import subprocess
        pdf_file = str(out_file).replace('.xlsx', '.pdf')
        result = subprocess.run(
            ['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir', str(Path(__file__).parent), str(out_file)],
            capture_output=True, timeout=30
        )
        if result.returncode == 0:
            print(f"✅ Lưu PDF: {pdf_file}")
        else:
            print(f"⚠️  Lỗi PDF: {result.stderr.decode()[:100]}")
    except Exception as e:
        print(f"⚠️  Không thể xuất PDF: {e}")

    print("\n" + "=" * 70)
    print(f"✅ HOÀN THÀNH: {out_file}")
    print("=" * 70)
    return True

if __name__ == "__main__":
    print("📋 EXPORT BÁOGIÁ - App Hưng Thịnh Smart")
    print("=" * 70)

    # Kiểm tra clipboard hoặc file data.json
    data_json = None

    # Cách 1: Kiểm tra file data.json
    data_file = Path(__file__).parent / "quote_data.json"
    if data_file.exists():
        with open(data_file, 'r', encoding='utf-8') as f:
            data_json = f.read()
        print(f"✅ Đọc từ: {data_file}")
    else:
        # Cách 2: Yêu cầu nhập từ clipboard
        print("\n📌 Hướng dẫn:")
        print("1. Mở app_quote_embedded.html")
        print("2. Tạo báo giá + thêm sản phẩm")
        print("3. Click 'Xuất dữ liệu' → Dữ liệu sẽ copy vào clipboard")
        print("4. Paste dữ liệu vào terminal sau đó nhấn Ctrl+D (macOS/Linux) hoặc Ctrl+Z+Enter (Windows)")
        print("\n💬 Paste dữ liệu JSON (Ctrl+D khi xong):")

        try:
            lines = []
            while True:
                line = input()
                lines.append(line)
        except EOFError:
            data_json = '\n'.join(lines)

    if not data_json:
        print("❌ Không có dữ liệu báo giá")
        sys.exit(1)

    # Xuất báo giá
    success = export_quote(data_json)
    sys.exit(0 if success else 1)
