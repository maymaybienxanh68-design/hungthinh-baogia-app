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
import re
import sys
from datetime import datetime
from pathlib import Path

# xlsx_lib.py và prep_image.py nằm cùng thư mục scripts/ này
sys.path.insert(0, str(Path(__file__).parent))

try:
    from xlsx_lib import XlsxQuote
    from prep_image import prep_image
except ImportError as e:
    print(f"❌ Không tìm thấy xlsx_lib hoặc prep_image: {e}", file=sys.stderr)
    print("Vui lòng đảm bảo file nằm trong cùng thư mục.", file=sys.stderr)
    sys.exit(1)

def format_vn_money(n):
    """Định dạng số tiền dùng DẤU CHẤM ngăn cách hàng nghìn, vd 850000 -> '850.000'."""
    try:
        n = int(round(float(n)))
    except (ValueError, TypeError):
        return str(n)
    return f"{n:,}".replace(",", ".")


_VN_DIGITS = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
_VN_UNITS = ["", "nghìn", "triệu", "tỷ"]


def _vn_read_3(n, full=True):
    """Đọc 1 nhóm 3 chữ số ra chữ (0-999)."""
    tram, chuc, dv = n // 100, (n % 100) // 10, n % 10
    words = []
    if tram > 0 or full:
        words.append(_VN_DIGITS[tram] + " trăm")
    if chuc == 0:
        if tram > 0 and dv > 0:
            words.append("linh")
    elif chuc == 1:
        words.append("mười")
    else:
        words.append(_VN_DIGITS[chuc] + " mươi")
    if dv > 0:
        if chuc >= 2 and dv == 1:
            words.append("mốt")
        elif chuc >= 1 and dv == 5:
            words.append("lăm")
        else:
            words.append(_VN_DIGITS[dv])
    return " ".join(words)


def so_tien_bang_chu(n):
    """Đọc số tiền VNĐ ra chữ tiếng Việt, vd 1250000 -> 'Một triệu hai trăm năm mươi nghìn đồng'."""
    try:
        n = int(round(float(n)))
    except (ValueError, TypeError):
        return ""
    if n == 0:
        return "Không đồng"
    negative = n < 0
    n = abs(n)

    groups = []
    tmp = n
    while tmp > 0:
        groups.append(tmp % 1000)
        tmp //= 1000
    # groups[0] = hàng đơn vị (0-999), groups[1] = nghìn, groups[2] = triệu, groups[3] = tỷ ...

    parts = []
    for i in range(len(groups) - 1, -1, -1):
        g = groups[i]
        if g == 0:
            continue
        is_first = (i == len(groups) - 1)
        text = _vn_read_3(g, full=not is_first)
        unit = _VN_UNITS[i] if i < len(_VN_UNITS) else ("tỷ " * (i - 2))
        parts.append(text + (" " + unit if unit else ""))

    result = " ".join(parts).strip()
    result = " ".join(result.split())  # gộp khoảng trắng thừa
    result = result[0].upper() + result[1:] + " đồng"
    if negative:
        result = "Âm " + result
    return result


def download_image(url, local_path):
    """Lưu ảnh sản phẩm về local_path. Hỗ trợ 2 dạng:
    - data:image/...;base64,xxxx  (ảnh người dùng tải lên trực tiếp trong app)
    - http(s)://...                (ảnh từ URL bên ngoài)"""
    if not url:
        return None
    if url.startswith('data:'):
        try:
            import base64
            header, b64data = url.split(',', 1)
            raw = base64.b64decode(b64data)
            with open(local_path, 'wb') as f:
                f.write(raw)
            return local_path
        except Exception as e:
            print(f"⚠️  Không thể giải mã ảnh base64: {e}", file=sys.stderr)
            return None

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

    # Mở mẫu (fixed version) - ưu tiên tìm trong scripts/, sau đó templates/
    mau_file = Path(__file__).parent / "Bao_Gia_Mau_FIX.xlsx"
    if not mau_file.exists():
        mau_file = Path(__file__).parent.parent / "templates" / "Bao_Gia_Mau_FIX.xlsx"
    if not mau_file.exists():
        mau_file = Path(__file__).parent / "Bao_Gia_Mẫu.xlsx"
    if not mau_file.exists():
        print(f"❌ Không tìm thấy mẫu (đã thử: {Path(__file__).parent}, {Path(__file__).parent.parent / 'templates'})", file=sys.stderr)
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

    total_amount = 0
    for idx, item in enumerate(data['items'][:8], start=15):  # Max 8 dòng
        row = idx
        qty = item.get('qty', 1) or 0
        price = item.get('price', 0) or 0
        total_amount += qty * price

        wb.set_text(f"C{row}", item.get('name', ''))
        wb.set_text(f"D{row}", item.get('description', ''))
        wb.set_text(f"F{row}", item.get('unit', ''))
        wb.set_number(f"G{row}", qty)
        wb.set_number(f"H{row}", price)
        # Cột Thành tiền (I) đã có sẵn công thức =G*H trong mẫu; ghi cache giá trị
        # để hiện ngay số tiền kể cả khi phần mềm xem không tự tính lại công thức
        wb.set_formula_cache(f"I{row}", qty * price)

        # Tự động xuống dòng + tự co giãn chiều cao theo nội dung (tên/mô tả dài)
        wb.set_row_autoheight(row)
        # Cột tiền: luôn hiện dấu CHẤM ngăn cách hàng nghìn, không phụ thuộc máy/locale
        wb.set_number_format(f"H{row}", "[$-42A]#,##0")
        wb.set_number_format(f"I{row}", "[$-42A]#,##0")

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

    # Ẩn các dòng sản phẩm TRỐNG (STT chưa dùng) để báo giá không còn ô trống.
    # Chỉ ẩn (không xóa) nên công thức tổng và ảnh vẫn nguyên; dòng ẩn không in ra PDF.
    n_items = min(len(data['items']), 8)
    for r in range(15 + n_items, 23):
        try:
            wb.set_row_hidden(r)
        except Exception:
            pass

    # Ẩn cột "Mô tả sản phẩm" (D) và "Hình ảnh" (E) nếu KHÔNG mặt hàng nào có nội dung.
    # Khi ẩn: nới rộng cột "Tên hàng" (C) để GIỮ chiều ngang trang (tiêu đề không bị cắt),
    # và cho các ô thông tin ngân hàng tự co chữ để vừa khung hẹp hơn.
    _items8 = data['items'][:8]
    _C_W = 25.140625      # rộng gốc cột C
    _freed = 0.0
    if not any((it.get('description') or '').strip() for it in _items8):
        try:
            wb.hide_col('D'); _freed += 23.42578125
        except Exception:
            pass
    if not any((it.get('image') or '').strip() for it in _items8):
        try:
            wb.hide_col('E'); _freed += 16.28515625
        except Exception:
            pass
    if _freed:
        try:
            wb.set_col_width('C', round(_C_W + _freed, 4))
        except Exception:
            pass
        for _c in ('D39', 'D40', 'D41'):
            try:
                wb.set_shrink_to_fit(_c)
            except Exception:
                pass

    # Ghi cache giá trị cho dòng "Cộng tiền hàng" (I23 = SUM), VAT (I24, mặc định 0)
    # và "TỔNG CỘNG THANH TOÁN" (I25 = I23+I24) để hiện số ngay, không bị trống
    vat_amount = 0
    grand_total = total_amount + vat_amount
    try:
        wb.set_formula_cache("I23", total_amount)
        wb.set_number("I24", vat_amount)
        wb.set_formula_cache("I25", grand_total)
    except Exception as e:
        print(f"⚠️  Lỗi ghi cache tổng cộng: {e}", file=sys.stderr)

    # Định dạng dấu chấm cho các dòng tổng cộng (Cộng tiền hàng / VAT / Tổng cộng)
    for cell in ("I23", "I24", "I25"):
        try:
            wb.set_number_format(cell, "[$-42A]#,##0")
        except Exception:
            pass

    # Số tiền bằng chữ - điền đầy đủ, hiện trên 1 hàng ngang (không xuống dòng, tự co chữ nếu dài)
    bang_chu = so_tien_bang_chu(total_amount)
    wb.set_text("B27", f"Bằng chữ: {bang_chu}")
    try:
        wb.set_shrink_to_fit("B27")
    except Exception:
        pass

    # Bắt buộc tính lại toàn bộ công thức khi mở file (phòng khi phần mềm nào đó
    # không đọc cache <v>, ví dụ Excel/LibreOffice desktop)
    try:
        wb.force_recalc()
    except Exception:
        pass

    # Lưu Excel - loại bỏ ký tự không hợp lệ trong tên file (/, \, :, ?, ", <, >, |)
    # để tên khách hàng có ký tự đặc biệt không làm hỏng đường dẫn file
    safe_customer = re.sub(r'[\\/:*?"<>|]+', '_', customer).strip() or "KhachHang"
    safe_customer = re.sub(r'\s+', '_', safe_customer)
    out_name = f"BaoGia_{safe_customer}_{datetime.now().strftime('%d%m%Y%H%M%S')}.xlsx"
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
