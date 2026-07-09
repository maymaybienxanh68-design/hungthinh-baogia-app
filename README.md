# 🏭 Hưng Thịnh Smart - Báo giá App

## 📋 Cấu trúc Folder

```
HungThinh_App/
├── 📱 app/                    - App báo giá (chạy offline)
│   └── app_quote_embedded.html
├── 📊 templates/              - Mẫu báo giá Excel
│   └── Bao_Gia_Mau_FIX.xlsx
├── 📤 output/                 - Báo giá đã xuất (PDF/Excel)
├── 📋 data/                   - Dữ liệu sản phẩm, NCC, khách hàng
├── 🛠️  scripts/               - Script Python xuất báo giá
├── 📖 docs/                   - Tài liệu hướng dẫn
└── README.md                  - File này
```

## 🚀 Cách sử dụng

### 1. Chạy App
```
1. Mở D:\HungThinh_App\app\app_quote_embedded.html bằng Chrome
2. Đăng nhập PIN: 0000 (Chủ) hoặc 1111 (Nhân viên)
3. Tạo báo giá, thêm sản phẩm, xuất dữ liệu
```

### 2. Xuất PDF/Excel
```
1. Trong app → Click "Xuất dữ liệu (PDF/Excel)"
2. Dữ liệu copy vào clipboard
3. Mở Command Prompt: cd D:\HungThinh_App
4. Chạy: python scripts\export_quote_to_excel.py
5. Paste dữ liệu → Enter
6. File PDF + Excel được tạo ở folder output/
```

## ✅ Tính năng

- ✅ Tra cứu sản phẩm (offline)
- ✅ Tạo báo giá + thêm sản phẩm
- ✅ Tính toán tự động (Thành tiền, Tổng cộng)
- ✅ Định dạng tiền tệ chuẩn (dấu phẩy + ₫)
- ✅ Lưu báo giá vào localStorage
- ✅ Xuất PDF + Excel theo mẫu công ty
- ✅ Role-based access (Chủ/Nhân viên)

## 📞 Liên hệ

**Công ty:** HƯNG THỊNH SMART  
**Email:** maymaybienxanh68@gmail.com  
**ĐT:** 0982829806
