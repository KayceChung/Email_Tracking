# 📖 Hướng Dẫn Cài Đặt Gmail Email Logger

## 1. YÊU CẦU TIÊN QUYẾT
- Python 3.8+
- Tài khoản Google (với Gmail và Google Sheets)
- Google Cloud Project với Gmail API & Sheets API bật

## 2. BƯỚC CÀI ĐẶT

### 2.1 Tạo Google Cloud Project
1. Vào https://console.cloud.google.com/
2. Tạo project mới
3. Bật 2 API:
   - Gmail API
   - Google Sheets API
4. Tạo OAuth 2.0 Client ID (Desktop app)
5. Download file JSON → lưu thành `credentials.json` cùng folder script

### 2.2 Cài Đặt Python Libraries
```bash
pip install google-auth-oauthlib google-auth-httplib2 google-api-python-client gspread
```

### 2.3 Cấu Hình Script
Mở `gmail_email_logger.py` và sửa:
```python
SPREADSHEET_ID = "10DpSr-N4jOU5lhNp-8m0F_0KH5iIKBuTlsAZ3wOtNq0"  # Thay ID spreadsheet của bạn
```

## 3. CHẠY SCRIPT

### Lần đầu tiên
```bash
python gmail_email_logger.py
```
Sẽ mở browser để xác thực Gmail → Chọn tài khoản → Cho phép quyền

### Lần sau
```bash
python gmail_email_logger.py
```
Script sẽ sử dụng token lưu trong `token.pickle`

## 4. TỰ ĐỘNG HÓA (Optional)

### Trên Windows (Task Scheduler)
```batch
# Chạy mỗi ngày lúc 8:00 AM
py C:\path\to\gmail_email_logger.py
```

### Trên Mac/Linux (Cron)
```bash
# Mở crontab
crontab -e

# Thêm dòng (chạy mỗi ngày lúc 8:00)
0 8 * * * cd /path/to/script && python gmail_email_logger.py >> log.txt 2>&1
```

## 5. CẤU TRÚC LABELS (Cần Tạo Trong Gmail)

### Booking Emails
- `KLOOK_BOOKING`
- `12GO_BOOKING`
- `BAOLAU_BOOKING`

### Confirmed Emails
- `KLOOK_CONFIRMED`
- `12GO_CONFIRMED`
- `BAOLAU_CONFIRMED`
- `Klook Độc Quyền - Confirmed`

### Cancel Emails
- `KLOOK_CONFIRMED_CANCEL`
- `KLOOK_UNCONFIRMED_CANCEL`
- `BAOLAU_CANCEL`
- `12GO_CANCEL`
- `Klook Độc Quyền - Cancel`

### Khác
- `CTRIP`, `CTRIP CANCEL`
- `Klook Độc Quyền`
- `KLOOK REVIEW`, `12GO_REVIEW`
- `SeatOS-Booking`, `SeatOS-Pending`, `SeatOS - Unchecked`

## 6. TROUBLESHOOTING

### ❌ "credentials.json not found"
→ Tải credentials.json từ Google Cloud Console

### ❌ "Sheet not found"
→ Kiểm tra tên sheet chính xác trong `SHEETS_CONFIG`

### ❌ "Gmail API not enabled"
→ Vào Google Cloud Console → Enable Gmail API

### ❌ "Token expired"
→ Xóa `token.pickle` và chạy lại script

## 7. ỲU ĐIỂM SO VỚI GOOGLE APPS SCRIPT

| Tính Năng | Google Apps Script | Python Script |
|-----------|-------------------|---------------|
| Quota hàng ngày | 20,000 calls | Không giới hạn* |
| Tốc độ | Chậm | Nhanh hơn 10x |
| Sửa lỗi | Khó | Dễ hơn |
| Tự động hóa | Trigger | Cron/Task Scheduler |
| Chi phí | Miễn phí (có giới hạn) | Miễn phí |

*Gmail API: 1 tỉ requests/ngày/project

## 8. TÙYCHỈNH THÊM

### Thay Đổi Thời Gian Quét Email
Tìm trong code:
```python
messages = get_gmail_messages(gmail_service, config['labels'], hours_back=48)
```
Đổi `48` thành số giờ muốn quét (ví dụ: `24` = 1 ngày)

### Thêm Trích Xuất Dữ Liệu Mới
Sửa hàm `extract_order_id()` hoặc `extract_route_points()`

### Giới Hạn Số Email
Tìm:
```python
maxResults=20
```
Đổi thành số khác nếu cần

## 📞 CẦN GIÚP ĐỠ?

Kiểm tra:
1. File `token.pickle` có tồn tại không?
2. Credentials có quyền đúng không?
3. Tên label trong Gmail chính xác không?
4. Google Sheets API đã được bật?
