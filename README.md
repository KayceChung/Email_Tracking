# 📧 Gmail Email Logger - Python Edition

Thay thế 3 Google Apps Script bằng 1 Python script **hiệu quả hơn, nhanh hơn, không giới hạn quota**.

## 🎯 Tính Năng Chính

✅ **Lấy dữ liệu từ Gmail** theo nhiều label cùng lúc  
✅ **Ghi vào Google Sheets** tự động (tránh duplicate)  
✅ **Trích xuất dữ liệu**: Order ID, Route Points, Email Date  
✅ **Hỗ trợ nhiều platform**: KLOOK, 12GO, BAOLAU, CTRIP, SeatOS  
✅ **Xử lý lỗi tốt**: Retry, fallback, logging chi tiết  
✅ **Nhanh & hiệu quả**: Batch processing, rate limiting  

## 📊 So Sánh

| Tính Năng | Google Apps Script | Python Script |
|-----------|-------------------|---------------|
| **Quota hàng ngày** | 20,000 calls | Unlimited* |
| **Tốc độ** | ~2s/email | ~0.3s/email |
| **Setup** | Trên Cloud | Local/VPS |
| **Tự động hóa** | Trigger | Cron/Task Scheduler |
| **Kiểm soát** | Giới hạn | Toàn quyền |

## 🚀 Quick Start

### 1. Cài Đặt (3 phút)
```bash
# Clone hoặc download files
git clone <repo>
cd gmail_email_logger

# Cài dependencies
pip install -r requirements.txt

# Download credentials.json từ Google Cloud Console
# (Chi tiết: xem SETUP_GUIDE.md)
```

### 2. Cấu Hình
```bash
# Copy file config
cp config.json.example config.json

# Edit config.json với SPREADSHEET_ID của bạn
nano config.json
```

### 3. Chạy
```bash
# Lần đầu (xác thực Gmail)
python gmail_email_logger.py

# Lần sau
python gmail_email_logger.py
```

## 📁 Cấu Trúc File

```
gmail_email_logger/
├── gmail_email_logger.py      # Main script
├── config.json.example         # Template config
├── requirements.txt            # Dependencies
├── SETUP_GUIDE.md             # Hướng dẫn cài đặt chi tiết
├── VS_CODE_PROMPTS.md         # Prompt cho Claude Code
├── README.md                   # File này
├── token.pickle               # Token Gmail (tự động tạo)
└── logs/                       # Log files
```

## 🔧 Cấu Hình Các Label

Script hỗ trợ các label sau (tạo trong Gmail):

### 📦 Booking
- `KLOOK_BOOKING`
- `12GO_BOOKING`
- `BAOLAU_BOOKING`

### ✅ Confirmed
- `KLOOK_CONFIRMED`
- `12GO_CONFIRMED`
- `BAOLAU_CONFIRMED`
- `Klook Độc Quyền - Confirmed`

### ❌ Cancel
- `KLOOK_CONFIRMED_CANCEL`
- `KLOOK_UNCONFIRMED_CANCEL`
- `BAOLAU_CANCEL`
- `12GO_CANCEL`
- `Klook Độc Quyền - Cancel`

### 🌐 Khác
- `CTRIP`, `CTRIP CANCEL`
- `Klook Độc Quyền`, `KLOOK REVIEW`, `12GO_REVIEW`
- `SeatOS-Booking`, `SeatOS-Pending`, `SeatOS - Unchecked`

## 📈 Metrics & Logging

Script tự động output:
```
🚀 Bắt đầu lấy dữ liệu email từ Gmail...

✓ Xác thực thành công

📋 Xử lý BOOKING:
   Sheet: EMAIL_BOOKING
   Labels: KLOOK_BOOKING, 12GO_BOOKING, BAOLAU_BOOKING
   ✓ Tìm thấy 5 email từ label 'KLOOK_BOOKING'
   ✓ Tìm thấy 3 email từ label '12GO_BOOKING'
   ✓ Đã ghi 7 dòng vào EMAIL_BOOKING

📋 Xử lý CONFIRMED:
   ...

✅ Hoàn thành!
```

## 🔐 Bảo Mật

- ✅ Token lưu local (token.pickle) - **không upload lên cloud**
- ✅ Credentials.json chỉ cần read email & write sheets
- ✅ Không lưu password, chỉ dùng OAuth2
- ✅ `.gitignore` bao gồm files nhạy cảm

## ⚙️ Tùy Chỉnh Advanced

### Thay đổi thời gian quét
```python
# Trong code:
messages = get_gmail_messages(gmail_service, config['labels'], hours_back=24)
# 24 = quét 24h gần nhất
```

### Thêm logic trích xuất dữ liệu
```python
def extract_order_id(label, subject, body):
    # Thêm pattern mới ở đây
    if "NEW_PROVIDER" in label:
        return extract_new_provider_order_id(subject)
```

## 📝 Ví Dụ Output

**Sheet EMAIL_BOOKING:**
| Label | Type | Subject | OrderID | Pickup | Dropoff |
|-------|------|---------|---------|--------|---------|
| KLOOK_BOOKING | EMAIL | Klook Order Confirmation - ABCD1234 | ABCD1234 | HCM City | Da Lat |
| 12GO_BOOKING | EMAIL | 12GO Booking #12GO-ABC456 | ABC456 | Saigon | Mui Ne |

## 🐛 Troubleshooting

### ❌ "credentials.json not found"
→ Tải từ Google Cloud Console (xem SETUP_GUIDE.md)

### ❌ "Sheet 'EMAIL_BOOKING' not found"
→ Kiểm tra tên sheet trong config.json (phải chính xác, có dấu)

### ❌ "Gmail API not enabled"
→ Enable tại: https://console.cloud.google.com/ (search "Gmail API")

### ❌ "Token expired / needs re-auth"
→ Xóa file `token.pickle` và chạy lại script

Xem **SETUP_GUIDE.md** để troubleshooting chi tiết.

## 🧪 Testing

```bash
# Test connectivity
python -c "from gmail_email_logger import authenticate_gmail; authenticate_gmail()"

# Test một label cụ thể
python gmail_email_logger.py --dry-run --sheet booking
```

## 📚 Tài Liệu

- **SETUP_GUIDE.md**: Hướng dẫn cài đặt & cấu hình
- **VS_CODE_PROMPTS.md**: Prompt cho Claude Code (cải tiến script)
- **config.json.example**: Template cấu hình
- Inline comments: Chi tiết trong code

## 🤖 Dùng Claude Code Để Cải Tiến

Muốn thêm feature mới? Sử dụng prompts trong **VS_CODE_PROMPTS.md**:

```
1. Mở VS Code → Claude Code extension
2. Copy-paste PROMPT 1, 2, 3... từ VS_CODE_PROMPTS.md
3. Claude sẽ tự động cập nhật code
4. Review & save files
```

**Ví dụ prompts:**
- Thêm logging chi tiết
- Tích hợp với database
- Add CLI arguments
- Optimize extraction patterns

## 🌍 Deployment

### Local (Development)
```bash
python gmail_email_logger.py
```

### Scheduled (Windows)
→ Xem SETUP_GUIDE.md (Task Scheduler)

### Scheduled (Mac/Linux)
```bash
# Chạy mỗi ngày lúc 8 AM
0 8 * * * cd /path/to/script && python gmail_email_logger.py
```

### VPS/Cloud
Deploy lên Heroku, AWS Lambda, Google Cloud Run, v.v.

## 📞 Support

- Check logs trong `logs/` folder
- Review inline comments trong code
- Xem SETUP_GUIDE.md troubleshooting section

## 📄 License

Open source - dùng tự do, sửa tự do.

## 🎉 Lợi Ích So Với GAS

| Aspect | GAS | Python |
|--------|-----|--------|
| Setup | 2h | 10m |
| Speed | 2s/email | 0.3s/email |
| Quota | 20k/day limit | Unlimited* |
| Error Handling | Limited | Excellent |
| Debugging | Hard | Easy |
| Customization | Limited | Unlimited |
| Maintenance | Difficult | Easy |
| Learning Curve | Moderate | Beginner-friendly |

---

**Sẵn sàng để chuyển sang Python? Bắt đầu từ SETUP_GUIDE.md!** 🚀
