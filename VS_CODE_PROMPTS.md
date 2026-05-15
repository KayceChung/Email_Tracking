# 🎯 PROMPT DÀNH CHO VS CODE (Claude Code Extension)

## HƯỚNG DẪN: Copy-paste prompt này vào Claude Code trong VS Code

---

## PROMPT 1: TỐI ƯU HÓA TOÀN BỘ SCRIPT

```
Tôi có 3 Google Apps Script (GAS) dùng để lấy dữ liệu email từ Gmail và ghi vào Google Sheets.
Tôi muốn chuyển sang Python để tránh quota limit của Google Apps Script.

Dữ liệu email hiện tại từ các label sau:
- BOOKING: KLOOK_BOOKING, 12GO_BOOKING, BAOLAU_BOOKING
- CONFIRMED: KLOOK_CONFIRMED, 12GO_CONFIRMED, BAOLAU_CONFIRMED  
- CANCEL: KLOOK_CONFIRMED_CANCEL, KLOOK_UNCONFIRMED_CANCEL, BAOLAU_CANCEL, 12GO_CANCEL
- KHÁC: CTRIP, CTRIP CANCEL, Klook Độc Quyền, KLOOK REVIEW, 12GO_REVIEW
- SEATOS: SeatOS-Booking, SeatOS-Pending, SeatOS - Unchecked

YÊU CẦU:
1. Tạo Python script đơn nhất thay thế cả 3 GAS script hiện tại
2. Hỗ trợ tất cả label email trên
3. Trích xuất Order ID và Route Points (điểm đón/trả) cho mỗi loại
4. Ghi vào Google Sheets tương ứng (EMAIL_BOOKING, EMAIL_XÁC_NHẬN, EMAIL_HỦY, etc.)
5. Tránh lấy trùng email (check threadId đã tồn tại)
6. Xử lý error gracefully
7. Thêm logging rõ ràng (mấy email tìm được, mấy email ghi được)

Tối ưu hóa:
- Batch processing (lấy nhiều email, ghi 1 lần)
- Rate limiting hợp lý (0.1s giữa các request)
- Reuse Gmail connection
- Check email reply để skip
- Decode email body correctly (base64)

Output: File gmail_email_logger.py có thể chạy: python gmail_email_logger.py
```

---

## PROMPT 2: THÊM CHỨC NĂNG ADVANCED

```
Dựa trên Python script email logger hiện tại, hãy thêm các tính năng sau:

1. **Xử Lý Lỗi & Logging Tốt Hơn**
   - Tạo file log chi tiết (log_YYYY-MM-DD.txt)
   - Ghi lại: email nào được xử lý, lỗi gì xảy ra, timestamp
   - Color-coded console output (✓ success, ❌ error, ⚠️ warning, ℹ️ info)

2. **Cấu Hình Config**
   - Tách config ra file riêng (config.json)
   - Cho phép tùy chỉnh: SPREADSHEET_ID, labels, hours_back, maxResults
   - Ví dụ:
     {
       "spreadsheet_id": "...",
       "sheets": {
         "booking": {...},
         "confirmed": {...}
       },
       "gmail": {
         "hours_back": 48,
         "max_results": 20,
         "batch_size": 10
       }
     }

3. **Database Caching**
   - Lưu list thread_id đã xử lý vào SQLite (thay vì check sheet mỗi lần)
   - Nhanh hơn, ít call Google Sheets API

4. **Email Body Extraction Tốt Hơn**
   - Extract text từ cả HTML (strip HTML tags)
   - Limit text to 500 chars (hiện tại)
   - Option để lưu full body nếu cần

5. **Resume & Dry-Run Mode**
   - Dry-run: test mà không ghi vào sheet
   - Resume: lấy tiếp từ email cuối cùng đã xử lý

6. **Dashboard Simple**
   - In ra summary: tổng email tìm được, ghi được, skip, lỗi
   - Timing: mấy giây xử lý xong

Output: Code với struktur tốt, comments rõ ràng, có thể config dễ dàng
```

---

## PROMPT 3: THÊM CLI INTERFACE & SCHEDULING

```
Mở rộng Python script email logger với:

1. **CLI Arguments (Command Line)**
   ```
   python gmail_email_logger.py --help
   
   Usage:
     python gmail_email_logger.py                    # Chạy bình thường
     python gmail_email_logger.py --dry-run         # Test, không ghi
     python gmail_email_logger.py --sheet booking   # Chỉ process sheet "booking"
     python gmail_email_logger.py --hours 24        # Lấy email trong 24h gần nhất
     python gmail_email_logger.py --verbose         # Log chi tiết
     python gmail_email_logger.py --re-auth         # Xác thực lại Gmail
   ```

2. **Scheduled Task Helper**
   - Tạo file setup_schedule.py (1 lần chạy duy nhất)
   - Windows: Tạo Task Scheduler job tự động
   - Mac/Linux: Tạo cron job tự động
   - User chỉ cần chạy: python setup_schedule.py

3. **Health Check**
   - Test Google Cloud credentials có valid không
   - Test access từng label
   - Test write permission vào Sheets
   - In ra "All systems OK" hoặc lỗi cụ thể

Output: gmail_email_logger.py improved, setup_schedule.py, requirements.txt updated
```

---

## PROMPT 4: OPTIMIZE EXTRACTION LOGIC

```
Cải thiện phần trích xuất dữ liệu (Order ID & Route Points):

1. **Order ID Extraction - Tổng Hợp Pattern Tốt Hơn**
   
   Pattern hiện tại chỉ match subject. Cần extend:
   
   KLOOK:
   - Subject pattern: "- ABC123XYZ" (cuối)
   - Body pattern: "Order ID: ABC123XYZ"
   - Body pattern: "Booking reference ID: ABC123XYZ"
   
   12GO:
   - Pattern: "12GO-XXXXXX" hoặc "12GO XXXXXX"
   - Multiple formats để match
   
   BAOLAU:
   - Pattern: "(12ABC)" hoặc "[12ABC]"
   - Fallback: extract từ body
   
   CTRIP:
   - Pattern: "CT_XXXXXX" hoặc "CTRIP_XXXXXX"
   
   SEATOS:
   - Pattern: "ST_XXXXXX" hoặc "SeatOS_XXXXXX"
   
   Hãy viết một function tổng hợp tất cả patterns, test với examples, in ra match results.

2. **Route Points Extraction - Phức Tạp Hơn**
   
   Hiện tại chỉ support 12GO & KLOOK. Cần thêm:
   
   BAOLAU, CTRIP, SEATOS:
   - Parse "From: X to: Y" pattern
   - Parse "pickup: X, dropoff: Y" pattern
   - Extract từ itinerary trong body
   
   Fallback strategy:
   - Nếu không tìm được, return empty string thay "N/A"
   - Option để skip columns này nếu không cần
   
   Output: function tổng hợp, test cases, documentation

3. **Email Subject & Body Parsing**
   - Normalize subject (trim, lowercase comparison)
   - Handle special characters trong tiếng Việt
   - Extract departure time (nếu có)
   - Extract passenger count (nếu có)

Output: Improved extraction functions với test cases & examples
```

---

## PROMPT 5: INTEGRATION VỚI EXTERNAL SERVICES

```
Mở rộng script để tích hợp với:

1. **Database Logging (Tùy Chọn)**
   - Insert/update records vào SQLite/PostgreSQL
   - Keep backup data
   - Query/report easier
   - Schema: (id, label, subject, order_id, pickup, dropoff, email_date, sheet_date, thread_link)

2. **Notification System**
   - Send email/Slack khi có order mới
   - Alert nếu có error
   - Daily report: X email processed

3. **Data Validation**
   - Check Order ID format validity
   - Validate route points (không empty, không invalid)
   - Flag suspicious emails

Output: Optional modules, documentation, examples
```

---

## PROMPT 6: TESTING & DOCUMENTATION

```
Hoàn thiện script với:

1. **Unit Tests**
   - Test extract_order_id() với 10+ cases mỗi label
   - Test extract_route_points() tương tự
   - Test decode_email_body() với HTML/plain text
   - Mock Gmail API responses
   - Run: pytest tests/

2. **Documentation**
   - README.md: overview, features, quick start
   - ARCHITECTURE.md: structure, flow diagram
   - API.md: functions & parameters
   - EXAMPLES.md: real-world examples
   - TROUBLESHOOTING.md: common issues

3. **Code Quality**
   - Type hints cho tất cả functions
   - Docstrings chi tiết
   - Error messages descriptive
   - Consistent formatting (Black)

Output: tests/, docs/, updated code dengan type hints & docstrings
```

---

## CÁCH SỬ DỤNG PROMPT VỚI CLAUDE CODE

1. **Mở VS Code** → Click "Claude Code" extension
2. **Paste prompt** nào bạn muốn vào chat
3. **Claude sẽ**:
   - Đề xuất changes
   - Tạo/edit files
   - Run tests nếu cần
4. **Review & confirm** trước mỗi change
5. **Lưu file** → Ready to use!

---

## 💡 TIPS DÙNG PROMPT HIỆU QUẢ

✅ **LÀM NÀY:**
```
"Hãy thêm logging chi tiết vào hàm decode_email_body(). 
Khi error xảy ra, cần log: email ID, error message, fallback value sử dụng"
```

❌ **KHÔNG NÀY:**
```
"Làm tốt hơn"
```

---

## 🎯 RECOMMENDED ORDER

1. **Prompt 1**: Tạo script cơ bản
2. **Prompt 2**: Add logging & config
3. **Prompt 3**: Add CLI & scheduling
4. **Prompt 4**: Optimize extraction
5. **Prompt 6**: Testing & docs
6. **Prompt 5**: Advanced integration (nếu cần)

---

## ❓ MẶC ĐỊNH QUESTIONS CÓ THỂ HỎI CLAUDE CODE

- "Tại sao regex này không match?"
- "Có cách nào để xử lý timeout không?"
- "Làm sao để debug email decode error?"
- "Performance có thể improve thêm không?"
- "Script này có security issues gì không?"

---

**Happy Coding! 🚀**
