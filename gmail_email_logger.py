#!/usr/bin/env python3
"""
Gmail Email Logger - Tối ưu hóa lấy dữ liệu tu Gmail và ghi vào Google Sheets
Thay thế 3 script Google Apps Script hiện tại
"""

import os
import pickle
import json
import re
import base64
import sys
from datetime import datetime, timedelta, timezone
from typing import Dict, List, Tuple, Set, Optional, Any, Callable
from urllib.parse import parse_qs, urlparse
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.oauth2.service_account import Credentials as ServiceAccountCredentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread
from gspread.auth import authorize
from gspread.exceptions import WorksheetNotFound
import time

# ==================== CẤU HÌNH ====================
IS_CI = os.getenv('CI', '').lower() in ('true', '1', 'yes')


def env_int(name: str, default: int) -> int:
    """Đọc biến môi trường int với fallback an toàn."""
    try:
        return int(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


def env_float(name: str, default: float) -> float:
    """Đọc biến môi trường float với fallback an toàn."""
    try:
        return float(os.getenv(name, str(default)))
    except (TypeError, ValueError):
        return default


def env_bool(name: str, default: bool) -> bool:
    """Đọc biến môi trường bool với fallback an toàn."""
    value = os.getenv(name)
    if value is None:
        return default
    return value.lower() in ('true', '1', 'yes', 'on')


def console_text(value: Any) -> str:
    """Chuyển text sang dạng an toàn để in ra terminal Windows."""
    return str(value).encode('unicode_escape').decode('ascii')


SCOPES = [
    'https://www.googleapis.com/auth/gmail.readonly',
    'https://www.googleapis.com/auth/spreadsheets'
]

# OAuth mode: "local" (default) để chỉ cần đăng nhập browser.
# Có thể đổi thành "manual" bằng biến môi trường GMAIL_OAUTH_MODE=manual.
GMAIL_OAUTH_MODE = os.getenv('GMAIL_OAUTH_MODE', 'local').lower()
GMAIL_HOURS_BACK = env_int('GMAIL_HOURS_BACK', 24)
GMAIL_DAYS_BACK = env_int('GMAIL_DAYS_BACK', 0)
GMAIL_MAX_RESULTS_PER_LABEL = env_int('GMAIL_MAX_RESULTS_PER_LABEL', 0)
# Mặc định quét theo cửa sổ 24h gần nhất. Có thể đặt GMAIL_FETCH_ALL=true để quét toàn bộ.
GMAIL_FETCH_ALL = env_bool('GMAIL_FETCH_ALL', False)
# Ngày bắt đầu lấy dữ liệu. Hỗ trợ: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY.
# Khi đặt, sẽ override GMAIL_FETCH_ALL.
GMAIL_AFTER_DATE = os.getenv('GMAIL_AFTER_DATE', '').strip()
# Múi giờ dùng để ghi DATE/DATE_TYPE và tính mốc GMAIL_AFTER_DATE (mặc định UTC+7).
GMAIL_TIMEZONE_OFFSET_HOURS = env_int('GMAIL_TIMEZONE_OFFSET_HOURS', 7)
LOGGER_IGNORE_STATE = env_bool('LOGGER_IGNORE_STATE', True)
# Delay giữa các BATCH (mỗi batch = 100 email), không phải per-message.
GMAIL_REQUEST_DELAY_SECONDS = env_float('GMAIL_REQUEST_DELAY_SECONDS', 0.0 if IS_CI else 0.3)
GSPREAD_REQUEST_TIMEOUT_SECONDS = env_float('GSPREAD_REQUEST_TIMEOUT_SECONDS', 30.0)
GSPREAD_OPEN_RETRIES = env_int('GSPREAD_OPEN_RETRIES', 5)
API_CALL_RETRIES = env_int('API_CALL_RETRIES', 5)
# Mặc định chạy daemon mỗi 30 phút.
LOGGER_RUN_MODE = os.getenv('LOGGER_RUN_MODE', 'daemon').lower()  # once | daemon
LOGGER_POLL_SECONDS = int(os.getenv('LOGGER_POLL_SECONDS', '1800'))
LOGGER_START_FROM_SHEET = os.getenv('LOGGER_START_FROM_SHEET', '').strip().lower()
STATE_FILE = os.getenv('LOGGER_STATE_FILE', 'sync_state.json')

SPREADSHEET_ID = os.getenv('SPREADSHEET_ID') or '10DpSr-N4jOU5lhNp-8m0F_0KH5iIKBuTlsAZ3wOtNq0'
SHEETS_CONFIG = {
    "booking": {
        "name": "EMAIL_BOOKING",
        "labels": ["KLOOK_BOOKING", "BAOLAU_BOOKING","SeatOS-Booking"]
    },
    "confirmed": {
        "name": "EMAIL_XÁC_NHẬN",
        "labels": ["KLOOK_CONFIRMED", "BAOLAU_CONFIRMED", "Klook Độc Quyền - Confirmed","SeatOS-Confirm"]
    },
    "cancel": {
        "name": "EMAIL_HỦY",
        "labels": ["KLOOK_CONFIRMED_CANCEL", "KLOOK_UNCONFIRMED_CANCEL", "BAOLAU_CANCEL", "Klook Độc Quyền - Cancel","SeatOS-Cancel"]
    },
    "ctrip": {
        "name": "EMAIL_CTRIP",
        "labels": ["CTRIP", "CTRIP CANCEL"]
    },
    "review": {
        "name": "EMAIL_REVIEW",
        "labels": ["KLOOK REVIEW", "12GO_REVIEW"]
    },
    "seatos": {
        "name": "EMAIL_SEATOS",
        "labels": [ "SeatOS-Pending", "SeatOS - Unchecked"]
    }
}

BOOKING_HEADERS = [
    "NEN_TANG",
    "PHAN_LOAI",
    "SUBJECT",
    "BODY",
    "LINK",
    "DATE",
    "WEEKNUM",
    "DATE_TYPE",
    "ID_DH",
    "EMAIL_PHAN_HOI",
    "P.I",
    "P.I_TIMES",
    "LATEST_UPDATE_BY",
    "LATEST_UPDATES",
    "Change_counter",
    "STATUS_TIME_CHANGE",
    "DURATION",
    "CHUYEN",
    "TRANG_THAI_DH",
    "Message_id",
    "DAM_NHAN_HO_TRO",
    "DIEM_DI",
    "DIEM_DON",
    "NOI_DUNG_CONFIRM",
    "THOI_GIAN_DAM_NHAN_H_TRO",
    "CONFIRM_TU_NHAN_VIEN"
]

STANDARD_HEADERS = [
    "NEN_TANG",
    "PHAN_LOAI",
    "SUBJECT",
    "BODY",
    "LINK",
    "DATE",
    "WEEKNUM",
    "DATE_TYPE",
    "ID_DH",
    "P.I",
    "P.I_TIME",
    "LATEST_UPDATE_BY",
    "LATEST_UPDATES",
    "CONFIRM_TIME"
]

SHEET_HEADERS = {
    "booking": BOOKING_HEADERS,
    "confirmed": STANDARD_HEADERS,
    "cancel": STANDARD_HEADERS,
    "ctrip": STANDARD_HEADERS,
    "review": STANDARD_HEADERS,
    "seatos": STANDARD_HEADERS,
}

# ==================== GMAIL AUTHENTICATION ====================
def authenticate_gmail():
    """Xac thuc Gmail API"""
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    # Detect non-interactive environment (GitHub Actions sets CI=true)
    is_ci = os.getenv('CI', '').lower() in ('true', '1', 'yes')

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
                print("[OK] Token da duoc refresh thanh cong")
            except Exception as refresh_error:
                if is_ci:
                    raise RuntimeError(
                        f"[ERR] Token hết hạn và không thể refresh: {refresh_error}\n"
                        "Hãy chạy lại script trên máy local để tạo token moi, "
                        "rồi cập nhật secret GMAIL_TOKEN_PICKLE_B64 trên GitHub."
                    )
                raise
        else:
            if is_ci:
                raise RuntimeError(
                    "[ERR] Khong co token hợp lệ trong môi trường CI (không có browser).\n"
                    "Chạy script trên máy local trước để tạo token.pickle, "
                    "rồi cập nhật secret GMAIL_TOKEN_PICKLE_B64 trên GitHub."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            if GMAIL_OAUTH_MODE == 'manual':
                creds = authenticate_gmail_manual(flow)
            else:
                try:
                    # Local callback OAuth flow.
                    creds = flow.run_local_server(port=0)
                except Exception as local_oauth_error:
                    print(f"[WARN] OAuth local server loi: {repr(local_oauth_error)}")
                    print("[=>] Chuyển sang chế độ nhập URL callback thủ công...")
                    creds = authenticate_gmail_manual(flow)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return build('gmail', 'v1', credentials=creds)


def authenticate_gmail_manual(flow: InstalledAppFlow) -> Any:
    """Fallback OAuth không cần local callback server."""
    auth_url, _ = flow.authorization_url(
        access_type='offline',
        include_granted_scopes='true',
        prompt='consent'
    )

    print("\nMở URL sau trong browser và cho phép quyền:")
    print(auth_url)
    print("\nSau khi Google redirect về trang localhost loi, copy TOÀN BỘ URL trên thanh địa chỉ và dán vào đây.")

    try:
        redirected_url = input("Redirect URL: ").strip()
    except EOFError:
        raise RuntimeError(
            "Khong đọc duoc input tu terminal (EOF). Hãy chạy script trong terminal tương tác và dán Redirect URL."
        )
    if not redirected_url:
        raise ValueError("Bạn chưa nhập Redirect URL")

    # Xac thuc URL có chứa authorization code.
    parsed = urlparse(redirected_url)
    query_params = parse_qs(parsed.query)
    if 'code' not in query_params:
        raise ValueError("Redirect URL không chứa 'code'. Hãy copy lại URL đầy đủ sau khi cấp quyền.")

    auth_code = query_params['code'][0]
    flow.fetch_token(code=auth_code)
    return flow.credentials

def authenticate_gspread():
    """Xac thuc Google Sheets"""
    creds: Optional[Any] = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)

    if creds is None:
        raise RuntimeError("Khong tìm thấy token.pickle hợp lệ để xác thực Google Sheets")
    client = authorize(creds)
    # Dat timeout cho request Google Sheets de tranh treo vo han khi mang cham.
    if hasattr(client, 'set_timeout'):
        client.set_timeout(GSPREAD_REQUEST_TIMEOUT_SECONDS)
    return client


def open_spreadsheet_with_retry(gs_client, spreadsheet_id: str):
    """Mở Google Sheet với retry + backoff để tránh fail/treo tạm thời."""
    last_error: Optional[Exception] = None
    for attempt in range(1, GSPREAD_OPEN_RETRIES + 1):
        try:
            return gs_client.open_by_key(spreadsheet_id)
        except KeyboardInterrupt:
            raise
        except Exception as e:
            last_error = e
            wait = min(30, 2 ** (attempt - 1))
            print(
                f"[WARN] Mở Google Sheet thất bại (lần {attempt}/{GSPREAD_OPEN_RETRIES}): {e}"
            )
            if attempt < GSPREAD_OPEN_RETRIES:
                print(f"[RETRY] Chờ {wait}s rồi thử lại...")
                time.sleep(wait)

    raise RuntimeError(f"Khong thể mở Google Sheet sau {GSPREAD_OPEN_RETRIES} lần: {last_error}")


def is_transient_network_error(error: Exception) -> bool:
    """Nhận diện lỗi mạng/DNS tạm thời để retry."""
    text = str(error).lower()
    keywords = (
        'name resolution',
        'failed to resolve',
        'getaddrinfo failed',
        'temporary failure',
        'connection reset',
        'connection aborted',
        'connectionerror',
        'read timed out',
        'timed out',
        'timeout',
        'ssl',
        'max retries exceeded',
        'httpsconnectionpool',
        'oauth2.googleapis.com',
        'sheets.googleapis.com'
    )
    return any(k in text for k in keywords)


def call_with_retry(func: Callable[[], Any], op_name: str, retries: int = API_CALL_RETRIES) -> Any:
    """Retry có backoff cho các API call hay lỗi mạng tạm thời."""
    last_error: Optional[Exception] = None
    for attempt in range(1, max(1, retries) + 1):
        try:
            return func()
        except KeyboardInterrupt:
            raise
        except Exception as e:
            last_error = e
            if not is_transient_network_error(e) or attempt >= max(1, retries):
                raise
            wait = min(30, 2 ** attempt)
            print(
                f"[WARN] {op_name} lỗi mạng tạm thời (lan {attempt}/{retries}): {e}"
            )
            print(f"[RETRY] Chờ {wait}s rồi thử lại {op_name}...")
            time.sleep(wait)

    raise RuntimeError(f"{op_name} thất bại sau retry: {last_error}")

# ==================== LẤY DỮ LIỆU EMAIL ====================
def decode_email_body(payload: Dict) -> str:
    """Giải mã nội dung email tu payload"""
    try:
        def decode_part(part):
            if part.get('body', {}).get('data'):
                data = part['body']['data']
                decoded = base64.urlsafe_b64decode(data + '==')
                return decoded.decode('utf-8', errors='ignore')
            return ''
        
        # Tìm plain text trước
        parts = payload.get('parts', [])
        for part in parts:
            if part.get('mimeType') == 'text/plain':
                return decode_part(part)
        
        # Nếu không có plain text, tìm HTML
        for part in parts:
            if part.get('mimeType') == 'text/html':
                return decode_part(part)
        
        return payload.get('snippet', '')
    except Exception as e:
        print(f"Error decoding email: {e}")
        return payload.get('snippet', '')

_BATCH_SIZE = 20   # Giới hạn concurrent requests để tránh 429 "Too many concurrent requests"

def get_gmail_messages(
    service,
    labels: List[str],
    hours_back: int = 48,
    max_results_per_label: int = 200,
    after_epoch_seconds: Optional[int] = None,
    label_name_to_id: Optional[Dict[str, str]] = None,
    on_message_batch: Optional[Callable[[List[Dict], str], None]] = None
) -> List[Dict]:
    """Lấy danh sách email tu Gmail với các label cụ thể (dùng Batch API)"""
    all_messages = []

    if label_name_to_id is None:
        labels_result = service.users().labels().list(userId='me').execute()
        label_name_to_id = {
            lbl['name']: lbl['id']
            for lbl in labels_result.get('labels', [])
            if 'name' in lbl and 'id' in lbl
        }

    for label in labels:
        try:
            label_id = label_name_to_id.get(label)
            if not label_id:
                print(f"[WARN] Label '{console_text(label)}' khong tim thay")
                continue

            # Xây query thời gian.
            if after_epoch_seconds:
                query = f"after:{after_epoch_seconds} -in:draft"
            elif hours_back == 0:
                query = "-in:draft"
            else:
                query = f"newer_than:{hours_back}h -in:draft"

            # Lấy danh sách message ID (phân trang).
            unlimited = (max_results_per_label == 0)
            messages = []
            next_page_token = None
            while True:
                if unlimited:
                    page_size = 100
                else:
                    page_size = min(100, max(1, max_results_per_label - len(messages)))
                    if page_size <= 0:
                        break

                results = service.users().messages().list(
                    userId='me',
                    labelIds=[label_id],
                    q=query,
                    maxResults=page_size,
                    pageToken=next_page_token
                ).execute()

                messages.extend(results.get('messages', []))
                next_page_token = results.get('nextPageToken')
                if not next_page_token:
                    break
                if not unlimited and len(messages) >= max_results_per_label:
                    break

            print(f"[OK] Tim thay {len(messages)} email tu label '{console_text(label)}'")

            # Luon fetch chi tiet tat ca message id trong cua so thoi gian de khong bo sot
            # email moi trong cung thread. Dedupe duoc xu ly o tang build row truoc khi ghi sheet.
            msgs_to_fetch = messages

            if not msgs_to_fetch:
                continue

            print(f"  [>] Can fetch chi tiet {len(msgs_to_fetch)} email (batch {_BATCH_SIZE}/lan)")

            id_to_msg = {m['id']: m for m in msgs_to_fetch}
            pending = list(msgs_to_fetch)
            processed_count = 0

            # Batch fetch chi tiet email với retry cho tung message loi 429.
            def _fetch_chunk(chunk: List[Dict]) -> Tuple[List[Dict], List[str]]:
                ok_responses: Dict[str, Any] = {}
                fail_ids: List[str] = []

                def _cb(request_id: str, response: Any, exception: Any):
                    if exception is not None:
                        err_str = str(exception)
                        if '429' in err_str or 'rateLimitExceeded' in err_str:
                            fail_ids.append(request_id)
                        else:
                            print(f"  [WARN] Batch loi message {request_id}: {exception}")
                    elif response is not None:
                        ok_responses[request_id] = response

                batch = service.new_batch_http_request(callback=_cb)
                for m in chunk:
                    batch.add(
                        service.users().messages().get(userId='me', id=m['id'], format='full'),
                        request_id=m['id']
                    )
                batch.execute()

                ok_messages: List[Dict] = []
                for m in chunk:
                    detail = ok_responses.get(m['id'])
                    if detail is None:
                        continue
                    ok_messages.append({
                        'label': label,
                        'id': m['id'],
                        'threadId': detail['threadId'],
                        'payload': detail['payload'],
                        'internalDate': int(detail['internalDate'])
                    })

                return ok_messages, fail_ids

            for attempt in range(6):  # tối đa 6 vòng retry
                if not pending:
                    break
                if attempt > 0:
                    wait = 5 * (2 ** (attempt - 1))  # 5s, 10s, 20s, 40s, 80s
                    print(f"  [RETRY] Retry lan {attempt}: {len(pending)} message bi 429, cho {wait}s...")
                    time.sleep(wait)

                failed_ids: List[str] = []
                for i in range(0, len(pending), _BATCH_SIZE):
                    chunk = pending[i:i + _BATCH_SIZE]
                    ok_messages, chunk_failed_ids = _fetch_chunk(chunk)
                    if ok_messages:
                        if on_message_batch:
                            on_message_batch(ok_messages, label)
                        else:
                            all_messages.extend(ok_messages)
                        processed_count += len(ok_messages)

                    failed_ids.extend(chunk_failed_ids)
                    if processed_count and processed_count % 500 == 0:
                        print(f"  [...] Da fetch chi tiet {processed_count}/{len(msgs_to_fetch)} email")

                    # Delay giữa các batch để tránh quota.
                    if i + _BATCH_SIZE < len(pending):
                        time.sleep(GMAIL_REQUEST_DELAY_SECONDS if GMAIL_REQUEST_DELAY_SECONDS > 0 else 0.5)

                pending = [id_to_msg[mid] for mid in failed_ids if mid in id_to_msg]

            if pending:
                print(f"  [WARN] Bo qua {len(pending)} message sau khi retry het lan")
            print(f"  [OK] Fetch xong chi tiet {processed_count}/{len(msgs_to_fetch)} email")

        except Exception as e:
            print(f"[ERR] Loi xu ly label '{console_text(label)}': {e}")

    return all_messages

# ==================== TRÍCH XUẤT DỮ LIỆU ====================
def extract_order_id(label: str, subject: str, body: str = "") -> str:
    """Trích xuất Order ID theo tung loại label"""
    patterns = {
        "KLOOK": r"-\s([A-Z0-9]+)$",
        "12GO": r"12GO\s*([\w-]+)",
        "BAOLAU": r"\((.*?)\)",
        "CTRIP": r"(?:CTRIP|booking)\s*([\w-]+)",
        "SEATIOS": r"(?:SeatOS|booking)\s*([\w-]+)"
    }
    
    search_text = f"{subject} {body}"
    
    for key, pattern in patterns.items():
        if key.lower() in label.lower():
            match = re.search(pattern, subject, re.IGNORECASE)
            if match:
                return match.group(1)
    
    return "N/A"

def extract_route_points(label: str, subject: str, body: str = "") -> Tuple[str, str]:
    """Trích xuất điểm đón/trả theo tung loại label"""
    
    if "12GO" in label.upper():
        # Format: location1 - location2
        parts = re.search(r"-\s*#12GO\d+\s*(.*)$", subject, re.IGNORECASE)
        if parts and parts.group(1):
            info = parts.group(1).strip()
            # Bỏ ngày/gio
            cleaned = re.sub(r'^\d{1,2}\s+\w+\s+\d{1,2}:\d{2}\s+\w+\s+', '', info)
            
            # Tìm "Hostel" làm điểm ngăn cách
            idx = cleaned.find("Hostel")
            if idx != -1:
                return (cleaned[:idx].strip(), cleaned[idx:].strip())
            
            # Nếu không, chia đôi
            words = cleaned.split()
            if len(words) >= 2:
                return (' '.join(words[:-2]), ' '.join(words[-2:]))
        
        return ("N/A", "N/A")
    
    if "KLOOK" in label.upper():
        # KLOOK: tìm "Pick-up location" và "Drop-off location"
        pickup_match = re.search(r"Pick-up location:\s*(.*?)(?:\n|Drop-off)", body, re.IGNORECASE)
        dropoff_match = re.search(r"Drop-off location:\s*(.*?)(?:\n|$)", body, re.IGNORECASE)
        
        pickup = pickup_match.group(1).strip() if pickup_match else "N/A"
        dropoff = dropoff_match.group(1).strip() if dropoff_match else "N/A"
        
        return (pickup, dropoff)
    
    return ("N/A", "N/A")

def extract_headers(payload: Dict, header_names: List[str]) -> Dict[str, str]:
    """Trích xuất các header cụ thể tu email"""
    headers = payload.get('headers', [])
    result = {}
    
    for header in headers:
        if header['name'] in header_names:
            result[header['name']] = header['value']
    
    return result

# ==================== GHI VÀO GOOGLE SHEETS ====================
def get_existing_row_keys(worksheet, sheet_type: str) -> Set[str]:
    """Lấy khóa đã tồn tại để tránh trùng lặp theo từng sheet."""
    try:
        # KLOOK_SPECIAL cần giữ nhiều email trong cùng thread nếu khác ngày.
        if sheet_type == "klook_special":
            links = worksheet.col_values(5)
            dates = worksheet.col_values(6)
            if len(links) > 1:
                existing_keys = set()
                for index in range(1, min(len(links), len(dates))):
                    link = links[index]
                    date_value = dates[index]
                    if link and date_value:
                        existing_keys.add(f"{link}|{date_value[:10]}")
                return existing_keys

        # CTRIP cần giữ nhiều email trong cùng thread (khác timestamp).
        if sheet_type == "ctrip":
            links = worksheet.col_values(5)
            dates = worksheet.col_values(6)
            if len(links) > 1:
                existing_keys = set()
                for index in range(1, min(len(links), len(dates))):
                    link = links[index]
                    date_value = dates[index]
                    if link and date_value:
                        existing_keys.add(f"{link}|{date_value}")
                return existing_keys

        # Cac sheet khac dedupe theo thread link + timestamp de van lay du email moi
        # trong cung thread nhung khong tao ban ghi trung lap khi chay lai.
        links = worksheet.col_values(5)
        dates = worksheet.col_values(6)
        if len(links) > 1 and len(dates) > 1:
            existing_keys = set()
            for index in range(1, min(len(links), len(dates))):
                link = links[index]
                date_value = dates[index]
                if link and date_value:
                    existing_keys.add(f"{link}|{date_value}")
            return existing_keys
    except Exception:
        pass
    return set()

def get_local_timezone() -> timezone:
    """Trả về timezone cố định theo offset giờ cấu hình."""
    return timezone(timedelta(hours=GMAIL_TIMEZONE_OFFSET_HOURS))


def format_datetime(timestamp_ms: int) -> Tuple[str, int, str]:
    """Định dạng datetime tu Gmail internalDate"""
    local_tz = get_local_timezone()
    dt = datetime.fromtimestamp(timestamp_ms / 1000, tz=timezone.utc).astimezone(local_tz)
    formatted = dt.strftime("%Y-%m-%d %H:%M:%S")
    
    # Tính tuần trong năm
    year_start = datetime(dt.year, 1, 1, tzinfo=dt.tzinfo)
    week_num = ((dt - year_start).days // 7) + 1
    
    date_type = dt.strftime("%Y-%m-%d")
    
    return formatted, week_num, date_type

_SHEET_WRITE_CHUNK = 500   # Số dong tối đa mỗi lần append_rows
_SHEET_WRITE_DELAY = 1.2   # Giây chờ giữa các chunk để tránh quota 429
_SHEET_MAX_RETRIES = 5     # Số lần retry khi gặp quota error


def _count_nonempty_first_col(worksheet) -> int:
    """Đếm số ô có dữ liệu ở cột A (bao gồm header)."""
    values = worksheet.col_values(1)
    return sum(1 for v in values if str(v).strip() != "")


def _append_rows_with_retry(worksheet, chunk: List[List]):
    """Gửi 1 chunk với retry exponential backoff khi gặp 429/quota."""
    delay = 2.0
    for attempt in range(_SHEET_MAX_RETRIES):
        try:
            print(f"      [>] Append {len(chunk)} hàng vào worksheet '{console_text(worksheet.title)}'...")
            worksheet.append_rows(chunk, value_input_option='USER_ENTERED', table_range='A1')
            print(f"      [OK] Append thành công")
            return
        except Exception as e:
            err = str(e).lower()
            is_quota = any(k in err for k in ('429', 'quota', 'rate', 'resource exhausted'))
            if is_quota and attempt < _SHEET_MAX_RETRIES - 1:
                wait = delay * (2 ** attempt)
                print(f"    [WARN] Quota exceeded, chờ {wait:.0f}s rồi thử lại ({attempt + 1}/{_SHEET_MAX_RETRIES})...")
                time.sleep(wait)
            else:
                print(f"      [ERR] Append failed: {repr(e)}")
                if not is_quota:
                    print("      [DEBUG] Lỗi không xác định, kiểm tra lại logic hoặc kết nối API.")
                raise


def _normalize_rows_for_sheet(rows: List[List], expected_columns: int) -> List[List]:
    """Chuẩn hóa số cột mỗi row để tránh lệch cột khi ghi Sheets."""
    normalized: List[List] = []
    for row in rows:
        if len(row) < expected_columns:
            normalized.append(row + [""] * (expected_columns - len(row)))
        else:
            normalized.append(row[:expected_columns])
    return normalized


def _canonical_header_name(value: str) -> str:
    """Chuẩn hóa tên header để so khớp ổn định giữa code và sheet."""
    return str(value or '').strip().upper()


def _remap_rows_to_worksheet_headers(
    rows: List[List],
    source_headers: List[str],
    worksheet_headers: List[str]
) -> List[List]:
    """Đổi thứ tự row theo header thực tế của worksheet để tránh lệch cột."""
    # Nếu header khớp nhau, không cần remap
    source_canon = [_canonical_header_name(h) for h in source_headers]
    ws_canon = [_canonical_header_name(h) for h in worksheet_headers]
    
    if source_canon == ws_canon:
        return rows
    
    source_index = {canon: idx for idx, canon in enumerate(source_canon)}
    remapped: List[List] = []
    
    for row in rows:
        output_row = [""] * len(worksheet_headers)
        for out_idx, ws_header_canon in enumerate(ws_canon):
            src_idx = source_index.get(ws_header_canon)
            if src_idx is not None and src_idx < len(row):
                output_row[out_idx] = row[src_idx]
        remapped.append(output_row)
    return remapped


def append_to_sheet(worksheet, rows: List[List], sheet_type: str) -> int:
    """Ghi nhiều dòng vào sheet theo chunk để tránh vượt payload và quota."""
    if not rows:
        print(f"  [INFO] Không có dữ liệu mới cho {sheet_type}")
        return 0

    # Lấy toàn bộ dữ liệu hiện có trong bảng tính
    existing_data = worksheet.get_all_values()
    existing_message_ids = set(row[STANDARD_HEADERS.index("Message_id")] for row in existing_data[1:] if len(row) > STANDARD_HEADERS.index("Message_id"))

    # Lọc các dòng dữ liệu mới không trùng lặp
    filtered_rows = [row for row in rows if row[STANDARD_HEADERS.index("Message_id")] not in existing_message_ids]

    if not filtered_rows:
        print(f"  [INFO] Không có dữ liệu mới sau khi lọc trùng lặp cho {sheet_type}")
        return 0

    source_headers = SHEET_HEADERS.get(sheet_type, STANDARD_HEADERS)
    filtered_rows = _normalize_rows_for_sheet(filtered_rows, len(source_headers))

    # Lấy header từ worksheet - chỉ lấy phần header thực tế (không lấy cột trống đầu)
    worksheet_headers_raw = worksheet.row_values(1) if worksheet.row_count > 0 else []
    worksheet_headers = worksheet_headers_raw[:len(source_headers)] if worksheet_headers_raw else []

    # Kiểm tra xem header có khớp không, nếu khớp thì không cần remap
    if worksheet_headers and _canonical_header_name(worksheet_headers[0] if worksheet_headers else '') == _canonical_header_name(source_headers[0]):
        if len(worksheet_headers) >= len(source_headers):
            filtered_rows = _remap_rows_to_worksheet_headers(filtered_rows, source_headers, worksheet_headers)
        else:
            filtered_rows = _normalize_rows_for_sheet(filtered_rows, len(source_headers))
    else:
        filtered_rows = _normalize_rows_for_sheet(filtered_rows, len(source_headers))

    print(f"  [WRITE] Bắt đầu ghi {len(filtered_rows)} dòng vào sheet '{console_text(worksheet.title)}' ({sheet_type})")
    total = len(filtered_rows)
    written = 0
    for i in range(0, total, _SHEET_WRITE_CHUNK):
        chunk = filtered_rows[i:i + _SHEET_WRITE_CHUNK]
        print(f"    Chunk {i//(_SHEET_WRITE_CHUNK)+1}: ghi dòng {i+1}-{min(i+_SHEET_WRITE_CHUNK, total)}")
        try:
            before_count = _count_nonempty_first_col(worksheet)
            _append_rows_with_retry(worksheet, chunk)
            expected_after = before_count + len(chunk)

            actual_after = before_count
            for _ in range(3):
                actual_after = _count_nonempty_first_col(worksheet)
                if actual_after >= expected_after:
                    break
                time.sleep(1)

            if actual_after < expected_after:
                raise RuntimeError(
                    "Append reported success but sheet row count did not increase "
                    f"as expected. before={before_count}, expected_after={expected_after}, "
                    f"actual_after={actual_after}"
                )

            written += len(chunk)
            print(f"  [OK] [{sheet_type}] Da ghi {written}/{total} dong")
        except Exception as e:
            print(f"  [ERR] [{sheet_type}] Loi ghi chunk {i}--{i + len(chunk)}: {e}")
            raise RuntimeError(
                f"Khong the ghi du lieu vao sheet '{worksheet.title}' ({sheet_type})"
            ) from e
        # Delay giữa các chunk để tránh quota.
        if i + _SHEET_WRITE_CHUNK < total:
            time.sleep(_SHEET_WRITE_DELAY)

    return written


def ensure_sheet_headers(spreadsheet):
    """Tạo sheet còn thiếu và luôn đồng bộ header theo cấu trúc chuẩn."""
    for sheet_type, config in SHEETS_CONFIG.items():
        headers = SHEET_HEADERS.get(sheet_type, STANDARD_HEADERS)
        sheet_name = config['name']
        is_new_sheet = False

        try:
            worksheet = call_with_retry(
                lambda: spreadsheet.worksheet(sheet_name),
                f"Sheets worksheet('{sheet_name}')"
            )
        except WorksheetNotFound:
            worksheet = call_with_retry(
                lambda: spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=len(headers)),
                f"Sheets add_worksheet('{sheet_name}')"
            )
            is_new_sheet = True
            print(f"  [OK] Da tao sheet moi: {console_text(sheet_name)}")

        if worksheet.col_count < len(headers):
            call_with_retry(
                lambda: worksheet.add_cols(len(headers) - worksheet.col_count),
                f"Sheets add_cols('{sheet_name}')"
            )

        # Luôn cập nhật header từ A1 để đảm bảo không bị lệch cột
        # Tính range A1:Z1 dựa trên số lượng header
        end_col = chr(ord('A') + len(headers) - 1) if len(headers) < 26 else 'Z'
        header_range = f"A1:{end_col}1"
        print(f"  [UPDATE] Cap nhat header cho sheet '{console_text(sheet_name)}' (range: {header_range})...")
        call_with_retry(
            lambda: worksheet.update(
                range_name=header_range,
                values=[headers],
                value_input_option='USER_ENTERED'
            ),
            f"Sheets update header('{sheet_name}')"
        )
        print(f"  [OK] Header da duoc cap nhat cho sheet: {console_text(sheet_name)}")


def handle_auth_error(error: Exception):
    """In loi xác thực theo dạng dễ xu ly."""
    error_text = str(error)
    error_repr = repr(error)
    joined = f"{error_repr}\n{error_text}"

    if "sheets.googleapis.com" in joined and "disabled" in joined.lower():
        print("[ERR] Loi xac thuc: Google Sheets API chua duoc bat cho project OAuth hien tai.")
        print("[=>] Vao link sau va bam Enable:")
        print("   https://console.developers.google.com/apis/api/sheets.googleapis.com/overview?project=492061659472")
        print("[=>] Sau khi bật, đợi 2-5 phút rồi chạy lại script.")
        return

    if isinstance(error, PermissionError):
        print("[ERR] Loi xac thuc: Khong co quyen mo Google Sheet.")
        print("[=>] Kiem tra lai SPREADSHEET_ID hoac tai khoan OAuth co quyen truy cap file Sheet.")
        return

    print(f"[ERR] Loi xac thuc: {error_repr}")


def load_state() -> Dict:
    """Đọc trạng thái lần đồng bộ gan nhat."""
    if not os.path.exists(STATE_FILE):
        return {}
    try:
        with open(STATE_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception:
        return {}


def save_state(state: Dict):
    """Lưu trạng thái đồng bộ để chạy incremental."""
    with open(STATE_FILE, 'w', encoding='utf-8') as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


_PROCESS_CHUNK = 200   # Xu ly và ghi sheet sau mỗi N email (tránh mất dữ liệu khi loi)


def _build_row(msg: Dict, sheet_type: str, existing_keys: Set[str]) -> Tuple[Optional[List], Optional[str]]:
    """Tạo 1 row tu message. Trả lý do skip nếu cần bỏ qua."""
    thread_link = f"https://mail.google.com/mail/u/0/#inbox/{msg['threadId']}"

    headers = extract_headers(msg['payload'], ['Subject', 'In-Reply-To', 'References'])
    # CTRIP/CTRIP CANCEL thường đi theo luồng reply, không bỏ qua các mail này.
    if sheet_type != "ctrip" and ('In-Reply-To' in headers or 'References' in headers):
        return None, "reply-thread"

    subject = headers.get('Subject', 'No Subject')
    body = decode_email_body(msg['payload'])
    order_id = extract_order_id(msg['label'], subject, body)
    pickup, dropoff = extract_route_points(msg['label'], subject, body)
    formatted_date, week_num, date_type = format_datetime(msg['internalDate'])

    if sheet_type == "klook_special":
        row_key = f"{thread_link}|{date_type}"
    elif sheet_type == "ctrip":
        row_key = f"{thread_link}|{formatted_date}"
    else:
        row_key = f"{thread_link}|{formatted_date}"
    if row_key in existing_keys:
        return None, f"duplicate-row-key:{row_key}"

    if sheet_type == "booking":
        row: List[Any] = [""] * len(BOOKING_HEADERS)
        row[0] = msg['label']
        row[1] = "EMAIL"
        row[2] = subject
        row[3] = body[:500]
        row[4] = thread_link
        row[5] = formatted_date
        row[6] = week_num
        row[7] = date_type
        row[8] = order_id
        row[19] = msg['id']
        row[21] = pickup
        row[22] = dropoff
        return row, None
    else:
        # STANDARD_HEADERS: NEN_TANG, PHAN_LOAI, SUBJECT, BODY, LINK, DATE, WEEKNUM, DATE_TYPE, ID_DH, P.I, P.I_TIME, LATEST_UPDATE_BY, LATEST_UPDATES, CONFIRM_TIME
        row: List[Any] = [""] * len(STANDARD_HEADERS)
        row[0] = msg['label']  # NEN_TANG
        row[1] = "EMAIL"       # PHAN_LOAI
        row[2] = subject       # SUBJECT
        row[3] = body[:500]    # BODY
        row[4] = thread_link   # LINK
        row[5] = formatted_date  # DATE
        row[6] = week_num      # WEEKNUM
        row[7] = date_type     # DATE_TYPE
        row[8] = order_id      # ID_DH
        # row[9] = ""  # P.I (empty)
        # row[10] = "" # P.I_TIME (empty)
        # row[11] = "" # LATEST_UPDATE_BY (empty)
        # row[12] = "" # LATEST_UPDATES (empty)
        # row[13] = "" # CONFIRM_TIME (empty)
        return row, None


def run_once() -> bool:
    """Chạy 1 vòng đồng bộ."""
    print("[START] Bat dau lay du lieu email tu Gmail...\n")
    state = load_state()
    if LOGGER_IGNORE_STATE:
        state = {}
        print("[MODE] Bo qua state cu theo LOGGER_IGNORE_STATE=true\n")

    after_epoch_seconds: Optional[int] = None
    days_back_anchor_epoch: Optional[int] = None

    if GMAIL_AFTER_DATE:
        after_dt: Optional[datetime] = None
        accepted_formats = ('%Y-%m-%d', '%d/%m/%Y', '%d-%m-%Y')
        for fmt in accepted_formats:
            try:
                after_dt = datetime.strptime(GMAIL_AFTER_DATE, fmt)
                break
            except ValueError:
                continue

        if after_dt is not None:
            # Diễn giải mốc ngày theo timezone cấu hình (mặc định UTC+7),
            # tránh lệch ngày khi chạy trên runner UTC.
            local_tz = get_local_timezone()
            after_epoch_seconds = int(after_dt.replace(tzinfo=local_tz).timestamp())
            print(
                f"[MODE] Lay du lieu tu ngay {after_dt.strftime('%Y-%m-%d')} den hien tai, "
                "khong gioi han so mail\n"
            )
        else:
            print(
                f"[WARN] GMAIL_AFTER_DATE khong hop le '{GMAIL_AFTER_DATE}'. "
                "Dinh dang hop le: YYYY-MM-DD, DD/MM/YYYY, DD-MM-YYYY. Bo qua."
            )
    elif GMAIL_FETCH_ALL:
        after_epoch_seconds = None
    elif GMAIL_HOURS_BACK > 0:
        after_epoch_seconds = int((datetime.now() - timedelta(hours=GMAIL_HOURS_BACK)).timestamp())
    elif not LOGGER_IGNORE_STATE:
        after_epoch_seconds = state.get('last_success_epoch_seconds')
        if not after_epoch_seconds and GMAIL_DAYS_BACK > 0:
            days_back_anchor_epoch = int((datetime.now() - timedelta(days=GMAIL_DAYS_BACK)).timestamp())
            after_epoch_seconds = days_back_anchor_epoch

    if GMAIL_AFTER_DATE and after_epoch_seconds:
        dt_since = datetime.fromtimestamp(after_epoch_seconds)
        print(f"[TIME] Quet tu moc ngay cau hinh: {dt_since.strftime('%Y-%m-%d %H:%M:%S')} den hien tai\n")
    elif after_epoch_seconds and not GMAIL_FETCH_ALL:
        dt_since = datetime.fromtimestamp(after_epoch_seconds)
        if GMAIL_HOURS_BACK > 0:
            print(f"[TIME] Che do quet lui {GMAIL_HOURS_BACK} gio gan nhat tu: {dt_since.strftime('%Y-%m-%d %H:%M:%S')}\n")
        elif days_back_anchor_epoch:
            print(
                f"[TIME] Che do quet {GMAIL_DAYS_BACK} ngay gan nhat tu: "
                f"{dt_since.strftime('%Y-%m-%d %H:%M:%S')}\n"
            )
        else:
            print(f"[TIME] Che do incremental tu: {dt_since.strftime('%Y-%m-%d %H:%M:%S')}\n")
    elif GMAIL_FETCH_ALL:
        print("[MODE] Che do FETCH ALL - lay toan bo email khong gioi han thoi gian\n")
    else:
        print(f"[MODE] Che do full scan theo cua so {GMAIL_HOURS_BACK} gio gan nhat\n")

    # Xac thuc
    try:
        gmail_service = authenticate_gmail()
        gs = authenticate_gspread()
        spreadsheet = open_spreadsheet_with_retry(gs, SPREADSHEET_ID)
        print(
            f"[INFO] Dang ghi vao spreadsheet: title='{console_text(spreadsheet.title)}', "
            f"id='{SPREADSHEET_ID}'"
        )
        call_with_retry(
            lambda: ensure_sheet_headers(spreadsheet),
            "Sheets ensure_sheet_headers"
        )
        print("[OK] Xac thuc thanh cong\n")
    except Exception as e:
        handle_auth_error(e)
        return False

    run_max_internal_date_ms = 0
    total_new_rows = 0

    labels_result = call_with_retry(
        lambda: gmail_service.users().labels().list(userId='me').execute(),
        "Gmail labels.list"
    )
    label_name_to_id = {
        lbl['name']: lbl['id']
        for lbl in labels_result.get('labels', [])
        if 'name' in lbl and 'id' in lbl
    }

    _hours_back = 0 if (GMAIL_FETCH_ALL or GMAIL_AFTER_DATE) else GMAIL_HOURS_BACK
    if GMAIL_FETCH_ALL or GMAIL_AFTER_DATE:
        _max_results = 0
    else:
        _max_results = GMAIL_MAX_RESULTS_PER_LABEL

    started_from_sheet = (LOGGER_START_FROM_SHEET == "")
    if LOGGER_START_FROM_SHEET:
        print(f"[MODE] Bat dau xu ly tu sheet_type='{LOGGER_START_FROM_SHEET}'\n")

    # Xu ly tung loại sheet
    for sheet_type, config in SHEETS_CONFIG.items():
        if not started_from_sheet:
            if sheet_type == LOGGER_START_FROM_SHEET:
                started_from_sheet = True
            else:
                print(f"[SKIP] Bo qua sheet '{sheet_type}' vi LOGGER_START_FROM_SHEET='{LOGGER_START_FROM_SHEET}'")
                continue

        print(f"\n[SHEET] Xu ly {sheet_type.upper()}:")
        print(f"   Sheet: {console_text(config['name'])}")
        print(f"   Labels: {', '.join(console_text(label) for label in config['labels'])}")

        # Đọc checkpoint: label nào đã xong trong lần chạy FETCH_ALL này.
        done_labels: Set[str] = set(state.get(f'done_labels_{sheet_type}', []))

        try:
            worksheet = call_with_retry(
                lambda: spreadsheet.worksheet(config['name']),
                f"Sheets worksheet('{config['name']}')"
            )
        except Exception as e:
            print(f"  [ERR] Sheet '{console_text(config['name'])}' khong tim thay: {e}")
            continue

        existing_keys = get_existing_row_keys(worksheet, sheet_type)
        worksheet_url = f"https://docs.google.com/spreadsheets/d/{SPREADSHEET_ID}/edit#gid={worksheet.id}"
        print(f"   Tab URL: {worksheet_url}")
        print(f"   Existing dedupe keys: {len(existing_keys)}")
        sheet_new_rows = 0

        for label in config['labels']:
            if label in done_labels:
                print(f"  [OK] Label '{console_text(label)}' da hoan thanh (checkpoint), bo qua")
                continue

            buffer: List[List] = []
            label_new = 0
            fetched_any = False

            def _on_msg_batch(batch_messages: List[Dict], _label: str):
                nonlocal buffer, label_new, sheet_new_rows, run_max_internal_date_ms, fetched_any
                print(f"    [CALLBACK] _on_msg_batch called with {len(batch_messages)} messages from label '{console_text(_label)}'")
                if batch_messages:
                    fetched_any = True

                for msg in batch_messages:
                    row, skip_reason = _build_row(msg, sheet_type, existing_keys)
                    if row is None:
                        print(f"      [SKIP] {skip_reason}")
                        continue

                    thread_link = f"https://mail.google.com/mail/u/0/#inbox/{msg['threadId']}"
                    formatted_date, _, date_type = format_datetime(msg['internalDate'])
                    if sheet_type == "klook_special":
                        row_key = f"{thread_link}|{date_type}"
                    elif sheet_type == "ctrip":
                        row_key = f"{thread_link}|{formatted_date}"
                    else:
                        row_key = f"{thread_link}|{formatted_date}"
                    existing_keys.add(row_key)
                    buffer.append(row)
                    run_max_internal_date_ms = max(run_max_internal_date_ms, msg['internalDate'])
                    print(f"      [ADD] Row added to buffer. Buffer size: {len(buffer)}/{_PROCESS_CHUNK}")

                    if len(buffer) >= _PROCESS_CHUNK:
                        print(f"      [FLUSH] Buffer reached limit, calling append_to_sheet...")
                        written_now = append_to_sheet(worksheet, buffer, sheet_type)
                        label_new += written_now
                        sheet_new_rows += written_now
                        print(f"  [...] [{console_text(label)}] Da ghi tam {label_new} dong")
                        buffer = []

            # Lấy email cho 1 label và xu ly ngay theo batch fetch
            print(f"    [GET_MSGS] Calling get_gmail_messages for label '{console_text(label)}'...")
            get_gmail_messages(
                gmail_service,
                [label],
                hours_back=_hours_back,
                max_results_per_label=_max_results,
                after_epoch_seconds=after_epoch_seconds,
                label_name_to_id=label_name_to_id,
                on_message_batch=_on_msg_batch
            )
            print(f"    [GET_MSGS] Completed. fetched_any={fetched_any}, buffer size={len(buffer)}")

            if not fetched_any and not buffer:
                print(f"  [INFO] Không có email mới từ '{console_text(label)}'")
                done_labels.add(label)
                state[f'done_labels_{sheet_type}'] = list(done_labels)
                save_state(state)
                continue

            # Ghi phần còn lại
            print(f"    [FINAL] Buffer final check: fetched_any={fetched_any}, buffer size={len(buffer)}")
            if buffer:
                print(f"    [FINAL_APPEND] Appending final {len(buffer)} rows to sheet...")
                written_now = append_to_sheet(worksheet, buffer, sheet_type)
                label_new += written_now
                sheet_new_rows += written_now
                print(f"    [FINAL_APPEND] Completed")

            print(f"  [OK] Label '{console_text(label)}': {label_new} dong moi")

            # Lưu checkpoint label này đã xong
            done_labels.add(label)
            state[f'done_labels_{sheet_type}'] = list(done_labels)
            save_state(state)

        total_new_rows += sheet_new_rows
        print(f"  [STATS] Tong {console_text(config['name'])}: {sheet_new_rows} dong moi")

        # Xoá checkpoint label khi sheet này hoàn thành
        state.pop(f'done_labels_{sheet_type}', None)
        save_state(state)

    # Lưu mốc sync để lần sau chạy incremental, tránh quét lại toàn bộ cửa sổ bootstrap.
    if run_max_internal_date_ms > 0:
        new_after_epoch_seconds = max(0, (run_max_internal_date_ms // 1000) - 60)
        state['last_success_epoch_seconds'] = new_after_epoch_seconds
        state['last_success_iso'] = datetime.fromtimestamp(new_after_epoch_seconds).isoformat()
        save_state(state)

    print(f"\n[DONE] Hoan thanh! Tong dong moi: {total_new_rows}")
    return True


def run_daemon():
    """Chạy liên tục để tự động đồng bộ gần realtime."""
    print(f"🔁 Daemon mode đang chạy, chu kỳ {LOGGER_POLL_SECONDS}s")
    while True:
        try:
            run_once()
        except Exception as e:
            print(f"[ERR] Lỗi vòng lặp daemon: {repr(e)}")
        time.sleep(max(5, LOGGER_POLL_SECONDS))

# ==================== MAIN FUNCTION ====================
def main():
    """Hàm chính"""
    if LOGGER_RUN_MODE == 'daemon':
        run_daemon()
        return

    ok = run_once()
    if not ok:
        sys.exit(1)

if __name__ == "__main__":
    main()
