# app_gv_attendance_monthly.py
import os
import io
import re
import time
import base64
import urllib.parse
import unicodedata
import datetime
from calendar import monthrange
from difflib import get_close_matches

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from PIL import Image
import qrcode
import pandas as pd
import altair as alt

try:
    from streamlit_geolocation import streamlit_geolocation
except Exception:
    streamlit_geolocation = None

try:
    from geopy.distance import geodesic
except Exception:
    geodesic = None


# ===================== CẤU HÌNH CHUNG =====================
# App điểm danh GIẢNG VIÊN.
# Giảng viên KHÔNG cần đăng nhập.
# Quản trị đăng nhập để tạo QR cố định theo cơ sở, tìm kiếm, thống kê, kiểm tra dữ liệu.
MSGV_PREFIX = st.secrets.get("SESSION_PREFIX", "0607")
SHEET_KEY = st.secrets["SHEET_KEY"]
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

BASE_INFO_COLUMNS = ["MSGV", "Họ và tên", "Đơn vị", "Bộ môn"]

LOCATIONS = {
    "Cơ sở 1 - Hồng Bàng": {
        "code": "CS1",
        "lat": 10.754665,
        "lon": 106.663381,
        "radius": 500,
        "address": "217 Hồng Bàng, Phường Chợ Lớn, TP.HCM",
    },
    "Cơ sở 2 - Đinh Tiên Hoàng": {
        "code": "CS2",
        "lat": 10.785434,
        "lon": 106.702667,
        "radius": 400,
        "address": "43 Đinh Tiên Hoàng, Phường Sài Gòn, TP.HCM",
    },
}

LOCATION_BY_CODE = {v["code"]: k for k, v in LOCATIONS.items()}

st.set_page_config(page_title="Điểm danh giảng viên QR", layout="wide")


# ===================== CSS IN MINH CHỨNG =====================
st.markdown("""
<style>
html, body, .stApp, [class*="css"] {
    font-size: 18px !important;
    color: #000000 !important;
}
h1 {
    font-size: 38px !important;
    font-weight: 900 !important;
    color: #000000 !important;
}
h2, h3 {
    font-size: 28px !important;
    font-weight: 900 !important;
    color: #000000 !important;
}
p, span, div, label {
    color: #000000 !important;
}
label {
    font-size: 18px !important;
    font-weight: 800 !important;
}
input, textarea {
    font-size: 18px !important;
    font-weight: 700 !important;
    color: #000000 !important;
}
button {
    font-size: 18px !important;
    font-weight: 800 !important;
}
[data-testid="stMetricLabel"] {
    font-size: 18px !important;
    font-weight: 900 !important;
    color: #000000 !important;
}
[data-testid="stMetricValue"] {
    font-size: 40px !important;
    font-weight: 900 !important;
    color: #000000 !important;
}
[data-testid="stCaptionContainer"], small {
    font-size: 16px !important;
    font-weight: 700 !important;
    color: #000000 !important;
}
[data-testid="stAlert"] {
    font-size: 18px !important;
    font-weight: 800 !important;
}
section[data-testid="stSidebar"] * {
    font-size: 16px !important;
    font-weight: 700 !important;
    color: #000000 !important;
}
.footer-dhn {
    font-size: 14px !important;
    font-weight: 800 !important;
    color: #000000 !important;
}
@media print {
    html, body, .stApp {
        background: #ffffff !important;
        color: #000000 !important;
    }
    * {
        color: #000000 !important;
        text-shadow: none !important;
        box-shadow: none !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ===================== NGÀY THÁNG =====================
def now_vn():
    return datetime.datetime.now(VN_TZ)


def today_date():
    return now_vn().date()


def today_header():
    return today_date().strftime("%d/%m/%Y")


def month_sheet_name(d=None):
    if d is None:
        d = today_date()
    return d.strftime("%B-%Y")


def previous_month_sheet_name(d=None):
    if d is None:
        d = today_date()
    first_day = d.replace(day=1)
    prev_day = first_day - datetime.timedelta(days=1)
    return month_sheet_name(prev_day)


def date_header(d: datetime.date) -> str:
    return d.strftime("%d/%m/%Y")


def current_month_working_day_headers(d=None):
    if d is None:
        d = today_date()
    year, month = d.year, d.month
    last_day = monthrange(year, month)[1]
    headers = []
    for day in range(1, last_day + 1):
        cur = datetime.date(year, month, day)
        if cur.weekday() != 6:  # bỏ Chủ nhật
            headers.append(date_header(cur))
    return headers


# ===================== TIỆN ÍCH =====================
def get_query_params():
    if hasattr(st, "query_params"):
        return dict(st.query_params)
    raw = st.experimental_get_query_params()
    return {k: (v[0] if isinstance(v, list) and v else v) for k, v in raw.items()}


def normalize_name(name: str):
    return " ".join(w.capitalize() for w in (name or "").strip().split())


def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", s)


def norm_search(s: str) -> str:
    return " ".join(strip_accents(s).lower().split())


def _google_api_retry(callable_fn, retries=3, delay=1.5):
    last_error = None
    for attempt in range(retries):
        try:
            return callable_fn()
        except Exception as e:
            last_error = e
            msg = str(e)
            transient = any(code in msg for code in ["[500]", "[503]", "[429]", "Internal error", "Quota", "timeout", "Timeout"])
            if not transient or attempt == retries - 1:
                raise
            time.sleep(delay * (attempt + 1))
    raise last_error


def get_base_url():
    return (
        st.secrets.get("WRAPPER_URL")
        or st.secrets.get("APP_BASE_URL")
        or st.secrets.get("google_service_account", {}).get("app_base_url")
        or "https://qrlecturer.streamlit.app"
    )


def attendance_value(time_str: str, campus_name: str) -> str:
    return f"{time_str} | {campus_name}"


def parse_attendance_value(value: str):
    value = str(value or "").strip()
    if not value:
        return "", ""
    if "|" in value:
        parts = [p.strip() for p in value.split("|", 1)]
        return parts[0], parts[1]
    return value, ""


def is_date_col(header: str) -> bool:
    return bool(re.match(r"^\d{2}/\d{2}/\d{4}$", str(header or "").strip()))


# ===================== ĐĂNG NHẬP QUẢN TRỊ =====================
def _get_admin_pw():
    if "ADMIN_PASSWORD" in st.secrets:
        return st.secrets["ADMIN_PASSWORD"]
    if "teacher_password" in st.secrets:
        return st.secrets["teacher_password"]
    if "google_service_account" in st.secrets:
        maybe = st.secrets["google_service_account"].get("teacher_password")
        if maybe:
            return maybe
    return os.getenv("ADMIN_PASSWORD") or os.getenv("TEACHER_PASSWORD")


def admin_unlocked() -> bool:
    return bool(st.session_state.get("admin_unlocked"))


def render_admin_auth():
    with st.sidebar:
        st.header("Quản trị")
        if admin_unlocked():
            st.success("Đã đăng nhập quản trị")
            if st.button("Đăng xuất"):
                st.session_state.clear()
                st.rerun()
        else:
            pw_input = st.text_input("Mật khẩu quản trị", type="password", key="pw_admin")
            if st.button("Đăng nhập", type="primary", use_container_width=True):
                if _get_admin_pw() and pw_input == _get_admin_pw():
                    st.session_state["admin_unlocked"] = True
                    st.rerun()
                else:
                    st.warning("Sai mật khẩu hoặc chưa cấu hình ADMIN_PASSWORD trong Secrets/ENV.")


# ===================== GOOGLE SHEETS =====================
@st.cache_resource
def _get_gspread_client():
    if "google_service_account" not in st.secrets:
        raise RuntimeError("Thiếu block [google_service_account] trong Secrets.")

    cred = dict(st.secrets["google_service_account"])
    pk = cred.get("private_key", "")
    if not pk:
        raise RuntimeError("Secrets thiếu private_key.")

    if "\\n" in pk:
        pk = pk.replace("\\n", "\n")
    pk = pk.replace("\r\n", "\n").replace("\r", "\n")

    header = "-----BEGIN PRIVATE KEY-----"
    footer = "-----END PRIVATE KEY-----"
    if header not in pk or footer not in pk:
        raise RuntimeError("private_key thiếu BEGIN/END.")

    lines = [ln.strip() for ln in pk.split("\n")]
    h_idx = lines.index(header)
    f_idx = lines.index(footer)
    body_raw = re.sub(r"[^A-Za-z0-9+/=]", "", "".join([ln for ln in lines[h_idx + 1:f_idx] if ln]))
    body = body_raw.replace("=", "")
    if not body:
        raise RuntimeError("private_key rỗng sau khi làm sạch.")

    rem = len(body) % 4
    if rem:
        body += "=" * (4 - rem)
    base64.b64decode(body, validate=True)

    pk_clean = header + "\n" + "\n".join(body[i:i+64] for i in range(0, len(body), 64)) + "\n" + footer + "\n"
    cred["private_key"] = pk_clean

    creds = Credentials.from_service_account_info(cred, scopes=SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    client = _get_gspread_client()
    return _google_api_retry(lambda: client.open_by_key(SHEET_KEY))


def get_worksheet_or_none(ss, title: str):
    try:
        return _google_api_retry(lambda: ss.worksheet(title))
    except Exception:
        return None


def get_sheet_headers(ws):
    return _google_api_retry(lambda: ws.row_values(1))


def load_records(ws):
    return _google_api_retry(lambda: ws.get_all_records(expected_headers=None, default_blank=""))


def find_header_col(ws, header_name):
    return _google_api_retry(lambda: ws.find(header_name)).col


def find_staff_row(ws, msgv_full: str):
    try:
        return _google_api_retry(lambda: ws.find(msgv_full))
    except Exception:
        return None


def ensure_month_headers(ws, d=None):
    if d is None:
        d = today_date()

    headers = get_sheet_headers(ws)
    if not headers:
        headers = []

    changed = False
    for col in BASE_INFO_COLUMNS:
        if col not in headers:
            headers.append(col)
            changed = True

    for col in current_month_working_day_headers(d):
        if col not in headers:
            headers.append(col)
            changed = True

    if changed:
        _google_api_retry(lambda: ws.update("1:1", [headers]))

    return headers


def copy_staff_list_from_previous_sheet(ss, new_ws, d=None):
    if d is None:
        d = today_date()

    prev_title = previous_month_sheet_name(d)
    prev_ws = get_worksheet_or_none(ss, prev_title)

    if prev_ws is None:
        return False

    prev_headers = get_sheet_headers(prev_ws)
    records = load_records(prev_ws)

    base_cols = [c for c in BASE_INFO_COLUMNS if c in prev_headers]
    if not base_cols:
        return False

    new_headers = BASE_INFO_COLUMNS + current_month_working_day_headers(d)
    rows = [new_headers]

    for r in records:
        rows.append([r.get(c, "") for c in BASE_INFO_COLUMNS] + [""] * len(current_month_working_day_headers(d)))

    _google_api_retry(lambda: new_ws.clear())
    _google_api_retry(lambda: new_ws.update("A1", rows))
    return True


def create_blank_month_sheet(ws, d=None):
    if d is None:
        d = today_date()

    headers = BASE_INFO_COLUMNS + current_month_working_day_headers(d)
    _google_api_retry(lambda: ws.update("A1", [headers]))


def get_or_create_month_sheet(d=None):
    if d is None:
        d = today_date()

    ss = get_spreadsheet()
    title = month_sheet_name(d)
    ws = get_worksheet_or_none(ss, title)

    if ws is not None:
        ensure_month_headers(ws, d)
        return ws

    ws = _google_api_retry(lambda: ss.add_worksheet(title=title, rows=300, cols=50))

    copied = copy_staff_list_from_previous_sheet(ss, ws, d)
    if not copied:
        create_blank_month_sheet(ws, d)

    ensure_month_headers(ws, d)
    return ws


def ensure_today_column(ws):
    if today_date().weekday() == 6:
        raise RuntimeError("Hôm nay là Chủ nhật, hệ thống không mở cột điểm danh.")

    headers = ensure_month_headers(ws, today_date())
    h = today_header()

    if h not in headers:
        headers.append(h)
        _google_api_retry(lambda: ws.update("1:1", [headers]))

    return headers.index(h) + 1


# ===================== KIỂM TRA GPS THEO CƠ SỞ =====================
def render_location_check(campus_code: str):
    campus_name = LOCATION_BY_CODE.get(campus_code)
    if not campus_name:
        st.error("Cơ sở điểm danh không hợp lệ.")
        st.stop()

    campus = LOCATIONS[campus_name]
    st.info(f"Cơ sở điểm danh: {campus_name} - {campus['address']}")

    if streamlit_geolocation is None or geodesic is None:
        st.error("Ứng dụng chưa cài đủ thư viện kiểm tra vị trí. Cần cài streamlit-geolocation và geopy.")
        st.code("pip install streamlit-geolocation geopy")
        st.stop()

    st.caption("Vui lòng cho phép trình duyệt truy cập vị trí để xác thực điểm danh.")
    location = streamlit_geolocation()

    if not location:
        st.warning("Chưa nhận được vị trí. Vui lòng bật định vị và cho phép trình duyệt truy cập vị trí.")
        st.stop()

    lat = location.get("latitude")
    lon = location.get("longitude")

    if lat is None or lon is None:
        st.warning("Không lấy được tọa độ GPS từ thiết bị. Vui lòng thử lại.")
        st.stop()

    staff_loc = (float(lat), float(lon))
    campus_loc = (campus["lat"], campus["lon"])
    distance = geodesic(staff_loc, campus_loc).meters

    st.caption(f"Khoảng cách đến cơ sở: {distance:.0f} m. Phạm vi cho phép: {campus['radius']} m.")

    if distance > campus["radius"]:
        st.error(f"Bạn đang ngoài phạm vi điểm danh của {campus_name}.")
        st.stop()

    st.success("Vị trí hợp lệ. Bạn có thể tiếp tục điểm danh.")


# ===================== GIAO DIỆN QUẢN TRỊ =====================
def render_tab_qr():
    st.subheader("Tạo mã QR cố định theo cơ sở")
    st.caption("Mỗi cơ sở chỉ cần tạo QR một lần. QR tự ghi nhận theo ngày hiện hành và sheet tháng hiện hành.")

    campus_name = st.selectbox("Chọn cơ sở", list(LOCATIONS.keys()))
    campus_code = LOCATIONS[campus_name]["code"]

    st.info(f"Cơ sở: {campus_name} - {LOCATIONS[campus_name]['address']}")
    st.info(f"Sheet tháng hiện tại: {month_sheet_name()} | Ngày hiện tại: {today_header()}")

    if st.button("Tạo mã QR cố định", type="primary", use_container_width=True):
        base_url = get_base_url()
        qr_data = (
            f"{base_url}/?gv=1"
            f"&coso={urllib.parse.quote(campus_code)}"
        )

        qr = qrcode.make(qr_data)
        buf = io.BytesIO()
        qr.save(buf, format="PNG")
        buf.seek(0)
        img = Image.open(buf)

        st.image(img, caption=f"QR cố định cho {campus_name}", width=380)
        st.caption("Có thể in/dán QR này tại cơ sở. Không cần tạo lại mỗi ngày.")
        with st.expander("Xem link QR"):
            st.code(qr_data)


def find_staff_candidates(records, query: str):
    q = (query or "").strip()
    if not q:
        return []

    if q.isdigit() and len(q) == 4:
        return [r for r in records if str(r.get("MSGV", "")).strip().endswith(q)]

    qn = norm_search(q)
    name_col = "Họ và tên"
    contains = [r for r in records if qn in norm_search(r.get(name_col, ""))]
    if contains:
        return contains

    names = [r.get(name_col, "") for r in records]
    name_map = {n: r for n, r in zip(names, records)}
    close = get_close_matches(q, names, n=5, cutoff=0.6)
    if not close:
        names_no = [norm_search(n) for n in names]
        name_map_no = {norm_search(n): n for n in names}
        close_no = get_close_matches(qn, names_no, n=5, cutoff=0.6)
        close = [name_map_no[c] for c in close_no]

    return [name_map[n] for n in close if n in name_map]


def render_tab_search():
    st.subheader("Tìm kiếm thông tin giảng viên")
    st.caption(f"Đang tra cứu trong sheet tháng: {month_sheet_name()}")

    q = st.text_input("Nhập 4 số cuối MSGV hoặc họ tên", placeholder="Ví dụ: 1234 hoặc Nguyễn Văn A")

    if st.button("Tìm", use_container_width=True):
        try:
            ws = get_or_create_month_sheet()
            records = load_records(ws)
            results = find_staff_candidates(records, q)

            if not results:
                st.warning("Không tìm thấy kết quả phù hợp.")
                return

            st.success(f"Tìm thấy {len(results)} kết quả.")
            df = pd.DataFrame(results)
            st.dataframe(df, use_container_width=True)

        except Exception as e:
            st.error(f"Lỗi khi tìm kiếm: {e}")


def render_tab_stats():
    st.subheader("Thống kê điểm danh giảng viên")
    st.caption("Thống kê theo ngày, giờ và cơ sở. Không thống kê có mặt/vắng mặt.")

    try:
        ss = get_spreadsheet()
        worksheet_titles = [ws.title for ws in _google_api_retry(lambda: ss.worksheets())]
        month_titles = [t for t in worksheet_titles if re.match(r"^[A-Za-z]+-\d{4}$", t)]

        current_title = month_sheet_name()
        if current_title not in month_titles:
            get_or_create_month_sheet()
            month_titles.append(current_title)

        month_titles = sorted(set(month_titles))
        selected_month = st.selectbox(
            "Chọn tháng",
            month_titles,
            index=month_titles.index(current_title) if current_title in month_titles else len(month_titles) - 1,
        )

        ws = get_worksheet_or_none(ss, selected_month)
        if ws is None:
            st.warning("Không tìm thấy sheet tháng đã chọn.")
            return

        headers = ensure_month_headers(ws)
        records = load_records(ws)

        date_cols = [h for h in headers if is_date_col(h)]
        if not date_cols:
            st.warning("Chưa có cột ngày điểm danh.")
            return

        selected_day = st.selectbox(
            "Chọn ngày",
            date_cols,
            index=date_cols.index(today_header()) if today_header() in date_cols else 0,
        )

        rows = []
        for r in records:
            raw = str(r.get(selected_day, "") or "").strip()
            if not raw:
                continue

            hour, campus = parse_attendance_value(raw)
            rows.append({
                "Tháng": selected_month,
                "Ngày": selected_day,
                "Giờ điểm danh": hour,
                "Cơ sở": campus,
                "MSGV": r.get("MSGV", ""),
                "Họ và tên": r.get("Họ và tên", ""),
                "Đơn vị": r.get("Đơn vị", ""),
                "Bộ môn": r.get("Bộ môn", ""),
            })

        st.metric("Số lượt đã điểm danh", len(rows))

        if not rows:
            st.info("Ngày này chưa có dữ liệu điểm danh.")
            return

        df = pd.DataFrame(rows)
        st.dataframe(df, use_container_width=True)

        campus_df = df.groupby("Cơ sở", dropna=False).size().reset_index(name="Số lượt")
        if not campus_df.empty:
            chart = alt.Chart(campus_df).mark_bar().encode(
                x=alt.X("Cơ sở:N", title="Cơ sở"),
                y=alt.Y("Số lượt:Q", title="Số lượt điểm danh"),
                tooltip=["Cơ sở", "Số lượt"],
            ).properties(height=360)
            st.altair_chart(chart, use_container_width=True)

    except Exception as e:
        st.error(f"Lỗi khi lấy thống kê: {e}")


def render_tab_data_setup():
    st.subheader("Cấu trúc dữ liệu theo tháng")
    st.caption("Hệ thống tự dùng/tạo sheet tháng hiện hành, ví dụ June-2026, July-2026. Mỗi tháng tự có các cột ngày, bỏ Chủ nhật.")

    try:
        ws = get_or_create_month_sheet()
        headers = ensure_month_headers(ws)

        st.success(f"Đã kiểm tra/cập nhật sheet tháng: {ws.title}")
        st.write("Các cột hiện có:")
        st.dataframe(pd.DataFrame({"Tên cột": headers}), use_container_width=True)

    except Exception as e:
        st.error(f"Lỗi khi kiểm tra cấu trúc Sheet: {e}")


# ===================== MÀN HÌNH GIẢNG VIÊN ĐIỂM DANH =====================
def render_gv_attendance():
    qp = get_query_params()
    campus_code = qp.get("coso", "CS1")

    st.title("Điểm danh giảng viên")

    if today_date().weekday() == 6:
        st.error("Hôm nay là Chủ nhật, hệ thống không mở điểm danh.")
        st.stop()

    campus_name = LOCATION_BY_CODE.get(campus_code)
    if not campus_name:
        st.error("Cơ sở điểm danh không hợp lệ.")
        st.stop()

    st.info(f"Ngày điểm danh: {today_header()}")
    st.info(f"Sheet tháng: {month_sheet_name()}")

    render_location_check(campus_code)

    msgv_suffix = st.text_input(
        "Nhập 4 số cuối MSGV",
        placeholder="Ví dụ: 1234",
        max_chars=4,
        help=f"Mã đầy đủ sẽ là {MSGV_PREFIX} + 4 số cuối bạn nhập",
    )
    hoten = st.text_input("Nhập họ và tên")

    if msgv_suffix.strip().isdigit():
        st.caption(f"MSGV đầy đủ: {MSGV_PREFIX}{msgv_suffix.strip().zfill(4)}")

    if st.button("Xác nhận điểm danh", type="primary", use_container_width=True):
        if not msgv_suffix.strip().isdigit() or len(msgv_suffix.strip()) != 4:
            st.warning("Vui lòng nhập đúng 4 số cuối MSGV.")
            st.stop()

        if not hoten.strip():
            st.warning("Vui lòng nhập họ và tên.")
            st.stop()

        msgv_full = f"{MSGV_PREFIX}{msgv_suffix.strip().zfill(4)}"

        try:
            ws = get_or_create_month_sheet()
            today_col = ensure_today_column(ws)

            staff_cell = find_staff_row(ws, msgv_full)
            if not staff_cell:
                st.error(f"Không tìm thấy MSGV {msgv_full} trong danh sách.")
                st.stop()

            name_col = find_header_col(ws, "Họ và tên")
            name_in_sheet = (_google_api_retry(lambda: ws.cell(staff_cell.row, name_col)).value or "").strip()

            if normalize_name(name_in_sheet) != normalize_name(hoten):
                st.error("Họ tên không khớp với MSGV trong danh sách.")
                st.stop()

            current_value = (_google_api_retry(lambda: ws.cell(staff_cell.row, today_col)).value or "").strip()
            if current_value:
                hour, campus = parse_attendance_value(current_value)
                st.info(f"MSGV {msgv_full} đã điểm danh ngày {today_header()} lúc {hour} tại {campus}.")
                st.stop()

            time_str = now_vn().strftime("%H:%M:%S")
            value = attendance_value(time_str, campus_name)
            _google_api_retry(lambda: ws.update_cell(staff_cell.row, today_col, value))

            st.success(f"Điểm danh thành công! MSGV {msgv_full}, ngày {today_header()}, lúc {time_str}, tại {campus_name}.")

        except Exception as e:
            st.error(f"Lỗi khi điểm danh: {e}")


# ===================== ĐIỀU HƯỚNG =====================
qp = get_query_params()

if qp.get("gv") == "1":
    render_gv_attendance()
    st.stop()

render_admin_auth()

st.title("Hệ thống điểm danh QR cho giảng viên")

if not admin_unlocked():
    st.error("Vui lòng đăng nhập quản trị để sử dụng các chức năng quản lý.")
    st.stop()

with st.sidebar:
    st.markdown("**Điều hướng**")
    menu = st.radio(
        "Chọn mục",
        options=["Tạo QR cố định", "Tìm kiếm giảng viên", "Thống kê điểm danh", "Cấu trúc dữ liệu"],
        index=0,
        label_visibility="collapsed",
    )

if menu == "Tạo QR cố định":
    render_tab_qr()
elif menu == "Tìm kiếm giảng viên":
    render_tab_search()
elif menu == "Thống kê điểm danh":
    render_tab_stats()
else:
    render_tab_data_setup()

st.markdown(
    """
    <style>
    .footer-dhn {
        position: fixed;
        left: 0; right: 0; bottom: 0;
        padding: 8px 16px;
        background: rgba(0,0,0,0.04);
        color: #000000;
        font-size: 12px;
        text-align: center;
        z-index: 1000;
        border-top: 1px solid rgba(0,0,0,0.1);
        width: 100%;
    }
    </style>
    <div class="footer-dhn">Copyright © 2025 Bản quyền thuộc về <strong>TS. Đào Hồng Nam - Đại học Y Dược Thành phố Hồ Chí Minh</strong></div>
    """,
    unsafe_allow_html=True,
)
