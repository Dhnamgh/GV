# app.py
import os, io, re, time, base64, urllib.parse, unicodedata, datetime
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


MSGV_PREFIX = st.secrets.get("SESSION_PREFIX", "0607")
SHEET_KEY = st.secrets["SHEET_KEY"]
STAFF_SHEET_NAME = st.secrets.get("STAFF_SHEET_NAME", "NhanSu")
LOG_SHEET_NAME = st.secrets.get("LOG_SHEET_NAME", "Log")
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

STAFF_COLUMNS = ["MSGV", "Họ và tên", "Đơn vị", "Bộ môn"]
LOG_COLUMNS = ["Ngày", "MSGV", "Họ và tên", "Đơn vị", "Bộ môn", "CS", "Ca", "IN/OUT", "Giờ", "Timestamp"]

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

st.set_page_config(page_title="Điểm danh giảng viên", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
<style>
html, body, .stApp, [class*="css"] {
    color: #000000 !important;
}
h1 {
    font-weight: 900 !important;
    line-height: 1.15 !important;
}
label, p, span, div {
    color: #000000 !important;
}
input {
    font-weight: 700 !important;
    color: #000000 !important;
}
.stButton > button, div[data-testid="stButton"] > button, button[kind="primary"] {
    font-weight: 900 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-align: center !important;
}
#MainMenu, footer, header,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
.stDeployButton {
    display: none !important;
    visibility: hidden !important;
}

/* Chỉ tối ưu mạnh cho màn hình điện thoại khi GV quét QR */
@media (max-width: 768px) {
    html, body, .stApp, .main, [data-testid="stAppViewContainer"] {
        max-width: 100vw !important;
        overflow-x: hidden !important;
    }
    .block-container {
        width: 100% !important;
        max-width: 100% !important;
        padding-top: 0.8rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        padding-bottom: 9rem !important;
        box-sizing: border-box !important;
    }
    h1 {
        font-size: 1.85rem !important;
        margin-bottom: 0.2rem !important;
    }
    input {
        font-size: 1.05rem !important;
        min-height: 3rem !important;
    }
    .stButton, .stButton > button, div[data-testid="stButton"], div[data-testid="stButton"] > button {
        width: 100% !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }
    .stButton > button, div[data-testid="stButton"] > button, button[kind="primary"] {
        min-height: 3.4rem !important;
        font-size: 1.05rem !important;
        margin: .25rem 0 1rem 0 !important;
    }
    [data-testid="stAlert"] {
        font-size: 1rem !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }
    div[role="radiogroup"] {
        display: flex !important;
        gap: 0.5rem !important;
        flex-wrap: wrap !important;
    }
}
</style>
""", unsafe_allow_html=True)


def now_vn():
    return datetime.datetime.now(VN_TZ)

def today_date():
    return now_vn().date()

def today_str():
    return today_date().strftime("%d/%m/%Y")

def time_str():
    return now_vn().strftime("%H:%M:%S")

def timestamp_str():
    return now_vn().strftime("%Y-%m-%d %H:%M:%S")

def infer_shift():
    return "Sáng" if now_vn().hour < 12 else "Chiều"

def get_query_params():
    if hasattr(st, "query_params"):
        return dict(st.query_params)
    raw = st.experimental_get_query_params()
    return {k: (v[0] if isinstance(v, list) and v else v) for k, v in raw.items()}

def strip_accents(s: str) -> str:
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", s)

def norm_search(s: str) -> str:
    return " ".join(strip_accents(s).lower().split())

def _norm_digits(value):
    return re.sub(r"\D", "", str(value or "").strip())

def _norm_header(value):
    return norm_search(str(value or "").strip()).replace(" ", "")

def normalize_sheet_name(name: str) -> str:
    return re.sub(r"\s+", "", str(name or "")).lower()

def _google_api_retry(callable_fn, retries=3, delay=1.2):
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
        or "https://giangvien.streamlit.app"
    )

def _get_admin_pw():
    return (
        st.secrets.get("ADMIN_PASSWORD")
        or st.secrets.get("teacher_password")
        or st.secrets.get("google_service_account", {}).get("teacher_password")
        or os.getenv("ADMIN_PASSWORD")
        or os.getenv("TEACHER_PASSWORD")
    )

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
                    st.warning("Sai mật khẩu hoặc chưa cấu hình ADMIN_PASSWORD.")

@st.cache_resource
def _get_gspread_client():
    if "google_service_account" not in st.secrets:
        raise RuntimeError("Thiếu block [google_service_account] trong Secrets.")

    cred = dict(st.secrets["google_service_account"])
    pk = cred.get("private_key", "")
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

def get_or_create_ws(ss, title, rows=1000, cols=20):
    worksheets = _google_api_retry(lambda: ss.worksheets())
    wanted = normalize_sheet_name(title)
    for ws in worksheets:
        if normalize_sheet_name(ws.title) == wanted:
            return ws
    return _google_api_retry(lambda: ss.add_worksheet(title=title, rows=rows, cols=cols))

def ensure_header(ws, headers):
    current = _google_api_retry(lambda: ws.row_values(1))
    if not current:
        _google_api_retry(lambda: ws.update("A1", [headers]))
        return headers
    merged = current[:]
    changed = False
    for h in headers:
        if h not in merged:
            merged.append(h)
            changed = True
    if changed:
        _google_api_retry(lambda: ws.update("1:1", [merged]))
    return _google_api_retry(lambda: ws.row_values(1))

def get_staff_ws():
    ss = get_spreadsheet()
    ws = get_or_create_ws(ss, STAFF_SHEET_NAME, rows=300, cols=10)
    ensure_header(ws, STAFF_COLUMNS)
    return ws

def get_log_ws():
    ss = get_spreadsheet()
    ws = get_or_create_ws(ss, LOG_SHEET_NAME, rows=5000, cols=12)
    ensure_header(ws, LOG_COLUMNS)
    return ws

def load_records(ws):
    return _google_api_retry(lambda: ws.get_all_records(expected_headers=None, default_blank=""))

def find_staff_by_msgv(msgv_full):
    ws = get_staff_ws()
    values = _google_api_retry(lambda: ws.get_all_values())
    if not values or len(values) < 2:
        return None

    headers = values[0]
    header_norms = [_norm_header(h) for h in headers]
    try:
        msgv_idx = header_norms.index("msgv")
    except ValueError:
        msgv_idx = 0

    def find_col(names, fallback):
        wanted = [_norm_header(x) for x in names]
        for i, h in enumerate(header_norms):
            if h in wanted:
                return i
        return fallback

    name_idx = find_col(["Họ và tên", "Ho va ten", "Họ tên", "Ho ten"], 1)
    unit_idx = find_col(["Đơn vị", "Don vi"], 2)
    dept_idx = find_col(["Bộ môn", "Bo mon"], 3)

    target_digits = _norm_digits(msgv_full)
    target_last4 = target_digits[-4:]

    for row in values[1:]:
        raw = row[msgv_idx] if msgv_idx < len(row) else ""
        raw_digits = _norm_digits(raw)
        if not raw_digits:
            continue

        raw_padded = raw_digits.zfill(len(target_digits))
        if raw_digits == target_digits or raw_padded == target_digits or raw_digits.endswith(target_last4):
            return {
                "MSGV": msgv_full,
                "Họ và tên": row[name_idx] if name_idx < len(row) else "",
                "Đơn vị": row[unit_idx] if unit_idx < len(row) else "",
                "Bộ môn": row[dept_idx] if dept_idx < len(row) else "",
            }
    return None

def load_log_records():
    return load_records(get_log_ws())

def get_today_logs_for_msgv(msgv_full):
    return [
        r for r in load_log_records()
        if str(r.get("Ngày", "")).strip() == today_str()
        and _norm_digits(r.get("MSGV", "")) == _norm_digits(msgv_full)
    ]

def next_action_for_msgv(msgv_full):
    shift = infer_shift()
    logs = get_today_logs_for_msgv(msgv_full)
    has_in = any(str(r.get("IN/OUT", "")).strip() == "IN" and str(r.get("Ca", "")).strip() == shift for r in logs)
    has_out = any(str(r.get("IN/OUT", "")).strip() == "OUT" and str(r.get("Ca", "")).strip() == shift for r in logs)
    if not has_in:
        return "IN", shift
    if not has_out:
        return "OUT", shift
    return None, shift

def append_log(row_dict):
    ws = get_log_ws()
    headers = ensure_header(ws, LOG_COLUMNS)
    row_values = [row_dict.get(c, "") for c in headers]
    _google_api_retry(lambda: ws.append_row(row_values, value_input_option="USER_ENTERED"))

def calculate_work_hours(logs):
    if not logs:
        return pd.DataFrame()
    df = pd.DataFrame(logs)
    if df.empty:
        return df
    for c in LOG_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    rows = []
    group_cols = ["Ngày", "MSGV", "Họ và tên", "Đơn vị", "Bộ môn", "Ca"]
    for keys, g in df.groupby(group_cols, dropna=False):
        ngay, msgv, hoten, donvi, bomon, ca = keys
        in_times = g[g["IN/OUT"] == "IN"]["Giờ"].tolist()
        out_times = g[g["IN/OUT"] == "OUT"]["Giờ"].tolist()
        cs_values = sorted(set([str(x) for x in g["CS"].tolist() if str(x).strip()]))
        in_time = min(in_times) if in_times else ""
        out_time = max(out_times) if out_times else ""
        hours = ""
        if in_time and out_time:
            try:
                dt_in = datetime.datetime.strptime(in_time, "%H:%M:%S")
                dt_out = datetime.datetime.strptime(out_time, "%H:%M:%S")
                delta = dt_out - dt_in
                hours = round(delta.total_seconds() / 3600, 2) if delta.total_seconds() >= 0 else ""
            except Exception:
                hours = ""
        rows.append({
            "Ngày": ngay, "MSGV": msgv, "Họ và tên": hoten, "Đơn vị": donvi,
            "Bộ môn": bomon, "Ca": ca, "Vào ca": in_time, "Ra ca": out_time,
            "Giờ có mặt": hours, "Cơ sở": ", ".join(cs_values),
        })
    return pd.DataFrame(rows)

def render_location_check(campus_code: str):
    campus_name = LOCATION_BY_CODE.get(campus_code)
    if not campus_name:
        st.error("Cơ sở điểm danh không hợp lệ.")
        st.stop()
    campus = LOCATIONS[campus_name]
    st.info(f"Cơ sở: {campus_name}")

    if streamlit_geolocation is None or geodesic is None:
        st.error("Ứng dụng chưa cài đủ thư viện kiểm tra vị trí.")
        st.stop()

    st.caption("Cho phép truy cập vị trí để xác thực điểm danh.")
    location = streamlit_geolocation()
    if not location:
        st.warning("Chưa nhận được vị trí. Vui lòng bật định vị và cho phép truy cập vị trí.")
        st.stop()

    lat, lon = location.get("latitude"), location.get("longitude")
    if lat is None or lon is None:
        st.warning("Không lấy được tọa độ GPS từ thiết bị. Vui lòng thử lại.")
        st.stop()

    distance = geodesic((float(lat), float(lon)), (campus["lat"], campus["lon"])).meters
    if distance > campus["radius"]:
        st.error(f"Bạn đang ngoài phạm vi điểm danh của {campus_name}.")
        st.stop()
    st.success("Vị trí hợp lệ. Có thể tiếp tục điểm danh.")

def render_gv_attendance():
    qp = get_query_params()
    campus_code = qp.get("coso", "CS1")
    campus_name = LOCATION_BY_CODE.get(campus_code)

    st.title("Điểm danh giảng viên")

    if today_date().weekday() == 6:
        st.error("Chủ nhật không mở điểm danh.")
        st.stop()

    if not campus_name:
        st.error("Cơ sở điểm danh không hợp lệ.")
        st.stop()

    st.info(f"Ngày: {today_str()}")
    render_location_check(campus_code)

    shift = infer_shift()
    st.info(f"Ca hiện tại: {shift}")

    action_label = st.radio(
        "Chọn loại điểm danh",
        ["Vào ca", "Ra ca"],
        horizontal=True,
        key="action_label",
    )
    action = "IN" if action_label == "Vào ca" else "OUT"

    msgv_suffix = st.text_input(
        "4 số cuối MSGV",
        placeholder="VD: 1234",
        max_chars=4,
        help=None,
        key="msgv_suffix",
    )

    if msgv_suffix.strip().isdigit() and len(msgv_suffix.strip()) == 4:
        msgv_full_preview = f"{MSGV_PREFIX}{msgv_suffix.strip().zfill(4)}"
        st.caption(f"MSGV: {msgv_full_preview}")

    if st.button("Xác nhận điểm danh", type="primary", use_container_width=True):
        if not msgv_suffix.strip().isdigit() or len(msgv_suffix.strip()) != 4:
            st.warning("Vui lòng nhập đúng 4 số cuối MSGV.")
            st.stop()

        msgv_full = f"{MSGV_PREFIX}{msgv_suffix.strip().zfill(4)}"

        try:
            staff = find_staff_by_msgv(msgv_full)
            if not staff:
                st.error(f"Không tìm thấy MSGV {msgv_full}.")
                st.stop()

            # Chặn ghi trùng: cùng ngày + cùng MSGV + cùng ca + cùng loại IN/OUT chỉ được ghi 1 lần.
            today_logs = get_today_logs_for_msgv(msgv_full)
            already_logged = any(
                str(r.get("Ca", "")).strip() == shift
                and str(r.get("IN/OUT", "")).strip() == action
                for r in today_logs
            )

            if already_logged:
                label = "vào ca" if action == "IN" else "ra ca"
                st.info(f"MSGV {msgv_full} đã điểm danh {label} {shift} hôm nay. Hệ thống không ghi trùng.")
                st.stop()

            # Nếu chọn ra ca nhưng chưa có vào ca trong cùng ca thì nhắc lại, không ghi.
            if action == "OUT":
                has_in = any(
                    str(r.get("Ca", "")).strip() == shift
                    and str(r.get("IN/OUT", "")).strip() == "IN"
                    for r in today_logs
                )
                if not has_in:
                    st.warning(f"Chưa có dữ liệu vào ca {shift}. Vui lòng điểm danh vào ca trước.")
                    st.stop()

            t = time_str()
            append_log({
                "Ngày": today_str(),
                "MSGV": msgv_full,
                "Họ và tên": staff.get("Họ và tên", ""),
                "Đơn vị": staff.get("Đơn vị", ""),
                "Bộ môn": staff.get("Bộ môn", ""),
                "CS": campus_code,
                "Ca": shift,
                "IN/OUT": action,
                "Giờ": t,
                "Timestamp": timestamp_str(),
            })

            st.success(
                f"{action_label} thành công!

"
                f"MSGV: {msgv_full}

"
                f"Ca: {shift}

"
                f"Giờ: {t}

"
                f"Cơ sở: {campus_code}"
            )

        except Exception as e:
            st.error(f"Lỗi khi điểm danh: {e}")


def render_tab_qr():
    st.subheader("Tạo QR cố định theo cơ sở")
    campus_name = st.selectbox("Chọn cơ sở", list(LOCATIONS.keys()))
    campus_code = LOCATIONS[campus_name]["code"]
    if st.button("Tạo QR cố định", type="primary", use_container_width=True):
        qr_data = f"{get_base_url()}/?gv=1&coso={urllib.parse.quote(campus_code)}"
        qr = qrcode.make(qr_data)
        buf = io.BytesIO()
        qr.save(buf, format="PNG")
        buf.seek(0)
        st.image(Image.open(buf), caption=f"QR cố định cho {campus_code}", width=380)
        st.code(qr_data)

def find_staff_candidates(records, query: str):
    q = (query or "").strip()
    if not q:
        return []
    if q.isdigit() and len(q) == 4:
        return [r for r in records if _norm_digits(r.get("MSGV", "")).endswith(q)]
    qn = norm_search(q)
    return [r for r in records if qn in norm_search(r.get("Họ và tên", ""))]

def render_tab_search():
    st.subheader("Tìm kiếm giảng viên")
    q = st.text_input("Nhập 4 số cuối MSGV hoặc họ tên")
    if st.button("Tìm", use_container_width=True):
        records = load_records(get_staff_ws())
        results = find_staff_candidates(records, q)
        if results:
            st.dataframe(pd.DataFrame(results), use_container_width=True)
        else:
            st.warning("Không tìm thấy kết quả phù hợp.")

def render_tab_stats():
    st.subheader("Thống kê log điểm danh")
    logs = load_log_records()
    if not logs:
        st.info("Chưa có dữ liệu điểm danh.")
        return
    df = pd.DataFrame(logs)
    for c in LOG_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    date_values = sorted([x for x in df["Ngày"].dropna().astype(str).unique() if x])
    selected_date = st.selectbox("Chọn ngày", date_values, index=len(date_values)-1 if date_values else 0)
    filtered = df[df["Ngày"].astype(str) == selected_date].copy()
    st.metric("Số lượt log", len(filtered))
    st.dataframe(filtered, use_container_width=True)
    summary = calculate_work_hours(filtered.to_dict("records"))
    if not summary.empty:
        st.subheader("Tổng hợp giờ có mặt")
        st.dataframe(summary, use_container_width=True)
        campus_df = filtered.groupby("CS", dropna=False).size().reset_index(name="Số lượt")
        chart = alt.Chart(campus_df).mark_bar().encode(
            x=alt.X("CS:N", title="Cơ sở"),
            y=alt.Y("Số lượt:Q", title="Số lượt"),
            tooltip=["CS", "Số lượt"],
        ).properties(height=320)
        st.altair_chart(chart, use_container_width=True)

def render_tab_setup():
    st.subheader("Kiểm tra cấu trúc dữ liệu")
    staff_ws = get_staff_ws()
    log_ws = get_log_ws()
    ensure_header(log_ws, LOG_COLUMNS)
    st.success("Đã kiểm tra xong.")
    st.write("Sheet danh sách:", staff_ws.title)
    st.write("Sheet log:", log_ws.title)
    st.write("Cột log:", LOG_COLUMNS)

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
    render_tab_setup()
