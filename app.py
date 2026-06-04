# app.py
import os
import io
import re
import time
import base64
import urllib.parse
import unicodedata
import datetime

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


# ===================== CẤU HÌNH =====================
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

# ===================== CSS =====================
st.markdown("""
<style>
html, body, .stApp {
    color: #000000 !important;
    overflow-x: hidden !important;
}
#MainMenu, footer, header,
[data-testid="stToolbar"],
[data-testid="stDecoration"],
[data-testid="stStatusWidget"],
.stDeployButton {
    display: none !important;
    visibility: hidden !important;
}
h1, h2, h3 {
    font-weight: 900 !important;
}
label, p, span, div {
    color: #000000 !important;
}
input {
    font-weight: 700 !important;
    color: #000000 !important;
}
.stButton > button, div[data-testid="stButton"] > button {
    font-weight: 900 !important;
    white-space: normal !important;
    overflow: visible !important;
    text-align: center !important;
}
@media (max-width: 768px) {
    .block-container {
        width: 100% !important;
        max-width: 100% !important;
        padding-top: 0.8rem !important;
        padding-left: 0.85rem !important;
        padding-right: 0.85rem !important;
        padding-bottom: 8rem !important;
        box-sizing: border-box !important;
    }
    h1 {
        font-size: 1.8rem !important;
        line-height: 1.15 !important;
        margin-bottom: 0.25rem !important;
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
    .stButton > button, div[data-testid="stButton"] > button {
        min-height: 3.4rem !important;
        font-size: 1.05rem !important;
        margin: .25rem 0 1rem 0 !important;
    }
    div[role="radiogroup"] {
        display: flex !important;
        gap: 1.2rem !important;
        flex-wrap: wrap !important;
    }
    [data-testid="stAlert"] {
        font-size: 1rem !important;
        max-width: 100% !important;
        box-sizing: border-box !important;
    }
}
</style>
""", unsafe_allow_html=True)


# ===================== TIỆN ÍCH =====================
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
    s = unicodedata.normalize("NFD", str(s or ""))
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", s)


def norm_search(s: str) -> str:
    return " ".join(strip_accents(s).lower().split())


def norm_header(s: str) -> str:
    return norm_search(s).replace(" ", "")


def norm_sheet_name(name: str) -> str:
    return re.sub(r"\s+", "", str(name or "")).lower()


def norm_digits(value) -> str:
    return re.sub(r"\D", "", str(value or ""))


def safe_str(value) -> str:
    return str(value or "").strip()


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


# ===================== ĐĂNG NHẬP QUẢN TRỊ =====================
def get_admin_pw():
    return (
        st.secrets.get("ADMIN_PASSWORD")
        or st.secrets.get("teacher_password")
        or st.secrets.get("google_service_account", {}).get("teacher_password")
        or os.getenv("ADMIN_PASSWORD")
        or os.getenv("TEACHER_PASSWORD")
    )


def admin_unlocked():
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
            pw = st.text_input("Mật khẩu quản trị", type="password")
            if st.button("Đăng nhập", type="primary", use_container_width=True):
                if get_admin_pw() and pw == get_admin_pw():
                    st.session_state["admin_unlocked"] = True
                    st.rerun()
                else:
                    st.warning("Sai mật khẩu hoặc chưa cấu hình ADMIN_PASSWORD.")


# ===================== GOOGLE SHEETS =====================
@st.cache_resource
def get_gspread_client():
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
    body_raw = re.sub(r"[^A-Za-z0-9+/=]", "", "".join(lines[h_idx + 1:f_idx]))
    body = body_raw.replace("=", "")
    rem = len(body) % 4
    if rem:
        body += "=" * (4 - rem)
    base64.b64decode(body, validate=True)

    cred["private_key"] = header + "\n" + "\n".join(body[i:i+64] for i in range(0, len(body), 64)) + "\n" + footer + "\n"
    creds = Credentials.from_service_account_info(cred, scopes=SCOPES)
    return gspread.authorize(creds)


def get_spreadsheet():
    return _google_api_retry(lambda: get_gspread_client().open_by_key(SHEET_KEY))


def get_or_create_ws(title, rows=1000, cols=20):
    ss = get_spreadsheet()
    wanted = norm_sheet_name(title)
    worksheets = _google_api_retry(lambda: ss.worksheets())
    for ws in worksheets:
        if norm_sheet_name(ws.title) == wanted:
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


def staff_ws():
    ws = get_or_create_ws(STAFF_SHEET_NAME, rows=300, cols=10)
    ensure_header(ws, STAFF_COLUMNS)
    return ws


def log_ws():
    ws = get_or_create_ws(LOG_SHEET_NAME, rows=5000, cols=12)
    ensure_header(ws, LOG_COLUMNS)
    return ws


def get_all_records_by_header(ws):
    values = _google_api_retry(lambda: ws.get_all_values())
    if not values:
        return []
    headers = values[0]
    out = []
    for row in values[1:]:
        item = {}
        for i, h in enumerate(headers):
            item[h] = row[i] if i < len(row) else ""
        if any(str(v).strip() for v in item.values()):
            out.append(item)
    return out


def find_staff_by_msgv(msgv_full):
    ws = staff_ws()
    values = _google_api_retry(lambda: ws.get_all_values())
    if not values or len(values) < 2:
        return None

    headers = values[0]
    hn = [norm_header(h) for h in headers]

    def col_index(names, default):
        wanted = [norm_header(x) for x in names]
        for i, h in enumerate(hn):
            if h in wanted:
                return i
        return default

    msgv_i = col_index(["MSGV"], 0)
    name_i = col_index(["Họ và tên", "Ho va ten", "Họ tên", "Ho ten"], 1)
    unit_i = col_index(["Đơn vị", "Don vi"], 2)
    dept_i = col_index(["Bộ môn", "Bo mon"], 3)

    target_full = norm_digits(msgv_full)
    target_last4 = target_full[-4:]

    for row in values[1:]:
        raw = row[msgv_i] if msgv_i < len(row) else ""
        raw_digits = norm_digits(raw)
        if not raw_digits:
            continue

        raw_padded = raw_digits.zfill(len(target_full))
        if raw_digits == target_full or raw_padded == target_full or raw_digits.endswith(target_last4):
            return {
                "MSGV": msgv_full,
                "Họ và tên": row[name_i] if name_i < len(row) else "",
                "Đơn vị": row[unit_i] if unit_i < len(row) else "",
                "Bộ môn": row[dept_i] if dept_i < len(row) else "",
            }

    return None


def load_logs():
    return get_all_records_by_header(log_ws())


def logs_for_msgv_today(msgv_full):
    target = norm_digits(msgv_full)
    result = []
    for r in load_logs():
        if safe_str(r.get("Ngày")) == today_str() and norm_digits(r.get("MSGV")) == target:
            result.append(r)
    return result


def append_log(row):
    ws = log_ws()
    headers = ensure_header(ws, LOG_COLUMNS)
    values = [row.get(h, "") for h in headers]
    _google_api_retry(lambda: ws.append_row(values, value_input_option="USER_ENTERED"))


def summarize_hours(records):
    if not records:
        return pd.DataFrame()
    df = pd.DataFrame(records)
    for c in LOG_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    rows = []
    for keys, g in df.groupby(["Ngày", "MSGV", "Họ và tên", "Đơn vị", "Bộ môn", "Ca"], dropna=False):
        ngay, msgv, hoten, donvi, bomon, ca = keys
        ins = g[g["IN/OUT"] == "IN"]["Giờ"].tolist()
        outs = g[g["IN/OUT"] == "OUT"]["Giờ"].tolist()
        vao = min(ins) if ins else ""
        ra = max(outs) if outs else ""
        hours = ""
        if vao and ra:
            try:
                d1 = datetime.datetime.strptime(vao, "%H:%M:%S")
                d2 = datetime.datetime.strptime(ra, "%H:%M:%S")
                sec = (d2 - d1).total_seconds()
                hours = round(sec / 3600, 2) if sec >= 0 else ""
            except Exception:
                hours = ""
        rows.append({
            "Ngày": ngay, "MSGV": msgv, "Họ và tên": hoten, "Đơn vị": donvi,
            "Bộ môn": bomon, "Ca": ca, "Vào ca": vao, "Ra ca": ra,
            "Giờ có mặt": hours, "Cơ sở": ", ".join(sorted(set(g["CS"].astype(str)))),
        })
    return pd.DataFrame(rows)


# ===================== GPS =====================
def render_location_check(campus_code):
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
    loc = streamlit_geolocation()
    if not loc:
        st.warning("Chưa nhận được vị trí. Vui lòng bật định vị và cho phép truy cập vị trí.")
        st.stop()

    lat, lon = loc.get("latitude"), loc.get("longitude")
    if lat is None or lon is None:
        st.warning("Không lấy được tọa độ GPS từ thiết bị. Vui lòng thử lại.")
        st.stop()

    distance = geodesic((float(lat), float(lon)), (campus["lat"], campus["lon"])).meters
    if distance > campus["radius"]:
        st.error(f"Bạn đang ngoài phạm vi điểm danh của {campus_name}.")
        st.stop()

    st.success("Vị trí hợp lệ. Có thể tiếp tục điểm danh.")


# ===================== GIAO DIỆN GV =====================
def render_gv_attendance():
    qp = get_query_params()
    campus_code = qp.get("coso", "CS1")

    st.title("Điểm danh giảng viên")

    if today_date().weekday() == 6:
        st.error("Chủ nhật không mở điểm danh.")
        st.stop()

    st.info(f"Ngày: {today_str()}")
    render_location_check(campus_code)

    shift = infer_shift()
    st.info(f"Ca hiện tại: {shift}")

    action_label = st.radio("Chọn loại điểm danh", ["Vào ca", "Ra ca"], horizontal=True)
    action = "IN" if action_label == "Vào ca" else "OUT"

    suffix = st.text_input("4 số cuối MSGV", placeholder="VD: 1234", max_chars=4)
    if suffix.strip().isdigit() and len(suffix.strip()) == 4:
        st.caption(f"MSGV: {MSGV_PREFIX}{suffix.strip().zfill(4)}")

    if st.button("Xác nhận điểm danh", type="primary", use_container_width=True):
        if not suffix.strip().isdigit() or len(suffix.strip()) != 4:
            st.warning("Vui lòng nhập đúng 4 số cuối MSGV.")
            st.stop()

        msgv_full = f"{MSGV_PREFIX}{suffix.strip().zfill(4)}"
        staff = find_staff_by_msgv(msgv_full)

        if not staff:
            st.error(f"Không tìm thấy MSGV {msgv_full}.")
            st.stop()

        current_logs = logs_for_msgv_today(msgv_full)

        already = any(
            safe_str(r.get("Ca")) == shift and safe_str(r.get("IN/OUT")) == action
            for r in current_logs
        )

        if already:
            label = "vào ca" if action == "IN" else "ra ca"
            st.info(f"MSGV {msgv_full} đã điểm danh {label} {shift} hôm nay. Hệ thống không ghi trùng.")
            st.stop()

        if action == "OUT":
            has_in = any(
                safe_str(r.get("Ca")) == shift and safe_str(r.get("IN/OUT")) == "IN"
                for r in current_logs
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

        st.success(f"{action_label} thành công!")
        st.write(f"MSGV: **{msgv_full}**")
        st.write(f"Ca: **{shift}**")
        st.write(f"Giờ: **{t}**")
        st.write(f"Cơ sở: **{campus_code}**")


# ===================== GIAO DIỆN QUẢN TRỊ =====================
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


def render_tab_search():
    st.subheader("Tìm kiếm giảng viên")
    q = st.text_input("Nhập 4 số cuối MSGV hoặc họ tên")
    if st.button("Tìm", use_container_width=True):
        rows = get_all_records_by_header(staff_ws())
        if q.isdigit() and len(q) == 4:
            rows = [r for r in rows if norm_digits(r.get("MSGV")).endswith(q)]
        else:
            rows = [r for r in rows if norm_search(q) in norm_search(r.get("Họ và tên"))]
        if rows:
            st.dataframe(pd.DataFrame(rows), use_container_width=True)
        else:
            st.warning("Không tìm thấy kết quả phù hợp.")


def render_tab_stats():
    st.subheader("Thống kê điểm danh")
    logs = load_logs()
    if not logs:
        st.info("Chưa có dữ liệu điểm danh.")
        return
    df = pd.DataFrame(logs)
    for c in LOG_COLUMNS:
        if c not in df.columns:
            df[c] = ""

    dates = sorted([x for x in df["Ngày"].dropna().astype(str).unique() if x])
    selected = st.selectbox("Chọn ngày", dates, index=len(dates)-1)
    filtered = df[df["Ngày"].astype(str) == selected].copy()
    st.metric("Số lượt log", len(filtered))
    st.dataframe(filtered, use_container_width=True)

    summary = summarize_hours(filtered.to_dict("records"))
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
    sw = staff_ws()
    lw = log_ws()
    ensure_header(lw, LOG_COLUMNS)
    st.success("Đã kiểm tra xong.")
    st.write("Sheet danh sách:", sw.title)
    st.write("Sheet log:", lw.title)
    st.write("Cột log:", LOG_COLUMNS)


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
        ["Tạo QR cố định", "Tìm kiếm giảng viên", "Thống kê điểm danh", "Cấu trúc dữ liệu"],
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
