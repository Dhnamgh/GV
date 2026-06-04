# app_gv_attendance_log.py
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

# ===================== CONFIG =====================
MSGV_PREFIX = st.secrets.get("SESSION_PREFIX", "0607")
SHEET_KEY = st.secrets["SHEET_KEY"]
STAFF_SHEET_NAME = st.secrets.get("STAFF_SHEET_NAME", "NhanSu")
LOG_SHEET_NAME = st.secrets.get("LOG_SHEET_NAME", "Log")
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
STAFF_COLUMNS = ["MSGV", "Họ và tên", "Đơn vị", "Bộ môn"]
LOG_COLUMNS = ["Ngày", "MSGV", "Họ và tên", "Đơn vị", "Bộ môn", "CS", "Ca", "IN/OUT", "Giờ", "Timestamp"]

LOCATIONS = {
    "Cơ sở 1 - Hồng Bàng": {"code": "CS1", "lat": 10.754665, "lon": 106.663381, "radius": 500, "address": "217 Hồng Bàng, Phường Chợ Lớn, TP.HCM"},
    "Cơ sở 2 - Đinh Tiên Hoàng": {"code": "CS2", "lat": 10.785434, "lon": 106.702667, "radius": 400, "address": "43 Đinh Tiên Hoàng, Phường Sài Gòn, TP.HCM"},
}
LOCATION_BY_CODE = {v["code"]: k for k, v in LOCATIONS.items()}

st.set_page_config(page_title="Điểm danh giảng viên QR", layout="wide")

st.markdown("""
<style>
html, body, .stApp, [class*="css"] {font-size:17px!important;color:#000!important;}
h1 {font-size:32px!important;font-weight:900!important;color:#000!important;}
h2,h3 {font-size:24px!important;font-weight:900!important;color:#000!important;}
p,span,div,label {color:#000!important;}
label {font-size:17px!important;font-weight:800!important;}
input,textarea {font-size:17px!important;font-weight:700!important;color:#000!important;min-height:44px!important;}
.stButton, div[data-testid="stButton"] {width:100%!important;max-width:100%!important;}
.stButton button, div[data-testid="stButton"] > button, button[kind="primary"] {
  width:100%!important;
  max-width:100%!important;
  min-height:56px!important;
  font-size:18px!important;
  font-weight:900!important;
  white-space:normal!important;
  overflow:visible!important;
  display:block!important;
  text-align:center!important;
}
[data-testid="stMetricLabel"] {font-size:17px!important;font-weight:900!important;}
[data-testid="stMetricValue"] {font-size:36px!important;font-weight:900!important;}
[data-testid="stCaptionContainer"], small {font-size:14px!important;font-weight:700!important;}
[data-testid="stAlert"] {font-size:17px!important;font-weight:800!important;}
section[data-testid="stSidebar"] * {font-size:16px!important;font-weight:700!important;}
.block-container {padding-bottom:8rem!important;max-width:100vw!important;}
#MainMenu, footer, header, [data-testid="stToolbar"], [data-testid="stDecoration"], [data-testid="stStatusWidget"], .stDeployButton {display:none!important;visibility:hidden!important;}
@media(max-width:768px){
  html,body,.stApp,[class*="css"]{font-size:15px!important;}
  h1{font-size:25px!important;line-height:1.2!important;margin-bottom:.4rem!important;}
  h2,h3{font-size:20px!important;}
  .block-container{padding:1rem 1rem 8rem 1rem!important;max-width:100vw!important;}
}
</style>
""", unsafe_allow_html=True)

# ===================== HELPERS =====================
def now_vn(): return datetime.datetime.now(VN_TZ)
def today_str(): return now_vn().strftime("%d/%m/%Y")
def time_str(): return now_vn().strftime("%H:%M:%S")
def timestamp_str(): return now_vn().strftime("%Y-%m-%d %H:%M:%S")
def infer_shift(): return "Sáng" if now_vn().hour < 12 else "Chiều"

def get_query_params():
    if hasattr(st, "query_params"):
        return dict(st.query_params)
    raw = st.experimental_get_query_params()
    return {k: (v[0] if isinstance(v, list) and v else v) for k, v in raw.items()}

def normalize_name(name): return " ".join(w.capitalize() for w in (name or "").strip().split())
def strip_accents(s):
    s = unicodedata.normalize("NFD", s or "")
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return unicodedata.normalize("NFC", s)
def norm_search(s): return " ".join(strip_accents(s).lower().split())

def _google_api_retry(fn, retries=3, delay=1.5):
    last = None
    for i in range(retries):
        try: return fn()
        except Exception as e:
            last = e
            msg = str(e)
            transient = any(x in msg for x in ["[500]", "[503]", "[429]", "Internal error", "Quota", "timeout", "Timeout"])
            if not transient or i == retries - 1: raise
            time.sleep(delay * (i + 1))
    raise last

def get_base_url():
    return st.secrets.get("WRAPPER_URL") or st.secrets.get("APP_BASE_URL") or st.secrets.get("google_service_account", {}).get("app_base_url") or "https://giangvien.streamlit.app"

# ===================== AUTH =====================
def _get_admin_pw():
    if "ADMIN_PASSWORD" in st.secrets: return st.secrets["ADMIN_PASSWORD"]
    if "teacher_password" in st.secrets: return st.secrets["teacher_password"]
    if "google_service_account" in st.secrets:
        maybe = st.secrets["google_service_account"].get("teacher_password")
        if maybe: return maybe
    return os.getenv("ADMIN_PASSWORD") or os.getenv("TEACHER_PASSWORD")

def admin_unlocked(): return bool(st.session_state.get("admin_unlocked"))

def render_admin_auth():
    with st.sidebar:
        st.header("Quản trị")
        if admin_unlocked():
            st.success("Đã đăng nhập quản trị")
            if st.button("Đăng xuất"):
                st.session_state.clear(); st.rerun()
        else:
            pw = st.text_input("Mật khẩu quản trị", type="password", key="pw_admin")
            if st.button("Đăng nhập", type="primary", use_container_width=True):
                if _get_admin_pw() and pw == _get_admin_pw():
                    st.session_state["admin_unlocked"] = True; st.rerun()
                else:
                    st.warning("Sai mật khẩu hoặc chưa cấu hình ADMIN_PASSWORD trong Secrets/ENV.")

# ===================== GOOGLE SHEETS =====================
@st.cache_resource
def _get_gspread_client():
    if "google_service_account" not in st.secrets: raise RuntimeError("Thiếu block [google_service_account] trong Secrets.")
    cred = dict(st.secrets["google_service_account"])
    pk = cred.get("private_key", "")
    if not pk: raise RuntimeError("Secrets thiếu private_key.")
    if "\\n" in pk: pk = pk.replace("\\n", "\n")
    pk = pk.replace("\r\n", "\n").replace("\r", "\n")
    header, footer = "-----BEGIN PRIVATE KEY-----", "-----END PRIVATE KEY-----"
    if header not in pk or footer not in pk: raise RuntimeError("private_key thiếu BEGIN/END.")
    lines = [ln.strip() for ln in pk.split("\n")]
    h_idx, f_idx = lines.index(header), lines.index(footer)
    body_raw = re.sub(r"[^A-Za-z0-9+/=]", "", "".join([ln for ln in lines[h_idx+1:f_idx] if ln]))
    body = body_raw.replace("=", "")
    if not body: raise RuntimeError("private_key rỗng sau khi làm sạch.")
    rem = len(body) % 4
    if rem: body += "=" * (4 - rem)
    base64.b64decode(body, validate=True)
    cred["private_key"] = header + "\n" + "\n".join(body[i:i+64] for i in range(0, len(body), 64)) + "\n" + footer + "\n"
    creds = Credentials.from_service_account_info(cred, scopes=SCOPES)
    return gspread.authorize(creds)

def get_spreadsheet(): return _google_api_retry(lambda: _get_gspread_client().open_by_key(SHEET_KEY))
def _norm_ws_title(name):
    return re.sub(r"\s+", "", str(name or "")).lower()

def get_or_create_ws(ss, title, rows=1000, cols=20):
    # Dò danh sách worksheet trước để tránh lỗi: A sheet with the name already exists.
    wanted = _norm_ws_title(title)
    worksheets = _google_api_retry(lambda: ss.worksheets())
    for ws in worksheets:
        if _norm_ws_title(ws.title) == wanted:
            return ws

    try:
        return _google_api_retry(lambda: ss.add_worksheet(title=title, rows=rows, cols=cols))
    except Exception as e:
        msg = str(e).lower()
        if "already exists" in msg or "sheet with the name" in msg:
            worksheets = _google_api_retry(lambda: ss.worksheets())
            for ws in worksheets:
                if _norm_ws_title(ws.title) == wanted:
                    return ws
        raise

def ensure_header(ws, headers):
    cur = _google_api_retry(lambda: ws.row_values(1))
    if not cur:
        _google_api_retry(lambda: ws.update("A1", [headers])); return headers
    merged, changed = cur[:], False
    for h in headers:
        if h not in merged:
            merged.append(h); changed = True
    if changed: _google_api_retry(lambda: ws.update("1:1", [merged]))
    return _google_api_retry(lambda: ws.row_values(1))

def get_staff_ws():
    ws = get_or_create_ws(get_spreadsheet(), STAFF_SHEET_NAME, rows=300, cols=10)
    ensure_header(ws, STAFF_COLUMNS)
    return ws

def get_log_ws():
    ws = get_or_create_ws(get_spreadsheet(), LOG_SHEET_NAME, rows=5000, cols=12)
    ensure_header(ws, LOG_COLUMNS)
    return ws

def load_records(ws): return _google_api_retry(lambda: ws.get_all_records(expected_headers=None, default_blank=""))

def find_staff_by_msgv(msgv):
    for r in load_records(get_staff_ws()):
        if str(r.get("MSGV", "")).strip() == msgv: return r
    return None

def load_log_records(): return load_records(get_log_ws())

def today_logs(msgv):
    return [r for r in load_log_records() if str(r.get("Ngày", "")).strip() == today_str() and str(r.get("MSGV", "")).strip() == msgv]

def next_action(msgv):
    shift = infer_shift()
    logs = today_logs(msgv)
    has_in = any(str(r.get("IN/OUT", "")).strip() == "IN" and str(r.get("Ca", "")).strip() == shift for r in logs)
    has_out = any(str(r.get("IN/OUT", "")).strip() == "OUT" and str(r.get("Ca", "")).strip() == shift for r in logs)
    if not has_in: return "IN", shift
    if not has_out: return "OUT", shift
    return None, shift

# ===================== GPS =====================
def render_location_check(campus_code):
    campus_name = LOCATION_BY_CODE.get(campus_code)
    if not campus_name: st.error("Cơ sở điểm danh không hợp lệ."); st.stop()
    campus = LOCATIONS[campus_name]
    st.info(f"Cơ sở: {campus_name}")
    if streamlit_geolocation is None or geodesic is None:
        st.error("Ứng dụng chưa cài đủ thư viện kiểm tra vị trí."); st.code("pip install streamlit-geolocation geopy"); st.stop()
    st.caption("Cho phép truy cập vị trí để xác thực điểm danh.")
    loc = streamlit_geolocation()
    if not loc: st.warning("Chưa nhận được vị trí. Vui lòng bật định vị và cho phép truy cập vị trí."); st.stop()
    lat, lon = loc.get("latitude"), loc.get("longitude")
    if lat is None or lon is None: st.warning("Không lấy được tọa độ GPS từ thiết bị. Vui lòng thử lại."); st.stop()
    dist = geodesic((float(lat), float(lon)), (campus["lat"], campus["lon"])).meters
    if dist > campus["radius"]: st.error(f"Bạn đang ngoài phạm vi điểm danh của {campus_name}."); st.stop()
    st.success("Vị trí hợp lệ. Có thể tiếp tục điểm danh.")

# ===================== GV UI =====================
def render_gv_attendance():
    qp = get_query_params(); campus_code = qp.get("coso", "CS1")
    st.title("Điểm danh giảng viên")
    if now_vn().weekday() == 6: st.error("Chủ nhật không mở điểm danh."); st.stop()
    if campus_code not in LOCATION_BY_CODE: st.error("Cơ sở điểm danh không hợp lệ."); st.stop()
    st.info(f"Ngày: {today_str()}")
    render_location_check(campus_code)
    suffix = st.text_input("4 số cuối MSGV", placeholder="VD: 1234", max_chars=4, help=None)
    if suffix.strip().isdigit() and len(suffix.strip()) == 4:
        st.caption(f"MSGV: {MSGV_PREFIX}{suffix.strip().zfill(4)}")
    if st.button("Xác nhận điểm danh", type="primary", use_container_width=True):
        if not suffix.strip().isdigit() or len(suffix.strip()) != 4:
            st.warning("Vui lòng nhập đúng 4 số cuối MSGV."); st.stop()
        msgv = f"{MSGV_PREFIX}{suffix.strip().zfill(4)}"
        try:
            staff = find_staff_by_msgv(msgv)
            if not staff: st.error(f"Không tìm thấy MSGV {msgv}."); st.stop()
            action, shift = next_action(msgv)
            if action is None:
                st.info(f"MSGV {msgv} đã đủ vào/ra ca {shift} hôm nay."); st.stop()
            row = {
                "Ngày": today_str(), "MSGV": msgv, "Họ và tên": staff.get("Họ và tên", ""),
                "Đơn vị": staff.get("Đơn vị", ""), "Bộ môn": staff.get("Bộ môn", ""),
                "CS": campus_code, "Ca": shift, "IN/OUT": action, "Giờ": time_str(), "Timestamp": timestamp_str()
            }
            ws = get_log_ws(); headers = ensure_header(ws, LOG_COLUMNS)
            _google_api_retry(lambda: ws.append_row([row.get(c, "") for c in headers], value_input_option="USER_ENTERED"))
            label = "Vào ca" if action == "IN" else "Ra ca"
            st.success(f"{label} thành công!\n\nMSGV: {msgv}\n\nCa: {shift}\n\nGiờ: {row['Giờ']}\n\nCơ sở: {campus_code}")
        except Exception as e: st.error(f"Lỗi khi điểm danh: {e}")

# ===================== ADMIN UI =====================
def render_tab_qr():
    st.subheader("Tạo QR cố định theo cơ sở")
    campus_name = st.selectbox("Chọn cơ sở", list(LOCATIONS.keys()))
    campus_code = LOCATIONS[campus_name]["code"]
    st.info(f"Cơ sở: {campus_name} - {LOCATIONS[campus_name]['address']}")
    if st.button("Tạo QR cố định", type="primary", use_container_width=True):
        qr_data = f"{get_base_url()}/?gv=1&coso={urllib.parse.quote(campus_code)}"
        qr = qrcode.make(qr_data); buf = io.BytesIO(); qr.save(buf, format="PNG"); buf.seek(0)
        st.image(Image.open(buf), caption=f"QR cố định cho {campus_code}", width=380)
        with st.expander("Xem link QR"): st.code(qr_data)

def find_staff_candidates(records, q):
    q = (q or "").strip()
    if not q: return []
    if q.isdigit() and len(q) == 4: return [r for r in records if str(r.get("MSGV", "")).strip().endswith(q)]
    qn = norm_search(q); res = [r for r in records if qn in norm_search(r.get("Họ và tên", ""))]
    if res: return res
    names = [r.get("Họ và tên", "") for r in records]; mp = {n: r for n, r in zip(names, records)}
    return [mp[n] for n in get_close_matches(q, names, n=5, cutoff=0.6) if n in mp]

def render_tab_search():
    st.subheader("Tìm kiếm giảng viên")
    q = st.text_input("Nhập 4 số cuối MSGV hoặc họ tên")
    if st.button("Tìm", use_container_width=True):
        try:
            res = find_staff_candidates(load_records(get_staff_ws()), q)
            if not res: st.warning("Không tìm thấy kết quả phù hợp."); return
            st.dataframe(pd.DataFrame(res), use_container_width=True)
        except Exception as e: st.error(f"Lỗi khi tìm kiếm: {e}")

def calculate_work_hours(logs):
    if not logs: return pd.DataFrame()
    df = pd.DataFrame(logs)
    for c in LOG_COLUMNS:
        if c not in df.columns: df[c] = ""
    rows = []
    for keys, g in df.groupby(["Ngày", "MSGV", "Họ và tên", "Đơn vị", "Bộ môn", "Ca"], dropna=False):
        ngay, msgv, hoten, donvi, bomon, ca = keys
        ins = g[g["IN/OUT"] == "IN"]["Giờ"].tolist(); outs = g[g["IN/OUT"] == "OUT"]["Giờ"].tolist()
        in_time = min(ins) if ins else ""; out_time = max(outs) if outs else ""; hours = ""
        if in_time and out_time:
            try:
                delta = datetime.datetime.strptime(out_time, "%H:%M:%S") - datetime.datetime.strptime(in_time, "%H:%M:%S")
                hours = round(delta.total_seconds()/3600, 2) if delta.total_seconds() >= 0 else ""
            except Exception: pass
        rows.append({"Ngày": ngay, "MSGV": msgv, "Họ và tên": hoten, "Đơn vị": donvi, "Bộ môn": bomon, "Ca": ca, "Vào ca": in_time, "Ra ca": out_time, "Giờ có mặt": hours, "Cơ sở": ", ".join(sorted(set(g["CS"].astype(str).tolist())))})
    return pd.DataFrame(rows)

def render_tab_stats():
    st.subheader("Thống kê log điểm danh")
    try:
        logs = load_log_records()
        if not logs: st.info("Chưa có dữ liệu điểm danh."); return
        df = pd.DataFrame(logs)
        dates = sorted([x for x in df["Ngày"].dropna().astype(str).unique() if x])
        selected = st.selectbox("Chọn ngày", dates, index=len(dates)-1)
        filtered = df[df["Ngày"].astype(str) == selected].copy()
        st.metric("Số lượt log", len(filtered)); st.dataframe(filtered, use_container_width=True)
        summary = calculate_work_hours(filtered.to_dict("records"))
        if not summary.empty:
            st.subheader("Tổng hợp giờ có mặt"); st.dataframe(summary, use_container_width=True)
            campus_df = filtered.groupby("CS", dropna=False).size().reset_index(name="Số lượt")
            st.altair_chart(alt.Chart(campus_df).mark_bar().encode(x="CS:N", y="Số lượt:Q", tooltip=["CS", "Số lượt"]).properties(height=320), use_container_width=True)
    except Exception as e: st.error(f"Lỗi thống kê: {e}")

def render_tab_setup():
    st.subheader("Kiểm tra cấu trúc dữ liệu")
    try:
        get_staff_ws(); get_log_ws(); st.success("Đã kiểm tra xong.")
        st.write("Sheet danh sách:", STAFF_SHEET_NAME); st.write("Sheet log:", LOG_SHEET_NAME)
        st.write("Cột danh sách:", STAFF_COLUMNS); st.write("Cột log:", LOG_COLUMNS)
    except Exception as e: st.error(f"Lỗi cấu trúc dữ liệu: {e}")

# ===================== ROUTING =====================
if get_query_params().get("gv") == "1":
    render_gv_attendance(); st.stop()

render_admin_auth(); st.title("Hệ thống điểm danh QR cho giảng viên")
if not admin_unlocked():
    st.error("Vui lòng đăng nhập quản trị để sử dụng các chức năng quản lý."); st.stop()
with st.sidebar:
    st.markdown("**Điều hướng**")
    menu = st.radio("Chọn mục", ["Tạo QR cố định", "Tìm kiếm giảng viên", "Thống kê điểm danh", "Cấu trúc dữ liệu"], index=0, label_visibility="collapsed")
if menu == "Tạo QR cố định": render_tab_qr()
elif menu == "Tìm kiếm giảng viên": render_tab_search()
elif menu == "Thống kê điểm danh": render_tab_stats()
else: render_tab_setup()
