# main.py
import os
import io
import re
import time
import base64
import urllib.parse
import unicodedata
import datetime
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
QR_SLOT_SECONDS = 30          # đổi 1 chỗ cho toàn app (30 giây là khuyến nghị)
UNLOCK_TTL = 120              # ân hạn phiên SV sau khi mở form (giây)
MSSV_PREFIX = st.secrets.get("SESSION_PREFIX", "51125")  # đọc từ Secrets, không cần sửa Secrets

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]
SHEET_KEY = st.secrets["SHEET_KEY"]  # đọc từ Secrets
WORKSHEET_NAME = "D25C"                                     # Đổi nếu cần
VN_TZ = datetime.timezone(datetime.timedelta(hours=7))

# Cấu hình kiểm soát vị trí theo cơ sở
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

st.set_page_config(page_title="QR Lecturer", layout="wide")

# ===================== GIAO DIỆN IN MINH CHỨNG =====================
# Chỉ làm chữ to, đậm, rõ hơn khi chụp/in; không thay đổi logic, không thay đổi Secrets.
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

[data-testid="stDataFrame"], table {
    font-size: 17px !important;
    color: #000000 !important;
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


# ===================== TIỆN ÍCH CHUNG =====================
def get_query_params():
    # Streamlit mới: st.query_params; fallback: experimental
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

def attendance_flag(val) -> bool:
    return str(val or "").strip() != ""

# ===================== MẬT KHẨU GV (Secrets/ENV) =====================
def _get_teacher_pw():
    # Đọc đúng tên biến đang có trong Secrets của người dùng.
    # Không yêu cầu đổi, thêm hoặc xóa bất kỳ dòng nào trong Secrets.
    if "ADMIN_PASSWORD" in st.secrets:
        return st.secrets["ADMIN_PASSWORD"]
    if "teacher_password" in st.secrets:
        return st.secrets["teacher_password"]
    if "google_service_account" in st.secrets:
        maybe = st.secrets["google_service_account"].get("teacher_password")
        if maybe:
            return maybe
    return os.getenv("ADMIN_PASSWORD") or os.getenv("TEACHER_PASSWORD")

def gv_unlocked() -> bool:
    return bool(st.session_state.get("gv_unlocked"))

def render_gv_auth():
    with st.sidebar:
        st.header("🔒 Đăng nhập Giảng viên")
        if gv_unlocked():
            st.success("Đã đăng nhập")
            if st.button("Đăng xuất"):
                st.session_state.clear()
                st.rerun()
        else:
            pw_input = st.text_input("Mật khẩu", type="password", key="pw_gv")
            if st.button("Đăng nhập", type="primary", use_container_width=True):
                if _get_teacher_pw() and pw_input == _get_teacher_pw():
                    st.session_state["gv_unlocked"] = True
                    st.rerun()
                else:
                    st.warning("Sai mật khẩu hoặc chưa cấu hình `ADMIN_PASSWORD` trong Secrets/ENV.")

# ===================== KẾT NỐI GOOGLE SHEETS =====================
@st.cache_resource
def _get_gspread_client():
    if "google_service_account" not in st.secrets:
        raise RuntimeError("Thiếu block [google_service_account] trong Secrets.")
    cred = dict(st.secrets["google_service_account"])
    pk = cred.get("private_key", "")
    if not pk:
        raise RuntimeError("Secrets thiếu private_key.")

    # Chuẩn hóa xuống dòng
    if "\\n" in pk:
        pk = pk.replace("\\n", "\n")
    pk = pk.replace("\r\n", "\n").replace("\r", "\n")

    header = "-----BEGIN PRIVATE KEY-----"
    footer = "-----END PRIVATE KEY-----"
    if header not in pk or footer not in pk:
        raise RuntimeError("private_key thiếu BEGIN/END.")

    # Làm sạch nội dung & padding base64
    lines = [ln.strip() for ln in pk.split("\n")]
    h_idx = lines.index(header)
    f_idx = lines.index(footer)
    body_raw = re.sub(r"[^A-Za-z0-9+/=]", "", "".join([ln for ln in lines[h_idx+1:f_idx] if ln]))
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

def get_sheet():
    client = _get_gspread_client()
    ss = client.open_by_key(SHEET_KEY)
    return ss.worksheet(WORKSHEET_NAME)

def load_records(sheet):
    return sheet.get_all_records(expected_headers=None, default_blank="")

def find_header_col(sheet, header_name):
    return sheet.find(header_name).col

# ===================== CỘT THỜI GIAN CẠNH CỘT BUỔI =====================
def find_or_create_time_col(sheet, buoi_col: int, buoi_header: str) -> int:
    headers = sheet.row_values(1)
    n_cols = len(headers)
    nxt = buoi_col + 1
    if nxt <= n_cols:
        h = (headers[nxt-1] or "").lower()
        if "thời gian" in h or "time" in h:
            return nxt
    m = re.search(r"(\d+)", buoi_header or "", flags=re.I)
    idx = m.group(1) if m else None
    if idx:
        for i, h in enumerate(headers, start=1):
            hl = (h or "").lower()
            if (("thời gian" in hl) or ("time" in hl)) and re.search(rf"\b{idx}\b", hl):
                return i
    # tạo mới ở cột kế bên
    sheet.update_cell(1, nxt, f"Thời gian {buoi_header}")
    return nxt

# ===================== TOKEN QR =====================
def current_slot(now=None, step=QR_SLOT_SECONDS):
    import time as _t
    return int((_t.time() if now is None else now) // step)

def token_valid(t_str: str, step=QR_SLOT_SECONDS, strict=True) -> bool:
    if not t_str or not str(t_str).isdigit():
        return False
    t = int(t_str)
    now_slot = current_slot(step=step)
    if strict:
        return t == now_slot
    return abs(t - now_slot) <= 1  # chấp nhận lệch ±1 slot nếu cần

# ===================== KIỂM TRA VỊ TRÍ THEO CƠ SỞ =====================
def render_location_check(campus_code: str):
    campus_name = LOCATION_BY_CODE.get(campus_code)
    if not campus_name:
        st.error("Cơ sở điểm danh không hợp lệ. Vui lòng quét lại mã QR mới.")
        st.stop()

    campus = LOCATIONS[campus_name]
    st.info(f"Cơ sở điểm danh: **{campus_name}** - {campus['address']}")

    if streamlit_geolocation is None or geodesic is None:
        st.error("Ứng dụng chưa cài đủ thư viện kiểm tra vị trí GPS.")
        st.code("pip install streamlit-geolocation geopy")
        st.stop()

    st.caption("Vui lòng cho phép trình duyệt truy cập vị trí để xác thực điểm danh.")
    location = streamlit_geolocation()

    if not location:
        st.warning("Chưa nhận được vị trí. Vui lòng bật GPS và cho phép trình duyệt truy cập vị trí.")
        st.stop()

    lat = location.get("latitude")
    lon = location.get("longitude")
    if lat is None or lon is None:
        st.warning("Không lấy được tọa độ GPS từ thiết bị. Vui lòng bật định vị và thử lại.")
        st.stop()

    student_loc = (float(lat), float(lon))
    campus_loc = (campus["lat"], campus["lon"])
    distance = geodesic(student_loc, campus_loc).meters

    st.caption(f"Khoảng cách hiện tại đến {campus_name}: **{distance:.0f} m**. Phạm vi cho phép: **{campus['radius']} m**.")
    if distance > campus["radius"]:
        st.error(f"Bạn đang ngoài phạm vi điểm danh của {campus_name}. Không thể điểm danh.")
        st.stop()

    st.success("Vị trí hợp lệ. Bạn có thể tiếp tục điểm danh.")

# ===================== CÁC MỤC GIAO DIỆN =====================
def render_tab_gv():
    # Lấy tên lớp theo tên worksheet (động)
    try:
        class_name = get_sheet().title
    except Exception:
        class_name = WORKSHEET_NAME

    st.subheader(f"📸 Mã QR điểm danh lớp {class_name} (QR động mỗi {QR_SLOT_SECONDS} giây)")

    col1, col2 = st.columns(2)
    with col1:
        co_so = st.selectbox("Chọn cơ sở giảng dạy", list(LOCATIONS.keys()), index=0)
    with col2:
        buoi = st.selectbox(
            "Chọn buổi học",
            ["Buổi 1", "Buổi 2", "Buổi 3", "Buổi 4", "Buổi 5", "Buổi 6"],
            index=0, key="buoi_gv_select",
        )

    auto = st.toggle("Tự đổi QR mỗi 30 giây", value=True)
    show_link = st.toggle("🔎 Hiển thị link chi tiết (ẩn/hiện)", value=False,
                          help="Bật khi cần xem toàn bộ URL để debug")
    go = st.button("Tạo mã QR", use_container_width=True, type="primary")
    campus_code = LOCATIONS[co_so]["code"]

    if go:
        qr_slot = st.empty()
        link_slot = st.empty()
        timer_slot = st.empty()
        campus_slot = st.empty()
        try:
            while True:
                now = int(time.time())
                token = f"{now // QR_SLOT_SECONDS}"
                base_url = st.secrets.get("WRAPPER_URL") or st.secrets.get("APP_BASE_URL") or st.secrets.get("google_service_account", {}).get(
                    "app_base_url", "https://qrlecturer.streamlit.app"
                )
                qr_data = (
                    f"{base_url}/?sv=1"
                    f"&buoi={urllib.parse.quote(buoi)}"
                    f"&coso={urllib.parse.quote(campus_code)}"
                    f"&t={token}"
                )

                qr = qrcode.make(qr_data)
                buf = io.BytesIO(); qr.save(buf, format="PNG"); buf.seek(0)
                img = Image.open(buf)

                campus_slot.info(f"{co_so} - {LOCATIONS[co_so]['address']}")
                qr_slot.image(img, caption="📱 Quét mã để điểm danh", width=360)

                if show_link:
                    with link_slot.container():
                        st.markdown(
                            f'<a href="{qr_data}" target="_blank" rel="noopener noreferrer">🌐 Mở link hiện tại</a>',
                            unsafe_allow_html=True
                        )
                        st.code(qr_data)
                else:
                    link_slot.empty()

                remain = QR_SLOT_SECONDS - (now % QR_SLOT_SECONDS)
                timer_slot.markdown(f"⏳ QR đổi sau: **{remain} giây**  •  Buổi: **{buoi}**  •  Cơ sở: **{co_so}**")

                if not auto:
                    break
                time.sleep(1)
        except Exception as e:
            st.error(f"❌ Lỗi khi tạo QR: {e}")

def find_student_candidates(records, query: str):
    q = (query or "").strip()
    if not q:
        return []
    if q.isdigit() and len(q) == 4:
        return [r for r in records if str(r.get("MSSV", "")).strip().endswith(q)]
    qn = norm_search(q)
    contains = [r for r in records if qn in norm_search(r.get("Họ và Tên", ""))]
    if contains:
        return contains
    names = [r.get("Họ và Tên", "") for r in records]
    name_map = {n: r for n, r in zip(names, records)}
    close = get_close_matches(q, names, n=5, cutoff=0.6)
    if not close:
        names_no = [norm_search(n) for n in names]
        name_map_no = {norm_search(n): n for n in names}
        close_no = get_close_matches(qn, names_no, n=5, cutoff=0.6)
        close = [name_map_no[c] for c in close_no]
    return [name_map[n] for n in close]

def render_tab_search():
    st.subheader("🔎 Tìm sinh viên (4 số cuối MSSV hoặc họ và tên)")
    q = st.text_input("Nhập từ khóa tìm kiếm", placeholder="VD: 1234 hoặc 'Nguyen Van A'")
    run = st.button("Tìm", use_container_width=True)

    if run and q.strip():
        try:
            sheet = get_sheet()
            records = load_records(sheet)
            results = find_student_candidates(records, q)

            if not results:
                st.warning("🙁 Không tìm thấy kết quả phù hợp.")
            else:
                st.success(f"Tìm thấy {len(results)} kết quả:")
                show_cols = list(records[0].keys()) if records else []
                pref = ["MSSV", "Họ và Tên", "Tổ"]
                buoi_cols = [c for c in show_cols if norm_search(c).startswith("buổi ")]
                cols = [c for c in pref if c in show_cols] + buoi_cols

                tidy = []
                for r in results:
                    row = {c: r.get(c, "") for c in cols}
                    for bc in buoi_cols:
                        row[bc] = "✅" if attendance_flag(r.get(bc, "")) else ""
                    tidy.append(row)
                st.dataframe(tidy, use_container_width=True)
        except Exception as e:
            st.error(f"❌ Lỗi khi tìm kiếm: {e}")

def render_tab_stats():
    st.subheader("📊 Thống kê điểm danh theo buổi & theo Tổ")
    try:
        sheet = get_sheet()
        headers = sheet.row_values(1)
        buoi_list = [h for h in headers if norm_search(h).startswith("buổi ")]
        buoi_chon = st.selectbox("Chọn buổi", buoi_list or ["Buổi 1"], index=0)
        records = load_records(sheet)

        present, absent = 0, 0
        by_group = {}
        for r in records:
            flag = attendance_flag(r.get(buoi_chon, ""))
            if flag: present += 1
            else: absent += 1
            group = str(r.get("Tổ", "")).strip() or "Chưa rõ"
            if group not in by_group:
                by_group[group] = {"present": 0, "absent": 0}
            by_group[group]["present" if flag else "absent"] += 1

        c1, c2, c3 = st.columns(3)
        with c1: st.metric("✅ Có mặt", present)
        with c2: st.metric("❌ Vắng", absent)
        with c3:
            total = present + absent
            st.metric("📈 Tỷ lệ có mặt", f"{(present/total*100):.1f}%" if total else "-")

        rows = []
        for g, v in sorted(by_group.items()):
            total_g = v["present"] + v["absent"]
            rate = (v["present"]/total_g*100) if total_g else 0.0
            rows.append({
                "Tổ": g, "Có mặt": v["present"], "Vắng": v["absent"],
                "Tổng": total_g, "Tỷ lệ (%)": round(rate, 1),
                "Nhãn": f"{v['present']} ({rate:.1f}%)"
            })
        df = pd.DataFrame(rows)
        if not df.empty:
            base = alt.Chart(df).encode(
                x=alt.X('Tổ:N', sort=None, title='Tổ'),
                y=alt.Y('Có mặt:Q', title='Số SV có mặt'),
                color=alt.Color('Tổ:N', legend=None),
                tooltip=[
                    alt.Tooltip('Tổ:N', title='Tổ'),
                    alt.Tooltip('Có mặt:Q', title='Có mặt'),
                    alt.Tooltip('Vắng:Q', title='Vắng'),
                    alt.Tooltip('Tổng:Q', title='Tổng'),
                    alt.Tooltip('Tỷ lệ (%):Q', title='Tỷ lệ (%)')
                ]
            )
            bars = base.mark_bar()
            text = base.mark_text(dy=-5).encode(text='Nhãn:N')
            chart = (bars + text).properties(height=420)
            st.altair_chart(chart, use_container_width=True)
        else:
            st.info("Không có dữ liệu để vẽ biểu đồ.")

        table = []
        for g, v in sorted(by_group.items()):
            total_g = v["present"] + v["absent"]
            rate_g = f"{(v['present']/total_g*100):.1f}%" if total_g else "-"
            table.append({"Tổ": g, "Có mặt": v["present"], "Vắng": v["absent"], "Tỷ lệ có mặt": rate_g})
        st.dataframe(table, use_container_width=True)
    except Exception as e:
        st.error(f"❌ Lỗi khi lấy thống kê: {e}")

# ===== Trợ lý AI (nâng cấp) – chạy nội bộ, không dùng API ngoài =====
def render_tab_ai():
    import unicodedata, re, datetime
    from difflib import get_close_matches

    st.subheader("🤖 Trợ lý AI (nội bộ, không dùng API ngoài)")
    st.caption(
        "Ví dụ: “Buổi 3 có bao nhiêu SV đi học?”, “Tổ 2 buổi 5 có bao nhiêu SV có mặt?”, "
        "“Ai đi học sớm nhất buổi 2?”, “Ai đến muộn nhất buổi 4?”, "
        "“Buổi 1 Thái có đi học không?”, “MSSV 5112xxxx đi mấy buổi?”, “Nguyen Van A có vắng không?”"
    )
    q_raw = st.text_input("Câu hỏi của bạn", placeholder="Nhập câu hỏi tiếng Việt (có thể gõ không dấu)...")

    # ===== Helpers NLP =====
    def lv_norm(s: str) -> str:
        s = (s or "").strip().lower()
        s = unicodedata.normalize("NFD", s)
        s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
        s = unicodedata.normalize("NFC", s)
        s = re.sub(r"\s+", " ", s)
        return s

    def fuzzy_has(text_norm: str, variants: list[str], thresh: float = 0.8) -> bool:
        from difflib import SequenceMatcher
        for v in variants:
            v2 = lv_norm(v)
            if v2 in text_norm:
                return True
            if SequenceMatcher(None, text_norm, v2).ratio() >= thresh:
                return True
        return False

    def extract_buoi(text_norm: str, buoi_cols: list[str]) -> str | None:
        for b in buoi_cols:
            if lv_norm(b) in text_norm:
                return b
        m = re.search(r"\bbuoi\s*(\d+)\b", text_norm)
        if m:
            num = m.group(1)
            for b in buoi_cols:
                if re.search(rf"\b{num}\b", lv_norm(b)):
                    return b
        return None

    def extract_to(text_norm: str) -> str | None:
        m = re.search(r"\bto\s*([a-z0-9]+)\b", text_norm)
        return m.group(1) if m else None

    def looks_like_mssv(s: str) -> bool:
        s = re.sub(r"\D", "", s or "")
        return len(s) >= 7

    def extract_mssv(text_norm: str) -> str | None:
        m = re.search(r"(?:mssv|sv|student)\s*([0-9]{6,})", text_norm)
        if m: return m.group(1)
        m2 = re.search(r"\b([0-9]{7,})\b", text_norm)
        return m2.group(1) if m2 else None

    def find_student_row(records: list[dict], mssv_or_name: str) -> dict | None:
        if looks_like_mssv(mssv_or_name):
            ms = re.sub(r"\D", "", mssv_or_name)
            for r in records:
                if re.sub(r"\D", "", str(r.get("MSSV",""))) == ms:
                    return r
        target = lv_norm(mssv_or_name)
        names = [r.get("Họ và Tên","") for r in records]
        name_map = {n: r for n, r in zip(names, records)}
        for n in names:
            if target and target in lv_norm(n):
                return name_map[n]
        cand = get_close_matches(mssv_or_name, names, n=1, cutoff=0.6)
        return name_map[cand[0]] if cand else None

    def extract_name_candidate(text_norm: str) -> str | None:
        stop = {
            "buoi","buổi","to","tổ","mssv","sv","student",
            "di","đi","hoc","học","co","có","mat","mặt","vang","vắng",
            "khong","không","ai","nhat","nhất","som","sớm","muon","muộn",
            "den","đến","tre","trễ","bao","nhiu","nhieu","bao nhieu",
            "ty","le","ty le","chuyen","can","chuyên","cần","trung","binh","trung binh",
            "la","là","khong di","co di","khong co mat"
        }
        tokens = re.findall(r"[a-zA-ZÀ-ỹ0-9]+", text_norm)
        remain = [t for t in tokens if t not in stop and not t.isdigit()]
        name = " ".join(remain).strip()
        return name if name else None

    # ===== Dò cột Buổi + Thời gian (linh hoạt) =====
    def detect_buoi_columns(headers: list[str]) -> list[str]:
        cols = []
        for h in headers:
            hn = norm_search(h).replace("_", " ").replace("-", " ")
            if re.match(r"^(b|bu|buoi)\s*\d+$", hn):
                cols.append(h); continue
            if hn.startswith("buoi ") and re.search(r"\d+", hn):
                cols.append(h); continue
            if norm_search(h).startswith("buổi ") and re.search(r"\d+", norm_search(h)):
                cols.append(h); continue
        seen, out = set(), []
        for h in cols:
            if h not in seen:
                seen.add(h); out.append(h)
        return out

    def build_time_map(headers: list[str], buoi_cols: list[str]) -> dict[str,int|None]:
        name_to_idx = {h: i+1 for i, h in enumerate(headers)}
        time_map = {}
        for b in buoi_cols:
            idx = name_to_idx[b]
            m = re.search(r"(\d+)", b)
            num = m.group(1) if m else None
            tcol = None
            if num:
                for i, h in enumerate(headers, start=1):
                    hn = norm_search(h)
                    if (("thời gian" in h.lower()) or ("thoi gian" in hn) or ("time" in h.lower())) and re.search(rf"\b{num}\b", hn):
                        tcol = i; break
            if not tcol and idx < len(headers):
                right = headers[idx]  # cột bên phải (1-based -> headers[idx])
                hn = norm_search(right)
                if ("thời gian" in right.lower()) or ("time" in right.lower()) or ("thoi gian" in hn):
                    tcol = idx + 1
            time_map[b] = tcol
        return time_map

    def parse_time(val: str) -> datetime.datetime | None:
        if not val: return None
        val = str(val).strip()
        fmts = ["%Y-%m-%d %H:%M:%S", "%d/%m/%Y %H:%M:%S", "%H:%M:%S", "%H:%M"]
        for fmt in fmts:
            try:
                dt = datetime.datetime.strptime(val, fmt)
                if fmt in ("%H:%M:%S", "%H:%M"):
                    today = datetime.datetime.now(VN_TZ).date()
                    dt = datetime.datetime.combine(today, dt.time())
                return dt.replace(tzinfo=VN_TZ)
            except Exception:
                continue
        return None

    def answer(q_user: str) -> str:
        qn = norm_search(q_user)  # bỏ dấu + lower
        sheet = get_sheet()
        records = load_records(sheet)
        if not records:
            return "Không có dữ liệu trong Sheet."

        headers = sheet.row_values(1)
        buoi_cols = detect_buoi_columns(headers)
        if not buoi_cols:
            return "Không tìm thấy các cột 'Buổi ...' trong Sheet."
        time_map = build_time_map(headers, buoi_cols)
        total_sv = len(records)

        # ------ sớm nhất / muộn nhất theo cột thời gian ------
        ask_earliest = fuzzy_has(qn, ["som nhat", "sớm nhất", "den som nhat", "som nhut", "somnha"])
        ask_latest   = fuzzy_has(qn, ["muon nhat", "muộn nhất", "den muon nhat", "den tre nhat", "tre nhat"])
        if ask_earliest or ask_latest:
            b = extract_buoi(qn, buoi_cols) or buoi_cols[-1]
            t_col = time_map.get(b)
            if not t_col:
                return f"Không tìm thấy cột thời gian ứng với “{b}”."

            best_row, best_time = None, None
            for r_idx, r in enumerate(records, start=2):
                if not attendance_flag(r.get(b,"")):
                    continue
                try:
                    t_val = sheet.cell(r_idx, t_col).value
                except Exception:
                    t_val = ""
                t_parsed = parse_time(t_val)
                if not t_parsed:
                    continue
                if best_time is None:
                    best_time, best_row = t_parsed, r
                else:
                    if ask_earliest and t_parsed < best_time:
                        best_time, best_row = t_parsed, r
                    if ask_latest and t_parsed > best_time:
                        best_time, best_row = t_parsed, r
            if best_row is None:
                return f"Chưa có dữ liệu thời gian hợp lệ cho {b}."
            name = best_row.get("Họ và Tên","(không tên)")
            ms   = best_row.get("MSSV","?")
            kind = "sớm nhất" if ask_earliest else "muộn nhất"
            return f"👤 {name} ({ms}) là người {kind} trong {b}: {best_time.strftime('%Y-%m-%d %H:%M:%S')}."

        # ------ “Buổi X <tên> có đi học không?” → liệt kê tất cả tên khớp ------
        if any(w in qn for w in ["di hoc","co mat","vang","khong"]):
            b = extract_buoi(qn, buoi_cols)
            if b:
                name_guess = extract_name_candidate(qn)
                if name_guess:
                    target_norm = norm_search(name_guess)
                    matches = [r for r in records if target_norm in norm_search(r.get("Họ và Tên",""))]
                    if not matches:
                        names = [r.get("Họ và Tên","") for r in records]
                        close = get_close_matches(name_guess, names, n=5, cutoff=0.6)
                        name_map = {n: r for n, r in zip(names, records)}
                        matches = [name_map[n] for n in close]
                    if not matches:
                        return f"Không tìm thấy sinh viên nào khớp với “{name_guess}”."
                    lines = []
                    for r in matches:
                        flag = "✅" if attendance_flag(r.get(b,"")) else "❌"
                        lines.append(f"- {r.get('Họ và Tên','(không tên)')} ({r.get('MSSV','?')}): {flag} tại {b}")
                    return f"Kết quả cho “{name_guess}” ở {b}:\n" + "\n".join(lines)

        # ------ Thống kê theo buổi / tổng quan ------
        if any(w in qn for w in ["bao nhieu","di hoc","co mat","vang"]):
            present = {b: sum(1 for r in records if attendance_flag(r.get(b,""))) for b in buoi_cols}
            b = extract_buoi(qn, buoi_cols)
            if b:
                p = present[b]; a = total_sv - p
                return f"{b}: {p}/{total_sv} có mặt, {a} vắng ({(p/total_sv*100):.1f}%)."
            total_present_all = sum(present.values())
            total_slots = total_sv * len(buoi_cols)
            rate = total_present_all/total_slots*100 if total_slots else 0
            return f"Tổng tất cả buổi: {total_present_all}/{total_slots} lượt có mặt (~{rate:.1f}%)."

        # ------ Theo tổ ------
        if " to " in f" {qn} " or re.search(r"\bto\b", qn):
            b = extract_buoi(qn, buoi_cols) or buoi_cols[-1]
            target_to = extract_to(qn)
            stats = {}
            for r in records:
                g = str(r.get("Tổ","")).strip() or "Chưa rõ"
                stats.setdefault(g, {"present":0,"absent":0})
                if attendance_flag(r.get(b,"")):
                    stats[g]["present"] += 1
                else:
                    stats[g]["absent"] += 1
            if target_to and target_to in stats:
                v = stats[target_to]; tot = v["present"]+v["absent"]
                rate = v["present"]/tot*100 if tot else 0
                return f"{b} - Tổ {target_to}: {v['present']}/{tot} có mặt ({rate:.1f}%)."
            lines = []
            for g, v in sorted(stats.items()):
                tot = v["present"]+v["absent"]; rate = v["present"]/tot*100 if tot else 0
                lines.append(f"Tổ {g}: {v['present']}/{tot} ({rate:.1f}%)")
            return f"📊 {b} theo tổ:\n" + "\n".join(lines)

        # ------ Một sinh viên cụ thể (MSSV hoặc tên) ------
        if "mssv" in qn or re.search(r"\b[0-9]{7,}\b", qn) or any(k in qn for k in ["sv ","sinh vien","sinhvien"]):
            mssv = extract_mssv(qn)
            target = mssv if mssv else q_raw
            row = find_student_row(records, target)
            if not row:
                return "Không tìm thấy sinh viên tương ứng."
            name = row.get("Họ và Tên","(không tên)")
            ms   = row.get("MSSV","?")
            presents = 0; marks = []
            for b in buoi_cols:
                flag = "✅" if attendance_flag(row.get(b,"")) else "❌"
                if flag == "✅": presents += 1
                marks.append(f"{b}:{flag}")
            return f"{name} ({ms}) — {presents}/{len(buoi_cols)} buổi có mặt.\n" + ", ".join(marks)

        # ------ Tỷ lệ chuyên cần trung bình ------
        if "chuyen can" in qn or ("ty le" in qn and "buoi" not in qn):
            total_present_all = sum(
                sum(1 for r in records if attendance_flag(r.get(b,"")))
                for b in buoi_cols
            )
            total_slots = total_sv * len(buoi_cols)
            rate = total_present_all/total_slots*100 if total_slots else 0
            return f"📈 Tỷ lệ chuyên cần trung bình: {rate:.1f}%."

        # ------ Danh sách vắng quá N buổi ------
        m = re.search(r"vang\s+qua\s+(\d+)\s*buoi", qn)
        if m:
            limit = int(m.group(1))
            rows = []
            for r in records:
                vangs = sum(1 for b in buoi_cols if not attendance_flag(r.get(b,"")))
                if vangs > limit:
                    rows.append(f"- {r.get('Họ và Tên','(không tên)')} ({r.get('MSSV','?')}): {vangs} buổi")
            return "Danh sách vắng quá {} buổi:\n".format(limit) + ("\n".join(rows) if rows else "Không có.")

        # ------ fallback ------
        return ("🤔 Tôi chưa chắc ý bạn. Bạn có thể hỏi: "
                "“Ai đi học sớm nhất buổi 2?”, “Buổi 1 Thái có đi học không?”, "
                "“Buổi 3 có bao nhiêu SV đi học?”, “MSSV 5112xxxx đi mấy buổi?”")

    if st.button("Hỏi trợ lý", use_container_width=True) and q_raw.strip():
        try:
            st.markdown(f"**Trả lời:**\n\n{answer(q_raw)}")
        except Exception as e:
            st.error(f"❌ Lỗi khi xử lý câu hỏi: {e}")

# ===================== luồng trang: SV / GV =====================
qp = get_query_params()

# ---------- MÀN HÌNH SINH VIÊN ----------
if qp.get("sv") == "1":
    buoi_sv = qp.get("buoi", "Buổi 1")
    token_qr = qp.get("t", "")
    campus_code = qp.get("coso", "CS1")

    lock_key = f"locked_{buoi_sv}_{campus_code}"
    info_key = f"lock_info_{buoi_sv}_{campus_code}"

    st.title("🎓 Điểm danh sinh viên")
    st.info(f"Bạn đang điểm danh cho **{buoi_sv}**")

    # Nếu đã khóa vì đã điểm danh
    if st.session_state.get(lock_key):
        st.success(st.session_state.get(info_key, "Bạn đã điểm danh thành công."))
        st.stop()

    # ===== Mở khóa phiên theo session_state để tránh 'hết hạn' khi rerun =====
    unlock_key = f"sv_unlocked_{buoi_sv}_{campus_code}"      # lưu {'ts': epoch, 't': token}
    now_epoch = time.time()
    uinfo = st.session_state.get(unlock_key)

    if not uinfo:
        # Lần đầu vào: buộc token hợp lệ (nới lỏng ±1 slot để tránh sát ranh)
        if not token_valid(token_qr, step=QR_SLOT_SECONDS, strict=False):
            st.error("⏳ Link điểm danh đã hết hạn hoặc không hợp lệ. Vui lòng quét mã QR mới.")
            remain = QR_SLOT_SECONDS - (int(time.time()) % QR_SLOT_SECONDS)
            st.caption(f"Gợi ý: mã QR đổi sau khoảng {remain} giây.")
            st.stop()
        # Token hợp lệ -> mở khóa phiên với TTL
        st.session_state[unlock_key] = {"ts": now_epoch, "t": str(token_qr)}
    else:
        # Đã mở khóa -> cho dùng tiếp trong TTL mà KHÔNG kiểm token nữa
        if now_epoch - uinfo["ts"] > UNLOCK_TTL:
            st.warning("Phiên điểm danh đã hết thời gian. Vui lòng quét mã QR mới.")
            del st.session_state[unlock_key]
            st.stop()

    # ===== Kiểm tra vị trí GPS theo cơ sở trong mã QR =====
    render_location_check(campus_code)

    # ======= Form điểm danh: SV chỉ nhập 4 số cuối MSSV =======
    mssv_suffix = st.text_input(
        "Nhập **4 số cuối** MSSV",
        placeholder="VD: 1234",
        max_chars=4,
        help=f"Mã đầy đủ sẽ là {MSSV_PREFIX} + 4 số cuối bạn nhập"
    )
    hoten = st.text_input("Nhập họ và tên")

    if mssv_suffix.strip().isdigit():
        full_mssv_preview = f"{MSSV_PREFIX}{mssv_suffix.strip().zfill(4)}"
        st.caption(f"MSSV đầy đủ: **{full_mssv_preview}**")

    if st.button("✅ Xác nhận điểm danh", use_container_width=True):
        if not mssv_suffix.strip().isdigit() or len(mssv_suffix.strip()) != 4:
            st.warning("⚠️ Vui lòng nhập **đúng 4 số cuối** MSSV (chỉ số).")
            st.stop()
        if not hoten.strip():
            st.warning("⚠️ Vui lòng nhập họ và tên.")
            st.stop()

        full_mssv = f"{MSSV_PREFIX}{mssv_suffix.strip().zfill(4)}"

        try:
            sheet = get_sheet()
            col_buoi = find_header_col(sheet, buoi_sv)

            # Tìm hàng theo MSSV đầy đủ; nếu không thấy, fallback quét records
            try:
                cell_mssv = sheet.find(full_mssv)
            except Exception:
                cell_mssv = None

            if not cell_mssv:
                records = load_records(sheet)
                target_row = None
                for idx, r in enumerate(records, start=2):
                    ms = str(r.get("MSSV", "")).strip()
                    ms_norm = re.sub(r"\D", "", ms) if ms else ""
                    if ms_norm.startswith(MSSV_PREFIX) and ms_norm.endswith(mssv_suffix.strip().zfill(4)):
                        target_row = idx
                        break
                if target_row:
                    class DummyCell:
                        def __init__(self, row): self.row = row
                    cell_mssv = DummyCell(target_row)

            if not cell_mssv:
                st.error(f"❌ Không tìm thấy MSSV **{full_mssv}** trong danh sách.")
                st.stop()

            # Kiểm tra họ tên khớp
            hoten_sheet = sheet.cell(cell_mssv.row, find_header_col(sheet, "Họ và Tên")).value
            if normalize_name(hoten_sheet or "") != normalize_name(hoten):
                st.error("❌ Họ tên không khớp với MSSV trong danh sách.")
                st.stop()

            # Kiểm tra đã điểm danh trước đó
            curr_mark = (sheet.cell(cell_mssv.row, col_buoi).value or "").strip()
            time_col = find_or_create_time_col(sheet, col_buoi, buoi_sv)
            if curr_mark:
                exist_time = sheet.cell(cell_mssv.row, time_col).value or ""
                msg = f"✅ MSSV **{full_mssv}** đã điểm danh trước đó" + (f" lúc **{exist_time}**." if exist_time else ".")
                st.info(msg)
                st.session_state[lock_key] = True
                st.session_state[info_key] = msg
                st.rerun()

            # Ghi ✅ và thời gian thực
            sheet.update_cell(cell_mssv.row, col_buoi, "✅")
            now_str = datetime.datetime.now(VN_TZ).strftime("%Y-%m-%d %H:%M:%S")
            sheet.update_cell(cell_mssv.row, time_col, now_str)

            msg = f"🎉 Điểm danh thành công! MSSV **{full_mssv}** ({now_str})."
            st.success(msg)
            st.session_state[lock_key] = True
            st.session_state[info_key] = msg
            st.rerun()

        except Exception as e:
            st.error(f"❌ Lỗi khi điểm danh: {e}")

    st.stop()

# ---------- MÀN HÌNH GIẢNG VIÊN ----------
render_gv_auth()
st.title("📋 Hệ thống điểm danh QR")

if not gv_unlocked():
    st.error("🔒 Bạn chưa đăng nhập Giảng viên. Vào **Sidebar → Đăng nhập Giảng viên** để mở khóa.")
    st.stop()

# Điều hướng ở Sidebar
with st.sidebar:
    st.markdown("---")
    st.markdown("**📂 Điều hướng**")
    menu = st.radio(
        "Chọn mục",
        options=["👨‍🏫 Giảng viên (QR động)", "🔎 Tìm kiếm", "📊 Thống kê", "🤖 Trợ lý AI"],
        index=0,
        label_visibility="collapsed"
    )

# Nội dung ở khung chính
if menu == "👨‍🏫 Giảng viên (QR động)":
    render_tab_gv()
elif menu == "🔎 Tìm kiếm":
    render_tab_search()
elif menu == "📊 Thống kê":
    render_tab_stats()
else:
    render_tab_ai()

# ---------- FOOTER (bản quyền, căn giữa) ----------
st.markdown(
    """
    <style>
    .footer-dhn {
        position: fixed;
        left: 0; right: 0; bottom: 0;
        padding: 8px 16px;
        background: rgba(0,0,0,0.04);
        color: #444;
        font-size: 12px;
        text-align: center;
        z-index: 1000;
        border-top: 1px solid rgba(0,0,0,0.1);
        width: 100%;
    }
    </style>
    <div class="footer-dhn"> Copyright © 2025 Bản quyền thuộc về <strong>TS. Đào Hồng Nam - Đại học Y Dược Thành phố Hồ Chí Minh</strong></div>
    """,
    unsafe_allow_html=True
)



