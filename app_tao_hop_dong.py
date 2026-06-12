import streamlit as st
import io, os, re, unicodedata
from datetime import datetime, date
from docx import Document

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="MobiFone Invoice – Tạo Hợp Đồng",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
[data-testid="stAppViewContainer"] { background: #f0f4f8; }
[data-testid="stSidebar"] { background: #1a3a5c; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] hr { border-color: #2d5a8c !important; }

.header-banner {
    background: linear-gradient(135deg, #1a3a5c 0%, #2563a8 60%, #1a3a5c 100%);
    border-radius: 14px; padding: 24px 32px; margin-bottom: 20px;
    display: flex; align-items: center; gap: 18px;
    box-shadow: 0 8px 32px rgba(0,0,0,0.18);
}
.header-banner h1 { color:#fff; margin:0; font-size:1.6rem; font-weight:800; }
.header-banner p  { color:#bfdbfe; margin:4px 0 0; font-size:0.88rem; }

.section-card {
    background:#fff; border-radius:12px; padding:20px 24px 16px;
    margin-bottom:16px; box-shadow:0 2px 10px rgba(0,0,0,0.07);
    border-left:4px solid #2563a8;
}
.section-title {
    font-weight:700; font-size:0.95rem; color:#1a3a5c;
    margin-bottom:12px; letter-spacing:0.3px;
}

.preview-badge {
    background:#eff6ff; border:1px solid #93c5fd; border-radius:8px;
    padding:10px 14px; font-family:monospace; font-size:0.82rem;
    color:#1d4ed8; word-break:break-all;
}
.ma-kh-badge {
    background:#f0fdf4; border:1px solid #86efac; border-radius:6px;
    padding:6px 12px; font-family:monospace; font-size:0.8rem; color:#166534;
    display:inline-block; margin-top:6px;
}
.error-item {
    background:#fef2f2; border:1px solid #fca5a5; border-radius:7px;
    padding:8px 12px; font-size:0.82rem; color:#991b1b; margin-bottom:5px;
}
.history-item {
    background:#1e3a5c; border-radius:8px; padding:9px 13px;
    margin-bottom:6px; font-size:0.79rem; color:#93c5e8;
}
.history-item span { color:#e2e8f0; font-weight:600; }
.history-item .time { color:#64748b; font-size:0.75rem; }

.metric-strip { display:flex; gap:10px; margin-bottom:18px; flex-wrap:wrap; }
.metric-box {
    background:white; border-radius:10px; padding:12px 18px;
    flex:1; min-width:110px; box-shadow:0 2px 8px rgba(0,0,0,0.06);
    text-align:center;
}
.metric-box .val { font-size:1.45rem; font-weight:800; color:#1a3a5c; }
.metric-box .lbl { font-size:0.71rem; color:#94a3b8; margin-top:2px; }

.stDownloadButton > button {
    background:linear-gradient(135deg,#16693a,#15803d) !important;
    color:white !important; border:none !important;
    border-radius:10px !important; font-weight:700 !important;
    font-size:1.05rem !important; padding:14px 28px !important;
    width:100% !important; box-shadow:0 4px 16px rgba(22,105,58,0.35) !important;
}
.stDownloadButton > button:hover { transform:translateY(-1px) !important; }

div[data-testid="stButton"] > button[kind="primary"] {
    background:linear-gradient(135deg,#1a3a5c,#2563a8) !important;
    color:white !important; border-radius:8px !important;
    font-weight:700 !important; border:none !important;
    font-size:1rem !important;
}

/* Required star */
.req { color:#dc2626; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# SESSION STATE
# ─────────────────────────────────────────────
for k, default in [
    ("history", []),
    ("draft", {}),
    ("generated", False),
    ("result_bytes", None),
    ("result_fname", ""),
]:
    if k not in st.session_state:
        st.session_state[k] = default

# ─────────────────────────────────────────────
# CORE LOGIC  (ported 1-to-1 từ tkinter v7)
# ─────────────────────────────────────────────
def bo_dau(text: str) -> str:
    text = unicodedata.normalize("NFD", text)
    return "".join(c for c in text if unicodedata.category(c) != "Mn")

def tao_ma_hd(ten_hkd: str) -> str:
    return bo_dau(ten_hkd.upper()).replace(" ", "")

def filename_preview(so_hd: str, ten_hkd: str) -> str:
    if so_hd and ten_hkd:
        return f"HD_{so_hd}_{tao_ma_hd(ten_hkd)}.docx"
    return "HD_???.docx"

def replace_in_para(para, old: str, new: str):
    """Replace placeholder trong paragraph, giữ nguyên formatting."""
    full = "".join(r.text for r in para.runs)
    if old not in full:
        return
    replaced = False
    for run in para.runs:
        if old in run.text:
            run.text = run.text.replace(old, new)
            replaced = True
            break
    if not replaced and para.runs:
        para.runs[0].text = full.replace(old, new)
        for run in para.runs[1:]:
            run.text = ""

def get_all_paras(doc):
    paras = list(doc.paragraphs)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                paras.extend(cell.paragraphs)
    return paras

def _xoa_ngay_ky(doc, ngay: str, thang: str, nam: str):
    """Smart replace ngày ký — giống hệt logic tkinter v7."""
    ngay_str = ngay if ngay else "  "
    all_paras = get_all_paras(doc)
    for para in all_paras:
        full = "".join(r.text for r in para.runs)
        if (
            "lập vào ngày" in full
            or ("từ ngày" in full and "28" in full)
            or (
                "Ngày" in full and "tháng" in full and nam in full
                and "Luật" not in full and "Quốc Hội" not in full
            )
        ):
            for run in para.runs:
                if run.text == "28":
                    run.text = ngay_str
            replace_in_para(para, "lập vào ngày 28", f"lập vào ngày {ngay_str}")
            replace_in_para(para, "28/05/",           f"{ngay_str}/{thang}/")
            replace_in_para(para, "Ngày 28 tháng",    f"Ngày {ngay_str} tháng")

def tao_mot_hop_dong(kh: dict, template_bytes: bytes) -> bytes:
    """Tạo hợp đồng, trả về bytes của file docx."""
    doc = Document(io.BytesIO(template_bytes))

    email_kh = kh.get("email_kh", "")
    ep = email_kh.split("@")[0] if "@" in email_kh else email_kh
    es = ("@" + email_kh.split("@")[1]) if "@" in email_kh else ""

    ten_hkd    = kh.get("ten_hkd", "")
    ten_hkd_ng = ten_hkd[4:].strip() if ten_hkd.upper().startswith("HKD ") else ten_hkd
    ma_hd      = tao_ma_hd(ten_hkd)
    tru_so     = kh.get("tru_so", "")
    dia_chi_gd = kh.get("dia_chi_gd", "") or tru_so
    ngay       = kh.get("ngay_ky", "")
    thang      = kh.get("thang_ky", "")
    nam        = kh.get("nam_ky", "2026")

    replacements = [
        ("EMAIL_KH_PLACEHOLDER",                       ep),
        ("@GMAIL.COM",                                  es),
        ("101",                                         kh.get("so_hd", "")),
        ("HKDKIMKIEN",                                  ma_hd),
        ("NHÀ NGHỈ KIM LIÊN",                          ten_hkd_ng),
        ("142/2 Phan Châu Trinh P.Hải Châu, Tp Đà Nẵng", tru_so),
        ("0904725978",                                  kh.get("dien_thoai", "")),
        ("049189011144",                                kh.get("ma_so_thue", "")),
        ("PHẠM THỊ KIM LIÊN",                          kh.get("dai_dien", "")),
        ("Chủ HKD",                                     kh.get("chuc_vu", "")),
        ("Trương Thị Mỹ Châu",                         kh.get("ten_nv_ben_b", "")),
        ("0901959799",                                  kh.get("sdt_ben_b", "")),
        ("cuong.danghuy.ctv@mobifone.vn",               kh.get("email_ben_b", "")),
        ("tháng 05",                                    f"tháng {thang}"),
        ("năm 2026",                                    f"năm {nam}"),
        ("/2026",                                       f"/{nam}"),
    ]

    all_paras = get_all_paras(doc)
    for old, new in replacements:
        if old and new and old != new:
            for para in all_paras:
                replace_in_para(para, old, new)

    # Địa chỉ giao dịch khác trụ sở → replace lần xuất hiện thứ 2
    if dia_chi_gd != tru_so and tru_so:
        found = 0
        for para in get_all_paras(doc):
            if tru_so in "".join(r.text for r in para.runs):
                found += 1
                if found == 2:
                    replace_in_para(para, tru_so, dia_chi_gd)
                    break

    _xoa_ngay_ky(doc, ngay, thang, nam)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()

def validate(kh: dict) -> list:
    req_map = {
        "so_hd":       "Số hợp đồng",
        "ten_hkd":     "Tên HKD",
        "tru_so":      "Trụ sở chính",
        "dien_thoai":  "Số điện thoại",
        "ma_so_thue":  "Mã số thuế",
        "dai_dien":    "Họ tên đại diện",
        "chuc_vu":     "Chức vụ",
        "email_kh":    "Email khách hàng",
        "thang_ky":    "Tháng ký",
        "nam_ky":      "Năm ký",
        "ten_nv_ben_b":"Tên nhân viên Bên B",
        "email_ben_b": "Email nhân viên Bên B",
        "sdt_ben_b":   "SĐT nhân viên Bên B",
    }
    errors = []
    for k, label in req_map.items():
        if not kh.get(k, "").strip():
            errors.append(f"Thiếu: **{label}**")
    if kh.get("dien_thoai") and not re.match(r"^0\d{9}$", kh["dien_thoai"]):
        errors.append("Số điện thoại KH phải 10 chữ số, bắt đầu bằng 0")
    if kh.get("sdt_ben_b") and not re.match(r"^0\d{9}$", kh["sdt_ben_b"]):
        errors.append("SĐT nhân viên Bên B phải 10 chữ số, bắt đầu bằng 0")
    if kh.get("email_kh") and "@" not in kh["email_kh"]:
        errors.append("Email khách hàng không đúng định dạng")
    if kh.get("email_ben_b") and "@" not in kh["email_ben_b"]:
        errors.append("Email nhân viên Bên B không đúng định dạng")
    return errors

# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 📁 File Template")
    template_file = st.file_uploader(
        "Upload hop_dong_template.docx",
        type=["docx"],
        label_visibility="collapsed",
    )
    st.markdown("---")

    st.markdown("### ⚙️ Bên B — Mặc định")
    ten_nv_ben_b  = st.text_input("Tên nhân viên *",  value=st.session_state.draft.get("ten_nv_ben_b", "Trương Thị Mỹ Châu"))
    email_ben_b   = st.text_input("Email nhân viên *", value=st.session_state.draft.get("email_ben_b",  "cuong.danghuy.ctv@mobifone.vn"))
    sdt_ben_b     = st.text_input("SĐT nhân viên *",  value=st.session_state.draft.get("sdt_ben_b",    "0901959799"))
    st.markdown("---")

    st.markdown("### 🕐 Lịch sử (5 gần nhất)")
    if not st.session_state.history:
        st.markdown("<div style='color:#475569;font-size:0.8rem'>Chưa có hợp đồng nào</div>", unsafe_allow_html=True)
    else:
        for item in reversed(st.session_state.history[-5:]):
            st.markdown(f"""
            <div class="history-item">
                <span>{item['fname']}</span><br>
                <span class="time">{item['time']}</span>
            </div>
            """, unsafe_allow_html=True)
        if st.button("🗑️ Xoá lịch sử", use_container_width=True):
            st.session_state.history = []
            st.rerun()

# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown("""
<div class="header-banner">
    <div style="font-size:2.8rem">📄</div>
    <div>
        <h1>Tool Tạo Hợp Đồng Tự Động</h1>
        <p>MobiFone Invoice · Minh Thông · 0788 563 777 · Phiên bản Web v7</p>
    </div>
</div>
""", unsafe_allow_html=True)

# METRICS
total       = len(st.session_state.history)
today_count = sum(1 for h in st.session_state.history
                  if h["time"].startswith(date.today().strftime("%d/%m/%Y")))
tpl_status  = "✅ Đã chọn" if template_file else "❌ Chưa có"
st.markdown(f"""
<div class="metric-strip">
    <div class="metric-box"><div class="val">{total}</div><div class="lbl">HĐ đã tạo</div></div>
    <div class="metric-box"><div class="val">{today_count}</div><div class="lbl">Hôm nay</div></div>
    <div class="metric-box"><div class="val" style="font-size:1rem">{tpl_status}</div><div class="lbl">Template</div></div>
</div>
""", unsafe_allow_html=True)

if not template_file:
    st.warning("⬅️  Upload file **hop_dong_template.docx** ở sidebar để bắt đầu.")
    st.stop()

template_bytes = template_file.read()

# ─────────────────────────────────────────────
# FORM
# ─────────────────────────────────────────────
col_l, col_r = st.columns([3, 2], gap="large")

with col_l:
    # ── SEC 1: HỢP ĐỒNG ──────────────────────
    st.markdown('<div class="section-card"><div class="section-title">📋 THÔNG TIN HỢP ĐỒNG</div></div>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        so_hd    = st.text_input("Số hợp đồng *", value=st.session_state.draft.get("so_hd",""), placeholder="101")
    with c2:
        ngay_ky  = st.text_input("Ngày ký",       value=st.session_state.draft.get("ngay_ky", datetime.now().strftime("%d")), placeholder="28")
    with c3:
        thang_ky = st.text_input("Tháng ký *",    value=st.session_state.draft.get("thang_ky", datetime.now().strftime("%m")), placeholder="05")
    with c4:
        nam_ky   = st.text_input("Năm ký *",      value=st.session_state.draft.get("nam_ky",   datetime.now().strftime("%Y")), placeholder="2026")

    st.markdown("---")

    # ── SEC 2: BÊN A ─────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">🏪 BÊN A — KHÁCH HÀNG</div></div>', unsafe_allow_html=True)

    ca1, ca2 = st.columns(2)
    with ca1:
        ten_hkd  = st.text_input("Tên HKD đầy đủ *", value=st.session_state.draft.get("ten_hkd",""),  placeholder="HKD Nhà Nghỉ Kim Liên")
    with ca2:
        tru_so   = st.text_input("Trụ sở chính *",    value=st.session_state.draft.get("tru_so",""),   placeholder="142/2 Phan Châu Trinh, P.Hải Châu, Tp Đà Nẵng")

    cb1, cb2 = st.columns(2)
    with cb1:
        dia_chi_gd = st.text_input("Địa chỉ giao dịch (nếu khác trụ sở)", value=st.session_state.draft.get("dia_chi_gd",""), placeholder="Để trống nếu giống trụ sở")
    with cb2:
        dien_thoai = st.text_input("Số điện thoại *", value=st.session_state.draft.get("dien_thoai",""), placeholder="0904725978")

    cc1, cc2 = st.columns(2)
    with cc1:
        ma_so_thue = st.text_input("Mã số thuế *",         value=st.session_state.draft.get("ma_so_thue",""), placeholder="049189011144")
    with cc2:
        dai_dien   = st.text_input("Họ tên đại diện *",    value=st.session_state.draft.get("dai_dien",""),   placeholder="PHẠM THỊ KIM LIÊN")

    cd1, cd2 = st.columns(2)
    with cd1:
        chuc_vu  = st.text_input("Chức vụ *",          value=st.session_state.draft.get("chuc_vu","Chủ HKD"), placeholder="Chủ HKD")
    with cd2:
        email_kh = st.text_input("Email khách hàng *", value=st.session_state.draft.get("email_kh",""),       placeholder="kimlienpham@gmail.com")

with col_r:
    # ── Preview tên file ─────────────────────
    st.markdown('<div class="section-card"><div class="section-title">👁️ Preview tên file</div></div>', unsafe_allow_html=True)
    fname_preview = filename_preview(so_hd, ten_hkd)
    st.markdown(f'<div class="preview-badge">📄 {fname_preview}</div>', unsafe_allow_html=True)
    if ten_hkd:
        ma = tao_ma_hd(ten_hkd)
        st.markdown(f'<div class="ma-kh-badge">Mã KH: {ma}</div>', unsafe_allow_html=True)

    st.markdown("---")

    # ── Nháp ─────────────────────────────────
    st.markdown('<div class="section-card"><div class="section-title">💾 Lưu nháp</div></div>', unsafe_allow_html=True)
    cs1, cs2 = st.columns(2)

    current_data = {
        "so_hd": so_hd, "ngay_ky": ngay_ky, "thang_ky": thang_ky, "nam_ky": nam_ky,
        "ten_hkd": ten_hkd, "tru_so": tru_so, "dia_chi_gd": dia_chi_gd,
        "dien_thoai": dien_thoai, "ma_so_thue": ma_so_thue,
        "dai_dien": dai_dien.upper() if dai_dien else "",
        "chuc_vu": chuc_vu, "email_kh": email_kh,
        "ten_nv_ben_b": ten_nv_ben_b, "email_ben_b": email_ben_b, "sdt_ben_b": sdt_ben_b,
    }

    with cs1:
        if st.button("💾 Lưu nháp", use_container_width=True):
            st.session_state.draft = current_data
            st.success("Đã lưu nháp!")
    with cs2:
        if st.button("🗑️ Xoá nháp", use_container_width=True):
            st.session_state.draft = {}
            st.session_state.generated = False
            st.rerun()

    st.markdown("---")

    # ── Validation realtime ───────────────────
    st.markdown('<div class="section-card"><div class="section-title">✅ Kiểm tra thông tin</div></div>', unsafe_allow_html=True)
    errors = validate(current_data)
    if errors:
        for e in errors:
            st.markdown(f'<div class="error-item">⚠️ {e}</div>', unsafe_allow_html=True)
    else:
        st.success("Tất cả hợp lệ — Sẵn sàng tạo HĐ ✅")

# ─────────────────────────────────────────────
# GENERATE
# ─────────────────────────────────────────────
st.markdown("---")
btn_col, dl_col = st.columns([1, 1], gap="large")

with btn_col:
    generate = st.button(
        "✦  Tạo Hợp Đồng",
        use_container_width=True,
        disabled=bool(errors),
        type="primary",
    )

if generate:
    try:
        result_bytes = tao_mot_hop_dong(current_data, template_bytes)
        fname = fname_preview
        st.session_state.result_bytes = result_bytes
        st.session_state.result_fname = fname
        st.session_state.generated    = True
        st.session_state.history.append({
            "fname": fname,
            "time":  datetime.now().strftime("%d/%m/%Y %H:%M"),
        })
        st.success(f"✅ Tạo thành công: **{fname}**")
    except Exception as ex:
        st.error(f"❌ Lỗi: {ex}")
        st.session_state.generated = False

if st.session_state.generated and st.session_state.result_bytes:
    with dl_col:
        st.download_button(
            label=f"⬇️  Tải xuống  {st.session_state.result_fname}",
            data=st.session_state.result_bytes,
            file_name=st.session_state.result_fname,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("""
<hr style="margin-top:36px;border-color:#e2e8f0">
<div style="text-align:center;color:#94a3b8;font-size:0.77rem;padding:10px 0">
    Tool Tạo Hợp Đồng · MobiFone Invoice v7 Web · Streamlit + python-docx
</div>
""", unsafe_allow_html=True)
