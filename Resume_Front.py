import os
import streamlit as st
from dotenv import load_dotenv
from Resume_Back import font_base64, logo_base64
from Resume_Back import ResumeEvaluatorBot
from Resume_Back import extract_text_from_pdf, extract_score, extract_candidate_name
from Resume_Back import build_resume_report, style_excel_sheet
load_dotenv()


# =============================================
# Streamlit Page Config
# =============================================
st.set_page_config(
    page_title="سامانه ارزیابی هوشمند رزومه‌های متقاضیان استخدام",
    layout="wide",
)
APP_VERSION = "نسخه 0.1.0"

# Brand palette
PURPLE = "#782DBE"
BLUE   = "#0064B2"
GOLD   = "#BF9000"
NAVY   = "#1F3864"   # رنگ سرمه‌ای برای فیلدهای اطلاعات شغلی

# =============================================
# Global Style (light theme) + Hero
# =============================================
st.markdown(f"""
<style>
@font-face {{
    font-family: 'BKoodkBD';
    src: url(data:font/ttf;base64,{font_base64}) format('truetype');
}}
* {{ margin:0; padding:0; box-sizing:border-box; }}

html, body, [class*="css"], .stApp, .stApp * {{
    font-family: 'BKoodkBD', Cambria, sans-serif !important;
    direction: rtl !important;
}}

/* منوی باز-شونده‌ی گزینه‌های select (مثل «مدت سابقه کاری») در یک پرتال جدا و بیرون از .stApp
   رندر می‌شود، پس قانون بالا به آن نمی‌رسد و باید جداگانه هدف‌گیری شود */
[data-baseweb="popover"], [data-baseweb="popover"] *,
[data-baseweb="menu"], [data-baseweb="menu"] *,
[data-testid="stSelectboxVirtualDropdown"], [data-testid="stSelectboxVirtualDropdown"] * {{
    font-family: 'BKoodkBD', Cambria, sans-serif !important;
}}
[data-baseweb="menu"] li, [data-testid="stSelectboxVirtualDropdown"] li {{
    direction: rtl !important;
    font-size:0.9rem;
    text-align: right !important;
}}

.stApp {{
    background: #F7F7FB !important;
    min-height: 100vh;
    color: #1F2430;
}}

header[data-testid="stHeader"] {{ background: transparent !important; border-bottom: none; }}
#MainMenu, footer, [data-testid="stToolbar"] {{ visibility: hidden; }}
footer[data-testid] {{ display:none; }}
[data-testid="InputInstructions"] {{ display:none !important; }}

.block-container {{
    padding-top: 0.6rem !important;
    padding-bottom: 3rem !important;
    max-width: 1120px !important;
    margin: 0 auto;
}}

/* ---------- Hero : لوگو سمت راست / عنوان در بالا ---------- */
.hero {{ display:flex; justify-content:center; padding: 8px 0 22px; }}
.hero-box {{
    background:#fff;
    border:1px solid #ECEAF3;
    border-top:4px solid {PURPLE};
    border-radius:18px;
    padding:18px 32px;
    box-shadow:0 6px 24px rgba(31,36,48,0.06);
    width:100%;
}}
.hero-row {{
    display:flex;
    flex-direction:row;
    align-items:center;
    justify-content:space-between;
    gap:24px;
}}
.hero-logo {{ flex:0 0 auto; order:1; }}      /* در RTL اولین آیتم سمت راست قرار می‌گیرد */
.hero-titles {{ flex:1 1 auto; order:2; text-align:center; }}
.hero-box h1 {{
    font-size:1.6rem;
    font-weight:800;
    color:#1F2430;
    line-height:1.7;
    margin:0;
}}
.hero-box h1 span {{ color:{PURPLE}; }}
.hero-sub {{ margin-top:6px; color:#7A8090; font-size:0.98rem; }}

/* ---------- Steps ---------- */
.steps {{ display:flex; justify-content:center; align-items:center; margin: 4px 0 22px; }}
.step {{ display:flex; align-items:center; gap:8px; font-size:1rem; color:#9AA0AE; font-weight:700; }}
.step-num {{
    width:30px; height:30px; border-radius:50%;
    display:flex; align-items:center; justify-content:center;
    font-weight:800; background:#EFEFF5; color:#9AA0AE; border:1px solid #E2E1EC;
}}
.step.active {{ color:{PURPLE}; }}
.step.active .step-num {{ background:{PURPLE}; border-color:{PURPLE}; color:#fff; }}
.step.done {{ color:{BLUE}; }}
.step.done .step-num {{ background:{BLUE}; border-color:{BLUE}; color:#fff; }}
.step-line {{ width:90px; height:2px; background:#E2E1EC; margin:0 10px; border-radius:2px; }}

/* ---------- Cards & titles ---------- */
.card {{
    background:#fff;
    border:1px solid #ECEAF3;
    border-radius:16px;
    padding:18px 20px 20px;
    box-shadow:0 4px 18px rgba(31,36,48,0.05);
    margin-bottom:16px;
}}
.section-title {{
    display:flex; align-items:center; gap:10px;
    font-size:1.1rem; font-weight:800; color:#1F2430;
    margin:0 0 14px;
}}
.section-title::before {{
    content:''; width:6px; height:20px; border-radius:4px;
    background:linear-gradient(180deg,{PURPLE},{BLUE});
}}
.section-title::after {{ content:''; flex:1; height:1px; background:#EFEFF5; }}

/* Streamlit's own bordered containers used as cards */
[data-testid="stVerticalBlockBorderWrapper"] {{
    background:#fff;
    border-radius:16px !important;
    border:1px solid #ECEAF3 !important;
    box-shadow:0 4px 18px rgba(31,36,48,0.05);
}}

/* ---------- هم‌ارتفاع کردن ردیف «اطلاعات موقعیت شغلی» و «وزن‌دهی معیارها» ---------- */
.row-equal-height + div[data-testid="stHorizontalBlock"] {{
    align-items: stretch !important;
}}
.row-equal-height + div[data-testid="stHorizontalBlock"] > div[data-testid="column"] {{
    display:flex !important;
}}
.row-equal-height + div[data-testid="stHorizontalBlock"] > div[data-testid="column"] > div {{
    display:flex !important;
    flex-direction:column !important;
    width:100% !important;
}}
.row-equal-height + div[data-testid="stHorizontalBlock"] [data-testid="stVerticalBlockBorderWrapper"] {{
    height:100% !important;
}}

/* ---------- Inputs ---------- */
label[data-testid="stWidgetLabel"] p {{
    color:#4A5060 !important;
    font-size:1rem !important;
    font-weight:700 !important;
}}
.stTextInput input,
.stTextArea textarea,
[data-baseweb="select"] > div {{
    background:#FBFBFD !important;
    border:1px solid #E2E1EC !important;
    border-radius:12px !important;
    color:#1F2430 !important;
    font-family:'BKoodkBD' !important;
}}
.stTextInput input:focus,
.stTextArea textarea:focus {{
    border-color:{PURPLE} !important;
    background:#fff !important;
    box-shadow:0 0 0 3px rgba(120,45,190,0.12) !important;
}}
.stTextArea textarea {{ font-size:0.98rem !important; line-height:1.9 !important; }}

/* ---------- استایل متمایز فیلدهای «اطلاعات موقعیت شغلی» (حاشیه سرمه‌ای) ---------- */
.job-info-scope .stTextInput input,
.job-info-scope .stTextArea textarea,
.job-info-scope [data-baseweb="select"] > div {{
    border:2px solid {NAVY} !important;
    background:#FFFFFF !important;
    border-radius:12px !important;
}}
.job-info-scope .stTextInput input:focus,
.job-info-scope .stTextArea textarea:focus {{
    border-color:{NAVY} !important;
    box-shadow:0 0 0 3px rgba(31,56,100,0.15) !important;
}}
.job-info-scope label[data-testid="stWidgetLabel"] p {{
    color:{NAVY} !important;
}}

/* file uploader */
[data-testid="stFileUploader"] {{
    border:2px dashed rgba(120,45,190,0.35);
    border-radius:14px;
    background:rgba(120,45,190,0.04);
    padding:10px;
}}
[data-testid="stFileUploader"]:hover {{
    border-color:{PURPLE};
    background:rgba(120,45,190,0.08);
}}
[data-testid="stFileUploader"] * {{ color:#4A5060 !important; }}

/* weights sum badge */
.weights-sum {{
    margin-top:auto;
    padding-top:10px;
    font-size:0.9rem;
    font-weight:700;
    color:#7A8090;
    text-align:center;
}}
.weights-sum b {{ color:{PURPLE}; }}

/* sliders */
.stSlider [data-baseweb="slider"] {{ padding:0 4px !important; }}
.stSlider [data-testid="stThumbValue"] {{ color:{PURPLE} !important; font-weight:800 !important; }}
.stSlider [data-testid="stTickBarMin"], .stSlider [data-testid="stTickBarMax"] {{ color:#B0B5C2 !important; }}
.stSlider, .stSlider *, [data-baseweb="slider"], [data-baseweb="slider"] * {{ direction: ltr !important; }}

/* ---------- Buttons ---------- */
.stButton > button {{
    width:100%;
    padding:13px 14px;
    border-radius:12px;
    font-size:1rem;
    font-weight:800;
    font-family:'BKoodkBD' !important;
    transition:transform .15s, box-shadow .2s, background .2s;
}}
.stButton > button[kind="primary"] {{
    border:none !important;
    background:linear-gradient(135deg,{PURPLE},{BLUE}) !important;
    color:#fff !important;
    box-shadow:0 6px 16px rgba(120,45,190,0.25);
}}
.stButton > button[kind="primary"]:hover {{ transform:translateY(-1px); box-shadow:0 10px 22px rgba(120,45,190,0.32); }}
.stButton > button[kind="secondary"] {{
    background:#fff !important;
    color:{PURPLE} !important;
    border:1.5px solid #E2E1EC !important;
}}
.stButton > button[kind="secondary"]:hover {{ border-color:{PURPLE} !important; background:rgba(120,45,190,0.05) !important; }}

/* ---------- Score ring (وسط‌چین) ---------- */
.score-block {{
    display:flex;
    flex-direction:column;
    align-items:center;
    justify-content:center;
    gap:14px;
    padding:6px 0 2px;
}}
.score-ring {{ position:relative; width:170px; height:170px; }}
.score-ring svg {{ transform:rotate(-90deg); }}
.score-ring circle {{ fill:none; stroke-width:12; }}
.ring-bg {{ stroke:#EFEFF5; }}
.ring-fill {{ stroke:url(#scoreGrad); stroke-linecap:round; stroke-dasharray:345; transition:stroke-dashoffset 1s ease; }}
.score-text {{ position:absolute; inset:0; display:flex; flex-direction:column; align-items:center; justify-content:center; }}
.score-num {{ font-size:2.4rem; font-weight:800; line-height:1; }}
.score-lbl {{ font-size:0.85rem; color:#9AA0AE; margin-top:4px; }}

/* ---------- Verdict (زیر دایره نمره) ---------- */
.verdict-wrap {{ text-align:center; }}
.verdict {{
    display:inline-flex; align-items:center; gap:8px;
    padding:9px 22px; border-radius:30px;
    font-size:0.98rem; font-weight:800;
}}
.verdict.good {{ background:rgba(0,100,178,0.10); border:1px solid rgba(0,100,178,0.28); color:{BLUE}; }}
.verdict.mid  {{ background:rgba(191,144,0,0.12); border:1px solid rgba(191,144,0,0.30); color:{GOLD}; }}
.verdict.low  {{ background:rgba(214,48,49,0.10); border:1px solid rgba(214,48,49,0.28); color:#C0392B; }}

/* ---------- Disabled result textareas ---------- */
.stTextArea[data-testid] textarea[disabled] {{
    background:#FBFBFD !important;
    color:#2B303C !important;
    border:1px solid #ECEAF3 !important;
    opacity:1 !important;
    -webkit-text-fill-color:#2B303C !important;
}}

/* ---------- Misc ---------- */
.app-version {{
    position:fixed; bottom:12px; left:16px;
    font-size:0.85rem; color:#B0B5C2; z-index:10;
}}
[data-testid="stAlert"] {{ border-radius:12px; }}
</style>
<div class="app-version">{APP_VERSION}</div>

<div class="hero">
  <div class="hero-box">
    <div class="hero-row">
      <div class="hero-logo">
        <img src="data:image/png;base64,{logo_base64}" width="120">
      </div>
      <div class="hero-titles">
        <h1>سامانه ارزیابی هوشمند <span>رزومه‌های متقاضیان استخدام</span></h1>
        <div class="hero-sub">تحلیل مطابقت رزومه با شایستگی‌های مدنظر موقعیت شغلی</div>
      </div>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

# =============================================
# Session State
# =============================================
if "bot" not in st.session_state:
    API_KEY = os.getenv("OPENAI_API_KEY")
    BASE_URL = os.getenv("OPENAI_API_BASE")
    if not API_KEY or not BASE_URL:
        st.error("عدم اتصال به سرویس هوش مصنوعی.")
        st.stop()
    st.session_state.bot = ResumeEvaluatorBot(api_key=API_KEY, base_url=BASE_URL)

if "last_result" not in st.session_state:
    st.session_state.last_result = ""

# اعمال ضرایب نرمال‌شده روی اسلایدرها؛ این کار باید همین‌جا و پیش از ساخته‌شدن خود اسلایدرها انجام شود،
# چون Streamlit اجازه نمی‌دهد مقدار session_state یک ویجت بعد از ساخته‌شدنش در همان اجرا تغییر کند.
if "_pending_weights" in st.session_state:
    _pw = st.session_state.pop("_pending_weights")
    st.session_state["w_tech"] = _pw["tech"]
    st.session_state["w_exp"] = _pw["exp"]
    st.session_state["w_edu"] = _pw["edu"]
    st.session_state["w_proj"] = _pw["proj"]

# =============================================
# Step indicator (dynamic)
# =============================================
step1_cls = "active"
step2_cls = ""
step3_cls = ""
if st.session_state.last_result:
    step1_cls = "done"
    step2_cls = "done"
    step3_cls = "active"

if "batch_results" not in st.session_state:
    st.session_state.batch_results = []

st.markdown(f"""
<div class="steps">
  <div class="step {step1_cls}"><div class="step-num">۱</div><span>تعریف موقعیت شغلی</span></div>
  <div class="step-line"></div>
  <div class="step {step2_cls}"><div class="step-num">۲</div><span>بارگذاری رزومه</span></div>
  <div class="step-line"></div>
  <div class="step {step3_cls}"><div class="step-num">۳</div><span>نتیجه ارزیابی</span></div>
</div>
""", unsafe_allow_html=True)

# =============================================
# Helper: normalize criteria weights to sum to 100
# =============================================
def normalize_weights(raw: dict) -> dict:
    """
    ضرایب را طوری تغییر مقیاس می‌دهد که مجموعشان دقیقاً ۱۰۰ شود، بدون اینکه نسبت آن‌ها نسبت به هم تغییر کند.
    اگر مجموع ورودی صفر باشد (نسبتی برای حفظ کردن وجود ندارد)، همان مقادیر بدون تغییر برگردانده می‌شود.
    """
    total = sum(raw.values())
    if total == 0:
        return dict(raw)
    scaled = {k: round(v * 100 / total) for k, v in raw.items()}
    diff = 100 - sum(scaled.values())
    if diff:
        k_max = max(scaled, key=scaled.get)
        scaled[k_max] = max(0, min(100, scaled[k_max] + diff))
    return scaled


# =============================================
# Input Area : right = job info | left = weights ; upload below (full width)
# =============================================
st.markdown('<div class="row-equal-height"></div>', unsafe_allow_html=True)
right_col, left_col = st.columns([1.15, 1], gap="large")

# ---------- Section 1: Job Info (استایل متمایز سرمه‌ای) ----------
with right_col:
    with st.container(border=True):
        st.markdown('<div class="section-title">اطلاعات موقعیت شغلی</div>', unsafe_allow_html=True)
        st.markdown('<div class="job-info-scope">', unsafe_allow_html=True)

        jc1, jc2 = st.columns(2)
        with jc1:
            job_title = st.text_input("عنوان شغل", placeholder="مثال: کارشناس تحقیق و توسعه")
        with jc2:
            experience = st.selectbox(
                "مدت سابقه کاری",
                ["کمتر از ۲ سال", "بین ۲ تا ۵ سال", "بین ۵ تا ۱۰ سال", "بالای ۱۰ سال"],
                index=None,
                placeholder="یک گزینه را انتخاب کنید ..."
            )

        competencies_text = st.text_area(
            "شایستگی‌های شغلی مدنظر کارفرما",
            height=137,
            placeholder="شایستگی‌ها، مهارت‌ها و الزامات کلیدی موقعیت شغلی را وارد نمایید ..."
        )

        st.markdown('</div>', unsafe_allow_html=True)

with left_col:
    # ---------- Section 2: Weights (بدون جمع وزن‌ها) ----------
    with st.container(border=True):
        st.markdown('<div class="section-title">ضرایب معیارهای شایستگی</div>', unsafe_allow_html=True)

        wc1, wc2 = st.columns(2)
        with wc1:
            w_tech = st.slider("مهارت فنی", 0, 100, 25, format="%d%%", key="w_tech")
            w_edu  = st.slider("تحصیلات",        0, 100, 25, format="%d%%", key="w_edu")
        with wc2:
            w_exp  = st.slider("مدت سابقه کاری",     0, 100, 25, format="%d%%", key="w_exp")
            w_proj = st.slider("پروژه‌های انجام‌داده",        0, 100, 25, format="%d%%", key="w_proj")

        weights = {"tech": w_tech, "exp": w_exp, "edu": w_edu, "proj": w_proj}
        st.markdown(
            f'<div class="weights-sum">مجموع ضرایب: <b>{sum(weights.values())}%</b></div>',
            unsafe_allow_html=True
        )

        normalize_clicked = st.button(
            "نرمال‌سازی ضرایب", use_container_width=True, key="normalize_btn", type="secondary"
        )
        if normalize_clicked:
            if sum(weights.values()) == 0:
                st.warning("برای نرمال‌سازی، ابتدا حداقل یکی از ضرایب را بیشتر از صفر کنید.")
            else:
                st.session_state["_pending_weights"] = normalize_weights(weights)
                st.rerun()

# ---------- Section 3: Upload (سراسری — هم‌عرض با دکمه ارزیابی) ----------
with st.container(border=True):
    st.markdown('<div class="section-title">بارگذاری رزومه</div>', unsafe_allow_html=True)
    st.markdown("""
    <style>
    [data-testid="stFileUploaderDropzoneInstructions"] span,
    [data-testid="stFileUploaderDropzoneInstructions"] small { display: none !important; }

    [data-testid="stFileUploaderDropzoneInstructions"] > div::after {
        content: "لطفاً فایل‌های رزومه را بارگذاری نمایید (حداکثر 50 فایل)";
        display: block;
        direction: rtl;
        text-align: right;
        font-size: 1rem;
        font-weight: 600;
        color: #4A5060;
        line-height: 1.9;
    }

    [data-testid="stFileUploaderFileList"],
    [data-testid="stFileUploaderFile"],
    [data-testid="stFileUploaderPagination"],
    [data-testid="stFileUploaderDropzone"] ~ div { display: none !important; }
    </style>
    """, unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "فایل‌های رزومه را بارگذاری کنید (حداکثر ۵۰ فایل)",
        label_visibility="collapsed",
        type=["pdf"],
        accept_multiple_files=True,
        key="resume_uploader"
    )
    if uploaded_files:
        if len(uploaded_files) > 50:
            st.warning("حداکثر ۵۰ فایل قابل بارگذاری است. فقط ۵۰ فایل اول پردازش می‌شوند.")
            uploaded_files = uploaded_files[:50]
        st.success(f"✅ {len(uploaded_files)} فایل رزومه با موفقیت بارگذاری شد ...")

# =============================================
# Action Buttons
# =============================================
evaluate_clicked = st.button("ارزیابی", use_container_width=True, key="eval_btn", type="primary")

# =============================================
# Evaluation
# =============================================
if evaluate_clicked:
    if not uploaded_files:
        st.error("لطفاً حداقل یک فایل رزومه بارگذاری کنید.")
    elif not competencies_text.strip():
        st.error("لطفاً شایستگی‌های مورد نظر را وارد کنید.")
    elif sum(weights.values()) == 0:
        st.error("مجموع ضرایب معیارها نباید صفر باشد. لطفاً حداقل یک معیار را وزن‌دهی کنید.")
    else:
        # اگر کاربر فراموش کرده باشد دکمه‌ی «نرمال‌سازی ضرایب» را بزند، همین‌جا به‌صورت خودکار انجام می‌شود.
        # مقدار اسلایدرها همین الان تغییر نمی‌کند (چون در همین اجرا قبلاً ساخته شده‌اند)؛ بلکه به‌عنوان
        # «در انتظار» ذخیره می‌شود تا با rerun پایان ارزیابی، در ابتدای اجرای بعدی روی اسلایدرها اعمال شود.
        weights = normalize_weights(weights)
        st.session_state["_pending_weights"] = dict(weights)

        results = []
        # progress_bar = st.progress(0)
        status_text = st.empty()
        total = len(uploaded_files)

        for i, uploaded_file in enumerate(uploaded_files):
            status_text.text(f"در حال ارزیابی رزومه {i+1} از {total}")
            resume_text = extract_text_from_pdf(uploaded_file)
            if not resume_text.strip():
                results.append({"file_name": uploaded_file.name, "score_raw": -1, "result_text": "", "applicant_name": "-"})
            else:
                result_text = st.session_state.bot.evaluate(resume_text, competencies_text, weights)
                score = extract_score(result_text, weights)
                applicant_name = extract_candidate_name(result_text)
                try:
                    score_val = int(score)
                except (ValueError, TypeError):
                    score_val = -1
                results.append({"file_name": uploaded_file.name, "score_raw": score_val, "result_text": result_text, "applicant_name": applicant_name})
            # progress_bar.progress((i + 1) / total)

        status_text.text("ارزیابی تمام رزومه‌ها به پایان رسید.")
        st.session_state.batch_results = results
        st.session_state.last_result = "batch_done"
        st.rerun()

# =============================================
# Result Display
# =============================================
if st.session_state.get("last_result") == "batch_done" and st.session_state.get("batch_results"):
    import pandas as pd
    from io import BytesIO

    batch = st.session_state.batch_results

    # جداسازی رزومه‌های قابل پردازش و ناموفق
    valid = [r for r in batch if r["score_raw"] >= 0]
    invalid = [r for r in batch if r["score_raw"] < 0]

    # مرتب‌سازی نزولی بر اساس نمره
    valid_sorted = sorted(valid, key=lambda x: x["score_raw"], reverse=True)

    # ساخت DataFrame برای نمایش و Excel
    rows = []
    for rank, r in enumerate(valid_sorted, start=1):
        rows.append({
            "رتبه": rank,
            "نام متقاضی": r.get("applicant_name", "-"),
            "نام فایل رزومه": r["file_name"],
            "نمره ارزیابی": r["score_raw"]
        })

    df = pd.DataFrame(rows)

    st.markdown("ارزیابی با موفقیت انجام شد ...", text_alignment = "right")
    # st.dataframe(df, use_container_width=True, hide_index=True)

    if invalid:
        st.warning(f"⚠️ {len(invalid)} متن رزومه قابل استخراج نبود: " + "، ".join([r['file_name'] for r in invalid]))

    # خروجی Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="نتایج ارزیابی")
        style_excel_sheet(writer.sheets["نتایج ارزیابی"])
    output.seek(0)

    dl_col1, dl_col2 = st.columns(2)
    with dl_col1:
        st.download_button(
            label="دانلود نتایج رتبه‌بندی",
            data=output,
            file_name="resume_evaluation_results.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

    # خروجی Word (گزارش تحلیلی) - ترتیب رزومه‌ها دقیقاً مطابق رتبه‌ها در Excel
    with dl_col2:
        try:
            report_docx = build_resume_report(valid_sorted)
            st.download_button(
                label="دانلود گزارش تحلیلی",
                data=report_docx,
                file_name="resume_evaluation_report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
        except Exception as e:
            st.error(f"ساخت گزارش با خطا مواجه شد: {e}")
