import streamlit as st
import pandas as pd
import sqlite3
import json
import hashlib
from datetime import datetime
import os

# تنظیمات اولیه صفحه Streamlit
st.set_page_config(
    page_title="سامانه مدیریت و بایگانی ساخت افزایشی (SLM)",
    page_icon="logo.png" if os.path.exists("logo.png") else "⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- استایل یکپارچه، تمیز و کاملاً راست‌چین (RTL) ---
st.markdown("""
<style>
    @font-face {
        font-family: 'B Nazanin';
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot');
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot?#iefix') format('embedded-opentype'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.woff') format('woff'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.ttf') format('truetype');
        font-weight: normal;
        font-style: normal;
    }

    /* ۱. راست‌چین‌سازی کامل کل برنامه، متون و فرم ورود */
    html, body, .stApp, div, p, span, label, input, select, button, textarea {
        font-family: 'B Nazanin', 'Vazir', sans-serif !important;
        direction: rtl !important;
        text-align: right !important;
    }

    p, span, label, input, select, button {
        font-size: 1.2rem !important;
    }
    
    h1, h2, h3, h4, h5, h6 {
        font-family: 'B Nazanin', 'Vazir' !important;
        direction: rtl !important;
        text-align: right !important;
        font-weight: bold;
    }

    h1 { font-size: 2.2rem !important; }
    h2 { font-size: 1.8rem !important; }
    h3 { font-size: 1.5rem !important; }

    /* ۲. مخفی‌سازی متون راهنمای اضافی داخل کادرها */
    [data-testid="stInputInstruction"], 
    input::placeholder,
    textarea::placeholder {
        display: none !important;
        opacity: 0 !important;
    }

    /* ۳. اصلاح آیکون‌های سیستم جهت جلوگیری از نمایش متن کد لاتین آن‌ها */
    [data-testid="stIconMaterial"], .aria-hidden, i, [class^="st-"] {
        font-family: 'Material Symbols Outlined', 'Material Icons' !important;
        direction: ltr !important;
    }

    /* ۴. راست‌چین کردن کامل و اجباری تمامی جداول (Headers & Data Cells) */
    table, [data-testid="stTable"], .stTable {
        direction: rtl !important;
        width: 100% !important;
        font-size: 1.15rem !important;
        border-collapse: collapse !important;
        text-align: right !important;
    }
    
    th, td, table th, table td {
        text-align: right !important;
        direction: rtl !important;
        padding: 10px !important;
    }

    /* ۵. استایل‌دهی دکمه‌های منوی سایدبار */
    [data-testid="stSidebar"][aria-expanded="false"] [data-testid="stSidebarUserContent"] {
        display: none !important;
    }

    [data-testid="stSidebar"][aria-expanded="true"] [data-testid="stSidebarUserContent"] {
        direction: rtl !important;
        text-align: right !important;
    }

    div[data-testid="stSidebar"] [data-testid="stRadio"] > div {
        display: flex;
        flex-direction: column;
        gap: 8px;
    }

    div[data-testid="stSidebar"] [data-testid="stRadio"] label {
        background-color: transparent !important;
        border-radius: 8px !important;
        padding: 8px 16px !important;
        cursor: pointer !important;
        transition: all 0.2s ease !important;
        border: 1px solid transparent !important;
    }

    div[data-testid="stSidebar"] [data-testid="stRadio"] label:hover {
        background-color: #2a374e !important;
        border: 1px solid #3b82f6 !important;
    }

    div[data-testid="stSidebar"] [data-testid="stRadio"] label[data-checked="true"] {
        background-color: #334155 !important;
        border: 1px solid #475569 !important;
        box-shadow: 0 2px 4px rgba(0,0,0,0.3) !important;
    }

    div[data-testid="stSidebar"] [data-testid="stRadio"] input[type="radio"] {
        display: none !important;
    }

    /* ۶. کارت‌های شاخص آمار */
    [data-testid="stMetric"] {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
        padding: 15px !important;
        text-align: center !important;
    }
    [data-testid="stMetricLabel"] > div {
        color: #94a3b8 !important;
        font-size: 1.2rem !important;
        justify-content: center !important;
    }
    [data-testid="stMetricValue"] > div {
        color: #38bdf8 !important;
        font-size: 2.2rem !important;
        justify-content: center !important;
    }

    .stButton>button {
        width: 100%;
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        font-size: 1.3rem !important;
        padding: 10px 20px !important;
        border: none !important;
    }
</style>
""", unsafe_allow_html=True)

FARSI_HEADERS_MAP = {
    'powder_code': 'شماره ظرف پودر',
    'material': 'جنس / نوع متریال پودر',
    'weight_g': 'وزن پودر (گرم)',
    'date': 'تاریخ ثبت',
    'id': 'شناسه',
    'recycled_batch_code': 'کد پارت بازیافت',
    'input_powder_g': 'پودر ورودی به دستگاه (گرم)',
    'unrecyclable_powder_g': 'پودر غیرقابل بازیافت (گرم)',
    'recycled_powder_g': 'پودر بازیافتی قابل استفاده (گرم)',
    'notes': 'توضیحات و ملاحظات',
    'part_code': 'کد/شناسه قطعه',
    'part_name': 'نام قطعه',
    'quantity': 'تعداد روی صفحه',
    'machine_model': 'مدل دستگاه',
    'build_time_hrs': 'زمان تولید (ساعت)',
    'downtime_hrs': 'زمان توقف (ساعت)',
    'start_date': 'تاریخ شروع',
    'start_time': 'ساعت شروع',
    'end_date': 'تاریخ پایان',
    'end_time': 'ساعت پایان',
    'setup_time_hrs': 'زمان آماده‌سازی (ساعت)',
    'cleaning_time_hrs': 'زمان تمیزکاری (ساعت)',
    'waste_powder_g': 'پودر غیرقابل بازیافت (گرم)',
    'part_with_support_g': 'وزن با ساپورت (گرم)',
    'final_part_g': 'وزن قطعه نهایی (گرم)',
    'filter_percentage': 'درصد فیلتر دستگاه (%)',
    'build_plate_code': 'کد صفحه ساخت',
    'build_plate_init_wt_g': 'وزن اولیه صفحه ساخت (گرم)',
    'build_plate_post_wt_g': 'وزن صفحه ساخت بعد پرداخت (گرم)',
    'engraving_qty': 'تعداد حکاکی لیزر',
    'delivery_date': 'تاریخ تحویل',
    'qc_inspector': 'بازرس کنترل کیفیت',
    'qc_engineer': 'مسئول مهندسی کیفیت',
    'qa_manager': 'مدیر تضمین کیفیت',
    'powder_type': 'نوع پودر فلزی',
    'volume_cm3': 'حجم قطعه (cm3)',
    'net_weight_g': 'وزن خالص (گرم)',
    'support_volume_cm3': 'حجم ساپورت (cm3)',
    'support_weight_g': 'وزن ساپورت (گرم)',
    'machine_type': 'نوع دستگاه',
    'parts_on_plate': 'تعداد قطعات روی صفحه',
    'print_time_hrs': 'زمان چاپ (ساعت)',
    'design_time_hrs': 'زمان طراحی (ساعت)',
    'post_process_time_hrs': 'زمان پرداخت (ساعت)',
    'overhead_pct': 'ضریب سربار (%)',
    'powder_cost_total': 'هزینه پودر (ریال)',
    'argon_cost_total': 'هزینه گاز آرگون (ریال)',
    'depreciation_cost_total': 'هزینه استهلاک دستگاه (ریال)',
    'power_cost_total': 'هزینه برق (ریال)',
    'engineering_cost_total': 'هزینه طراحی/مهندسی (ریال)',
    'operator_cost_total': 'هزینه اپراتور (ریال)',
    'post_process_cost_total': 'هزینه پرداخت‌کاری (ریال)',
    'qc_cost_total': 'هزینه کنترل کیفیت (ریال)',
    'utility_ventilation': 'هزینه تهویه (ریال)',
    'utility_chiller': 'هزینه چیلر (ریال)',
    'total_production_cost': 'بهای تمام شده کل (ریال)',
    'overhead_cost': 'مبلغ سربار (ریال)',
    'final_price': 'قیمت نهایی قابل ارائه به مشتری (ریال)'
}

DB_NAME = "slm_management.db"

DEFAULT_RATES = {
    "powder_price_Steel_316": 120000000,
    "powder_price_Ti6Al4V": 500000000,
    "powder_price_Inconel_718": 300000000,
    "powder_price_Hastelloy_X": 400000000,
    "machine_depr_M120": 1250000,
    "machine_depr_M300": 3125000,
    "power_kw_M120": 7,
    "power_kw_M300": 15,
    "wage_designer": 2500000,
    "wage_operator": 1870000,
    "wage_qc": 2180000,
    "argon_rate": 300000,
    "electricity_rate": 50000,
    "post_process_rate": 500000,
    "qc_fixed_cost": 40000000,
    "ventilation_rate": 200000,
    "chiller_rate": 2500
}

def hash_password(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    c.execute('''CREATE TABLE IF NOT EXISTS users (
                    username TEXT PRIMARY KEY,
                    password_hash TEXT,
                    full_name TEXT,
                    role TEXT
                )''')
    
    default_users = [
        ('admin', hash_password('admin123'), 'مدیر سیستم', 'مدیریت'),
        ('operator', hash_password('op123'), 'اپراتور و طراح', 'اپراتور و طراح'),
        ('qc_user', hash_password('qc123'), 'کارشناس QC', 'کنترل کیفیت'),
        ('commerce', hash_password('com123'), 'کارشناس بازرگانی', 'بازرگانی')
    ]
    for u, p, f, r in default_users:
        c.execute("INSERT OR IGNORE INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)", (u, p, f, r))

    c.execute('''CREATE TABLE IF NOT EXISTS system_settings (
                    key TEXT PRIMARY KEY,
                    value REAL
                )''')
    
    for key, val in DEFAULT_RATES.items():
        c.execute("INSERT OR IGNORE INTO system_settings (key, value) VALUES (?, ?)", (key, val))

    c.execute('''CREATE TABLE IF NOT EXISTS powders (
                    powder_code TEXT PRIMARY KEY,
                    material TEXT,
                    weight_g REAL,
                    date TEXT,
                    checklist_json TEXT
                )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS recycled_powders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    powder_code TEXT,
                    recycled_batch_code TEXT,
                    input_powder_g REAL,
                    unrecyclable_powder_g REAL,
                    recycled_powder_g REAL,
                    date TEXT,
                    notes TEXT,
                    FOREIGN KEY(powder_code) REFERENCES powders(powder_code)
                )''')

    c.execute('''CREATE TABLE IF NOT EXISTS nora_powders (
                    powder_code TEXT PRIMARY KEY,
                    material TEXT,
                    weight_g REAL,
                    date TEXT
                )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS production (
                    part_code TEXT PRIMARY KEY,
                    part_name TEXT,
                    powder_code TEXT,
                    quantity INTEGER,
                    machine_model TEXT,
                    date TEXT,
                    build_time_hrs REAL,
                    downtime_hrs REAL,
                    start_date TEXT,
                    start_time TEXT,
                    end_date TEXT,
                    end_time TEXT,
                    setup_time_hrs REAL,
                    cleaning_time_hrs REAL,
                    input_powder_g REAL,
                    waste_powder_g REAL,
                    part_with_support_g REAL,
                    final_part_g REAL,
                    filter_percentage REAL,
                    build_plate_code TEXT,
                    build_plate_init_wt_g REAL,
                    build_plate_post_wt_g REAL,
                    finishing_json TEXT,
                    engraving_qty INTEGER,
                    delivery_date TEXT,
                    FOREIGN KEY(powder_code) REFERENCES powders(powder_code)
                )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS qc (
                    part_code TEXT PRIMARY KEY,
                    part_name TEXT,
                    material TEXT,
                    machine_model TEXT,
                    date TEXT,
                    qc_checks_json TEXT,
                    qc_inspector TEXT,
                    qc_engineer TEXT,
                    qa_manager TEXT,
                    FOREIGN KEY(part_code) REFERENCES production(part_code)
                )''')
    
    c.execute('''CREATE TABLE IF NOT EXISTS cost_calculator (
                    part_code TEXT PRIMARY KEY,
                    part_name TEXT,
                    powder_type TEXT,
                    volume_cm3 REAL,
                    net_weight_g REAL,
                    support_volume_cm3 REAL,
                    support_weight_g REAL,
                    machine_type TEXT,
                    parts_on_plate INTEGER,
                    print_time_hrs REAL,
                    design_time_hrs REAL,
                    post_process_time_hrs REAL,
                    overhead_pct REAL,
                    powder_cost_total REAL,
                    argon_cost_total REAL,
                    depreciation_cost_total REAL,
                    power_cost_total REAL,
                    engineering_cost_total REAL,
                    operator_cost_total REAL,
                    post_process_cost_total REAL,
                    qc_cost_total REAL,
                    utility_ventilation REAL,
                    utility_chiller REAL,
                    total_production_cost REAL,
                    overhead_cost REAL,
                    final_price REAL,
                    FOREIGN KEY(part_code) REFERENCES production(part_code)
                )''')
    
    conn.commit()
    conn.close()

init_db()

def get_db_connection():
    return sqlite3.connect(DB_NAME)

def load_system_rates():
    conn = get_db_connection()
    df = pd.read_sql_query("SELECT * FROM system_settings", conn)
    conn.close()
    rates = DEFAULT_RATES.copy()
    for _, row in df.iterrows():
        rates[row['key']] = row['value']
    return rates

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.session_state["user_role"] = ""
    st.session_state["full_name"] = ""

def login_user(username, password):
    conn = get_db_connection()
    c = conn.cursor()
    c.execute("SELECT full_name, role FROM users WHERE username=? AND password_hash=?", (username, hash_password(password)))
    result = c.fetchone()
    conn.close()
    if result:
        st.session_state["authenticated"] = True
        st.session_state["username"] = username
        st.session_state["full_name"] = result[0]
        st.session_state["user_role"] = result[1]
        return True
    return False

def logout_user():
    st.session_state["authenticated"] = False
    st.session_state["username"] = ""
    st.session_state["user_role"] = ""
    st.session_state["full_name"] = ""
    st.rerun()

# ---------------------------------------------------------
# صفحه ورود (کاملاً راست‌چین)
# ---------------------------------------------------------
if not st.session_state["authenticated"]:
    col_login_1, col_login_2, col_login_3 = st.columns([1, 2, 1])
    with col_login_2:
        st.markdown("<br>", unsafe_allow_html=True)
        if os.path.exists("logo.png"):
            st.image("logo.png", width=160)
        
        st.title("🔐 ورود به سامانه ساخت افزایشی (SLM)")
        st.caption("لطفاً نام کاربری و رمز عبور خود را وارد کنید:")
        
        with st.form("login_form"):
            username_input = st.text_input("نام کاربری (Username)", placeholder="")
            password_input = st.text_input("رمز عبور (Password)", type="password", placeholder="")
            submit_login = st.form_submit_button("🔑 ورود به سامانه")
            
            if submit_login:
                if login_user(username_input, password_input):
                    st.success(f"خوش آمدید {st.session_state['full_name']}!")
                    st.rerun()
                else:
                    st.error("نام کاربری یا رمز عبور اشتباه است.")

# ---------------------------------------------------------
# برنامه اصلی پس از ورود
# ---------------------------------------------------------
else:
    role = st.session_state["user_role"]
    
    if os.path.exists("logo.png"):
        st.sidebar.image("logo.png", width=140)
        
    st.sidebar.markdown(f"👤 **کاربر:** {st.session_state['full_name']}")
    st.sidebar.markdown(f"مقام/سطح: `{role}`")
    if st.sidebar.button("🚪 خروج از حساب"):
        logout_user()
    st.sidebar.markdown("---")

    if role == "مدیریت":
        menu_options = [
            "🏠 خانه",
            "📦 پودر",
            "🏭 فرم تولید",
            "❇️ کنترل کیفیت",
            "💰 محاسبه‌گر هزینه",
            "⚙️ تنظیمات نرخ‌ها",
            "🔍 بایگانی",
            "👥 مدیریت کاربران"
        ]
    elif role == "اپراتور و طراح":
        menu_options = [
            "🏠 خانه",
            "📦 پودر",
            "🏭 فرم تولید"
        ]
    elif role == "کنترل کیفیت":
        menu_options = [
            "🏠 خانه",
            "❇️ کنترل کیفیت",
            "🔍 بایگانی"
        ]
    elif role == "بازرگانی":
        menu_options = [
            "🏠 خانه",
            "📦 پودر",
            "💰 محاسبه‌گر هزینه",
            "🔍 بایگانی"
        ]
    else:
        menu_options = ["🏠 خانه"]

    choice = st.sidebar.radio("", menu_options, index=0)

    # ---------------------------------------------------------
    # ۰. خانه
    # ---------------------------------------------------------
    if choice == "🏠 خانه":
        head_c1, head_col2 = st.columns([6, 1])
        with head_c1:
            st.title("🧩 سامانه بایگانی و مدیریت تولید قطعات چاپ سه بعدی (SLM)")
            st.caption("این سامانه فرم‌های «پودر»، «فرم تولید»، «کنترل کیفیت (QC)» و «محاسبه‌گر قیمت» را به هم مرتبط کرده و امکان بایگانی را فراهم می‌سازد.")
        with head_col2:
            if os.path.exists("logo.png"):
                st.image("logo.png", width=120)

        st.markdown("<br>", unsafe_allow_html=True)
        
        conn = get_db_connection()
        powder_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM powders", conn)['cnt'].values[0] + pd.read_sql_query("SELECT COUNT(*) as cnt FROM nora_powders", conn)['cnt'].values[0]
        prod_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM production", conn)['cnt'].values[0]
        qc_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM qc", conn)['cnt'].values[0]
        cost_count = pd.read_sql_query("SELECT COUNT(*) as cnt FROM cost_calculator", conn)['cnt'].values[0]
        conn.close()

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("رکوردهای پودر", powder_count)
        m2.metric("فرم‌های تولید", prod_count)
        m3.metric("فرم‌های کنترل کیفیت", qc_count)
        m4.metric("محاسبات هزینه", cost_count)

        st.markdown("---")
        st.subheader("راهنمای گردش کار")
        st.markdown("""
        1. **📦 پودر** — ثبت و آنالیز ظروف پودر خریداری شده اولیه، بازیافتی یا پودرهای نورا.
        2. **🏭 فرم تولید** — تعریف «کد قطعه» و اتصال آن به «کد ظرف پودر» و ثبت زمان‌های آماده‌سازی و فرآیند.
        3. **❇️ کنترل کیفیت (QC)** — ثبت نتایج تست‌های چشمی، ابعادی و مکانیکی برای کد قطعه.
        4. **💰 محاسبه‌گر هزینه** — برآورد خودکار هزینه‌های تولید و قیمت فروش نهایی بر اساس نرخ‌های روز.
        5. **🔍 بایگانی و جستجو** — استعلام شناسنامه جامع قطعات، مشاهده جداول دیتابیس و خروجی اکسل.
        
        <br>
        <p style='color: #94a3b8; font-size: 1.1rem;'>از منوی سمت راست بین صفحات جابه‌جا شوید.</p>
        """, unsafe_allow_html=True)

    # ---------------------------------------------------------
    # ۱. پودر
    # ---------------------------------------------------------
    elif choice == "📦 پودر":
        st.header("🧪 فرم مدیریت، آنالیز و بایگانی پودر")
        
        conn = get_db_connection()
        powders_df = pd.read_sql_query("SELECT * FROM powders", conn)
        nora_df = pd.read_sql_query("SELECT * FROM nora_powders", conn)
        
        sub_tab1, sub_tab2, sub_tab3 = st.tabs(["📦 ۱- پودرهای خریداری شده اولیه", "♻️ ۲- پودرهای بازیافت شده", "🏭 ۳- پودرهای خریداری شده از نورا"])
        
        with sub_tab1:
            with st.expander("➕ ثبت / ویرایش پودر جدید", expanded=True):
                col1, col2 = st.columns(2)
                with col1:
                    powder_code = st.text_input("کد/شماره ظرف پودر فلز (مانند: PWD-316-01)", placeholder="")
                    material = st.selectbox("جنس پودر", ["Steel 316", "Ti6Al4V", "Inconel 718", "Hastelloy X", "سایر"])
                with col2:
                    weight_g = st.number_input("وزن پودر (گرم)", min_value=0.0, value=10000.0)
                    date_str = st.date_input("تاریخ ورود/تست").strftime("%Y-%m-%d")
                    
                st.markdown("---")
                st.subheader("📋 چک‌لیست آزمون‌های خواص پودر")
                checklist = [
                    "بررسی خلوص شیمیایی", "اندازه‌گیری میزان رطوبت", "اندازه‌گیری چگالی ضربه‌ای",
                    "اندازه‌گیری چگالی ظاهری", "بررسی مورفولوژی ذرات", "آنالیز توزیع اندازه ذرات",
                    "وضعیت بسته‌بندی و عدم آلودگی", "بررسی گواهینامه کیفیت تامین‌کننده"
                ]
                
                qc_results = {}
                for item in checklist:
                    c1, c2, c3 = st.columns([3, 2, 4])
                    with c1: st.write(f"**{item}**")
                    with c2: status = st.selectbox("وضعیت", ["تایید", "رد"], key=item)
                    with c3: note = st.text_input("ملاحظات / توضیحات", key=f"note_{item}", placeholder="")
                    qc_results[item] = {"status": status, "note": note}
                    
                if st.button("💾 ذخیره در بایگانی پودر اولیه"):
                    if powder_code:
                        c = conn.cursor()
                        c.execute("""INSERT OR REPLACE INTO powders 
                                     (powder_code, material, weight_g, date, checklist_json)
                                     VALUES (?, ?, ?, ?, ?)""",
                                  (powder_code, material, weight_g, date_str, json.dumps(qc_results, ensure_ascii=False)))
                        conn.commit()
                        st.success(f"اطلاعات پودر {powder_code} با موفقیت ذخیره شد.")
                        st.rerun()
                    else:
                        st.error("لطفاً کد/شماره ظرف پودر را وارد کنید.")
                        
            st.markdown("---")
            st.subheader("📂 جدول پودرهای اولیه خریداری شده")
            if not powders_df.empty:
                disp_powders = powders_df[['powder_code', 'material', 'weight_g', 'date']].rename(columns=FARSI_HEADERS_MAP)
                st.table(disp_powders)
            else:
                st.info("هیچ پودر اولیه ای در سیستم ثبت نشده است.")

        with sub_tab2:
            st.subheader("♻️ مدیریت و ثبت پودرهای بازیافت شده")
            
            p_initial_list = [f"{code} [اولیه]" for code in powders_df['powder_code'].tolist()] if not powders_df.empty else []
            p_nora_list = [f"{code} [نورا]" for code in nora_df['powder_code'].tolist()] if not nora_df.empty else []
            combined_source_powders = p_initial_list + p_nora_list
            
            if not combined_source_powders:
                st.warning("ابتدا باید حداقل یک ظرف پودر اولیه یا پودر نورا ثبت کرده باشید.")
            else:
                with st.form("recycled_powder_form"):
                    rc1, rc2 = st.columns(2)
                    with rc1:
                        selected_source = st.selectbox("شماره/کد ظرف پودر مبدا (خریداری‌شده اولیه / نورا)", combined_source_powders)
                        clean_source_code = selected_source.split(" [")[0] if selected_source else ""
                        recycled_batch_code = st.text_input("شناسه پارت بازیافت (مانند: REC-PWD-01)", placeholder="")
                        input_powder_g = st.number_input("پودر ورودی به دستگاه (گرم)", min_value=0.0, value=5000.0)
                    with rc2:
                        unrecyclable_powder_g = st.number_input("پودر مصرف شده غیر قابل بازیافت (گرم)", min_value=0.0, value=200.0)
                        recycled_powder_g = input_powder_g - unrecyclable_powder_g
                        st.metric("مقدار پودر بازیافت‌شده قابل استفاده (گرم)", f"{recycled_powder_g:,.1f}")
                        rec_date = st.date_input("تاریخ بازیافت").strftime("%Y-%m-%d")
                        rec_notes = st.text_input("توضیحات و ملاحظات غربال‌گری / الک", placeholder="")
                    
                    rec_submit = st.form_submit_button("💾 ثبت رکورد پودر بازیافتی")
                    if rec_submit:
                        c = conn.cursor()
                        c.execute("""INSERT INTO recycled_powders 
                                     (powder_code, recycled_batch_code, input_powder_g, unrecyclable_powder_g, recycled_powder_g, date, notes)
                                     VALUES (?, ?, ?, ?, ?, ?, ?)""",
                                  (clean_source_code, recycled_batch_code, input_powder_g, unrecyclable_powder_g, recycled_powder_g, rec_date, rec_notes))
                        conn.commit()
                        st.success("اطلاعات پودر بازیافتی با موفقیت ثبت شد.")
                        st.rerun()

            st.markdown("---")
            st.subheader("📊 جدول بایگانی پودرهای بازیافت شده")
            recycled_df = pd.read_sql_query("SELECT * FROM recycled_powders", conn)
            
            if not recycled_df.empty:
                display_recycled_df = recycled_df.rename(columns=FARSI_HEADERS_MAP)
                st.table(display_recycled_df)
            else:
                st.info("هنوز رکوردی برای پودر بازیافت شده ثبت نشده است.")

        with sub_tab3:
            st.subheader("🏭 پودرهای خریداری شده از نورا (ثبت سریع)")
            
            with st.form("nora_powder_form"):
                nc1, nc2 = st.columns(2)
                with nc1:
                    nora_powder_code = st.text_input("کد/شماره ظرف پودر نورا (مانند: NORA-PWD-01)", placeholder="")
                    nora_material = st.selectbox("نوع متریال / جنس پودر", ["Steel 316", "Ti6Al4V", "Inconel 718", "Hastelloy X", "سایر"], key="nora_mat")
                with nc2:
                    nora_weight_g = st.number_input("مقدار / وزن پودر (گرم)", min_value=0.0, value=10000.0, key="nora_wt")
                    nora_date = st.date_input("تاریخ ورود/تحویل", key="nora_dt").strftime("%Y-%m-%d")
                    
                nora_submit = st.form_submit_button("💾 ثبت پودر نورا در بایگانی")
                if nora_submit:
                    if nora_powder_code:
                        c = conn.cursor()
                        c.execute("""INSERT OR REPLACE INTO nora_powders 
                                     (powder_code, material, weight_g, date)
                                     VALUES (?, ?, ?, ?)""",
                                  (nora_powder_code, nora_material, nora_weight_g, nora_date))
                        conn.commit()
                        st.success(f"پودر نورا با کد {nora_powder_code} با موفقیت ثبت شد.")
                        st.rerun()
                    else:
                        st.error("لطفاً کد/شماره ظرف پودر را وارد کنید.")

            st.markdown("---")
            st.subheader("📊 جدول پودرهای خریداری شده از نورا")
            if not nora_df.empty:
                disp_nora = nora_df.rename(columns=FARSI_HEADERS_MAP)
                st.table(disp_nora)
            else:
                st.info("هنوز پودری از نورا ثبت نشده است.")

        conn.close()

    # ---------------------------------------------------------
    # ۲. فرم تولید
    # ---------------------------------------------------------
    elif choice == "🏭 فرم تولید":
        st.header("🏭 فرم رکورد تولید (Production Form.xlsx)")
        
        conn = get_db_connection()
        powders_list = pd.read_sql_query("SELECT powder_code FROM powders", conn)['powder_code'].tolist()
        nora_powders_list = pd.read_sql_query("SELECT powder_code FROM nora_powders", conn)['powder_code'].tolist()
        all_powders = list(set(powders_list + nora_powders_list))
        
        st.subheader("حالت")
        mode = st.radio("", ["ثبت قطعه جدید", "ویرایش قطعه موجود"], horizontal=True, key="prod_mode")
        
        edit_data = {}
        if mode == "ویرایش قطعه موجود":
            existing_parts = pd.read_sql_query("SELECT part_code FROM production", conn)['part_code'].tolist()
            if existing_parts:
                selected_edit_code = st.selectbox("انتخاب قطعه جهت ویرایش:", existing_parts)
                edit_data = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(selected_edit_code,)).iloc[0].to_dict()
            else:
                st.warning("هنوز هیچ قطعه‌ای جهت ویرایش ثبت نشده است.")

        if not all_powders:
            st.warning("جهت ثبت فرم تولید، ابتدا باید حداقل یک پودر در بخش ۱ ثبت شده باشد.")
        else:
            with st.form("prod_form"):
                col1, col2, col3 = st.columns(3)
                with col1:
                    default_part_code = edit_data.get("part_code", "") if mode == "ویرایش قطعه موجود" else ""
                    part_code = st.text_input("کد قطعه (شناسه یکتا)", value=default_part_code, disabled=(mode == "ویرایش قطعه موجود"), placeholder="")
                    part_name = st.text_input("نام قطعه", value=edit_data.get("part_name", ""), placeholder="")
                with col2:
                    default_powder_index = all_powders.index(edit_data["powder_code"]) if mode == "ویرایش قطعه موجود" and edit_data.get("powder_code") in all_powders else 0
                    powder_code = st.selectbox("شماره ظرف پودر مصرفی", all_powders, index=default_powder_index)
                    
                    machine_options = ["M120", "M300", "سایر"]
                    default_machine_index = machine_options.index(edit_data["machine_model"]) if mode == "ویرایش قطعه موجود" and edit_data.get("machine_model") in machine_options else 0
                    machine_model = st.selectbox("مدل دستگاه", machine_options, index=default_machine_index)
                with col3:
                    quantity = st.number_input("تعداد روی صفحه ساخت", min_value=1, value=int(edit_data.get("quantity", 1)))
                    date_str = st.date_input("تاریخ ساخت", value=datetime.strptime(edit_data.get("date", datetime.now().strftime("%Y-%m-%d")), "%Y-%m-%d")).strftime("%Y-%m-%d")
                    
                st.markdown("---")
                st.subheader("⏱️ ۱- زمان آماده‌سازی و ساخت")
                tc1, tc2, tc3 = st.columns(3)
                with tc1:
                    build_time_hrs = st.number_input("زمان تولید (ساعت)", min_value=0.0, value=float(edit_data.get("build_time_hrs", 0.0)))
                    downtime_hrs = st.number_input("زمان توقف حین ساخت (ساعت)", min_value=0.0, value=float(edit_data.get("downtime_hrs", 0.0)))
                with tc2:
                    start_date = st.text_input("تاریخ شروع", value=edit_data.get("start_date", date_str), placeholder="")
                    start_time = st.text_input("ساعت شروع", value=edit_data.get("start_time", "08:00"), placeholder="")
                    end_date = st.text_input("تاریخ پایان", value=edit_data.get("end_date", date_str), placeholder="")
                    end_time = st.text_input("ساعت پایان", value=edit_data.get("end_time", "16:00"), placeholder="")
                with tc3:
                    setup_time_hrs = st.number_input("زمان آماده‌سازی دستگاه (ساعت)", min_value=0.0, value=float(edit_data.get("setup_time_hrs", 0.0)))
                    cleaning_time_hrs = st.number_input("زمان تمیزکاری دستگاه (ساعت)", min_value=0.0, value=float(edit_data.get("cleaning_time_hrs", 0.0)))
                    
                st.markdown("---")
                st.subheader("⚖️ ۲- پارامترهای متریال و وزن")
                mc1, mc2 = st.columns(2)
                with mc1:
                    input_powder_g = st.number_input("پودر ورودی به دستگاه (گرم)", min_value=0.0, value=float(edit_data.get("input_powder_g", 0.0)))
                    waste_powder_g = st.number_input("پودر مصرف شده غیر قابل بازیافت (گرم)", min_value=0.0, value=float(edit_data.get("waste_powder_g", 0.0)))
                    part_with_support_g = st.number_input("وزن قطعه با ساپورت (گرم)", min_value=0.0, value=float(edit_data.get("part_with_support_g", 0.0)))
                    final_part_g = st.number_input("وزن قطعه نهایی (گرم)", min_value=0.0, value=float(edit_data.get("final_part_g", 0.0)))
                    filter_pct = st.number_input("درصد فیلتر دستگاه (%)", min_value=0.0, max_value=100.0, value=float(edit_data.get("filter_percentage", 0.0)))
                with mc2:
                    plate_code = st.text_input("کد صفحه ساخت", value=edit_data.get("build_plate_code", ""), placeholder="")
                    plate_init_wt = st.number_input("وزن اولیه صفحه ساخت (گرم)", min_value=0.0, value=float(edit_data.get("build_plate_init_wt_g", 0.0)))
                    plate_post_wt = st.number_input("وزن صفحه ساخت پس از پرداخت (گرم)", min_value=0.0, value=float(edit_data.get("build_plate_post_wt_g", 0.0)))
                    
                submitted = st.form_submit_button("💾 ذخیره تغییرات / ثبت رکورد تولید")
                if submitted:
                    target_code = default_part_code if mode == "ویرایش قطعه موجود" else part_code
                    if target_code:
                        c = conn.cursor()
                        c.execute("""INSERT OR REPLACE INTO production 
                            (part_code, part_name, powder_code, quantity, machine_model, date,
                             build_time_hrs, downtime_hrs, start_date, start_time, end_date, end_time,
                             setup_time_hrs, cleaning_time_hrs, input_powder_g, waste_powder_g,
                             part_with_support_g, final_part_g, filter_percentage, build_plate_code,
                             build_plate_init_wt_g, build_plate_post_wt_g)
                            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                            (target_code, part_name, powder_code, quantity, machine_model, date_str,
                             build_time_hrs, downtime_hrs, start_date, start_time, end_date, end_time,
                             setup_time_hrs, cleaning_time_hrs, input_powder_g, waste_powder_g,
                             part_with_support_g, final_part_g, filter_pct, plate_code,
                             plate_init_wt, plate_post_wt))
                        conn.commit()
                        st.success(f"اطلاعات تولید قطعه {target_code} با موفقیت ذخیره شد.")
                        st.rerun()
                    else:
                        st.error("لطفاً کد قطعه را مشخص کنید.")
        conn.close()

    # ---------------------------------------------------------
    # ۳. کنترل کیفیت
    # ---------------------------------------------------------
    elif choice == "❇️ کنترل کیفیت":
        st.header("🔬 فرم کنترل کیفیت (QC.xlsx)")
        
        conn = get_db_connection()
        parts_list = pd.read_sql_query("SELECT part_code, part_name FROM production", conn)
        
        if parts_list.empty:
            st.warning("هیچ قطعه‌ای در بخش فرم تولید ثبت نشده است.")
        else:
            selected_part = st.selectbox("انتخاب کد قطعه جهت ارزیابی کیفیت", parts_list['part_code'])
            part_info = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(selected_part,)).iloc[0]
            
            st.info(f"**نام قطعه:** {part_info['part_name']} | **دستگاه:** {part_info['machine_model']} | **شماره ظرف پودر:** {part_info['powder_code']}")
            
            with st.form("qc_form"):
                st.subheader("📋 تست‌ها و بازرسی‌های کیفی")
                tests = [
                    ("1.1.1 ظاهر سطح و عیوب قابل رؤیت", "چشمی"),
                    ("1.1.2 کیفیت حذف ساپورت", "چشمی"),
                    ("1.1.3 زبری سطح (Ra)", "ابزار زبری‌سنج"),
                    ("2.1.1 ابعاد بحرانی", "کولیس/CMM"),
                    ("2.1.2 تختی، هم‌محوری، عمودیت و GD&T", "CMM/ژئومتریک"),
                    ("2.1.3 انطباق با مدل CAD", "اسکن سه بعدی"),
                    ("3.1.1 تخلخل/ترک/ناپیوستگی", "NDT"),
                    ("4.1.1 استحکام کششی / تسلیم", "تست مکانیکی"),
                    ("4.1.2 سختی", "سختی‌سنج"),
                    ("4.1.3 چگالی نسبی / تخلخل", "دانسیته‌سنج"),
                    ("4.1.4 ریزساختار و کیفیت ذوب", "متالوگرافی")
                ]
                
                qc_data = {}
                for t_title, t_type in tests:
                    q1, q2, q3 = st.columns([3, 2, 4])
                    with q1: st.write(f"**{t_title}**")
                    with q2: res = st.selectbox("نتیجه", ["تایید", "رد"], key=t_title)
                    with q3: note = st.text_input("ملاحظات", key=f"qc_n_{t_title}", placeholder="")
                    qc_data[t_title] = {"result": res, "type": t_type, "note": note}
                    
                st.markdown("---")
                st.subheader("👥 مسئولین و تاییدکنندگان")
                sc1, sc2, sc3 = st.columns(3)
                with sc1: inspector = st.text_input("بازرس کنترل کیفیت", placeholder="")
                with sc2: engineer = st.text_input("مسئول فنی / مهندسی کیفیت", placeholder="")
                with sc3: manager = st.text_input("مدیر تضمین کیفیت", placeholder="")
                
                qc_submit = st.form_submit_button("💾 ثبت نهایی فرم QC")
                if qc_submit:
                    c = conn.cursor()
                    c.execute("""INSERT OR REPLACE INTO qc 
                                 (part_code, part_name, material, machine_model, date, qc_checks_json, qc_inspector, qc_engineer, qa_manager)
                                 VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                              (selected_part, part_info['part_name'], part_info['powder_code'], part_info['machine_model'],
                               datetime.now().strftime("%Y-%m-%d"), json.dumps(qc_data, ensure_ascii=False), inspector, engineer, manager))
                    conn.commit()
                    st.success("نتایج ارزیابی QC با موفقیت ذخیره شد.")
        conn.close()

    # ---------------------------------------------------------
    # ۴. محاسبه‌گر هزینه
    # ---------------------------------------------------------
    elif choice == "💰 محاسبه‌گر هزینه":
        st.header("💰 محاسبه‌گر بهای تمام شده و قیمت فروش (Cost Calculator.xlsx)")
        
        SYSTEM_RATES = load_system_rates()
        
        RATES = {
            "powder_price_per_kg": {
                "Steel 316": SYSTEM_RATES["powder_price_Steel_316"],
                "Ti6Al4V": SYSTEM_RATES["powder_price_Ti6Al4V"],
                "Inconel 718": SYSTEM_RATES["powder_price_Inconel_718"],
                "Hastelloy X": SYSTEM_RATES["powder_price_Hastelloy_X"]
            },
            "machine_depreciation_hr": {
                "M120": SYSTEM_RATES["machine_depr_M120"],
                "M300": SYSTEM_RATES["machine_depr_M300"]
            },
            "power_kw_hr": {
                "M120": SYSTEM_RATES["power_kw_M120"],
                "M300": SYSTEM_RATES["power_kw_M300"]
            },
            "wages": {
                "designer_hr": SYSTEM_RATES["wage_designer"],
                "operator_hr": SYSTEM_RATES["wage_operator"],
                "qc_hr": SYSTEM_RATES["wage_qc"]
            },
            "argon_hr": SYSTEM_RATES["argon_rate"],
            "electricity_kwh": SYSTEM_RATES["electricity_rate"],
            "post_process_hr": SYSTEM_RATES["post_process_rate"],
            "qc_fixed_cost": SYSTEM_RATES["qc_fixed_cost"],
            "ventilation_hr": SYSTEM_RATES["ventilation_rate"],
            "chiller_hr": SYSTEM_RATES["chiller_rate"],
            "density": {"Steel 316": 8.0, "Ti6Al4V": 4.43, "Inconel 718": 8.19, "Hastelloy X": 8.22}
        }
        
        conn = get_db_connection()
        parts_list = pd.read_sql_query("SELECT part_code, part_name FROM production", conn)
        
        selected_part = st.selectbox("فراخوانی قطعه از بخش تولید (یا وارد کردن قطعه جدید)", ["جدید"] + list(parts_list['part_code']))
        
        def_part_name = ""
        def_machine = "M300"
        def_print_time = 10.0
        
        if selected_part != "جدید":
            p_row = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(selected_part,)).iloc[0]
            def_part_name = p_row['part_name']
            def_machine = p_row['machine_model'] if p_row['machine_model'] in ["M120", "M300"] else "M300"
            def_print_time = float(p_row['build_time_hrs']) if p_row['build_time_hrs'] else 10.0

        st.subheader("📥 مشخصات فنی و ورودی‌های قطعه")
        c1, c2, c3 = st.columns(3)
        with c1:
            p_code = st.text_input("شناسه/کد قطعه", value=selected_part if selected_part != "جدید" else "", placeholder="")
            p_name = st.text_input("نام قطعه", value=def_part_name, placeholder="")
            powder_type = st.selectbox("نوع پودر فلزی", list(RATES["density"].keys()))
        with c2:
            vol_cm3 = st.number_input("حجم قطعه (cm3)", min_value=0.0, value=50.0)
            sup_vol_cm3 = st.number_input("حجم ساپورت (cm3)", min_value=0.0, value=10.0)
            machine_type = st.selectbox("نوع دستگاه", ["M120", "M300"], index=1 if def_machine=="M300" else 0)
        with c3:
            parts_on_plate = st.number_input("تعداد قطعات روی صفحه", min_value=1, value=1)
            print_time_hrs = st.number_input("زمان کل چاپ (ساعت)", min_value=0.0, value=def_print_time)
            design_time_hrs = st.number_input("زمان طراحی (ساعت)", min_value=0.0, value=2.0)
            post_time_hrs = st.number_input("زمان پرداخت‌کاری (ساعت)", min_value=0.0, value=3.0)
            overhead_pct = st.number_input("ضریب سربار (%)", min_value=0.0, value=35.0)

        density = RATES["density"][powder_type]
        net_weight_g = vol_cm3 * density
        support_weight_g = sup_vol_cm3 * density
        total_weight_kg = (net_weight_g + support_weight_g) / 1000.0
        
        cost_powder = total_weight_kg * RATES["powder_price_per_kg"][powder_type]
        cost_argon = print_time_hrs * RATES["argon_hr"]
        cost_depreciation = print_time_hrs * RATES["machine_depreciation_hr"][machine_type]
        cost_power = print_time_hrs * RATES["power_kw_hr"][machine_type] * RATES["electricity_kwh"]
        cost_engineering = design_time_hrs * RATES["wages"]["designer_hr"]
        cost_operator = print_time_hrs * RATES["wages"]["operator_hr"]
        cost_post_process = post_time_hrs * RATES["post_process_hr"]
        cost_qc = RATES["qc_fixed_cost"]
        cost_ventilation = print_time_hrs * RATES["ventilation_hr"]
        cost_chiller = print_time_hrs * RATES["chiller_hr"]
        
        total_production_cost = (cost_powder + cost_argon + cost_depreciation + cost_power + 
                                 cost_engineering + cost_operator + cost_post_process + 
                                 cost_qc + cost_ventilation + cost_chiller)
        
        overhead_cost = total_production_cost * (overhead_pct / 100.0)
        final_price = total_production_cost + overhead_cost
        
        st.markdown("---")
        st.subheader("📊 تفکیک و خلاصه هزینه‌های برآورد شده")
        
        mc1, mc2, mc3 = st.columns(3)
        mc1.metric("وزن خالص قطعه", f"{net_weight_g:.1f} گرم")
        mc2.metric("بهای تمام شده کل", f"{total_production_cost:,.0f} ریال")
        mc3.metric("قیمت نهایی قابل ارائه به مشتری", f"{final_price:,.0f} ریال")
        
        if st.button("💾 ذخیره برآورد قیمت"):
            if p_code:
                c = conn.cursor()
                c.execute("""INSERT OR REPLACE INTO cost_calculator 
                    (part_code, part_name, powder_type, volume_cm3, net_weight_g, support_volume_cm3, support_weight_g,
                     machine_type, parts_on_plate, print_time_hrs, design_time_hrs, post_process_time_hrs, overhead_pct,
                     powder_cost_total, argon_cost_total, depreciation_cost_total, power_cost_total, engineering_cost_total,
                     operator_cost_total, post_process_cost_total, qc_cost_total, utility_ventilation, utility_chiller,
                     total_production_cost, overhead_cost, final_price)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                    (p_code, p_name, powder_type, vol_cm3, net_weight_g, sup_vol_cm3, support_weight_g,
                     machine_type, parts_on_plate, print_time_hrs, design_time_hrs, post_time_hrs, overhead_pct,
                     cost_powder, cost_argon, cost_depreciation, cost_power, cost_engineering,
                     cost_operator, cost_post_process, cost_qc, cost_ventilation, cost_chiller,
                     total_production_cost, overhead_cost, final_price))
                conn.commit()
                st.success("محاسبه هزینه با موفقیت ثبت شد.")
            else:
                st.error("لطفاً شناسه قطعه را مشخص کنید.")
        conn.close()

    # ---------------------------------------------------------
    # ۵. تنظیمات نرخ‌ها
    # ---------------------------------------------------------
    elif choice == "⚙️ تنظیمات نرخ‌ها":
        st.header("⚙️ تنظیمات و ویرایش نرخ‌های پایه محاسبات")
        st.info("تغییرات در این بخش بلافاصله در «محاسبه‌گر هزینه» و فرآیندهای مالی اعمال می‌شوند.")
        
        current_rates = load_system_rates()
        
        with st.form("settings_form"):
            st.subheader("۱. قیمت پودرهای فلزی (ریال به ازای هر کیلوگرم)")
            c1, c2 = st.columns(2)
            with c1:
                r_steel = st.number_input("پودر Steel 316", min_value=0.0, value=float(current_rates.get("powder_price_Steel_316", 120000000)))
                r_ti = st.number_input("پودر Ti6Al4V", min_value=0.0, value=float(current_rates.get("powder_price_Ti6Al4V", 500000000)))
            with c2:
                r_inconel = st.number_input("پودر Inconel 718", min_value=0.0, value=float(current_rates.get("powder_price_Inconel_718", 300000000)))
                r_hastelloy = st.number_input("پودر Hastelloy X", min_value=0.0, value=float(current_rates.get("powder_price_Hastelloy_X", 400000000)))

            st.markdown("---")
            st.subheader("۲. نرخ ساعتی استهلاک دستگاه‌ها (ریال / ساعت)")
            m1, m2 = st.columns(2)
            with m1:
                r_depr_m120 = st.number_input("نرخ استهلاک دستگاه M120", min_value=0.0, value=float(current_rates.get("machine_depr_M120", 1250000)))
            with m2:
                r_depr_m300 = st.number_input("نرخ استهلاک دستگاه M300", min_value=0.0, value=float(current_rates.get("machine_depr_M300", 3125000)))

            st.markdown("---")
            st.subheader("۳. دستمزدهای نیروی انسانی (ریال / نفرساعت)")
            w1, w2, w3 = st.columns(3)
            with w1:
                r_wage_designer = st.number_input("دستمزد طراح مهندسی", min_value=0.0, value=float(current_rates.get("wage_designer", 2500000)))
            with w2:
                r_wage_operator = st.number_input("دستمزد اپراتور دستگاه", min_value=0.0, value=float(current_rates.get("wage_operator", 1870000)))
            with w3:
                r_wage_qc = st.number_input("دستمزد مسئول کنترل کیفیت", min_value=0.0, value=float(current_rates.get("wage_qc", 2180000)))

            st.markdown("---")
            st.subheader("۴. هزینه‌های جانبی، انرژی و مصارف کارگاهی")
            u1, u2, u3 = st.columns(3)
            with u1:
                r_argon = st.number_input("نرخ گاز آرگون (ریال/ساعت)", min_value=0.0, value=float(current_rates.get("argon_rate", 300000)))
                r_electricity = st.number_input("نرخ برق مصرفی (ریال/کیلووات‌ساعت)", min_value=0.0, value=float(current_rates.get("electricity_rate", 50000)))
            with u2:
                r_post_process = st.number_input("عملیات پرداخت سطحی (ریال/ماشین‌ساعت)", min_value=0.0, value=float(current_rates.get("post_process_rate", 500000)))
                r_qc_fixed = st.number_input("تست‌های QC ثابت (ریال/قطعه)", min_value=0.0, value=float(current_rates.get("qc_fixed_cost", 40000000)))
            with u3:
                r_ventilation = st.number_input("سیستم تهویه (ریال/ساعت)", min_value=0.0, value=float(current_rates.get("ventilation_rate", 200000)))
                r_chiller = st.number_input("آب خنک‌کاری چیلر (ریال/ساعت)", min_value=0.0, value=float(current_rates.get("chiller_rate", 2500)))

            save_rates_submit = st.form_submit_button("💾 ذخیره تغییرات نرخ‌ها در پایگاه داده")
            if save_rates_submit:
                conn = get_db_connection()
                c = conn.cursor()
                updated_pairs = [
                    ("powder_price_Steel_316", r_steel),
                    ("powder_price_Ti6Al4V", r_ti),
                    ("powder_price_Inconel_718", r_inconel),
                    ("powder_price_Hastelloy_X", r_hastelloy),
                    ("machine_depr_M120", r_depr_m120),
                    ("machine_depr_M300", r_depr_m300),
                    ("wage_designer", r_wage_designer),
                    ("wage_operator", r_wage_operator),
                    ("wage_qc", r_wage_qc),
                    ("argon_rate", r_argon),
                    ("electricity_rate", r_electricity),
                    ("post_process_rate", r_post_process),
                    ("qc_fixed_cost", r_qc_fixed),
                    ("ventilation_rate", r_ventilation),
                    ("chiller_rate", r_chiller)
                ]
                for k, val in updated_pairs:
                    c.execute("INSERT OR REPLACE INTO system_settings (key, value) VALUES (?, ?)", (k, val))
                conn.commit()
                conn.close()
                st.success("تمام نرخ‌های جدید با موفقیت ذخیره شدند و در بخش «محاسبه‌گر هزینه» اعمال گردیدند.")
                st.rerun()

    # ---------------------------------------------------------
    # ۶. بایگانی و گزارش‌گیری اکسل
    # ---------------------------------------------------------
    elif choice == "🔍 بایگانی":
        st.header("🔍 بایگانی جامع، استعلام پرونده‌ها و مدیریت رکوردها")
        
        conn = get_db_connection()
        
        all_parts_df = pd.read_sql_query("SELECT part_code, part_name, powder_code, machine_model, quantity, date FROM production", conn)
        
        st.subheader("📋 لیست کلی تمامی قطعات ثبت‌شده")
        if not all_parts_df.empty:
            disp_all_parts = all_parts_df.rename(columns={
                'part_code': 'کد قطعه',
                'part_name': 'نام قطعه',
                'powder_code': 'شماره ظرف پودر مصرفی',
                'machine_model': 'مدل دستگاه',
                'quantity': 'تعداد روی صفحه',
                'date': 'تاریخ ساخت'
            })
            st.table(disp_all_parts)
        else:
            st.info("هنوز هیچ قطعه‌ای در سامانه ثبت نشده است.")
            
        st.markdown("---")
        st.subheader("🔎 استعلام پرونده جامع قطعه")
        search_code = st.text_input("کد قطعه را جهت استعلام کامل وارد کنید:", placeholder="")
        
        if search_code:
            prod_df = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(search_code,))
            qc_df = pd.read_sql_query("SELECT * FROM qc WHERE part_code=?", conn, params=(search_code,))
            cost_df = pd.read_sql_query("SELECT * FROM cost_calculator WHERE part_code=?", conn, params=(search_code,))
            
            if not prod_df.empty:
                st.success(f"اطلاعات قطعه {search_code} یافت شد.")
                tab1, tab2, tab3, tab4 = st.tabs(["📌 پارامترهای تولید", "🧪 اطلاعات پودر", "🔬 کنترل کیفیت (QC)", "💰 برآورد مالی"])
                
                json_cols_to_drop = ['checklist_json', 'qc_checks_json', 'finishing_json']
                
                with tab1:
                    st.subheader("مشخصات فنی و فرآیند ساخت")
                    clean_prod = prod_df.drop(columns=[c for c in json_cols_to_drop if c in prod_df.columns])
                    disp_prod = clean_prod.rename(columns=FARSI_HEADERS_MAP).T
                    disp_prod.columns = ["مقدار / مقدار ثبت شده"]
                    st.table(disp_prod)
                    
                with tab2:
                    powder_code = prod_df['powder_code'].values[0]
                    powder_df = pd.read_sql_query("SELECT * FROM powders WHERE powder_code=?", conn, params=(powder_code,))
                    if powder_df.empty:
                        powder_df = pd.read_sql_query("SELECT * FROM nora_powders WHERE powder_code=?", conn, params=(powder_code,))
                    
                    if not powder_df.empty:
                        st.write(f"**کد ظرف پودر استفاده شده:** `{powder_code}`")
                        clean_powder = powder_df.drop(columns=[c for c in json_cols_to_drop if c in powder_df.columns])
                        disp_powder = clean_powder.rename(columns=FARSI_HEADERS_MAP).T
                        disp_powder.columns = ["مشخصات پودر"]
                        st.table(disp_powder)
                    else:
                        st.warning("اطلاعات پودر متناظر یافت نشد.")
                        
                with tab3:
                    if not qc_df.empty:
                        clean_qc = qc_df.drop(columns=[c for c in json_cols_to_drop if c in qc_df.columns])
                        disp_qc = clean_qc.rename(columns=FARSI_HEADERS_MAP).T
                        disp_qc.columns = ["اطلاعات ارزیابی کیفیت"]
                        st.table(disp_qc)
                    else:
                        st.info("فرم کنترل کیفیت برای این قطعه هنوز ثبت نشده است.")
                        
                with tab4:
                    if not cost_df.empty:
                        disp_cost = cost_df.rename(columns=FARSI_HEADERS_MAP).T
                        disp_cost.columns = ["جزئیات مالی و برآورد هزینه"]
                        st.table(disp_cost)
                        st.metric("قیمت نهایی فروش (ریال)", f"{cost_df['final_price'].values[0]:,.0f}")
                    else:
                        st.info("محاسبه هزینه برای این قطعه ثبت نشده است.")
            else:
                st.error("قطعه‌ای با این کد یافت نشد.")
                
        st.markdown("---")
        st.subheader("📂 مشاهده و دانلود جداول پایگاه داده")
        target_table = st.selectbox("انتخاب جدول جهت مشاهده، دانلود Excel یا حذف رکورد:", 
                                    ["آنالیز پودر اولیه (powders)", "پودرهای بازیافت شده (recycled_powders)", "پودرهای خریداری شده از نورا (nora_powders)", "رکوردهای تولید (production)", "کنترل کیفیت (qc)", "محاسبه هزینه (cost_calculator)"])
        
        table_map = {
            "آنالیز پودر اولیه (powders)": ("powders", "powder_code"),
            "پودرهای بازیافت شده (recycled_powders)": ("recycled_powders", "id"),
            "پودرهای خریداری شده از نورا (nora_powders)": ("nora_powders", "powder_code"),
            "رکوردهای تولید (production)": ("production", "part_code"),
            "کنترل کیفیت (qc)": ("qc", "part_code"),
            "محاسبه هزینه (cost_calculator)": ("cost_calculator", "part_code")
        }
        
        selected_tbl, pkey_col = table_map[target_table]
        df = pd.read_sql_query(f"SELECT * FROM {selected_tbl}", conn)
        
        json_cols_to_drop = ['checklist_json', 'qc_checks_json', 'finishing_json']
        clean_df = df.drop(columns=[col for col in json_cols_to_drop if col in df.columns])
        
        if not clean_df.empty:
            farsi_df = clean_df.rename(columns=FARSI_HEADERS_MAP)
            st.table(farsi_df)
            
            col_down, col_del = st.columns([2, 2])
            
            with col_down:
                excel_filename = f"{selected_tbl}_export.xlsx"
                farsi_df.to_excel(excel_filename, index=False)
                with open(excel_filename, "rb") as f:
                    st.download_button(
                        label="📥 دانلود فایل اکسل جدول",
                        data=f,
                        file_name=excel_filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
            with col_del:
                if role == "مدیریت":
                    with st.expander("🗑️ مدیریت و حذف رکورد از این جدول"):
                        item_options = df[pkey_col].tolist()
                        item_to_delete = st.selectbox(f"انتخاب شناسه/کد جهت حذف از جدول:", item_options)
                        
                        if st.button("⚠️ حذف رکورد انتخاب شده"):
                            c = conn.cursor()
                            c.execute(f"DELETE FROM {selected_tbl} WHERE {pkey_col}=?", (item_to_delete,))
                            conn.commit()
                            st.success(f"رکورد با کد {item_to_delete} با موفقیت حذف شد.")
                            st.rerun()
        else:
            st.info("اطلاعاتی در این جدول ثبت نشده است.")
            
        conn.close()

    # ---------------------------------------------------------
    # ۷. مدیریت کامل کاربران
    # ---------------------------------------------------------
    elif choice == "👥 مدیریت کاربران":
        st.header("👥 مدیریت کامل کاربران و سطوح دسترسی")
        
        conn = get_db_connection()
        users_df = pd.read_sql_query("SELECT username, full_name, role FROM users", conn)
        
        st.subheader("📋 لیست کاربران فعلی سامانه")
        st.table(users_df.rename(columns={'username': 'نام کاربری', 'full_name': 'نام و نام خانوادگی', 'role': 'سطح دسترسی'}))
        
        st.markdown("---")
        
        u_tab1, u_tab2, u_tab3 = st.tabs(["➕ ۱. تعریف کاربر جدید", "✏️ ۲. ویرایش کاربر موجود", "🗑️ ۳. حذف کاربر"])
        
        with u_tab1:
            with st.form("add_user_form"):
                u_col1, u_col2 = st.columns(2)
                with u_col1:
                    new_username = st.text_input("نام کاربری جدید (لاتین)", placeholder="")
                    new_password = st.text_input("رمز عبور", type="password", placeholder="")
                with u_col2:
                    new_fullname = st.text_input("نام و نام خانوادگی", placeholder="")
                    new_role = st.selectbox("سطح دسترسی / نقش", ["مدیریت", "اپراتور و طراح", "کنترل کیفیت", "بازرگانی"])
                
                submit_new_user = st.form_submit_button("💾 ثبت کاربر جدید")
                if submit_new_user:
                    if new_username and new_password:
                        c = conn.cursor()
                        c.execute("INSERT OR REPLACE INTO users (username, password_hash, full_name, role) VALUES (?, ?, ?, ?)",
                                  (new_username, hash_password(new_password), new_fullname, new_role))
                        conn.commit()
                        st.success(f"کاربر {new_username} با موفقیت تعریف شد.")
                        st.rerun()
                    else:
                        st.error("نام کاربری و رمز عبور الزامی است.")

        with u_tab2:
            all_usernames = users_df['username'].tolist()
            if all_usernames:
                selected_edit_user = st.selectbox("انتخاب کاربر جهت ویرایش:", all_usernames)
                user_info = pd.read_sql_query("SELECT * FROM users WHERE username=?", conn, params=(selected_edit_user,)).iloc[0]
                
                with st.form("edit_user_form"):
                    e_col1, e_col2 = st.columns(2)
                    with e_col1:
                        st.text_input("نام کاربری", value=user_info['username'], disabled=True, placeholder="")
                        edit_password = st.text_input("رمز عبور جدید (در صورت عدم تغییر خالی بگذارید)", type="password", placeholder="")
                    with e_col2:
                        edit_fullname = st.text_input("نام و نام خانوادگی", value=user_info['full_name'], placeholder="")
                        roles_list = ["مدیریت", "اپراتور و طراح", "کنترل کیفیت", "بازرگانی"]
                        default_role_idx = roles_list.index(user_info['role']) if user_info['role'] in roles_list else 0
                        edit_role = st.selectbox("سطح دسترسی / نقش", roles_list, index=default_role_idx)
                    
                    submit_edit_user = st.form_submit_button("💾 ذخیره تغییرات کاربر")
                    if submit_edit_user:
                        c = conn.cursor()
                        if edit_password:
                            c.execute("UPDATE users SET full_name=?, role=?, password_hash=? WHERE username=?",
                                      (edit_fullname, edit_role, hash_password(edit_password), selected_edit_user))
                        else:
                            c.execute("UPDATE users SET full_name=?, role=? WHERE username=?",
                                      (edit_fullname, edit_role, selected_edit_user))
                        conn.commit()
                        st.success(f"اطلاعات کاربر {selected_edit_user} با موفقیت به‌روزرسانی شد.")
                        st.rerun()

        with u_tab3:
            all_usernames = users_df['username'].tolist()
            deletable_users = [u for u in all_usernames if u != st.session_state["username"]]
            
            if deletable_users:
                selected_delete_user = st.selectbox("انتخاب کاربر جهت حذف از سیستم:", deletable_users)
                if st.button("⚠️ حذف کاربر انتخاب شده"):
                    c = conn.cursor()
                    c.execute("DELETE FROM users WHERE username=?", (selected_delete_user,))
                    conn.commit()
                    st.success(f"کاربر {selected_delete_user} با موفقیت حذف شد.")
                    st.rerun()
            else:
                st.info("کاربر دیگری جهت حذف وجود ندارد.")

        conn.close()
