import streamlit as st
import pandas as pd
import sqlite3
import json
from datetime import datetime

# تنظیمات اولیه صفحه Streamlit
st.set_page_config(
    page_title="سامانه مدیریت و بایگانی ساخت افزایشی (SLM)",
    page_icon="⚙️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- استایل اختصاصی: بزرگ‌تر کردن فونت عمومی، اصلاح دقیق منوی سایدبار و کارت‌ها ---
st.markdown("""
<style>
    /* تعریف و فراخوانی فونت B Nazanin */
    @font-face {
        font-family: 'B Nazanin';
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot');
        src: url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.eot?#iefix') format('embedded-opentype'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.woff') format('woff'),
             url('https://cdn.fontcdn.ir/Font/Persian/BNazanin/BNazanin.ttf') format('truetype');
        font-weight: normal;
        font-style: normal;
    }

    /* ۱. اعمال فونت B Nazanin، راست‌چین‌سازی و افزایش سایز فونت عمومی */
    html, body, p, h1, h2, h3, h4, h5, h6, input, button, select, label {
        font-family: 'B Nazanin', 'Vazir', sans-serif !important;
        direction: rtl;
        text-align: right;
    }

    /* افزایش سایز متون عمومی و عناوین */
    p, span, label, input, select {
        font-size: 1.25rem !important;
    }
    h1 { font-size: 2.3rem !important; }
    h2 { font-size: 1.9rem !important; }
    h3 { font-size: 1.6rem !important; }

    /* ۲. تنظیم سایدبار سرمه‌ای با فونت درشت */
    [data-testid="stSidebar"] {
        background-color: #1a2536 !important;
        direction: rtl;
        text-align: right;
        border-left: 1px solid #2c3e50;
    }

    /* ۳. اصلاح جعبه‌های انتخاب (Selectbox) در سایدبار جهت جلوگیری از سفید شدن متن */
    [data-testid="stSidebar"] div[data-baseweb="select"] > div {
        background-color: #1e293b !important;
        color: #ffffff !important;
        border: 1px solid #334155 !important;
        border-radius: 8px !important;
    }

    [data-testid="stSidebar"] div[data-baseweb="select"] span {
        color: #ffffff !important;
        font-weight: bold !important;
        font-size: 1.2rem !important;
    }

    [data-testid="stSidebar"] label, [data-testid="stSidebar"] p {
        color: #f8f9fa !important;
        font-size: 1.3rem !important;
        font-weight: bold !important;
    }

    /* ۴. کارت‌های آمار و شاخص‌ها با فونت درشت */
    [data-testid="stMetric"] {
        background-color: #1e293b !important;
        border: 1px solid #334155 !important;
        border-radius: 10px !important;
        padding: 15px !important;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.2) !important;
    }
    [data-testid="stMetricLabel"] > div {
        color: #94a3b8 !important;
        font-weight: bold !important;
        font-size: 1.2rem !important;
    }
    [data-testid="stMetricValue"] > div {
        color: #38bdf8 !important;
        font-weight: bold !important;
        font-size: 1.8rem !important;
    }

    /* ۵. استایل دکمه اصلی */
    .stButton>button {
        width: 100%;
        background-color: #2563eb !important;
        color: #ffffff !important;
        border-radius: 8px !important;
        font-weight: bold !important;
        font-size: 1.3rem !important;
        padding: 10px 20px !important;
        border: none !important;
        transition: all 0.3s ease !important;
    }
    .stButton>button:hover {
        background-color: #1d4ed8 !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.4) !important;
    }

    /* ۶. آکاردئون‌ها */
    .streamlit-expanderHeader {
        background-color: #1e293b !important;
        color: #f8f9fa !important;
        border-radius: 6px !important;
        border: 1px solid #334155 !important;
        font-size: 1.3rem !important;
    }

    /* ۷. افزایش سایز فونت جداول */
    table {
        font-size: 1.15rem !important;
    }
</style>
""", unsafe_allow_html=True)

# --- مدیریت پایگاه داده SQLite ---
DB_NAME = "slm_management.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # جدول پودرها (خریداری شده اولیه)
    c.execute('''CREATE TABLE IF NOT EXISTS powders (
                    powder_code TEXT PRIMARY KEY,
                    material TEXT,
                    weight_g REAL,
                    date TEXT,
                    checklist_json TEXT
                )''')
    
    # جدول پودرهای بازیافت شده
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
    
    # جدول فرم‌های تولید
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
    
    # جدول کنترل کیفیت (QC)
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
    
    # جدول محاسبه‌گر هزینه
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

# --- منوی اصلی وب اپلیکیشن ---
st.title("⚙️ سامانه مدیریت و بایگانی ساخت افزایشی (SLM)")
st.caption("سیستم یکپارچه متصل‌کننده فرم‌های آنالیز پودر، تولید، کنترل کیفیت و برآورد قیمت")

menu = [
    "🔍 داشبورد و جستجوی جامع",
    "🧪 ۱. آنالیز و بایگانی پودر",
    "🏭 ۲. فرم ثبت تولید",
    "🔬 ۳. فرم کنترل کیفیت (QC)",
    "💰 ۴. محاسبه‌گر قیمت قطعه",
    "📂 ۵. خروجی و گزارش‌گیری اکسل"
]

choice = st.sidebar.selectbox("📋 منوی دسترسی بخش‌ها", menu)

# ---------------------------------------------------------
# ۱. داشبورد و جستجوی جامع
# ---------------------------------------------------------
if choice == "🔍 داشبورد و جستجوی جامع":
    st.header("🔍 استعلام و بایگانی کامل قطعه")
    
    search_code = st.text_input("کد قطعه را جهت جستجو وارد کنید:")
    
    if search_code:
        conn = get_db_connection()
        prod_df = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(search_code,))
        qc_df = pd.read_sql_query("SELECT * FROM qc WHERE part_code=?", conn, params=(search_code,))
        cost_df = pd.read_sql_query("SELECT * FROM cost_calculator WHERE part_code=?", conn, params=(search_code,))
        
        if not prod_df.empty:
            st.success(f"اطلاعات قطعه {search_code} یافت شد.")
            
            tab1, tab2, tab3, tab4 = st.tabs(["📌 پارامترهای تولید", "🧪 اطلاعات پودر", "🔬 کنترل کیفیت (QC)", "💰 برآورد مالی"])
            
            with tab1:
                st.subheader("مشخصات فنی و فرآیند ساخت")
                st.table(prod_df.T)
                
            with tab2:
                powder_code = prod_df['powder_code'].values[0]
                powder_df = pd.read_sql_query("SELECT * FROM powders WHERE powder_code=?", conn, params=(powder_code,))
                if not powder_df.empty:
                    st.write(f"**کد ظرف پودر استفاده شده:** `{powder_code}`")
                    st.table(powder_df.T)
                else:
                    st.warning("اطلاعات پودر متناظر یافت نشد.")
                    
            with tab3:
                if not qc_df.empty:
                    st.table(qc_df.T)
                else:
                    st.info("فرم کنترل کیفیت برای این قطعه هنوز ثبت نشده است.")
                    
            with tab4:
                if not cost_df.empty:
                    st.table(cost_df.T)
                    st.metric("قیمت نهایی فروش (ریال)", f"{cost_df['final_price'].values[0]:,.0f}")
                else:
                    st.info("محاسبه هزینه برای این قطعه ثبت نشده است.")
        else:
            st.error("قطعه‌ای با این کد یافت نشد.")
        conn.close()

# ---------------------------------------------------------
# ۲. آنالیز و بایگانی پودر
# ---------------------------------------------------------
elif choice == "🧪 ۱. آنالیز و بایگانی پودر":
    st.header("🧪 فرم مدیریت، آنالیز و بایگانی پودر")
    
    conn = get_db_connection()
    powders_df = pd.read_sql_query("SELECT * FROM powders", conn)
    
    sub_tab1, sub_tab2 = st.tabs(["📦 ۱- پودرهای خریداری شده اولیه", "♻️ ۲- پودرهای بازیافت شده"])
    
    with sub_tab1:
        with st.expander("➕ ثبت / ویرایش پودر جدید", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                powder_code = st.text_input("کد/شماره ظرف پودر فلز (مانند: PWD-316-01)")
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
                with c3: note = st.text_input("ملاحظات / توضیحات", key=f"note_{item}")
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
            disp_powders = powders_df[['powder_code', 'material', 'weight_g', 'date']].rename(columns={
                'powder_code': 'شماره ظرف پودر',
                'material': 'جنس پودر',
                'weight_g': 'وزن (گرم)',
                'date': 'تاریخ ورود'
            })
            st.table(disp_powders)
        else:
            st.info("هیچ پودر اولیه ای در سیستم ثبت نشده است.")

    with sub_tab2:
        st.subheader("♻️ مدیریت و ثبت پودرهای بازیافت شده")
        
        powders_list = powders_df['powder_code'].tolist() if not powders_df.empty else []
        
        if not powders_list:
            st.warning("ابتدا باید حداقل یک ظرف پودر اولیه ثبت کرده باشید.")
        else:
            with st.form("recycled_powder_form"):
                rc1, rc2 = st.columns(2)
                with rc1:
                    selected_powder_code = st.selectbox("شماره/کد ظرف پودر مبدا", powders_list)
                    recycled_batch_code = st.text_input("شناسه پارت بازیافت (مانند: REC-PWD-01)")
                    input_powder_g = st.number_input("پودر ورودی به دستگاه (گرم)", min_value=0.0, value=5000.0)
                with rc2:
                    unrecyclable_powder_g = st.number_input("پودر مصرف شده غیر قابل بازیافت (گرم)", min_value=0.0, value=200.0)
                    recycled_powder_g = input_powder_g - unrecyclable_powder_g
                    st.metric("مقدار پودر بازیافت‌شده قابل استفاده (گرم)", f"{recycled_powder_g:,.1f}")
                    rec_date = st.date_input("تاریخ بازیافت").strftime("%Y-%m-%d")
                    rec_notes = st.text_input("توضیحات و ملاحظات غربال‌گری / الک")
                
                rec_submit = st.form_submit_button("💾 ثبت رکورد پودر بازیافتی")
                if rec_submit:
                    c = conn.cursor()
                    c.execute("""INSERT INTO recycled_powders 
                                 (powder_code, recycled_batch_code, input_powder_g, unrecyclable_powder_g, recycled_powder_g, date, notes)
                                 VALUES (?, ?, ?, ?, ?, ?, ?)""",
                              (selected_powder_code, recycled_batch_code, input_powder_g, unrecyclable_powder_g, recycled_powder_g, rec_date, rec_notes))
                    conn.commit()
                    st.success("اطلاعات پودر بازیافتی با موفقیت ثبت شد.")
                    st.rerun()

        st.markdown("---")
        st.subheader("📊 جدول بایگانی پودرهای بازیافت شده")
        recycled_df = pd.read_sql_query("SELECT * FROM recycled_powders", conn)
        
        if not recycled_df.empty:
            display_recycled_df = recycled_df.rename(columns={
                'id': 'شناسه',
                'powder_code': 'کد ظرف پودر مبدا',
                'recycled_batch_code': 'کد بازیافت',
                'input_powder_g': 'پودر ورودی (گرم)',
                'unrecyclable_powder_g': 'غیرقابل بازیافت (گرم)',
                'recycled_powder_g': 'بازیافت شده (گرم)',
                'date': 'تاریخ بازیافت',
                'notes': 'توضیحات'
            })
            st.table(display_recycled_df)
        else:
            st.info("هنوز رکوردی برای پودر بازیافت شده ثبت نشده است.")

    conn.close()

# ---------------------------------------------------------
# ۳. فرم ثبت تولید
# ---------------------------------------------------------
elif choice == "🏭 ۲. فرم ثبت تولید":
    st.header("🏭 فرم رکورد تولید (Production Form.xlsx)")
    
    conn = get_db_connection()
    powders_list = pd.read_sql_query("SELECT powder_code FROM powders", conn)['powder_code'].tolist()
    
    if not powders_list:
        st.warning("جهت ثبت فرم تولید، ابتدا باید حداقل یک پودر در بخش ۱ ثبت شده باشد.")
    else:
        with st.form("prod_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                part_code = st.text_input("کد قطعه (شناسه یکتا)")
                part_name = st.text_input("نام قطعه")
            with col2:
                powder_code = st.selectbox("شماره ظرف پودر مصرفی", powders_list)
                machine_model = st.selectbox("مدل دستگاه", ["M120", "M300", "سایر"])
            with col3:
                quantity = st.number_input("تعداد روی صفحه ساخت", min_value=1, value=1)
                date_str = st.date_input("تاریخ ساخت").strftime("%Y-%m-%d")
                
            st.markdown("---")
            st.subheader("⏱️ ۱- زمان آماده‌سازی و ساخت")
            tc1, tc2, tc3 = st.columns(3)
            with tc1:
                build_time_hrs = st.number_input("زمان تولید (ساعت)", min_value=0.0)
                downtime_hrs = st.number_input("زمان توقف حین ساخت (ساعت)", min_value=0.0)
            with tc2:
                start_date = st.text_input("تاریخ شروع", value=date_str)
                start_time = st.text_input("ساعت شروع", value="08:00")
                end_date = st.text_input("تاریخ پایان", value=date_str)
                end_time = st.text_input("ساعت پایان", value="16:00")
            with tc3:
                setup_time_hrs = st.number_input("زمان آماده‌سازی دستگاه (ساعت)", min_value=0.0)
                cleaning_time_hrs = st.number_input("زمان تمیزکاری دستگاه (ساعت)", min_value=0.0)
                
            st.markdown("---")
            st.subheader("⚖️ ۲- پارامترهای متریال و وزن")
            mc1, mc2 = st.columns(2)
            with mc1:
                input_powder_g = st.number_input("پودر ورودی به دستگاه (گرم)", min_value=0.0)
                waste_powder_g = st.number_input("پودر مصرف شده غیر قابل بازیافت (گرم)", min_value=0.0)
                part_with_support_g = st.number_input("وزن قطعه با ساپورت (گرم)", min_value=0.0)
                final_part_g = st.number_input("وزن قطعه نهایی (گرم)", min_value=0.0)
                filter_pct = st.number_input("درصد فیلتر دستگاه (%)", min_value=0.0, max_value=100.0)
            with mc2:
                plate_code = st.text_input("کد صفحه ساخت")
                plate_init_wt = st.number_input("وزن اولیه صفحه ساخت (گرم)", min_value=0.0)
                plate_post_wt = st.number_input("وزن صفحه ساخت پس از پرداخت (گرم)", min_value=0.0)
                
            submitted = st.form_submit_button("💾 ثبت رکورد تولید")
            if submitted:
                if part_code:
                    c = conn.cursor()
                    c.execute("""INSERT OR REPLACE INTO production 
                        (part_code, part_name, powder_code, quantity, machine_model, date,
                         build_time_hrs, downtime_hrs, start_date, start_time, end_date, end_time,
                         setup_time_hrs, cleaning_time_hrs, input_powder_g, waste_powder_g,
                         part_with_support_g, final_part_g, filter_percentage, build_plate_code,
                         build_plate_init_wt_g, build_plate_post_wt_g)
                        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                        (part_code, part_name, powder_code, quantity, machine_model, date_str,
                         build_time_hrs, downtime_hrs, start_date, start_time, end_date, end_time,
                         setup_time_hrs, cleaning_time_hrs, input_powder_g, waste_powder_g,
                         part_with_support_g, final_part_g, filter_pct, plate_code,
                         plate_init_wt, plate_post_wt))
                    conn.commit()
                    st.success(f"اطلاعات تولید قطعه {part_code} ثبت شد.")
                else:
                    st.error("لطفاً کد قطعه را مشخص کنید.")
    conn.close()

# ---------------------------------------------------------
# ۴. فرم کنترل کیفیت (QC)
# ---------------------------------------------------------
elif choice == "🔬 ۳. فرم کنترل کیفیت (QC)":
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
                with q3: note = st.text_input("ملاحظات", key=f"qc_n_{t_title}")
                qc_data[t_title] = {"result": res, "type": t_type, "note": note}
                
            st.markdown("---")
            st.subheader("👥 مسئولین و تاییدکنندگان")
            sc1, sc2, sc3 = st.columns(3)
            with sc1: inspector = st.text_input("بازرس کنترل کیفیت")
            with sc2: engineer = st.text_input("مسئول فنی / مهندسی کیفیت")
            with sc3: manager = st.text_input("مدیر تضمین کیفیت")
            
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
# ۵. محاسبه‌گر قیمت قطعه
# ---------------------------------------------------------
elif choice == "💰 ۴. محاسبه‌گر قیمت قطعه":
    st.header("💰 محاسبه‌گر بهای تمام شده و قیمت فروش (Cost Calculator.xlsx)")
    
    RATES = {
        "powder_price_per_kg": {"Steel 316": 120000000, "Ti6Al4V": 500000000, "Inconel 718": 300000000, "Hastelloy X": 400000000},
        "machine_depreciation_hr": {"M120": 1250000, "M300": 3125000},
        "power_kw_hr": {"M120": 7, "M300": 15},
        "wages": {"designer_hr": 2500000, "operator_hr": 1870000, "qc_hr": 2180000},
        "argon_hr": 300000,
        "electricity_kwh": 50000,
        "post_process_hr": 500000,
        "qc_fixed_cost": 40000000,
        "ventilation_hr": 200000,
        "chiller_hr": 2500,
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
        p_code = st.text_input("شناسه/کد قطعه", value=selected_part if selected_part != "جدید" else "")
        p_name = st.text_input("نام قطعه", value=def_part_name)
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
# ۶. خروجی و گزارش‌گیری اکسل (حذف ستون‌های JSON خام و ترجمه کامل فارسی)
# ---------------------------------------------------------
elif choice == "📂 ۵. خروجی و گزارش‌گیری اکسل":
    st.header("📂 بایگانی اطلاعات و خروجی گزارش‌ها")
    
    conn = get_db_connection()
    
    target_table = st.selectbox("انتخاب جدول جهت مشاهده و دریافت Excel:", 
                                ["آنالیز پودر اولیه (powders)", "پودرهای بازیافت شده (recycled_powders)", "رکوردهای تولید (production)", "کنترل کیفیت (qc)", "محاسبه هزینه (cost_calculator)"])
    
    table_map = {
        "آنالیز پودر اولیه (powders)": "powders",
        "پودرهای بازیافت شده (recycled_powders)": "recycled_powders",
        "رکوردهای تولید (production)": "production",
        "کنترل کیفیت (qc)": "qc",
        "محاسبه هزینه (cost_calculator)": "cost_calculator"
    }
    
    selected_tbl = table_map[target_table]
    df = pd.read_sql_query(f"SELECT * FROM {selected_tbl}", conn)
    
    # حذف ستون‌های حاوی ساختار خام JSON از دید کاربر در جدول بایگانی
    json_cols_to_drop = ['checklist_json', 'qc_checks_json', 'finishing_json']
    clean_df = df.drop(columns=[col for col in json_cols_to_drop if col in df.columns])
    
    # دیکشنری نگاشت جامع تمام ستون‌های لاتین به فارسی
    farsi_headers_map = {
        # فرم پودر
        'powder_code': 'شماره ظرف پودر',
        'material': 'جنس پودر',
        'weight_g': 'وزن پودر (گرم)',
        'date': 'تاریخ ثبت',
        
        # پودر بازیافتی
        'id': 'شناسه',
        'recycled_batch_code': 'کد پارت بازیافت',
        'input_powder_g': 'پودر ورودی به دستگاه (گرم)',
        'unrecyclable_powder_g': 'پودر غیرقابل بازیافت (گرم)',
        'recycled_powder_g': 'پودر بازیافتی قابل استفاده (گرم)',
        'notes': 'توضیحات و ملاحظات',
        
        # فرم تولید
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
        
        # کنترل کیفیت QC
        'qc_inspector': 'بازرس کنترل کیفیت',
        'qc_engineer': 'مسئول مهندسی کیفیت',
        'qa_manager': 'مدیر تضمین کیفیت',
        
        # محاسبه‌گر هزینه
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
    
    if not clean_df.empty:
        farsi_df = clean_df.rename(columns=farsi_headers_map)
        st.table(farsi_df)
        
        excel_filename = f"{selected_tbl}_export.xlsx"
        farsi_df.to_excel(excel_filename, index=False)
        with open(excel_filename, "rb") as f:
            st.download_button(
                label="📥 دانلود فایل اکسل جدول انتخابی",
                data=f,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.info("اطلاعاتی در این جدول ثبت نشده است.")
        
    conn.close()
