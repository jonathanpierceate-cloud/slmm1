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

# استایل CSS سفارشی برای راست‌چین‌سازی (RTL) و زیباسازی UI
st.markdown("""
<style>
    @import url('https://v1.fontapi.ir/css/Vazir');
    html, body, [class*="css"] {
        font-family: 'Vazir', sans-serif;
        direction: rtl;
        text-align: right;
    }
    .stButton>button {
        width: 100%;
        background-color: #1f77b4;
        color: white;
        border-radius: 8px;
        font-weight: bold;
    }
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 15px;
        border-radius: 10px;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

# --- مدیریت پایگاه داده SQLite ---
DB_NAME = "slm_management.db"

def init_db():
    conn = sqlite3.connect(DB_NAME)
    c = conn.cursor()
    
    # جدول پودرها
    c.execute('''CREATE TABLE IF NOT EXISTS powders (
                    powder_code TEXT PRIMARY KEY,
                    material TEXT,
                    weight_g REAL,
                    date TEXT,
                    checklist_json TEXT
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
st.title("⚙️ سامانه یکپارچه مدیریت و بایگانی ساخت افزایشی (SLM)")
st.caption("اتصال خودکار فرم‌های پودر، تولید، کنترل کیفیت (QC) و برآورد قیمت قطعه")

menu = [
    "🔍 داشبورد و جستجوی جامع",
    "🧪 ۱. آنالیز و بایگانی پودر",
    "🏭 ۲. فرم ثبت تولید",
    "🔬 ۳. فرم کنترل کیفیت (QC)",
    "💰 ۴. محاسبه‌گر قیمت قطعه",
    "📂 ۵. خروجی و گزارش‌گیری اکسل"
]

choice = st.sidebar.selectbox("منوی دسترسی", menu)

# ---------------------------------------------------------
# ۱. داشبورد و جستجوی جامع
# ---------------------------------------------------------
if choice == "🔍 داشبورد و جستجوی جامع":
    st.header("🔍 جستجوی شناسنامه کامل قطعه")
    search_code = st.text_input("کد قطعه را جهت استعلام وارد کنید:")
    
    if search_code:
        conn = get_db_connection()
        prod_df = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(search_code,))
        qc_df = pd.read_sql_query("SELECT * FROM qc WHERE part_code=?", conn, params=(search_code,))
        cost_df = pd.read_sql_query("SELECT * FROM cost_calculator WHERE part_code=?", conn, params=(search_code,))
        
        if not prod_df.empty:
            st.success(f"اطلاعات قطعه {search_code} با موفقیت پیدا شد.")
            
            tab1, tab2, tab3, tab4 = st.tabs(["اطلاعات ساخت", "آنالیز پودر مصرفی", "کنترل کیفیت (QC)", "برآورد هزینه"])
            
            with tab1:
                st.subheader("مشخصات تولید")
                st.dataframe(prod_df.T, use_container_width=True)
                
            with tab2:
                powder_code = prod_df['powder_code'].values[0]
                powder_df = pd.read_sql_query("SELECT * FROM powders WHERE powder_code=?", conn, params=(powder_code,))
                if not powder_df.empty:
                    st.write(f"**کد پودر استفاده شده:** {powder_code}")
                    st.dataframe(powder_df.T, use_container_width=True)
                else:
                    st.warning("اطلاعات پودر متناظر یافت نشد.")
                    
            with tab3:
                if not qc_df.empty:
                    st.dataframe(qc_df.T, use_container_width=True)
                else:
                    st.info("فرم کنترل کیفیت برای این قطعه هنوز تکمیل نشده است.")
                    
            with tab4:
                if not cost_df.empty:
                    st.dataframe(cost_df.T, use_container_width=True)
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
    st.header("🧪 فرم آنالیز پودر خریداری شده")
    
    conn = get_db_connection()
    powders_df = pd.read_sql_query("SELECT * FROM powders", conn)
    
    with st.expander("➕ افزودن / ویرایش ظرف پودر جدید", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            powder_code = st.text_input("کد ظرف پودر فلز (مانند: PWD-316-01)")
            material = st.selectbox("جنس پودر", ["Steel 316", "Ti6Al4V", "Inconel 718", "Hastelloy X", "سایر"])
        with col2:
            weight_g = st.number_input("وزن پودر (گرم)", min_value=0.0, value=10000.0)
            date_str = st.date_input("تاریخ ثبت").strftime("%Y-%m-%d")
            
        st.markdown("### آزمون‌های کیفیت پودر")
        checklist = [
            "بررسی خلوص شیمیایی", "اندازه‌گیری میزان رطوبت", "اندازه‌گیری چگالی ضربه‌ای",
            "اندازه‌گیری چگالی ظاهری", "بررسی مورفولوژی ذرات", "آنالیز توزیع اندازه ذرات",
            "وضعیت بسته‌بندی و عدم آلودگی", "بررسی گواهینامه کیفیت تامین‌کننده"
        ]
        
        qc_results = {}
        for item in checklist:
            c1, c2, c3 = st.columns([2, 1, 2])
            with c1: st.write(item)
            with c2: status = st.selectbox(f"وضعیت {item}", ["تایید", "رد"], key=item)
            with c3: note = st.text_input(f"توضیحات {item}", key=f"note_{item}")
            qc_results[item] = {"status": status, "note": note}
            
        if st.button("ذخیره در بایگانی پودر"):
            if powder_code:
                c = conn.cursor()
                c.execute("""INSERT OR REPLACE INTO powders 
                             (powder_code, material, weight_g, date, checklist_json)
                             VALUES (?, ?, ?, ?, ?)""",
                          (powder_code, material, weight_g, date_str, json.dumps(qc_results, ensure_ascii=False)))
                conn.commit()
                st.success(f"پودر با کد {powder_code} با موفقیت ثبت شد.")
                st.rerun()
            else:
                st.error("لطفاً کد ظرف پودر را وارد کنید.")
                
    st.subheader("📋 بایگانی پودرهای ثبت شده")
    st.dataframe(powders_df[['powder_code', 'material', 'weight_g', 'date']], use_container_width=True)
    conn.close()

# ---------------------------------------------------------
# ۳. فرم ثبت تولید
# ---------------------------------------------------------
elif choice == "🏭 ۲. فرم ثبت تولید":
    st.header("🏭 فرم رکورد تولید (Production Form)")
    
    conn = get_db_connection()
    powders_list = pd.read_sql_query("SELECT powder_code FROM powders", conn)['powder_code'].tolist()
    
    if not powders_list:
        st.warning("ابتدا باید حداقل یک پودر در سیستم ثبت کنید.")
    else:
        with st.form("prod_form"):
            col1, col2, col3 = st.columns(3)
            with col1:
                part_code = st.text_input("کد قطعه (شناسه یکتا)")
                part_name = st.text_input("نام قطعه")
            with col2:
                powder_code = st.selectbox("انتخاب کد ظرف پودر فلز", powders_list)
                machine_model = st.selectbox("مدل دستگاه", ["M120", "M300", "سایر"])
            with col3:
                quantity = st.number_input("تعداد روی صفحه ساخت", min_value=1, value=1)
                date_str = st.date_input("تاریخ ساخت").strftime("%Y-%m-%d")
                
            st.subheader("۱- زمان آماده‌سازی و ساخت")
            tc1, tc2, tc3 = st.columns(3)
            with tc1:
                build_time_hrs = st.number_input("زمان تولید قطعه (ساعت)", min_value=0.0)
                downtime_hrs = st.number_input("زمان توقف حین ساخت (ساعت)", min_value=0.0)
            with tc2:
                start_date = st.text_input("تاریخ شروع", value=date_str)
                start_time = st.text_input("ساعت شروع", value="08:00")
                end_date = st.text_input("تاریخ پایان", value=date_str)
                end_time = st.text_input("ساعت پایان", value="16:00")
            with tc3:
                setup_time_hrs = st.number_input("زمان آماده‌سازی دستگاه (ساعت)", min_value=0.0)
                cleaning_time_hrs = st.number_input("زمان تمیزکاری دستگاه (ساعت)", min_value=0.0)
                
            st.subheader("۲- متریال و وزنی")
            mc1, mc2 = st.columns(2)
            with mc1:
                input_powder_g = st.number_input("پودر ورودی (گرم)", min_value=0.0)
                waste_powder_g = st.number_input("پودر ضایعات / غیرقابل بازیافت (گرم)", min_value=0.0)
                part_with_support_g = st.number_input("وزن قطعه با ساپورت (گرم)", min_value=0.0)
                final_part_g = st.number_input("وزن قطعه نهایی (گرم)", min_value=0.0)
                filter_pct = st.number_input("درصد فیلتر دستگاه (%)", min_value=0.0, max_value=100.0)
            with mc2:
                plate_code = st.text_input("کد صفحه ساخت")
                plate_init_wt = st.number_input("وزن اولیه صفحه ساخت (گرم)", min_value=0.0)
                plate_post_wt = st.number_input("وزن صفحه ساخت پس از پرداخت (گرم)", min_value=0.0)
                
            submitted = st.form_submit_button("ذخیره رکورد تولید")
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
                    st.success(f"فرم تولید قطعه {part_code} ثبت شد.")
                else:
                    st.error("وارد کردن کد قطعه الزامی است.")
    conn.close()

# ---------------------------------------------------------
# ۴. فرم کنترل کیفیت (QC)
# ---------------------------------------------------------
elif choice == "🔬 ۳. فرم کنترل کیفیت (QC)":
    st.header("🔬 فرم کنترل کیفیت (QC)")
    
    conn = get_db_connection()
    parts_list = pd.read_sql_query("SELECT part_code, part_name FROM production", conn)
    
    if parts_list.empty:
        st.warning("هیچ قطعه‌ای در بخش تولید ثبت نشده است.")
    else:
        selected_part = st.selectbox("انتخاب کد قطعه جهت بررسی QC", parts_list['part_code'])
        part_info = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(selected_part,)).iloc[0]
        
        st.info(f"نام قطعه: {part_info['part_name']} | مدل دستگاه: {part_info['machine_model']}")
        
        with st.form("qc_form"):
            st.subheader("چک‌لیست آزمون‌ها")
            tests = [
                ("1.1.1 ظاهر سطح و عیوب قابل رؤیت", "چشمی"),
                ("1.1.2 کیفیت حذف ساپورت", "چشمی"),
                ("1.1.3 زبری سطح (Ra)", "ابزار زبری‌سنج"),
                ("2.1.1 ابعاد بحرانی", "ککولیس/CMM"),
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
                q1, q2, q3 = st.columns([3, 2, 3])
                with q1: st.write(t_title)
                with q2: res = st.selectbox("وضعیت", ["تایید", "رد"], key=t_title)
                with q3: note = st.text_input("ملاحظات", key=f"qc_n_{t_title}")
                qc_data[t_title] = {"result": res, "type": t_type, "note": note}
                
            st.subheader("تاییدکنندگان")
            sc1, sc2, sc3 = st.columns(3)
            with sc1: inspector = st.text_input("بازرس کنترل کیفیت")
            with sc2: engineer = st.text_input("مسئول فنی / مهندسی کیفیت")
            with sc3: manager = st.text_input("مدیر تضمین کیفیت")
            
            qc_submit = st.form_submit_button("ثبت نتیجه QC")
            if qc_submit:
                c = conn.cursor()
                c.execute("""INSERT OR REPLACE INTO qc 
                             (part_code, part_name, material, machine_model, date, qc_checks_json, qc_inspector, qc_engineer, qa_manager)
                             VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                          (selected_part, part_info['part_name'], part_info['powder_code'], part_info['machine_model'],
                           datetime.now().strftime("%Y-%m-%d"), json.dumps(qc_data, ensure_ascii=False), inspector, engineer, manager))
                conn.commit()
                st.success("فرم QC با موفقیت ذخیره شد.")
    conn.close()

# ---------------------------------------------------------
# ۵. محاسبه‌گر قیمت قطعه
# ---------------------------------------------------------
elif choice == "💰 ۴. محاسبه‌گر قیمت قطعه":
    st.header("💰 محاسبه‌گر قیمت و بهای تمام شده SLM")
    
    # نرخ‌های پایه ثابت طبق فایل اکسل Cost Calculator
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
    
    col1, col2 = st.columns(2)
    with col1:
        selected_part = st.selectbox("انتخاب کد قطعه ثبت‌شده (جهت فراخوانی خودکار)", ["جدید"] + list(parts_list['part_code']))
    
    # مقادیر پیش‌فرض
    def_part_name = ""
    def_machine = "M300"
    def_powder = "Steel 316"
    def_print_time = 10.0
    
    if selected_part != "جدید":
        p_row = pd.read_sql_query("SELECT * FROM production WHERE part_code=?", conn, params=(selected_part,)).iloc[0]
        def_part_name = p_row['part_name']
        def_machine = p_row['machine_model'] if p_row['machine_model'] in ["M120", "M300"] else "M300"
        def_print_time = float(p_row['build_time_hrs']) if p_row['build_time_hrs'] else 10.0

    st.subheader("مشخصات ورودی قطعه")
    c1, c2, c3 = st.columns(3)
    with c1:
        p_code = st.text_input("شناسه/کد قطعه", value=selected_part if selected_part != "جدید" else "")
        p_name = st.text_input("نام قطعه", value=def_part_name)
        powder_type = st.selectbox("نوع پودر فلزی", list(RATES["density"].keys()))
    with c2:
        vol_cm3 = st.number_input("حجم قطعه (cm3)", min_value=0.0, value=50.0)
        sup_vol_cm3 = st.number_input("حجم ساپورت (cm3)", min_value=0.0, value=10.0)
        machine_type = st.selectbox("نوع دستگاه چاپی", ["M120", "M300"], index=1 if def_machine=="M300" else 0)
    with c3:
        parts_on_plate = st.number_input("تعداد قطعات روی صفحه", min_value=1, value=1)
        print_time_hrs = st.number_input("زمان کل چاپ (ساعت)", min_value=0.0, value=def_print_time)
        design_time_hrs = st.number_input("زمان طراحی/آماده‌سازی (ساعت)", min_value=0.0, value=2.0)
        post_time_hrs = st.number_input("زمان پرداخت و فینیشینگ (ساعت)", min_value=0.0, value=3.0)
        overhead_pct = st.number_input("ضریب هزینه سربار (%)", min_value=0.0, value=35.0)

    # محاسبات لحظه‌ای
    density = RATES["density"][powder_type]
    net_weight_g = vol_cm3 * density
    support_weight_g = sup_vol_cm3 * density
    total_weight_kg = (net_weight_g + support_weight_g) / 1000.0
    
    # محاسبه ریز هزینه‌ها
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
    st.subheader("📊 خلاصه برآورد مالی")
    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("وزن خالص قطعه", f"{net_weight_g:.1f} گرم")
    mc2.metric("بهای تمام شده تولید", f"{total_production_cost:,.0f} ریال")
    mc3.metric("قیمت نهایی به مشتری", f"{final_price:,.0f} ریال")
    
    if st.button("ذخیره برآورد قیمت در پایگاه داده"):
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
# ۶. خروجی و گزارش‌گیری اکسل
# ---------------------------------------------------------
elif choice == "📂 ۵. خروجی و گزارش‌گیری اکسل":
    st.header("📂 مدیریت بایگانی و دریافت خروجی Excel")
    
    conn = get_db_connection()
    
    target_table = st.selectbox("جدول مورد نظر را انتخاب کنید:", 
                                ["آنالیز پودر (powders)", "رکوردهای تولید (production)", "کنترل کیفیت (qc)", "محاسبه هزینه (cost_calculator)"])
    
    table_map = {
        "آنالیز پودر (powders)": "powders",
        "رکوردهای تولید (production)": "production",
        "کنترل کیفیت (qc)": "qc",
        "محاسبه هزینه (cost_calculator)": "cost_calculator"
    }
    
    df = pd.read_sql_query(f"SELECT * FROM {table_map[target_table]}", conn)
    st.dataframe(df, use_container_width=True)
    
    # خروجی Excel
    if not df.empty:
        excel_filename = f"{table_map[target_table]}_export.xlsx"
        df.to_excel(excel_filename, index=False)
        with open(excel_filename, "rb") as f:
            st.download_button(
                label="📥 دانلود خروجی اکسل این جدول",
                data=f,
                file_name=excel_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    conn.close()