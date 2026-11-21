import pandas as pd
import streamlit as st
from datetime import date
import hashlib
import os
from sqlalchemy import create_engine, text

# يقرأ رابط قاعدة البيانات السحابية من إعدادات Streamlit Secrets بأمان
DB_URL = st.secrets["DATABASE_URL"]
engine = create_engine(DB_URL, pool_pre_ping=True)

TYPES = ["أصول", "خصوم", "حقوق ملكية", "إيرادات", "مصروفات"]

DEFAULT_ACCOUNTS = [
    ("1", "الأصول", "أصول", None, 0),
    ("11", "الأصول المتداولة والمشروعات", "أصول", 1, 0),
    ("111", "الصندوق / الخزينة", "أصول", 2, 1),
    ("112", "البنك", "أصول", 2, 1),
    ("113", "عملاء وأقساط مستحقة (أوراق قبض)", "أصول", 2, 1),
    ("114", "أراضي بغرض التطوير", "أصول", 2, 1),
    ("115", "مشروعات تحت التنفيذ (إنشاءات ومرافق)", "أصول", 2, 1),
    ("116", "مخزون الوحدات الجاهزة للبيع", "أصول", 2, 1),
    ("117", "عهد وسلف مهندسين ومواقع", "أصول", 2, 1),
    ("12", "الأصول الثابتة", "أصول", 1, 0),
    ("121", "مقر الشركة والمباني الإدارية", "أصول", 10, 1),
    ("122", "السيارات والمعدات الهندسية", "أصول", 10, 1),
    ("123", "الأثاث والأجهزة المكتبية", "أصول", 10, 1),
    ("2", "الخصوم والالتزامات", "خصوم", None, 0),
    ("21", "الخصوم المتداولة", "خصوم", 14, 0),
    ("211", "مقاولي الباطن والموردين", "خصوم", 15, 1),
    ("212", "دفعات مقدمة من العملاء (حجوزات تحت التسليم)", "خصوم", 15, 1),
    ("213", "تأمينات محجوزة لمقاولي الباطن", "خصوم", 15, 1),
    ("214", "مصروفات مستحقة", "خصوم", 15, 1),
    ("22", "الخصوم غير المتداولة", "خصوم", 14, 0),
    ("221", "قروض وتسهيلات تمويل عقاري", "خصوم", 20, 1),
    ("3", "حقوق الملكية", "حقوق ملكية", None, 0),
    ("31", "رأس المال", "حقوق ملكية", 22, 1),
    ("32", "الأرباح المبقاة / المدورة", "حقوق ملكية", 22, 1),
    ("33", "جاري الشركاء / المسحوبات", "حقوق ملكية", 22, 1),
    ("4", "الإيرادات", "إيرادات", None, 0),
    ("41", "إيرادات مبيعات وحدات عقارية", "إيرادات", 26, 1),
    ("42", "إيرادات إيجارات واستثمارات عقارية", "إيرادات", 26, 1),
    ("43", "إيرادات تنازلات وخدمات إدارية", "إيرادات", 26, 1),
    ("5", "المصروفات والتكاليف", "مصروفات", None, 0),
    ("51", "تكلفة الوحدات العقارية المباعة", "مصروفات", 30, 1),
    ("52", "المصروفات البيعية والعمومية", "مصروفات", 30, 0),
    ("521", "الرواتب والأجور الهندسية والإدارية", "مصروفات", 32, 1),
    ("522", "عمولات بيعية وتسويق عقاري", "مصروفات", 32, 1),
    ("523", "إيجارات ومرافق وخدمات", "مصروفات", 32, 1),
    ("524", "رسوم تراخيص واستشارات هندسية", "مصروفات", 32, 1)
]
def make_hashes(password):
    return hashlib.sha256(str.encode(password)).hexdigest()

def check_hashes(password, hashed_text):
    return make_hashes(password) == hashed_text

def to_excel(df):
    return df.to_csv(index=False).encode('utf-8-sig')

def generate_entry_html(entry_id, entry_date, description, debit_acc, credit_acc, amount):
    html_content = f"""
    <!DOCTYPE html>
    <html dir="rtl" lang="ar">
    <head>
        <meta charset="UTF-8">
        <title>سند قيد رقم #{entry_id}</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 40px; line-height: 1.6; }}
            .voucher-box {{ border: 2px solid #333; padding: 20px; border-radius: 8px; max-width: 700px; margin: auto; }}
            .header {{ text-align: center; border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 20px; }}
            .row {{ display: flex; justify-content: space-between; margin-bottom: 10px; }}
            table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
            th, td {{ border: 1px solid #ddd; padding: 10px; text-align: right; }}
            th {{ background-color: #f2f2f2; }}
            .footer {{ margin-top: 40px; display: flex; justify-content: space-between; font-weight: bold; }}
        </style>
    </head>
    <body>
        <div class="voucher-box">
            <div class="header">
                <h2>سند قيد محاسبي / Accounting Voucher</h2>
            </div>
            <div class="row">
                <span><strong>رقم القيد:</strong> #{entry_id}</span>
                <span><strong>التاريخ:</strong> {entry_date}</span>
            </div>
            <div class="row">
                <span><strong>البيان:</strong> {description}</span>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>البيان / الحساب</th>
                        <th>مدين</th>
                        <th>دائن</th>
                    </tr>
                </thead>
                <tbody>
                    <tr>
                        <td>من حـ/ {debit_acc}</td>
                        <td>{amount:,.2f}</td>
                        <td>0.00</td>
                    </tr>
                    <tr>
                        <td>إلى حـ/ {credit_acc}</td>
                        <td>0.00</td>
                        <td>{amount:,.2f}</td>
                    </tr>
                </tbody>
            </table>
            <div class="footer">
                <div>إعداد: ________________</div>
                <div>اعتماد: ________________</div>
            </div>
        </div>
    </body>
    </html>
    """
    return html_content.encode('utf-8')

def init_db():
    with engine.begin() as conn:
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL
            )
        '''))
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS accounts (
                id SERIAL PRIMARY KEY,
                code TEXT UNIQUE,
                name TEXT NOT NULL,
                account_type TEXT NOT NULL,
                parent_id INTEGER,
                allow_posting INTEGER DEFAULT 1,
                balance REAL DEFAULT 0.0
            )
        '''))
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS projects (
                id SERIAL PRIMARY KEY,
                name TEXT UNIQUE NOT NULL
            )
        '''))
        conn.execute(text('''
            CREATE TABLE IF NOT EXISTS journal_entries (
                id SERIAL PRIMARY KEY,
                entry_date TEXT NOT NULL,
                description TEXT,
                project_name TEXT DEFAULT 'عام / الإدارة'
            )
        '''))
conn.execute(text("ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS project_name TEXT DEFAULT 'عام / الإدارة'"))
        conn.execute(text("INSERT INTO projects (name) VALUES ('عام / الإدارة') ON CONFLICT DO NOTHING"))
conn.execute(text('''       
            CREATE TABLE IF NOT EXISTS journal_details (
                id SERIAL PRIMARY KEY,
                entry_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
                account_id INTEGER NOT NULL REFERENCES accounts(id),
                debit REAL DEFAULT 0.0,
                credit REAL DEFAULT 0.0
            )
        '''))
        res = conn.execute(text("SELECT COUNT(*) FROM users")).scalar()
        if res == 0:
            conn.execute(
                text("INSERT INTO users (username, password, role) VALUES (:u, :p, :r)"),
                {"u": "admin", "p": make_hashes("admin123"), "r": "مدير النظام"}
            )

def load_default_accounts():
    with engine.begin() as conn:
        for code, name, acc_type, parent_id, allow_post in DEFAULT_ACCOUNTS:
            conn.execute(
                text("""
                    INSERT INTO accounts (code, name, account_type, parent_id, allow_posting)
                    VALUES (:c, :n, :t, :p, :a)
                    ON CONFLICT (code) DO NOTHING
                """),
                {"c": code, "n": name, "t": acc_type, "p": parent_id, "a": allow_post}
            )

init_db()

st.set_page_config(page_title="برنامجي المحاسبي السحابي", layout="wide")

if 'logged_in' not in st.session_state:
    st.session_state['logged_in'] = False
    st.session_state['user_role'] = ''
    st.session_state['username'] = ''

if not st.session_state['logged_in']:
    col_l1, col_l2, col_l3 = st.columns([1, 2, 1])
    with col_l2:
        st.title("🔒 تسجيل الدخول للنظام السحابي")
        username = st.text_input("اسم المستخدم")
        password = st.text_input("كلمة السر", type="password")
        
        if st.button("🔑 دخول إلى النظام", type="primary", use_container_width=True):
            if not username or not password:
                st.error("يرجى إدخال اسم المستخدم وكلمة السر!")
            else:
                with engine.connect() as conn:
                    user_res = conn.execute(
                        text("SELECT password, role FROM users WHERE username = :u"),
                        {"u": username}
                    ).fetchone()

                if user_res and check_hashes(password, user_res[0]):
                    st.session_state['logged_in'] = True
                    st.session_state['username'] = username
                    st.session_state['user_role'] = user_res[1]
                    st.success(f"مرحباً بك {username}!")
                    st.rerun()
                else:
                    st.error("❌ اسم المستخدم أو كلمة السر غير صحيحة!")

else:
    st.sidebar.markdown(f"👤 **المستخدم الحالي:** {st.session_state['username']}")
    st.sidebar.markdown(f"🔰 **الصلاحية:** {st.session_state['user_role']}")
    st.sidebar.success("☁️ متصل بقاعدة البيانات السحابية")
    
    if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state['logged_in'] = False
        st.session_state['user_role'] = ''
        st.session_state['username'] = ''
        st.rerun()

    st.title("📊 برنامجي المحاسبي السحابي")

    tabs_list = [
        "🌳 شجرة الحسابات", 
        "🏁 الأرصدة الافتتاحية",
        "📝 إضافة قيد يومي", 
        "📖 دفتر اليومية والطباعة", 
        "⚖️ ميزان المراجعة", 
        "📑 كشف حساب",
        "📈 قائمة الدخل",
        "🏛️ الميزانية العمومية"
        "🏗️ المشاريع ومراكز التكلفة"
    ]
    
    if st.session_state['user_role'] == "مدير النظام":
        tabs_list.append("👥 إدارة المستخدمين")
        tabs_list.append("🤖 الـ AI وتحديث الكود")

    tabs = st.tabs(tabs_list)

    # --- 1. شجرة الحسابات ---
    with tabs[0]:
        st.subheader("🌳 إدارة وتأسيس شجرة الحسابات")
        if st.button("⚡ انشاء شجرة حسابات قياسية كاملاً", type="primary"):
            load_default_accounts()
            st.success("✅ تم بناء شجرة الحسابات القياسية بنجاح!")
            st.rerun()

        st.markdown("---")
        col_acc1, col_acc2 = st.columns(2)
        
        with col_acc1:
            st.subheader("➕ إضافة حساب جديد")
            accounts_df = pd.read_sql("SELECT id, code, name FROM accounts", engine)
            
            parent_options = {"بدون (حساب رئيسي)": None}
            for _, row in accounts_df.iterrows():
                parent_options[f"{row['code']} - {row['name']}"] = int(row['id'])
                
            with st.form("account_form", clear_on_submit=True):
                code = st.text_input("كود الحساب (مثال: 1111)")
                name = st.text_input("اسم الحساب (مثال: خزانة الفرع الرئيسي)")
                account_type = st.selectbox("نوع الحساب", TYPES)
                parent_label = st.selectbox("الحساب الأب", list(parent_options.keys()))
                allow_post = st.selectbox("طبيعة التسجيل:", ["متاح للتسجيل (فرعي)", "ممنوع التسجيل (رئيسي)"])
                    
                submitted = st.form_submit_button("حفظ الحساب")
                if submitted:
                    if not code or not name:
                        st.error("يرجى إدخال الكود والاسم!")
                    else:
                        try:
                            parent_id = parent_options[parent_label]
                            allow_val = 1 if allow_post == "متاح للتسجيل (فرعي)" else 0
                            with engine.begin() as conn:
                                conn.execute(
                                    text("INSERT INTO accounts (code, name, account_type, parent_id, allow_posting) VALUES (:c, :n, :t, :p, :a)"),
                                    {"c": code, "n": name, "t": account_type, "p": parent_id, "a": allow_val}
                                )
                            st.success(f"تم إضافة الحساب '{name}' بنجاح!")
                            st.rerun()
                        except Exception:
                            st.error("كود الحساب هذا مُسجّل مسبقاً!")

        with col_acc2:
            st.subheader("⚙️ تعديل بيانات الحساب (الاسم / الكود / الحالة)")
            all_accs = pd.read_sql("SELECT id, code, name, parent_id, allow_posting FROM accounts ORDER BY code ASC", engine)
            
            if not all_accs.empty:
                acc_options = {f"{row['code']} - {row['name']}": (int(row['id']), row['code'], row['name'], int(row['allow_posting'])) for _, row in all_accs.iterrows()}
                selected_edit_acc = st.selectbox("اختر الحساب المراد تعديله:", list(acc_options.keys()), key="edit_acc_sel")
                
                target_id, target_code, target_name, curr_posting = acc_options[selected_edit_acc]
                curr_status_index = 0 if curr_posting == 1 else 1

                with st.form("edit_details_form"):
                    updated_code = st.text_input("كود الحساب المعدل", value=target_code)
                    updated_name = st.text_input("اسم الحساب المعدل", value=target_name)
                    new_status = st.radio("طبيعة التسجيل لهذا الحساب:", ["متاح للتسجيل (فرعي)", "ممنوع التسجيل (رئيسي)"], index=curr_status_index)
                    
                    save_update = st.form_submit_button("💾 حفظ التعديلات")
                    
                    if save_update:
                        if not updated_code or not updated_name:
                            st.error("يرجى إدخال كود واسم الحساب!")
                        else:
                            new_val = 1 if new_status == "متاح للتسجيل (فرعي)" else 0
                            with engine.connect() as conn:
                                entries_count = conn.execute(
                                    text("SELECT COUNT(*) FROM journal_details WHERE account_id = :id"),
                                    {"id": target_id}
                                ).scalar()
                            
                            if new_val == 0 and curr_posting == 1 and entries_count > 0:
                                st.error("🛑 لا يمكن تغيير طبيعة الحساب لـ [ممنوع التسجيل] لأن عليه قيوداً مسجلة!")
                            else:
                                try:
                                    with engine.begin() as conn:
                                        conn.execute(
                                            text("UPDATE accounts SET code = :c, name = :n, allow_posting = :a WHERE id = :id"),
                                            {"c": updated_code, "n": updated_name, "a": new_val, "id": target_id}
                                        )
                                    st.success(f"✅ تم تحديث بيانات الحساب إلى '{updated_name}' بنجاح!")
                                    st.rerun()
                                except Exception:
                                    st.error("كود الحساب الجديد مُسجل مسبقاً لحساب آخر!")
            else:
                st.info("لا توجد حسابات مسجلة بعد.")

        st.markdown("---")
        st.subheader("📋 دليل شجرة الحسابات المعتمدة")
        df_acc = pd.read_sql("""
            SELECT a.code AS "كود الحساب", a.name AS "اسم الحساب", a.account_type AS "النوع", 
                   COALESCE(p.name, 'رئيسي') AS "الحساب الأب",
                   CASE WHEN a.allow_posting = 1 THEN '✅ متاح للتسجيل' ELSE '❌ ممنوع التسجيل (رئيسي)' END AS "حالة التسجيل"
            FROM accounts a
            LEFT JOIN accounts p ON a.parent_id = p.id
            ORDER BY a.code ASC
        """, engine)
        
        if not df_acc.empty:
            st.dataframe(df_acc, use_container_width=True)

    # --- 2. الأرصدة الافتتاحية ---
    with tabs[1]:
        st.subheader("🏁 إدخال الأرصدة الافتتاحية")
        acc_list = pd.read_sql("SELECT id, code, name, account_type, allow_posting FROM accounts ORDER BY code ASC", engine)

        if acc_list.empty:
            st.warning("⚠️ يرجى إنشاء شجرة الحسابات أولاً.")
        else:
            acc_dict = {}
            for _, row in acc_list.iterrows():
                status_str = "✅ متاح" if row['allow_posting'] == 1 else "❌ ممنوع التسجيل"
                acc_dict[f"{row['code']} - {row['name']} ({status_str})"] = (int(row['id']), row['name'], int(row['allow_posting']))
            
            with st.form("opening_balance_form", clear_on_submit=True):
                col_ob1, col_ob2 = st.columns(2)
                with col_ob1:
                    selected_acc_label = st.selectbox("الحساب:", list(acc_dict.keys()))
                    balance_type = st.radio("طبيعة الرصيد:", ["مدين (نقدية/أصول)", "دائن (رأس مال/موردين)"])
                with col_ob2:
                    ob_date = st.date_input("تاريخ الرصيد", value=date(2026, 1, 1))
                    ob_amount = st.number_input("المبلغ الافتتاحي", min_value=0.0, step=500.0)

                submit_ob = st.form_submit_button("💾 حفظ الرصيد الافتتاحي", type="primary")

                if submit_ob:
                    target_acc_id, target_acc_name, is_allowed = acc_dict[selected_acc_label]
                    if is_allowed == 0:
                        st.error("🛑 هذا الحساب مصنف [ممنوع التسجيل]! يرجى اختيار حساب فرعي متاح.")
                    elif ob_amount <= 0:
                        st.error("يرجى إدخال مبلغ أكبر من الصفر!")
                    else:
                        desc = f"رصيد افتتاحي - {target_acc_name}"
                        with engine.begin() as conn:
                            entry_id = conn.execute(
                                text("INSERT INTO journal_entries (entry_date, description) VALUES (:d, :desc) RETURNING id"),
                                {"d": str(ob_date), "desc": desc}
                            ).scalar()

                            if "مدين" in balance_type:
                                conn.execute(
                                    text("INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, :amt, 0)"),
                                    {"eid": entry_id, "aid": target_acc_id, "amt": ob_amount}
                                )
                            else:
                                conn.execute(
                                    text("INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, 0, :amt)"),
                                    {"eid": entry_id, "aid": target_acc_id, "amt": ob_amount}
                                )
                        st.success("✅ تم حفظ الرصيد الافتتاحي بنجاح!")
                        st.rerun()

            st.markdown("---")
            ob_df = pd.read_sql(text("""
                SELECT e.entry_date AS "التاريخ", a.code AS "كود الحساب", a.name AS "اسم الحساب",
                       d.debit AS "رصيد مدين", d.credit AS "رصيد دائن", e.description AS "البيان"
                FROM journal_entries e
                JOIN journal_details d ON e.id = d.entry_id
                JOIN accounts a ON d.account_id = a.id
                WHERE e.description LIKE 'رصيد افتتاحي%'
                ORDER BY e.id DESC
            """), engine)
            if not ob_df.empty:
                st.dataframe(ob_df, use_container_width=True)

    # --- 3. إضافة قيد يومي ---
    with tabs[2]:
        st.subheader("📝 تسجيل قيد محاسبي جديد")
        acc_list = pd.read_sql("SELECT id, code, name, allow_posting FROM accounts ORDER BY code ASC", engine)
        
        if not acc_list.empty:
            acc_dict = {}
            for _, row in acc_list.iterrows():
                status_str = "✅ متاح" if row['allow_posting'] == 1 else "❌ ممنوع التسجيل"
                acc_dict[f"{row['code']} - {row['name']} ({status_str})"] = (int(row['id']), int(row['allow_posting']))

            projects_df = pd.read_sql("SELECT name FROM projects ORDER BY id ASC", engine)
            projects_list = projects_df['name'].tolist() if not projects_df.empty else ["عام / الإدارة"]
            
            col_top1, col_top2 = st.columns(2)
            with col_top1:
                entry_date = st.date_input("تاريخ القيد", value=date.today())
            with col_top2:
                selected_project = st.selectbox("🏗️ المشروع / مركز التكلفة", projects_list)
                
            description = st.text_input("البيان / شرح القيد")
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                debit_acc = st.selectbox("حساب المدين", list(acc_dict.keys()), key="deb_acc")
                debit_val = st.number_input("المبلغ المدين", min_value=0.0, step=100.0, key="deb_val")
            with col_m2:
                credit_acc = st.selectbox("حساب الدائن", list(acc_dict.keys()), key="cred_acc")
                credit_val = st.number_input("المبلغ الدائن", min_value=0.0, step=100.0, key="cred_val")

            if st.button("حفظ القيد اليومي", type="primary"):
                deb_id, deb_allowed = acc_dict[debit_acc]
                cred_id, cred_allowed = acc_dict[credit_acc]

                if deb_allowed == 0:
                    st.error("🛑 حساب المدين المحدد [ممنوع التسجيل]! يرجى اختيار حساب فرعي.")
                elif cred_allowed == 0:
                    st.error("🛑 حساب الدائن المحدد [ممنوع التسجيل]! يرجى اختيار حساب فرعي.")
                elif debit_val <= 0 or credit_val <= 0:
                    st.error("يرجى إدخال مبالغ أكبر من الصفر.")
                elif debit_val != credit_val:
                    st.error("❌ القيد غير متوازن!")
                elif debit_acc == credit_acc:
                    st.error("❌ لا يمكن اختيار نفس الحساب!")
                else:
                    with engine.begin() as conn:
                        entry_id = conn.execute(
                            text("INSERT INTO journal_entries (entry_date, description, project_name) VALUES (:d, :desc, :proj) RETURNING id"),
                            {"d": str(entry_date), "desc": description, "proj": selected_project}
                        ).scalar()
                        conn.execute(
                            text("INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, :amt, 0)"),
                            {"eid": entry_id, "aid": deb_id, "amt": debit_val}
                        )
                        conn.execute(
                            text("INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, 0, :amt)"),
                            {"eid": entry_id, "aid": cred_id, "amt": credit_val}
                        )
                    st.success("✅ تم حفظ القيد اليومي بنجاح!")
                    st.rerun()

    # --- 4. دفتر اليومية والطباعة والحذف ---
    with tabs[3]:
        st.subheader("📖 دفتر اليومية وتصدير السندات")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            start_date_j = st.date_input("من تاريخ", value=date(2020, 1, 1), key="j_start")
        with col_f2:
            end_date_j = st.date_input("إلى تاريخ", value=date.today(), key="j_end")

        journal_df = pd.read_sql(
            text("""
                SELECT e.id AS "رقم القيد", e.entry_date AS "التاريخ", e.description AS "البيان",
                       a.code AS "كود الحساب", a.name AS "اسم الحساب", d.debit AS "مدين", d.credit AS "دائن"
                FROM journal_entries e
                JOIN journal_details d ON e.id = d.entry_id
                JOIN accounts a ON d.account_id = a.id
                WHERE e.entry_date BETWEEN :s AND :e
                ORDER BY e.id DESC
            """),
            engine,
            params={"s": str(start_date_j), "e": str(end_date_j)}
        )
        
        if not journal_df.empty:
            st.dataframe(journal_df, use_container_width=True)
            excel_journal = to_excel(journal_df)
            st.download_button(label="📥 تصدير دفتر اليومية إلى Excel", data=excel_journal, file_name="journal_entries.csv", mime="text/csv")

            st.markdown("---")
            col_p1, col_p2 = st.columns(2)
            
            with col_p1:
                st.subheader("🖨️ طباعة سند قيد / فاتورة")
                entry_ids = journal_df['رقم القيد'].unique()
                selected_entry_id = st.selectbox("اختر رقم القيد لطباعة السند:", entry_ids)

                if selected_entry_id:
                    e_rows = journal_df[journal_df['رقم القيد'] == selected_entry_id]
                    e_date = e_rows.iloc[0]['التاريخ']
                    e_desc = e_rows.iloc[0]['البيان']
                    deb_row = e_rows[e_rows['مدين'] > 0]
                    cred_row = e_rows[e_rows['دائن'] > 0]

                    if not deb_row.empty and not cred_row.empty:
                        deb_name = deb_row.iloc[0]['اسم الحساب']
                        cred_name = cred_row.iloc[0]['اسم الحساب']
                        amt = deb_row.iloc[0]['مدين']
                        html_data = generate_entry_html(selected_entry_id, e_date, e_desc, deb_name, cred_name, amt)
                        st.download_button(label=f"📄 تحميل سند قيد رقم #{selected_entry_id}", data=html_data, file_name=f"Voucher_{selected_entry_id}.html", mime="text/html", type="primary")

            with col_p2:
                st.subheader("🗑️ إلغاء / حذف قيد يومية")
                if st.session_state['user_role'] == "مدير النظام":
                    all_entry_ids = journal_df['رقم القيد'].unique()
                    entry_to_delete = st.selectbox("اختر رقم القيد المراد حذفه نهائياً:", all_entry_ids, key="del_entry_sel")
                    
                    if st.button("❌ حذف القيد المحدد الآن"):
                        with engine.begin() as conn:
                            conn.execute(text("DELETE FROM journal_details WHERE entry_id = :id"), {"id": int(entry_to_delete)})
                            conn.execute(text("DELETE FROM journal_entries WHERE id = :id"), {"id": int(entry_to_delete)})
                        st.success(f"✅ تم حذف القيد رقم #{entry_to_delete} بنجاح!")
                        st.rerun()
                else:
                    st.warning("🔒 إلغاء القيود متاح فقط لـ 'مدير النظام'.")

    # --- 5. ميزان المراجعة ---
    with tabs[4]:
        st.subheader("⚖️ ميزان المراجعة")
        col_tb1, col_tb2 = st.columns(2)
        with col_tb1:
            start_date_tb = st.date_input("من تاريخ", value=date(2020, 1, 1), key="tb_start")
        with col_tb2:
            end_date_tb = st.date_input("إلى تاريخ", value=date.today(), key="tb_end")

        tb_df = pd.read_sql(
            text("""
                SELECT a.code AS "كود الحساب", a.name AS "اسم الحساب", a.account_type AS "النوع",
                       COALESCE(SUM(d.debit), 0) AS "مجموع المدين", COALESCE(SUM(d.credit), 0) AS "مجموع الدائن",
                       (COALESCE(SUM(d.debit), 0) - COALESCE(SUM(d.credit), 0)) AS "الرصيد الصافي"
                FROM accounts a
                LEFT JOIN journal_details d ON a.id = d.account_id
                LEFT JOIN journal_entries e ON d.entry_id = e.id AND e.entry_date BETWEEN :s AND :e
                GROUP BY a.id, a.code, a.name, a.account_type
                ORDER BY a.code ASC
            """),
            engine,
            params={"s": str(start_date_tb), "e": str(end_date_tb)}
        )

        if not tb_df.empty:
            st.dataframe(tb_df, use_container_width=True)
            tot_debit = tb_df['مجموع المدين'].sum()
            tot_credit = tb_df['مجموع الدائن'].sum()
            
            col_t1, col_t2 = st.columns(2)
            col_t1.metric("إجمالي مجموع المدين", f"{tot_debit:,.2f}")
            col_t2.metric("إجمالي مجموع الدائن", f"{tot_credit:,.2f}")
            
            if abs(tot_debit - tot_credit) < 0.01:
                st.success("✅ ميزان المراجعة متوازن!")
            else:
                st.error("❌ يوجد عدم توازن في الميزان!")
                
            excel_tb = to_excel(tb_df)
            st.download_button(label="📥 تصدير ميزان المراجعة إلى Excel", data=excel_tb, file_name="trial_balance.csv", mime="text/csv")

    # --- 6. كشف حساب ---
    with tabs[5]:
        st.subheader("📑 كشف حساب تفصيلي")
        acc_df = pd.read_sql("SELECT id, code, name FROM accounts ORDER BY code ASC", engine)

        if not acc_df.empty:
            acc_dict_ledger = {f"{row['code']} - {row['name']}": int(row['id']) for _, row in acc_df.iterrows()}
            selected_account = st.selectbox("اختر الحساب:", list(acc_dict_ledger.keys()))
            
            col_l1, col_l2 = st.columns(2)
            with col_l1:
                start_date_l = st.date_input("من تاريخ", value=date(2020, 1, 1), key="l_start")
            with col_l2:
                end_date_l = st.date_input("إلى تاريخ", value=date.today(), key="l_end")

            if selected_account:
                acc_id = acc_dict_ledger[selected_account]
                ledger_df = pd.read_sql(
                    text("""
                        SELECT e.entry_date AS "التاريخ", e.id AS "رقم القيد", e.description AS "البيان",
                               d.debit AS "مدين", d.credit AS "دائن"
                        FROM journal_details d
                        JOIN journal_entries e ON d.entry_id = e.id
                        WHERE d.account_id = :aid AND e.entry_date BETWEEN :s AND :e
                        ORDER BY e.id ASC
                    """),
                    engine,
                    params={"aid": acc_id, "s": str(start_date_l), "e": str(end_date_l)}
                )

                if not ledger_df.empty:
                    ledger_df['الرصيد التراكمي'] = (ledger_df['مدين'] - ledger_df['دائن']).cumsum()
                    st.dataframe(ledger_df, use_container_width=True)
                    final_balance = ledger_df['الرصيد التراكمي'].iloc[-1]
                    st.info(f"💡 **الرصيد الحالي:** {final_balance:,.2f}")
                    excel_ledger = to_excel(ledger_df)
                    st.download_button(label="📥 تصدير كشف الحساب", data=excel_ledger, file_name=f"ledger_{selected_account}.csv", mime="text/csv")

    # --- 7. قائمة الدخل ---
    with tabs[6]:
        st.subheader("📈 قائمة الدخل (الإيرادات والمصروفات)")
        col_inc1, col_inc2 = st.columns(2)
        with col_inc1:
            start_date_inc = st.date_input("من تاريخ", value=date(2020, 1, 1), key="inc_start")
        with col_inc2:
            end_date_inc = st.date_input("إلى تاريخ", value=date.today(), key="inc_end")

        inc_df = pd.read_sql(
            text("""
                SELECT a.account_type AS "النوع", a.code AS "كود الحساب", a.name AS "اسم الحساب",
                       COALESCE(SUM(d.credit - d.debit), 0) AS "صافي الإيراد",
                       COALESCE(SUM(d.debit - d.credit), 0) AS "صافي المصروف"
                FROM accounts a
                JOIN journal_details d ON a.id = d.account_id
                JOIN journal_entries e ON d.entry_id = e.id
                WHERE a.account_type IN ('إيرادات', 'مصروفات') AND e.entry_date BETWEEN :s AND :e
                GROUP BY a.id, a.account_type, a.code, a.name
                ORDER BY a.code ASC
            """),
            engine,
            params={"s": str(start_date_inc), "e": str(end_date_inc)}
        )

        if not inc_df.empty:
            revenues_df = inc_df[inc_df['النوع'] == 'إيرادات'][['كود الحساب', 'اسم الحساب', 'صافي الإيراد']]
            expenses_df = inc_df[inc_df['النوع'] == 'مصروفات'][['كود الحساب', 'اسم الحساب', 'صافي المصروف']]

            st.markdown("#### 🟢 الإيرادات")
            if not revenues_df.empty:
                st.dataframe(revenues_df, use_container_width=True)
            st.markdown("#### 🔴 المصروفات")
            if not expenses_df.empty:
                st.dataframe(expenses_df, use_container_width=True)

            total_rev = revenues_df['صافي الإيراد'].sum() if not revenues_df.empty else 0.0
            total_exp = expenses_df['صافي المصروف'].sum() if not expenses_df.empty else 0.0
            net_profit = total_rev - total_exp

            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.metric("إجمالي الإيرادات", f"{total_rev:,.2f}")
            c2.metric("إجمالي المصروفات", f"{total_exp:,.2f}")
            c3.metric("صافي الدخل", f"{net_profit:,.2f}", delta=f"{net_profit:,.2f}")

    # --- 8. الميزانية العمومية ---
    with tabs[7]:
        st.subheader("🏛️ الميزانية العمومية (قائمة المركز المالي)")
        bs_date = st.date_input("الميزانية في تاريخ:", value=date.today(), key="bs_date")

        profit_query = pd.read_sql(
            text("""
                SELECT COALESCE(SUM(CASE WHEN a.account_type = 'إيرادات' THEN (d.credit - d.debit) ELSE 0 END), 0) -
                       COALESCE(SUM(CASE WHEN a.account_type = 'مصروفات' THEN (d.debit - d.credit) ELSE 0 END), 0) AS "net_profit"
                FROM journal_details d
                JOIN accounts a ON d.account_id = a.id
                JOIN journal_entries e ON d.entry_id = e.id
                WHERE e.entry_date <= :d
            """),
            engine,
            params={"d": str(bs_date)}
        )
        current_net_profit = profit_query.iloc[0]['net_profit'] if not profit_query.empty else 0.0

        bs_df = pd.read_sql(
            text("""
                SELECT a.account_type AS "النوع", a.code AS "كود الحساب", a.name AS "اسم الحساب",
                       COALESCE(SUM(d.debit - d.credit), 0) AS "الرصيد"
                FROM accounts a
                LEFT JOIN journal_details d ON a.id = d.account_id
                LEFT JOIN journal_entries e ON d.entry_id = e.id AND e.entry_date <= :d
                WHERE a.account_type IN ('أصول', 'خصوم', 'حقوق ملكية')
                GROUP BY a.id, a.account_type, a.code, a.name
                ORDER BY a.code ASC
            """),
            engine,
            params={"d": str(bs_date)}
        )

        col_assets, col_liab = st.columns(2)
        with col_assets:
            st.markdown("### 🔹 الأصول")
            assets_df = bs_df[bs_df['النوع'] == 'أصول'][['كود الحساب', 'اسم الحساب', 'الرصيد']]
            if not assets_df.empty:
                st.dataframe(assets_df, use_container_width=True)
            total_assets = assets_df['الرصيد'].sum() if not assets_df.empty else 0.0
            st.metric("إجمالي الأصول", f"{total_assets:,.2f}")

        with col_liab:
            st.markdown("### 🔸 الخصوم وحقوق الملكية")
            liab_df = bs_df[bs_df['النوع'] == 'خصوم'].copy()
            liab_df['الرصيد'] = -liab_df['الرصيد']
            eq_df = bs_df[bs_df['النوع'] == 'حقوق ملكية'].copy()
            eq_df['الرصيد'] = -eq_df['الرصيد']

            if not liab_df.empty:
                st.dataframe(liab_df[['كود الحساب', 'اسم الحساب', 'الرصيد']], use_container_width=True)
            if not eq_df.empty:
                st.dataframe(eq_df[['كود الحساب', 'اسم الحساب', 'الرصيد']], use_container_width=True)

            total_liab = liab_df['الرصيد'].sum() if not liab_df.empty else 0.0
            total_eq = eq_df['الرصيد'].sum() if not eq_df.empty else 0.0
            total_liab_eq = total_liab + total_eq + current_net_profit
            st.metric("إجمالي الخصوم وحقوق الملكية", f"{total_liab_eq:,.2f}")

        st.markdown("---")
        if abs(total_assets - total_liab_eq) < 0.01:
            st.success(f"⚖️ الميزانية العمومية متوازنة تماماً! ({total_assets:,.2f})")
        else:
            diff = total_assets - total_liab_eq
            st.error(f"❌ الفرق: {diff:,.2f}")
# --- 9. المشاريع ومراكز التكلفة ---
    with tabs[9]:
        st.subheader("🏗️ إدارة المشاريع وتقارير مراكز التكلفة")
        
        col_pr1, col_pr2 = st.columns([1, 2])
        with col_pr1:
            st.markdown("#### ➕ إضافة مشروع عقاري جديد")
            with st.form("add_project_form", clear_on_submit=True):
                new_proj_name = st.text_input("اسم المشروع (مثال: كمبوند الياسمين - برج النيل)")
                if st.form_submit_button("حفظ المشروع", type="primary"):
                    if new_proj_name.strip():
                        try:
                            with engine.begin() as conn:
                                conn.execute(
                                    text("INSERT INTO projects (name) VALUES (:n)"),
                                    {"n": new_proj_name.strip()}
                                )
                            st.success(f"✅ تم إضافة '{new_proj_name}' بنجاح!")
                            st.rerun()
                        except Exception:
                            st.error("هذا المشروع مسجل بالفعل!")
                    else:
                        st.error("يرجى كتابة اسم المشروع.")

        with col_pr2:
            st.markdown("#### 📊 ملخص موقف المشاريع (التكاليف والإيرادات والأرباح)")
            summary_proj_df = pd.read_sql(text("""
                SELECT 
                    COALESCE(e.project_name, 'عام / الإدارة') AS "اسم المشروع",
                    COALESCE(SUM(CASE WHEN a.account_type = 'إيرادات' THEN (d.credit - d.debit) ELSE 0 END), 0) AS "إجمالي الإيرادات",
                    COALESCE(SUM(CASE WHEN a.account_type = 'مصروفات' THEN (d.debit - d.credit) ELSE 0 END), 0) AS "إجمالي المصروفات",
                    COALESCE(SUM(CASE WHEN a.code = '115' THEN (d.debit - d.credit) ELSE 0 END), 0) AS "تكلفة تحت التنفيذ (WIP)",
                    (COALESCE(SUM(CASE WHEN a.account_type = 'إيرادات' THEN (d.credit - d.debit) ELSE 0 END), 0) -
                     COALESCE(SUM(CASE WHEN a.account_type = 'مصروفات' THEN (d.debit - d.credit) ELSE 0 END), 0)) AS "صافي ربح المشروع"
                FROM journal_entries e
                JOIN journal_details d ON e.id = d.entry_id
                JOIN accounts a ON d.account_id = a.id
                GROUP BY COALESCE(e.project_name, 'عام / الإدارة')
                ORDER BY "صافي ربح المشروع" DESC
            """), engine)
            
            if not summary_proj_df.empty:
                st.dataframe(summary_proj_df, use_container_width=True)
            else:
                st.info("لا توجد حركات مسجلة على المشاريع بعد.")

        st.markdown("---")
        st.markdown("#### 🔍 كشف حركات تفصيلي لمشروع محدد")
        all_projs = pd.read_sql("SELECT name FROM projects ORDER BY id ASC", engine)
        if not all_projs.empty:
            chosen_proj = st.selectbox("اختر المشروع لعرض كافة قيوده ومصروفاته:", all_projs['name'].tolist())
            proj_details_df = pd.read_sql(
                text("""
                    SELECT e.entry_date AS "التاريخ", e.id AS "رقم القيد", a.code AS "كود الحساب",
                           a.name AS "اسم الحساب", e.description AS "البيان", d.debit AS "مدين", d.credit AS "دائن"
                    FROM journal_entries e
                    JOIN journal_details d ON e.id = d.entry_id
                    JOIN accounts a ON d.account_id = a.id
                    WHERE COALESCE(e.project_name, 'عام / الإدارة') = :p
                    ORDER BY e.id DESC
                """),
                engine,
                params={"p": chosen_proj}
            )
            if not proj_details_df.empty:
                st.dataframe(proj_details_df, use_container_width=True)
                st.download_button(
                    label=f"📥 تصدير تقرير ({chosen_proj}) إلى Excel",
                    data=to_excel(proj_details_df),
                    file_name=f"project_{chosen_proj}.csv",
                    mime="text/csv"
                )
    # --- 9. إدارة المستخدمين ---
    if st.session_state['user_role'] == "مدير النظام":
        with tabs[8]:
            col_u_add, col_u_edit = st.columns(2)
            with col_u_add:
                st.subheader("➕ إضافة مستخدم جديد")
                with st.form("new_user_form", clear_on_submit=True):
                    new_username = st.text_input("اسم المستخدم")
                    new_password = st.text_input("كلمة السر", type="password")
                    new_role = st.selectbox("الصلاحية", ["محاسب / مدخل بيانات", "مدير النظام"])
                    if st.form_submit_button("إضافة المستخدم"):
                        if new_username and new_password:
                            try:
                                with engine.begin() as conn:
                                    conn.execute(
                                        text("INSERT INTO users (username, password, role) VALUES (:u, :p, :r)"),
                                        {"u": new_username, "p": make_hashes(new_password), "r": new_role}
                                    )
                                st.success("✅ تم الإضافة بنجاح!")
                                st.rerun()
                            except Exception:
                                st.error("المستخدم موجود مسبقاً!")

            with col_u_edit:
                st.subheader("⚙️ تعديل بيانات مستخدم")
                all_users_df = pd.read_sql("SELECT id, username, role FROM users", engine)

                if not all_users_df.empty:
                    user_dict = {f"{row['username']} ({row['role']})": row['username'] for _, row in all_users_df.iterrows()}
                    selected_u_label = st.selectbox("اختر المستخدم:", list(user_dict.keys()))
                    target_username = user_dict[selected_u_label]

                    with st.form("edit_user_form"):
                        updated_username = st.text_input("الاسم المعدل", value=target_username)
                        updated_password = st.text_input("كلمة السر الجديدة", type="password")
                        updated_role = st.selectbox("الصلاحية الجديدة", ["محاسب / مدخل بيانات", "مدير النظام"])

                        if st.form_submit_button("حفظ التغييرات"):
                            with engine.begin() as conn:
                                if updated_password:
                                    conn.execute(
                                        text("UPDATE users SET username = :nu, password = :np, role = :nr WHERE username = :ou"),
                                        {"nu": updated_username, "np": make_hashes(updated_password), "nr": updated_role, "ou": target_username}
                                    )
                                else:
                                    conn.execute(
                                        text("UPDATE users SET username = :nu, role = :nr WHERE username = :ou"),
                                        {"nu": updated_username, "nr": updated_role, "ou": target_username}
                                    )
                            st.success("✅ تم التحديث بنجاح!")
                            st.rerun()

            st.markdown("---")
            users_df = pd.read_sql('SELECT id AS "رقم", username AS "اسم المستخدم", role AS "الصلاحية" FROM users', engine)
            st.dataframe(users_df, use_container_width=True)

        # --- 10. الـ AI والتحديث التلقائي ---
        with tabs[10]:
            st.subheader("🤖 تعديل وتحديث الكود البرمجي تلقائياً بواسطة AI")
            st.info("ملاحظة: عند التشغيل السحابي، يولد الذكاء الاصطناعي الكود المحدث لتقوم بنسخه وتحديثه في ملفك على GitHub.")

            api_key = st.text_input("مفتاح OpenAI API Key الخاص بك:", type="password")
            ai_instruction = st.text_area("ما التعديل أو الميزة التي ترغب في إضافتها إلى الكود؟")

            if st.button("🚀 توليد الكود المعدل", type="primary"):
                if not api_key or not ai_instruction:
                    st.error("يرجى إدخال API Key والتعديل المطلوب!")
                else:
                    try:
                        from openai import OpenAI
                        client = OpenAI(api_key=api_key)
                        with st.spinner("جار كتابة الكود المحدث..."):
                            with open(__file__, "r", encoding="utf-8") as f:
                                current_code = f.read()

                            prompt = f"""
                            أنت خبير برمجة Streamlit و PostgreSQL المحاسبي.
                            إليك كود تطبيق محاسبي كاملاً:
                            ```python
                            {current_code}
                            ```
                            المطلوب إضافة التعديل التالي: "{ai_instruction}"
                            أعد الكود كاملاً داخل مجسم python فقط بدون أي شرح.
                            """
                            response = client.chat.completions.create(
                                model="gpt-4o",
                                messages=[{"role": "user", "content": prompt}],
                                temperature=0.2
                            )
                            new_code = response.choices[0].message.content
                            if "```python" in new_code:
                                new_code = new_code.split("```python")[1].split("```")[0]
                            elif "```" in new_code:
                                new_code = new_code.split("```")[1].split("```")[0]

                        st.success("✅ تم توليد الكود بنجاح! انسخه وضعه في ملف app.py على GitHub:")
                        st.code(new_code.strip(), language="python")
                    except Exception as e:
                        st.error(f"❌ حدث خطأ: {str(e)}")
