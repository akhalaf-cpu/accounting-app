from datetime import date, datetime, time as dt_time
from email.message import EmailMessage
import hashlib
import io
import json
import smtplib
import threading
import pandas as pd
from sqlalchemy import create_engine, text
import streamlit as st

# يقرأ رابط قاعدة البيانات السحابية من إعدادات Streamlit Secrets بأمان
DB_URL = st.secrets["DATABASE_URL"]
engine = create_engine(
    DB_URL, connect_args={"prepare_threshold": None}, pool_pre_ping=True
)

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
    ("524", "رسوم تراخيص واستشارات هندسية", "مصروفات", 32, 1),
]


def make_hashes(password):
    return hashlib.sha256(str.encode(password)).hexdigest()


def check_hashes(password, hashed_text):
    return make_hashes(password) == hashed_text


def to_excel(df):
    """تصدير ملف Excel حقيقي (.xlsx) مع fallback إلى CSV عند الحاجة"""
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Sheet1")
        return output.getvalue()
    except Exception:
        return df.to_csv(index=False).encode("utf-8-sig")


def build_backup_payload():
    """تجميع نسخة احتياطية كاملة من جميع جداول قاعدة البيانات بصيغتي JSON و Excel"""
    tables = [
        "users",
        "accounts",
        "projects",
        "units",
        "journal_entries",
        "journal_details",
        "installments",
    ]
    data_dict = {}
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        for tbl in tables:
            df_tbl = pd.read_sql(f"SELECT * FROM {tbl}", engine)
            data_dict[tbl] = df_tbl.to_dict(orient="records")
            sheet_name = tbl[:31]
            df_tbl.to_excel(writer, index=False, sheet_name=sheet_name)

    backup_meta = {
        "exported_at": datetime.now().isoformat(),
        "tables": data_dict,
    }
    json_bytes = json.dumps(
        backup_meta, ensure_ascii=False, indent=2, default=str
    ).encode("utf-8")
    return json_bytes, excel_buf.getvalue()


def send_backup_email(
    smtp_server: str,
    smtp_port: int,
    sender_email: str,
    sender_password: str,
    recipient_email: str,
    subject_prefix: str = "نسخة احتياطية يومية",
):
  """إرسال ملف النسخة الاحتياطية (JSON + Excel) عبر البريد الإلكتروني"""
  json_bytes, excel_bytes = build_backup_payload()
  now_str = datetime.now().strftime("%Y-%m-%d %H:%M")

  msg = EmailMessage()
  msg["Subject"] = f"{subject_prefix} - النظام المحاسبي ({now_str})"
  msg["From"] = sender_email
  msg["To"] = recipient_email
  msg.set_content(
      f"مرحباً،\n\nمرفق النسخة الاحتياطية لقاعدة بيانات النظام المحاسبي"
      f" بتاريخ: {now_str}\n\n- ملف JSON: نسخة كاملة مهيكلة لجميع"
      " الجداول.\n- ملف XLSX: جداول البيانات قابلة للفتح في Excel.\n"
  )

  msg.add_attachment(
      json_bytes,
      maintype="application",
      subtype="json",
      filename=f"accounting_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
  )
  msg.add_attachment(
      excel_bytes,
      maintype="application",
      subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      filename=f"accounting_backup_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
  )

  clean_password = sender_password.replace(" ", "").strip()
  port_num = int(smtp_port)

  if port_num == 465:
    with smtplib.SMTP_SSL(smtp_server.strip(), port_num, timeout=30) as server:
      server.login(sender_email.strip(), clean_password)
      server.send_message(msg)
  else:
    with smtplib.SMTP(smtp_server.strip(), port_num, timeout=30) as server:
      server.starttls()
      server.login(sender_email.strip(), clean_password)
      server.send_message(msg)

def get_smtp_config():
    """قراءة إعدادات البريد من st.secrets إن وجدت"""
    sec = st.secrets.get("backup_email", {})
    return {
        "smtp_server": sec.get("smtp_server", "smtp.gmail.com"),
        "smtp_port": int(sec.get("smtp_port", 587)),
        "sender_email": sec.get("sender_email", ""),
        "sender_password": sec.get("sender_password", ""),
        "recipient_email": sec.get("recipient_email", ""),
    }


def ensure_backup_table():
    with engine.begin() as conn:
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS backup_settings (
                id INTEGER PRIMARY KEY,
                auto_enabled INTEGER DEFAULT 0,
                send_hour INTEGER DEFAULT 23,
                send_minute INTEGER DEFAULT 55,
                smtp_server TEXT DEFAULT 'smtp.gmail.com',
                smtp_port INTEGER DEFAULT 587,
                sender_email TEXT,
                sender_password TEXT,
                recipient_email TEXT,
                last_sent_date TEXT
            )
        """)
        )
        conn.execute(
            text(
                "INSERT INTO backup_settings (id, auto_enabled, send_hour, send_minute) VALUES (1, 0, 23, 55) ON CONFLICT (id) DO NOTHING"
            )
        )


def _daily_backup_worker():
    """مراقب في الخلفية يرسل النسخة الاحتياطية تلقائياً في الموعد اليومي المحدد"""
    import time as py_time

    while True:
        try:
            with engine.connect() as conn:
                row = conn.execute(
                    text("""
                        SELECT auto_enabled, send_hour, send_minute, smtp_server, smtp_port,
                               sender_email, sender_password, recipient_email, last_sent_date
                        FROM backup_settings WHERE id = 1
                    """)
                ).fetchone()

            if row and row[0] == 1:
                now = datetime.now()
                today_str = now.strftime("%Y-%m-%d")
                target_reached = (now.hour > int(row[1])) or (
                    now.hour == int(row[1]) and now.minute >= int(row[2])
                )
                if target_reached and row[8] != today_str:
                    s_email = row[5] or ""
                    s_pass = row[6] or ""
                    r_email = row[7] or ""
                    if s_email and s_pass and r_email:
                        send_backup_email(
                            smtp_server=row[3] or "smtp.gmail.com",
                            smtp_port=int(row[4] or 587),
                            sender_email=s_email,
                            sender_password=s_pass,
                            recipient_email=r_email,
                            subject_prefix="نسخة احتياطية تلقائية نهاية اليوم",
                        )
                        with engine.begin() as conn:
                            conn.execute(
                                text(
                                    "UPDATE backup_settings SET last_sent_date = :d WHERE id = 1"
                                ),
                                {"d": today_str},
                            )
        except Exception:
            pass
        py_time.sleep(60)


@st.cache_resource
def start_backup_scheduler():
    t = threading.Thread(target=_daily_backup_worker, daemon=True)
    t.start()
    return True


def import_accounts_from_df(df_upload: pd.DataFrame):
    """استيراد حسابات إلى جدول accounts من ملف Excel أو CSV"""
    # توحيد أسماء الأعمدة الشائعة بالعربية والإنجليزية
    col_map = {
        "كود الحساب": "code",
        "الكود": "code",
        "code": "code",
        "اسم الحساب": "name",
        "الاسم": "name",
        "name": "name",
        "نوع الحساب": "account_type",
        "النوع": "account_type",
        "account_type": "account_type",
        "كود الحساب الأب": "parent_code",
        "parent_code": "parent_code",
        "parent_id": "parent_id",
        "allow_posting": "allow_posting",
        "طبيعة التسجيل": "allow_posting",
    }
    renamed = df_upload.rename(
        columns={c: col_map.get(str(c).strip(), str(c).strip()) for c in df_upload.columns}
    )

    if not {"code", "name", "account_type"}.issubset(set(renamed.columns)):
        return (
            0,
            0,
            "الملف يجب أن يحتوي على الأعمدة التالية على الأقل: code (أو كود الحساب), name (أو اسم الحساب), account_type (أو نوع الحساب).",
        )

    added = 0
    skipped = 0

    with engine.begin() as conn:
        existing_rows = conn.execute(
            text("SELECT id, code FROM accounts WHERE code IS NOT NULL")
        ).fetchall()
        code_to_id = {str(r[1]).strip(): int(r[0]) for r in existing_rows}

        for _, r in renamed.iterrows():
            raw_code = r.get("code")
            raw_name = r.get("name")
            raw_type = r.get("account_type")
            if pd.isna(raw_code) or pd.isna(raw_name) or pd.isna(raw_type):
                skipped += 1
                continue

            acc_code = str(raw_code).strip()
            acc_name = str(raw_name).strip()
            acc_type = str(raw_type).strip()

            if not acc_code or not acc_name or acc_type not in TYPES:
                skipped += 1
                continue

            if acc_code in code_to_id:
                skipped += 1
                continue

            parent_id = None
            if "parent_code" in renamed.columns and pd.notna(r.get("parent_code")):
                p_code = str(r.get("parent_code")).strip()
                parent_id = code_to_id.get(p_code)
            elif "parent_id" in renamed.columns and pd.notna(r.get("parent_id")):
                try:
                    parent_id = int(r.get("parent_id"))
                except Exception:
                    parent_id = None

            allow_val = 1
            if "allow_posting" in renamed.columns and pd.notna(r.get("allow_posting")):
                val_str = str(r.get("allow_posting")).strip()
                if val_str in ["0", "ممنوع", "ممنوع التسجيل (رئيسي)", "False", "false"]:
                    allow_val = 0

            new_id = conn.execute(
                text("""
                    INSERT INTO accounts (code, name, account_type, parent_id, allow_posting)
                    VALUES (:c, :n, :t, :p, :a)
                    RETURNING id
                """),
                {
                    "c": acc_code,
                    "n": acc_name,
                    "t": acc_type,
                    "p": parent_id,
                    "a": allow_val,
                },
            ).scalar()
            code_to_id[acc_code] = int(new_id)
            added += 1

    return added, skipped, None


def generate_entry_html(
    entry_id, entry_date, description, debit_acc, credit_acc, amount
):
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
    return html_content.encode("utf-8")


def init_db():
    with engine.begin() as conn:
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS users (
                id SERIAL PRIMARY KEY,
                username TEXT UNIQUE NOT NULL,
                password TEXT NOT NULL,
                role TEXT NOT NULL
            )
        """)
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS accounts (
                id SERIAL PRIMARY KEY,
                code TEXT UNIQUE,
                name TEXT NOT NULL,
                account_type TEXT NOT NULL,
                parent_id INTEGER,
                allow_posting INTEGER DEFAULT 1,
                balance REAL DEFAULT 0.0
            )
        """)
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS projects (
                id SERIAL PRIMARY KEY,
                name TEXT UNIQUE NOT NULL,
                project_code TEXT,
                tower_no TEXT,
                project_type TEXT DEFAULT 'سكني'
            )
        """)
        )
        conn.execute(
            text(
                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS project_code TEXT"
            )
        )
        conn.execute(
            text("ALTER TABLE projects ADD COLUMN IF NOT EXISTS tower_no TEXT")
        )
        conn.execute(
            text(
                "ALTER TABLE projects ADD COLUMN IF NOT EXISTS project_type TEXT DEFAULT 'سكني'"
            )
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS units (
                id SERIAL PRIMARY KEY,
                project_name TEXT,
                unit_code TEXT UNIQUE,
                area REAL DEFAULT 0.0,
                price_per_meter REAL DEFAULT 0.0,
                total_price REAL DEFAULT 0.0,
                status TEXT DEFAULT 'متاحة للبيع',
                customer_name TEXT
            )
        """)
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS journal_entries (
                id SERIAL PRIMARY KEY,
                entry_date TEXT NOT NULL,
                description TEXT,
                project_name TEXT DEFAULT 'عام / الإدارة'
            )
        """)
        )
        conn.execute(
            text(
                "ALTER TABLE journal_entries ADD COLUMN IF NOT EXISTS project_name TEXT DEFAULT 'عام / الإدارة'"
            )
        )
        conn.execute(
            text(
                "INSERT INTO projects (name) VALUES ('عام / الإدارة') ON CONFLICT DO NOTHING"
            )
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS journal_details (
                id SERIAL PRIMARY KEY,
                entry_id INTEGER NOT NULL REFERENCES journal_entries(id) ON DELETE CASCADE,
                account_id INTEGER NOT NULL REFERENCES accounts(id),
                debit REAL DEFAULT 0.0,
                credit REAL DEFAULT 0.0
            )
        """)
        )
        conn.execute(
            text("""
            CREATE TABLE IF NOT EXISTS installments (
                id SERIAL PRIMARY KEY,
                project_name TEXT,
                customer_name TEXT NOT NULL,
                unit_code TEXT NOT NULL,
                contract_no TEXT,
                installment_type TEXT,
                due_date TEXT NOT NULL,
                amount REAL DEFAULT 0.0,
                paid_amount REAL DEFAULT 0.0,
                collector_name TEXT,
                status TEXT DEFAULT 'Unpaid'
            )
        """)
        )
        res = conn.execute(text("SELECT COUNT(*) FROM users")).scalar()
        if res == 0:
            conn.execute(
                text(
                    "INSERT INTO users (username, password, role) VALUES (:u, :p, :r)"
                ),
                {
                    "u": "admin",
                    "p": make_hashes("admin123"),
                    "r": "مدير النظام",
                },
            )
    ensure_backup_table()


def load_default_accounts():
    with engine.begin() as conn:
        for code, name, acc_type, parent_id, allow_post in DEFAULT_ACCOUNTS:
            conn.execute(
                text("""
                    INSERT INTO accounts (code, name, account_type, parent_id, allow_posting)
                    VALUES (:c, :n, :t, :p, :a)
                    ON CONFLICT (code) DO NOTHING
                """),
                {
                    "c": code,
                    "n": name,
                    "t": acc_type,
                    "p": parent_id,
                    "a": allow_post,
                },
            )


st.set_page_config(page_title="برنامجي المحاسبي السحابي", layout="wide")

if "db_initialized" not in st.session_state:
    init_db()
    start_backup_scheduler()
    st.session_state["db_initialized"] = True

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
    st.session_state["user_role"] = ""
    st.session_state["username"] = ""

if not st.session_state["logged_in"]:
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
                        {"u": username},
                    ).fetchone()

                if user_res and check_hashes(password, user_res[0]):
                    st.session_state["logged_in"] = True
                    st.session_state["username"] = username
                    st.session_state["user_role"] = user_res[1]
                    st.success(f"مرحباً بك {username}!")
                    st.rerun()
                else:
                    st.error("❌ اسم المستخدم أو كلمة السر غير صحيحة!")

else:
    st.sidebar.markdown(f"👤 **المستخدم الحالي:** {st.session_state['username']}")
    st.sidebar.markdown(f"🔰 **الصلاحية:** {st.session_state['user_role']}")
    st.sidebar.success("☁️ متصل بقاعدة البيانات السحابية")

    if st.sidebar.button("🚪 تسجيل الخروج", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["user_role"] = ""
        st.session_state["username"] = ""
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
        "🏛️ الميزانية العمومية",
        "🏗️ المشاريع ومراكز التكلفة",
        "📅 أقساط العملاء",
    ]

    is_admin = st.session_state["user_role"] == "مدير النظام"
    if is_admin:
        tabs_list.append("👥 إدارة المستخدمين")
        tabs_list.append("🤖 الـ AI وتحديث الكود")
        tabs_list.append("💾 النسخ الاحتياطي")

    tabs = st.tabs(tabs_list)

    # --- 1. شجرة الحسابات ---
    with tabs[0]:
        st.subheader("🌳 إدارة وتأسيس شجرة الحسابات")
        col_top_b1, col_top_b2 = st.columns([1, 2])
        with col_top_b1:
            if st.button("⚡ انشاء شجرة حسابات قياسية كاملاً", type="primary"):
                load_default_accounts()
                st.success("✅ تم بناء شجرة الحسابات القياسية بنجاح!")
                st.rerun()

        with st.expander("📥 استيراد شجرة حسابات من ملف Excel / CSV"):
            st.caption(
                "الأعمدة المطلوبة: `code` (أو كود الحساب)، `name` (أو اسم الحساب)، `account_type` (أصول / خصوم / حقوق ملكية / إيرادات / مصروفات). اختياري: `parent_code`."
            )
            template_df = pd.DataFrame(
                [
                    {
                        "code": "118",
                        "name": "مقاولون تحت الحساب",
                        "account_type": "أصول",
                        "parent_code": "11",
                        "allow_posting": 1,
                    }
                ]
            )
            st.download_button(
                label="📄 تحميل نموذج Excel جاهز للتعبئة",
                data=to_excel(template_df),
                file_name="chart_of_accounts_template.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            uploaded_coa = st.file_uploader(
                "اختر ملف شجرة الحسابات (.xlsx, .xls, .csv)",
                type=["xlsx", "xls", "csv"],
                key="coa_upload",
            )
            if uploaded_coa is not None:
                try:
                    if uploaded_coa.name.lower().endswith(".csv"):
                        up_df = pd.read_csv(uploaded_coa)
                    else:
                        up_df = pd.read_excel(uploaded_coa)
                    st.dataframe(up_df.head(10), use_container_width=True)
                    if st.button("🚀 تنفيذ استيراد الحسابات للملف", type="primary"):
                        added_cnt, skipped_cnt, err_msg = import_accounts_from_df(up_df)
                        if err_msg:
                            st.error(f"❌ {err_msg}")
                        else:
                            st.success(
                                f"✅ تم استيراد {added_cnt} حساب جديد، وتم تخطي {skipped_cnt} صف (مكرر أو غير مكتمل)."
                            )
                            st.rerun()
                except Exception as e:
                    st.error(f"❌ تعذر قراءة الملف: {str(e)}")

        st.markdown("---")
        col_acc1, col_acc2 = st.columns(2)

        with col_acc1:
            st.subheader("➕ إضافة حساب جديد")
            accounts_df = pd.read_sql("SELECT id, code, name FROM accounts", engine)

            parent_options = {"بدون (حساب رئيسي)": None}
            for _, row in accounts_df.iterrows():
                parent_options[f"{row['code']} - {row['name']}"] = int(row["id"])

            with st.form("account_form", clear_on_submit=True):
                code = st.text_input("كود الحساب (مثال: 1111)")
                name = st.text_input("اسم الحساب (مثال: خزانة الفرع الرئيسي)")
                account_type = st.selectbox("نوع الحساب", TYPES)
                parent_label = st.selectbox(
                    "الحساب الأب", list(parent_options.keys())
                )
                allow_post = st.selectbox(
                    "طبيعة التسجيل:",
                    ["متاح للتسجيل (فرعي)", "ممنوع التسجيل (رئيسي)"],
                )

                submitted = st.form_submit_button("حفظ الحساب")
                if submitted:
                    if not code or not name:
                        st.error("يرجى إدخال الكود والاسم!")
                    else:
                        try:
                            parent_id = parent_options[parent_label]
                            allow_val = (
                                1 if allow_post == "متاح للتسجيل (فرعي)" else 0
                            )
                            with engine.begin() as conn:
                                conn.execute(
                                    text(
                                        "INSERT INTO accounts (code, name, account_type, parent_id, allow_posting) VALUES (:c, :n, :t, :p, :a)"
                                    ),
                                    {
                                        "c": code,
                                        "n": name,
                                        "t": account_type,
                                        "p": parent_id,
                                        "a": allow_val,
                                    },
                                )
                            st.success(f"تم إضافة الحساب '{name}' بنجاح!")
                            st.rerun()
                        except Exception:
                            st.error("كود الحساب هذا مُسجّل مسبقاً!")

        with col_acc2:
            st.subheader("⚙️ تعديل بيانات الحساب (الاسم / الكود / الحالة)")
            all_accs = pd.read_sql(
                "SELECT id, code, name, parent_id, allow_posting FROM accounts ORDER BY code ASC",
                engine,
            )

            if not all_accs.empty:
                acc_options = {
                    f"{row['code']} - {row['name']}": (
                        int(row["id"]),
                        row["code"],
                        row["name"],
                        int(row["allow_posting"]),
                    )
                    for _, row in all_accs.iterrows()
                }
                selected_edit_acc = st.selectbox(
                    "اختر الحساب المراد تعديله:",
                    list(acc_options.keys()),
                    key="edit_acc_sel",
                )

                target_id, target_code, target_name, curr_posting = acc_options[
                    selected_edit_acc
                ]
                curr_status_index = 0 if curr_posting == 1 else 1

                with st.form("edit_details_form"):
                    updated_code = st.text_input(
                        "كود الحساب المعدل", value=target_code
                    )
                    updated_name = st.text_input(
                        "اسم الحساب المعدل", value=target_name
                    )
                    new_status = st.radio(
                        "طبيعة التسجيل لهذا الحساب:",
                        ["متاح للتسجيل (فرعي)", "ممنوع التسجيل (رئيسي)"],
                        index=curr_status_index,
                    )

                    save_update = st.form_submit_button("💾 حفظ التعديلات")

                    if save_update:
                        if not updated_code or not updated_name:
                            st.error("يرجى إدخال كود واسم الحساب!")
                        else:
                            new_val = (
                                1 if new_status == "متاح للتسجيل (فرعي)" else 0
                            )
                            with engine.connect() as conn:
                                entries_count = conn.execute(
                                    text(
                                        "SELECT COUNT(*) FROM journal_details WHERE account_id = :id"
                                    ),
                                    {"id": target_id},
                                ).scalar()

                            if (
                                new_val == 0
                                and curr_posting == 1
                                and entries_count > 0
                            ):
                                st.error(
                                    "🛑 لا يمكن تغيير طبيعة الحساب لـ [ممنوع التسجيل] لأن عليه قيوداً مسجلة!"
                                )
                            else:
                                try:
                                    with engine.begin() as conn:
                                        conn.execute(
                                            text(
                                                "UPDATE accounts SET code = :c, name = :n, allow_posting = :a WHERE id = :id"
                                            ),
                                            {
                                                "c": updated_code,
                                                "n": updated_name,
                                                "a": new_val,
                                                "id": target_id,
                                            },
                                        )
                                    st.success(
                                        f"✅ تم تحديث بيانات الحساب إلى '{updated_name}' بنجاح!"
                                    )
                                    st.rerun()
                                except Exception:
                                    st.error(
                                        "كود الحساب الجديد مُسجل مسبقاً لحساب آخر!"
                                    )
            else:
                st.info("لا توجد حسابات مسجلة بعد.")

        st.markdown("---")
        st.subheader("📋 دليل شجرة الحسابات المعتمدة")
        df_acc = pd.read_sql(
            """
            SELECT a.code AS "كود الحساب", a.name AS "اسم الحساب", a.account_type AS "النوع", 
                   COALESCE(p.name, 'رئيسي') AS "الحساب الأب",
                   CASE WHEN a.allow_posting = 1 THEN '✅ متاح للتسجيل' ELSE '❌ ممنوع التسجيل (رئيسي)' END AS "حالة التسجيل"
            FROM accounts a
            LEFT JOIN accounts p ON a.parent_id = p.id
            ORDER BY a.code ASC
        """,
            engine,
        )

        if not df_acc.empty:
            st.dataframe(df_acc, use_container_width=True)
            st.download_button(
                label="📥 تصدير شجرة الحسابات إلى Excel",
                data=to_excel(df_acc),
                file_name="chart_of_accounts.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # --- 2. الأرصدة الافتتاحية ---
    with tabs[1]:
        st.subheader("🏁 إدخال الأرصدة الافتتاحية")
        acc_list = pd.read_sql(
            "SELECT id, code, name, account_type, allow_posting FROM accounts ORDER BY code ASC",
            engine,
        )

        if acc_list.empty:
            st.warning("⚠️ يرجى إنشاء شجرة الحسابات أولاً.")
        else:
            acc_dict = {}
            for _, row in acc_list.iterrows():
                status_str = (
                    "✅ متاح" if row["allow_posting"] == 1 else "❌ ممنوع التسجيل"
                )
                acc_dict[f"{row['code']} - {row['name']} ({status_str})"] = (
                    int(row["id"]),
                    row["name"],
                    int(row["allow_posting"]),
                )

            with st.form("opening_balance_form", clear_on_submit=True):
                col_ob1, col_ob2 = st.columns(2)
                with col_ob1:
                    selected_acc_label = st.selectbox(
                        "الحساب:", list(acc_dict.keys())
                    )
                    balance_type = st.radio(
                        "طبيعة الرصيد:",
                        ["مدين (نقدية/أصول)", "دائن (رأس مال/موردين)"],
                    )
                with col_ob2:
                    ob_date = st.date_input(
                        "تاريخ الرصيد", value=date(2026, 1, 1)
                    )
                    ob_amount = st.number_input(
                        "المبلغ الافتتاحي", min_value=0.0, step=500.0
                    )

                submit_ob = st.form_submit_button(
                    "💾 حفظ الرصيد الافتتاحي", type="primary"
                )

                if submit_ob:
                    target_acc_id, target_acc_name, is_allowed = acc_dict[
                        selected_acc_label
                    ]
                    if is_allowed == 0:
                        st.error(
                            "🛑 هذا الحساب مصنف [ممنوع التسجيل]! يرجى اختيار حساب فرعي متاح."
                        )
                    elif ob_amount <= 0:
                        st.error("يرجى إدخال مبلغ أكبر من الصفر!")
                    else:
                        desc = f"رصيد افتتاحي - {target_acc_name}"
                        with engine.begin() as conn:
                            entry_id = conn.execute(
                                text(
                                    "INSERT INTO journal_entries (entry_date, description) VALUES (:d, :desc) RETURNING id"
                                ),
                                {"d": str(ob_date), "desc": desc},
                            ).scalar()

                            if "مدين" in balance_type:
                                conn.execute(
                                    text(
                                        "INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, :amt, 0)"
                                    ),
                                    {
                                        "eid": entry_id,
                                        "aid": target_acc_id,
                                        "amt": ob_amount,
                                    },
                                )
                            else:
                                conn.execute(
                                    text(
                                        "INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, 0, :amt)"
                                    ),
                                    {
                                        "eid": entry_id,
                                        "aid": target_acc_id,
                                        "amt": ob_amount,
                                    },
                                )
                        st.success("✅ تم حفظ الرصيد الافتتاحي بنجاح!")
                        st.rerun()

            st.markdown("---")
            ob_df = pd.read_sql(
                text("""
                SELECT e.entry_date AS "التاريخ", a.code AS "كود الحساب", a.name AS "اسم الحساب",
                       d.debit AS "رصيد مدين", d.credit AS "رصيد دائن", e.description AS "البيان"
                FROM journal_entries e
                JOIN journal_details d ON e.id = d.entry_id
                JOIN accounts a ON d.account_id = a.id
                WHERE e.description LIKE 'رصيد افتتاحي%'
                ORDER BY e.id DESC
            """),
                engine,
            )
            if not ob_df.empty:
                st.dataframe(ob_df, use_container_width=True)

    # --- 3. إضافة قيد يومي ---
    with tabs[2]:
        st.subheader("📝 تسجيل قيد محاسبي جديد")
        acc_list = pd.read_sql(
            "SELECT id, code, name, allow_posting FROM accounts ORDER BY code ASC",
            engine,
        )

        if not acc_list.empty:
            acc_dict = {}
            for _, row in acc_list.iterrows():
                status_str = (
                    "✅ متاح" if row["allow_posting"] == 1 else "❌ ممنوع التسجيل"
                )
                acc_dict[f"{row['code']} - {row['name']} ({status_str})"] = (
                    int(row["id"]),
                    int(row["allow_posting"]),
                )

            projects_df = pd.read_sql(
                "SELECT name FROM projects ORDER BY id ASC", engine
            )
            projects_list = (
                projects_df["name"].tolist()
                if not projects_df.empty
                else ["عام / الإدارة"]
            )

            col_top1, col_top2 = st.columns(2)
            with col_top1:
                entry_date = st.date_input("تاريخ القيد", value=date.today())
            with col_top2:
                selected_project = st.selectbox(
                    "🏗️ المشروع / مركز التكلفة", projects_list
                )

            description = st.text_input("البيان / شرح القيد")
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                debit_acc = st.selectbox(
                    "حساب المدين", list(acc_dict.keys()), key="deb_acc"
                )
                debit_val = st.number_input(
                    "المبلغ المدين", min_value=0.0, step=100.0, key="deb_val"
                )
            with col_m2:
                credit_acc = st.selectbox(
                    "حساب الدائن", list(acc_dict.keys()), key="cred_acc"
                )
                credit_val = st.number_input(
                    "المبلغ الدائن", min_value=0.0, step=100.0, key="cred_val"
                )

            if st.button("حفظ القيد اليومي", type="primary"):
                deb_id, deb_allowed = acc_dict[debit_acc]
                cred_id, cred_allowed = acc_dict[credit_acc]

                if deb_allowed == 0:
                    st.error(
                        "🛑 حساب المدين المحدد [ممنوع التسجيل]! يرجى اختيار حساب فرعي."
                    )
                elif cred_allowed == 0:
                    st.error(
                        "🛑 حساب الدائن المحدد [ممنوع التسجيل]! يرجى اختيار حساب فرعي."
                    )
                elif debit_val <= 0 or credit_val <= 0:
                    st.error("يرجى إدخال مبالغ أكبر من الصفر.")
                elif debit_val != credit_val:
                    st.error("❌ القيد غير متوازن!")
                elif debit_acc == credit_acc:
                    st.error("❌ لا يمكن اختيار نفس الحساب!")
                else:
                    with engine.begin() as conn:
                        entry_id = conn.execute(
                            text(
                                "INSERT INTO journal_entries (entry_date, description, project_name) VALUES (:d, :desc, :proj) RETURNING id"
                            ),
                            {
                                "d": str(entry_date),
                                "desc": description,
                                "proj": selected_project,
                            },
                        ).scalar()
                        conn.execute(
                            text(
                                "INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, :amt, 0)"
                            ),
                            {"eid": entry_id, "aid": deb_id, "amt": debit_val},
                        )
                        conn.execute(
                            text(
                                "INSERT INTO journal_details (entry_id, account_id, debit, credit) VALUES (:eid, :aid, 0, :amt)"
                            ),
                            {"eid": entry_id, "aid": cred_id, "amt": credit_val},
                        )
                    st.success("✅ تم حفظ القيد اليومي بنجاح!")
                    st.rerun()

    # --- 4. دفتر اليومية والطباعة والحذف ---
    with tabs[3]:
        st.subheader("📖 دفتر اليومية وتصدير السندات")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            start_date_j = st.date_input(
                "من تاريخ", value=date(2020, 1, 1), key="j_start"
            )
        with col_f2:
            end_date_j = st.date_input(
                "إلى تاريخ", value=date.today(), key="j_end"
            )

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
            params={"s": str(start_date_j), "e": str(end_date_j)},
        )

        if not journal_df.empty:
            st.dataframe(journal_df, use_container_width=True)
            excel_journal = to_excel(journal_df)
            st.download_button(
                label="📥 تصدير دفتر اليومية إلى Excel",
                data=excel_journal,
                file_name="journal_entries.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            st.markdown("---")
            col_p1, col_p2 = st.columns(2)

            with col_p1:
                st.subheader("🖨️ طباعة سند قيد / فاتورة")
                entry_ids = journal_df["رقم القيد"].unique()
                selected_entry_id = st.selectbox(
                    "اختر رقم القيد لطباعة السند:", entry_ids
                )

                if selected_entry_id:
                    e_rows = journal_df[
                        journal_df["رقم القيد"] == selected_entry_id
                    ]
                    e_date = e_rows.iloc[0]["التاريخ"]
                    e_desc = e_rows.iloc[0]["البيان"]
                    deb_row = e_rows[e_rows["مدين"] > 0]
                    cred_row = e_rows[e_rows["دائن"] > 0]

                    if not deb_row.empty and not cred_row.empty:
                        deb_name = deb_row.iloc[0]["اسم الحساب"]
                        cred_name = cred_row.iloc[0]["اسم الحساب"]
                        amt = deb_row.iloc[0]["مدين"]
                        html_data = generate_entry_html(
                            selected_entry_id,
                            e_date,
                            e_desc,
                            deb_name,
                            cred_name,
                            amt,
                        )
                        st.download_button(
                            label=f"📄 تحميل سند قيد رقم #{selected_entry_id}",
                            data=html_data,
                            file_name=f"Voucher_{selected_entry_id}.html",
                            mime="text/html",
                            type="primary",
                        )

            with col_p2:
                st.subheader("🗑️ إلغاء / حذف قيد يومية")
                if is_admin:
                    all_entry_ids = journal_df["رقم القيد"].unique()
                    entry_to_delete = st.selectbox(
                        "اختر رقم القيد المراد حذفه نهائياً:",
                        all_entry_ids,
                        key="del_entry_sel",
                    )

                    if st.button("❌ حذف القيد المحدد الآن"):
                        with engine.begin() as conn:
                            conn.execute(
                                text(
                                    "DELETE FROM journal_details WHERE entry_id = :id"
                                ),
                                {"id": int(entry_to_delete)},
                            )
                            conn.execute(
                                text(
                                    "DELETE FROM journal_entries WHERE id = :id"
                                ),
                                {"id": int(entry_to_delete)},
                            )
                        st.success(
                            f"✅ تم حذف القيد رقم #{entry_to_delete} بنجاح!"
                        )
                        st.rerun()
                else:
                    st.warning("🔒 إلغاء القيود متاح فقط لـ 'مدير النظام'.")

    # --- 5. ميزان المراجعة ---
    with tabs[4]:
        st.subheader("⚖️ ميزان المراجعة")
        col_tb1, col_tb2 = st.columns(2)
        with col_tb1:
            start_date_tb = st.date_input(
                "من تاريخ", value=date(2020, 1, 1), key="tb_start"
            )
        with col_tb2:
            end_date_tb = st.date_input(
                "إلى تاريخ", value=date.today(), key="tb_end"
            )

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
            params={"s": str(start_date_tb), "e": str(end_date_tb)},
        )

        if not tb_df.empty:
            st.dataframe(tb_df, use_container_width=True)
            tot_debit = tb_df["مجموع المدين"].sum()
            tot_credit = tb_df["مجموع الدائن"].sum()

            col_t1, col_t2 = st.columns(2)
            col_t1.metric("إجمالي مجموع المدين", f"{tot_debit:,.2f}")
            col_t2.metric("إجمالي مجموع الدائن", f"{tot_credit:,.2f}")

            if abs(tot_debit - tot_credit) < 0.01:
                st.success("✅ ميزان المراجعة متوازن!")
            else:
                st.error("❌ يوجد عدم توازن في الميزان!")

            excel_tb = to_excel(tb_df)
            st.download_button(
                label="📥 تصدير ميزان المراجعة إلى Excel",
                data=excel_tb,
                file_name="trial_balance.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # --- 6. كشف حساب ---
    with tabs[5]:
        st.subheader("📑 كشف حساب تفصيلي")
        acc_df = pd.read_sql(
            "SELECT id, code, name FROM accounts ORDER BY code ASC", engine
        )

        if not acc_df.empty:
            acc_dict_ledger = {
                f"{row['code']} - {row['name']}": int(row["id"])
                for _, row in acc_df.iterrows()
            }
            selected_account = st.selectbox(
                "اختر الحساب:", list(acc_dict_ledger.keys())
            )

            col_l1, col_l2 = st.columns(2)
            with col_l1:
                start_date_l = st.date_input(
                    "من تاريخ", value=date(2020, 1, 1), key="l_start"
                )
            with col_l2:
                end_date_l = st.date_input(
                    "إلى تاريخ", value=date.today(), key="l_end"
                )

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
                    params={
                        "aid": acc_id,
                        "s": str(start_date_l),
                        "e": str(end_date_l),
                    },
                )

                if not ledger_df.empty:
                    ledger_df["الرصيد التراكمي"] = (
                        ledger_df["مدين"] - ledger_df["دائن"]
                    ).cumsum()
                    st.dataframe(ledger_df, use_container_width=True)
                    final_balance = ledger_df["الرصيد التراكمي"].iloc[-1]
                    st.info(f"💡 **الرصيد الحالي:** {final_balance:,.2f}")
                    excel_ledger = to_excel(ledger_df)
                    st.download_button(
                        label="📥 تصدير كشف الحساب إلى Excel",
                        data=excel_ledger,
                        file_name=f"ledger_{selected_account}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

    # --- 7. قائمة الدخل ---
    with tabs[6]:
        st.subheader("📈 قائمة الدخل (الإيرادات والمصروفات)")
        col_inc1, col_inc2 = st.columns(2)
        with col_inc1:
            start_date_inc = st.date_input(
                "من تاريخ", value=date(2020, 1, 1), key="inc_start"
            )
        with col_inc2:
            end_date_inc = st.date_input(
                "إلى تاريخ", value=date.today(), key="inc_end"
            )

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
            params={"s": str(start_date_inc), "e": str(end_date_inc)},
        )

        if not inc_df.empty:
            revenues_df = inc_df[inc_df["النوع"] == "إيرادات"][
                ["كود الحساب", "اسم الحساب", "صافي الإيراد"]
            ]
            expenses_df = inc_df[inc_df["النوع"] == "مصروفات"][
                ["كود الحساب", "اسم الحساب", "صافي المصروف"]
            ]

            st.markdown("#### 🟢 الإيرادات")
            if not revenues_df.empty:
                st.dataframe(revenues_df, use_container_width=True)
            st.markdown("#### 🔴 المصروفات")
            if not expenses_df.empty:
                st.dataframe(expenses_df, use_container_width=True)

            total_rev = (
                revenues_df["صافي الإيراد"].sum()
                if not revenues_df.empty
                else 0.0
            )
            total_exp = (
                expenses_df["صافي المصروف"].sum()
                if not expenses_df.empty
                else 0.0
            )
            net_profit = total_rev - total_exp

            st.markdown("---")
            c1, c2, c3 = st.columns(3)
            c1.metric("إجمالي الإيرادات", f"{total_rev:,.2f}")
            c2.metric("إجمالي المصروفات", f"{total_exp:,.2f}")
            c3.metric(
                "صافي الدخل", f"{net_profit:,.2f}", delta=f"{net_profit:,.2f}"
            )

    # --- 8. الميزانية العمومية ---
    with tabs[7]:
        st.subheader("🏛️ الميزانية العمومية (قائمة المركز المالي)")
        bs_date = st.date_input(
            "الميزانية في تاريخ:", value=date.today(), key="bs_date"
        )

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
            params={"d": str(bs_date)},
        )
        current_net_profit = (
            profit_query.iloc[0]["net_profit"]
            if not profit_query.empty
            else 0.0
        )

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
            params={"d": str(bs_date)},
        )

        col_assets, col_liab = st.columns(2)
        with col_assets:
            st.markdown("### 🔹 الأصول")
            assets_df = bs_df[bs_df["النوع"] == "أصول"][
                ["كود الحساب", "اسم الحساب", "الرصيد"]
            ]
            if not assets_df.empty:
                st.dataframe(assets_df, use_container_width=True)
            total_assets = (
                assets_df["الرصيد"].sum() if not assets_df.empty else 0.0
            )
            st.metric("إجمالي الأصول", f"{total_assets:,.2f}")

        with col_liab:
            st.markdown("### 🔸 الخصوم وحقوق الملكية")
            liab_df = bs_df[bs_df["النوع"] == "خصوم"].copy()
            liab_df["الرصيد"] = -liab_df["الرصيد"]
            eq_df = bs_df[bs_df["النوع"] == "حقوق ملكية"].copy()
            eq_df["الرصيد"] = -eq_df["الرصيد"]

            if not liab_df.empty:
                st.dataframe(
                    liab_df[["كود الحساب", "اسم الحساب", "الرصيد"]],
                    use_container_width=True,
                )
            if not eq_df.empty:
                st.dataframe(
                    eq_df[["كود الحساب", "اسم الحساب", "الرصيد"]],
                    use_container_width=True,
                )

            total_liab = liab_df["الرصيد"].sum() if not liab_df.empty else 0.0
            total_eq = eq_df["الرصيد"].sum() if not eq_df.empty else 0.0
            total_liab_eq = total_liab + total_eq + current_net_profit
            st.metric("إجمالي الخصوم وحقوق الملكية", f"{total_liab_eq:,.2f}")

        st.markdown("---")
        if abs(total_assets - total_liab_eq) < 0.01:
            st.success(
                f"⚖️ الميزانية العمومية متوازنة تماماً! ({total_assets:,.2f})"
            )
        else:
            diff = total_assets - total_liab_eq
            st.error(f"❌ الفرق: {diff:,.2f}")

 # --- 9. المشاريع والوحدات العقارية ---
    with tabs[8]:
        st.subheader("🏗️ إدارة المشاريع العقارية والوحدات")
        proj_sub1, proj_sub2, proj_sub3 = st.tabs([
            "🏢 بيانات المشاريع والأبراج",
            "🏠 وحدات المشروع والأسعار",
            "📊 تقرير أرباح وتكاليف المشروع",
        ])

        with proj_sub1:
            col_pr1, col_pr2 = st.columns([1, 2])
            with col_pr1:
                st.markdown("#### ➕ إضافة مشروع جديد")
                with st.form("add_project_form", clear_on_submit=True):
                    new_proj_code = st.text_input(
                        "كود المشروع (Project Code)", placeholder="مثال: PRJ-01"
                    )
                    new_proj_name = st.text_input(
                        "اسم المشروع", placeholder="مثال: كمبوند الياسمين"
                    )
                    new_tower_no = st.text_input(
                        "رقم البرج / المرحلة (Tower No)",
                        placeholder="مثال: برج A1",
                    )
                    new_proj_type = st.selectbox(
                        "نوع المشروع",
                        ["سكني", "تجاري", "إداري", "سكني تجاري"],
                    )
                    if st.form_submit_button("حفظ المشروع", type="primary"):
                        if new_proj_name.strip() and new_proj_code.strip():
                            try:
                                with engine.begin() as conn:
                                    conn.execute(
                                        text(
                                            "INSERT INTO projects (project_code, name, tower_no, project_type) VALUES (:c, :n, :t, :tp)"
                                        ),
                                        {
                                            "c": new_proj_code.strip(),
                                            "n": new_proj_name.strip(),
                                            "t": new_tower_no.strip(),
                                            "tp": new_proj_type,
                                        },
                                    )
                                st.success(
                                    f"✅ تم إضافة المشروع '{new_proj_name}' بكود '{new_proj_code}' بنجاح!"
                                )
                                st.rerun()
                            except Exception:
                                st.error("⚠️ اسم المشروع أو الكود مسجل بالفعل!")
                        else:
                            st.error("يرجى إدخال كود المشروع واسمه.")
            with col_pr2:
                st.markdown(
                    "#### 📋 قائمة المشاريع المسجلة (البحث بالكود أو الاسم)"
                )
                search_code = st.text_input(
                    "🔍 ادخل كود المشروع للوصول السريع:",
                    placeholder="اكتب كود المشروع هنا...",
                )
                projs_df = pd.read_sql(
                    'SELECT id, project_code AS "كود المشروع", name AS "اسم المشروع", tower_no AS "رقم البرج", project_type AS "النوع" FROM projects',
                    engine,
                )
                if search_code.strip():
                    projs_df = projs_df[
                        projs_df["كود المشروع"]
                        .fillna("")
                        .str.contains(search_code.strip(), case=False)
                    ]
                st.dataframe(projs_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("#### ✏️ تعديل أو حذف مشروع محدد")
            all_projs_edit = pd.read_sql(
                "SELECT id, project_code, name, tower_no, project_type FROM projects WHERE name != 'عام / الإدارة' ORDER BY id ASC",
                engine,
            )

            if not all_projs_edit.empty:
                proj_edit_map = {
                    f"{r['project_code'] or 'بدون كود'} - {r['name']}": int(r["id"])
                    for _, r in all_projs_edit.iterrows()
                }
                sel_edit_proj_label = st.selectbox(
                    "اختر المشروع للتعديل أو الحذف:",
                    list(proj_edit_map.keys()),
                    key="edit_del_proj_sel",
                )
                target_proj_id = proj_edit_map[sel_edit_proj_label]
                target_proj_row = all_projs_edit[
                    all_projs_edit["id"] == target_proj_id
                ].iloc[0]

                with st.form("edit_project_form"):
                    ep1, ep2, ep3, ep4 = st.columns(4)
                    upd_proj_code = ep1.text_input(
                        "كود المشروع", value=target_proj_row["project_code"] or ""
                    )
                    upd_proj_name = ep2.text_input(
                        "اسم المشروع", value=target_proj_row["name"] or ""
                    )
                    upd_tower_no = ep3.text_input(
                        "رقم البرج / المرحلة", value=target_proj_row["tower_no"] or ""
                    )
                    type_options = ["سكني", "تجاري", "إداري", "سكني تجاري"]
                    curr_type_idx = (
                        type_options.index(target_proj_row["project_type"])
                        if target_proj_row["project_type"] in type_options
                        else 0
                    )
                    upd_proj_type = ep4.selectbox(
                        "نوع المشروع", type_options, index=curr_type_idx
                    )

                    col_btn_u, col_btn_d = st.columns(2)
                    save_proj_btn = col_btn_u.form_submit_button(
                        "💾 حفظ تعديل المشروع", type="primary"
                    )
                    del_proj_btn = col_btn_d.form_submit_button("🗑️ حذف المشروع نهائياً")

                    if save_proj_btn:
                        if not upd_proj_name.strip() or not upd_proj_code.strip():
                            st.error("يرجى إدخال كود المشروع واسمه!")
                        else:
                            try:
                                old_name = target_proj_row["name"]
                                new_name = upd_proj_name.strip()
                                with engine.begin() as conn:
                                    conn.execute(
                                        text("""
                                            UPDATE projects
                                            SET project_code = :c, name = :n, tower_no = :t, project_type = :tp
                                            WHERE id = :id
                                        """),
                                        {
                                            "c": upd_proj_code.strip(),
                                            "n": new_name,
                                            "t": upd_tower_no.strip(),
                                            "tp": upd_proj_type,
                                            "id": int(target_proj_id),
                                        },
                                    )
                                    if old_name != new_name:
                                        conn.execute(
                                            text(
                                                "UPDATE units SET project_name = :nn WHERE project_name = :on"
                                            ),
                                            {"nn": new_name, "on": old_name},
                                        )
                                        conn.execute(
                                            text(
                                                "UPDATE installments SET project_name = :nn WHERE project_name = :on"
                                            ),
                                            {"nn": new_name, "on": old_name},
                                        )
                                        conn.execute(
                                            text(
                                                "UPDATE journal_entries SET project_name = :nn WHERE project_name = :on"
                                            ),
                                            {"nn": new_name, "on": old_name},
                                        )
                                st.success("✅ تم تحديث بيانات المشروع بنجاح!")
                                st.rerun()
                            except Exception:
                                st.error("⚠️ تعذر التعديل؛ ربما يكون كود أو اسم المشروع مستخدماً.")

                    if del_proj_btn:
                        proj_name_to_del = target_proj_row["name"]
                        with engine.connect() as conn:
                            has_units = (
                                conn.execute(
                                    text("SELECT COUNT(*) FROM units WHERE project_name = :p"),
                                    {"p": proj_name_to_del},
                                ).scalar()
                                or 0
                            )
                            has_entries = (
                                conn.execute(
                                    text(
                                        "SELECT COUNT(*) FROM journal_entries WHERE project_name = :p"
                                    ),
                                    {"p": proj_name_to_del},
                                ).scalar()
                                or 0
                            )
                            has_inst = (
                                conn.execute(
                                    text(
                                        "SELECT COUNT(*) FROM installments WHERE project_name = :p"
                                    ),
                                    {"p": proj_name_to_del},
                                ).scalar()
                                or 0
                            )

                        if has_units > 0 or has_entries > 0 or has_inst > 0:
                            st.error(
                                "🛑 لا يمكن حذف المشروع لارتباطه بوحدات أو قيود يومية أو أقساط مسجلة. عدّل بياناته أو احذف السجلات المرتبطة به أولاً."
                            )
                        else:
                            with engine.begin() as conn:
                                conn.execute(
                                    text("DELETE FROM projects WHERE id = :id"),
                                    {"id": int(target_proj_id)},
                                )
                            st.success("✅ تم حذف المشروع بنجاح!")
                            st.rerun()
            else:
                st.info("لا توجد مشاريع إضافية قابلة للتعديل أو الحذف حالياً.")

        with proj_sub2:
            st.markdown("#### 🏠 إضافة وإدارة وحدات المشروع")
            all_projs = pd.read_sql(
                "SELECT name, project_code FROM projects", engine
            )
            proj_options = [
                f"{r['project_code']} - {r['name']}"
                if r["project_code"]
                else r["name"]
                for _, r in all_projs.iterrows()
            ]
            proj_map = {
                (
                    f"{r['project_code']} - {r['name']}"
                    if r["project_code"]
                    else r["name"]
                ): r["name"]
                for _, r in all_projs.iterrows()
            }

            if proj_options:
                sel_proj_label = st.selectbox(
                    "اختر المشروع (بالكود أو الاسم):",
                    proj_options,
                    key="unit_proj_sel",
                )
                sel_proj_name = proj_map[sel_proj_label]

                with st.expander("➕ إضافة وحدة جديدة للمشروع", expanded=True):
                    with st.form("add_unit_form", clear_on_submit=True):
                        u_c1, u_c2, u_c3, u_c4 = st.columns(4)
                        u_code = u_c1.text_input(
                            "كود / رقم الوحدة", placeholder="مثال: A1-102"
                        )
                        u_area = u_c2.number_input(
                            "مساحة الوحدة (م²)",
                            min_value=1.0,
                            value=100.0,
                            step=5.0,
                        )
                        u_ppm = u_c3.number_input(
                            "سعر المتر",
                            min_value=0.0,
                            value=10000.0,
                            step=500.0,
                        )
                        u_status = u_c4.selectbox(
                            "حالة الوحدة", ["متاحة للبيع", "محجوزة", "مباعة"]
                        )
                        u_cust = st.text_input(
                            "اسم العميل (في حالة الحجز أو البيع - اختياري)"
                        )
                        if st.form_submit_button("💾 حفظ الوحدة", type="primary"):
                            if u_code.strip():
                                total_p = u_area * u_ppm
                                try:
                                    with engine.begin() as conn:
                                        conn.execute(
                                            text("""INSERT INTO units (project_name, unit_code, area, price_per_meter, total_price, status, customer_name)
                                                    VALUES (:p, :uc, :ar, :ppm, :tp, :st, :cn)"""),
                                            {
                                                "p": sel_proj_name,
                                                "uc": u_code.strip(),
                                                "ar": u_area,
                                                "ppm": u_ppm,
                                                "tp": total_p,
                                                "st": u_status,
                                                "cn": u_cust.strip(),
                                            },
                                        )
                                    st.success(
                                        f"✅ تم إضافة الوحدة '{u_code}' بسعر إجمالي {total_p:,.2f}"
                                    )
                                    st.rerun()
                                except Exception:
                                    st.error("⚠️ كود الوحدة مسجل بالفعل!")
                            else:
                                st.error("يرجى إدخال كود الوحدة.")

                units_df = pd.read_sql(
                    text("""
                    SELECT unit_code AS "كود الوحدة", area AS "المساحة (م²)", price_per_meter AS "سعر المتر",
                           total_price AS "سعر الوحدة الكامل", status AS "الحالة", customer_name AS "العميل"
                    FROM units WHERE project_name = :p
                """),
                    engine,
                    params={"p": sel_proj_name},
                )

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("إجمالي الوحدات", len(units_df))
                m2.metric(
                    "🟢 متاحة للبيع",
                    len(units_df[units_df["الحالة"] == "متاحة للبيع"]),
                )
                m3.metric(
                    "🟡 محجوزة", len(units_df[units_df["الحالة"] == "محجوزة"])
                )
                m4.metric(
                    "🔴 مباعة", len(units_df[units_df["الحالة"] == "مباعة"])
                )
                st.dataframe(units_df, use_container_width=True, hide_index=True)

        with proj_sub3:
            st.markdown("#### 📊 تقرير تكاليف وإيرادات كل مشروع على حدة")
            projs_list = pd.read_sql("SELECT name FROM projects", engine)[
                "name"
            ].tolist()
            if projs_list:
                chosen_proj = st.selectbox(
                    "اختر المشروع لعرض حساباته المستقلة:",
                    projs_list,
                    key="rep_proj_sel",
                )
                proj_details_df = pd.read_sql(
                    text("""
                    SELECT je.entry_date AS "التاريخ", a.code AS "كود الحساب", a.name AS "اسم الحساب",
                           a.account_type AS "نوع الحساب", je.description AS "البيان",
                           jd.debit AS "مدين", jd.credit AS "دائن"
                    FROM journal_details jd
                    JOIN journal_entries je ON jd.entry_id = je.id
                    JOIN accounts a ON jd.account_id = a.id
                    WHERE je.project_name = :p
                    ORDER BY je.entry_date DESC, je.id DESC
                """),
                    engine,
                    params={"p": chosen_proj},
                )

                tot_rev = (
                    proj_details_df[proj_details_df["نوع الحساب"] == "إيرادات"][
                        "دائن"
                    ].sum()
                    - proj_details_df[
                        proj_details_df["نوع الحساب"] == "إيرادات"
                    ]["مدين"].sum()
                )
                tot_exp = (
                    proj_details_df[
                        proj_details_df["نوع الحساب"] == "مصروفات"
                    ]["مدين"].sum()
                    - proj_details_df[
                        proj_details_df["نوع الحساب"] == "مصروفات"
                    ]["دائن"].sum()
                )
                c_p1, c_p2, c_p3 = st.columns(3)
                c_p1.metric("إيرادات المشروع", f"{tot_rev:,.2f}")
                c_p2.metric("مصروفات وتكاليف المشروع", f"{tot_exp:,.2f}")
                c_p3.metric("صافي ربح المشروع", f"{(tot_rev - tot_exp):,.2f}")
                st.dataframe(
                    proj_details_df, use_container_width=True, hide_index=True
                )

    # --- 10. نظام أقساط العملاء ---
    with tabs[9]:
        st.subheader("🗓️ إدارة ومتابعة أقساط العملاء")
        with st.expander("➕ إنشاء خطة أقساط جديدة لعقد عميل", expanded=False):
            with st.form("generate_installments_form"):
                c1, c2, c3, c4 = st.columns(4)
                proj_list = pd.read_sql("SELECT name FROM projects", engine)[
                    "name"
                ].tolist()
                inst_proj = c1.selectbox(
                    "المشروع (Project)",
                    proj_list if proj_list else ["عام / الإدارة"],
                )
                inst_cust = c2.text_input("اسم العميل (Customer)")
                inst_unit = c3.text_input(
                    "كود الوحدة (Unit Code)", placeholder="مثال: OB-QST-S048A"
                )
                inst_contract = c4.text_input(
                    "رقم العقد (Contract No)", placeholder="مثال: CNTRCT/078"
                )

                c5, c6, c7, c8 = st.columns(4)
                down_payment = c5.number_input(
                    "الدفعة المقدمة (Down Payment)",
                    min_value=0.0,
                    step=1000.0,
                )
                down_date = c6.date_input(
                    "تاريخ الدفعة المقدمة", value=date.today()
                )
                reg_amount = c7.number_input(
                    "قيمة القسط الدوري (Regular)",
                    min_value=0.0,
                    step=1000.0,
                )
                reg_count = c8.number_input(
                    "عدد الأقساط الدورية",
                    min_value=1,
                    max_value=120,
                    value=12,
                    step=1,
                )

                c9, c10, c11 = st.columns(3)
                first_reg_date = c9.date_input(
                    "تاريخ أول قسط دوري", value=date.today()
                )
                freq_label = c10.selectbox(
                    "دورية السداد",
                    [
                        "ربع سنوي (كل 3 شهور)",
                        "شهري",
                        "نصف سنوي (كل 6 شهور)",
                        "سنوي",
                    ],
                )
                collector_name = c11.text_input("مسؤول التحصيل (Collector)")

                submitted_plan = st.form_submit_button(
                    "🚀 توليد وحفظ جدول الأقساط", type="primary"
                )
                if submitted_plan:
                    if not inst_cust or not inst_unit:
                        st.error("يرجى إدخال اسم العميل وكود الوحدة!")
                    else:
                        freq_months = {
                            "شهري": 1,
                            "ربع سنوي (كل 3 شهور)": 3,
                            "نصف سنوي (كل 6 شهور)": 6,
                            "سنوي": 12,
                        }[freq_label]
                        with engine.begin() as conn:
                            if down_payment > 0:
                                conn.execute(
                                    text("""INSERT INTO installments 
                                            (project_name, customer_name, unit_code, contract_no, installment_type, due_date, amount, paid_amount, collector_name, status)
                                            VALUES (:p, :c, :u, :cn, 'Down Payment', :dd, :amt, 0, :col, 'Unpaid')"""),
                                    {
                                        "p": inst_proj,
                                        "c": inst_cust,
                                        "u": inst_unit,
                                        "cn": inst_contract,
                                        "dd": str(down_date),
                                        "amt": down_payment,
                                        "col": collector_name,
                                    },
                                )
                            for i in range(int(reg_count)):
                                months_add = i * freq_months
                                y = (
                                    first_reg_date.year
                                    + (first_reg_date.month + months_add - 1)
                                    // 12
                                )
                                m = (
                                    first_reg_date.month + months_add - 1
                                ) % 12 + 1
                                d = min(first_reg_date.day, 28)
                                due_d = date(y, m, d)
                                conn.execute(
                                    text("""INSERT INTO installments 
                                            (project_name, customer_name, unit_code, contract_no, installment_type, due_date, amount, paid_amount, collector_name, status)
                                            VALUES (:p, :c, :u, :cn, 'Regular', :dd, :amt, 0, :col, 'Unpaid')"""),
                                    {
                                        "p": inst_proj,
                                        "c": inst_cust,
                                        "u": inst_unit,
                                        "cn": inst_contract,
                                        "dd": str(due_d),
                                        "amt": reg_amount,
                                        "col": collector_name,
                                    },
                                )
                            conn.execute(
                                text(
                                    "UPDATE units SET status = 'مباعة', customer_name = :c WHERE unit_code = :u"
                                ),
                                {"c": inst_cust, "u": inst_unit},
                            )
                        st.success(
                            f"✅ تم توليد وحفظ خطة الأقساط للعميل {inst_cust} بنجاح!"
                        )
                        st.rerun()

        df_inst = pd.read_sql(
            "SELECT * FROM installments ORDER BY due_date ASC, id ASC", engine
        )
        if df_inst.empty:
            st.info(
                "لا توجد أقساط مسجلة حالياً. يمكنك إضافة أول خطة أقساط من القائمة أعلاه."
            )
        else:
            df_inst["remaining"] = df_inst["amount"] - df_inst["paid_amount"]

            f1, f2, f3, f4 = st.columns(4)
            sel_p = f1.selectbox(
                "فلتر المشروع",
                ["الكل"]
                + sorted(df_inst["project_name"].dropna().unique().tolist()),
            )
            sel_c = f2.selectbox(
                "فلتر العميل",
                ["الكل"]
                + sorted(df_inst["customer_name"].dropna().unique().tolist()),
            )
            sel_u = f3.text_input("بحث بكود الوحدة", "")
            sel_s = f4.selectbox(
                "حالة السداد", ["الكل", "Unpaid", "Partially Paid", "Paid"]
            )

            filtered = df_inst.copy()
            if sel_p != "الكل":
                filtered = filtered[filtered["project_name"] == sel_p]
            if sel_c != "الكل":
                filtered = filtered[filtered["customer_name"] == sel_c]
            if sel_u:
                filtered = filtered[
                    filtered["unit_code"].str.contains(
                        sel_u, case=False, na=False
                    )
                ]
            if sel_s != "الكل":
                filtered = filtered[filtered["status"] == sel_s]

            m1, m2, m3 = st.columns(3)
            m1.metric(
                "إجمالي قيمة الأقساط", f"{filtered['amount'].sum():,.2f}"
            )
            m2.metric(
                "إجمالي المحصل فعلياً", f"{filtered['paid_amount'].sum():,.2f}"
            )
            m3.metric(
                "إجمالي المتبقي للتحصيل", f"{filtered['remaining'].sum():,.2f}"
            )

            display_df = filtered[
                [
                    "id",
                    "project_name",
                    "customer_name",
                    "unit_code",
                    "installment_type",
                    "due_date",
                    "amount",
                    "paid_amount",
                    "remaining",
                    "collector_name",
                    "contract_no",
                    "status",
                ]
            ].rename(
                columns={
                    "id": "رقم القسط",
                    "project_name": "Project",
                    "customer_name": "Customer",
                    "unit_code": "Unit Code",
                    "installment_type": "Installment Type",
                    "due_date": "Due Date",
                    "amount": "Amount",
                    "paid_amount": "Paid Amount",
                    "remaining": "Remaining (المتبقي)",
                    "collector_name": "Collector",
                    "contract_no": "Contract No",
                    "status": "Status",
                }
            )
            st.dataframe(display_df, use_container_width=True, hide_index=True)

            st.markdown("---")
            st.markdown("#### 💵 تسجيل تحصيل دفعة من قسط")
            unpaid_df = filtered[filtered["status"] != "Paid"]
            if not unpaid_df.empty:
                col_p1, col_p2, col_p3 = st.columns([2, 1, 1])
                options = {
                    f"#{row['id']} | {row['customer_name']} - {row['unit_code']} ({row['installment_type']} - {row['due_date']}) | المتبقي: {row['remaining']:,.2f}": row[
                        "id"
                    ]
                    for _, row in unpaid_df.iterrows()
                }
                chosen_label = col_p1.selectbox(
                    "اختر القسط المراد تحصيله", list(options.keys())
                )
                chosen_id = options[chosen_label]
                chosen_row = unpaid_df[unpaid_df["id"] == chosen_id].iloc[0]
                pay_now = col_p2.number_input(
                    "المبلغ المحصل الآن",
                    min_value=0.0,
                    max_value=float(chosen_row["remaining"]),
                    value=float(chosen_row["remaining"]),
                    step=500.0,
                )
                col_p3.write("")
                col_p3.write("")
                if col_p3.button("✅ تأكيد التحصيل", type="primary"):
                    new_paid = float(chosen_row["paid_amount"]) + pay_now
                    new_status = (
                        "Paid"
                        if new_paid >= float(chosen_row["amount"]) - 0.01
                        else ("Partially Paid" if new_paid > 0 else "Unpaid")
                    )
                    with engine.begin() as conn:
                        conn.execute(
                            text(
                                "UPDATE installments SET paid_amount = :pa, status = :st WHERE id = :id"
                            ),
                            {
                                "pa": new_paid,
                                "st": new_status,
                                "id": int(chosen_id),
                            },
                        )
                    st.success("تم تسجيل التحصيل وتحديث حالة القسط بنجاح!")
                    st.rerun()

            st.markdown("---")
            st.markdown("#### 🧾 كشف حساب تفصيلي للعميل")
            cust_list = sorted(
                df_inst["customer_name"].dropna().unique().tolist()
            )
            if cust_list:
                chosen_cust = st.selectbox(
                    "اختر اسم العميل لعرض كشف حسابه:",
                    cust_list,
                    key="cust_stmt_sel",
                )
                cust_df = df_inst[
                    df_inst["customer_name"] == chosen_cust
                ].copy()
                cust_units = ", ".join(
                    cust_df["unit_code"].dropna().unique().tolist()
                )
                cust_projects = ", ".join(
                    cust_df["project_name"].dropna().unique().tolist()
                )

                st.info(
                    f"**العميل:** {chosen_cust} | **المشروع:** {cust_projects} | **الوحدة / الوحدات:** {cust_units}"
                )

                c_tot = cust_df["amount"].sum()
                c_paid = cust_df["paid_amount"].sum()
                c_rem = cust_df["remaining"].sum()

                cs1, cs2, cs3 = st.columns(3)
                cs1.metric("إجمالي قيمة التعاقد / الأقساط", f"{c_tot:,.2f}")
                cs2.metric("إجمالي المسدد من العميل", f"{c_paid:,.2f}")
                cs3.metric("إجمالي المتبقي على العميل", f"{c_rem:,.2f}")

                stmt_view = cust_df[
                    [
                        "id",
                        "project_name",
                        "unit_code",
                        "contract_no",
                        "installment_type",
                        "due_date",
                        "amount",
                        "paid_amount",
                        "remaining",
                        "status",
                        "collector_name",
                    ]
                ].rename(
                    columns={
                        "id": "رقم القسط",
                        "project_name": "المشروع",
                        "unit_code": "كود الوحدة",
                        "contract_no": "رقم العقد",
                        "installment_type": "نوع القسط",
                        "due_date": "تاريخ الاستحقاق",
                        "amount": "قيمة القسط",
                        "paid_amount": "المسدد",
                        "remaining": "المتبقي",
                        "status": "الحالة",
                        "collector_name": "المحصّل",
                    }
                )
                st.dataframe(stmt_view, use_container_width=True, hide_index=True)
                st.download_button(
                    label=f"📥 تصدير كشف حساب العميل ({chosen_cust}) إلى Excel",
                    data=to_excel(stmt_view),
                    file_name=f"statement_{chosen_cust}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

st.markdown("---")
st.markdown("#### ✏️ تعديل قسط محدد")
        all_inst_df = pd.read_sql(
            """
            SELECT id, customer_name, unit_code, contract_no, installment_type,
                   due_date, amount, paid_amount, collector_name
            FROM installments
            ORDER BY id DESC
        """,
            engine,
        )

        if not all_inst_df.empty:
          inst_edit_options = {
              (
                  f"#{int(r['id'])} | {r['customer_name']} | {r['unit_code']} |"
                  f" {r['installment_type']} | استحقاق: {r['due_date']} |"
                  f" قيمة: {float(r['amount'] or 0):,.2f}"
              ): int(r["id"])
              for _, r in all_inst_df.iterrows()
          }
          sel_inst_label = st.selectbox(
              "اختر القسط للتعديل:",
              list(inst_edit_options.keys()),
              key="edit_inst_sel",
          )
          target_inst_id = inst_edit_options[sel_inst_label]
          target_inst_row = all_inst_df[
              all_inst_df["id"] == target_inst_id
          ].iloc[0]

          with st.form("edit_installment_form"):
            ic1, ic2, ic3, ic4 = st.columns(4)
            upd_cust = ic1.text_input(
                "اسم العميل", value=target_inst_row["customer_name"] or ""
            )
            upd_unit = ic2.text_input(
                "كود الوحدة", value=target_inst_row["unit_code"] or ""
            )
            upd_contract = ic3.text_input(
                "رقم العقد", value=target_inst_row["contract_no"] or ""
            )
            inst_types = ["Down Payment", "Regular"]
            curr_type_i = (
                inst_types.index(target_inst_row["installment_type"])
                if target_inst_row["installment_type"] in inst_types
                else 1
            )
            upd_inst_type = ic4.selectbox(
                "نوع القسط", inst_types, index=curr_type_i
            )

            ic5, ic6, ic7, ic8 = st.columns(4)
            try:
              parsed_due = date.fromisoformat(
                  str(target_inst_row["due_date"])[:10]
              )
            except Exception:
              parsed_due = date.today()
            upd_due_date = ic5.date_input("تاريخ الاستحقاق", value=parsed_due)
            upd_amount = ic6.number_input(
                "قيمة القسط",
                min_value=0.0,
                value=float(target_inst_row["amount"] or 0.0),
                step=500.0,
            )
            upd_paid = ic7.number_input(
                "المبلغ المسدد",
                min_value=0.0,
                value=float(target_inst_row["paid_amount"] or 0.0),
                step=500.0,
            )
            upd_collector = ic8.text_input(
                "مسؤول التحصيل",
                value=target_inst_row["collector_name"] or "",
            )

            save_inst_btn = st.form_submit_button(
                "💾 حفظ تعديل القسط", type="primary"
            )

            if save_inst_btn:
              if not upd_cust.strip() or not upd_unit.strip():
                st.error("يرجى إدخال اسم العميل وكود الوحدة!")
              else:
                calc_status = (
                    "Paid"
                    if upd_paid >= upd_amount - 0.01 and upd_amount > 0
                    else ("Partially Paid" if upd_paid > 0 else "Unpaid")
                )
                with engine.begin() as conn:
                  conn.execute(
                      text("""
                                    UPDATE installments
                                    SET customer_name = :c,
                                        unit_code = :u,
                                        contract_no = :cn,
                                        installment_type = :it,
                                        due_date = :dd,
                                        amount = :amt,
                                        paid_amount = :pa,
                                        collector_name = :col,
                                        status = :st
                                    WHERE id = :id
                                """),
                      {
                          "c": upd_cust.strip(),
                          "u": upd_unit.strip(),
                          "cn": upd_contract.strip(),
                          "it": upd_inst_type,
                          "dd": str(upd_due_date),
                          "amt": float(upd_amount),
                          "pa": float(upd_paid),
                          "col": upd_collector.strip(),
                          "st": calc_status,
                          "id": int(target_inst_id),
                      },
                  )
                st.success(f"✅ تم تحديث القسط رقم #{target_inst_id} بنجاح!")
                st.rerun()
        else:
          st.info("لا توجد أقساط مسجلة حالياً للتعديل.")

        st.markdown("---")
        st.markdown("#### 🗑️ حذف قسط واحد أو عدة أقساط دفعة واحدة")
        all_inst_del_df = pd.read_sql(
            """
            SELECT id, project_name, customer_name, unit_code, contract_no,
                   installment_type, due_date, amount, paid_amount, status
            FROM installments
            ORDER BY due_date ASC, id ASC
        """,
            engine,
        )

        if not all_inst_del_df.empty:
          multi_inst_map = {
              (
                  f"#{int(r['id'])} | {r['customer_name']} | {r['unit_code']} |"
                  f" {r['installment_type']} | استحقاق: {r['due_date']} |"
                  f" قيمة: {float(r['amount'] or 0):,.2f} | الحالة:"
                  f" {r['status']}"
              ): int(r["id"])
              for _, r in all_inst_del_df.iterrows()
          }

          selected_inst_labels = st.multiselect(
              "اختر الأقساط المطلوب حذفها (يمكنك تحديد أكثر من قسط):",
              options=list(multi_inst_map.keys()),
              key="multi_delete_inst_sel",
          )

          if selected_inst_labels:
            selected_ids = [
                int(multi_inst_map[lbl]) for lbl in selected_inst_labels
            ]
            selected_rows_df = all_inst_del_df[
                all_inst_del_df["id"].isin(selected_ids)
            ]
            total_del_amt = float(selected_rows_df["amount"].sum())

            st.warning(
                f"تم تحديد **{len(selected_ids)}** قسط بإجمالي قيمة"
                f" **{total_del_amt:,.2f}**."
            )
            st.dataframe(
                selected_rows_df[[
                    "id",
                    "customer_name",
                    "unit_code",
                    "installment_type",
                    "due_date",
                    "amount",
                    "paid_amount",
                    "status",
                ]],
                use_container_width=True,
                hide_index=True,
            )

            confirm_bulk_del = st.checkbox(
                "أؤكد رغبتي في حذف الأقساط المحددة نهائياً من قاعدة البيانات",
                key="confirm_bulk_inst_del",
            )

            if st.button(
                f"🗑️ حذف ({len(selected_ids)}) قسط محدد",
                type="primary",
                disabled=not confirm_bulk_del,
                key="btn_bulk_delete_inst",
            ):
              with engine.begin() as conn:
                for inst_id in selected_ids:
                  conn.execute(
                      text("DELETE FROM installments WHERE id = :id"),
                      {"id": inst_id},
                  )
              st.success(f"✅ تم حذف {len(selected_ids)} قسط بنجاح!")
              st.rerun()
        else:
          st.info("لا توجد أقساط مسجلة حالياً للحذف.")
        # --- 11. إدارة المستخدمين ---
    if is_admin:
        with tabs[10]:
            col_u_add, col_u_edit = st.columns(2)
            with col_u_add:
                st.subheader("➕ إضافة مستخدم جديد")
                with st.form("new_user_form", clear_on_submit=True):
                    new_username = st.text_input("اسم المستخدم")
                    new_password = st.text_input("كلمة السر", type="password")
                    new_role = st.selectbox(
                        "الصلاحية", ["مدير النظام", "محاسب / مدخل بيانات"]
                    )
                    if st.form_submit_button("إضافة المستخدم"):
                        if new_username and new_password:
                            try:
                                with engine.begin() as conn:
                                    conn.execute(
                                        text(
                                            "INSERT INTO users (username, password, role) VALUES (:u, :p, :r)"
                                        ),
                                        {
                                            "u": new_username,
                                            "p": make_hashes(new_password),
                                            "r": new_role,
                                        },
                                    )
                                st.success("✅ تم الإضافة بنجاح!")
                                st.rerun()
                            except Exception:
                                st.error("المستخدم موجود مسبقاً!")

            with col_u_edit:
                st.subheader("⚙️ تعديل بيانات مستخدم")
                all_users_df = pd.read_sql(
                    "SELECT id, username, role FROM users", engine
                )
                if not all_users_df.empty:
                    user_dict = {
                        f"{row['username']} ({row['role']})": row["username"]
                        for _, row in all_users_df.iterrows()
                    }
                    selected_u_label = st.selectbox(
                        "اختر المستخدم:", list(user_dict.keys())
                    )
                    target_username = user_dict[selected_u_label]
                    with st.form("edit_user_form"):
                        updated_username = st.text_input(
                            "الاسم المعدل", value=target_username
                        )
                        updated_password = st.text_input(
                            "كلمة السر الجديدة", type="password"
                        )
                        updated_role = st.selectbox(
                            "الصلاحية الجديدة",
                            ["مدير النظام", "محاسب / مدخل بيانات"],
                        )
                        if st.form_submit_button("حفظ التغييرات"):
                            with engine.begin() as conn:
                                if updated_password:
                                    conn.execute(
                                        text(
                                            "UPDATE users SET username = :nu, password = :np, role = :nr WHERE username = :ou"
                                        ),
                                        {
                                            "nu": updated_username,
                                            "np": make_hashes(updated_password),
                                            "nr": updated_role,
                                            "ou": target_username,
                                        },
                                    )
                                else:
                                    conn.execute(
                                        text(
                                            "UPDATE users SET username = :nu, role = :nr WHERE username = :ou"
                                        ),
                                        {
                                            "nu": updated_username,
                                            "nr": updated_role,
                                            "ou": target_username,
                                        },
                                    )
                            st.success("✅ تم التحديث بنجاح!")
                            st.rerun()

            st.markdown("---")
            users_df = pd.read_sql(
                'SELECT id AS "رقم", username AS "اسم المستخدم", role AS "الصلاحية" FROM users',
                engine,
            )
            st.dataframe(users_df, use_container_width=True)

        # --- 12. التحديث التلقائي الـ AI ---
        with tabs[11]:
            st.subheader(
                "🤖 تعديل وتحديث الكود البرمجي تلقائياً بواسطة AI"
            )
            st.info(
                "ملاحظة: عند التشغيل السحابي، يولد الذكاء الاصطناعي الكود المحدث لتقوم بنسخه وتحديثه في ملفك على GitHub."
            )

            api_key = st.text_input(
                "مفتاح OpenAI API Key الخاص بك:", type="password"
            )
            ai_instruction = st.text_area(
                "ما التعديل أو الميزة التي ترغب في إضافتها إلى الكود؟"
            )

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
                                temperature=0.2,
                            )
                            new_code = response.choices[0].message.content
                            if "```python" in new_code:
                                new_code = new_code.split("```python")[1].split(
                                    "```"
                                )[0]
                            elif "```" in new_code:
                                new_code = new_code.split("```")[1].split("```")[
                                    0
                                ]

                        st.success(
                            "✅ تم توليد الكود بنجاح! انسخه وضعه في ملف app.py على GitHub:"
                        )
                        st.code(new_code.strip(), language="python")
                    except Exception as e:
                        st.error(f"❌ حدث خطأ: {str(e)}")

        # --- 13. النسخ الاحتياطي التلقائي عبر البريد الإلكتروني ---
        with tabs[12]:
            st.subheader("💾 النسخ الاحتياطي وإرساله عبر البريد الإلكتروني")
            st.info(
                "يتم حفظ إعدادات الإرسال اليومي في قاعدة البيانات، ويُرسل ملف النسخة الاحتياطية (JSON + Excel) تلقائياً في الموعد اليومي المحدد (مثل نهاية اليوم 23:55) طالما التطبيق نشط."
            )

            sec_cfg = get_smtp_config()
            with engine.connect() as conn:
                cfg_row = conn.execute(
                    text("""
                        SELECT auto_enabled, send_hour, send_minute, smtp_server, smtp_port,
                               sender_email, sender_password, recipient_email, last_sent_date
                        FROM backup_settings WHERE id = 1
                    """)
                ).fetchone()

            cur_auto = bool(cfg_row[0] == 1) if cfg_row else False
            cur_hour = int(cfg_row[1]) if cfg_row and cfg_row[1] is not None else 23
            cur_min = int(cfg_row[2]) if cfg_row and cfg_row[2] is not None else 55
            cur_server = (cfg_row[3] if cfg_row and cfg_row[3] else sec_cfg["smtp_server"])
            cur_port = int(cfg_row[4]) if cfg_row and cfg_row[4] else sec_cfg["smtp_port"]
            cur_sender = (cfg_row[5] if cfg_row and cfg_row[5] else sec_cfg["sender_email"])
            cur_pass = (cfg_row[6] if cfg_row and cfg_row[6] else sec_cfg["sender_password"])
            cur_rec = (cfg_row[7] if cfg_row and cfg_row[7] else sec_cfg["recipient_email"])
            last_sent = cfg_row[8] if cfg_row and cfg_row[8] else "لم يتم الإرسال بعد"

            st.caption(f"آخر إرسال تلقائي ناجح: **{last_sent}**")

            with st.form("backup_email_settings_form"):
                auto_enabled = st.checkbox(
                    "⏰ تفعيل الإرسال التلقائي اليومي في نهاية اليوم",
                    value=cur_auto,
                )
                col_tm1, col_tm2 = st.columns(2)
                send_time_val = col_tm1.time_input(
                    "موعد الإرسال اليومي",
                    value=dt_time(hour=cur_hour, minute=cur_min),
                )
                smtp_port_val = col_tm2.number_input(
                    "منفذ SMTP Port",
                    min_value=1,
                    max_value=65535,
                    value=cur_port,
                )

                smtp_server_val = st.text_input(
                    "خادم SMTP Server", value=cur_server
                )
                col_em1, col_em2 = st.columns(2)
                sender_email_val = col_em1.text_input(
                    "البريد المرسل منه (Sender Email)", value=cur_sender
                )
                sender_pass_val = col_em2.text_input(
                    "كلمة مرور التطبيقات (App Password)",
                    value=cur_pass,
                    type="password",
                )
                recipient_email_val = st.text_input(
                    "البريد المستلم للنسخة الاحتياطية (Recipient Email)",
                    value=cur_rec,
                )

                save_bkp_cfg = st.form_submit_button(
                    "💾 حفظ إعدادات النسخ الاحتياطي", type="primary"
                )
                if save_bkp_cfg:
                    with engine.begin() as conn:
                        conn.execute(
                            text("""
                                UPDATE backup_settings
                                SET auto_enabled = :ae,
                                    send_hour = :sh,
                                    send_minute = :sm,
                                    smtp_server = :srv,
                                    smtp_port = :prt,
                                    sender_email = :se,
                                    sender_password = :sp,
                                    recipient_email = :re
                                WHERE id = 1
                            """),
                            {
                                "ae": 1 if auto_enabled else 0,
                                "sh": int(send_time_val.hour),
                                "sm": int(send_time_val.minute),
                                "srv": smtp_server_val.strip(),
                                "prt": int(smtp_port_val),
                                "se": sender_email_val.strip(),
                                "sp": sender_pass_val,
                                "re": recipient_email_val.strip(),
                            },
                        )
                    st.success("✅ تم حفظ إعدادات النسخ الاحتياطي بنجاح!")
                    st.rerun()

            st.markdown("---")
            col_act1, col_act2 = st.columns(2)
            with col_act1:
                if st.button("📧 إرسال نسخة احتياطية الآن عبر الإيميل"):
                    if not sender_email_val or not sender_pass_val or not recipient_email_val:
                        st.error("يرجى إدخال البريد المرسل وكلمة المرور والبريد المستلم أولاً.")
                    else:
                        try:
                            with st.spinner("جارٍ إنشاء وإرسال النسخة الاحتياطية..."):
                                send_backup_email(
                                    smtp_server=smtp_server_val.strip(),
                                    smtp_port=int(smtp_port_val),
                                    sender_email=sender_email_val.strip(),
                                    sender_password=sender_pass_val,
                                    recipient_email=recipient_email_val.strip(),
                                    subject_prefix="نسخة احتياطية يدوية",
                                )
                            st.success(
                                f"✅ تم إرسال النسخة الاحتياطية إلى {recipient_email_val.strip()} بنجاح!"
                            )
                        except Exception as e:
                            st.error(f"❌ فشل إرسال البريد: {str(e)}")

            with col_act2:
                json_bkp, excel_bkp = build_backup_payload()
                st.download_button(
                    label="⬇️ تحميل نسخة احتياطية فورية (Excel)",
                    data=excel_bkp,
                    file_name=f"full_backup_{date.today()}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                st.download_button(
                    label="⬇️ تحميل نسخة احتياطية فورية (JSON)",
                    data=json_bkp,
                    file_name=f"full_backup_{date.today()}.json",
                    mime="application/json",
                )
