
import streamlit as st
import pandas as pd
import sqlite3
from datetime import datetime, date
from io import BytesIO
import hashlib
import altair as alt

# -----------------------
# Config & constants
# -----------------------
DB_PATH = "data.db"
APP_TITLE = "IARPF – MICROTROL"
DEFAULT_ADMIN_PASS = "admin123"
MONTH_NAMES = ["January","February","March","April","May","June","July","August","September","October","November","December"]

# -----------------------
# Database helpers
# -----------------------
def get_conn():
    return sqlite3.connect(DB_PATH, check_same_thread=False)

def init_db():
    con = get_conn(); cur = con.cursor()
    cur.execute("""CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE,
        pass_hash TEXT,
        is_admin INTEGER DEFAULT 0,
        created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS bescom (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER, month INTEGER, kwh REAL, bill REAL, md_kva REAL, notes TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS process_analysis (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT, year INTEGER, month INTEGER,
        customer_name TEXT, product_name TEXT, mode TEXT,
        length_mm REAL, width_mm REAL, height_mm REAL, weight_kg REAL, density_gcc REAL,
        speed_m_min REAL, no_of_passes INTEGER, irradiation_side INTEGER, batch_time_min REAL,
        processed_boxes INTEGER, batches INTEGER, productive_hrs REAL, productive_hrs_gs REAL,
        total_productive_hrs REAL, throughput_per_hr REAL, no_of_batches INTEGER,
        notes TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS procurement (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER, month INTEGER, item_desc TEXT, amount REAL, date TEXT, notes TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS maintenance (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        date TEXT, year INTEGER, month INTEGER, category TEXT, amount REAL, remarks TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS dosimeter (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER, month INTEGER, num_used INTEGER, batch_type TEXT, notes TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS rrcat (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER, month INTEGER, item_desc TEXT, amount REAL, date TEXT, notes TEXT, created_at TEXT
    )""")
    cur.execute("""CREATE TABLE IF NOT EXISTS infra (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        year INTEGER, month INTEGER, work_desc TEXT, vendor TEXT, amount REAL, date TEXT, notes TEXT, created_at TEXT
    )""")
    # default admin
    cur.execute("SELECT COUNT(*) FROM users")
    try:
        cnt = cur.fetchone()[0]
    except Exception:
        cnt = 0
    if cnt == 0:
        h = hashlib.sha256(DEFAULT_ADMIN_PASS.encode("utf-8")).hexdigest()
        cur.execute("INSERT OR IGNORE INTO users (username, pass_hash, is_admin, created_at) VALUES (?,?,1,?)",
                    ("admin", h, datetime.utcnow().isoformat()))
    con.commit(); con.close()

def hash_pw(pw: str) -> str:
    return hashlib.sha256(pw.encode("utf-8")).hexdigest()

def check_login(username: str, password: str) -> bool:
    con = get_conn(); cur = con.cursor()
    cur.execute("SELECT pass_hash FROM users WHERE username = ?", (username,))
    r = cur.fetchone(); con.close()
    if not r: return False
    return hash_pw(password) == r[0]

def is_admin_user(username: str) -> bool:
    con = get_conn(); cur = con.cursor()
    cur.execute("SELECT is_admin FROM users WHERE username = ?", (username,))
    r = cur.fetchone(); con.close()
    return bool(r and r[0] == 1)

# -----------------------
# Utilities
# -----------------------
def to_excel_bytes(df: pd.DataFrame) -> bytes:
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="data")
    return out.getvalue()

def fetch_df(table_name, limit=10000):
    con = get_conn()
    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name} ORDER BY created_at DESC LIMIT {limit}", con)
    except Exception:
        df = pd.DataFrame()
    con.close(); return df

def fetch_df_full(table_name):
    con = get_conn()
    try:
        df = pd.read_sql_query(f"SELECT * FROM {table_name} ORDER BY created_at DESC", con)
    except Exception:
        df = pd.DataFrame()
    con.close(); return df

def delete_row(table, row_id):
    con = get_conn(); cur = con.cursor()
    cur.execute(f"DELETE FROM {table} WHERE id = ?", (row_id,))
    con.commit(); con.close()

def delete_all(table):
    con = get_conn(); cur = con.cursor()
    cur.execute(f"DELETE FROM {table}")
    con.commit(); con.close()

# -----------------------
# App init
# -----------------------
st.set_page_config(page_title=APP_TITLE, layout="wide")
init_db()
if "logged_in" not in st.session_state: st.session_state.logged_in = False
if "username" not in st.session_state: st.session_state.username = ""
if "pending_delete" not in st.session_state: st.session_state.pending_delete = None
if "pending_table" not in st.session_state: st.session_state.pending_table = None
if "selected_analysis" not in st.session_state: st.session_state.selected_analysis = None

# -----------------------
# Login
# -----------------------
if not st.session_state.logged_in:
    st.title(APP_TITLE)
    st.write("Operations & Process Tracking")
    with st.form("login"):
        uname = st.text_input("Username", value="admin")
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Login"):
            if check_login(uname.strip(), pw.strip()):
                st.session_state.logged_in = True; st.session_state.username = uname.strip(); st.rerun()
            else:
                st.error("Invalid credentials")
    st.stop()

# -----------------------
# Sidebar/header
# -----------------------
st.sidebar.title(APP_TITLE)
menu = st.sidebar.radio("Navigation", ("Dashboard","Data Entry","Visualization","Admin"))

h1, h2 = st.columns([4,1])
with h1: st.title("Dashboard")  # heading required as user asked
with h2:
    st.write(f"**{st.session_state.username}**")
    if st.button("Logout"):
        st.session_state.logged_in = False; st.session_state.username = ""; st.rerun()

st.markdown("---")

# pending delete confirmation
if st.session_state.pending_table:
    st.warning("Delete action pending. Enter admin credentials to confirm.")
    with st.form("confirm_delete"):
        user = st.text_input("Admin username"); pwd = st.text_input("Admin password", type="password")
        if st.form_submit_button("Confirm Delete"):
            if user and pwd and check_login(user.strip(), pwd.strip()) and is_admin_user(user.strip()):
                if st.session_state.pending_delete is None:
                    delete_all(st.session_state.pending_table); st.success(f"All records deleted from {st.session_state.pending_table}")
                else:
                    delete_row(st.session_state.pending_table, int(st.session_state.pending_delete)); st.success(f"Deleted id={st.session_state.pending_delete} from {st.session_state.pending_table}")
                st.session_state.pending_delete = None; st.session_state.pending_table = None; st.rerun()
            else:
                st.error("Invalid admin credentials or not admin")
        if st.form_submit_button("Cancel"):
            st.session_state.pending_delete = None; st.session_state.pending_table = None; st.rerun()

# -----------------------
# Dashboard content (horizontal analysis select below heading)
# -----------------------
if menu == "Dashboard":
    # analysis selector (below heading as requested)
    analysis_type = st.selectbox("Choose analysis", [
        "Process Analysis (monthly)","BESCOM Analysis (monthly)","Maintenance Analysis (monthly)",
        "Dosimeter Analysis (monthly)","Procurement Analysis (monthly)","RRCAT Analysis (monthly)",
        "Infrastructure Analysis (monthly)"], index=0)
    st.session_state.selected_analysis = analysis_type

    # load data
    bes = fetch_df("bescom"); proc = fetch_df("process_analysis"); maint = fetch_df("maintenance")
    dos = fetch_df("dosimeter"); pr = fetch_df("procurement"); rrc = fetch_df("rrcat"); infra = fetch_df("infra")

    # compute latest monthly aggregates helper
    def latest_month_agg(df, cols):
        if df.empty: return None, None, None
        d = df.copy()
        if "date" in d.columns:
            d["date"]=pd.to_datetime(d["date"], errors="coerce"); d["year"]=d["date"].dt.year; d["month"]=d["date"].dt.month
        if "year" not in d.columns or "month" not in d.columns: return None, None, None
        y = int(d["year"].max()); m = int(d[d["year"]==y]["month"].max())
        filt = d[(d["year"]==y)&(d["month"]==m)]
        agg = {c: float(filt[c].sum()) if c in filt.columns else 0.0 for c in cols}
        return y, m, agg

    p_y,p_m,p_agg = latest_month_agg(proc, ["processed_boxes","batches","productive_hrs"])
    b_y,b_m,b_agg = latest_month_agg(bes, ["kwh","bill","md_kva"])

    # KPI row horizontally
    cols = st.columns(3)
    if analysis_type == "Process Analysis (monthly)":
        cols[0].markdown(f"**Processed Boxes**\n\n### {int(p_agg.get('processed_boxes',0)) if p_agg else '—'}")
        cols[1].markdown(f"**Batches**\n\n### {int(p_agg.get('batches',0)) if p_agg else '—'}")
        cols[2].markdown(f"**Productive Hrs**\n\n### {p_agg.get('productive_hrs',0) if p_agg else '—'}")
        # main chart
        if not proc.empty:
            dfp = proc.copy(); dfp["date"]=pd.to_datetime(dfp["date"], errors="coerce"); dfp["year_month"]=dfp["date"].dt.to_period("M").astype(str)
            monthly = dfp.groupby("year_month", as_index=False)["processed_boxes"].sum().sort_values("year_month")
            try: st.altair_chart(alt.Chart(monthly).mark_line(point=True).encode(x="year_month:N", y="processed_boxes:Q"), use_container_width=True)
            except Exception: pass

    elif analysis_type == "BESCOM Analysis (monthly)":
        cols[0].markdown(f"**kWh**\n\n### {b_agg.get('kwh',0) if b_agg else '—'}")
        cols[1].markdown(f"**Bill (₹)**\n\n### {b_agg.get('bill',0) if b_agg else '—'}")
        cols[2].markdown(f"**MD (kVA)**\n\n### {b_agg.get('md_kva',0) if b_agg else '—'}")
        if not bes.empty:
            btmp = bes.copy(); btmp["date"]=pd.to_datetime(btmp.get("date", None), errors="coerce"); btmp["year_month"]=btmp["date"].dt.to_period("M").astype(str)
            monthly = btmp.groupby("year_month", as_index=False)["kwh"].sum().sort_values("year_month")
            try: st.altair_chart(alt.Chart(monthly).mark_bar().encode(x="year_month:N", y="kwh:Q"), use_container_width=True)
            except Exception: pass

    else:
        # other analysis show a single KPI in first column
        if analysis_type == "Maintenance Analysis (monthly)":
            val = maint["amount"].sum() if not maint.empty else 0.0
            cols[0].markdown(f"**Maintenance (₹)**\n\n### {val:.2f}")
        if analysis_type == "Procurement Analysis (monthly)":
            val = pr["amount"].sum() if not pr.empty else 0.0
            cols[0].markdown(f"**Procurement (₹)**\n\n### {val:.2f}")
        if analysis_type == "Dosimeter Analysis (monthly)":
            val = dos["num_used"].sum() if not dos.empty else 0
            cols[0].markdown(f"**Dosimeters used**\n\n### {val}")

    # details table and downloads
    if analysis_type == "Process Analysis (monthly)":
        if not proc.empty:
            dfp = proc.copy(); dfp["date"]=pd.to_datetime(dfp["date"], errors="coerce"); dfp["year"]=dfp["date"].dt.year; dfp["month"]=dfp["date"].dt.month
            monthly = dfp.groupby(["year","month"], as_index=False).agg({"processed_boxes":"sum","batches":"sum","productive_hrs":"sum"})
            monthly["month_name"] = monthly["month"].apply(lambda m: MONTH_NAMES[m-1])
            st.dataframe(monthly, use_container_width=True); st.download_button("Export Process Monthly", data=to_excel_bytes(monthly), file_name="process_monthly.xlsx")
        else: st.info("No process data.")

    if analysis_type == "BESCOM Analysis (monthly)":
        if not bes.empty:
            bes2 = bes.copy(); bes2["month_name"] = bes2["month"].apply(lambda m: MONTH_NAMES[m-1])
            st.dataframe(bes2, use_container_width=True); st.download_button("Export BESCOM Monthly", data=to_excel_bytes(bes2), file_name="bescom_monthly.xlsx")
        else: st.info("No BESCOM data.")

# -----------------------
# Data Entry
# -----------------------
if menu == "Data Entry":
    st.header("Data Entry")
    section = st.selectbox("Select Section", [
        "BESCOM (monthly)","Process Analysis (daily)","Procurement (monthly)","Maintenance (daily)","Dosimeter (monthly)","RRCAT (monthly)","Building / Infrastructure (monthly)"])
    today = date.today()

    def show_table_with_delete(table_name):
        df = fetch_df(table_name)
        if df.empty: st.info("No records found."); return
        st.write("### Existing Records")
        for i,row in df.iterrows():
            c1,c2 = st.columns([7,1])
            with c1: st.json(row.to_dict())
            with c2:
                if st.button("Delete", key=f"del_{table_name}_{row['id']}"):
                    st.session_state.pending_delete = int(row['id']); st.session_state.pending_table = table_name
        st.warning("Delete ALL records (irreversible!)")
        if st.button(f"Delete ALL from {table_name}_btn"):
            st.session_state.pending_delete = None; st.session_state.pending_table = table_name

    # BESCOM entry
    if section == "BESCOM (monthly)":
        st.subheader("BESCOM (monthly)")
        with st.form("bescom_form"):
            byear = st.number_input("Year", value=today.year, min_value=2000, max_value=2100)
            selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
            bmonth = MONTH_NAMES.index(selected) + 1
            kwh = st.number_input("kWh", min_value=0.0, format="%.3f")
            bill = st.number_input("Electricity bill (₹)", min_value=0.0, format="%.2f")
            md_kva = st.number_input("MD (kVA)", min_value=0.0, format="%.3f")
            notes = st.text_area("Notes (optional)")
            if st.form_submit_button("Save BESCOM"):
                con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO bescom (year,month,kwh,bill,md_kva,notes,created_at) VALUES (?,?,?,?,?,?,?)", (int(byear), int(bmonth), float(kwh), float(bill), float(md_kva), notes, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("Saved BESCOM record")
        show_table_with_delete("bescom")

    # Process Analysis (daily)
    if section == "Process Analysis (daily)":
        st.subheader("Process Analysis — Daily Entry")
        with st.form("proc_form"):
            p_date = st.date_input("Date", value=today)
            p_year = p_date.year; p_month = p_date.month
            customer_name = st.text_input("Customer name"); product_name = st.text_input("Product name")
            mode = st.selectbox("Mode", ["R&D/Trial","DE","DM","RP"])
            lcol,mcol,rcol = st.columns(3)
            length_mm = lcol.number_input("Length (mm)", min_value=0.0); width_mm = mcol.number_input("Width (mm)", min_value=0.0); height_mm = rcol.number_input("Height (mm)", min_value=0.0)
            weight_kg = st.number_input("Weight (kg)", min_value=0.0); density_gcc = st.number_input("Density (g/cc)", min_value=0.0)
            speed_m_min = st.number_input("Speed (m/min)", min_value=0.0); no_of_passes = st.number_input("No. of passes", min_value=0)
            irradiation_side = st.selectbox("Irradiation side", [1,2]); batch_time_min = st.number_input("Batch time (min)", min_value=0.0)
            processed_boxes = st.number_input("No. of processed boxes", min_value=0); batches = st.number_input("No. of batches", min_value=0)
            productive_hrs = st.number_input("Productive Hrs", min_value=0.0); productive_hrs_gs = st.number_input("Productive Hrs - GS", min_value=0.0)
            total_productive_hrs = st.number_input("Total Productive Hrs", min_value=0.0); throughput_per_hr = st.number_input("Throughput / Hr", min_value=0.0)
            no_of_batches = st.number_input("No. of batches (repeat)", min_value=0); notes = st.text_area("Notes (optional)")
            if st.form_submit_button("Save daily record"):
                rec = {"date": p_date.isoformat(), "year": int(p_year), "month": int(p_month),
                       "customer_name": customer_name, "product_name": product_name, "mode": mode,
                       "length_mm": length_mm, "width_mm": width_mm, "height_mm": height_mm,
                       "weight_kg": weight_kg, "density_gcc": density_gcc, "speed_m_min": speed_m_min,
                       "no_of_passes": int(no_of_passes), "irradiation_side": int(irradiation_side),
                       "batch_time_min": float(batch_time_min), "processed_boxes": int(processed_boxes),
                       "batches": int(batches), "productive_hrs": float(productive_hrs), "productive_hrs_gs": float(productive_hrs_gs),
                       "total_productive_hrs": float(total_productive_hrs), "throughput_per_hr": float(throughput_per_hr),
                       "no_of_batches": int(no_of_batches), "notes": notes}
                con = get_conn(); cur = con.cursor()
                cur.execute("""INSERT INTO process_analysis (date,year,month,customer_name,product_name,mode,length_mm,width_mm,height_mm,weight_kg,density_gcc,speed_m_min,no_of_passes,irradiation_side,batch_time_min,processed_boxes,batches,productive_hrs,productive_hrs_gs,total_productive_hrs,throughput_per_hr,no_of_batches,notes,created_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                            (rec["date"],rec["year"],rec["month"],rec["customer_name"],rec["product_name"],rec["mode"],rec["length_mm"],rec["width_mm"],rec["height_mm"],rec["weight_kg"],rec["density_gcc"],rec["speed_m_min"],rec["no_of_passes"],rec["irradiation_side"],rec["batch_time_min"],rec["processed_boxes"],rec["batches"],rec["productive_hrs"],rec["productive_hrs_gs"],rec["total_productive_hrs"],rec["throughput_per_hr"],rec["no_of_batches"],rec["notes"],datetime.utcnow().isoformat()))
                con.commit(); con.close(); st.success("Saved daily process record")
        show_table_with_delete("process_analysis")

    # Procurement
    if section == "Procurement (monthly)":
        st.subheader("Procurement Entry (monthly)")
        with st.form("procure_form"):
            year = st.number_input("Year", value=today.year)
            selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
            month = MONTH_NAMES.index(selected) + 1
            item_desc = st.text_input("Item Description")
            amount = st.number_input("Amount (₹)", min_value=0.0)
            date_iso = st.date_input("Date").isoformat()
            notes = st.text_area("Notes (optional)")
            if st.form_submit_button("Save Procurement"):
                con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO procurement (year,month,item_desc,amount,date,notes,created_at) VALUES (?,?,?,?,?,?,?)", (int(year), int(month), item_desc, float(amount), date_iso, notes, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("Procurement Record Saved!")
        show_table_with_delete("procurement")

    # Maintenance
    if section == "Maintenance (daily)":
        st.subheader("Maintenance Entry (daily)")
        with st.form("maint_form"):
            d_date = st.date_input("Date").isoformat()
            year = st.number_input("Year", value=today.year)
            selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
            month = MONTH_NAMES.index(selected) + 1
            category = st.selectbox("Category", ["Electrical","Mechanical","Vacuum","Other"])
            amount = st.number_input("Amount (₹)", min_value=0.0)
            remarks = st.text_area("Remarks / Description")
            if st.form_submit_button("Save Maintenance"):
                con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO maintenance (date,year,month,category,amount,remarks,created_at) VALUES (?,?,?,?,?,?,?)", (d_date, int(year), int(month), category, float(amount), remarks, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("Maintenance Record Saved!")
        show_table_with_delete("maintenance")

    # Dosimeter
    if section == "Dosimeter (monthly)":
        st.subheader("Dosimeter Entry (monthly)")
        with st.form("dos_form"):
            year = st.number_input("Year", value=today.year)
            selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
            month = MONTH_NAMES.index(selected) + 1
            num_used = st.number_input("No. of Dosimeters Used", min_value=0)
            batch_type = st.text_input("Dosimeter Batch / Type")
            notes = st.text_area("Notes (optional)")
            if st.form_submit_button("Save Dosimeter"):
                con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO dosimeter (year,month,num_used,batch_type,notes,created_at) VALUES (?,?,?,?,?,?)", (int(year), int(month), int(num_used), batch_type, notes, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("Dosimeter Record Saved!")
        show_table_with_delete("dosimeter")

    # RRCAT
        if section == "RRCAT (monthly)":
            st.subheader("RRCAT Entry (monthly)")
            with st.form("rrcat_form"):
                year = st.number_input("Year", value=today.year)
                selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
                month = MONTH_NAMES.index(selected) + 1
                item_desc = st.text_input("Item Description")
                amount = st.number_input("Amount (₹)", min_value=0.0)
                date_iso = st.date_input("Date").isoformat()
                notes = st.text_area("Notes (optional)")
                if st.form_submit_button("Save RRCAT"):
                    con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO rrcat (year,month,item_desc,amount,date,notes,created_at) VALUES (?,?,?,?,?,?,?)", (int(year), int(month), item_desc, float(amount), date_iso, notes, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("RRCAT Record Saved!")
        show_table_with_delete("rrcat")

    # Infra
    if section == "Building / Infrastructure (monthly)":
        st.subheader("Building / Infrastructure Entry (monthly)")
        with st.form("infra_form"):
            year = st.number_input("Year", value=today.year)
            selected = st.selectbox("Month", MONTH_NAMES, index=today.month-1)
            month = MONTH_NAMES.index(selected) + 1
            work_desc = st.text_input("Work Description")
            vendor = st.text_input("Vendor")
            amount = st.number_input("Amount (₹)", min_value=0.0)
            date_iso = st.date_input("Date").isoformat()
            notes = st.text_area("Notes (optional)")
            if st.form_submit_button("Save Infra"):
                con = get_conn(); cur = con.cursor(); cur.execute("INSERT INTO infra (year,month,work_desc,vendor,amount,date,notes,created_at) VALUES (?,?,?,?,?,?,?,?)", (int(year), int(month), work_desc, vendor, float(amount), date_iso, notes, datetime.utcnow().isoformat())); con.commit(); con.close(); st.success("Infrastructure Record Saved!")
        show_table_with_delete("infra")

# -----------------------
# Visualization (basic) and Admin (full)
# -----------------------
if menu == "Visualization":
    st.header("Visualization & Reports")
    bes = fetch_df_full("bescom"); proc = fetch_df_full("process_analysis")
    if not proc.empty:
        p = proc.copy(); p["date"]=pd.to_datetime(p["date"], errors="coerce"); p["year"]=p["date"].dt.year; p["month"]=p["date"].dt.month
        monthly = p.groupby(["year","month"], as_index=False).agg({"processed_boxes":"sum","batches":"sum","productive_hrs":"sum"})
        monthly["month_name"] = monthly["month"].apply(lambda m: MONTH_NAMES[m-1])
        monthly["year_month"] = monthly["year"].astype(str) + "-" + monthly["month"].apply(lambda m: f"{m:02d}")
        st.subheader("Process — Monthly Aggregates")
        st.dataframe(monthly, use_container_width=True)
        st.download_button("Export Process Monthly", data=to_excel_bytes(monthly), file_name="process_monthly.xlsx")
        try:
            st.altair_chart(alt.Chart(monthly).mark_line(point=True).encode(x="year_month:N", y="processed_boxes:Q"), use_container_width=True)
        except Exception:
            pass
    else:
        st.info("No process data to visualize")

    if not bes.empty:
        st.subheader("BESCOM — Monthly Aggregates")
        btmp = bes.copy()
        if "date" in btmp.columns:
            btmp["date"]=pd.to_datetime(btmp["date"], errors="coerce")
            if "month" in btmp.columns:
                btmp["month_name"] = btmp["month"].apply(lambda m: MONTH_NAMES[m-1])
        st.dataframe(btmp, use_container_width=True)
        st.download_button("Export BESCOM", data=to_excel_bytes(btmp), file_name="bescom_all.xlsx")

# -----------------------
# ADMIN
# -----------------------
if menu == "Admin":
    st.header("Admin — User Management (Create / Remove / Change Password)")
    if not is_admin_user(st.session_state.username):
        st.error("Only admin users can access this area. Login as admin to manage users.")
    else:
        # Create user
        st.subheader("Create new user")
        with st.form("admin_create_user"):
            new_username = st.text_input("Username")
            new_password = st.text_input("Password", type="password")
            make_admin = st.checkbox("Make user an admin")
            submitted = st.form_submit_button("Create user")
            if submitted:
                if not new_username or not new_password:
                    st.error("Provide username and password")
                else:
                    con = get_conn(); cur = con.cursor()
                    try:
                        cur.execute("INSERT INTO users (username, pass_hash, is_admin, created_at) VALUES (?,?,?,?)",
                                    (new_username.strip(), hash_pw(new_password.strip()), 1 if make_admin else 0, datetime.utcnow().isoformat()))
                        con.commit(); st.success(f"User '{new_username}' created.")
                    except Exception as e:
                        st.error(f"Could not create user: {e}")
                    finally:
                        con.close()

        # Change password
        st.subheader("Change user password")
        with st.form("admin_change_pw"):
            ch_user = st.text_input("Username to change password")
            ch_pw = st.text_input("New password", type="password")
            do_change = st.form_submit_button("Change password")
            if do_change:
                if not ch_user or not ch_pw:
                    st.error("Provide username and new password")
                else:
                    con = get_conn(); cur = con.cursor()
                    cur.execute("SELECT id FROM users WHERE username = ?", (ch_user.strip(),))
                    if cur.fetchone() is None:
                        st.error("User not found")
                    else:
                        cur.execute("UPDATE users SET pass_hash = ? WHERE username = ?", (hash_pw(ch_pw.strip()), ch_user.strip()))
                        con.commit(); st.success("Password changed.")
                    con.close()

        # Delete user
        st.subheader("Delete user")
        with st.form("admin_delete_user"):
            del_user = st.text_input("Username to delete")
            confirm = st.checkbox("I confirm deletion of this user")
            do_del = st.form_submit_button("Delete user")
            if do_del:
                if not del_user or not confirm:
                    st.error("Provide username and confirm deletion")
                else:
                    con = get_conn(); cur = con.cursor()
                    # prevent deleting last admin
                    cur.execute("SELECT is_admin FROM users WHERE username = ?", (del_user.strip(),))
                    row = cur.fetchone()
                    if row and row[0] == 1:
                        cur.execute("SELECT COUNT(*) FROM users WHERE is_admin = 1"); admin_count = cur.fetchone()[0]
                        if admin_count <= 1:
                            st.error("Cannot delete the last admin user.")
                        else:
                            cur.execute("DELETE FROM users WHERE username = ?", (del_user.strip(),)); con.commit(); st.success("User deleted.")
                    else:
                        cur.execute("DELETE FROM users WHERE username = ?", (del_user.strip(),)); con.commit(); st.success("User deleted.")
                    con.close()

        # List users & show basic info (no password shown, only hashed preview)
        st.subheader("Existing users")
        users_df = fetch_df_full("users")
        if users_df.empty:
            st.info("No users found")
        else:
            # show id, username, is_admin, created_at
            display = users_df[["id","username","is_admin","created_at"]].copy()
            st.dataframe(display, use_container_width=True)
            # show hashed preview (first 8 chars) if needed
            if st.checkbox("Show password hash preview (first 8 chars)"):
                users_df["hash_preview"] = users_df["pass_hash"].str[:8]
                st.dataframe(users_df[["username","hash_preview"]], use_container_width=True)

        # Admin actions: reset password to default
        st.subheader("Admin utilities")
        with st.form("admin_utils"):
            reset_user = st.text_input("Reset password for username (optional)")
            make_admin_user = st.text_input("Make user admin (username) (optional)")
            do_utils = st.form_submit_button("Apply utilities")
            if do_utils:
                con = get_conn(); cur = con.cursor()
                if reset_user:
                    cur.execute("UPDATE users SET pass_hash = ? WHERE username = ?", (hash_pw(DEFAULT_ADMIN_PASS), reset_user.strip())); st.success(f"Password for {reset_user} reset to default")
                if make_admin_user:
                    cur.execute("UPDATE users SET is_admin = 1 WHERE username = ?", (make_admin_user.strip(),)); st.success(f"{make_admin_user} is now admin")
                con.commit(); con.close()

st.markdown("---")
st.caption("IARPF – MICROTROL — Rebuilt clean app")
