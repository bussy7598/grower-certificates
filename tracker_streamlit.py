import streamlit as st
import pandas as pd
import difflib
from datetime import datetime, timedelta
import os

# ===== SETTINGS =====
CERT_FILE = r"C:\Users\SeanBuss\Project_Cert\Grower Certifications.xlsx"
CONTACT_FILE = r"C:\Users\SeanBuss\Project_Cert\Contact Log.xlsx"   # will be created if missing
EXPIRY_WARNING_DAYS = 60  # threshold for "Expiring Soon"

# ===== Helpers for robust column mapping =====
EXPECTED = {
    "Supplier": "Supplier",
    "Certification Body": "Certification Body",
    "Certificate": "Certificate",
    "Expiry Date": "Expiry Date"
}

def load_and_map_certificates(path):
    # try reading common excel forms; if header row is not row0, we attempt header=2 as fallback
    try:
        df = pd.read_excel(path)
    except Exception:
        df = pd.read_excel(path, header=2)

    # tidy column names
    df.columns = [str(c).strip() for c in df.columns]
    # drop unnamed helper cols
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False)]

    # fuzzy match and rename
    col_map = {}
    for expected_col in EXPECTED.keys():
        match = difflib.get_close_matches(expected_col, df.columns, n=1, cutoff=0.5)
        if match:
            col_map[match[0]] = EXPECTED[expected_col]
    df = df.rename(columns=col_map)

    # ensure our canonical columns exist
    for canonical in EXPECTED.values():
        if canonical not in df.columns:
            df[canonical] = ""

    # parse expiry date column robustly
    df['Expiry Date'] = pd.to_datetime(df['Expiry Date'], errors='coerce')

    # compute derived fields
    today = pd.Timestamp.today()
    df['Days_Until_Expiry'] = (df['Expiry Date'] - today).dt.days
    def status_from_days(days):
        if pd.isna(days):
            return "Unknown"
        if days < 0:
            return "Expired"
        if days <= EXPIRY_WARNING_DAYS:
            return "Expiring Soon"
        return "Valid"
    df['Status'] = df['Days_Until_Expiry'].apply(status_from_days)

    # keep useful columns and return
    return df

def load_contact_log(path):
    cols = ["Date", "Supplier", "Action", "Notes"]
    if os.path.exists(path):
        try:
            log = pd.read_excel(path)
        except Exception:
            log = pd.read_csv(path)
        for col in cols:
            if col not in log.columns:
                log[col] = ""
    else:
        log = pd.DataFrame(columns=["Date","Supplier","Action","Notes"])
    return log

def save_contact_log(df, path):
    # prefer Excel for readability; fallback to csv if write fails
    try:
        df.to_excel(path, index=False)
    except Exception:
        csv_path = os.path.splitext(path)[0] + ".csv"
        df.to_csv(csv_path, index=False)

def style_row(row):
    """Colour Rows By Status"""
    status = row['Status']
    if status == "Expired":
        return ['background-color: #ff0000'] * len(row)
    if status == "Expiring Soon":
        return ['background-color: #ff9966'] * len(row)
    if status == "Valid":
        return ['background-color: #000000'] * len(row)
    return [''] * len(row)

# ===== App UI =====
st.set_page_config(page_title="Certification Tracker", layout="wide")
st.title("Certification Tracker — Interactive")

# Load data
st.sidebar.markdown("### Configuration")
st.sidebar.write(f"Cert file: `{CERT_FILE}`")
st.sidebar.write(f"Contact file: `{CONTACT_FILE}`")

df = load_and_map_certificates(CERT_FILE)
contact_log = load_contact_log(CONTACT_FILE)

# Grower selector
growers = sorted(df['Supplier'].dropna().unique().tolist())
selected = st.selectbox("Select grower / supplier", ["(All growers)"] + growers)

# Filters row
col1, col2, col3 = st.columns([2,1,1])
with col1:
    search = st.text_input("Quick search supplier (contains):")
with col2:
    show_only = st.selectbox("Show only", ["All", "Valid", "Expiring Soon", "Expired", "Unknown"])
with col3:
    refresh = st.button("Refresh data")

if refresh:
    df = load_and_map_certificates(CERT_FILE)
    st.experimental_rerun()

# Filter dataframe
filtered = df.copy()
if selected != "(All growers)":
    filtered = filtered[filtered['Supplier'] == selected]
if search:
    filtered = filtered[filtered['Supplier'].str.contains(search, case=False, na=False)]
if show_only != "All":
    filtered = filtered[filtered['Status'] == show_only]

# Summary metrics
total = len(filtered)
valid_count = (filtered['Status'] == "Valid").sum()
expiring_count = (filtered['Status'] == "Expiring Soon").sum()
expired_count = (filtered['Status'] == "Expired").sum()
unknown_count = (filtered['Status'] == "Unknown").sum()
valid_pct = (valid_count / total * 100) if total > 0 else 0.0

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total Certificates", total)
m2.metric("Valid", f"{valid_count}", f"{valid_pct:.1f}% valid")
m3.metric("Expiring Soon", expiring_count)
m4.metric("Expired", expired_count)
m5.metric("Unknown", unknown_count)

st.subheader("Certificates")
display_cols = ['Supplier', 'Certification Body','Certificate','Expiry Date','Days_Until_Expiry','Status']
display_df = filtered[display_cols].copy()
# format expiry for nicer view
display_df['Expiry Date'] = display_df['Expiry Date'].dt.strftime('%Y-%m-%d')
st.dataframe(display_df.style.apply(style_row, axis=1), use_container_width=True)

# Export buttons
colA, colB = st.columns([1,1])
with colA:
    if st.button("Export certificates (selected) to Excel"):
        out_path = os.path.join(os.getcwd(), f"{selected}_certs.xlsx" if selected!="(All growers)" else "all_certs.xlsx")
        filtered.to_excel(out_path, index=False)
        st.success(f"Exported to `{out_path}`")
with colB:
    if st.button("Export contact log (selected) to Excel"):
        if selected == "(All growers)":
            out_path = os.path.join(os.getcwd(), "all_contact_log.xlsx")
            contact_log.to_excel(out_path, index=False)
        else:
            out_path = os.path.join(os.getcwd(), f"{selected}_contact_log.xlsx")
            contact_log[contact_log['Supplier'] == selected].to_excel(out_path, index=False)
        st.success(f"Exported to `{out_path}`")

# Contact log section
st.subheader("Contact log")
if selected == "(All growers)":
    st.write("Showing all contact entries")
    if "Date" in contact_log.columns:
        st.dataframe(contact_log.sort_values("Date", ascending=False), use_container_width=True)
    else:
        st.dataframe(contact_log, use_container_width=True)
else:
    entries = contact_log[contact_log['Supplier'] == selected]
    if "Date" in entries.columns:
        entries = entries.sort_values("Date", ascending=False)
    st.dataframe(entries, use_container_width=True)

# Add new contact entry (form)
st.markdown("### Log new contact")
with st.form("contact_form", clear_on_submit=True):
    c_action = st.selectbox("Action", ["Email","Call","Meeting","Other"])
    c_date = st.date_input("Date", value=datetime.today())
    c_notes = st.text_area("Notes", height=120)
    submitted = st.form_submit_button("Save contact")
    if submitted:
        if selected == "(All growers)":
            st.error("Please select a specific grower to log a contact.")
        else:
            new_row = {
                "Date": pd.to_datetime(c_date).strftime("%Y-%m-%d"),
                "Supplier": selected,
                "Action": c_action,
                "Notes": c_notes
            }
            contact_log = pd.concat([contact_log, pd.DataFrame([new_row])], ignore_index=True)
            save_contact_log(contact_log, CONTACT_FILE)
            st.success("Contact saved")

# Small footer
st.markdown("---")
st.caption("Local app — Contact log is saved to the CONTACT_FILE on disk. For multi-user concurrency use a shared DB or spreadsheet host (SharePoint/OneDrive).")
