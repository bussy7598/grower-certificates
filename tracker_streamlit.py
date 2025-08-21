import streamlit as st
import pandas as pd
import difflib
from datetime import datetime, timedelta
import io

# ===== SETTINGS =====
EXPIRY_WARNING_DAYS = 60  # threshold for "Expiring Soon"

# ===== Helpers for robust column mapping =====
EXPECTED = {
    "Supplier": "Supplier",
    "Certification Body": "Certification Body",
    "Certificate": "Certificate",
    "Expiry Date": "Expiry Date"
}

def load_and_map_certificates(file):
    """
    Load and clean certificates from an Excel file.
    'file' can be a file-like object (Streamlit upload) or path
    """
    try:
        df = pd.read_excel(file)
    except Exception:
        df = pd.read_excel(file, header=2)

    # tidy column names
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains('^Unnamed', case=False)]

    # fuzzy match and rename
    col_map = {}
    for expected_col in EXPECTED.keys():
        match = difflib.get_close_matches(expected_col, df.columns, n=1, cutoff=0.5)
        if match:
            col_map[match[0]] = EXPECTED[expected_col]
    df = df.rename(columns=col_map)

    for canonical in EXPECTED.values():
        if canonical not in df.columns:
            df[canonical] = ""

    # parse expiry date
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
    return df

def load_contact_log(file=None):
    """
    Load contact log from an uploaded file or create empty DataFrame
    """
    cols = ["Date", "Supplier", "Action", "Notes"]
    if file:
        try:
            df = pd.read_excel(file)
        except Exception:
            df = pd.read_csv(file)
        for col in cols:
            if col not in df.columns:
                df[col] = ""
    else:
        df = pd.DataFrame(columns=cols)
    return df

def save_contact_log(df):
    """
    Returns a BytesIO object of Excel file for download
    """
    output = io.BytesIO()
    df.to_excel(output, index=False)
    output.seek(0)
    return output

def style_row(row):
    """Colour Rows By Status"""
    status = row['Status']
    if status == "Expired":
        return ['background-color: #ff0000'] * len(row)
    if status == "Expiring Soon":
        return ['background-color: #ff9966'] * len(row)
    if status == "Valid":
        return ['background-color: #00cc00'] * len(row)
    return [''] * len(row)

# ===== App UI =====
st.set_page_config(page_title="Certification Tracker", layout="wide")
st.title("Certification Tracker — Interactive")

st.sidebar.markdown("### Upload Files")
cert_file = st.sidebar.file_uploader("Upload Grower Certifications Excel", type=["xlsx"])
contact_file = st.sidebar.file_uploader("Upload Contact Log Excel (optional)", type=["xlsx","csv"])

if cert_file:
    df = load_and_map_certificates(cert_file)
else:
    st.warning("Please upload a certifications Excel file to continue.")
    st.stop()

contact_log = load_contact_log(contact_file)

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

if refresh and cert_file:
    df = load_and_map_certificates(cert_file)
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
display_df['Expiry Date'] = display_df['Expiry Date'].dt.strftime('%Y-%m-%d')
st.dataframe(display_df.style.apply(style_row, axis=1), use_container_width=True)

# Export buttons
colA, colB = st.columns([1,1])
with colA:
    if st.button("Export certificates (selected) to Excel"):
        out_file = save_contact_log(filtered)
        st.download_button("Download Certificates Excel", out_file, file_name=f"{selected}_certs.xlsx" if selected!="(All growers)" else "all_certs.xlsx")
with colB:
    if st.button("Export contact log (selected) to Excel"):
        if selected == "(All growers)":
            out_file = save_contact_log(contact_log)
        else:
            out_file = save_contact_log(contact_log[contact_log['Supplier'] == selected])
        st.download_button("Download Contact Log Excel", out_file, file_name=f"{selected}_contact_log.xlsx" if selected!="(All growers)" else "all_contact_log.xlsx")

# Contact log section
st.subheader("Contact log")
if selected == "(All growers)":
    st.write("Showing all contact entries")
    st.dataframe(contact_log.sort_values("Date", ascending=False), use_container_width=True)
else:
    entries = contact_log[contact_log['Supplier'] == selected]
    entries = entries.sort_values("Date", ascending=False) if "Date" in entries.columns else entries
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
            st.success("Contact saved (in memory, download updated log below)")
            # offer updated download
            st.download_button("Download Updated Contact Log", save_contact_log(contact_log), file_name="updated_contact_log.xlsx")

# Small footer
st.markdown("---")
st.caption("Streamlit Cloud app — Upload files via sidebar. Downloads are provided for certificates and contact log. Multi-user concurrency requires shared storage or database.")
