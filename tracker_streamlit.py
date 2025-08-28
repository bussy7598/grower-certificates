import streamlit as st
import pandas as pd
import difflib
from datetime import datetime, timedelta
import io

# =========================
# Settings
# =========================
EXPIRY_WARNING_DAYS = 60  # threshold for "Expiring Soon"

EXPECTED = {
    "Supplier": "Supplier",
    "Certification Body": "Certification Body",
    "Certificate": "Certificate",
    "Expiry Date": "Expiry Date",
}

# =========================
# Helpers
# =========================
def load_and_map_certificates(file):
    """Load and clean certificates from an Excel file handle (uploaded)."""
    try:
        df = pd.read_excel(file)
    except Exception:
        # some exports put headers on row 3
        df = pd.read_excel(file, header=2)

    # tidy cols & drop Unnamed
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", case=False)]

    # fuzzy match & rename to canonical names
    col_map = {}
    for expected_col in EXPECTED.keys():
        match = difflib.get_close_matches(expected_col, df.columns, n=1, cutoff=0.5)
        if match:
            col_map[match[0]] = EXPECTED[expected_col]
    df = df.rename(columns=col_map)

    # ensure required columns
    for col in EXPECTED.values():
        if col not in df.columns:
            df[col] = ""

    # parse dates & compute status
    df["Expiry Date"] = pd.to_datetime(df["Expiry Date"], errors="coerce")
    today = pd.Timestamp.today()
    df["Days_Until_Expiry"] = (df["Expiry Date"] - today).dt.days

    def status_from_days(days):
        if pd.isna(days):
            return "Unknown"
        if days < 0:
            return "Expired"
        if days <= EXPIRY_WARNING_DAYS:
            return "Expiring Soon"
        return "Valid"

    df["Status"] = df["Days_Until_Expiry"].apply(status_from_days)
    return df


def init_contact_log(df_in=None):
    """Ensure contact log has the right columns & types."""
    cols = ["Date", "Supplier", "Action", "Notes"]
    if df_in is None:
        log = pd.DataFrame(columns=cols)
    else:
        log = df_in.copy()
        for c in cols:
            if c not in log.columns:
                log[c] = ""
        log = log[cols]
    # parse/normalize Date for sorting
    log["Date"] = pd.to_datetime(log["Date"], errors="coerce")
    return log


def df_to_excel_bytes(df):
    """Return a BytesIO excel blob from a DataFrame."""
    output = io.BytesIO()
    # if df is styled, get the underlying data
    if hasattr(df, "data"):
        df = df.data
    df.to_excel(output, index=False)
    output.seek(0)
    return output


def style_row(row):
    """Softer row colours by Status."""
    status = str(row["Status"]).strip()
    if status == "Expired":
        return ["background-color: #f8d7da"] * len(row)   # light red
    if status == "Expiring Soon":
        return ["background-color: #fff3cd"] * len(row)   # light yellow
    if status == "Valid":
        return ["background-color: #d4edda"] * len(row)   # light green
    return [""] * len(row)


# =========================
# App UI
# =========================
st.set_page_config(page_title="Certification Tracker", layout="wide")
st.title("Certification Tracker — Interactive (Local Uploads)")

# ---- Uploads
st.sidebar.markdown("### Upload Files")
cert_file = st.sidebar.file_uploader("Upload Grower Certifications Excel", type=["xlsx"])
contact_file = st.sidebar.file_uploader("Upload Contact Log (xlsx/csv, optional)", type=["xlsx", "csv"])

if not cert_file:
    st.warning("Please upload a certifications Excel file to continue.")
    st.stop()

# ---- Load data
df = load_and_map_certificates(cert_file)

# contact log: load once into session_state so new entries persist
if "contact_log" not in st.session_state:
    if contact_file is not None:
        try:
            tmp = pd.read_excel(contact_file)
        except Exception:
            tmp = pd.read_csv(contact_file)
        st.session_state.contact_log = init_contact_log(tmp)
    else:
        st.session_state.contact_log = init_contact_log()

contact_log = st.session_state.contact_log

# ---- Grower selector & filters
growers = sorted(df["Supplier"].dropna().unique().tolist())
selected = st.selectbox("Select grower / supplier", ["(All growers)"] + growers)

col1, col2, col3 = st.columns([2, 1, 1])
with col1:
    search = st.text_input("Quick search supplier (contains):")
with col2:
    show_only = st.selectbox("Show only", ["All", "Valid", "Expiring Soon", "Expired", "Unknown"])
with col3:
    refresh = st.button("Recompute statuses")

if refresh:
    # recompute status (useful if you changed EXPIRY_WARNING_DAYS)
    df = load_and_map_certificates(cert_file)

# ---- Apply filters
filtered = df.copy()
if selected != "(All growers)":
    filtered = filtered[filtered["Supplier"] == selected]
if search:
    filtered = filtered[filtered["Supplier"].str.contains(search, case=False, na=False)]
if show_only != "All":
    filtered = filtered[filtered["Status"] == show_only]

# ---- Metrics
total = len(filtered)
valid_count = (filtered["Status"] == "Valid").sum()
expiring_count = (filtered["Status"] == "Expiring Soon").sum()
expired_count  = (filtered["Status"] == "Expired").sum()
unknown_count  = (filtered["Status"] == "Unknown").sum()
valid_pct = (valid_count / total * 100) if total > 0 else 0.0

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Total Certificates", total)
m2.metric("Valid", f"{valid_count}", f"{valid_pct:.1f}% valid")
m3.metric("Expiring Soon", expiring_count)
m4.metric("Expired", expired_count)
m5.metric("Unknown", unknown_count)

# ---- Certificates table
st.subheader("Certificates")
display_cols = ["Supplier", "Certification Body", "Certificate", "Expiry Date", "Days_Until_Expiry", "Status"]
display_df = filtered[display_cols].copy()
display_df["Expiry Date"] = display_df["Expiry Date"].dt.strftime("%Y-%m-%d")
styled = display_df.style.apply(style_row, axis=1)
st.dataframe(styled, use_container_width=True)

# ---- Export buttons
st.subheader("Export Data")
colA, colB = st.columns(2)
with colA:
    st.download_button(
        "Download Certificates Excel",
        df_to_excel_bytes(display_df),
        file_name=(f"{selected}_certs.xlsx" if selected != "(All growers)" else "all_certs.xlsx"),
    )
with colB:
    # filter contact log for selected grower
    to_log = contact_log if selected == "(All growers)" else contact_log[contact_log["Supplier"] == selected]
    # show newest first
    to_log = to_log.sort_values("Date", ascending=False)
    st.download_button(
        "Download Contact Log Excel",
        df_to_excel_bytes(to_log),
        file_name=(f"{selected}_contact_log.xlsx" if selected != "(All growers)" else "all_contact_log.xlsx"),
    )

# ---- Contact log table
st.subheader("Contact log")
to_show = contact_log if selected == "(All growers)" else contact_log[contact_log["Supplier"] == selected]
to_show = to_show.sort_values("Date", ascending=False)
st.dataframe(to_show, use_container_width=True)

# ---- Add new contact entry
st.markdown("### Log new contact")
with st.form("contact_form", clear_on_submit=True):
    c_action = st.selectbox("Action", ["Email", "Call", "Meeting", "Other"])
    c_date = st.date_input("Date", value=datetime.today())
    c_notes = st.text_area("Notes", height=120)
    submitted = st.form_submit_button("Save contact")
    if submitted:
        if selected == "(All growers)":
            st.error("Please select a specific grower to log a contact.")
        else:
            new_row = {
                "Date": pd.to_datetime(c_date),
                "Supplier": selected,
                "Action": c_action,
                "Notes": c_notes,
            }
            st.session_state.contact_log = pd.concat(
                [st.session_state.contact_log, pd.DataFrame([new_row])], ignore_index=True
            )
            st.success("Contact saved in session. Use 'Download Contact Log Excel' to save to disk.")
            # refresh on-screen table
            contact_log = st.session_state.contact_log
            to_show = contact_log if selected == "(All growers)" else contact_log[contact_log["Supplier"] == selected]
            st.dataframe(to_show.sort_values("Date", ascending=False), use_container_width=True)

st.markdown("---")
st.caption(
    "Uploads are in-memory for this session. New contacts are stored in session_state and included in downloads. "
    "For shared/multi-user storage, point reads/writes to a shared Excel/CSV/SQLite on OneDrive/SharePoint."
)
