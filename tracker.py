import pandas as pd
import PySimpleGUI as sg
from datetime import datetime, timedelta
import os
import difflib

#Settings
CERT_FILE = r"C:\Users\SeanBuss\Project_Cert\Grower Certifications.xlsx"
CONTACT_FILE = r"C:\Users\SeanBuss\Project_Cert\Contact Log.xlsx"
EXPIRY_WARNING_DAYS = 60

#Expected_Columns
EXPECTED_COLUMNS = ["Supplier", "Certification Body", "Certificate No", "Expiry Date"]

#Load Certifications
df = pd.read_excel(CERT_FILE, header=2)
df.columns = [c.strip() for c in df.columns]
print("Columns found in Excel file:")
print(df.columns.tolist())

#Map Similar Columns and add missing expected columns if not found

col_map = {}
for col in EXPECTED_COLUMNS:
    match = difflib.get_close_matches(col, df.columns, n=1, cutoff=0.6)
    if match:
        col_map[match[0]] = col
df = df.rename(columns=col_map)

for col in EXPECTED_COLUMNS:
    if col not in df.columns:
        df[col] = ""

#Function to Check Status
def get_status(expiry_str):
    try:
        expiry = pd.to_datetime(expiry_str)
        if expiry < datetime.now():
            return "Expired"
        elif expiry < datetime.now() + timedelta(days=EXPIRY_WARNING_DAYS):
            return "Expiring Soon"
        else:
            return "Valid"
    except:
        return "Unknown"

df["Status"] = df["Expiry Date"].apply(get_status)

#Load or create contact log
if os.path.exists(CONTACT_FILE):
    log_df = pd.read_excel(CONTACT_FILE)
else:
    log_df = pd.DataFrame(columns=["Date", "Supplier","Action","Notes"])


#GUI Layout

sg.theme("SystemDefault")

layout = [
    [sg.Text("Search Supplier:"), sg.Input(key="-SEARCH-", enable_events=True)],
    [sg.Table(
        values=df[["Supplier", "Certification Body", "Certificate No","Expiry Date", "Status"]].values.tolist(),
        headings=["Supplier", "Cert Body", "Cert No", "Expiry Date", "Status"],
        key="-TABLE-",
        enable_events=True,
        auto_size_columns=True,
        justification="Left",
        num_rows=15
    )],
    [sg.Button("Log Contact"), sg.Button("Exit")]
]

window = sg.Window("Certification Tracker", layout)

#Event Loop

while True:
    event, values = window.read()
    if event in (sg.WIN_CLOSED, "Exit"):
        break

    #Search Filter
    if event == "-SEARCH-":
        search_text = values["-SEARCH-"].lower()
        filtered = df[df["Supplier"].str.lower().str.contains(search_text)]
        window["-TABLE-"].update(filtered[["Supplier","Certification Body", "Certification No","Expiry Date", "Status"]].values.tolist())

    #Log Contact
    if event == "Log Contact":
        selected_rows = values["-TABLE-"]
        if not selected_rows:
            sg.popup("Please select a supplier from the table.")
            continue

        supplier = df.iloc[selected_rows[0]]["Supplier"]
        layout_log = [
            [sg.Text(f"Logging contact for: {supplier}")],
            [sg.Text("Action:"), sg.Combo(["Email", "Call", "Meeting"], key="-ACTION-")],
            [sg.Text("Notes:"),sg.Multiline(size=(40,5), key="-NOTES-")],
            [sg.Button("Save"), sg.Button("Cancel")]
        ]
        log_win = sg.Window("Log Contact", layout_log)
        while True:
            ev, vals = log_win.read()
            if ev in (sg.WIN_CLOSED, "Cancel"):
                break
            if ev == "Save":
                new_entry = {
                    "Date": datetime.now().strftime("%Y-%m-%d"),
                    "Supplier": supplier,
                    "Action": vals["-ACTION-"],
                    "Notes": vals["-NOTES-"]
                }
                log_df = pd.concat([log_df, pd.DataFrame([new_entry])], ignore_index=True)
                log_df.to_excel(CONTACT_FILE, index=False)
                sg.popup("Contact Saved")
                break
            log_win.close()

window.close()
