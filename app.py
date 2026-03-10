import streamlit as st
import pandas as pd
import urllib.parse
import os

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="DHL Operations – Email Dashboard",
    page_icon="📦",
    layout="wide",
)

# ── Constants ─────────────────────────────────────────────────────────────────
RECIPIENTS_CSV = os.path.join(os.path.dirname(__file__), "data", "recipients.csv")

COMMON_COLS = ["dhlParcelId", "Customer Name", "Date", "Remarks"]

SHEETS = [
    {
        "name": "Damage",
        "icon": "🔴",
        "color": "#C00000",
        "extra_cols": ["Damage Type", "Damage Description", "Item Value (RM)"],
        "subject": "DHL Damage Report – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "Please be informed that a damage case has been identified for the following shipment:\n\n"
            "  Parcel ID    : {dhlParcelId}\n"
            "  Customer     : {Customer Name}\n"
            "  Date          : {Date}\n"
            "  Damage Type : {Damage Type}\n"
            "  Description  : {Damage Description}\n"
            "  Item Value    : RM {Item Value (RM)}\n"
            "  Remarks       : {Remarks}\n\n"
            "Kindly investigate and advise on the next course of action.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "MPS Incomplete",
        "icon": "🟠",
        "color": "#ED7D31",
        "extra_cols": ["Total Pieces", "Pieces Received", "Missing Pieces"],
        "subject": "MPS Incomplete – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "An incomplete Multi-Piece Shipment (MPS) has been detected. Details below:\n\n"
            "  Parcel ID       : {dhlParcelId}\n"
            "  Customer        : {Customer Name}\n"
            "  Date             : {Date}\n"
            "  Total Pieces    : {Total Pieces}\n"
            "  Pieces Received : {Pieces Received}\n"
            "  Missing Pieces  : {Missing Pieces}\n"
            "  Remarks          : {Remarks}\n\n"
            "Please trace the missing pieces and update accordingly.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "Suspect Lost",
        "icon": "🟣",
        "color": "#7030A0",
        "extra_cols": ["Last Scan Location", "Last Scan Date", "Investigation Status"],
        "subject": "Suspect Lost Shipment – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "The following shipment is flagged as SUSPECT LOST and requires immediate attention:\n\n"
            "  Parcel ID            : {dhlParcelId}\n"
            "  Customer             : {Customer Name}\n"
            "  Date                  : {Date}\n"
            "  Last Scan Location  : {Last Scan Location}\n"
            "  Last Scan Date      : {Last Scan Date}\n"
            "  Investigation Status : {Investigation Status}\n"
            "  Remarks               : {Remarks}\n\n"
            "Please escalate and initiate a full trace immediately.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "Duplicate",
        "icon": "🟡",
        "color": "#F4B942",
        "extra_cols": ["Duplicate Parcel ID", "Root Cause"],
        "subject": "Duplicate Shipment Alert – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "A duplicate shipment entry has been identified. Please review:\n\n"
            "  Parcel ID           : {dhlParcelId}\n"
            "  Customer            : {Customer Name}\n"
            "  Date                 : {Date}\n"
            "  Duplicate Parcel ID : {Duplicate Parcel ID}\n"
            "  Root Cause          : {Root Cause}\n"
            "  Remarks              : {Remarks}\n\n"
            "Please verify and take corrective action to avoid billing discrepancies.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "EMY Zipcode (Reroute)",
        "icon": "🔵",
        "color": "#0070C0",
        "extra_cols": ["Original Zipcode", "Correct Zipcode", "Reroute Status"],
        "subject": "EMY Zipcode Reroute – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "The following shipment requires rerouting due to an EMY zipcode issue:\n\n"
            "  Parcel ID        : {dhlParcelId}\n"
            "  Customer         : {Customer Name}\n"
            "  Date              : {Date}\n"
            "  Original Zipcode : {Original Zipcode}\n"
            "  Correct Zipcode  : {Correct Zipcode}\n"
            "  Reroute Status   : {Reroute Status}\n"
            "  Remarks           : {Remarks}\n\n"
            "Kindly action the reroute at the earliest to avoid further delay.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "Missing DO",
        "icon": "🟤",
        "color": "#C55A11",
        "extra_cols": ["DO Number", "Origin Station", "Action Required"],
        "subject": "Missing Delivery Order – {dhlParcelId} | DO: {DO Number}",
        "body": (
            "Dear Team,\n\n"
            "A Delivery Order (DO) is missing for the following shipment:\n\n"
            "  Parcel ID       : {dhlParcelId}\n"
            "  Customer        : {Customer Name}\n"
            "  Date             : {Date}\n"
            "  DO Number       : {DO Number}\n"
            "  Origin Station  : {Origin Station}\n"
            "  Action Required : {Action Required}\n"
            "  Remarks          : {Remarks}\n\n"
            "Please provide or reissue the DO urgently to proceed with delivery.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "Cancelled Shipment",
        "icon": "⚫",
        "color": "#595959",
        "extra_cols": ["Cancellation Reason", "Cancelled By", "Refund Required"],
        "subject": "Cancelled Shipment Notice – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "The following shipment has been cancelled. Please take note:\n\n"
            "  Parcel ID            : {dhlParcelId}\n"
            "  Customer             : {Customer Name}\n"
            "  Date                  : {Date}\n"
            "  Cancellation Reason  : {Cancellation Reason}\n"
            "  Cancelled By         : {Cancelled By}\n"
            "  Refund Required      : {Refund Required}\n"
            "  Remarks               : {Remarks}\n\n"
            "Please ensure all related records are updated and refund (if applicable) is processed.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
    {
        "name": "Prealert Shipment",
        "icon": "🟢",
        "color": "#375623",
        "extra_cols": ["Origin Country", "Expected Arrival", "Prealert Reference"],
        "subject": "Prealert Notification – {dhlParcelId} | {Customer Name}",
        "body": (
            "Dear Team,\n\n"
            "A prealert has been received for the following inbound shipment:\n\n"
            "  Parcel ID           : {dhlParcelId}\n"
            "  Customer            : {Customer Name}\n"
            "  Date                 : {Date}\n"
            "  Origin Country      : {Origin Country}\n"
            "  Expected Arrival    : {Expected Arrival}\n"
            "  Prealert Reference  : {Prealert Reference}\n"
            "  Remarks              : {Remarks}\n\n"
            "Please prepare for receiving and ensure customs documentation is in order.\n\n"
            "Regards,\nDHL Operations Team"
        ),
    },
]

SHEET_MAP = {s["name"]: s for s in SHEETS}

# ── Helpers ───────────────────────────────────────────────────────────────────
@st.cache_data
def load_recipients():
    try:
        df = pd.read_csv(RECIPIENTS_CSV, dtype=str).fillna("")
        return df
    except Exception:
        return pd.DataFrame(columns=["sheet", "to", "cc", "bcc"])

def get_recipients(sheet_name):
    df = load_recipients()
    row = df[df["sheet"].str.strip() == sheet_name]
    if row.empty:
        return "", "", ""
    r = row.iloc[0]
    return str(r.get("to","")).strip(), str(r.get("cc","")).strip(), str(r.get("bcc","")).strip()

def fill_template(template, row):
    result = template
    for col, val in row.items():
        result = result.replace("{" + str(col) + "}", str(val) if pd.notna(val) else "")
    return result

def build_mailto(to, cc, bcc, subject, body):
    params = {}
    if cc:  params["cc"]  = cc
    if bcc: params["bcc"] = bcc
    params["subject"] = subject
    params["body"]    = body
    query = urllib.parse.urlencode(params, quote_via=urllib.parse.quote)
    to_encoded = urllib.parse.quote(to)
    return f"mailto:{to_encoded}?{query}"

# ── Styles ────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { background: #F4F6F9; }
    .block-container { padding-top: 1.5rem; padding-bottom: 2rem; }
    .dashboard-title {
        font-size: 1.7rem; font-weight: 700; color: #1F3864;
        margin-bottom: 0.1rem;
    }
    .dashboard-sub {
        font-size: 0.92rem; color: #666; margin-bottom: 1.5rem;
    }
    .issue-card {
        background: white;
        border-radius: 10px;
        padding: 1rem 1.2rem 0.8rem 1.2rem;
        margin-bottom: 1rem;
        box-shadow: 0 1px 4px rgba(0,0,0,0.08);
        border-left: 5px solid #ccc;
    }
    .card-title {
        font-size: 1rem; font-weight: 700; margin-bottom: 0.3rem;
    }
    .card-count {
        font-size: 0.82rem; color: #888; margin-bottom: 0.6rem;
    }
    .pill {
        display: inline-block;
        padding: 2px 10px;
        border-radius: 20px;
        font-size: 0.75rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    .stDataFrame { border-radius: 8px; }
    div[data-testid="stExpander"] { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="dashboard-title">📦 DHL Operations – Email Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="dashboard-sub">Upload your data file, review each issue type, and generate Outlook draft emails in one click.</div>', unsafe_allow_html=True)

# ── File upload ───────────────────────────────────────────────────────────────
uploaded = st.file_uploader(
    "Upload Excel (.xlsx) or CSV file",
    type=["xlsx", "csv"],
    help="File should have a sheet/column matching each issue type, with dhlParcelId, Customer Name, Date, Remarks columns."
)

if not uploaded:
    # Show empty dashboard with instructions
    st.info("👆 Upload a file above to populate the dashboard. Each sheet tab in the Excel file should match an issue type name.")

    col1, col2, col3, col4 = st.columns(4)
    for i, sh in enumerate(SHEETS):
        with [col1, col2, col3, col4][i % 4]:
            st.markdown(f"""
            <div class="issue-card" style="border-left-color:{sh['color']}">
                <div class="card-title">{sh['icon']} {sh['name']}</div>
                <div class="card-count">No data loaded</div>
                <span class="pill" style="background:{sh['color']}22;color:{sh['color']}">Waiting for upload</span>
            </div>
            """, unsafe_allow_html=True)
    st.stop()

# ── Load uploaded file ────────────────────────────────────────────────────────
@st.cache_data
def load_file(file):
    name = file.name
    if name.endswith(".csv"):
        df = pd.read_csv(file, dtype=str).fillna("")
        return {"CSV Data": df}
    else:
        xls = pd.ExcelFile(file)
        return {sheet: xls.parse(sheet, dtype=str).fillna("") for sheet in xls.sheet_names}

all_sheets_data = load_file(uploaded)

# Map uploaded sheet names to issue types (case-insensitive match)
def match_sheet(upload_names, issue_name):
    for n in upload_names:
        if n.strip().lower() == issue_name.lower():
            return n
    return None

# ── Summary row ───────────────────────────────────────────────────────────────
st.markdown("### 📊 Summary")
cols_summary = st.columns(len(SHEETS))
for i, sh in enumerate(SHEETS):
    matched = match_sheet(all_sheets_data.keys(), sh["name"])
    count = len(all_sheets_data[matched]) if matched else 0
    with cols_summary[i]:
        st.markdown(f"""
        <div style="background:white;border-radius:8px;padding:0.6rem 0.8rem;
                    border-top:4px solid {sh['color']};text-align:center;
                    box-shadow:0 1px 3px rgba(0,0,0,0.07);">
            <div style="font-size:1.4rem;font-weight:700;color:{sh['color']}">{count}</div>
            <div style="font-size:0.72rem;color:#555;font-weight:600">{sh['icon']} {sh['name']}</div>
        </div>
        """, unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ── Issue cards ───────────────────────────────────────────────────────────────
st.markdown("### ✉️ Issue Breakdown")

for sh in SHEETS:
    matched_key = match_sheet(all_sheets_data.keys(), sh["name"])
    df = all_sheets_data[matched_key].copy() if matched_key else pd.DataFrame()
    count = len(df)

    with st.expander(f"{sh['icon']}  **{sh['name']}**  —  {count} record{'s' if count != 1 else ''}", expanded=count > 0):

        to_addr, cc_addr, bcc_addr = get_recipients(sh["name"])

        # Recipient status
        if not to_addr:
            st.warning("⚠️ No **To** recipient configured. Edit `data/recipients.csv` to add one.")
        else:
            rcpt_parts = [f"**To:** {to_addr}"]
            if cc_addr:  rcpt_parts.append(f"**CC:** {cc_addr}")
            if bcc_addr: rcpt_parts.append(f"**BCC:** {bcc_addr}")
            st.markdown("  &nbsp;|&nbsp;  ".join(rcpt_parts))

        if df.empty:
            st.markdown(f"_No data found for **{sh['name']}** in the uploaded file. Make sure the sheet/tab name matches exactly._")
            continue

        # Show data table
        st.dataframe(df, use_container_width=True, height=min(200, 60 + 35 * len(df)))

        # Generate email buttons per row
        st.markdown(f"**Generate Outlook draft for each row:**")

        all_cols = COMMON_COLS[:3] + sh["extra_cols"] + [COMMON_COLS[3]]

        btn_cols = st.columns(min(len(df), 4))
        for idx, (_, row) in enumerate(df.iterrows()):
            parcel_id = str(row.get("dhlParcelId", f"Row {idx+1}")).strip() or f"Row {idx+1}"
            customer  = str(row.get("Customer Name", "")).strip()

            subject = fill_template(sh["subject"], row)
            body    = fill_template(sh["body"], row)
            mailto  = build_mailto(to_addr, cc_addr, bcc_addr, subject, body)

            label = f"✉️ {parcel_id}"
            if customer:
                label += f" · {customer[:18]}"

            with btn_cols[idx % 4]:
                st.markdown(
                    f'<a href="{mailto}" target="_blank" style="'
                    f'display:block;text-align:center;padding:0.45rem 0.6rem;'
                    f'background:{sh["color"]};color:white;border-radius:6px;'
                    f'font-size:0.8rem;font-weight:600;text-decoration:none;'
                    f'margin-bottom:0.4rem;">{label}</a>',
                    unsafe_allow_html=True,
                )

        # Bulk – generate one email for ALL rows in this sheet
        st.markdown("---")
        if to_addr and not df.empty:
            # Build combined body
            combined_body = f"Dear Team,\n\nPlease find below all {sh['name']} cases requiring attention:\n\n"
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                combined_body += f"{'='*50}\nRecord {idx}\n{'='*50}\n"
                combined_body += fill_template(sh["body"], row) + "\n\n"
            combined_subj  = f"[BULK] {sh['name']} – {count} records – DHL Operations"
            bulk_mailto    = build_mailto(to_addr, cc_addr, bcc_addr, combined_subj, combined_body)

            st.markdown(
                f'<a href="{bulk_mailto}" target="_blank" style="'
                f'display:inline-block;padding:0.5rem 1.2rem;'
                f'background:{sh["color"]};color:white;border-radius:6px;'
                f'font-weight:700;text-decoration:none;font-size:0.85rem;">'
                f'✉️ Generate Bulk Email — all {count} records</a>',
                unsafe_allow_html=True,
            )

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown(
    "<div style='text-align:center;color:#aaa;font-size:0.78rem;'>"
    "DHL Operations Email Dashboard &nbsp;|&nbsp; Edit <code>data/recipients.csv</code> to manage recipients"
    "</div>",
    unsafe_allow_html=True,
)
