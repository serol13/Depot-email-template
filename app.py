import streamlit as st
import pandas as pd
import urllib.parse
import os

st.set_page_config(page_title="DHL Operations – Email Dashboard", layout="wide")

RECIPIENTS_CSV = os.path.join(os.path.dirname(__file__), "data", "recipients.csv")
COMMON_COLS    = ["dhlParcelId", "Customer Name", "Date", "Remarks"]

SHEETS = [
    {
        "name": "Damage",
        "color": "#C00000",
        "extra_cols": ["Damage Type", "Damage Description", "Item Value (RM)"],
        "subject": "DHL Damage Report – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA damage case has been identified for the following shipment:\n\n  Parcel ID    : {dhlParcelId}\n  Customer     : {Customer Name}\n  Date          : {Date}\n  Damage Type : {Damage Type}\n  Description  : {Damage Description}\n  Item Value    : RM {Item Value (RM)}\n  Remarks       : {Remarks}\n\nKindly investigate and advise on the next course of action.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "MPS Incomplete",
        "color": "#ED7D31",
        "extra_cols": ["Total Pieces", "Pieces Received", "Missing Pieces"],
        "subject": "MPS Incomplete – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nAn incomplete Multi-Piece Shipment (MPS) has been detected:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  Total Pieces    : {Total Pieces}\n  Pieces Received : {Pieces Received}\n  Missing Pieces  : {Missing Pieces}\n  Remarks          : {Remarks}\n\nPlease trace the missing pieces and update accordingly.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Suspect Lost",
        "color": "#7030A0",
        "extra_cols": ["Last Scan Location", "Last Scan Date", "Investigation Status"],
        "subject": "Suspect Lost Shipment – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment is flagged as SUSPECT LOST:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Last Scan Location  : {Last Scan Location}\n  Last Scan Date      : {Last Scan Date}\n  Investigation Status : {Investigation Status}\n  Remarks               : {Remarks}\n\nPlease escalate and initiate a full trace immediately.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Duplicate",
        "color": "#D4A017",
        "extra_cols": ["Duplicate Parcel ID", "Root Cause"],
        "subject": "Duplicate Shipment Alert – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA duplicate shipment entry has been identified:\n\n  Parcel ID           : {dhlParcelId}\n  Customer            : {Customer Name}\n  Date                 : {Date}\n  Duplicate Parcel ID : {Duplicate Parcel ID}\n  Root Cause          : {Root Cause}\n  Remarks              : {Remarks}\n\nPlease verify and take corrective action.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "EMY Zipcode (Reroute)",
        "color": "#0070C0",
        "extra_cols": ["Original Zipcode", "Correct Zipcode", "Reroute Status"],
        "subject": "EMY Zipcode Reroute – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment requires rerouting due to an EMY zipcode issue:\n\n  Parcel ID        : {dhlParcelId}\n  Customer         : {Customer Name}\n  Date              : {Date}\n  Original Zipcode : {Original Zipcode}\n  Correct Zipcode  : {Correct Zipcode}\n  Reroute Status   : {Reroute Status}\n  Remarks           : {Remarks}\n\nKindly action the reroute at the earliest.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Missing DO",
        "color": "#C55A11",
        "extra_cols": ["DO Number", "Origin Station", "Action Required"],
        "subject": "Missing Delivery Order – {dhlParcelId} | DO: {DO Number}",
        "body": "Dear Team,\n\nA Delivery Order (DO) is missing for the following shipment:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  DO Number       : {DO Number}\n  Origin Station  : {Origin Station}\n  Action Required : {Action Required}\n  Remarks          : {Remarks}\n\nPlease provide or reissue the DO urgently.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Cancelled Shipment",
        "color": "#595959",
        "extra_cols": ["Cancellation Reason", "Cancelled By", "Refund Required"],
        "subject": "Cancelled Shipment Notice – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment has been cancelled:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Cancellation Reason  : {Cancellation Reason}\n  Cancelled By         : {Cancelled By}\n  Refund Required      : {Refund Required}\n  Remarks               : {Remarks}\n\nPlease ensure all related records are updated.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Prealert Shipment",
        "color": "#375623",
        "extra_cols": ["Origin Country", "Expected Arrival", "Prealert Reference"],
        "subject": "Prealert Notification – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA prealert has been received for the following inbound shipment:\n\n  Parcel ID           : {dhlParcelId}\n  Customer            : {Customer Name}\n  Date                 : {Date}\n  Origin Country      : {Origin Country}\n  Expected Arrival    : {Expected Arrival}\n  Prealert Reference  : {Prealert Reference}\n  Remarks              : {Remarks}\n\nPlease prepare for receiving and ensure documentation is in order.\n\nRegards,\nDHL Operations Team",
    },
]

SHEET_MAP = {s["name"]: s for s in SHEETS}

# ── Helpers ───────────────────────────────────────────────────────────────────
@st.cache_data
def load_recipients():
    try:
        return pd.read_csv(RECIPIENTS_CSV, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame(columns=["sheet", "to", "cc", "bcc"])

def get_recipients(sheet_name):
    df  = load_recipients()
    row = df[df["sheet"].str.strip() == sheet_name]
    if row.empty:
        return "", "", ""
    r = row.iloc[0]
    return r.get("to","").strip(), r.get("cc","").strip(), r.get("bcc","").strip()

def fill_template(template, row):
    out = template
    for k, v in row.items():
        out = out.replace("{" + str(k) + "}", str(v) if pd.notna(v) else "")
    return out

def mailto_link(to, cc, bcc, subject, body):
    p = {}
    if cc:  p["cc"]  = cc
    if bcc: p["bcc"] = bcc
    p["subject"] = subject
    p["body"]    = body
    return f"mailto:{urllib.parse.quote(to)}?{urllib.parse.urlencode(p, quote_via=urllib.parse.quote)}"

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');

    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { padding: 2rem 2.5rem 3rem; max-width: 1200px; }

    /* Hide default streamlit header/footer */
    #MainMenu, footer, header { visibility: hidden; }

    .page-title {
        font-size: 1.5rem; font-weight: 700;
        color: #1a1a2e; margin-bottom: 0.2rem;
    }
    .page-sub {
        font-size: 0.875rem; color: #888;
        margin-bottom: 2rem;
    }

    /* Case cards grid */
    .cards-grid {
        display: grid;
        grid-template-columns: repeat(4, 1fr);
        gap: 1rem;
        margin-bottom: 2rem;
    }
    .case-card {
        background: #ffffff;
        border: 1.5px solid #e8e8e8;
        border-radius: 10px;
        padding: 1.1rem 1.2rem;
        cursor: pointer;
        transition: all 0.15s ease;
        border-left: 4px solid #ccc;
    }
    .case-card:hover {
        box-shadow: 0 4px 12px rgba(0,0,0,0.1);
        transform: translateY(-2px);
    }
    .case-card.active {
        box-shadow: 0 4px 16px rgba(0,0,0,0.12);
        border-top: 1.5px solid #e8e8e8;
        border-right: 1.5px solid #e8e8e8;
        border-bottom: 1.5px solid #e8e8e8;
    }
    .card-label {
        font-size: 0.8rem; font-weight: 600;
        color: #444; margin-bottom: 0.2rem;
        text-transform: uppercase; letter-spacing: 0.04em;
    }
    .card-name {
        font-size: 1rem; font-weight: 700; color: #1a1a2e;
    }

    /* Section divider */
    .section-title {
        font-size: 0.8rem; font-weight: 600;
        text-transform: uppercase; letter-spacing: 0.08em;
        color: #aaa; margin: 1.5rem 0 1rem;
    }

    /* Input mode tabs */
    .stTabs [data-baseweb="tab-list"] {
        gap: 0; border-bottom: 2px solid #eee;
    }
    .stTabs [data-baseweb="tab"] {
        font-size: 0.85rem; font-weight: 600;
        padding: 0.5rem 1.2rem; color: #888;
    }
    .stTabs [aria-selected="true"] { color: #1a1a2e; }

    /* Form styling */
    .stTextInput input, .stTextArea textarea {
        border-radius: 7px; font-size: 0.88rem;
    }
    label { font-size: 0.82rem !important; font-weight: 600 !important; color: #555 !important; }

    /* Recipient bar */
    .rcpt-bar {
        background: #f7f8fc;
        border: 1px solid #e8e8e8;
        border-radius: 8px;
        padding: 0.6rem 1rem;
        font-size: 0.82rem;
        color: #555;
        margin-bottom: 1.2rem;
    }
    .rcpt-bar span { font-weight: 600; color: #1a1a2e; }

    /* Email preview box */
    .preview-box {
        background: #f9f9f9;
        border: 1px solid #e8e8e8;
        border-radius: 8px;
        padding: 1rem 1.2rem;
        font-size: 0.83rem;
        white-space: pre-wrap;
        color: #333;
        line-height: 1.6;
        max-height: 320px;
        overflow-y: auto;
    }
    .preview-subject {
        font-size: 0.82rem; font-weight: 600;
        color: #1a1a2e; margin-bottom: 0.5rem;
    }

    /* Open in Outlook button */
    .outlook-btn {
        display: inline-block;
        padding: 0.6rem 1.4rem;
        border-radius: 7px;
        font-size: 0.88rem;
        font-weight: 700;
        color: white !important;
        text-decoration: none !important;
        margin-top: 1rem;
    }

    /* Data table styling */
    .stDataFrame { border-radius: 8px; }

    /* Submit button */
    div.stButton > button {
        width: 100%;
        background: #1a1a2e;
        color: white;
        border: none;
        border-radius: 7px;
        font-weight: 600;
        font-size: 0.88rem;
        padding: 0.55rem;
    }
    div.stButton > button:hover { background: #2d2d4e; }

    /* Upload button area */
    .uploadedFile { border-radius: 8px; }
</style>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────────────────────
if "selected" not in st.session_state:
    st.session_state.selected = None
if "preview" not in st.session_state:
    st.session_state.preview = None  # dict: {subject, body, to, cc, bcc}

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="page-title">DHL Operations — Email Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Select a case type, fill in the details or upload a file, then open a draft in Outlook.</div>', unsafe_allow_html=True)

# ── Case cards ────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Case Types</div>', unsafe_allow_html=True)

cols = st.columns(4)
for i, sh in enumerate(SHEETS):
    with cols[i % 4]:
        is_active = st.session_state.selected == sh["name"]
        border_style = f"border-left: 4px solid {sh['color']};"
        bg = f"background: {sh['color']}08;" if is_active else ""
        st.markdown(
            f'<div class="case-card {"active" if is_active else ""}" style="{border_style}{bg}">'
            f'<div class="card-label">Case</div>'
            f'<div class="card-name" style="color:{sh["color"] if is_active else "#1a1a2e"}">{sh["name"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if st.button("Select", key=f"btn_{sh['name']}", use_container_width=True):
            st.session_state.selected = sh["name"]
            st.session_state.preview  = None
            st.rerun()

# ── Detail panel ──────────────────────────────────────────────────────────────
if not st.session_state.selected:
    st.info("Select a case type above to get started.")
    st.stop()

sh     = SHEET_MAP[st.session_state.selected]
to_a, cc_a, bcc_a = get_recipients(sh["name"])

st.markdown("---")
st.markdown(f'<div class="section-title">{sh["name"]}</div>', unsafe_allow_html=True)

# Recipient bar
if to_a:
    parts = [f"<span>To:</span> {to_a}"]
    if cc_a:  parts.append(f"<span>CC:</span> {cc_a}")
    if bcc_a: parts.append(f"<span>BCC:</span> {bcc_a}")
    st.markdown(f'<div class="rcpt-bar">{" &nbsp;&nbsp;|&nbsp;&nbsp; ".join(parts)}</div>', unsafe_allow_html=True)
else:
    st.warning("No recipient configured for this case type. Edit `data/recipients.csv` in your GitHub repo.")

# ── Input tabs ────────────────────────────────────────────────────────────────
tab_form, tab_upload = st.tabs(["Fill in Details", "Upload File"])

# ── TAB 1: Form ───────────────────────────────────────────────────────────────
with tab_form:
    all_cols = COMMON_COLS[:3] + sh["extra_cols"] + [COMMON_COLS[3]]

    with st.form(key=f"form_{sh['name']}"):
        # Common fields
        c1, c2, c3 = st.columns(3)
        with c1: pid  = st.text_input("dhlParcelId", placeholder="e.g. 1234567890")
        with c2: cust = st.text_input("Customer Name", placeholder="e.g. Ahmad bin Ali")
        with c3: date = st.text_input("Date", placeholder="e.g. 2026-03-10")

        # Extra fields
        if sh["extra_cols"]:
            ecols = st.columns(len(sh["extra_cols"]))
            extra_vals = {}
            for i, col in enumerate(sh["extra_cols"]):
                with ecols[i]:
                    extra_vals[col] = st.text_input(col, placeholder=col)
        else:
            extra_vals = {}

        remarks = st.text_area("Remarks", placeholder="Any additional notes...", height=80)
        submit  = st.form_submit_button("Preview Email")

    if submit:
        if not pid:
            st.error("dhlParcelId is required.")
        else:
            row  = {"dhlParcelId": pid, "Customer Name": cust, "Date": date,
                    "Remarks": remarks, **extra_vals}
            subj = fill_template(sh["subject"], row)
            body = fill_template(sh["body"], row)
            st.session_state.preview = {"subject": subj, "body": body,
                                        "to": to_a, "cc": cc_a, "bcc": bcc_a,
                                        "source": "single"}
            st.rerun()

# ── TAB 2: Upload ─────────────────────────────────────────────────────────────
with tab_upload:
    st.caption(f"Upload an Excel or CSV file with columns: {', '.join(COMMON_COLS[:3] + sh['extra_cols'] + [COMMON_COLS[3]])}")

    uploaded = st.file_uploader("Choose file", type=["xlsx", "csv"],
                                label_visibility="collapsed")

    if uploaded:
        try:
            if uploaded.name.endswith(".csv"):
                df = pd.read_csv(uploaded, dtype=str).fillna("")
            else:
                xls = pd.ExcelFile(uploaded)
                # Try to match sheet name, fallback to first sheet
                sheet_names = [s.lower() for s in xls.sheet_names]
                matched = next((xls.sheet_names[i] for i, s in enumerate(sheet_names)
                                if s == sh["name"].lower()), xls.sheet_names[0])
                df = xls.parse(matched, dtype=str).fillna("")

            st.dataframe(df, use_container_width=True, height=min(200, 55 + 35 * len(df)))

            if st.button("Generate Bulk Email", key="bulk_btn"):
                if not to_a:
                    st.error("No recipient configured. Edit data/recipients.csv first.")
                else:
                    combined = f"Dear Team,\n\nPlease find below all {sh['name']} cases requiring attention:\n\n"
                    for idx, (_, row) in enumerate(df.iterrows(), 1):
                        combined += f"{'='*46}\nRecord {idx}\n{'='*46}\n"
                        combined += fill_template(sh["body"], row) + "\n\n"
                    bulk_subj = f"[BULK] {sh['name']} – {len(df)} records – DHL Operations"
                    st.session_state.preview = {"subject": bulk_subj, "body": combined,
                                                "to": to_a, "cc": cc_a, "bcc": bcc_a,
                                                "source": "bulk", "count": len(df)}
                    st.rerun()

            # Per-row buttons
            st.markdown('<div class="section-title" style="margin-top:1rem">Generate per row</div>', unsafe_allow_html=True)
            rcols = st.columns(min(len(df), 4))
            for idx, (_, row) in enumerate(df.iterrows()):
                pid_val  = str(row.get("dhlParcelId", f"Row {idx+1}")).strip() or f"Row {idx+1}"
                cust_val = str(row.get("Customer Name","")).strip()
                label    = pid_val + (f" · {cust_val[:15]}" if cust_val else "")
                with rcols[idx % 4]:
                    if st.button(label, key=f"row_{idx}"):
                        subj = fill_template(sh["subject"], row)
                        body = fill_template(sh["body"], row)
                        st.session_state.preview = {"subject": subj, "body": body,
                                                    "to": to_a, "cc": cc_a, "bcc": bcc_a,
                                                    "source": "single"}
                        st.rerun()
        except Exception as e:
            st.error(f"Could not read file: {e}")

# ── Preview panel ─────────────────────────────────────────────────────────────
if st.session_state.preview:
    p = st.session_state.preview
    st.markdown("---")
    st.markdown('<div class="section-title">Email Preview</div>', unsafe_allow_html=True)

    left, right = st.columns([3, 1])
    with left:
        st.markdown(f'<div class="preview-subject">Subject: {p["subject"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="preview-box">{p["body"]}</div>', unsafe_allow_html=True)
    with right:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if p["to"]:
            link = mailto_link(p["to"], p["cc"], p["bcc"], p["subject"], p["body"])
            st.markdown(
                f'<a href="{link}" target="_blank" class="outlook-btn" '
                f'style="background:{sh["color"]};">Open in Outlook</a>',
                unsafe_allow_html=True,
            )
        else:
            st.warning("Set recipients in data/recipients.csv first.")
