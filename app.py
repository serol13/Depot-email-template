import streamlit as st
import pandas as pd
import urllib.parse
import os

st.set_page_config(page_title="DHL Operations – Email Dashboard", page_icon="📦", layout="wide")

RECIPIENTS_CSV = os.path.join(os.path.dirname(__file__), "data", "recipients.csv")
COMMON_COLS    = ["dhlParcelId", "Customer Name", "Date", "Remarks"]

SHEETS = [
    {
        "name": "Damage", "icon": "🔴", "color": "#C00000",
        "extra_cols": ["Damage Type", "Damage Description", "Item Value (RM)"],
        "subject": "DHL Damage Report – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA damage case has been identified:\n\n  Parcel ID    : {dhlParcelId}\n  Customer     : {Customer Name}\n  Date          : {Date}\n  Damage Type : {Damage Type}\n  Description  : {Damage Description}\n  Item Value    : RM {Item Value (RM)}\n  Remarks       : {Remarks}\n\nKindly investigate and advise.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "MPS Incomplete", "icon": "🟠", "color": "#ED7D31",
        "extra_cols": ["Total Pieces", "Pieces Received", "Missing Pieces"],
        "subject": "MPS Incomplete – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nAn incomplete MPS has been detected:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  Total Pieces    : {Total Pieces}\n  Pieces Received : {Pieces Received}\n  Missing Pieces  : {Missing Pieces}\n  Remarks          : {Remarks}\n\nPlease trace the missing pieces.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Suspect Lost", "icon": "🟣", "color": "#7030A0",
        "extra_cols": ["Last Scan Location", "Last Scan Date", "Investigation Status"],
        "subject": "Suspect Lost Shipment – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThis shipment is flagged as SUSPECT LOST:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Last Scan Location  : {Last Scan Location}\n  Last Scan Date      : {Last Scan Date}\n  Investigation Status : {Investigation Status}\n  Remarks               : {Remarks}\n\nPlease escalate and initiate a full trace immediately.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Duplicate", "icon": "🟡", "color": "#D4A017",
        "extra_cols": ["Duplicate Parcel ID", "Root Cause"],
        "subject": "Duplicate Shipment Alert – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA duplicate shipment entry has been identified:\n\n  Parcel ID           : {dhlParcelId}\n  Customer            : {Customer Name}\n  Date                 : {Date}\n  Duplicate Parcel ID : {Duplicate Parcel ID}\n  Root Cause          : {Root Cause}\n  Remarks              : {Remarks}\n\nPlease verify and take corrective action.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "EMY Zipcode (Reroute)", "icon": "🔵", "color": "#0070C0",
        "extra_cols": ["Original Zipcode", "Correct Zipcode", "Reroute Status"],
        "subject": "EMY Zipcode Reroute – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThis shipment requires rerouting due to an EMY zipcode issue:\n\n  Parcel ID        : {dhlParcelId}\n  Customer         : {Customer Name}\n  Date              : {Date}\n  Original Zipcode : {Original Zipcode}\n  Correct Zipcode  : {Correct Zipcode}\n  Reroute Status   : {Reroute Status}\n  Remarks           : {Remarks}\n\nKindly action the reroute at the earliest.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Missing DO", "icon": "🟤", "color": "#C55A11",
        "extra_cols": ["DO Number", "Origin Station", "Action Required"],
        "subject": "Missing Delivery Order – {dhlParcelId} | DO: {DO Number}",
        "body": "Dear Team,\n\nA Delivery Order (DO) is missing for this shipment:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  DO Number       : {DO Number}\n  Origin Station  : {Origin Station}\n  Action Required : {Action Required}\n  Remarks          : {Remarks}\n\nPlease provide or reissue the DO urgently.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Cancelled Shipment", "icon": "⚫", "color": "#595959",
        "extra_cols": ["Cancellation Reason", "Cancelled By", "Refund Required"],
        "subject": "Cancelled Shipment Notice – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment has been cancelled:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Cancellation Reason  : {Cancellation Reason}\n  Cancelled By         : {Cancelled By}\n  Refund Required      : {Refund Required}\n  Remarks               : {Remarks}\n\nPlease update all related records.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Prealert Shipment", "icon": "🟢", "color": "#375623",
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

def email_btn(label, to, cc, bcc, subject, body, color, width="auto"):
    if not to:
        st.warning("⚠️ No recipient set. Edit `data/recipients.csv`.")
        return
    st.markdown(
        f'<a href="{mailto_link(to,cc,bcc,subject,body)}" target="_blank" style="'
        f'display:inline-block;width:{width};text-align:center;padding:0.48rem 1rem;'
        f'background:{color};color:#fff;border-radius:7px;font-size:0.83rem;'
        f'font-weight:700;text-decoration:none;margin-bottom:0.35rem;">{label}</a>',
        unsafe_allow_html=True,
    )

def match_sheet(names, issue):
    for n in names:
        if n.strip().lower() == issue.lower():
            return n
    return None

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.block-container{padding-top:1.3rem;padding-bottom:2rem}
.dhl-title{font-size:1.65rem;font-weight:800;color:#1F3864;margin-bottom:0}
.dhl-sub{font-size:0.88rem;color:#777;margin-bottom:1rem}
.sec{font-size:0.95rem;font-weight:700;color:#1F3864;
     border-left:4px solid #2E75B6;padding-left:0.55rem;margin:1.1rem 0 0.5rem}
.sum-card{background:#fff;border-radius:9px;padding:0.65rem 0.5rem;
          border-top:4px solid #ccc;text-align:center;
          box-shadow:0 1px 4px rgba(0,0,0,0.07)}
.sum-n{font-size:1.45rem;font-weight:800}
.sum-l{font-size:0.68rem;color:#555;font-weight:600;margin-top:1px}
</style>
""", unsafe_allow_html=True)

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="dhl-title">📦 DHL Operations – Email Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="dhl-sub">Upload a file <i>or</i> fill in a form — then open a pre-filled Outlook draft in one click.</div>', unsafe_allow_html=True)

mode = st.radio("Mode", ["📂 Upload File (Excel / CSV)", "✏️ Fill Form Manually"],
                horizontal=True, label_visibility="collapsed")
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# MODE 1 – FILE UPLOAD
# ══════════════════════════════════════════════════════════════════════════════
if mode == "📂 Upload File (Excel / CSV)":

    uploaded = st.file_uploader("Upload Excel (.xlsx) or CSV", type=["xlsx","csv"],
        help="Excel: each sheet tab name must match the issue type exactly.\nCSV: single-sheet upload, matched by column headers.")

    @st.cache_data
    def load_file(f):
        if f.name.endswith(".csv"):
            return {"CSV Data": pd.read_csv(f, dtype=str).fillna("")}
        xls = pd.ExcelFile(f)
        return {s: xls.parse(s, dtype=str).fillna("") for s in xls.sheet_names}

    all_data = load_file(uploaded) if uploaded else {}

    # Summary bar
    st.markdown('<div class="sec">Summary</div>', unsafe_allow_html=True)
    scols = st.columns(len(SHEETS))
    for i, sh in enumerate(SHEETS):
        mk = match_sheet(all_data.keys(), sh["name"]) if all_data else None
        n  = len(all_data[mk]) if mk else 0
        with scols[i]:
            st.markdown(
                f'<div class="sum-card" style="border-top-color:{sh["color"]}">'
                f'<div class="sum-n" style="color:{sh["color"]}">{n}</div>'
                f'<div class="sum-l">{sh["icon"]} {sh["name"]}</div></div>',
                unsafe_allow_html=True)

    if not uploaded:
        st.info("👆 Upload a file to populate the dashboard.")
        st.stop()

    # Issue cards
    st.markdown('<div class="sec">Issue Breakdown</div>', unsafe_allow_html=True)
    for sh in SHEETS:
        mk    = match_sheet(all_data.keys(), sh["name"])
        df    = all_data[mk].copy() if mk else pd.DataFrame()
        count = len(df)
        to_a, cc_a, bcc_a = get_recipients(sh["name"])

        with st.expander(f"{sh['icon']}  **{sh['name']}**  —  {count} record{'s' if count!=1 else ''}", expanded=count>0):
            if not to_a:
                st.warning("⚠️ No **To** recipient. Edit `data/recipients.csv`.")
            else:
                parts = [f"**To:** {to_a}"]
                if cc_a:  parts.append(f"**CC:** {cc_a}")
                if bcc_a: parts.append(f"**BCC:** {bcc_a}")
                st.caption("  |  ".join(parts))

            if df.empty:
                st.info(f"No data found for **{sh['name']}**. Sheet/tab name must match exactly.")
                continue

            st.dataframe(df, use_container_width=True, height=min(220, 55+35*count))

            st.markdown("**Generate draft per row:**")
            bcols = st.columns(min(count, 4))
            for idx, (_, row) in enumerate(df.iterrows()):
                pid   = str(row.get("dhlParcelId", f"Row {idx+1}")).strip() or f"Row {idx+1}"
                cust  = str(row.get("Customer Name","")).strip()
                subj  = fill_template(sh["subject"], row)
                body  = fill_template(sh["body"], row)
                label = f"✉️ {pid}" + (f" · {cust[:15]}" if cust else "")
                with bcols[idx % 4]:
                    email_btn(label, to_a, cc_a, bcc_a, subj, body, sh["color"])

            st.markdown("---")
            cb = f"Dear Team,\n\nAll {sh['name']} cases requiring attention:\n\n"
            for idx, (_, row) in enumerate(df.iterrows(), 1):
                cb += f"{'='*46}\nRecord {idx}\n{'='*46}\n" + fill_template(sh["body"], row) + "\n\n"
            email_btn(f"✉️ Bulk Email — all {count} records", to_a, cc_a, bcc_a,
                      f"[BULK] {sh['name']} – {count} records – DHL Ops", cb, sh["color"])

# ══════════════════════════════════════════════════════════════════════════════
# MODE 2 – MANUAL FORM
# ══════════════════════════════════════════════════════════════════════════════
else:
    st.markdown('<div class="sec">Select Issue Type</div>', unsafe_allow_html=True)

    selected = st.selectbox("Issue Type", [s["name"] for s in SHEETS],
                            format_func=lambda x: f"{SHEET_MAP[x]['icon']}  {x}",
                            label_visibility="collapsed")
    sh = SHEET_MAP[selected]
    to_a, cc_a, bcc_a = get_recipients(sh["name"])

    if not to_a:
        st.warning("⚠️ No **To** recipient for this issue type. Edit `data/recipients.csv`.")
    else:
        parts = [f"**To:** {to_a}"]
        if cc_a:  parts.append(f"**CC:** {cc_a}")
        if bcc_a: parts.append(f"**BCC:** {bcc_a}")
        st.caption("  |  ".join(parts))

    st.markdown('<div class="sec">Fill in Details</div>', unsafe_allow_html=True)

    with st.form("manual_form", clear_on_submit=False):
        c1, c2, c3 = st.columns(3)
        with c1: pid  = st.text_input("dhlParcelId *", placeholder="e.g. 1234567890")
        with c2: cust = st.text_input("Customer Name *", placeholder="e.g. Ahmad bin Ali")
        with c3: date = st.text_input("Date *", placeholder="e.g. 2026-03-10")

        extra_vals = {}
        if sh["extra_cols"]:
            ecols = st.columns(len(sh["extra_cols"]))
            for i, col in enumerate(sh["extra_cols"]):
                with ecols[i]:
                    extra_vals[col] = st.text_input(col, placeholder=f"{col}")

        remarks = st.text_area("Remarks", placeholder="Any additional notes...", height=80)
        submit  = st.form_submit_button("Preview Email →", use_container_width=True)

    if submit:
        if not pid:
            st.error("dhlParcelId is required.")
        else:
            row = {"dhlParcelId": pid, "Customer Name": cust, "Date": date,
                   "Remarks": remarks, **extra_vals}
            subj = fill_template(sh["subject"], row)
            body = fill_template(sh["body"], row)

            st.markdown('<div class="sec">Email Preview</div>', unsafe_allow_html=True)
            pcol, bcol = st.columns([3, 1])
            with pcol:
                st.markdown(f"**Subject:** `{subj}`")
                st.text_area("Body preview", body, height=300, disabled=True)
            with bcol:
                st.markdown("<br><br><br>", unsafe_allow_html=True)
                email_btn("✉️ Open in Outlook", to_a, cc_a, bcc_a, subj, body, sh["color"], width="100%")

# ── Footer ────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown("<div style='text-align:center;color:#bbb;font-size:0.77rem;'>"
            "DHL Operations Email Dashboard &nbsp;|&nbsp; "
            "Edit <code>data/recipients.csv</code> in GitHub to manage recipients</div>",
            unsafe_allow_html=True)
