import streamlit as st
import pandas as pd
import urllib.parse
import os
import base64
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.base import MIMEBase
from email import encoders

st.set_page_config(page_title="DHL Operations – Email Dashboard", layout="wide")

RECIPIENTS_CSV = os.path.join(os.path.dirname(__file__), "recipients.csv")
COMMON_COLS    = ["dhlParcelId", "Customer Name", "Date", "Remarks"]

SHEETS = [
    {
        "name": "Damage",
        "color": "#C00000",
        "extra_cols": ["Damage Type", "Damage Description", "Item Value (RM)"],
        "has_attachment": True,
        "subject": "DHL Damage Report – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA damage case has been identified for the following shipment:\n\n  Parcel ID    : {dhlParcelId}\n  Customer     : {Customer Name}\n  Date          : {Date}\n  Damage Type : {Damage Type}\n  Description  : {Damage Description}\n  Item Value    : RM {Item Value (RM)}\n  Remarks       : {Remarks}\n\nPlease refer to the attached photo for visual evidence.\n\nKindly investigate and advise on the next course of action.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "MPS Incomplete",
        "color": "#ED7D31",
        "extra_cols": ["Total Pieces", "Pieces Received", "Missing Pieces"],
        "has_attachment": False,
        "subject": "MPS Incomplete – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nAn incomplete Multi-Piece Shipment (MPS) has been detected:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  Total Pieces    : {Total Pieces}\n  Pieces Received : {Pieces Received}\n  Missing Pieces  : {Missing Pieces}\n  Remarks          : {Remarks}\n\nPlease trace the missing pieces and update accordingly.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Suspect Lost",
        "color": "#7030A0",
        "extra_cols": ["Last Scan Location", "Last Scan Date", "Investigation Status"],
        "has_attachment": False,
        "subject": "Suspect Lost Shipment – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment is flagged as SUSPECT LOST:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Last Scan Location  : {Last Scan Location}\n  Last Scan Date      : {Last Scan Date}\n  Investigation Status : {Investigation Status}\n  Remarks               : {Remarks}\n\nPlease escalate and initiate a full trace immediately.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Duplicate",
        "color": "#D4A017",
        "extra_cols": ["Duplicate Parcel ID", "Root Cause"],
        "has_attachment": False,
        "subject": "Duplicate Shipment Alert – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA duplicate shipment entry has been identified:\n\n  Parcel ID           : {dhlParcelId}\n  Customer            : {Customer Name}\n  Date                 : {Date}\n  Duplicate Parcel ID : {Duplicate Parcel ID}\n  Root Cause          : {Root Cause}\n  Remarks              : {Remarks}\n\nPlease verify and take corrective action.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "EMY Zipcode (Reroute)",
        "color": "#0070C0",
        "extra_cols": ["Original Zipcode", "Correct Zipcode", "Reroute Status"],
        "has_attachment": False,
        "subject": "EMY Zipcode Reroute – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment requires rerouting due to an EMY zipcode issue:\n\n  Parcel ID        : {dhlParcelId}\n  Customer         : {Customer Name}\n  Date              : {Date}\n  Original Zipcode : {Original Zipcode}\n  Correct Zipcode  : {Correct Zipcode}\n  Reroute Status   : {Reroute Status}\n  Remarks           : {Remarks}\n\nKindly action the reroute at the earliest.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Missing DO",
        "color": "#C55A11",
        "extra_cols": ["DO Number", "Origin Station", "Action Required"],
        "has_attachment": False,
        "subject": "Missing Delivery Order – {dhlParcelId} | DO: {DO Number}",
        "body": "Dear Team,\n\nA Delivery Order (DO) is missing for the following shipment:\n\n  Parcel ID       : {dhlParcelId}\n  Customer        : {Customer Name}\n  Date             : {Date}\n  DO Number       : {DO Number}\n  Origin Station  : {Origin Station}\n  Action Required : {Action Required}\n  Remarks          : {Remarks}\n\nPlease provide or reissue the DO urgently.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Cancelled Shipment",
        "color": "#595959",
        "extra_cols": ["Cancellation Reason", "Cancelled By", "Refund Required"],
        "has_attachment": False,
        "subject": "Cancelled Shipment Notice – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nThe following shipment has been cancelled:\n\n  Parcel ID            : {dhlParcelId}\n  Customer             : {Customer Name}\n  Date                  : {Date}\n  Cancellation Reason  : {Cancellation Reason}\n  Cancelled By         : {Cancelled By}\n  Refund Required      : {Refund Required}\n  Remarks               : {Remarks}\n\nPlease ensure all related records are updated.\n\nRegards,\nDHL Operations Team",
    },
    {
        "name": "Prealert Shipment",
        "color": "#375623",
        "extra_cols": ["Origin Country", "Expected Arrival", "Prealert Reference"],
        "has_attachment": False,
        "subject": "Prealert Notification – {dhlParcelId} | {Customer Name}",
        "body": "Dear Team,\n\nA prealert has been received for the following inbound shipment:\n\n  Parcel ID           : {dhlParcelId}\n  Customer            : {Customer Name}\n  Date                 : {Date}\n  Origin Country      : {Origin Country}\n  Expected Arrival    : {Expected Arrival}\n  Prealert Reference  : {Prealert Reference}\n  Remarks              : {Remarks}\n\nPlease prepare for receiving and ensure documentation is in order.\n\nRegards,\nDHL Operations Team",
    },
]

SHEET_MAP = {s["name"]: s for s in SHEETS}

# ── Helpers ───────────────────────────────────────────────────────────────────
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

def build_eml(to, cc, bcc, subject, body, attachment_bytes=None, attachment_name=None):
    """Build a .eml file bytes with optional image attachment."""
    msg = MIMEMultipart()
    msg["To"]      = to
    msg["Subject"] = subject
    if cc:  msg["Cc"]  = cc
    if bcc: msg["Bcc"] = bcc

    msg.attach(MIMEText(body, "plain"))

    if attachment_bytes and attachment_name:
        # Detect mime type from extension
        ext = attachment_name.rsplit(".", 1)[-1].lower()
        mime_map = {"jpg": "jpeg", "jpeg": "jpeg", "png": "png", "gif": "gif", "bmp": "bmp", "webp": "webp"}
        img_type = mime_map.get(ext)

        if img_type:
            part = MIMEImage(attachment_bytes, _subtype=img_type)
        else:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_bytes)
            encoders.encode_base64(part)

        part.add_header("Content-Disposition", "attachment", filename=attachment_name)
        msg.attach(part)

    return msg.as_bytes()

# ── CSS ───────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
    .block-container { padding: 2rem 2.5rem 3rem; max-width: 1200px; }
    #MainMenu, footer, header { visibility: hidden; }

    .page-title { font-size: 1.5rem; font-weight: 700; color: #1a1a2e; margin-bottom: 0.2rem; }
    .page-sub   { font-size: 0.875rem; color: #888; margin-bottom: 2rem; }

    .case-card {
        background: #ffffff; border: 1.5px solid #e8e8e8;
        border-radius: 10px; padding: 1.1rem 1.2rem;
        cursor: pointer; border-left: 4px solid #ccc;
    }
    .card-label { font-size: 0.8rem; font-weight: 600; color: #aaa;
                  text-transform: uppercase; letter-spacing: 0.04em; margin-bottom: 0.2rem; }
    .card-name  { font-size: 1rem; font-weight: 700; color: #1a1a2e; }

    .section-title { font-size: 0.78rem; font-weight: 600; text-transform: uppercase;
                     letter-spacing: 0.08em; color: #aaa; margin: 1.5rem 0 0.8rem; }

    .stTabs [data-baseweb="tab-list"] { gap: 0; border-bottom: 2px solid #eee; }
    .stTabs [data-baseweb="tab"] { font-size: 0.85rem; font-weight: 600; padding: 0.5rem 1.2rem; color: #888; }
    .stTabs [aria-selected="true"] { color: #1a1a2e; }

    .stTextInput input, .stTextArea textarea { border-radius: 7px; font-size: 0.88rem; }
    label { font-size: 0.82rem !important; font-weight: 600 !important; color: #555 !important; }

    .rcpt-bar { background: #f7f8fc; border: 1px solid #e8e8e8; border-radius: 8px;
                padding: 0.6rem 1rem; font-size: 0.82rem; color: #555; margin-bottom: 1.2rem; }
    .rcpt-bar span { font-weight: 600; color: #1a1a2e; }

    .preview-box { background: #f9f9f9; border: 1px solid #e8e8e8; border-radius: 8px;
                   padding: 1rem 1.2rem; font-size: 0.83rem; white-space: pre-wrap;
                   color: #333; line-height: 1.6; max-height: 320px; overflow-y: auto; }
    .preview-subject { font-size: 0.82rem; font-weight: 600; color: #1a1a2e; margin-bottom: 0.5rem; }

    .outlook-btn { display: inline-block; padding: 0.6rem 1.4rem; border-radius: 7px;
                   font-size: 0.88rem; font-weight: 700; color: white !important;
                   text-decoration: none !important; margin-top: 0.5rem; }

    .attach-note { background: #fff8e1; border: 1px solid #ffe082; border-radius: 7px;
                   padding: 0.6rem 0.9rem; font-size: 0.8rem; color: #7a5c00; margin-top: 0.5rem; }

    .stDataFrame { border-radius: 8px; }

    div.stButton > button { width: 100%; background: #1a1a2e; color: white; border: none;
                            border-radius: 7px; font-weight: 600; font-size: 0.88rem; padding: 0.55rem; }
    div.stButton > button:hover { background: #2d2d4e; }
</style>
""", unsafe_allow_html=True)

# ── State ─────────────────────────────────────────────────────────────────────
if "selected"   not in st.session_state: st.session_state.selected   = None
if "preview"    not in st.session_state: st.session_state.preview    = None
if "attach_img" not in st.session_state: st.session_state.attach_img = None  # {"bytes", "name"}

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown('<div class="page-title">DHL Operations — Email Dashboard</div>', unsafe_allow_html=True)
st.markdown('<div class="page-sub">Select a case type, fill in the details or upload a file, then open a draft in Outlook.</div>', unsafe_allow_html=True)

# ── Case cards ────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Case Types</div>', unsafe_allow_html=True)
cols = st.columns(4)
for i, sh in enumerate(SHEETS):
    with cols[i % 4]:
        is_active    = st.session_state.selected == sh["name"]
        border_style = f"border-left: 4px solid {sh['color']};"
        bg           = f"background: {sh['color']}0d;" if is_active else ""
        st.markdown(
            f'<div class="case-card" style="{border_style}{bg}">'
            f'<div class="card-label">Case</div>'
            f'<div class="card-name" style="color:{sh["color"] if is_active else "#1a1a2e"}">{sh["name"]}</div>'
            f'</div>',
            unsafe_allow_html=True,
        )
        if st.button("Select", key=f"btn_{sh['name']}", use_container_width=True):
            st.session_state.selected   = sh["name"]
            st.session_state.preview    = None
            st.session_state.attach_img = None
            st.rerun()

# ── Guard ─────────────────────────────────────────────────────────────────────
if not st.session_state.selected:
    st.info("Select a case type above to get started.")
    st.stop()

sh = SHEET_MAP[st.session_state.selected]
to_a, cc_a, bcc_a = get_recipients(sh["name"])
is_damage = sh["name"] == "Damage"

st.markdown("---")
st.markdown(f'<div class="section-title">{sh["name"]}</div>', unsafe_allow_html=True)

# Recipient bar
if to_a:
    parts = [f"<span>To:</span> {to_a}"]
    if cc_a:  parts.append(f"<span>CC:</span> {cc_a}")
    if bcc_a: parts.append(f"<span>BCC:</span> {bcc_a}")
    st.markdown(f'<div class="rcpt-bar">{" &nbsp;&nbsp;|&nbsp;&nbsp; ".join(parts)}</div>', unsafe_allow_html=True)
else:
    st.warning("No recipient configured for this case type. Edit `recipients.csv` in your GitHub repo.")

# ── Tabs ──────────────────────────────────────────────────────────────────────
tab_form, tab_upload = st.tabs(["Fill in Details", "Upload File"])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — FORM
# ══════════════════════════════════════════════════════════════════════════════
with tab_form:

    # Image uploader for Damage — OUTSIDE form so it stays reactive
    img_file_form = None
    if is_damage:
        st.markdown('<div class="section-title">Damage Photo</div>', unsafe_allow_html=True)
        img_file_form = st.file_uploader(
            "Upload damage photo (JPG, PNG, etc.)",
            type=["jpg","jpeg","png","gif","bmp","webp"],
            key="img_form",
            label_visibility="collapsed"
        )
        if img_file_form:
            st.image(img_file_form, width=220, caption=img_file_form.name)

    with st.form(key=f"form_{sh['name']}"):
        c1, c2, c3 = st.columns(3)
        with c1: pid  = st.text_input("dhlParcelId",   placeholder="e.g. 1234567890")
        with c2: cust = st.text_input("Customer Name", placeholder="e.g. Ahmad bin Ali")
        with c3: date = st.text_input("Date",          placeholder="e.g. 2026-03-10")

        extra_vals = {}
        if sh["extra_cols"]:
            ecols = st.columns(len(sh["extra_cols"]))
            for i, col in enumerate(sh["extra_cols"]):
                with ecols[i]:
                    extra_vals[col] = st.text_input(col, placeholder=col)

        remarks = st.text_area("Remarks", placeholder="Any additional notes...", height=80)
        submit  = st.form_submit_button("Preview Email")

    if submit:
        if not pid:
            st.error("dhlParcelId is required.")
        else:
            row  = {"dhlParcelId": pid, "Customer Name": cust,
                    "Date": date, "Remarks": remarks, **extra_vals}
            subj = fill_template(sh["subject"], row)
            body = fill_template(sh["body"], row)

            attach = None
            if is_damage and img_file_form:
                attach = {"bytes": img_file_form.read(), "name": img_file_form.name}

            st.session_state.preview    = {"subject": subj, "body": body,
                                           "to": to_a, "cc": cc_a, "bcc": bcc_a}
            st.session_state.attach_img = attach
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — UPLOAD FILE
# ══════════════════════════════════════════════════════════════════════════════
with tab_upload:
    st.caption(f"Columns needed: {', '.join(COMMON_COLS[:3] + sh['extra_cols'] + [COMMON_COLS[3]])}")

    # Image uploader for Damage
    img_file_upload = None
    if is_damage:
        st.markdown('<div class="section-title">Damage Photo</div>', unsafe_allow_html=True)
        img_file_upload = st.file_uploader(
            "Upload damage photo",
            type=["jpg","jpeg","png","gif","bmp","webp"],
            key="img_upload",
            label_visibility="collapsed"
        )
        if img_file_upload:
            st.image(img_file_upload, width=220, caption=img_file_upload.name)

    data_file = st.file_uploader("Choose Excel or CSV", type=["xlsx","csv"],
                                 label_visibility="collapsed", key="data_upload")

    if data_file:
        try:
            if data_file.name.endswith(".csv"):
                df = pd.read_csv(data_file, dtype=str).fillna("")
            else:
                xls   = pd.ExcelFile(data_file)
                names = [s.lower() for s in xls.sheet_names]
                match = next((xls.sheet_names[i] for i,s in enumerate(names)
                              if s == sh["name"].lower()), xls.sheet_names[0])
                df = xls.parse(match, dtype=str).fillna("")

            st.dataframe(df, use_container_width=True, height=min(200, 55+35*len(df)))

            attach_upload = None
            if is_damage and img_file_upload:
                attach_upload = {"bytes": img_file_upload.read(), "name": img_file_upload.name}

            # Bulk email button
            if st.button("Generate Bulk Email", key="bulk_btn"):
                if not to_a:
                    st.error("No recipient configured. Edit recipients.csv first.")
                else:
                    combined  = f"Dear Team,\n\nPlease find below all {sh['name']} cases requiring attention:\n\n"
                    for idx, (_, row) in enumerate(df.iterrows(), 1):
                        combined += f"{'='*46}\nRecord {idx}\n{'='*46}\n"
                        combined += fill_template(sh["body"], row) + "\n\n"
                    bulk_subj = f"[BULK] {sh['name']} – {len(df)} records – DHL Operations"
                    st.session_state.preview    = {"subject": bulk_subj, "body": combined,
                                                   "to": to_a, "cc": cc_a, "bcc": bcc_a}
                    st.session_state.attach_img = attach_upload
                    st.rerun()

            # Per-row buttons
            st.markdown('<div class="section-title" style="margin-top:1rem">Generate per row</div>', unsafe_allow_html=True)
            rcols = st.columns(min(len(df), 4))
            for idx, (_, row) in enumerate(df.iterrows()):
                pid_v  = str(row.get("dhlParcelId", f"Row {idx+1}")).strip() or f"Row {idx+1}"
                cust_v = str(row.get("Customer Name","")).strip()
                label  = pid_v + (f" · {cust_v[:15]}" if cust_v else "")
                with rcols[idx % 4]:
                    if st.button(label, key=f"row_{idx}"):
                        subj = fill_template(sh["subject"], row)
                        body = fill_template(sh["body"], row)
                        st.session_state.preview    = {"subject": subj, "body": body,
                                                       "to": to_a, "cc": cc_a, "bcc": bcc_a}
                        st.session_state.attach_img = attach_upload
                        st.rerun()

        except Exception as e:
            st.error(f"Could not read file: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# PREVIEW PANEL
# ══════════════════════════════════════════════════════════════════════════════
if st.session_state.preview:
    p      = st.session_state.preview
    attach = st.session_state.attach_img

    st.markdown("---")
    st.markdown('<div class="section-title">Email Preview</div>', unsafe_allow_html=True)

    left, right = st.columns([3, 1])

    with left:
        st.markdown(f'<div class="preview-subject">Subject: {p["subject"]}</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="preview-box">{p["body"]}</div>', unsafe_allow_html=True)
        if attach:
            st.markdown(f'<div class="attach-note">Attachment: {attach["name"]}</div>', unsafe_allow_html=True)

    with right:
        st.markdown("<br><br>", unsafe_allow_html=True)
        if not p["to"]:
            st.warning("Set recipients in recipients.csv first.")
        elif is_damage and attach:
            # Build .eml with attachment → download button
            eml_bytes = build_eml(
                p["to"], p["cc"], p["bcc"],
                p["subject"], p["body"],
                attach["bytes"], attach["name"]
            )
            st.download_button(
                label="Download .eml (opens in Outlook with attachment)",
                data=eml_bytes,
                file_name="damage_report.eml",
                mime="message/rfc822",
                use_container_width=True,
            )
            st.markdown(
                '<div style="font-size:0.75rem;color:#888;margin-top:0.4rem;line-height:1.5;">'
                'Double-click the downloaded .eml file — Outlook opens it as a draft with the photo attached.'
                '</div>',
                unsafe_allow_html=True
            )
        else:
            # No attachment or non-damage — use mailto as usual
            link = mailto_link(p["to"], p["cc"], p["bcc"], p["subject"], p["body"])
            st.markdown(
                f'<a href="{link}" target="_blank" class="outlook-btn" '
                f'style="background:{sh["color"]};">Open in Outlook</a>',
                unsafe_allow_html=True,
            )
            if is_damage and not attach:
                st.markdown(
                    '<div style="font-size:0.75rem;color:#888;margin-top:0.5rem;">'
                    'No photo uploaded — opening without attachment.'
                    '</div>',
                    unsafe_allow_html=True
                )
