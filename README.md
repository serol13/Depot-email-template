# DHL Operations – Email Dashboard

Streamlit app for generating Outlook draft emails per issue type.

## Project Structure

```
dhl_streamlit/
├── app.py                  # Main Streamlit app
├── requirements.txt
├── data/
│   └── recipients.csv      # ← Edit this to manage To / CC / BCC per issue type
└── README.md
```

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Managing Recipients

Edit `data/recipients.csv` directly in your repo:

| Column  | Description                                      |
|---------|--------------------------------------------------|
| `sheet` | Must match the issue type name exactly           |
| `to`    | Required. Semicolon-separated emails             |
| `cc`    | Optional. Semicolon-separated emails             |
| `bcc`   | Optional. Semicolon-separated emails             |

Example:
```csv
sheet,to,cc,bcc
Damage,ops@dhl.com,manager@dhl.com,
MPS Incomplete,ops@dhl.com,,
```

## How It Works

1. Upload an Excel (`.xlsx`) or CSV file
   - Excel: each **sheet tab name** must match the issue type (e.g. `Damage`, `MPS Incomplete`)
   - CSV: single sheet mapped to whichever issue type matches
2. Dashboard shows a summary count per issue type
3. Click any **✉️ parcel button** → Outlook opens with a pre-filled draft
4. Or click **Generate Bulk Email** → one email with all records for that issue type

## Issue Types & Required Columns

| Issue Type            | Extra Columns (beyond common fields)              |
|-----------------------|---------------------------------------------------|
| Damage                | Damage Type, Damage Description, Item Value (RM)  |
| MPS Incomplete        | Total Pieces, Pieces Received, Missing Pieces     |
| Suspect Lost          | Last Scan Location, Last Scan Date, Investigation Status |
| Duplicate             | Duplicate Parcel ID, Root Cause                   |
| EMY Zipcode (Reroute) | Original Zipcode, Correct Zipcode, Reroute Status |
| Missing DO            | DO Number, Origin Station, Action Required        |
| Cancelled Shipment    | Cancellation Reason, Cancelled By, Refund Required|
| Prealert Shipment     | Origin Country, Expected Arrival, Prealert Reference |

**Common fields on all sheets:** `dhlParcelId`, `Customer Name`, `Date`, `Remarks`
