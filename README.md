# 🚢 Ship Report Validation System

A Streamlit-based web application for validating vessel reports, detecting anomalies, and sending automated email alerts to ship operators.

---

## Features

- **Automated Report Validation** — Detects rule violations across key engine and operational parameters
- **Calculated Metrics** — Automatically computes Report Hours, SFOC, and SCOC from raw data
- **Downloadable Results** — Export failed validation reports as Excel files
- **Email Notifications** — Send vessel-specific or bulk alerts via SMTP with attachments
- **Visual Summary** — Bar charts of failure reason frequencies

---

## Installation

### Prerequisites

- Python 3.8+
- pip

### Install Dependencies

```bash
pip install streamlit pandas numpy openpyxl
```

### Run the App

```bash
streamlit run app.py
```

---

## Input File Format

Upload an Excel file containing a sheet named **`All Reports`** with the following columns:

| Column | Description |
|--------|-------------|
| `Ship Name` | Vessel name |
| `IMO_No` | IMO number |
| `Report Type` | `At Sea`, `At Port`, or `At Anchorage` |
| `Start Date` / `End Date` | Report period dates |
| `Start Time` / `End Time` | Report period times |
| `Time Shift` | Time zone offset in hours |
| `Average Load [kW]` | Main engine average load |
| `Average Load [%]` | Main engine load percentage |
| `ME Rhrs (From Last Report)` | Main engine running hours |
| `Avg. Speed` | Average vessel speed (knots) |
| `Fuel Cons. [MT] (ME Cons 1/2/3)` | Main engine fuel consumption |
| `Cyl. Oil Cons. [Ltrs]` | Cylinder oil consumption |
| `Exh. Temp [°C] (Main Engine Unit 1–16)` | Exhaust temperatures per unit |
| `A.E. 1–6 Last Report [Rhrs]` | Auxiliary engine running hours |
| Sub-consumer fields | Tank Cleaning, Cargo Transfer, Maintaining Cargo Temp, etc. |

---

## Validation Rules

### Rule 1 — SFOC (Specific Fuel Oil Consumption)
- **Applies to:** At Sea reports where ME Rhrs > 12
- **Valid range:** 150–200 g/kWh
- **Formula:** `(ME Cons 1 + ME Cons 2 + ME Cons 3) × 1,000,000 / (Avg Load [kW] × ME Rhrs)`

### Rule 2 — Average Speed
- **Applies to:** At Sea reports where ME Rhrs > 12
- **Valid range:** 0–20 knots

### Rule 3 — Exhaust Temperature Deviation
- **Applies to:** At Sea reports where ME Rhrs > 12
- **Condition:** Each unit's exhaust temp must be within ±50°C of the fleet average across Units 1–16

### Rule 4 — ME Running Hours vs Report Hours
- **Applies to:** All report types
- **Condition:** ME Rhrs must not exceed calculated Report Hours by more than 1 hour
- **Report Hours formula:** `(End Date/Time − Start Date/Time) + Time Shift`

### Rule 5 — Auxiliary Engines & Sub-Consumers
- **Applies to:** At Sea reports where ME Load > 40%
- **Condition:** If `AE Rhrs Sum / Report Hours > 1.25` (indicating 2+ AEs running), at least one sub-consumer must be non-zero
- **Sub-consumers checked:** Tank Cleaning, Cargo Transfer, Maintaining Cargo Temp, Shaft Gen. Propulsion, Raising Cargo Temp, Burning Sludge, Ballast Transfer, Fresh Water Production, Others, EGCS Consumption

### Rule 6 — SCOC (Specific Cylinder Oil Consumption)
- **Applies to:** At Sea reports where ME Rhrs > 12
- **Valid range:** 0.8–1.5 g/kWh
- **Formula:** `Cyl. Oil Cons. [Ltrs] × 1000 / (Avg Load [kW] × ME Rhrs)`

---

## Email Notifications

### Single Vessel
Select a vessel from the dropdown, enter recipient and CC emails (comma-separated), and send a report for that vessel only.

### Bulk Send to All Vessels
Upload a **Vessel Email Mapping** file (Excel or CSV) with the following structure:

| Ship Name | Email | CC1 | CC2 |
|-----------|-------|-----|-----|
| Vessel A | captain@vessel-a.com, chief@vessel-a.com | manager@company.com | office@company.com |
| Vessel B | vesselb@company.com | supervisor@company.com | |

- `Ship Name` column is required
- `Email` or `To` column is required (supports comma-separated multiple addresses)
- `CC1`, `CC2`, ... columns are optional

### SMTP Configuration (Sidebar)

| Field | Example |
|-------|---------|
| SMTP Server | `smtp.gmail.com` |
| SMTP Port | `587` |
| Sender Email | `your-email@company.com` |
| Password | App Password (see below) |

#### Gmail Setup
1. Enable 2-Factor Authentication on your Google account
2. Generate an App Password at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
3. Use the App Password (not your regular Gmail password)

#### Outlook / Office 365
- SMTP Server: `smtp.office365.com`
- SMTP Port: `587`

---

## Output

- **On-screen table** of all failed reports with reasons highlighted
- **Bar chart** summarising failure reason frequencies
- **Downloadable Excel** (`Failed_Validation.xlsx`) containing failed rows plus calculated SFOC, SCOC, and Report Hours columns
- **Email alerts** with the failed report attached as an Excel file per vessel

---

## Project Structure

```
app.py                  # Main Streamlit application
README.md               # This file
```

---

## Contact

For support, contact: sunilrrb73@gmail.com
