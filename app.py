# app.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# -------------------------
# Helper calculations
# -------------------------
def calculate_report_hours(df):
    """Calculate Report Hours from Start Date/Time, End Date/Time and Time Shift (vectorized-ish)"""
    report_hours = []
    for idx, row in df.iterrows():
        try:
            start_date = pd.to_datetime(row.get("Start Date"), errors='coerce')
            end_date = pd.to_datetime(row.get("End Date"), errors='coerce')
            start_time = str(row.get("Start Time", "00:00:00")).strip()
            end_time = str(row.get("End Time", "00:00:00")).strip()
            time_shift = row.get("Time Shift", 0) or 0

            if pd.notna(start_date) and pd.notna(end_date):
                # parse times safely
                try:
                    start_time_obj = pd.to_datetime(start_time, errors='coerce').time()
                except Exception:
                    start_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()
                try:
                    end_time_obj = pd.to_datetime(end_time, errors='coerce').time()
                except Exception:
                    end_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()

                start_datetime = datetime.combine(start_date.date(), start_time_obj)
                end_datetime = datetime.combine(end_date.date(), end_time_obj)

                total_hours = (end_datetime - start_datetime).total_seconds() / 3600.0 + float(time_shift)
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
        except Exception:
            report_hours.append(0)
    return report_hours

# -------------------------
# Aux Engine Validation (vectorized)
# -------------------------
def aux_engine_validation(df):
    """Vectorized validation to detect unusual Aux Engine operation at sea.

    Condition:
    (sum of AE unit Rhrs) / Report Period > 1.25
    AND Average Load [%] > 40
    AND sum(sub-consumers) == 0
    AND Report Type == 'At Sea'
    """

    ae_cols = [
        "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)",
        "A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)",
        "A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)",
        "A.E. 4 Total [Rhrs] (Aux Engine Unit 4)",
        "A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)",
        "A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)"
    ]
    sub_cols = [
        "Tank Cleaning [MT]",
        "Cargo Transfer [MT]",
        "Maintaining Cargo Temp. [MT]",
        "Shaft Gen. Propulsion [MT]",
        "Raising Cargo Temp. [MT]",
        "Burning Sludge [MT]",
        "Ballast Transfer [MT]",
        "Fresh Water Prod. [MT]",
        "Others [MT]",
        "EGCS Consumption [MT]",
    ]

    # Ensure columns exist to avoid KeyErrors; fill missing with 0
    for c in ae_cols + sub_cols + ["Average Load [%]", "Report Hours", "Report Type"]:
        if c not in df.columns:
            df[c] = 0

    # Convert numeric-like columns to numeric safely
    df[ae_cols] = df[ae_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    df[sub_cols] = df[sub_cols].apply(pd.to_numeric, errors='coerce').fillna(0)
    df["Average Load [%]"] = pd.to_numeric(df["Average Load [%]"], errors='coerce').fillna(0)
    df["Report Hours"] = pd.to_numeric(df["Report Hours"], errors='coerce').fillna(0)

    # Totals
    df["AE_Total_Rhrs"] = df[ae_cols].sum(axis=1)
    df["Sub_Consumption_Total"] = df[sub_cols].sum(axis=1)

    # Build condition
    cond = (
        df["Report Type"].astype(str).str.strip().eq("At Sea")
    ) & (
        df["Report Hours"] > 0
    ) & (
        (df["AE_Total_Rhrs"] / df["Report Hours"]) > 1.25
    ) & (
        df["Average Load [%]"] > 40
    ) & (
        df["Sub_Consumption_Total"] == 0
    )

    # Mark the flag
    df["Aux_Flag"] = False
    df.loc[cond, "Aux_Flag"] = True

    # Append clear reason for flagged rows
    aux_message = (
        "Two or more Aux Engines running at sea with ME Load > 40% and no sub-consumers reported. "
        "Please confirm operations and update relevant sub-consumption fields if applicable."
    )
    # Ensure Reason column exists
    if "Reason" not in df.columns:
        df["Reason"] = ""
    # Append message while preserving existing reasons
    df.loc[cond, "Reason"] = df.loc[cond, "Reason"].astype(str).apply(lambda x: (x + "; " + aux_message).strip("; ").strip())

    return df

# -------------------------
# Main validation function (keeps old rules + aux rule)
# -------------------------
def validate_reports(df):
    # --- Clean numeric columns (preserve earlier logic) ---
    numeric_cols = [
        "Average Load [kW]",
        "ME Rhrs (From Last Report)",
        "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)",
        "Fuel Cons. [MT] (ME Cons 2)",
        "Fuel Cons. [MT] (ME Cons 3)",
        "Time Shift"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.replace(",", "")
                .str.strip()
                .replace(["", "nan", "None"], np.nan)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --- Calculate Report Hours ---
    df["Report Hours"] = calculate_report_hours(df)

    # --- Calculate SFOC in g/kWh ---
    # Avoid division by zero by temporarily replacing zeros with NaN
    numerator = (
        df.get("Fuel Cons. [MT] (ME Cons 1)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 2)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 3)", 0).fillna(0)
    ) * 1_000_000
    denom = df.get("Average Load [kW]", 0).replace(0, np.nan) * df.get("ME Rhrs (From Last Report)", 0).replace(0, np.nan)
    df["SFOC"] = (numerator / denom).fillna(0)

    reasons = []
    fail_columns = set()

    # Existing row-by-row rules (SFOC, Avg Speed, Exhaust deviations, ME Rhrs vs Report Hours)
    for idx, row in df.iterrows():
        reason = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        report_hours = row.get("Report Hours", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        # Rule 1: SFOC (only for At Sea)
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")

        # Rule 2: Avg Speed (only for At Sea)
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")

        # Rule 3: Exhaust Temp deviation (Units 1–16, only At Sea)
        if report_type == "At Sea" and ME_Rhrs > 12:
            exhaust_cols = [
                f"Exh. Temp [°C] (Main Engine Unit {j})"
                for j in range(1, 17)
                if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns
            ]
            temps = [row[c] for c in exhaust_cols if pd.notna(row.get(c, None)) and row.get(c) != 0]
            if temps:
                avg_temp = np.mean(temps)
                for j, c in enumerate(exhaust_cols, start=1):
                    val = row.get(c, None)
                    if pd.notna(val) and val != 0 and abs(val - avg_temp) > 50:
                        reason.append(f"Exhaust temp deviation > ±50 from avg at Unit {j}")
                        fail_columns.add(c)

        # Rule 4: ME Rhrs should not exceed Report Hours (with ±1 hour margin)
        if report_hours > 0:
            hours_diff = ME_Rhrs - report_hours
            if hours_diff > 1.0:
                reason.append(f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {hours_diff:.2f}h (margin: ±1h)")
                fail_columns.add("ME Rhrs (From Last Report)")
                fail_columns.add("Report Hours")

        reasons.append("; ".join(reason))

    # Populate Reason column from existing rules
    df["Reason"] = reasons

    # Insert aux engine validation (vectorized) - this adds 'Aux_Flag' and appends reason where applicable
    df = aux_engine_validation(df)

    # Build failed dataframe (any non-empty Reason)
    failed = df[df["Reason"].astype(str).str.strip() != ""].copy()

    # Build columns to show in failed output (preserve helpful context)
    exhaust_cols_present = [
        f"Exh. Temp [°C] (Main Engine Unit {j})"
        for j in range(1, 17)
        if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns
    ]
    context_cols = [
        "Ship Name",
        "IMO_No",
        "Report Type",
        "Start Date",
        "Start Time",
        "End Date",
        "End Time",
        "Voyage Number",
        "Time Zone",
        "Distance - Ground [NM]",
        "Time Shift",
        "Distance - Sea [NM]",
        "Average Load [kW]",
        "Average RPM",
        "Average Load [%]",
        "ME Rhrs (From Last Report)",
        "Report Hours",
    ]

    # Combine columns while avoiding missing ones
    cols_to_keep = [c for c in (context_cols + exhaust_cols_present + ["AE_Total_Rhrs", "Sub_Consumption_Total", "Aux_Flag", "Reason"]) if c in failed.columns]

    failed = failed[cols_to_keep]

    return failed, df

# -------------------------
# Email utilities
# -------------------------
def send_email(smtp_server, smtp_port, sender_email, sender_password,
               recipient_emails, subject, body, attachment_data=None,
               attachment_name="Failed_Validation.xlsx", cc_emails=None):
    """Send email with optional attachment to multiple recipients"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email

        # To recipients
        if isinstance(recipient_emails, str):
            recipient_list = [email.strip() for email in recipient_emails.split(',') if email.strip()]
        else:
            recipient_list = recipient_emails or []

        msg['To'] = ', '.join(recipient_list)

        # CC handling
        cc_list = []
        if cc_emails:
            if isinstance(cc_emails, str):
                cc_list = [email.strip() for email in cc_emails.split(',') if email.strip()]
            else:
                cc_list = cc_emails
            if cc_list:
                msg['Cc'] = ', '.join(cc_list)

        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'html'))

        # Attachment
        if attachment_data is not None:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_data.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
            msg.attach(part)

        all_recipients = recipient_list + cc_list

        server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, all_recipients, msg.as_string())
        server.quit()

        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Failed to send email: {str(e)}"

def create_email_body(ship_name, failed_count, reasons_summary):
    """Create HTML email body"""
    body = f"""
    <html>
        <body style="font-family: Arial, sans-serif; line-height: 1.6; color: #333;">
            <h2 style="color: #2c3e50;">Vessel Report Validation Alert</h2>
            <p>Dear Captain and C/E of <strong>{ship_name}</strong>,</p>
            <p>This is an automated notification regarding recent validation failures in your vessel reports.</p>
            <div style="background-color: #fff3cd; border-left: 4px solid #ffc107; padding: 15px; margin: 20px 0;">
                <h3 style="margin-top: 0; color: #856404;">Validation Summary</h3>
                <p><strong>Failed Reports:</strong> {failed_count}</p>
            </div>
            <h3>Common Issues Detected:</h3>
            <ul>
    {reasons_summary}
            </ul>
            <p>Please review the attached Excel file for detailed information about the failed validations.</p>
            <h4 style="color: #2c3e50;">Action Required:</h4>
            <ol>
                <li>Review the attached report carefully</li>
                <li>Correct the identified issues</li>
                <li>Resubmit corrected reports</li>
                <li>Contact the technical team if you need assistance</li>
            </ol>
            <hr style="border: none; border-top: 1px solid #ddd; margin: 30px 0;">
            <p style="color: #7f8c8d; font-size: 0.9em;">
                For any queries, please contact us at <strong><a href="mailto:smartapp@enginelink.blue">smartapp@enginelink.blue</a></strong>
            </p>
            <p style="color: #7f8c8d; font-size: 0.85em; margin-top: 10px;">
                This is an automated message. Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            </p>
        </body>
    </html>
    """
    return body

# -------------------------
# Streamlit UI
# -------------------------
def main():
    st.set_page_config(page_title="Ship Report Validator", page_icon="🚢", layout="wide")

    # session state
    if 'validation_done' not in st.session_state:
        st.session_state.validation_done = False
    if 'failed_df' not in st.session_state:
        st.session_state.failed_df = None
    if 'df_with_calcs' not in st.session_state:
        st.session_state.df_with_calcs = None
    if 'original_df' not in st.session_state:
        st.session_state.original_df = None

    st.title("🚢 Ship Report Validation System")
    st.markdown("Upload your Excel file to validate ship reports and send automated alerts")

    # Sidebar
    with st.sidebar:
        st.header("📋 Validation Rules (summary)")
        st.markdown("""
        - SFOC (At Sea, ME Rhrs > 12): 150–200 g/kWh  
        - Avg Speed (At Sea, ME Rhrs > 12): 0–20 knots  
        - Exhaust Temp deviation (At Sea, ME Rhrs > 12): ±50°C from avg  
        - ME Rhrs must not exceed Report Hours by > 1 hour  
        - Aux Engine rule: Two or more AEs running at sea with ME Load > 40% and no sub-consumers reported -> flagged
        """)
        st.divider()
        st.header("📧 Email Configuration")
        with st.expander("SMTP Settings", expanded=False):
            smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
            smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
            sender_email = st.text_input("Sender Email", placeholder="your-email@company.com")
            sender_password = st.text_input("Password / App Password", type="password")

    uploaded_file = st.file_uploader("Choose an Excel file (sheet: All Reports)", type=["xlsx", "xls"])

    # handle file upload and validation
    if uploaded_file is not None:
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        if 'current_file_id' not in st.session_state or st.session_state.current_file_id != file_id:
            st.session_state.current_file_id = file_id
            st.session_state.validation_done = False
            st.session_state.failed_df = None
            st.session_state.df_with_calcs = None
            st.session_state.original_df = None

    if uploaded_file is not None and not st.session_state.validation_done:
        try:
            with st.spinner("Loading and validating file..."):
                df = pd.read_excel(uploaded_file, sheet_name="All Reports")
                st.session_state.original_df = df

                # Validate reports
                failed, df_with_calcs = validate_reports(df.copy())

                st.session_state.failed_df = failed
                st.session_state.df_with_calcs = df_with_calcs
                st.session_state.validation_done = True

            st.success(f"✅ File loaded and validated! Total rows: {len(df)}")
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.exception(e)

    # Results view
    if st.session_state.validation_done:
        df = st.session_state.original_df
        failed = st.session_state.failed_df
        df_with_calcs = st.session_state.df_with_calcs

        # dataset info
        with st.expander("📊 Dataset Information"):
            st.write(f"**Rows:** {len(df)}")
            st.write(f"**Columns:** {len(df.columns)}")
            st.write("**Column Names:**")
            st.write(df.columns.tolist())

        st.header("📈 Validation Results")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Reports", len(df))
        with col2:
            st.metric("Failed Reports", len(failed))
        with col3:
            pass_rate = ((len(df) - len(failed)) / len(df) * 100) if len(df) > 0 else 0
            st.metric("Pass Rate", f"{pass_rate:.1f}%")

        if not failed.empty:
            st.warning(f"⚠️ {len(failed)} reports failed validation")

            # highlight function for Aux_Flag only in failed section
            def highlight_aux_flag(row):
                if row.get("Aux_Flag") is True:
                    return ["background-color: #f8d7da; color: #000000"] * len(row)
                else:
                    return
