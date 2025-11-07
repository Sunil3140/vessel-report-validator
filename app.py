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

# ---------------- Helper Function ----------------
def calculate_report_hours(df):
    report_hours = []
    for _, row in df.iterrows():
        try:
            start_date = pd.to_datetime(row.get("Start Date"), errors="coerce")
            end_date = pd.to_datetime(row.get("End Date"), errors="coerce")
            start_time = str(row.get("Start Time", "00:00:00")).strip()
            end_time = str(row.get("End Time", "00:00:00")).strip()
            time_shift = row.get("Time Shift", 0) or 0

            if pd.notna(start_date) and pd.notna(end_date):
                start_dt = pd.to_datetime(f"{start_date.date()} {start_time}", errors="coerce")
                end_dt = pd.to_datetime(f"{end_date.date()} {end_time}", errors="coerce")
                total_hours = (end_dt - start_dt).total_seconds() / 3600 + float(time_shift)
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
        except Exception:
            report_hours.append(0)
    return report_hours


# ---------------- Aux Engine Validation ----------------
def aux_engine_validation(df):
    ae_cols = [
        "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)",
        "A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)",
        "A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)",
        "A.E. 4 Total [Rhrs] (Aux Engine Unit 4)",
        "A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)",
        "A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)",
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

    for c in ae_cols + sub_cols + ["Average Load [%]", "Report Hours", "Report Type"]:
        if c not in df.columns:
            df[c] = 0

    df[ae_cols + sub_cols] = df[ae_cols + sub_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    df["Average Load [%]"] = pd.to_numeric(df["Average Load [%]"], errors="coerce").fillna(0)
    df["Report Hours"] = pd.to_numeric(df["Report Hours"], errors="coerce").fillna(0)

    df["AE_Total_Rhrs"] = df[ae_cols].sum(axis=1)
    df["Sub_Consumption_Total"] = df[sub_cols].sum(axis=1)

    cond = (
        df["Report Type"].astype(str).str.strip().eq("At Sea")
        & (df["Report Hours"] > 0)
        & ((df["AE_Total_Rhrs"] / df["Report Hours"]) > 1.25)
        & (df["Average Load [%]"] > 40)
        & (df["Sub_Consumption_Total"] == 0)
    )

    aux_msg = (
        "Two or more Aux Engines running at sea with ME Load > 40% and no sub-consumers reported. "
        "Please confirm operations and update relevant sub-consumption fields if applicable."
    )

    df.loc[cond, "Reason"] = df.loc[cond, "Reason"].astype(str).apply(
        lambda x: (x + "; " + aux_msg).strip("; ").strip()
    )
    return df


# ---------------- Validation Rules ----------------
def validate_reports(df):
    numeric_cols = [
        "Average Load [kW]",
        "ME Rhrs (From Last Report)",
        "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)",
        "Fuel Cons. [MT] (ME Cons 2)",
        "Fuel Cons. [MT] (ME Cons 3)",
        "Time Shift",
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = (
                df[col].astype(str).str.replace(",", "").replace(["", "nan"], np.nan)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    df["Report Hours"] = calculate_report_hours(df)
    numerator = (
        df.get("Fuel Cons. [MT] (ME Cons 1)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 2)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 3)", 0).fillna(0)
    ) * 1_000_000
    denom = (
        df.get("Average Load [kW]", 0).replace(0, np.nan)
        * df.get("ME Rhrs (From Last Report)", 0).replace(0, np.nan)
    )
    df["SFOC"] = (numerator / denom).fillna(0)

    reasons = []
    for _, row in df.iterrows():
        reason = []
        rpt_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        rpt_hrs = row.get("Report Hours", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        if rpt_type == "At Sea" and ME_Rhrs > 12 and not (150 <= sfoc <= 200):
            reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")

        if rpt_type == "At Sea" and ME_Rhrs > 12 and not (0 <= avg_speed <= 20):
            reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")

        if rpt_hrs > 0 and (ME_Rhrs - rpt_hrs) > 1:
            reason.append(
                f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({rpt_hrs:.2f}) by {(ME_Rhrs - rpt_hrs):.2f}h"
            )

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    df = aux_engine_validation(df)
    failed = df[df["Reason"].astype(str).str.strip() != ""].copy()
    context_cols = [
        "Ship Name",
        "IMO_No",
        "Report Type",
        "Start Date",
        "End Date",
        "Average Load [kW]",
        "Average RPM",
        "Average Load [%]",
        "ME Rhrs (From Last Report)",
        "Report Hours",
        "SFOC",
        "Reason",
    ]
    cols_to_keep = [c for c in context_cols if c in failed.columns]
    failed = failed[cols_to_keep]
    return failed, df


# ---------------- Email Functions ----------------
def send_email(smtp_server, smtp_port, sender_email, sender_password, to, subject, body, attachment, filename, cc=None):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = to
        if cc:
            msg["Cc"] = cc
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))
        if attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.getvalue())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f"attachment; filename={filename}")
            msg.attach(part)

        recipients = to.split(",") + (cc.split(",") if cc else [])
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        return True, "Email sent successfully!"
    except Exception as e:
        return False, str(e)


def create_email_body(vessel, failed_count):
    return f"""
    <html><body style="font-family: Arial;">
    <h3>Vessel Report Validation Alert - {vessel}</h3>
    <p>{failed_count} failed report(s) detected.</p>
    <p>Please review the attached Excel file for details and update as necessary.</p>
    <p style='font-size:12px;color:gray;'>Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
    </body></html>
    """


# ---------------- Streamlit App ----------------
def main():
    st.set_page_config(page_title="Ship Report Validation System", layout="wide")
    st.title("🚢 Ship Report Validation System")
    st.markdown("Upload your Excel file to validate ship reports and send automated alerts")

    # Sidebar
    with st.sidebar:
        st.header("📘 Validation Rules")
        st.markdown("""
        **Rule 1: SFOC (Specific Fuel Oil Consumption)**  
        - At Sea (ME Rhrs > 12): 150–200 g/kWh  
        - At Port/Anchorage: No validation  

        **Rule 2: Average Speed**  
        - At Sea (ME Rhrs > 12): 0–20 knots  
        - At Port/Anchorage: No validation  

        **Rule 3: Exhaust Temperature**  
        - At Sea (ME Rhrs > 12): ±50°C deviation  
        - Applies to Units 1–16  

        **Rule 4: ME Running Hours**  
        - ME Rhrs must not exceed Report Hours by >1 hour  
        - Tolerance: ±1 hour  

        **Rule 5: Aux Engine Operation**  
        - At Sea: AE running hours >1.25× report hours  
        - ME Load >40%  
        - No sub-consumers reported  
        """)

        st.divider()
        st.header("📧 Email Configuration")
        smtp_server = st.text_input("SMTP Server", "smtp.gmail.com")
        smtp_port = st.number_input("SMTP Port", 587)
        sender_email = st.text_input("Sender Email")
        sender_password = st.text_input("App Password", type="password")

    uploaded_file = st.file_uploader("Choose Excel File (sheet: All Reports)", type=["xlsx", "xls"])
    if not uploaded_file:
        st.warning("⚠️ Please upload an Excel file to begin validation.")
        with st.expander("Expected Data Structure"):
            st.write("Sheet name: **All Reports**")
        with st.expander("Email Setup Guide"):
            st.write("Use your company email or Gmail App Password.")
        return

    df = pd.read_excel(uploaded_file, sheet_name="All Reports")
    failed, df_with_calcs = validate_reports(df.copy())
    st.success("✅ Validation completed")

    # Highlight all failed rows
    def highlight_failed_rows(row):
        return ["background-color: #f8d7da; color: black"] * len(row)

    st.subheader("📊 Failed Reports")
    if not failed.empty:
        st.dataframe(failed.style.apply(highlight_failed_rows, axis=1), use_container_width=True)
    else:
        st.success("🎉 All reports passed validation!")
        return

    # Download Failed
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        failed.to_excel(writer, index=False)
    buffer.seek(0)
    st.download_button("📥 Download Failed Reports", buffer, "Failed_Validation.xlsx")

    # Individual Email
    st.divider()
    st.subheader("📧 Send Email to Specific Vessel")
    vessels = failed["Ship Name"].unique()
    selected = st.selectbox("Select Vessel", vessels)
    to = st.text_input("To Email")
    cc = st.text_input("CC (optional)")
    if st.button("Send Email to Selected Vessel"):
        vessel_data = failed[failed["Ship Name"] == selected]
        subject = f"Vessel Report Validation Alert - {selected}"
        body = create_email_body(selected, len(vessel_data))
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            vessel_data.to_excel(writer, index=False)
        buf.seek(0)
        success, msg = send_email(
            smtp_server, smtp_port, sender_email, sender_password, to, subject, body, buf, f"{selected}_Failed.xlsx", cc
        )
        if success:
            st.success(msg)
        else:
            st.error(msg)

    # Bulk Email
    st.divider()
    st.subheader("📦 Bulk Email Sending")
    email_file = st.file_uploader("Upload Email Mapping (Ship Name, Email, CC)", type=["xlsx", "csv"])
    if email_file:
        email_df = pd.read_excel(email_file) if email_file.name.endswith(".xlsx") else pd.read_csv(email_file)
        if st.button("Send Bulk Emails"):
            for vessel in vessels:
                row = email_df[email_df["Ship Name"] == vessel]
                if row.empty:
                    continue
                to = row.iloc[0]["Email"]
                cc = row.iloc[0].get("CC", "")
                vessel_data = failed[failed["Ship Name"] == vessel]
                subject = f"Vessel Report Validation Alert - {vessel}"
                body = create_email_body(vessel, len(vessel_data))
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                    vessel_data.to_excel(writer, index=False)
                buf.seek(0)
                send_email(smtp_server, smtp_port, sender_email, sender_password, to, subject, body, buf, f"{vessel}_Failed.xlsx", cc)
            st.success("✅ Bulk emails sent successfully!")


if __name__ == "__main__":
    main()
