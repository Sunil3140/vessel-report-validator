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


# ---------------- Helper: Report Hours Calculation ----------------
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
                start_datetime = pd.to_datetime(
                    f"{start_date.date()} {start_time}", errors="coerce"
                )
                end_datetime = pd.to_datetime(
                    f"{end_date.date()} {end_time}", errors="coerce"
                )
                total_hours = (end_datetime - start_datetime).total_seconds() / 3600 + float(
                    time_shift
                )
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

    aux_message = (
        "Two or more Aux Engines running at sea with ME Load > 40% and no sub-consumers reported. "
        "Please confirm operations and update relevant sub-consumption fields if applicable."
    )

    if "Reason" not in df.columns:
        df["Reason"] = ""

    df.loc[cond, "Reason"] = df.loc[cond, "Reason"].astype(str).apply(
        lambda x: (x + "; " + aux_message).strip("; ").strip()
    )

    return df


# ---------------- Main Validation Function ----------------
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
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        report_hours = row.get("Report Hours", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")

        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")

        if report_hours > 0:
            diff = ME_Rhrs - report_hours
            if diff > 1.0:
                reason.append(
                    f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {diff:.2f}h (margin ±1h)"
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
        "Start Time",
        "End Date",
        "End Time",
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
def send_email(
    smtp_server,
    smtp_port,
    sender_email,
    sender_password,
    recipient_emails,
    subject,
    body,
    attachment_data=None,
    attachment_name="Failed_Validation.xlsx",
    cc_emails=None,
):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = recipient_emails
        if cc_emails:
            msg["Cc"] = cc_emails
        msg["Subject"] = subject

        msg.attach(MIMEText(body, "html"))

        if attachment_data is not None:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_data.getvalue())
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition", f"attachment; filename={attachment_name}"
            )
            msg.attach(part)

        recipients = recipient_emails.split(",") + (cc_emails.split(",") if cc_emails else [])

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()

        return True, "Email sent successfully!"
    except Exception as e:
        return False, f"Email sending failed: {str(e)}"


def create_email_body(vessel_name, failed_count):
    return f"""
    <html>
    <body style="font-family: Arial, sans-serif; color: #333;">
        <h3>Vessel Report Validation Alert</h3>
        <p>Dear Captain and Chief Engineer of <b>{vessel_name}</b>,</p>
        <p>This is an automated notification regarding <b>{failed_count}</b> failed report(s).</p>
        <p>Please review the attached file for detailed information and take corrective action.</p>
        <p style="color: gray; font-size: 12px;">This message was generated automatically on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}.</p>
    </body>
    </html>
    """


# ---------------- Streamlit App ----------------
def main():
    st.set_page_config(page_title="Ship Report Validator", layout="wide")
    st.title("🚢 Ship Report Validation System")

    uploaded_file = st.file_uploader("Upload Excel File (sheet: All Reports)", type=["xlsx", "xls"])

    smtp_server = st.sidebar.text_input("SMTP Server", value="smtp.gmail.com")
    smtp_port = st.sidebar.number_input("SMTP Port", value=587)
    sender_email = st.sidebar.text_input("Sender Email")
    sender_password = st.sidebar.text_input("Email Password", type="password")

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="All Reports")
        failed, df_with_calcs = validate_reports(df.copy())

        st.success("✅ Validation Completed")

        st.subheader("Failed Reports")
        st.write(f"Total Failed Reports: {len(failed)}")

        def highlight_failed_rows(_):
            return ["background-color: #f8d7da; color: black"] * len(failed.columns)

        if not failed.empty:
            st.dataframe(failed.style.apply(highlight_failed_rows, axis=1), use_container_width=True)

            # Download Failed Reports
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
            output.seek(0)
            st.download_button(
                label="📥 Download Failed Reports",
                data=output,
                file_name="Failed_Validation.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            # ---------------- Individual Email Sending ----------------
            st.divider()
            st.subheader("📧 Send Email to Individual Vessel")
            vessels = failed["Ship Name"].unique()
            selected_vessel = st.selectbox("Select Vessel", vessels)
            vessel_data = failed[failed["Ship Name"] == selected_vessel]
            vessel_email = st.text_input("To Email")
            vessel_cc = st.text_input("CC (optional)")

            if st.button("Send Email to Selected Vessel"):
                subject = f"Vessel Report Validation Alert - {selected_vessel}"
                body = create_email_body(selected_vessel, len(vessel_data))

                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                    vessel_data.to_excel(writer, index=False)
                buffer.seek(0)

                success, message = send_email(
                    smtp_server,
                    smtp_port,
                    sender_email,
                    sender_password,
                    vessel_email,
                    subject,
                    body,
                    buffer,
                    f"{selected_vessel}_Failed_Validation.xlsx",
                    cc_emails=vessel_cc,
                )

                if success:
                    st.success(message)
                else:
                    st.error(message)

            # ---------------- Bulk Email Sending ----------------
            st.divider()
            st.subheader("📦 Bulk Email Sending")

            email_mapping_file = st.file_uploader(
                "Upload Vessel Email Mapping (Ship Name, Email, CC optional)",
                type=["xlsx", "csv"],
            )

            if email_mapping_file:
                if email_mapping_file.name.endswith(".csv"):
                    email_df = pd.read_csv(email_mapping_file)
                else:
                    email_df = pd.read_excel(email_mapping_file)

                st.dataframe(email_df.head())

                if st.button("Send Bulk Emails"):
                    for vessel in vessels:
                        match = email_df[email_df["Ship Name"] == vessel]
                        if match.empty:
                            continue
                        vessel_email = match.iloc[0]["Email"]
                        vessel_cc = match.iloc[0].get("CC", "")

                        vessel_data = failed[failed["Ship Name"] == vessel]
                        subject = f"Vessel Report Validation Alert - {vessel}"
                        body = create_email_body(vessel, len(vessel_data))

                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                            vessel_data.to_excel(writer, index=False)
                        buffer.seek(0)

                        send_email(
                            smtp_server,
                            smtp_port,
                            sender_email,
                            sender_password,
                            vessel_email,
                            subject,
                            body,
                            buffer,
                            f"{vessel}_Failed_Validation.xlsx",
                            cc_emails=vessel_cc,
                        )

                    st.success("✅ Bulk Emails Sent Successfully!")

        else:
            st.success("🎉 All reports passed validation!")


if __name__ == "__main__":
    main()
