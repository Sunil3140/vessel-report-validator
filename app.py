# app.py
import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
import traceback
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# ---------------------------
# Helper: Report Hours
# ---------------------------
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
                hours = (end_dt - start_dt).total_seconds() / 3600.0
                total_hours = hours + float(time_shift)
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
        except Exception:
            report_hours.append(0)
    return report_hours


# ---------------------------
# Aux Engine validation (vectorized)
# ---------------------------
def apply_aux_engine_rule(df):
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

    # ensure columns exist
    for c in ae_cols + sub_cols + ["Average Load [%]", "Report Hours", "Report Type", "Reason"]:
        if c not in df.columns:
            df[c] = 0

    # numeric conversions
    df[ae_cols] = df[ae_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
    df[sub_cols] = df[sub_cols].apply(pd.to_numeric, errors="coerce").fillna(0)
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

    # append message keeping previous reasons
    df.loc[cond, "Reason"] = df.loc[cond, "Reason"].astype(str).apply(
        lambda x: (x + "; " + aux_message).strip("; ").strip()
    )

    # drop helper columns we don't want to show downstream
    df.drop(columns=["AE_Total_Rhrs", "Sub_Consumption_Total"], inplace=True, errors="ignore")

    return df


# ---------------------------
# Validation rules (existing + aux)
# ---------------------------
def validate_reports(df):
    # numeric cleaning
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
            df[col] = df[col].astype(str).str.replace(",", "").replace(["", "nan", "None"], np.nan)
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # report hours
    df["Report Hours"] = calculate_report_hours(df)

    # SFOC calculation (g/kWh)
    numerator = (
        df.get("Fuel Cons. [MT] (ME Cons 1)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 2)", 0).fillna(0)
        + df.get("Fuel Cons. [MT] (ME Cons 3)", 0).fillna(0)
    ) * 1_000_000
    denom = df.get("Average Load [kW]", 0).replace(0, np.nan) * df.get("ME Rhrs (From Last Report)", 0).replace(0, np.nan)
    df["SFOC"] = (numerator / denom).fillna(0)

    # row-by-row checks (some rules are easier in loop)
    reasons = []
    for _, row in df.iterrows():
        r = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0) or 0
        report_hours = row.get("Report Hours", 0) or 0
        sfoc = row.get("SFOC", 0) or 0
        avg_speed = row.get("Avg. Speed", 0) or 0

        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                r.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
            if not (0 <= avg_speed <= 20):
                r.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")

            # exhaust temp deviation rule across units (if present)
            exhaust_cols = [f"Exh. Temp [°C] (Main Engine Unit {j})" for j in range(1, 17) if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns]
            temps = [row[c] for c in exhaust_cols if pd.notna(row.get(c)) and row.get(c) != 0]
            if temps:
                avg_temp = np.mean(temps)
                for idx_col, c in enumerate(exhaust_cols, start=1):
                    val = row.get(c)
                    if pd.notna(val) and val != 0 and abs(val - avg_temp) > 50:
                        r.append(f"Exhaust temp deviation > ±50 at Unit {idx_col}")

        if report_hours > 0 and (ME_Rhrs - report_hours) > 1:
            r.append(f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {(ME_Rhrs - report_hours):.2f}h")

        reasons.append("; ".join(r))

    df["Reason"] = reasons

    # apply Aux Engine rule which appends to Reason when needed
    df = apply_aux_engine_rule(df)

    # build failed dataframe
    failed = df[df["Reason"].astype(str).str.strip() != ""].copy()

    # keep helpful context columns if present
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
        "SFOC",
        # "Reason" will be appended
    ]
    cols_to_keep = [c for c in context_cols if c in failed.columns] + ["Reason"]
    # ensure unique preserve order
    cols_to_keep = [c for i, c in enumerate(cols_to_keep) if c not in cols_to_keep[:i]]
    failed = failed[cols_to_keep]

    return failed, df


# ---------------------------
# Email utilities
# ---------------------------
def send_email(smtp_server, smtp_port, sender_email, sender_password, recipient_list, subject, body, attachment_buffer=None, attachment_name="Failed_Validation.xlsx", cc_list=None):
    try:
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = ", ".join(recipient_list) if isinstance(recipient_list, (list, tuple)) else recipient_list
        if cc_list:
            msg["Cc"] = ", ".join(cc_list) if isinstance(cc_list, (list, tuple)) else cc_list
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "html"))

        if attachment_buffer is not None:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment_buffer.getvalue())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", f'attachment; filename="{attachment_name}"')
            msg.attach(part)

        recipients = []
        if isinstance(recipient_list, (list, tuple)):
            recipients.extend(recipient_list)
        else:
            recipients.extend([e.strip() for e in str(recipient_list).split(",") if e.strip()])
        if cc_list:
            if isinstance(cc_list, (list, tuple)):
                recipients.extend(cc_list)
            else:
                recipients.extend([e.strip() for e in str(cc_list).split(",") if e.strip()])

        server = smtplib.SMTP(smtp_server, smtp_port, timeout=60)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, recipients, msg.as_string())
        server.quit()
        return True, "Email sent successfully"
    except Exception as e:
        return False, str(e)


def create_email_body(ship_name, failed_count, reasons_html=""):
    body = f"""
    <html>
      <body style="font-family: Arial, sans-serif; color: #333;">
        <h3>Vessel Report Validation Alert - {ship_name}</h3>
        <p>Failed reports: <b>{failed_count}</b></p>
        <p>Please find attached the failed validation report and take appropriate action.</p>
        <div>
          <h4>Common issues</h4>
          <ul>
            {reasons_html}
          </ul>
        </div>
        <p style="font-size:12px;color:gray;">Generated on {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
      </body>
    </html>
    """
    return body


# ---------------------------
# Streamlit UI
# ---------------------------
def main():
    st.set_page_config(page_title="Ship Report Validator", layout="wide")
    st.title("🚢 Ship Report Validation System")

    # Sidebar
    with st.sidebar:
        st.header("📘 Validation Rules")
        st.markdown(
            """
**Rule 1: SFOC (Specific Fuel Oil Consumption)**  
- At Sea (ME Rhrs > 12): 150–200 g/kWh

**Rule 2: Average Speed**  
- At Sea (ME Rhrs > 12): 0–20 knots

**Rule 3: Exhaust Temperature**  
- At Sea (ME Rhrs > 12): ±50°C deviation from avg (Units 1–16)

**Rule 4: ME Running Hours**  
- ME Rhrs must not exceed Report Hours by > 1 hour (±1h margin)

**Rule 5: Aux Engine Operation**  
- At Sea: (sum AE Rhrs / Report Hours) > 1.25 AND Average Load [%] > 40 AND sub-consumers = 0
"""
        )
        st.divider()
        st.header("📧 Email Configuration")
        smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com")
        smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
        sender_email = st.text_input("Sender Email")
        sender_password = st.text_input("Sender App Password", type="password", help="Use app password for Gmail")

    # File upload
    uploaded_file = st.file_uploader("Upload Excel (sheet: All Reports)", type=["xlsx", "xls"])
    if not uploaded_file:
        st.info("Upload the weekly data Excel (sheet name: 'All Reports') to begin.")
        return

    try:
        df = pd.read_excel(uploaded_file, sheet_name="All Reports")
    except Exception as e:
        st.error("Failed to read Excel. Ensure sheet is named 'All Reports'.")
        st.text(traceback.format_exc())
        return

    # Run validations
    failed, df_with_calcs = validate_reports(df.copy())

    # Summary metrics
    total_reports = len(df)
    failed_count = len(failed)
    pass_rate = ((total_reports - failed_count) / total_reports * 100) if total_reports else 0.0

    st.markdown("## Validation Results")
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Reports", total_reports)
    c2.metric("Failed Reports", failed_count)
    c3.metric("Pass Rate", f"{pass_rate:.1f}%")

    if failed_count > 0:
        st.warning(f"⚠️ {failed_count} reports failed validation")

        # Full table (all rows) with horizontal scroll — highlight all rows
        def highlight_all(row):
            return ["background-color: #f8d7da; color: black"] * len(row)

        st.subheader("Failed Reports (full table)")
        # Use styler and safe rendering fallback
        try:
            styled = failed.style.apply(highlight_all, axis=1)
            st.write(styled)  # preferred
        except Exception:
            try:
                html = failed.style.apply(highlight_all, axis=1).to_html()
                st.markdown("<div style='overflow:auto;'>" + html + "</div>", unsafe_allow_html=True)
            except Exception:
                st.dataframe(failed, use_container_width=True)

        # Failure reasons summary chart
        with st.expander("📊 Failure Reasons Summary"):
            all_reasons = []
            for r in failed["Reason"].astype(str).tolist():
                if r:
                    parts = [p.strip() for p in r.split(";") if p.strip()]
                    all_reasons.extend(parts)
            if all_reasons:
                reason_counts = pd.Series(all_reasons).value_counts()
                st.bar_chart(reason_counts)
                st.write(reason_counts)
            else:
                st.write("No detailed reasons parsed.")

        # Download failed and full data
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            out_failed = io.BytesIO()
            with pd.ExcelWriter(out_failed, engine="openpyxl") as writer:
                failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
            out_failed.seek(0)
            st.download_button("📥 Download Failed Reports", data=out_failed, file_name="Failed_Validation.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        with col_dl2:
            out_all = io.BytesIO()
            with pd.ExcelWriter(out_all, engine="openpyxl") as writer:
                df_with_calcs.to_excel(writer, index=False, sheet_name="All_Reports_Processed")
            out_all.seek(0)
            st.download_button("📥 Download All Data (with calculations)", data=out_all, file_name="All_Reports_With_Calculations.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("---")

        # Individual email sending
        st.subheader("📧 Send Email to Specific Vessel")
        if "Ship Name" in failed.columns:
            vessels = failed["Ship Name"].dropna().unique().tolist()
            selected_vessel = st.selectbox("Select Vessel", vessels)
            to_input = st.text_input("To (comma-separated)", placeholder="vessel@company.com")
            cc_input = st.text_input("CC (optional, comma-separated)")
            send_individual = st.button("Send Email to Selected Vessel")
            if send_individual:
                if not (smtp_server and sender_email and sender_password):
                    st.error("Fill SMTP settings in the sidebar before sending.")
                elif not to_input:
                    st.error("Enter recipient email(s).")
                else:
                    vessel_failed = failed[failed["Ship Name"] == selected_vessel]
                    # reasons summary
                    vessel_reasons = []
                    for r in vessel_failed["Reason"].astype(str):
                        if r:
                            parts = [p.strip() for p in r.split(";") if p.strip()]
                            vessel_reasons.extend(parts)
                    reasons_html = "".join(f"<li>{x}</li>" for x in pd.Series(vessel_reasons).value_counts().index.tolist())
                    body = create_email_body(selected_vessel, len(vessel_failed), reasons_html)
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                        vessel_failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
                    buffer.seek(0)
                    recipients = [e.strip() for e in to_input.split(",") if e.strip()]
                    cc_list = [e.strip() for e in cc_input.split(",") if e.strip()] if cc_input else None
                    ok, msg = send_email(smtp_server, smtp_port, sender_email, sender_password, recipients, f"Validation Alert - {selected_vessel}", body, buffer, f"{selected_vessel}_Failed_Validation.xlsx", cc_list)
                    if ok:
                        st.success(msg)
                    else:
                        st.error(msg)

        else:
            st.info("No 'Ship Name' column — cannot use vessel-specific emailing.")

        st.markdown("---")

        # Bulk email sending
        st.subheader("📦 Bulk Email Sending")
        st.markdown("Upload an email mapping file (must contain 'Ship Name' and 'Email' columns). Optional CC columns allowed (e.g., 'CC').")
        email_map_file = st.file_uploader("Upload Vessel Email Mapping (Excel/CSV)", type=["xlsx", "xls", "csv"], key="bulk_map")
        if email_map_file is not None:
            try:
                if str(email_map_file.name).lower().endswith(".csv"):
                    email_df = pd.read_csv(email_map_file)
                else:
                    email_df = pd.read_excel(email_map_file)
            except Exception as e:
                st.error("Failed to read mapping file.")
                st.text(traceback.format_exc())
                email_df = None

            if email_df is not None:
                st.dataframe(email_df.head(), use_container_width=True)
                if "Ship Name" not in email_df.columns or ("Email" not in email_df.columns and "To" not in email_df.columns):
                    st.error("Mapping file must have 'Ship Name' and 'Email' (or 'To') columns.")
                else:
                    email_col = "Email" if "Email" in email_df.columns else "To"
                    cc_cols = [c for c in email_df.columns if c.upper().startswith("CC")]
                    send_bulk = st.button("Send Bulk Emails")
                    if send_bulk:
                        if not (smtp_server and sender_email and sender_password):
                            st.error("Fill SMTP settings in the sidebar before sending.")
                        else:
                            progress = st.progress(0)
                            status_box = st.empty()
                            results = []
                            vessels = failed["Ship Name"].dropna().unique().tolist()
                            total = len(vessels)
                            for i, vessel in enumerate(vessels, start=1):
                                row = email_df[email_df["Ship Name"] == vessel]
                                if row.empty:
                                    results.append(f"❌ {vessel}: No mapping found")
                                    progress.progress(i / total)
                                    continue
                                to_val = row.iloc[0][email_col]
                                if pd.isna(to_val) or str(to_val).strip() == "":
                                    results.append(f"❌ {vessel}: No email present")
                                    progress.progress(i / total)
                                    continue
                                cc_list = []
                                for cc in cc_cols:
                                    val = row.iloc[0].get(cc)
                                    if pd.notna(val) and str(val).strip():
                                        cc_list.extend([e.strip() for e in str(val).split(",") if e.strip()])
                                cc_list = cc_list if cc_list else None

                                vessel_failed = failed[failed["Ship Name"] == vessel]
                                buffer = io.BytesIO()
                                with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                                    vessel_failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
                                buffer.seek(0)

                                # reasons summary
                                vessel_reasons = []
                                for r in vessel_failed["Reason"].astype(str):
                                    if r:
                                        vessel_reasons.extend([p.strip() for p in r.split(";") if p.strip()])
                                reasons_html = "".join(f"<li>{x}</li>" for x in pd.Series(vessel_reasons).value_counts().index.tolist())

                                body = create_email_body(vessel, len(vessel_failed), reasons_html)
                                recipients = [e.strip() for e in str(to_val).split(",") if e.strip()]
                                ok, msg = send_email(smtp_server, smtp_port, sender_email, sender_password, recipients, f"Validation Alert - {vessel}", body, buffer, f"{vessel}_Failed_Validation.xlsx", cc_list)
                                if ok:
                                    results.append(f"✅ {vessel}: Email sent")
                                else:
                                    results.append(f"❌ {vessel}: {msg}")
                                progress.progress(i / total)

                            status_box.write("### Bulk send results")
                            for r in results:
                                status_box.write(r)

    else:
        st.success("🎉 All reports passed validation!")

    st.markdown("---")
    st.caption("Tip: If the failed table is wide, use the horizontal scrollbar to view all columns.")

if __name__ == "__main__":
    main()
