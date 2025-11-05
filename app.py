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

def validate_reports(df):
    """Validate ship reports and return failed rows with reasons"""
    
    # --- Clean numeric columns ---
    numeric_cols = [
        "Average Load [kW]",
        "ME Rhrs (From Last Report)",
        "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)",
        "Fuel Cons. [MT] (ME Cons 2)",
        "Fuel Cons. [MT] (ME Cons 3)"
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

    # --- Calculate SFOC in g/kWh ---
    df["SFOC"] = (
        (
            df["Fuel Cons. [MT] (ME Cons 1)"]
            + df["Fuel Cons. [MT] (ME Cons 2)"]
            + df["Fuel Cons. [MT] (ME Cons 3)"]
        )
        * 1_000_000
        / (df["Average Load [kW]"].replace(0, np.nan)
           * df["ME Rhrs (From Last Report)"].replace(0, np.nan))
    )
    df["SFOC"] = df["SFOC"].fillna(0)

    reasons = []
    fail_columns = set()

    for idx, row in df.iterrows():
        reason = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        # --- Rule 1: SFOC ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")
        elif report_type in ["At Port", "At Anchorage"]:
            if abs(sfoc) > 0.0001:
                reason.append("SFOC not 0 at port/anchorage")
                fail_columns.add("SFOC")

        # --- Rule 2: Avg Speed ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")
        elif report_type == "At Port":
            if abs(avg_speed) > 0.0001:
                reason.append("Avg. Speed not 0 at port")
                fail_columns.add("Avg. Speed")

        # --- Rule 3: Exhaust Temp deviation (Units 1–16) ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            exhaust_cols = [
                f"Exh. Temp [°C] (Main Engine Unit {j})"
                for j in range(1, 17)
                if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns
            ]
            temps = [row[c] for c in exhaust_cols if pd.notna(row[c]) and row[c] != 0]
            if temps:
                avg_temp = np.mean(temps)
                for j, c in enumerate(exhaust_cols, start=1):
                    val = row[c]
                    if pd.notna(val) and val != 0 and abs(val - avg_temp) > 50:
                        reason.append(f"Exhaust temp deviation > ±50 from avg at Unit {j}")
                        fail_columns.add(c)

        # --- Rule 4: ME Rhrs always < 25 ---
        if ME_Rhrs > 25:
            reason.append("ME Rhrs > 25")
            fail_columns.add("ME Rhrs (From Last Report)")

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    failed = df[df["Reason"] != ""].copy()

    # --- Always include Ship Name and Exhaust Temp columns ---
    exhaust_cols = [
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
    ]

    # Combine all columns and remove duplicates while preserving order
    cols_to_keep = context_cols + exhaust_cols + list(fail_columns) + ["Reason"]
    
    # Remove duplicates while preserving order
    seen = set()
    cols_to_keep_unique = []
    for col in cols_to_keep:
        if col not in seen and col in failed.columns:
            seen.add(col)
            cols_to_keep_unique.append(col)
    
    # Move Ship Name to Column A
    if "Ship Name" in cols_to_keep_unique:
        cols_to_keep_unique.remove("Ship Name")
        cols_to_keep_unique = ["Ship Name"] + cols_to_keep_unique

    failed = failed[cols_to_keep_unique]

    return failed, df


def send_email(smtp_server, smtp_port, sender_email, sender_password, 
               recipient_email, subject, body, attachment_data=None, 
               attachment_name="Failed_Validation.xlsx"):
    """Send email with optional attachment"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        msg['To'] = recipient_email
        msg['Subject'] = subject
        
        msg.attach(MIMEText(body, 'html'))
        
        # Attach file if provided
        if attachment_data:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_data.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
            msg.attach(part)
        
        # Connect and send
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(msg)
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
            
            <p>Dear Captain and Crew of <strong>{ship_name}</strong>,</p>
            
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
                <strong>Validation Rules Reference:</strong><br>
                • SFOC: 150-200 g/kWh at sea (ME Rhrs > 12), 0 at port/anchorage<br>
                • Speed: 0-20 knots at sea (ME Rhrs > 12), 0 at port<br>
                • Exhaust Temp: Deviation ≤ ±50°C from average<br>
                • ME Rhrs: Must be < 25 hours
            </p>
            
            <p style="color: #7f8c8d; font-size: 0.85em; margin-top: 30px;">
                This is an automated message. Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}
            </p>
        </body>
    </html>
    """
    return body


def main():
    st.set_page_config(
        page_title="Ship Report Validator",
        page_icon="🚢",
        layout="wide"
    )
    
    st.title("🚢 Ship Report Validation System")
    st.markdown("Upload your Excel file to validate ship reports and send automated alerts")
    
    # Sidebar with validation rules and email settings
    with st.sidebar:
        st.header("📋 Validation Rules")
        st.markdown("""
        **Rule 1: SFOC (Specific Fuel Oil Consumption)**
        - At Sea (ME Rhrs > 12): 150–200 g/kWh
        - At Port/Anchorage: Must be 0
        
        **Rule 2: Average Speed**
        - At Sea (ME Rhrs > 12): 0–20 knots
        - At Port: Must be 0
        
        **Rule 3: Exhaust Temperature**
        - Deviation must be ≤ ±50°C from average
        - Applies to Units 1-16
        
        **Rule 4: ME Running Hours**
        - Must be < 25 hours
        """)
        
        st.divider()
        
        st.header("📧 Email Configuration")
        with st.expander("SMTP Settings", expanded=False):
            smtp_server = st.text_input("SMTP Server", value="smtp.gmail.com", 
                                       help="e.g., smtp.gmail.com, smtp.office365.com")
            smtp_port = st.number_input("SMTP Port", value=587, min_value=1, max_value=65535)
            sender_email = st.text_input("Sender Email", placeholder="your-email@company.com")
            sender_password = st.text_input("Password", type="password", 
                                           help="Use App Password for Gmail")
    
    # File uploader
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=["xlsx", "xls"],
        help="Upload the weekly data dump Excel file"
    )
    
    if uploaded_file is not None:
        try:
            # Read Excel file
            with st.spinner("Loading file..."):
                df = pd.read_excel(uploaded_file, sheet_name="All Reports")
            
            st.success(f"✅ File loaded successfully! Total rows: {len(df)}")
            
            # Show column info
            with st.expander("📊 Dataset Information"):
                st.write(f"**Rows:** {len(df)}")
                st.write(f"**Columns:** {len(df.columns)}")
                st.write("**Column Names:**")
                st.write(df.columns.tolist())
            
            # Validate reports
            with st.spinner("Validating reports..."):
                failed, df_with_sfoc = validate_reports(df.copy())
            
            # Display results
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
                
                # Show failed reports
                st.subheader("Failed Reports")
                st.dataframe(failed, use_container_width=True, height=400)
                
                # Failure reasons summary
                with st.expander("📊 Failure Reasons Summary"):
                    reasons_list = []
                    for reason_str in failed["Reason"]:
                        if reason_str:
                            reasons_list.extend(reason_str.split("; "))
                    
                    if reasons_list:
                        reason_counts = pd.Series(reasons_list).value_counts()
                        st.bar_chart(reason_counts)
                        st.write(reason_counts)
                
                # Create Excel file for download/email
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
                output.seek(0)
                
                # Download button
                st.download_button(
                    label="📥 Download Failed Reports",
                    data=output,
                    file_name="Failed_Validation.xlsx",
                    mime="application/vnd.openxmlx-officedocument.spreadsheetml.sheet"
                )
                
                # Email Section
                st.divider()
                st.header("📧 Send Email Notifications")
                
                # Get unique vessels
                if "Ship Name" in failed.columns:
                    vessels = failed["Ship Name"].unique()
                    
                    tab1, tab2 = st.tabs(["📤 Send to Specific Vessels", "📨 Bulk Send to All"])
                    
                    with tab1:
                        st.markdown("### Send validation report to specific vessels")
                        
                        selected_vessel = st.selectbox("Select Vessel", vessels)
                        vessel_email = st.text_input("Vessel Email Address", 
                                                     placeholder="vessel@company.com",
                                                     key="single_vessel_email")
                        
                        if st.button("📤 Send Email to Selected Vessel", type="primary"):
                            if not sender_email or not sender_password:
                                st.error("Please configure SMTP settings in the sidebar")
                            elif not vessel_email:
                                st.error("Please enter vessel email address")
                            else:
                                # Filter failed reports for this vessel
                                vessel_failed = failed[failed["Ship Name"] == selected_vessel]
                                
                                # Create vessel-specific Excel
                                vessel_output = io.BytesIO()
                                with pd.ExcelWriter(vessel_output, engine='openpyxl') as writer:
                                    vessel_failed.to_excel(writer, index=False, 
                                                          sheet_name="Failed_Validation")
                                vessel_output.seek(0)
                                
                                # Prepare reasons summary
                                vessel_reasons = []
                                for reason_str in vessel_failed["Reason"]:
                                    if reason_str:
                                        vessel_reasons.extend(reason_str.split("; "))
                                
                                reasons_html = ""
                                if vessel_reasons:
                                    reason_counts = pd.Series(vessel_reasons).value_counts()
                                    for reason, count in reason_counts.items():
                                        reasons_html += f"<li>{reason} ({count} occurrence{'s' if count > 1 else ''})</li>\n"
                                
                                # Create email
                                subject = f"Vessel Report Validation Alert - {selected_vessel}"
                                body = create_email_body(selected_vessel, len(vessel_failed), reasons_html)
                                
                                with st.spinner("Sending email..."):
                                    success, message = send_email(
                                        smtp_server, smtp_port, sender_email, sender_password,
                                        vessel_email, subject, body, vessel_output,
                                        f"Failed_Validation_{selected_vessel}.xlsx"
                                    )
                                
                                if success:
                                    st.success(f"✅ {message}")
                                else:
                                    st.error(f"❌ {message}")
                    
                    with tab2:
                        st.markdown("### Send validation reports to all vessels with failures")
                        
                        st.info(f"📊 {len(vessels)} vessel(s) have validation failures")
                        
                        # Upload vessel email mapping
                        email_mapping_file = st.file_uploader(
                            "Upload Vessel Email Mapping (Excel/CSV)",
                            type=["xlsx", "xls", "csv"],
                            help="File should have columns: 'Ship Name' and 'Email'",
                            key="email_mapping"
                        )
                        
                        if email_mapping_file:
                            try:
                                if email_mapping_file.name.endswith('.csv'):
                                    email_df = pd.read_csv(email_mapping_file)
                                else:
                                    email_df = pd.read_excel(email_mapping_file)
                                
                                st.success(f"✅ Loaded {len(email_df)} vessel email mappings")
                                st.dataframe(email_df.head(), use_container_width=True)
                                
                                if st.button("📨 Send Emails to All Vessels", type="primary"):
                                    if not sender_email or not sender_password:
                                        st.error("Please configure SMTP settings in the sidebar")
                                    else:
                                        progress_bar = st.progress(0)
                                        status_container = st.container()
                                        
                                        results = []
                                        for idx, vessel in enumerate(vessels):
                                            # Get vessel email
                                            vessel_email_row = email_df[email_df["Ship Name"] == vessel]
                                            
                                            if vessel_email_row.empty:
                                                results.append(f"❌ {vessel}: No email found in mapping")
                                                continue
                                            
                                            vessel_email = vessel_email_row.iloc[0]["Email"]
                                            
                                            # Filter and create report
                                            vessel_failed = failed[failed["Ship Name"] == vessel]
                                            vessel_output = io.BytesIO()
                                            with pd.ExcelWriter(vessel_output, engine='openpyxl') as writer:
                                                vessel_failed.to_excel(writer, index=False, 
                                                                      sheet_name="Failed_Validation")
                                            vessel_output.seek(0)
                                            
                                            # Prepare reasons
                                            vessel_reasons = []
                                            for reason_str in vessel_failed["Reason"]:
                                                if reason_str:
                                                    vessel_reasons.extend(reason_str.split("; "))
                                            
                                            reasons_html = ""
                                            if vessel_reasons:
                                                reason_counts = pd.Series(vessel_reasons).value_counts()
                                                for reason, count in reason_counts.items():
                                                    reasons_html += f"<li>{reason} ({count} occurrence{'s' if count > 1 else ''})</li>\n"
                                            
                                            # Send email
                                            subject = f"Vessel Report Validation Alert - {vessel}"
                                            body = create_email_body(vessel, len(vessel_failed), reasons_html)
                                            
                                            success, message = send_email(
                                                smtp_server, smtp_port, sender_email, sender_password,
                                                vessel_email, subject, body, vessel_output,
                                                f"Failed_Validation_{vessel}.xlsx"
                                            )
                                            
                                            if success:
                                                results.append(f"✅ {vessel}: Email sent successfully")
                                            else:
                                                results.append(f"❌ {vessel}: {message}")
                                            
                                            progress_bar.progress((idx + 1) / len(vessels))
                                        
                                        with status_container:
                                            st.subheader("Email Sending Results")
                                            for result in results:
                                                st.write(result)
                                
                            except Exception as e:
                                st.error(f"Error loading email mapping: {str(e)}")
                        else:
                            st.info("👆 Upload a vessel email mapping file to enable bulk sending")
                else:
                    st.warning("⚠️ 'Ship Name' column not found. Cannot send vessel-specific emails.")
            
            else:
                st.success("🎉 All reports passed validation!")
                st.balloons()
            
            # Option to view all data with SFOC
            with st.expander("🔍 View All Data (with calculated SFOC)"):
                st.dataframe(df_with_sfoc, use_container_width=True, height=400)
                
                # Download all data
                output_all = io.BytesIO()
                with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
                    df_with_sfoc.to_excel(writer, index=False, sheet_name="All_Reports_With_SFOC")
                output_all.seek(0)
                
                st.download_button(
                    label="📥 Download All Data with SFOC",
                    data=output_all,
                    file_name="All_Reports_With_SFOC.xlsx",
                    mime="application/vnd.openxmlx-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.exception(e)
    else:
        st.info("👆 Please upload an Excel file to begin validation")
        
        # Show sample data structure
        with st.expander("📝 Expected Data Structure"):
            st.markdown("""
            **Main Excel File** should contain a sheet named **"All Reports"** with columns:
            
            - Ship Name, IMO_No, Report Type (At Sea / At Port / At Anchorage)
            - Start Date, Start Time, End Date, End Time
            - Average Load [kW], ME Rhrs (From Last Report), Avg. Speed
            - Fuel Cons. [MT] (ME Cons 1, 2, 3)
            - Exh. Temp [°C] (Main Engine Unit 1-16)
            
            **Email Mapping File** (for bulk sending):
            - Must have columns: `Ship Name` and `Email`
            - Example:
            
            | Ship Name | Email |
            |-----------|-------|
            | Vessel A  | vessela@company.com |
            | Vessel B  | vesselb@company.com |
            """)
        
        with st.expander("📧 Email Setup Guide"):
            st.markdown("""
            ### Gmail Setup:
            1. Enable 2-Factor Authentication on your Google account
            2. Generate an App Password: [Google App Passwords](https://myaccount.google.com/apppasswords)
            3. Use these settings:
               - SMTP Server: `smtp.gmail.com`
               - SMTP Port: `587`
               - Your Gmail address as sender
               - App Password (not your regular password)
            
            ### Outlook/Office 365:
            - SMTP Server: `smtp.office365.com`
            - SMTP Port: `587`
            - Use your Office 365 credentials
            
            ### Other Email Providers:
            - Check your email provider's SMTP settings
            - Most use port 587 with TLS encryption
            """)


if __name__ == "__main__":
    main()
