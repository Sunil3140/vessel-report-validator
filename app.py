import streamlit as st
import pandas as pd
import numpy as np
import io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta

@st.cache_data
def calculate_report_hours_from_data(start_dates, end_dates, start_times, end_times, time_shifts):
    """Calculate Report Hours from Start Date/Time, End Date/Time and Time Shift"""
    report_hours = []
    
    for idx in range(len(start_dates)):
        try:
            # Get date and time components
            start_date = pd.to_datetime(start_dates[idx], errors='coerce')
            end_date = pd.to_datetime(end_dates[idx], errors='coerce')
            
            # Get time components (handle various formats)
            start_time = str(start_times[idx] if idx < len(start_times) else "00:00:00").strip()
            end_time = str(end_times[idx] if idx < len(end_times) else "00:00:00").strip()
            
            # Handle time shift (convert to hours)
            time_shift = time_shifts[idx] if idx < len(time_shifts) else 0
            if pd.isna(time_shift):
                time_shift = 0
            else:
                time_shift = float(time_shift)
            
            # Create datetime objects
            if pd.notna(start_date) and pd.notna(end_date):
                # Parse time strings
                try:
                    start_time_obj = pd.to_datetime(start_time, format='%H:%M:%S').time()
                except:
                    try:
                        start_time_obj = pd.to_datetime(start_time, format='%H:%M').time()
                    except:
                        start_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()
                
                try:
                    end_time_obj = pd.to_datetime(end_time, format='%H:%M:%S').time()
                except:
                    try:
                        end_time_obj = pd.to_datetime(end_time, format='%H:%M').time()
                    except:
                        end_time_obj = datetime.strptime("00:00:00", '%H:%M:%S').time()
                
                # Combine date and time
                start_datetime = datetime.combine(start_date.date(), start_time_obj)
                end_datetime = datetime.combine(end_date.date(), end_time_obj)
                
                # Calculate time difference
                time_diff = end_datetime - start_datetime
                hours_diff = time_diff.total_seconds() / 3600
                
                # Add time shift
                total_hours = hours_diff + time_shift
                
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
                
        except Exception as e:
            report_hours.append(0)
    
    return report_hours

def calculate_report_hours(df):
    """Calculate Report Hours from Start Date/Time, End Date/Time and Time Shift"""
    df = df.copy()
    start_dates = df.get("Start Date", pd.Series([None]*len(df))).tolist()
    end_dates = df.get("End Date", pd.Series([None]*len(df))).tolist()
    start_times = df.get("Start Time", pd.Series(["00:00:00"]*len(df))).tolist()
    end_times = df.get("End Time", pd.Series(["00:00:00"]*len(df))).tolist()
    time_shifts = df.get("Time Shift", pd.Series([0]*len(df))).tolist()
    
    return calculate_report_hours_from_data(
        tuple(start_dates), 
        tuple(end_dates), 
        tuple(start_times), 
        tuple(end_times), 
        tuple(time_shifts)
    )

@st.cache_data
def validate_reports(df):
    """Validate ship reports and return failed rows with reasons"""
    df = df.copy()
    
    # --- Clean numeric columns ---
    numeric_cols = [
        "Average Load [kW]",
        "ME Rhrs (From Last Report)",
        "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)",
        "Fuel Cons. [MT] (ME Cons 2)",
        "Fuel Cons. [MT] (ME Cons 3)",
        "Time Shift",
        "Average Load [%]",
        "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)",
        "A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)",
        "A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)",
        "A.E. 4 Total [Rhrs] (Aux Engine Unit 4)",
        "A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)",
        "A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)",
        "Tank Cleaning [MT]",
        "Cargo Transfer [MT]",
        "Maintaining Cargo Temp. [MT]",
        "Shaft Gen. Propulsion [MT]",
        "Raising Cargo Temp. [MT]",
        "Burning Sludge [MT]",
        "Ballast Transfer [MT]",
        "Fresh Water Prod. [MT]",
        "Others [MT]",
        "EGCS Consumption [MT]"
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
        report_hours = row.get("Report Hours", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)

        # --- Rule 1: SFOC (only for At Sea) ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150-200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")

        # --- Rule 2: Avg Speed (only for At Sea) ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0-20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")

        # --- Rule 3: Exhaust Temp deviation (Units 1-16, only At Sea) ---
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

        # --- Rule 4: ME Rhrs should not exceed Report Hours (with ±1 hour margin) ---
        if report_hours > 0:
            hours_diff = ME_Rhrs - report_hours
            if hours_diff > 1.0:
                reason.append(f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {hours_diff:.2f}h (margin: ±1h)")
                fail_columns.add("ME Rhrs (From Last Report)")
                fail_columns.add("Report Hours")

        # --- Rule 5: Multiple Aux Engines operating at sea without sub-consumers ---
        if report_type == "At Sea" and row.get("Average Load [%]", 0) > 40:
            # Sum all auxiliary engine running hours
            ae_rhrs_sum = (
                row.get("A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)", 0) +
                row.get("A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)", 0) +
                row.get("A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)", 0) +
                row.get("A.E. 4 Total [Rhrs] (Aux Engine Unit 4)", 0) +
                row.get("A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)", 0) +
                row.get("A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)", 0)
            )
            
            # Calculate AE running hours ratio
            if report_hours > 0:
                ae_ratio = ae_rhrs_sum / report_hours
            else:
                ae_ratio = 0
            
            # Sum all sub-consumers
            sub_consumers_sum = (
                row.get("Tank Cleaning [MT]", 0) +
                row.get("Cargo Transfer [MT]", 0) +
                row.get("Maintaining Cargo Temp. [MT]", 0) +
                row.get("Shaft Gen. Propulsion [MT]", 0) +
                row.get("Raising Cargo Temp. [MT]", 0) +
                row.get("Burning Sludge [MT]", 0) +
                row.get("Ballast Transfer [MT]", 0) +
                row.get("Fresh Water Prod. [MT]", 0) +
                row.get("Others [MT]", 0) +
                row.get("EGCS Consumption [MT]", 0)
            )
            
            # Check if 2+ Aux Engines operating (ratio > 1.25) with ME Load > 40% and no sub-consumers
            if ae_ratio > 1.25 and sub_consumers_sum == 0:
                reason.append(f"Multiple Aux Engines operating at sea (AE Rhrs/Report Hours = {ae_ratio:.2f}) with ME Load > 40% but no sub-consumers reported. Please confirm operations and update sub-consumption fields if applicable")
                fail_columns.add("Average Load [%]")
                fail_columns.add("A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)")
                fail_columns.add("A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)")
                fail_columns.add("A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)")
                fail_columns.add("Tank Cleaning [MT]")
                fail_columns.add("Cargo Transfer [MT]")

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
        "Report Hours",
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
               recipient_emails, subject, body, attachment_data=None, 
               attachment_name="Failed_Validation.xlsx", cc_emails=None):
    """Send email with optional attachment to multiple recipients"""
    try:
        msg = MIMEMultipart()
        msg['From'] = sender_email
        
        # Handle recipient emails
        if isinstance(recipient_emails, str):
            recipient_list = [email.strip() for email in recipient_emails.split(',') if email.strip()]
        else:
            recipient_list = recipient_emails
        
        msg['To'] = ', '.join(recipient_list)
        
        # Handle CC emails
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
        
        # Attach file if provided
        if attachment_data:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment_data.getvalue())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename={attachment_name}')
            msg.attach(part)
        
        # Combine To and CC for actual sending
        all_recipients = recipient_list + cc_list
        
        # Connect and send
        server = smtplib.SMTP(smtp_server, smtp_port)
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
                This is an automated message. Generated on {datetime.now().strftime('%Y-%m-%d %H:%M:%S UTC')}
            </p>
        </body>
    </html>
    """
    return body

@st.cache_data
def process_excel_file(file_bytes, file_name):
    """Process uploaded Excel file and return validation results"""
    df = pd.read_excel(io.BytesIO(file_bytes), sheet_name="All Reports")
    failed, df_with_calcs = validate_reports(df.copy())
    return df, failed, df_with_calcs


def main():
    st.set_page_config(
        page_title="Ship Report Validator",
        page_icon="🚢",
        layout="wide"
    )
    
    # Initialize session state
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
    
    # Sidebar with validation rules and email settings
    with st.sidebar:
        st.header("📋 Validation Rules")
        st.markdown("""
        **Rule 1: SFOC (Specific Fuel Oil Consumption)**
        - At Sea (ME Rhrs > 12): 150–200 g/kWh
        - At Port/Anchorage: No validation
        
        **Rule 2: Average Speed**
        - At Sea (ME Rhrs > 12): 0–20 knots
        - At Port/Anchorage: No validation
        
        **Rule 3: Exhaust Temperature**
        - At Sea (ME Rhrs > 12): Deviation ≤ ±50°C from average
        - Applies to Units 1-16
        - At Port/Anchorage: No validation
        
        **Rule 4: ME Running Hours**
        - ME Rhrs must not exceed Report Hours by more than 1 hour
        - Tolerance: ±1 hour margin
        
        **Rule 5: Auxiliary Engines & Sub-Consumers**
        - At Sea with ME Load > 40%
        - If AE Rhrs/Report Hours > 1.25 (indicating 2+ AEs running)
        - All sub-consumers must not be zero
        - Validates proper reporting of tank cleaning, cargo operations, etc.
        
        **Report Hours Calculation**
        - Calculated as: (End Date/Time - Start Date/Time) + Time Shift
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
    
    # Reset validation when new file is uploaded
    if uploaded_file is not None:
        # Create a unique identifier for the file
        file_id = f"{uploaded_file.name}_{uploaded_file.size}"
        
        # Check if this is a new file
        if 'current_file_id' not in st.session_state or st.session_state.current_file_id != file_id:
            st.session_state.current_file_id = file_id
            st.session_state.validation_done = False
            st.session_state.failed_df = None
            st.session_state.df_with_calcs = None
            st.session_state.original_df = None
    
    # Run validation only once when file is uploaded
    if uploaded_file is not None and not st.session_state.validation_done:
        try:
            # Read file bytes for caching
            file_bytes = uploaded_file.read()
            file_name = uploaded_file.name
            
            # Process file with caching
            with st.spinner("Loading and validating file..."):
                df, failed, df_with_calcs = process_excel_file(file_bytes, file_name)
                
                # Store in session state
                st.session_state.original_df = df
                st.session_state.failed_df = failed
                st.session_state.df_with_calcs = df_with_calcs
                st.session_state.validation_done = True
            
            st.success(f"✅ File loaded and validated! Total rows: {len(df)}")
            
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")
            st.exception(e)
    
    # Display results if validation is done
    if st.session_state.validation_done:
        df = st.session_state.original_df
        failed = st.session_state.failed_df
        df_with_calcs = st.session_state.df_with_calcs
        
        # Show column info
        with st.expander("📊 Dataset Information"):
            st.write(f"**Rows:** {len(df)}")
            st.write(f"**Columns:** {len(df.columns)}")
            st.write("**Column Names:**")
            st.write(df.columns.tolist())
        
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
                    
                    st.markdown("**Recipient Emails** (comma-separated for multiple)")
                    vessel_email = st.text_area("To:", 
                                                 placeholder="vessel1@company.com, vessel2@company.com",
                                                 key="single_vessel_email",
                                                 height=80)
                    
                    st.markdown("**CC Emails** (optional, comma-separated)")
                    vessel_cc = st.text_area("CC:", 
                                              placeholder="manager@company.com, office@company.com",
                                              key="single_vessel_cc",
                                              height=80)
                    
                    if st.button("📤 Send Email to Selected Vessel", type="primary"):
                        if not sender_email or not sender_password:
                            st.error("Please configure SMTP settings in the sidebar")
                        elif not vessel_email:
                            st.error("Please enter at least one recipient email address")
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
                                    f"Failed_Validation_{selected_vessel}.xlsx",
                                    cc_emails=vessel_cc if vessel_cc else None
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
                        help="File should have columns: 'Ship Name', 'Email' (or 'To'), and optionally 'CC1', 'CC2', 'CC3', etc.",
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
                            
                            # Check for required columns
                            if "Ship Name" not in email_df.columns:
                                st.error("❌ Email mapping file must have 'Ship Name' column")
                            elif "Email" not in email_df.columns and "To" not in email_df.columns:
                                st.error("❌ Email mapping file must have 'Email' or 'To' column")
                            else:
                                # Determine the email column name
                                email_col = "Email" if "Email" in email_df.columns else "To"
                                
                                # Detect CC columns
                                cc_columns = [col for col in email_df.columns if col.upper().startswith('CC')]
                                if cc_columns:
                                    st.info(f"📧 Found CC columns: {', '.join(cc_columns)}")
                                
                                if st.button("📨 Send Emails to All Vessels", type="primary"):
                                    if not sender_email or not sender_password:
                                        st.error("Please configure SMTP settings in the sidebar")
                                    else:
                                        progress_bar = st.progress(0)
                                        status_container = st.container()
                                        
                                        results = []
                                        for idx, vessel in enumerate(vessels):
                                            # Get vessel email row
                                            vessel_email_row = email_df[email_df["Ship Name"] == vessel]
                                            
                                            if vessel_email_row.empty:
                                                results.append(f"❌ {vessel}: No email found in mapping")
                                                continue
                                            
                                            # Get primary email(s)
                                            vessel_email = vessel_email_row.iloc[0][email_col]
                                            
                                            # Skip if no email
                                            if pd.isna(vessel_email) or str(vessel_email).strip() == "":
                                                results.append(f"❌ {vessel}: Email is empty")
                                                continue
                                            
                                            # Collect all CC emails from CC columns
                                            cc_emails_list = []
                                            for cc_col in cc_columns:
                                                cc_val = vessel_email_row.iloc[0].get(cc_col)
                                                if pd.notna(cc_val) and str(cc_val).strip():
                                                    cc_emails_list.extend([e.strip() for e in str(cc_val).split(',') if e.strip()])
                                            
                                            # Combine CC emails
                                            cc_emails_str = ', '.join(cc_emails_list) if cc_emails_list else None
                                            
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
                                                f"Failed_Validation_{vessel}.xlsx",
                                                cc_emails=cc_emails_str
                                            )
                                            
                                            if success:
                                                cc_info = f" (CC: {len(cc_emails_list)} recipients)" if cc_emails_list else ""
                                                results.append(f"✅ {vessel}: Email sent successfully{cc_info}")
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
        
        # Option to view all data with SFOC and Report Hours
        with st.expander("🔍 View All Data (with calculated SFOC and Report Hours)"):
            st.dataframe(df_with_calcs, use_container_width=True, height=400)
            
            # Download all data
            output_all = io.BytesIO()
            with pd.ExcelWriter(output_all, engine='openpyxl') as writer:
                df_with_calcs.to_excel(writer, index=False, sheet_name="All_Reports_Processed")
            output_all.seek(0)
            
            st.download_button(
                label="📥 Download All Data with Calculations",
                data=output_all,
                file_name="All_Reports_With_Calculations.xlsx",
                mime="application/vnd.openxmlx-officedocument.spreadsheetml.sheet"
            )
    
    elif uploaded_file is None:
        st.info("👆 Please upload an Excel file to begin validation")
        
        # Show sample data structure
        with st.expander("📄 Expected Data Structure"):
            st.markdown("""
            **Main Excel File** should contain a sheet named **"All Reports"** with columns:
            
            - Ship Name, IMO_No, Report Type (At Sea / At Port / At Anchorage)
            - Start Date, Start Time, End Date, End Time, Time Shift
            - Average Load [kW], ME Rhrs (From Last Report), Avg. Speed
            - Fuel Cons. [MT] (ME Cons 1, 2, 3)
            - Exh. Temp [°C] (Main Engine Unit 1-16)
            - A.E. 1-6 Last Report [Rhrs] (Aux Engine Units)
            - Sub-consumer fields: Tank Cleaning, Cargo Transfer, etc.
            
            **Email Mapping File** (for bulk sending):
            - Must have columns: `Ship Name` and `Email` (or `To`)
            - Optional CC columns: `CC1`, `CC2`, `CC3`, etc.
            - You can add multiple emails in one cell using commas
            - Example:
            
            | Ship Name | Email | CC1 | CC2 |
            |-----------|-------|-----|-----|
            | Vessel A  | captain@vessel-a.com, chief@vessel-a.com | manager@company.com | office@company.com |
            | Vessel B  | vesselb@company.com | supervisor@company.com | admin@company.com |
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
