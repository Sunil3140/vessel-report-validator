import streamlit as st
import pandas as pd
import numpy as np
import io

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

    cols_to_keep = context_cols + exhaust_cols + list(fail_columns) + ["Reason"]
    cols_to_keep = [c for c in cols_to_keep if c in failed.columns]

    # Move Ship Name to Column A
    if "Ship Name" in cols_to_keep:
        cols_to_keep.remove("Ship Name")
        cols_to_keep = ["Ship Name"] + cols_to_keep

    failed = failed[cols_to_keep]

    return failed, df


def main():
    st.set_page_config(
        page_title="Ship Report Validator",
        page_icon="🚢",
        layout="wide"
    )
    
    st.title("🚢 Ship Report Validation System")
    st.markdown("Upload your Excel file to validate ship reports against compliance rules")
    
    # Sidebar with validation rules
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
                
                # Download button for failed reports
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
                output.seek(0)
                
                st.download_button(
                    label="📥 Download Failed Reports",
                    data=output,
                    file_name="Failed_Validation.xlsx",
                    mime="application/vnd.openxmlx-officedocument.spreadsheetml.sheet"
                )
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
            The Excel file should contain a sheet named **"All Reports"** with the following columns:
            
            - Ship Name
            - IMO_No
            - Report Type (At Sea / At Port / At Anchorage)
            - Start Date, Start Time, End Date, End Time
            - Average Load [kW]
            - ME Rhrs (From Last Report)
            - Avg. Speed
            - Fuel Cons. [MT] (ME Cons 1, 2, 3)
            - Exh. Temp [°C] (Main Engine Unit 1-16)
            - And other standard ship report columns
            """)


if __name__ == "__main__":
    main()
