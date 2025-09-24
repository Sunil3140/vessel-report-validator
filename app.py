import pandas as pd
import numpy as np
import streamlit as st
from io import BytesIO

st.set_page_config(page_title="Vessel Report Validator", layout="wide")

st.title("🚢 Vessel Report Validator")
st.write("Upload your Excel report (same format as your `.xls` file).")

# --- validation function
def validate_reports(df):
    reasons = []

    for idx, row in df.iterrows():
        reason = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)
        fuel_pr = row.get("Fuel Oil. Pr. [bar]", 0)

        # --- Rule 1: SFOC
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
        elif report_type in ["At Port", "At Anchorage"]:
            if sfoc != 0:
                reason.append("SFOC not 0 at port/anchorage")

        # --- Rule 2: Avg Speed
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
        elif report_type == "At Port":
            if avg_speed != 0:
                reason.append("Avg. Speed not 0 at port")

        # --- Rule 3: Exhaust Temp deviation (Units 1–16)
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

        # --- Rule 4: ME Rhrs always < 25
        if ME_Rhrs >= 25:
            reason.append("ME Rhrs >= 25")

        # --- Rule 5: Fuel Oil Pressure 6–8 bar
        if not (6 <= fuel_pr <= 8):
            reason.append("Fuel Oil Pressure out of 6–8 bar")

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    failed = df[df["Reason"] != ""].copy()
    return failed

# --- file upload
uploaded = st.file_uploader("📂 Upload Excel File", type=["xls", "xlsx", "xlsm"])

if uploaded:
    df = pd.read_excel(uploaded, sheet_name="All Reports")
    st.write("✅ File loaded successfully")

    failed = validate_reports(df)

    st.metric("Total Rows", len(df))
    st.metric("Failed Rows", len(failed))

    if len(failed) > 0:
        st.subheader("❌ Failed Validations")
        st.dataframe(failed, use_container_width=True)

        # prepare Excel download
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
        excel_data = output.getvalue()

        st.download_button(
            label="⬇️ Download Failed Validation",
            data=excel_data,
            file_name="Failed_Validation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("🎉 All rows passed validation!")
