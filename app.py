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
        # fuel_pr = row.get("Fuel Oil. Pr. [bar]", 0)   # <-- removed since not used

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

