import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

st.set_page_config(page_title="Vessel Report Validator", layout="wide")

# --- Validation logic (same as we tested) ---
def validate_reports(df):
    # Clean numeric columns
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

    # Calculate SFOC in g/kWh
    df["SFOC"] = (
        (
            df["Fuel Cons. [MT] (ME Cons 1)"]
            + df["Fuel Cons. [MT] (ME Cons 2)"]
            + df["Fuel Cons. [MT] (ME Cons 3)"]
        )
        * 1_000_000
        / (
            df["Average Load [kW]"].replace(0, np.nan)
            * df["ME Rhrs (From Last Report)"].replace(0, np.nan)
        )
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

        # Rule 1: SFOC
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")
        elif report_type in ["At Port", "At Anchorage"]:
            if abs(sfoc) > 0.0001:
                reason.append("SFOC not 0 at port/anchorage")
                fail_columns.add("SFOC")

        # Rule 2: Avg Speed
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")
        elif report_type == "At Port":
            if abs(avg_speed) > 0.0001:
                reason.append("Avg. Speed not 0 at port")
                fail_columns.add("Avg. Speed")

        # Rule 3: Exhaust temps
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

        # Rule 4: ME Rhrs always < 25
        if ME_Rhrs > 25:
            reason.append("ME Rhrs > 25")
            fail_columns.add("ME Rhrs (From Last Report)")

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    failed = df[df["Reason"] != ""].copy()

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

    if "Ship Name" in cols_to_keep:
        cols_to_keep.remove("Ship Name")
        cols_to_keep = ["Ship Name"] + cols_to_keep

    failed = failed[cols_to_keep]
    return failed


# --- Streamlit UI ---
st.title("Vessel Report Validator")

st.markdown("Upload the Excel file (sheet name: 'All Reports') or use the default file path.")

uploaded_file = st.file_uploader("Drop Excel file here", type=["xlsx", "xls"])

use_default = st.checkbox("Use server filepath instead of upload", value=False)
server_path_input = st.text_input("Server file path (if using server file)", value="weekly-data-dump_31-Oct-25_07-Nov-25.xlsx")

if uploaded_file is None and not use_default:
    st.info("Waiting for file upload or enable 'Use server filepath' and provide path.")
    st.stop()

# Load and validate
try:
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, sheet_name="All Reports")
        st.success("File loaded from upload.")
    else:
        df = pd.read_excel(server_path_input, sheet_name="All Reports")
        st.success(f"File loaded from server path: {server_path_input}")

    st.write("Total rows in sheet:", len(df))

    with st.spinner("Running validation..."):
        failed = validate_reports(df)

    st.write("Failed rows:", len(failed))

    if not failed.empty:
        st.subheader("Preview — first 20 failed rows")
        st.dataframe(failed.head(20), use_container_width=True)

        # Provide download
        towrite = BytesIO()
        with pd.ExcelWriter(towrite, engine="openpyxl") as writer:
            failed.to_excel(writer, index=False, sheet_name="Failed_Validation")
        towrite.seek(0)
        st.download_button(
            label="Download Failed_Validation.xlsx",
            data=towrite,
            file_name="Failed_Validation.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.success("All rows passed validation!")

except Exception as e:
    st.error("An error occurred while processing the file.")
    st.exception(e)
    st.stop()
