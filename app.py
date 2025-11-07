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


def calculate_report_hours(df):
    report_hours = []
    for idx, row in df.iterrows():
        try:
            start_date = pd.to_datetime(row.get("Start Date"), errors='coerce')
            end_date = pd.to_datetime(row.get("End Date"), errors='coerce')
            start_time = str(row.get("Start Time", "00:00:00")).strip()
            end_time = str(row.get("End Time", "00:00:00")).strip()
            time_shift = row.get("Time Shift", 0) or 0

            if pd.notna(start_date) and pd.notna(end_date):
                start_time_obj = pd.to_datetime(start_time, errors='coerce').time() if start_time else datetime.strptime("00:00:00", '%H:%M:%S').time()
                end_time_obj = pd.to_datetime(end_time, errors='coerce').time() if end_time else datetime.strptime("00:00:00", '%H:%M:%S').time()
                start_datetime = datetime.combine(start_date.date(), start_time_obj)
                end_datetime = datetime.combine(end_date.date(), end_time_obj)
                total_hours = (end_datetime - start_datetime).total_seconds() / 3600 + float(time_shift)
                report_hours.append(round(total_hours, 2))
            else:
                report_hours.append(0)
        except Exception:
            report_hours.append(0)
    return report_hours


def aux_engine_validation(df):
    ae_cols = [
        "A.E. 1 Last Report [Rhrs] (Aux Engine Unit 1)",
        "A.E. 2 Last Report [Rhrs] (Aux Engine Unit 2)",
        "A.E. 3 Last Report [Rhrs] (Aux Engine Unit 3)",
        "A.E. 4 Total [Rhrs] (Aux Engine Unit 4)",
        "A.E. 5 Last Report [Rhrs] (Aux Engine Unit 5)",
        "A.E. 6 Last Report [Rhrs] (Aux Engine Unit 6)"
    ]
    sub_cols = [
        "Tank Cleaning [MT]", "Cargo Transfer [MT]", "Maintaining Cargo Temp. [MT]",
        "Shaft Gen. Propulsion [MT]", "Raising Cargo Temp. [MT]", "Burning Sludge [MT]",
        "Ballast Transfer [MT]", "Fresh Water Prod. [MT]", "Others [MT]", "EGCS Consumption [MT]"
    ]

    for col in ae_cols + sub_cols + ["Average Load [%]", "Report Hours", "Report Type"]:
        if col not in df.columns:
            df[col] = 0

    df["AE_Total_Rhrs"] = df[ae_cols].sum(axis=1)
    df["Sub_Consumption_Total"] = df[sub_cols].sum(axis=1)

    condition = (
        (df["Report Type"].str.strip().eq("At Sea")) &
        (df["Report Hours"] > 0) &
        ((df["AE_Total_Rhrs"] / df["Report Hours"]) > 1.25) &
        (df["Average Load [%]"] > 40) &
        (df["Sub_Consumption_Total"] == 0)
    )

    if "Reason" not in df.columns:
        df["Reason"] = ""

    df.loc[condition, "Reason"] = df["Reason"].astype(str) + "; " + (
        "Two or more Aux Engines running at sea with ME Load > 40% and no sub-consumers reported. "
        "Please confirm operations and update relevant sub-consumption fields if applicable."
    )

    return df


def validate_reports(df):
    numeric_cols = [
        "Average Load [kW]", "ME Rhrs (From Last Report)", "Avg. Speed",
        "Fuel Cons. [MT] (ME Cons 1)", "Fuel Cons. [MT] (ME Cons 2)", "Fuel Cons. [MT] (ME Cons 3)", "Time Shift"
    ]
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].astype(str).str.replace(",", "").replace(["", "nan", "None"], np.nan), errors="coerce").fillna(0)

    df["Report Hours"] = calculate_report_hours(df)
    df["SFOC"] = ((df["Fuel Cons. [MT] (ME Cons 1)"] + df["Fuel Cons. [MT] (ME Cons 2)"] + df["Fuel Cons. [MT] (ME Cons 3)"]) * 1_000_000) / (
        df["Average Load [kW]"].replace(0, np.nan) * df["ME Rhrs (From Last Report)"].replace(0, np.nan))
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

        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")

        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")

        if report_hours > 0:
            hours_diff = ME_Rhrs - report_hours
            if hours_diff > 1.0:
                reason.append(f"ME Rhrs ({ME_Rhrs:.2f}) exceeds Report Hours ({report_hours:.2f}) by {hours_diff:.2f}h")
                fail_columns.update(["ME Rhrs (From Last Report)", "Report Hours"])

        reasons.append("; ".join(reason))

    df["Reason"] = reasons
    df = aux_engine_validation(df)

    failed = df[df["Reason"] != ""].copy()
    return failed, df


def main():
    st.set_page_config(page_title="Ship Report Validator", page_icon="🚢", layout="wide")
    st.title("🚢 Ship Report Validation System")
    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if uploaded_file:
        df = pd.read_excel(uploaded_file, sheet_name="All Reports")
        failed, df_with_calcs = validate_reports(df.copy())
        st.dataframe(failed if not failed.empty else df_with_calcs)

if __name__ == "__main__":
    main()
