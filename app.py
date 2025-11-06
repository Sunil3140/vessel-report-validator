import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


def validate_reports(df):
    """Validate ship reports, skipping At Port and adding Report Hours"""

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
                df[col].astype(str).str.replace(",", "").str.strip().replace(["", "nan", "None"], np.nan)
            )
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --- Calculate Report Hours ---
    df["Report Hours"] = (
        (pd.to_datetime(df["End Date"].astype(str) + " " + df["End Time"].astype(str))
         - pd.to_datetime(df["Start Date"].astype(str) + " " + df["Start Time"].astype(str)))
        .dt.total_seconds() / 3600
    ) + df["Time Shift"].fillna(0)

    # --- Calculate SFOC in g/kWh ---
    df["SFOC"] = (
        (df["Fuel Cons. [MT] (ME Cons 1)"] + df["Fuel Cons. [MT] (ME Cons 2)"] + df["Fuel Cons. [MT] (ME Cons 3)"])
        * 1_000_000
        / (df["Average Load [kW]"].replace(0, np.nan) * df["ME Rhrs (From Last Report)"].replace(0, np.nan))
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

        # Skip validation for At Port
        if report_type == "At Port":
            reasons.append("")
            continue

        # --- Rule 1: SFOC ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")
        elif report_type == "At Anchorage":
            if abs(sfoc) > 0.0001:
                reason.append("SFOC not 0 at anchorage")
                fail_columns.add("SFOC")

        # --- Rule 2: Avg Speed ---
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
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

    # --- Keep context columns ---
    context_cols = [
        "Ship Name", "IMO_No", "Report Type", "Start Date", "Start Time", "End Date", "End Time",
        "Report Hours", "Voyage Number", "Time Zone", "Distance - Ground [NM]", "Time Shift",
        "Distance - Sea [NM]", "Average Load [kW]", "Average RPM", "Average Load [%]", "ME Rhrs (From Last Report)"
    ]

    exhaust_cols = [
        f"Exh. Temp [°C] (Main Engine Unit {j})"
        for j in range(1, 17)
        if f"Exh. Temp [°C] (Main Engine Unit {j})" in df.columns
    ]

    cols_to_keep = context_cols + exhaust_cols + list(fail_columns) + ["Reason"]

    seen = set()
    cols_to_keep_unique = []
    for col in cols_to_keep:
        if col not in seen and col in failed.columns:
            seen.add(col)
            cols_to_keep_unique.append(col)

    if "Ship Name" in cols_to_keep_unique:
        cols_to_keep_unique.remove("Ship Name")
        cols_to_keep_unique = ["Ship Name"] + cols_to_keep_unique

    failed = failed[cols_to_keep_unique]

    return failed, df


# The rest of the Streamlit app (main, email handling, UI, etc.) remains unchanged.
