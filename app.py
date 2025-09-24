import pandas as pd
import numpy as np
 
# --- validation function
def validate_reports(df):
    reasons = []
    fail_columns = set()   # keep track of which columns failed
 
    for idx, row in df.iterrows():
        reason = []
        report_type = str(row.get("Report Type", "")).strip()
        ME_Rhrs = row.get("ME Rhrs (From Last Report)", 0)
        sfoc = row.get("SFOC", 0)
        avg_speed = row.get("Avg. Speed", 0)
 
        # --- Rule 1: SFOC
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (150 <= sfoc <= 200):
                reason.append("SFOC out of 150–200 at sea with ME Rhrs > 12")
                fail_columns.add("SFOC")
        elif report_type in ["At Port", "At Anchorage"]:
            if sfoc != 0:
                reason.append("SFOC not 0 at port/anchorage")
                fail_columns.add("SFOC")
 
        # --- Rule 2: Avg Speed
        if report_type == "At Sea" and ME_Rhrs > 12:
            if not (0 <= avg_speed <= 20):
                reason.append("Avg. Speed out of 0–20 at sea with ME Rhrs > 12")
                fail_columns.add("Avg. Speed")
        elif report_type == "At Port":
            if avg_speed != 0:
                reason.append("Avg. Speed not 0 at port")
                fail_columns.add("Avg. Speed")
 
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
                        fail_columns.add(c)
 
        # --- Rule 4: ME Rhrs always < 25
        if ME_Rhrs >= 25:
            reason.append("ME Rhrs >= 25")
            fail_columns.add("ME Rhrs (From Last Report)")
 
        reasons.append("; ".join(reason))
 
    df["Reason"] = reasons
    failed = df[df["Reason"] != ""].copy()
 
    # Always keep these context columns in this order
    context_cols = [
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
 
    cols_to_keep = context_cols + list(fail_columns) + ["Reason"]
 
    # Filter only existing columns
    cols_to_keep = [c for c in cols_to_keep if c in failed.columns]
 
    failed = failed[cols_to_keep]
 
    return failed
 
 
if __name__ == "__main__":
    file_path = "weekly-data-dump_13-Sep-25_20-Sep-25.xlsx"
 
    df = pd.read_excel(file_path, sheet_name="All Reports")
    print("✅ File loaded successfully")
 
    failed = validate_reports(df)
 
    print(f"Total Rows: {len(df)}")
    print(f"Failed Rows: {len(failed)}")
 
    if not failed.empty:
        output_file = "Failed_Validation.xlsx"
        failed.to_excel(output_file, index=False, sheet_name="Failed_Validation")
        print(f"❌ Some rows failed validation. Saved to {output_file}")
        print("\n--- Failed Rows Preview ---")
        print(failed.head(10))  # print first 10 failed rows in console
    else:
        print("🎉 All rows passed validation!")
