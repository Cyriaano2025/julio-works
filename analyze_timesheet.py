import streamlit as st
import pandas as pd
import re
from datetime import datetime

st.title("🧠 Smart Timesheet Analyzer")
st.write("Upload the standardized Excel timesheet — and this AI-powered tool will analyze it week by week.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

def parse_timesheet(sheet_df):
    records = []
    current_day = None

    for i, row in sheet_df.iterrows():
        first_cell = str(row[0]).strip()

        # Detect day start
        if re.match(r'^\w+ \d{2} \w+', first_cell, re.IGNORECASE):
            current_day = first_cell
            continue

        # Skip reporting/closing metadata
        if str(first_cell).lower().startswith("reporting time"):
            continue

        # Detect header row
        if str(first_cell).lower() in ['brand - task description', 'task description']:
            continue

        # Skip fully empty rows
        if row.isnull().all():
            continue

        # Valid row: Append
        records.append({
            "Day": current_day,
            "Task": row[0],
            "Start": row[1],
            "End": row[2],
            "Duration": row[3]
        })

    return pd.DataFrame(records)

def analyze_timesheet(df):
    try:
        df["Start"] = pd.to_datetime(df["Start"], format="%H:%M", errors="coerce")
        df["End"] = pd.to_datetime(df["End"], format="%H:%M", errors="coerce")

        df["Duration (hrs)"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600
        df["Duration (hrs)"] = df["Duration (hrs)"].fillna(df["Duration"])

        df["Date"] = df["Day"].apply(lambda x: pd.to_datetime(x.strip(), format="%A %d %B", errors="coerce"))
        df = df.dropna(subset=["Date"])

        weekly_summary = df.groupby("Date").agg({
            "Duration (hrs)": "sum"
        }).rename(columns={"Duration (hrs)": "Total Hours"}).reset_index()

        total_hours = df["Duration (hrs)"].sum()
        average_per_day = weekly_summary["Total Hours"].mean()
        utilization = round((average_per_day / 8) * 100, 2)

        st.subheader("📅 Daily Task Log")
        st.dataframe(df[["Date", "Task", "Start", "End", "Duration (hrs)"]])

        st.subheader("📊 Weekly Summary")
        st.dataframe(weekly_summary)

        st.subheader("🧾 Performance Summary")
        st.write(f"**Total Hours Logged:** {total_hours:.2f} hrs")
        st.write(f"**Average per Day:** {average_per_day:.2f} hrs")
        st.write(f"**Utilization:** {utilization:.2f}% (based on 8hr/day)")

        if utilization < 50:
            st.warning("⚠️ Low utilization detected. Consider redistributing workload.")
        elif utilization > 100:
            st.warning("⚠️ Over-utilization! Possible burnout.")
        else:
            st.success("✅ Utilization is within healthy range.")
    except Exception as e:
        st.error("Failed to analyze.")
        st.exception(e)

if uploaded_file:
    try:
        xl = pd.ExcelFile(uploaded_file)
        for sheet in xl.sheet_names:
            st.markdown(f"### 📄 {sheet}")
            df_raw = xl.parse(sheet_name=sheet, header=None)
            parsed_df = parse_timesheet(df_raw)
            if parsed_df.empty:
                st.warning("No valid data extracted.")
            else:
                analyze_timesheet(parsed_df)
    except Exception as e:
        st.error("❌ Error reading Excel file.")
        st.exception(e)
