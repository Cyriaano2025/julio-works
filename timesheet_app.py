import re
import torch
import streamlit as st
import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer, util

st.set_page_config(page_title="Smart Timesheet Analyzer", layout="wide")
st.title("🧠 Smart Timesheet Analyzer")
st.write("Upload any Excel timesheet — no matter the format — and this AI-powered tool will analyze it.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

# Load AI embedding model once
model = SentenceTransformer('all-MiniLM-L6-v2')

# Helper to find best-matching header
def find_column(columns, keywords):
    for col in columns:
        for keyword in keywords:
            if keyword in col:
                return col
    return None

def analyze_timesheet(df):
    try:
        # Look through top 10 rows to identify header row
        potential_header_row = None
        for i in range(min(10, len(df))):
            row = df.iloc[i].astype(str).str.lower()
            if any(re.search(r"(start|report|login|in|check[- ]?in)", cell) for cell in row):
                potential_header_row = i
                break

        if potential_header_row is None:
            st.warning("⚠️ Could not detect a valid header row.")
            return

        df.columns = df.iloc[potential_header_row]
        df = df[potential_header_row + 1:]
        df = df.dropna(how="all")
        df.columns = df.columns.astype(str).str.strip().str.lower().str.replace(r"[^a-z0-9 ]", "", regex=True)

        # Match relevant columns
        start_col = find_column(df.columns, ["start", "report", "login", "in", "checkin"])
        end_col = find_column(df.columns, ["end", "close", "logout", "out", "checkout"])
        task_col = find_column(df.columns, ["task", "desc", "activity", "duty"])
        week_col = find_column(df.columns, ["week", "day", "date", "period"])

        if not start_col or not end_col:
            st.warning("⚠️ Could not detect 'Start' and 'End' time columns.")
            return

        df["Start"] = pd.to_datetime(df[start_col], errors='coerce')
        df["End"] = pd.to_datetime(df[end_col], errors='coerce')
        df["Task"] = df[task_col] if task_col else "N/A"
        df["Week"] = df[week_col] if week_col else "Unknown"

        df["Duration (hrs)"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600

        total_hours = df["Duration (hrs)"].sum()
        avg_per_day = df.groupby("Week")["Duration (hrs)"].sum().mean()
        utilization = round((avg_per_day / 8) * 100, 2)

        st.subheader("📊 Weekly Breakdown")
        st.dataframe(df[["Week", "Task", "Start", "End", "Duration (hrs)"]], use_container_width=True)

        st.subheader("🧾 Summary")
        st.write(f"**Total Hours Logged:** {total_hours:.2f} hrs")
        st.write(f"**Average per Week/Day:** {avg_per_day:.2f} hrs")
        st.write(f"**Utilization:** {utilization:.2f}% (based on 8hr/day)")

        if utilization < 50:
            st.warning("⚠️ Low utilization detected. Consider redistributing workload.")
        elif utilization > 100:
            st.warning("⚠️ Over-utilization! Possible burnout.")
        else:
            st.success("✅ Utilization is within a healthy range.")

    except Exception as e:
        st.error("❌ Could not process the timesheet.")
        st.exception(e)

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        for sheet in excel_file.sheet_names:
            st.markdown(f"## 📄 Sheet: {sheet}")
            df = excel_file.parse(sheet_name=sheet)
            analyze_timesheet(df)
    except Exception as e:
        st.error("❌ Failed to read the file.")
        st.exception(e)
