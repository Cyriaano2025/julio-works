import re
import torch
import streamlit as st
import pandas as pd
import numpy as np
from sentence_transformers import SentenceTransformer, util

st.title("🧠 Smart Timesheet Analyzer")
st.write("Upload any Excel timesheet — no matter the format — and this AI-powered tool will analyze it.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

# Load AI model
model = SentenceTransformer('all-MiniLM-L6-v2')

# Define expected header categories
target_headers = {
    "task": ["task", "activity", "description", "duty", "assignment"],
    "start": ["start time", "start", "reporting time", "begin", "login"],
    "end": ["end time", "finish", "closing time", "logout", "stop"],
    "week": ["week", "date", "day", "period"]
}

# AI-based header matcher
def match_headers(columns):
    embeddings = model.encode(columns, convert_to_tensor=True)
    matched = {}
    for key, options in target_headers.items():
        target_embed = model.encode(options, convert_to_tensor=True)
        cos_scores = util.pytorch_cos_sim(target_embed, embeddings)
        best_score, best_idx = torch.max(cos_scores.max(dim=0).values, dim=0)
        matched[key] = columns[best_idx]
    return matched

# Fallback column matcher using regex
def find_column(columns, keywords):
    for col in columns:
        for keyword in keywords:
            if re.search(keyword, str(col), re.IGNORECASE):
                return col
    return None

# Main analysis function
def analyze_timesheet(df):
    try:
        df.columns = df.iloc[0]  # assume first row is header
        df = df[1:]
        df = df.dropna(how="all")  # remove empty rows

        # Normalize columns to string
        columns = list(df.columns.astype(str))
        header_map = match_headers(columns)

        # Rename matched columns
        df = df.rename(columns={
            header_map["task"]: "Task",
            header_map["start"]: "Start",
            header_map["end"]: "End",
            header_map["week"]: "Week"
        })

        # Fallback detection in case matching missed something
        start_col = find_column(df.columns, ["start", "report", "in", "begin"])
        end_col = find_column(df.columns, ["end", "close", "out", "finish"])

        if start_col and end_col:
            df[start_col] = pd.to_datetime(df[start_col], errors='coerce')
            df[end_col] = pd.to_datetime(df[end_col], errors='coerce')
            df["Start"] = df[start_col]
            df["End"] = df[end_col]
        else:
            st.error("⚠️ Could not detect 'Start' and 'End' time columns.")
            st.stop()

        # Calculate duration
        df["Duration (hrs)"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600

        # Compute metrics
        total_hours = df["Duration (hrs)"].sum()
        avg_per_day = df.groupby("Week")["Duration (hrs)"].sum().mean()
        utilization = round((avg_per_day / 8) * 100, 2)

        # Display results
        st.subheader("📊 Weekly Summary")
        st.dataframe(df[["Week", "Task", "Start", "End", "Duration (hrs)"]])

        st.subheader("🧾 Performance Summary")
        st.write(f"**Total Hours Logged:** {total_hours:.2f} hrs")
        st.write(f"**Average per Week/Day:** {avg_per_day:.2f} hrs")
        st.write(f"**Utilization:** {utilization:.2f}% (based on 8hr/day)")

        # AI Recommendation
        if utilization < 50:
            st.warning("⚠️ Low utilization detected. Consider redistributing workload.")
        elif utilization > 100:
            st.warning("⚠️ Over-utilization! Possible burnout.")
        else:
            st.success("✅ Utilization is within healthy range.")

    except Exception as e:
        st.error("❌ Could not process the timesheet. Please check formatting.")
        st.exception(e)

# Streamlit file upload handler
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
