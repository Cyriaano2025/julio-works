import streamlit as st
import pandas as pd
import numpy as np
import torch
import re
from sentence_transformers import SentenceTransformer, util

st.set_page_config(page_title="Smart Timesheet Analyzer", layout="centered")
st.title("🧠 Smart Timesheet Analyzer")
st.markdown("Upload any Excel timesheet — no matter the format — and this AI-powered tool will analyze it.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

# Load semantic model
model = SentenceTransformer('all-MiniLM-L6-v2')

# Common header keyword mapping
target_headers = {
    "task": ["task", "activity", "description", "duty"],
    "start": ["start time", "start", "reporting", "login", "in"],
    "end": ["end time", "finish", "closing", "logout", "out"],
    "week": ["week", "day", "period", "date"]
}

def match_headers(columns):
    embeddings = model.encode([str(col) for col in columns], convert_to_tensor=True)
    matched = {}
    for key, options in target_headers.items():
        targets = model.encode(options, convert_to_tensor=True)
        cos_scores = util.pytorch_cos_sim(targets, embeddings)
        best_score, best_idx = torch.max(cos_scores.max(dim=0).values, dim=0)
        matched[key] = columns[best_idx]
    return matched

def robust_datetime_parse(series):
    return pd.to_datetime(series, errors='coerce', format=None, dayfirst=True, infer_datetime_format=True)

def analyze_timesheet(df, sheet_name):
    try:
        df.columns = df.iloc[0]
        df = df[1:].dropna(how="all")
        df = df.loc[:, ~df.columns.duplicated()]
        df.columns = df.columns.astype(str)

        header_map = match_headers(df.columns)

        # Rename to standard columns
        df = df.rename(columns={
            header_map["task"]: "Task",
            header_map["start"]: "Start",
            header_map["end"]: "End",
            header_map["week"]: "Week"
        })

        df["Start"] = robust_datetime_parse(df["Start"])
        df["End"] = robust_datetime_parse(df["End"])
        df["Week"] = df["Week"].fillna(method="ffill")

        if df["Start"].isna().all() or df["End"].isna().all():
            st.warning(f"⚠️ Time conversion failed in **{sheet_name}**.")
            return

        df["Duration (hrs)"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600
        df = df.dropna(subset=["Duration (hrs)"])

        st.subheader(f"📄 Sheet: {sheet_name}")
        st.dataframe(df[["Week", "Task", "Start", "End", "Duration (hrs)"]])

        total_hours = df["Duration (hrs)"].sum()
        avg_per_day = df.groupby("Week")["Duration (hrs)"].sum().mean()
        utilization = round((avg_per_day / 8) * 100, 2)

        st.markdown("### 📊 Performance Summary")
        st.markdown(f"- **Total Hours Logged:** {total_hours:.2f} hrs")
        st.markdown(f"- **Average per Week/Day:** {avg_per_day:.2f} hrs")
        st.markdown(f"- **Utilization:** {utilization:.2f}% (based on 8hr/day)")

        if utilization < 50:
            st.warning("⚠️ Low utilization detected. Consider redistributing workload.")
        elif utilization > 100:
            st.warning("⚠️ Over-utilization! Possible burnout.")
        else:
            st.success("✅ Utilization is within healthy range.")

    except Exception as e:
        st.error(f"❌ Failed to analyze {sheet_name}.")
        st.exception(e)

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        for sheet_name in xls.sheet_names:
            df = xls.parse(sheet_name=sheet_name, dtype=str)
            analyze_timesheet(df, sheet_name)
    except Exception as e:
        st.error("❌ Failed to load Excel file.")
        st.exception(e)
