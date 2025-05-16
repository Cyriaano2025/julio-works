import streamlit as st
import pandas as pd
import numpy as np
import torch
import re
from sentence_transformers import SentenceTransformer, util

st.set_page_config(page_title="Smart Timesheet Analyzer", layout="wide")

st.title("🧠 Smart Timesheet Analyzer")
st.write("Upload any Excel timesheet — no matter the format — and this AI-powered tool will analyze it.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

model = SentenceTransformer('all-MiniLM-L6-v2')

target_keywords = {
    "task": ["task", "activity", "description", "work", "assignment"],
    "start": ["start", "report", "login", "begin", "check-in"],
    "end": ["end", "close", "logout", "finish", "check-out"],
    "week": ["week", "date", "period", "day"]
}

def match_best_column(columns, keywords):
    try:
        column_embeddings = model.encode(columns, convert_to_tensor=True)
        keyword_embeddings = model.encode(keywords, convert_to_tensor=True)
        cos_scores = util.cos_sim(keyword_embeddings, column_embeddings)
        best_score_idx = torch.argmax(cos_scores.max(dim=0).values).item()
        return columns[best_score_idx]
    except Exception:
        return None

def analyze_sheet(df, sheet_name):
    st.markdown(f"### 📄 Sheet: {sheet_name}")
    df.columns = df.iloc[0] if df.iloc[0].isnull().sum() < len(df.columns) else df.columns
    df = df[1:] if df.iloc[0].isnull().sum() < len(df.columns) else df
    df = df.dropna(how="all")

    original_cols = list(df.columns.astype(str))

    task_col = match_best_column(original_cols, target_keywords["task"])
    start_col = match_best_column(original_cols, target_keywords["start"])
    end_col = match_best_column(original_cols, target_keywords["end"])
    week_col = match_best_column(original_cols, target_keywords["week"])

    if not start_col or not end_col:
        st.warning("⚠️ Could not detect 'Start' and 'End' time columns.")
        return

    rename_map = {}
    if task_col: rename_map[task_col] = "Task"
    rename_map[start_col] = "Start"
    rename_map[end_col] = "End"
    if week_col: rename_map[week_col] = "Week"

    df = df.rename(columns=rename_map)

    try:
        df["Start"] = pd.to_datetime(df["Start"], errors="coerce")
        df["End"] = pd.to_datetime(df["End"], errors="coerce")
        df["Duration (hrs)"] = (df["End"] - df["Start"]).dt.total_seconds() / 3600
    except:
        st.warning("⚠️ Time conversion failed.")
        return

    df = df.dropna(subset=["Duration (hrs)"])
    total_hours = df["Duration (hrs)"].sum()
    avg_duration = df["Duration (hrs)"].mean()
    utilization = round((avg_duration / 8) * 100, 2)

    st.success("✅ Timesheet analyzed!")
    st.dataframe(df[["Week", "Task", "Start", "End", "Duration (hrs)"]])

    st.markdown("### Summary")
    st.write(f"**Total Hours Logged:** {total_hours:.2f} hrs")
    st.write(f"**Average Duration:** {avg_duration:.2f} hrs")
    st.write(f"**Utilization (vs. 8-hr day):** {utilization:.2f}%")

    if utilization < 50:
        st.warning("⚠️ Low utilization. May indicate under-reporting.")
    elif utilization > 100:
        st.warning("⚠️ Over-utilization. Check for possible overtime.")
    else:
        st.info("✅ Utilization is healthy.")

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)
    for sheet in excel_file.sheet_names:
        df = excel_file.parse(sheet)
        analyze_sheet(df, sheet)
