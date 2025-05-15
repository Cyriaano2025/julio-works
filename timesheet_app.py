import streamlit as st
import pandas as pd
import numpy as np
import difflib

st.title("Weekly Timesheet Analyzer")
uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

# Define possible column names for matching
POSSIBLE_START = ['start time', 'in time', 'check in', 'reporting time']
POSSIBLE_END = ['finish time', 'end time', 'check out', 'closing time']
POSSIBLE_TASK = ['task', 'activity', 'task description', 'description']

def match_column(possible_names, columns):
    for name in possible_names:
        match = difflib.get_close_matches(name.lower(), [col.lower() for col in columns], n=1, cutoff=0.6)
        if match:
            return [col for col in columns if col.lower() == match[0]][0]
    return None

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        compiled_data = []

        for sheet_name in excel_file.sheet_names:
            df = excel_file.parse(sheet_name)
            df.columns = df.iloc[0]  # Assume first row is headers
            df = df.drop(df.index[0])

            start_col = match_column(POSSIBLE_START, df.columns)
            end_col = match_column(POSSIBLE_END, df.columns)
            task_col = match_column(POSSIBLE_TASK, df.columns)

            if not start_col or not end_col:
                st.warning(f"Sheet '{sheet_name}' skipped. Missing time columns.")
                continue

            df = df[[start_col, end_col] + ([task_col] if task_col else [])].copy()
            df.columns = ['Start', 'End'] + (['Task'] if task_col else [])
            df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
            df['End'] = pd.to_datetime(df['End'], errors='coerce')
            df['Hours'] = (df['End'] - df['Start']).dt.total_seconds() / 3600
            df['Week'] = sheet_name
            compiled_data.append(df)

        if compiled_data:
            all_data = pd.concat(compiled_data, ignore_index=True)
            all_data = all_data.dropna(subset=['Hours'])
            total_hours = all_data['Hours'].sum()
            avg_per_day = all_data.groupby('Week')['Hours'].mean().mean()
            utilization = round((avg_per_day / 8) * 100, 2)

            st.subheader("Performance Summary")
            st.write(f"**Total Hours Logged:** {round(total_hours, 2)} hrs")
            st.write(f"**Average per Day:** {round(avg_per_day, 2)} hrs/day")
            st.write(f"**Utilization:** {utilization}% of 8-hour workday")

        else:
            st.error("No usable data found in the uploaded Excel file.")

    except Exception as e:
        st.error(f"Could not process the timesheet. Please check formatting.\n\n{e}")
