import streamlit as st
import pandas as pd

st.title("Weekly Timesheet Analyzer")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        all_sheets = {name: excel_file.parse(name) for name in excel_file.sheet_names}
        
        def extract_relevant_data(sheet_name, df):
            df.columns = df.iloc[0]
            df = df.drop(index=0).reset_index(drop=True)

            start_col = next((col for col in df.columns if 'start' in str(col).lower()), None)
            end_col = next((col for col in df.columns if 'finish' in str(col).lower() or 'end' in str(col).lower()), None)
            task_col = next((col for col in df.columns if 'task' in str(col).lower() or 'description' in str(col).lower()), None)

            df['Week'] = sheet_name
            return df[[task_col, start_col, end_col, 'Week']].rename(columns={
                task_col: 'TASK',
                start_col: 'START TIME',
                end_col: 'FINISH TIME'
            })

        compiled_data = pd.concat([
            extract_relevant_data(name, df) for name, df in all_sheets.items()
        ], ignore_index=True)

        def compute_hours(row):
            try:
                start = pd.to_datetime(row['START TIME'], errors='coerce')
                end = pd.to_datetime(row['FINISH TIME'], errors='coerce')
                if pd.notnull(start) and pd.notnull(end) and end > start:
                    return (end - start).total_seconds() / 3600
            except:
                return None
            return None

        compiled_data['Duration (hrs)'] = compiled_data.apply(compute_hours, axis=1)
        compiled_data = compiled_data.dropna(subset=['Duration (hrs)'])

        total_hours = compiled_data['Duration (hrs)'].sum()
        avg_per_day = compiled_data.groupby('Week')['Duration (hrs)'].mean().mean()
        utilization = round((avg_per_day / 8) * 100, 2)

        st.subheader("Performance Summary")
        st.write(f"**Total Hours Logged**: {total_hours:.2f} hrs")
        st.write(f"**Average per Day**: {avg_per_day:.2f} hrs/day")
        st.write(f"**Utilization**: {utilization:.2f}% of 8-hour workday")

    except Exception as e:
        st.error("Could not process the timesheet. Please check formatting.")
        st.exception(e)
