import streamlit as st
import pandas as pd

st.set_page_config(page_title="Weekly Timesheet Analyzer", layout="centered")

st.title("Weekly Timesheet Analyzer")
st.markdown("Upload your Excel timesheet")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

if uploaded_file:
    try:
        excel_file = pd.ExcelFile(uploaded_file)

        def extract_relevant_data(sheet_name, df):
            df.columns = df.iloc[0]  # Set first row as header
            df = df.drop(index=0).reset_index(drop=True)

            # Auto-detect columns
            task_col = next((col for col in df.columns if isinstance(col, str) and 'task' in col.lower()), None)
            start_col = next((col for col in df.columns if isinstance(col, str) and 'start' in col.lower()), None)
            end_col = next((col for col in df.columns if isinstance(col, str) and ('end' in col.lower() or 'finish' in col.lower())), None)

            if not all([task_col, start_col, end_col]):
                raise ValueError(f"Missing required columns in sheet '{sheet_name}'. Found columns: {df.columns.tolist()}")

            df['Week'] = sheet_name
            return df[[task_col, start_col, end_col, 'Week']].rename(columns={
                task_col: 'TASK',
                start_col: 'START TIME',
                end_col: 'FINISH TIME'
            })

        # Process all sheets
        all_data = pd.concat([
            extract_relevant_data(sheet_name, excel_file.parse(sheet_name))
            for sheet_name in excel_file.sheet_names
        ], ignore_index=True)

        # Calculate duration in hours
        def compute_hours(row):
            try:
                start = pd.to_datetime(row['START TIME'], errors='coerce')
                end = pd.to_datetime(row['FINISH TIME'], errors='coerce')
                if pd.notnull(start) and pd.notnull(end):
                    return (end - start).total_seconds() / 3600
            except:
                return None
            return None

        all_data['Duration (hrs)'] = all_data.apply(compute_hours, axis=1)

        total_hours = all_data['Duration (hrs)'].sum()
        avg_per_day = all_data.groupby('Week')['Duration (hrs)'].mean().mean()
        utilization = round((avg_per_day / 8) * 100, 2)

        # Show performance summary
        st.subheader("Performance Summary")
        st.write(f"**Total Hours Logged:** {round(total_hours, 2)} hrs")
        st.write(f"**Average per Day:** {round(avg_per_day, 2)} hrs/day")
        st.write(f"**Utilization:** {utilization}% of 8-hour workday")

        # Optional insights
        if utilization < 50:
            st.warning("You've logged below-average hours. Try to log more consistently!")
        elif utilization >= 80:
            st.success("Great job! You're working at high efficiency.")
        else:
            st.info("You're doing okay. A little boost and you'll be there!")

    except Exception as e:
        st.error(f"Could not process the timesheet. Please check formatting.\n\nError: {e}")
