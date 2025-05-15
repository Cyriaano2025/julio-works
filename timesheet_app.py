import pandas as pd
import streamlit as st

st.title("Weekly Timesheet Analyzer")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

if uploaded_file:
    excel_file = pd.ExcelFile(uploaded_file)

    def extract_standardized_data(sheet_name, df):
        df.columns = df.iloc[1]
        df = df.drop([0, 1]).reset_index(drop=True)
        df['Week'] = sheet_name
        return df[['BRAND -  TASK DESCRIPTION ', 'START TIME', 'FINISH TIME', 'Week']]

    all_sheets = {name: excel_file.parse(name) for name in excel_file.sheet_names}
    compiled_data = pd.concat([extract_standardized_data(name, df) for name, df in all_sheets.items()], ignore_index=True)

    def compute_hours(row):
        try:
            start = pd.to_datetime(row['START TIME'], format='%H:%M:%S', errors='coerce')
            end = pd.to_datetime(row['FINISH TIME'], format='%H:%M:%S', errors='coerce')
            if pd.notnull(start) and pd.notnull(end) and end > start:
                return (end - start).total_seconds() / 3600
        except:
            return None
        return None

    compiled_data['Duration (hrs)'] = compiled_data.apply(compute_hours, axis=1)

    total_hours = compiled_data['Duration (hrs)'].sum()
    avg_per_day = compiled_data.groupby('Week')['Duration (hrs)'].mean().mean()
    utilization = round((avg_per_day / 8) * 100, 2)

    st.subheader("Performance Summary")
    st.write(f"**Total Hours Logged**: {total_hours:.2f} hrs")
    st.write(f"**Average per Day**: {avg_per_day:.2f} hrs/day")
    st.write(f"**Utilization**: {utilization:.2f}% of 8-hour workday")

    st.subheader("Observations & Recommendations")
    if avg_per_day > 6:
        st.success("✅ Great job maintaining consistency!")
        st.info("Keep up the strong momentum and continue delivering at this pace.")
    else:
        st.warning("⚠️ You’ve logged below-average work hours.")
        st.info("Try to log hours more consistently and take on more tasks.")

    st.markdown("🎯 Let’s aim for at least **6 productive hours/day** next week!")
    # Create feedback text
    feedback = f"""
Hi [Employee],

Here’s your time performance summary for the period reviewed:

- Total Hours Logged: {total_hours:.2f} hrs
- Average per Day: {avg_per_day:.2f} hrs/day
- Utilization of 8-hour Workday: {utilization:.2f}%

Observations:
{"✅ Great job maintaining consistency!" if avg_per_day > 6 else "⚠️ You’ve logged below-average work hours."}

Recommendations:
{"- Maintain your focus and task momentum." if avg_per_day > 6 else "- Log your hours more consistently and take on more tasks."}

🎯 Let’s aim for at least 6 productive hours/day next week!
"""

    # Download button
    st.download_button(
        label="Download Feedback as TXT",
        data=feedback,
        file_name="timesheet_feedback.txt",
        mime="text/plain"
    )

