import streamlit as st
import pandas as pd

st.title("🧠 Smart Timesheet Analyzer")
st.write("Upload the standardized Echo House timesheet to analyze daily logs from May 2025.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

def analyze_standard_timesheet(df):
    try:
        # Ensure correct columns
        expected_columns = ["Date", "Task Description", "Start Time", "End Time"]
        if not all(col in df.columns for col in expected_columns):
            st.error("❌ The uploaded sheet must contain these columns: Date, Task Description, Start Time, End Time.")
            return

        df = df.dropna(subset=["Date", "Start Time", "End Time"])
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
        df["Start Time"] = pd.to_datetime(df["Start Time"], errors="coerce")
        df["End Time"] = pd.to_datetime(df["End Time"], errors="coerce")

        df = df.dropna(subset=["Date", "Start Time", "End Time"])
        df["Duration (hrs)"] = (df["End Time"] - df["Start Time"]).dt.total_seconds() / 3600
        df = df[df["Duration (hrs)"] > 0]

        daily_summary = df.groupby("Date")["Duration (hrs)"].sum().reset_index()
        total_hours = df["Duration (hrs)"].sum()
        avg_daily_hours = daily_summary["Duration (hrs)"].mean()

        st.subheader("📅 Daily Summary")
        st.dataframe(daily_summary)

        st.subheader("🧾 Performance Overview")
        st.write(f"**Total Hours Logged:** {total_hours:.2f} hrs")
        st.write(f"**Average per Day:** {avg_daily_hours:.2f} hrs")

        if avg_daily_hours < 6:
            st.warning("⚠️ Daily average is low. Consider redistributing workload.")
        elif avg_daily_hours > 9:
            st.warning("⚠️ High daily average. Possible overworking.")
        else:
            st.success("✅ Working hours appear healthy.")

    except Exception as e:
        st.error("❌ An error occurred while analyzing the sheet.")
        st.exception(e)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, sheet_name="May 2025 Timesheet")
        st.markdown("### 📄 Sheet: May 2025 Timesheet")
        analyze_standard_timesheet(df)
    except Exception as e:
        st.error("❌ Failed to read the uploaded Excel file.")
        st.exception(e)
