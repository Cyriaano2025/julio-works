import streamlit as st
import pandas as pd
import difflib

st.set_page_config(page_title="Smart Timesheet Analyzer")
st.title("Smart Timesheet Analyzer")
st.markdown("Upload any Excel timesheet and get real-time insights.")

uploaded_file = st.file_uploader("Upload your Excel timesheet", type=["xlsx"])

# Define keyword sets
START_KEYWORDS = ['start', 'check in', 'report', 'sign in', 'in']
END_KEYWORDS = ['end', 'check out', 'finish', 'sign out', 'out']
TASK_KEYWORDS = ['task', 'description', 'activity', 'work done']

def find_best_match(columns, keyword_list):
    for keyword in keyword_list:
        match = difflib.get_close_matches(keyword.lower(), [str(col).lower() for col in columns], n=1, cutoff=0.6)
        if match:
            for col in columns:
                if col.lower() == match[0]:
                    return col
    return None

def detect_header_row(df):
    for i in range(min(10, len(df))):
        potential_headers = df.iloc[i].dropna().astype(str).str.lower().tolist()
        if any(any(key in col for key in START_KEYWORDS + END_KEYWORDS + TASK_KEYWORDS) for col in potential_headers):
            return i
    return None

def process_sheet(sheet_df, sheet_name):
    header_row = detect_header_row(sheet_df)
    if header_row is None:
        return None
    
    df = sheet_df.copy()
    df.columns = df.iloc[header_row]
    df = df.drop(index=range(header_row + 1)).reset_index(drop=True)

    start_col = find_best_match(df.columns, START_KEYWORDS)
    end_col = find_best_match(df.columns, END_KEYWORDS)
    task_col = find_best_match(df.columns, TASK_KEYWORDS)

    if not start_col or not end_col:
        return None

    df = df[[start_col, end_col] + ([task_col] if task_col else [])]
    df.columns = ['Start', 'End'] + (['Task'] if task_col else [])
    df['Start'] = pd.to_datetime(df['Start'], errors='coerce')
    df['End'] = pd.to_datetime(df['End'], errors='coerce')
    df = df.dropna(subset=['Start', 'End'])

    df['Hours'] = (df['End'] - df['Start']).dt.total_seconds() / 3600
    df['Week'] = sheet_name
    return df

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        all_data = []
        for sheet in xls.sheet_names:
            raw_df = xls.parse(sheet)
            processed = process_sheet(raw_df, sheet)
            if processed is not None:
                all_data.append(processed)

        if all_data:
            final_df = pd.concat(all_data, ignore_index=True)
            total_hours = final_df['Hours'].sum()
            avg_per_day = final_df.groupby('Week')['Hours'].mean().mean()
            utilization = round((avg_per_day / 8) * 100, 2)

            st.subheader("Performance Summary")
            st.write(f"**Total Hours Logged:** {round(total_hours, 2)} hrs")
            st.write(f"**Average per Day:** {round(avg_per_day, 2)} hrs/day")
            st.write(f"**Utilization:** {utilization}% of 8-hour workday")

            st.subheader("Recommendation")
            if utilization >= 80:
                st.success("Excellent work pace! You're utilizing your time very well.")
            elif utilization >= 50:
                st.info("Decent productivity, but there’s room for improvement.")
            else:
                st.warning("Low utilization. Try to log your hours consistently and take on more work.")

        else:
            st.error("No usable data found in this file. Please check formatting.")

    except Exception as e:
        st.error(f"Error processing file: {e}")
