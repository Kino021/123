import streamlit as st
import pandas as pd

st.set_page_config(layout="wide", page_title="Daily Remark Summary", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode
st.markdown(
    """
    <style>
    .reportview-container {
        background: #2E2E2E;
        color: white;
    }
    .sidebar .sidebar-content {
        background: #2E2E2E;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.title('Daily Remark Summary')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)

    # Convert 'Date' to datetime if it isn't already
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')

    # Exclude rows where the date is a Sunday (weekday() == 6)
    df = df[df['Date'].dt.weekday != 6]  # 6 corresponds to Sunday

    return df

uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)

    # Exclude rows where STATUS contains 'BP' (Broken Promise) or 'ABORT'
    df = df[~df['Status'].str.contains('ABORT', na=False)]

    # Exclude rows where REMARK contains certain keywords or phrases
    excluded_remarks = [
        "Broken Promise",
        "New files imported", 
        "Updates when case reassign to another collector", 
        "NDF IN ICS", 
        "FOR PULL OUT (END OF HANDLING PERIOD)", 
        "END OF HANDLING PERIOD"
    ]
    df = df[~df['Remark'].str.contains('|'.join(excluded_remarks), case=False, na=False)]

    # Display Per Cycle Manual Summary Table (Separated)
    st.write("## Per Cycle Manual Summary Table")
    
    cycles = df['Service No.'].unique()
    for cycle in cycles:
        cycle_df = df[df['Service No.'] == cycle]
        cycle_manual_summary = calculate_per_cycle_manual_summary(cycle_df)
        st.write(f"### Cycle {cycle}")
        st.write(cycle_manual_summary)

    # Summary Per Balance - Separated by Balance Ranges
    def calculate_summary_per_balance(df):
        balance_ranges = [
            (6000.00, 49999.00),
            (50000.00, 99999.00),
            (100000.00, float('inf'))  # 100,000 and up
        ]

        balance_tables = {}

        for (balance_min, balance_max) in balance_ranges:
            # Filter the dataframe by balance range
            df_filtered = df[(df['Balance'] >= balance_min) & (df['Balance'] <= balance_max)]

            summary_table = pd.DataFrame(columns=[ 
                'Day', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE (%)', 'CONNECTED #', 
                'CONNECTED RATE (%)', 'CONNECTED ACC', 'PTP ACC', 'PTP RATE', 'CALL DROP #', 
                'SYSTEM DROP', 'CALL DROP RATIO #'
            ]) 

            for date, group in df_filtered.groupby(df_filtered['Date'].dt.date):
                accounts = group['Account No.'].nunique()
                total_dialed = group['Account No.'].count()
                connected = group[group['Call Status'] == 'CONNECTED']['Account No.'].nunique()
                penetration_rate = (total_dialed / accounts * 100) if accounts != 0 else None
                connected_acc = group[group['Call Status'] == 'CONNECTED']['Account No.'].count()
                connected_rate = (connected_acc / total_dialed * 100) if total_dialed != 0 else None
                ptp_acc = group[(group['Status'].str.contains('PTP', na=False)) & (group['PTP Amount'] != 0)]['Account No.'].nunique()
                ptp_rate = (ptp_acc / connected * 100) if connected != 0 else None
                system_drop = group[(group['Status'].str.contains('DROPPED', na=False)) & (group['Remark By'] == 'SYSTEM')]['Account No.'].count()
                call_drop_count = group[(group['Status'].str.contains('NEGATIVE CALLOUTS - DROP CALL', na=False)) & 
                                        (~group['Remark By'].str.upper().isin(['SYSTEM']))]['Account No.'].count()
                call_drop_ratio = (system_drop / connected_acc * 100) if connected_acc != 0 else None

                summary_table = pd.concat([summary_table, pd.DataFrame([{
                    'Day': date,
                    'ACCOUNTS': accounts,
                    'TOTAL DIALED': total_dialed,
                    'PENETRATION RATE (%)': f"{round(penetration_rate)}%" if penetration_rate is not None else None,
                    'CONNECTED #': connected,
                    'CONNECTED RATE (%)': f"{round(connected_rate)}%" if connected_rate is not None else None,
                    'CONNECTED ACC': connected_acc,
                    'PTP ACC': ptp_acc,
                    'PTP RATE': f"{round(ptp_rate)}%" if ptp_rate is not None else None,
                    'CALL DROP #': call_drop_count,
                    'SYSTEM DROP': system_drop,
                    'CALL DROP RATIO #': f"{round(call_drop_ratio)}%" if call_drop_ratio is not None else None,
                }])], ignore_index=True)

            balance_tables[f"${balance_min:,.2f} - ${balance_max:,.2f}"] = summary_table

        return balance_tables

    # Example of overall summaries (you can modify these functions as per your requirement)
    def calculate_total_combined_summary(df):
        # Dummy function for the example; customize as needed
        return df.describe()

    def calculate_overall_predictive(df):
        # Dummy function for the example; customize as needed
        return df.groupby('Date')['Account No.'].nunique()

    def calculate_overall_manual(df):
        # Dummy function for the example; customize as needed
        return df.groupby('Service No.').agg({'Account No.': 'nunique'})

    def calculate_summary_per_cycle_predictive(df):
        # Dummy function for the example; customize as needed
        return df.groupby('Service No.').agg({'Balance': 'sum'})

    def calculate_summary_per_cycle_manual(df):
        # Dummy function for the example; customize as needed
        return df.groupby('Service No.').agg({'Call Status': 'count'})

    # Display all results
    st.write("## Total Combined Summary")
    st.write(calculate_total_combined_summary(df))

    st.write("## Overall Predictive Summary")
    st.write(calculate_overall_predictive(df))

    st.write("## Overall Manual Summary")
    st.write(calculate_overall_manual(df))

    st.write("## Summary Per Cycle (Predictive)")
    st.write(calculate_summary_per_cycle_predictive(df))

    st.write("## Summary Per Cycle (Manual)")
    st.write(calculate_summary_per_cycle_manual(df))

    st.write("## SUMMARY PER BALANCE")
    balance_tables = calculate_summary_per_balance(df)
    for balance_range, table in balance_tables.items():
        st.write(f"### Balance Range: {balance_range}")
        st.write(table)
