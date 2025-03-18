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

    # Exclude rows where 'Debtor' contains "DEFAULT_LEAD_"
    df = df[~df['Debtor'].str.contains('DEFAULT_LEAD_', na=False)]

    # Check if data is empty after filtering
    if df.empty:
        st.warning("No valid data available after filtering.")
    else:
        # Function to generate the summary table for a specific balance range and date
        def generate_balance_summary(df, balance_range_name, lower_limit, upper_limit):
            summary_table = pd.DataFrame(columns=[ 
                'Date', 'Balance Range', 'ACCOUNTS', 'TOTAL DIALED', 'PENETRATION RATE (%)', 'CONNECTED #', 
                'CONNECTED RATE (%)', 'CONNECTED ACC', 'PTP ACC', 'PTP RATE', 'CALL DROP #', 
                'SYSTEM DROP', 'CALL DROP RATIO #', 'TOTAL PTP AMOUNT', 'TOTAL BALANCE'
            ])

            # Filter data for current balance range
            balance_filtered_group = df[(df['Balance'] >= lower_limit) & (df['Balance'] <= upper_limit)]

            # Group by Date
            for date, group in balance_filtered_group.groupby(df['Date'].dt.date):
                # Calculate the various metrics
                accounts = group[group['Remark Type'].isin(['Predictive', 'Follow Up', 'Outgoing'])]['Account No.'].nunique()
                total_dialed = group[group['Remark Type'].isin(['Predictive', 'Follow Up', 'Outgoing'])]['Account No.'].count()
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

                # Calculate the TOTAL PTP AMOUNT and TOTAL BALANCE
                # First, filter the rows where PTP Amount is not zero or NaN
                ptp_data = group[group['PTP Amount'].notna() & (group['PTP Amount'] != 0)]

                # Sum the PTP Amount for unique Account No.
                total_ptp_amount = ptp_data.groupby('Account No.')['PTP Amount'].sum().sum()

                # Sum the Balance for those same unique Account No.'s
                total_balance = ptp_data.groupby('Account No.')['Balance'].sum().sum()

                summary_table = pd.concat([summary_table, pd.DataFrame([{
                    'Date': date,
                    'Balance Range': balance_range_name,
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
                    'TOTAL PTP AMOUNT': total_ptp_amount,
                    'TOTAL BALANCE': total_balance,
                }])], ignore_index=True)

            return summary_table

        # Balance ranges to consider
        balance_ranges = [
            ('6,000.00 - 49,999.99', 6000, 49999.99),
            ('50,000.00 - 99,999.99', 50000, 99999.99),
            ('100,000.00 - UP', 100000, float('inf'))
        ]

        # Display results for each balance range, grouped by Date
        for range_name, lower_limit, upper_limit in balance_ranges:
            st.subheader(f"Summary for Balance Range: {range_name}")
            balance_summary = generate_balance_summary(df, range_name, lower_limit, upper_limit)

            if not balance_summary.empty:
                st.dataframe(balance_summary)
            else:
                st.warning(f"No data available for the {range_name} range.")
