import pandas as pd
import streamlit as st

# Set up the page configuration
st.set_page_config(layout="wide", page_title="PRODUCTIVITY", page_icon="📊", initial_sidebar_state="expanded")

# Apply dark mode styling
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

# Title of the app
st.title('Daily Remark Summary')

# Data loading function with file upload support
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# File uploader for Excel file
uploaded_file = st.sidebar.file_uploader("Upload Daily Remark File", type="xlsx")

if uploaded_file is not None:
    df = load_data(uploaded_file)

    # Ensure 'Time' column is in datetime format
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce').dt.time

    # Filter out specific users based on 'Remark By'
    exclude_users = ['FGPANGANIBAN', 'KPILUSTRISIMO', 'BLRUIZ', 'MMMEJIA', 'SAHERNANDEZ', 'GPRAMOS',
                     'JGCELIZ', 'SPMADRID', 'RRCARLIT', 'MEBEJER',
                     'SEMIJARES', 'GMCARIAN', 'RRRECTO', 'EASORIANO', 'EUGALERA','JATERRADO','LMLABRADOR']
    df = df[~df['Remark By'].isin(exclude_users)]

    # Create the columns layout
    col1, col2 = st.columns(2)

    with col1:
        st.write("## Summary Table by Collector per Day")

        # Add date filter
        min_date = df['Date'].min().date()
        max_date = df['Date'].max().date()
        start_date, end_date = st.date_input("Select date range", [min_date, max_date], min_value=min_date, max_value=max_date)

        filtered_df = df[(df['Date'].dt.date >= start_date) & (df['Date'].dt.date <= end_date)]

        # Initialize an empty DataFrame for the summary table by collector
        collector_summary = pd.DataFrame(columns=[ 
            'Day', 'Collector', 'Total Connected', 'Total PTP', 'Total RPC', 'PTP Amount', 'Balance Amount'
        ])

        # Group by 'Date' and 'Remark By' (Collector)
        for (date, collector), collector_group in filtered_df[~filtered_df['Remark By'].str.upper().isin(['SYSTEM'])].groupby([filtered_df['Date'].dt.date, 'Remark By']):
            # Calculate the metrics
            total_connected = collector_group[collector_group['Call Status'] == 'CONNECTED']['Account No.'].count()
            total_ptp = collector_group[collector_group['Status'].str.contains('PTP', na=False) & (collector_group['PTP Amount'] != 0)]['Account No.'].nunique()
            total_rpc = collector_group[collector_group['Status'].str.contains('RPC', na=False)]['Account No.'].nunique()
            ptp_amount = collector_group[collector_group['Status'].str.contains('PTP', na=False) & (collector_group['PTP Amount'] != 0)]['PTP Amount'].sum()

            # Filter rows where PTP Amount is not zero for balance calculation
            balance_amount = collector_group[(collector_group['Status'].str.contains('PTP', na=False)) & (collector_group['PTP Amount'] != 0)]['Balance'].sum()

            # Add the row to the summary
            collector_summary = pd.concat([collector_summary, pd.DataFrame([{
                'Day': date,
                'Collector': collector,
                'Total Connected': total_connected,
                'Total PTP': total_ptp,
                'Total RPC': total_rpc,
                'PTP Amount': ptp_amount,
                'Balance Amount': balance_amount,
            }])], ignore_index=True)

        # Calculate and append totals for the collector summary
        total_row = pd.DataFrame([{
            'Day': 'Total',
            'Collector': '',
            'Total Connected': collector_summary['Total Connected'].sum(),
            'Total PTP': collector_summary['Total PTP'].sum(),
            'Total RPC': collector_summary['Total RPC'].sum(),
            'PTP Amount': collector_summary['PTP Amount'].sum(),
            'Balance Amount': collector_summary['Balance Amount'].sum(),
        }])

        collector_summary = pd.concat([collector_summary, total_row], ignore_index=True)

        # Round off numeric columns to 2 decimal places
        collector_summary[['PTP Amount', 'Balance Amount']] = collector_summary[['PTP Amount', 'Balance Amount']].round(2)

        st.write(collector_summary)

    with col2:
        st.write("## Summary Table by Time Interval per Cycle")

        # Define time intervals (6 AM - 6:59 AM, 7 AM - 7:59 AM, etc.)
        time_bins = [
            ("6 AM - 6:59 AM", pd.to_datetime("06:00:00").time(), pd.to_datetime("06:59:59").time()),
            ("7 AM - 7:59 AM", pd.to_datetime("07:00:00").time(), pd.to_datetime("07:59:59").time()),
            ("8 AM - 8:59 AM", pd.to_datetime("08:00:00").time(), pd.to_datetime("08:59:59").time()),
            ("9 AM - 9:59 AM", pd.to_datetime("09:00:00").time(), pd.to_datetime("09:59:59").time()),
            ("10 AM - 10:59 AM", pd.to_datetime("10:00:00").time(), pd.to_datetime("10:59:59").time()),
            ("11 AM - 11:59 AM", pd.to_datetime("11:00:00").time(), pd.to_datetime("11:59:59").time()),
            ("12 PM - 12:59 PM", pd.to_datetime("12:00:00").time(), pd.to_datetime("12:59:59").time()),
            ("1 PM - 1:59 PM", pd.to_datetime("13:00:00").time(), pd.to_datetime("13:59:59").time()),
            ("2 PM - 2:59 PM", pd.to_datetime("14:00:00").time(), pd.to_datetime("14:59:59").time()),
            ("3 PM - 3:59 PM", pd.to_datetime("15:00:00").time(), pd.to_datetime("15:59:59").time()),
            ("4 PM - 4:59 PM", pd.to_datetime("16:00:00").time(), pd.to_datetime("16:59:59").time()),
            ("5 PM - 5:59 PM", pd.to_datetime("17:00:00").time(), pd.to_datetime("17:59:59").time()),
            ("6 PM - 6:59 PM", pd.to_datetime("18:00:00").time(), pd.to_datetime("18:59:59").time()),
            ("7 PM - 7:59 PM", pd.to_datetime("19:00:00").time(), pd.to_datetime("19:59:59").time()),
            ("8 PM - 8:59 PM", pd.to_datetime("20:00:00").time(), pd.to_datetime("20:59:59").time()),
            ("9 PM - 9:59 PM", pd.to_datetime("21:00:00").time(), pd.to_datetime("21:59:59").time()),
            ("10 PM - 10:59 PM", pd.to_datetime("22:00:00").time(), pd.to_datetime("22:59:59").time()),
            ("11 PM - 11:59 PM", pd.to_datetime("23:00:00").time(), pd.to_datetime("23:59:59").time())
        ]

        # Create a dictionary to map time intervals to sortable integers
        time_interval_sort_order = {time[0]: i for i, time in enumerate(time_bins)}

        # Function to categorize the time into the intervals
        def categorize_time_interval(time_obj):
            if pd.isna(time_obj):
                return "Invalid Time"
            if isinstance(time_obj, pd.Timestamp):
                time_obj = time_obj.time()
            for label, start, end in time_bins:
                if start <= time_obj <= end:
                    return label
            return "Out of Range"

        # Apply categorization based on the existing "Time" column
        df['Time Interval'] = df['Time'].apply(categorize_time_interval)

        # Group the data by Cycle (Service No.) and generate separate tables
        for cycle, cycle_group in df.groupby('Service No.'):
            st.write(f"### Summary Table for Cycle: {cycle}")

            # Initialize an empty DataFrame for the summary table by time interval
            cycle_time_summary = pd.DataFrame(columns=[ 
                'Cycle', 'Time Interval', 'Total Connected', 'Total PTP', 'Total RPC', 'PTP Amount', 'Balance Amount'
            ])

            # Group by Time Interval within the current cycle
            for time_interval, time_interval_group in cycle_group.groupby('Time Interval'):
                # Calculate the metrics for the current time interval and cycle
                total_connected = time_interval_group[time_interval_group['Call Status'] == 'CONNECTED']['Account No.'].count()
                total_ptp = time_interval_group[time_interval_group['Status'].str.contains('PTP', na=False) & (time_interval_group['PTP Amount'] != 0)]['Account No.'].nunique()
                total_rpc = time_interval_group[time_interval_group['Status'].str.contains('RPC', na=False)]['Account No.'].nunique()
                ptp_amount = time_interval_group[time_interval_group['Status'].str.contains('PTP', na=False) & (time_interval_group['PTP Amount'] != 0)]['PTP Amount'].sum()

                # Filter rows where PTP Amount is not zero for balance calculation
                balance_amount = time_interval_group[(time_interval_group['Status'].str.contains('PTP', na=False)) & (time_interval_group['PTP Amount'] != 0)]['Balance'].sum()

                # Add the row to the summary
                cycle_time_summary = pd.concat([cycle_time_summary, pd.DataFrame([{
                    'Cycle': cycle,
                    'Time Interval': time_interval,
                    'Total Connected': total_connected,
                    'Total PTP': total_ptp,
                    'Total RPC': total_rpc,
                    'PTP Amount': ptp_amount,
                    'Balance Amount': balance_amount,
                }])], ignore_index=True)

            # Sort by the time interval to ensure the correct order
            cycle_time_summary['Time Interval'] = pd.Categorical(
                cycle_time_summary['Time Interval'], categories=time_interval_sort_order.keys(), ordered=True)
            cycle_time_summary = cycle_time_summary.sort_values('Time Interval')

            # Round off numeric columns to 2 decimal places
            cycle_time_summary[['PTP Amount', 'Balance Amount']] = cycle_time_summary[['PTP Amount', 'Balance Amount']].round(2)

            # Add totals row at the bottom for cycle-based summary
            totals_row_cycle = {
                'Cycle': cycle,
                'Time Interval': 'Total',
                'Total Connected': cycle_time_summary['Total Connected'].sum(),
                'Total PTP': cycle_time_summary['Total PTP'].sum(),
                'Total RPC': cycle_time_summary['Total RPC'].sum(),
                'PTP Amount': cycle_time_summary['PTP Amount'].sum(),
                'Balance Amount': cycle_time_summary['Balance Amount'].sum(),
            }
            cycle_time_summary = pd.concat([cycle_time_summary, pd.DataFrame([totals_row_cycle])], ignore_index=True)

            # Display the cycle-based summary table
            st.write(cycle_time_summary)
