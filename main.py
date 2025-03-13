import pandas as pd
import streamlit as st

# Set up the page configuration
st.set_page_config(layout="wide", page_title="QWERTY", page_icon="📊", initial_sidebar_state="expanded")

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

    # Ensure 'Time' column is in datetime format and 'Date' column is properly converted
    df['Time'] = pd.to_datetime(df['Time'], errors='coerce').dt.time
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce')  # Ensure the 'Date' column is in datetime format

    # Filter out specific users based on 'Remark By'
    exclude_users = ['JTAGRAMON']
    df = df[~df['Remark By'].isin(exclude_users)]

    # Create the columns layout
    col1, col2 = st.columns(2)

    with col1:
       # Get the total calls (Remark By contains "COLLECTOR" and exclude "SYSTEM")
total_calls = df[df['Remark By'].str.contains("COLLECTOR", case=False, na=False) & 
                 ~df['Remark By'].str.contains("SYSTEM", case=False, na=False)]
total_calls_count = total_calls.groupby('Remark By').size()

# Get the total connected (Call Status contains "CONNECTED" and exclude if Status contains "EMAIL")
total_connected = df[df['Call Status'].str.contains("CONNECTED", case=False, na=False) & 
                     ~df['Status'].str.contains("EMAIL", case=False, na=False)]
total_connected_count = total_connected.groupby('Remark By').size()

# Get the total RPC (Status contains "POSITIVE")
total_rpc = df[df['Status'].str.contains("POSITIVE", case=False, na=False)]
total_rpc_count = total_rpc.groupby('Remark By').size()

# Get the total PTP count (Status contains "PTP" and exclude if PTP Amount contains 0.00)
total_ptp = df[df['Status'].str.contains("PTP", case=False, na=False) & 
               ~df['PTP Amount'].astype(str).str.contains("0.00", na=False)]
total_ptp_count = total_ptp.groupby('Remark By').size()

# Get the total PTP amount (PTP Amount contains an amount and exclude if 0.00)
total_ptp_amount = df[df['PTP Amount'].astype(str).str.contains(r'\d+\.\d{2}', na=False) & 
                      ~df['PTP Amount'].astype(str).str.contains("0.00", na=False)]
total_ptp_amount_sum = total_ptp_amount.groupby('Remark By')['PTP Amount'].sum()

# Get the total OB (based on your data structure, if it's under the 'Balance' column or other logic)
total_ob = df[df['Balance'].notna() & (df['Balance'] > 0)]  # Adjust condition if OB is defined differently
total_ob_sum = total_ob.groupby('Remark By')['Balance'].sum()

# Get total talk time (Summing the 'Talk Time' column if present)
total_talk_time = df.groupby('Remark By')['Talk Time'].sum()
