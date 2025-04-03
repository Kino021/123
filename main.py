import streamlit as st
import pandas as pd

st.set_page_config(layout="wide", page_title="DIALER PRODUCTIVITY PER CRITERIA OF BALANCE", page_icon="📊", initial_sidebar_state="expanded")

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

st.title('DIALER REPORT PER CRITERIA OF BALANCE')

@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# File uploader
uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx'])

if uploaded_file is not None:
    # Load the data
    df = load_data(uploaded_file)
    
    # Create summary table
    summary_table = pd.DataFrame(columns=[
        'Date', 'ENVIRONMENT', 'COLLECTOR', 'TOTAL CONNECTED', 'TOTAL ACCOUNT', 'TOTAL TALK TIME'
    ])
    
    # Process the data
    # Convert date column to datetime
    df['Date'] = pd.to_datetime(df.iloc[:, 2], format='%d-%m-%Y')
    
    # Group by Date and Environment
    grouped = df.groupby([df['Date'].dt.date, df.iloc[:, 0]])  # Column C for date, Column A for environment
    
    # Calculate metrics
    summary_data = []
    for (date, env), group in grouped:
        # Count unique collectors (Column E)
        unique_collectors = group.iloc[:, 4].nunique()
        
        # Total connected (all rows excluding header)
        total_connected = len(group)
        
        # Total unique accounts (Column D - Customer Name)
        total_accounts = group.iloc[:, 3].nunique()
        
        # Convert talk time to timedelta and sum (Column I)
        talk_times = pd.to_timedelta(group.iloc[:, 8].astype(str))
        total_talk_time = talk_times.sum()
        
        # Format total talk time
        total_seconds = int(total_talk_time.total_seconds())
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        total_talk_time_str = f"{hours:02d}:{minutes:02d}:{seconds:02d}"
        
        summary_data.append({
            'Date': date,
            'ENVIRONMENT': env,
            'COLLECTOR': unique_collectors,
            'TOTAL CONNECTED': total_connected,
            'TOTAL ACCOUNT': total_accounts,
            'TOTAL TALK TIME': total_talk_time_str
        })
    
    # Create summary table
    summary_table = pd.DataFrame(summary_data)
    
    # Display the summary table
    st.subheader("Summary Report")
    st.dataframe(
        summary_table.style.format({
            'Date': '{:%d-%m-%Y}',
            'COLLECTOR': '{:,.0f}',
            'TOTAL CONNECTED': '{:,.0f}',
            'TOTAL ACCOUNT': '{:,.0f}'
        }),
        height=500,
        use_container_width=True
    )
    
    # Add download button for the summary
    csv = summary_table.to_csv(index=False)
    st.download_button(
        label="Download Summary as CSV",
        data=csv,
        file_name="dialer_summary_report.csv",
        mime="text/csv",
    )
else:
    st.info("Please upload an Excel file to generate the report.")

# Add some basic styling to the table
st.markdown(
    """
    <style>
    thead tr th {
        color: white !important;
        background-color: #4A4A4A !important;
    }
    tbody tr:nth-child(odd) {
        background-color: #3A3A3A;
    }
    </style>
    """,
    unsafe_allow_html=True
)
