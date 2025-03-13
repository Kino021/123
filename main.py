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
    exclude_users = ['FGPANGANIBAN', 'KPILUSTRISIMO', 'BLRUIZ', 'MMMEJIA', 'SAHERNANDEZ', 'GPRAMOS',
                     'JGCELIZ', 'SPMADRID', 'RRCARLIT', 'MEBEJER',
                     'SEMIJARES', 'GMCARIAN', 'RRRECTO', 'EASORIANO', 'EUGALERA','JATERRADO','LMLABRADOR']
    df = df[~df['Remark By'].isin(exclude_users)]

    # Create the columns layout
    col1, col2 = st.columns(2)

    with col1:
        # Date input from the user
        selected_date = st.date_input("Select Date", min_value=df['Date'].min(), max_value=df['Date'].max(), value=df['Date'].max())

        # Filter the dataframe by the selected date
        filtered_df = df[df['Date'] == pd.to_datetime(selected_date)]

        # Count rows where 'Call Status' contains 'CONNECTED' for the selected date
        connected_count = filtered_df['Call Status'].str.contains('CONNECTED', case=False, na=False).sum()
        
        st.subheader(f"Total Connected Calls on {selected_date}")
        st.write(connected_count)
